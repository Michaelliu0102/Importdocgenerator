"""
报关资料生成器 - 主程序
根据进口Invoice自动生成合同、报关单和申报要素
"""

import copy
import os
import re
import sys
import traceback
import argparse
import subprocess
from pathlib import Path
from datetime import datetime
from typing import List, Optional, Union

# 出口 PDF → export_templates/ 填充；若仍见「待实现」占位 txt，说明运行的是旧代码/未重打包的 .app
EXPORT_PIPELINE_VERSION = "export_templates-v2"
APP_VERSION_LABEL = "v4.3"


def _safe_filename(name: Optional[str]) -> str:
    """Replace characters illegal in file/dir names."""
    # invoice_no 等字段可能显式为 None；.get("k", "default") 在 k 存在且值为 None 时仍会得到 None。
    return re.sub(r'[\\/:*?"<>|]+', "_", str(name or ""))


def _write_export_run_info(
    out: Path,
    source_label: str,
    exp_dir: Path,
    generated_paths: list,
) -> str:
    """写入本次出口运行信息，便于确认是否为新版代码（含 export_templates 填充）。"""
    info = out / "_export_info.txt"
    info.write_text(
        f"pipeline: {EXPORT_PIPELINE_VERSION}\n"
        f"generator: customs_doc_generator (对应 GUI {APP_VERSION_LABEL})\n"
        f"source: {source_label}\n"
        f"export_templates: {exp_dir}\n"
        f"generated:\n"
        + ("\n".join(f"  - {p}" for p in generated_paths) if generated_paths else "  (none)\n")
        + "\n",
        encoding="utf-8",
    )
    return str(info)


from pdf_parser import InvoiceParser
from packing_slip_parser import PackingSlipParser
from docx_generator import DeclarationElementsGenerator
from config_loader import ConfigLoader
from excel_filler import ExcelFiller, _effective_invoice_no
from item_declaration_mapper import (
    build_declaration_groups,
    use_item_mapping_enabled,
)


class CustomsDocGenerator:
    """报关资料生成器"""

    def __init__(
        self,
        config_path: str = "data/supplier_product_mapping.yaml",
        templates_dir: str = "templates",
        export_templates_dir: str = "export_templates",
    ):
        self.config_loader = ConfigLoader(config_path)
        self.templates_dir = Path(templates_dir)
        self.export_templates_dir = Path(export_templates_dir)
        self.invoice_data = None
        self.supplier_info = None
        self.product_info = None

    def process_invoice(
        self,
        invoice_path: str,
        supplier_code: str = None,
        product_code: str = None,
        output_dir: str = "output",
        enable_ocr: bool = False,
        ocr_lang: str = "eng+ita",
    ) -> dict:
        print(f"开始处理Invoice: {invoice_path}")

        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        working_invoice_path = invoice_path
        if enable_ocr:
            working_invoice_path = self._prepare_invoice_with_ocr(
                invoice_path=invoice_path,
                output_path=output_path,
                ocr_lang=ocr_lang,
            )

        # 1. 解析Invoice
        print("\n1. 解析Invoice PDF...")
        parser = InvoiceParser(working_invoice_path)
        self.invoice_data = parser.parse()
        self._print_parsed_data()

        # 2. 匹配供应商
        print("\n2. 匹配供应商信息...")
        if supplier_code:
            self.supplier_info = self.config_loader.get_supplier_info(supplier_code)
        else:
            match_result = self.config_loader.match_supplier_by_name(
                self.invoice_data.get("supplier_name", "")
            )
            if match_result:
                supplier_code, self.supplier_info = match_result
                print(f"   匹配供应商: {self.supplier_info.get('name')} (code: {supplier_code})")
            else:
                print("   警告: 未找到匹配的供应商")

        # 3. 匹配产品
        print("\n3. 匹配产品信息...")
        if product_code:
            self.product_info = self.config_loader.get_product_info(product_code)
        else:
            self.product_info = self._match_product(supplier_code)

        if self.product_info:
            print(f"   匹配产品: {self.product_info.get('name')}")
            print(f"   HS Code: {self.product_info.get('hs_code', '')}")
        else:
            print("   警告: 未找到匹配的产品，使用默认值")

        # 4. 生成合同
        print("\n4. 生成合同...")
        contract_files = self._generate_contract(output_path)

        # 5. 生成报关单
        print("\n5. 生成报关单...")
        customs_files = self._generate_customs_declaration(output_path)

        # 6. 生成申报要素
        print("\n6. 生成申报要素...")
        element_files = self._generate_declaration_elements(output_path)

        result = {
            "contract": contract_files,
            "customs_declaration": customs_files,
            "declaration_elements": element_files,
        }

        print("\n" + "=" * 50)
        print("生成完成！")
        print("=" * 50)
        self._print_output_files(result)

        return result

    @staticmethod
    def _classify_export_pdfs(paths: List[str]) -> tuple[Optional[str], Optional[str]]:
        """区分 CustInvc 发票与 Packing Slip；返回 (invoice_pdf, packing_pdf)。"""
        invoice_path = None
        packing_path = None
        for p in paths:
            if Path(p).suffix.lower() != ".pdf":
                continue
            try:
                head = (InvoiceParser(p).extract_text() or "")[:8000]
            except Exception:
                continue
            if re.search(r"Packing\s+Slip", head, re.IGNORECASE):
                packing_path = p
            elif re.search(r"\bInvoice\b", head, re.IGNORECASE) and (
                "Bill To" in head or re.search(r"Invoice\s*#", head, re.IGNORECASE)
            ):
                invoice_path = p
        return invoice_path, packing_path

    def process_export_documents(
        self,
        source_paths: Union[str, List[str]],
        output_dir: str = "output",
        enable_ocr: bool = False,
        ocr_lang: str = "eng+ita",
        fx_units_per_eur: Optional[float] = None,
        generate_eur_invoice_pdf: bool = False,
    ) -> dict:
        """出口报关资料：解析发票 PDF（及可选装箱单 PDF），填充 export_templates/ 下模板。

        fx_units_per_eur: 若 >0 且发票币种非 EUR，则将金额换算为 EUR（1 EUR = 该数值单位发票货币），
        用于合同、报关单、申报要素等 Excel；与是否生成 EUR Invoice PDF 无关。

        generate_eur_invoice_pdf: 为 True 时，在 camari_cust 且换算后（或原票）为 EUR 时，
        额外生成 CustInvc_EUR_*.pdf；为 False 时不生成该 PDF。
        """
        paths = [source_paths] if isinstance(source_paths, str) else list(source_paths)
        source_label = ", ".join(paths)
        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)
        self.export_templates_dir.mkdir(parents=True, exist_ok=True)
        pdfs = [p for p in paths if Path(p).suffix.lower() == ".pdf"]
        stem = Path(pdfs[0]).stem if pdfs else "export"
        safe = _safe_filename(stem) or "export"
        exp_dir = self.export_templates_dir.resolve()

        if not pdfs:
            note = out / f"出口报关_说明_{safe}.txt"
            note.write_text(
                "【出口报关资料】当前仅支持 PDF。\n"
                f"已关联文件: {source_label}\n"
                f"出口模板目录: {exp_dir}\n",
                encoding="utf-8",
            )
            print(f"   出口: 无 PDF，已写入说明: {note}")
            info_path = _write_export_run_info(out, source_label, exp_dir, [str(note)])
            return {"export": [str(note), info_path]}

        inv_path, pack_path = self._classify_export_pdfs(pdfs)
        if not inv_path:
            try:
                t0 = (InvoiceParser(pdfs[0]).extract_text() or "")[:4000]
            except Exception:
                t0 = ""
            if "Packing Slip" in t0:
                note = out / f"出口报关_缺发票_{safe}.txt"
                note.write_text(
                    "【出口报关资料】仅检测到装箱单 PDF，缺少 Invoice PDF。\n"
                    "请在右侧出口列表同时添加 CustInvc 发票与 ItemShip 装箱单。\n"
                    f"已关联: {source_label}\n",
                    encoding="utf-8",
                )
                print(f"   出口: 缺发票: {note}")
                info_path = _write_export_run_info(out, source_label, exp_dir, [str(note)])
                return {"export": [str(note), info_path]}
            inv_path = pdfs[0]

        src = Path(inv_path)
        working_pdf = str(src)
        if enable_ocr:
            working_pdf = self._prepare_invoice_with_ocr(
                invoice_path=str(src),
                output_path=out,
                ocr_lang=ocr_lang,
            )

        packing_slip_data = None
        if pack_path:
            try:
                packing_slip_data = PackingSlipParser(pack_path).parse()
                print(
                    f"   出口装箱单: {Path(pack_path).name} "
                    f"(件数={packing_slip_data.get('pkg_qty')}, "
                    f"毛重={packing_slip_data.get('gross_weight_kg')}kg, "
                    f"净重={packing_slip_data.get('net_weight_kg')}kg)"
                )
            except Exception as e:
                print(f"   警告: 装箱单解析失败，报关单件数/重量将留空: {e}")

        print("\n[出口] 解析发票 PDF...")
        parser = InvoiceParser(working_pdf)
        self.invoice_data = parser.parse()
        if packing_slip_data:
            inv_no_pdf = (self.invoice_data.get("invoice_no") or "").strip()
            pref = str(packing_slip_data.get("invoice_ref") or "").strip()
            if pref and inv_no_pdf and pref != inv_no_pdf:
                print(
                    f"   警告: 装箱单编号 #{pref} 与发票号 {inv_no_pdf} 不一致，请确认是否为同一票货物。"
                )
        items = self.invoice_data.get("items") or []

        if not items:
            note = out / f"出口报关_解析失败_{safe}.txt"
            note.write_text(
                "【出口报关资料】未能从 PDF 中解析出商品明细。\n"
                f"文件: {source_label}\n"
                f"已识别版式: {self.invoice_data.get('format', 'unknown')}\n"
                f"出口模板目录: {exp_dir}\n"
                "请确认发票版式受支持，或开启 OCR 后重试。\n",
                encoding="utf-8",
            )
            print(f"   出口: 无商品行，已写入: {note}")
            info_path = _write_export_run_info(out, source_label, exp_dir, [str(note)])
            return {"export": [str(note), info_path]}

        print("[出口] 匹配供应商 / 产品...")
        match_result = self.config_loader.match_supplier_by_name(
            self.invoice_data.get("supplier_name", "") or ""
        )
        supplier_code = None
        if match_result:
            supplier_code, self.supplier_info = match_result
            print(f"   出口: 供应商 {self.supplier_info.get('name')} ({supplier_code})")
        else:
            self.supplier_info = {}
            print("   出口: 未匹配供应商配置，按发票字段与默认产品填充")

        self.product_info = self._match_product(supplier_code)
        product_info_map = self._build_product_info_map()

        print(
            f"   [出口] 生成 EUR Invoice PDF: "
            f"{'是' if generate_eur_invoice_pdf else '否'}"
        )

        # 汇率换算前先快照原票，供报关单/合同抬头与合同号使用（与 EUR 金额分离）
        invoice_before_fx = copy.deepcopy(self.invoice_data)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        _inv_key = _effective_invoice_no(invoice_before_fx) or (
            self.invoice_data.get("invoice_no") or ""
        ).strip()
        inv = _safe_filename(_inv_key or "UNKNOWN")
        generated: list[str] = []

        if fx_units_per_eur is not None and fx_units_per_eur > 0:
            src_cur = (self.invoice_data.get("currency") or "").strip().upper()
            # 币种为空时也必须换算：此前用「if src_cur and …」会跳过换算，导致永远无法生成 EUR 版 PDF。
            if src_cur != "EUR":
                from eur_invoice_converter import apply_eur_conversion

                self.invoice_data = apply_eur_conversion(
                    self.invoice_data, float(fx_units_per_eur)
                )
                if (
                    generate_eur_invoice_pdf
                    and self.invoice_data.get("format") != "camari_cust"
                ):
                    print(
                        "   提示: 当前发票版式非 camari_cust，已换算 EUR 数据，"
                        "未生成欧元 Invoice PDF（仅 camari_cust 可生成）。"
                    )
            else:
                print("   提示: 发票已为 EUR，跳过汇率换算。")

        cur_after = (self.invoice_data.get("currency") or "").strip().upper()
        _fmt = self.invoice_data.get("format")
        if generate_eur_invoice_pdf:
            print(
                f"   [EUR PDF 条件] 版式={_fmt!r}, 换算后币种={cur_after!r}, "
                f"汇率={fx_units_per_eur!r}, 发票PDF={Path(working_pdf).name}"
            )
            if cur_after == "EUR" and _fmt == "camari_cust":
                eur_pdf = out / f"CustInvc_EUR_{inv}_{timestamp}.pdf"
                eur_pdf_ok = False
                try:
                    from eur_invoice_pdf_overlay import (
                        can_use_overlay,
                        write_eur_invoice_pdf_from_original,
                    )

                    if can_use_overlay():
                        write_eur_invoice_pdf_from_original(
                            str(working_pdf),
                            str(eur_pdf),
                            self.invoice_data,
                            invoice_before_fx,
                        )
                        print(
                            f"   EUR 版 Invoice PDF（原 PDF 版式替换金额，可再解析）: {eur_pdf}"
                        )
                        eur_pdf_ok = True
                    else:
                        raise ImportError("PyMuPDF 未安装（pip install pymupdf）")
                except Exception as ex:
                    print(f"   EUR overlay 未成功，尝试 ReportLab 回退: {ex!r}")
                    traceback.print_exc()
                    try:
                        from eur_invoice_pdf import write_camari_cust_eur_invoice_pdf

                        write_camari_cust_eur_invoice_pdf(str(eur_pdf), self.invoice_data)
                        print(
                            f"   EUR 版 Invoice PDF（ReportLab 重排）: {eur_pdf}"
                        )
                        eur_pdf_ok = True
                    except Exception as ex2:
                        print(f"   EUR 版 Invoice PDF 生成失败（ReportLab）: {ex2}")
                        traceback.print_exc()
                if eur_pdf_ok:
                    generated.append(str(eur_pdf))
                else:
                    print(
                        "   错误: EUR 版 CustInvc PDF 未写出，请根据上方 Traceback 排查 "
                        "（常见：未 pip install pymupdf，或原 PDF 路径不可读）。"
                    )
            elif (
                _fmt == "camari_cust"
                and cur_after != "EUR"
                and (fx_units_per_eur is None or fx_units_per_eur <= 0)
            ):
                print(
                    "   提示: 未在「1 EUR = … 发票货币」中填写大于 0 的汇率，"
                    "发票仍为原币种，未生成 EUR 版 CustInvc PDF。"
                )
            elif cur_after == "EUR" and _fmt != "camari_cust":
                print(
                    f"   提示: 已换算为 EUR，但版式为 {_fmt!r}（非 camari_cust），"
                    "不生成 CustInvc EUR PDF；仅 camari 内部发票会生成该 PDF。"
                )
            elif _fmt == "camari_cust" and cur_after != "EUR" and (
                fx_units_per_eur is not None and fx_units_per_eur > 0
            ):
                print(
                    f"   提示: 已填汇率但解析币种仍为 {cur_after!r}，未生成 EUR PDF。"
                    "请确认发票币种解析是否正确（控制台见 [EUR PDF 条件]）。"
                )

        contract_tpl = self._find_export_template(["export_contract.xlsx"])
        if contract_tpl:
            outp = out / f"出口合同_{inv}_{timestamp}.xlsx"
            try:
                filler = ExcelFiller(str(contract_tpl))
                filler.fill_export_contract_template(
                    self.invoice_data,
                    self.supplier_info or {},
                    self.product_info or {},
                    str(outp),
                    party_invoice_data=invoice_before_fx,
                )
                print(f"   出口合同: {outp}")
                generated.append(str(outp))
            except Exception as e:
                print(f"   出口合同填充失败: {e}")

        decl_tpl = self._find_export_template(["export_declaration.xlsx"])
        if decl_tpl:
            outp = out / f"出口报关单_{inv}_{timestamp}.xlsx"
            try:
                deg = DeclarationElementsGenerator(self.config_loader.config_path)
                items_for_decl = (
                    self.invoice_data.get("customs_items")
                    or self.invoice_data.get("items", [])
                )
                if use_item_mapping_enabled(self.invoice_data):
                    line_groups = build_declaration_groups(
                        items_for_decl,
                        self.invoice_data,
                        self.product_info or {},
                        product_info_map,
                        self.config_loader.config_path,
                    )
                else:
                    line_groups = deg._group_items_by_product(
                        items_for_decl,
                        self.product_info or {},
                        product_info_map,
                    )
                filler = ExcelFiller(str(decl_tpl))
                filler.fill_export_declaration_template(
                    self.invoice_data,
                    self.supplier_info or {},
                    self.product_info or {},
                    str(outp),
                    line_groups=line_groups,
                    product_info_map=product_info_map,
                    packing_slip_data=packing_slip_data,
                    party_invoice_data=invoice_before_fx,
                )
                print(f"   出口报关单: {outp}")
                generated.append(str(outp))
            except Exception as e:
                print(f"   出口报关单填充失败: {e}")

        # 出口申报要素：必须使用 export_templates/ 下的壳（默认 申报要素总汇.docx），与进口申报要素生成方式不同
        docx_tpl = self._find_export_declaration_elements_template()
        if docx_tpl:
            outp = out / f"出口申报要素_{inv}_{timestamp}.docx"
            try:
                gen = DeclarationElementsGenerator(self.config_loader.config_path)
                gen.generate_from_export_template(
                    str(docx_tpl),
                    self.invoice_data,
                    self.product_info or {},
                    self.supplier_info or {},
                    str(outp),
                    product_info_map=product_info_map,
                )
                print(f"   出口申报要素: {outp}")
                generated.append(str(outp))
            except Exception as e:
                print(f"   出口申报要素生成失败: {e}")

        if not generated:
            note = out / f"出口报关_未生成_{safe}.txt"
            note.write_text(
                "【出口报关资料】未生成 Excel/Word。\n"
                f"文件: {source_label}\n"
                f"请在 {exp_dir} 中放置 export_contract.xlsx、"
                "export_declaration.xlsx、申报要素总汇.docx\n",
                encoding="utf-8",
            )
            generated.append(str(note))

        info_path = _write_export_run_info(out, source_label, exp_dir, generated)
        generated.append(info_path)

        return {"export": generated}

    @staticmethod
    def _get_ocr_env() -> dict:
        """Build environment with homebrew paths for ocrmypdf/tesseract."""
        env = os.environ.copy()
        brew_bin = "/opt/homebrew/bin"
        if brew_bin not in env.get("PATH", ""):
            env["PATH"] = f"{brew_bin}:{env.get('PATH', '')}"
        return env

    def _prepare_invoice_with_ocr(self, invoice_path: str, output_path: Path, ocr_lang: str) -> str:
        """If PDF has no text layer, run OCR and return OCR PDF path."""
        parser = InvoiceParser(invoice_path)
        text = parser.extract_text() or ""
        if len(text.strip()) > 20:
            print("   OCR检查: 检测到文本层，跳过OCR")
            return invoice_path

        print("   OCR检查: 未检测到可用文本层，开始OCR...")

        env = self._get_ocr_env()
        try:
            subprocess.run(
                ["ocrmypdf", "--version"],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                env=env,
            )
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("   警告: 未找到ocrmypdf，无法自动OCR。请先安装后重试。")
            print("   安装示例: brew install tesseract ocrmypdf")
            return invoice_path

        src = Path(invoice_path)
        stem = src.stem
        ocr_file = output_path / f"{stem}_ocr.pdf"

        try:
            subprocess.run(
                [
                    "ocrmypdf",
                    "--force-ocr",
                    "-l",
                    ocr_lang,
                    str(src),
                    str(ocr_file),
                ],
                check=True,
                env=env,
            )
            print(f"   OCR完成: {ocr_file}")
            return str(ocr_file)
        except subprocess.CalledProcessError as e:
            print(f"   警告: OCR失败，使用原始PDF继续。错误: {e}")
            return invoice_path

    def _match_product(self, supplier_code: str = None) -> dict:
        """从解析到的商品信息匹配产品配置"""
        items = self.invoice_data.get("items", [])

        # 优先: 按item_code前缀匹配 (Alcantara规则)
        for item in items:
            item_code = item.get("item_code", "")
            if item_code and len(item_code) >= 4:
                match = self.config_loader.match_product_by_item_code(item_code)
                if match:
                    _, info = match
                    return info

        # 次选: 按商品描述关键词匹配
        for item in items:
            desc = item.get("description", "")
            match = self.config_loader.match_product_by_keywords(desc, supplier_code)
            if match:
                _, info = match
                return info

        # 最后: 按成分匹配 (West Trading面料)
        for item in items:
            composition = item.get("composition", item.get("description", ""))
            match = self.config_loader.match_fabric_by_composition(composition, supplier_code)
            if match:
                _, info = match
                return info

        return None

    def _build_product_info_map(self) -> dict:
        """Build {hide_type: product_info} map for items with different HS codes."""
        products = self.config_loader.get_all_products()
        pim = {}
        for _, info in products.items():
            ht = info.get("hide_type", "")
            if ht:
                pim[ht] = info
        return pim

    def _generate_contract(self, output_dir: Path) -> list:
        template_path = self._find_template([
            "CONTRACT_CAMARI_PRETTY.xlsx",
            "CONTRACT_CAMARI_TEMPLATE.xlsx",
            "CONTRACT.xlsx",
            "合同模板.xlsx",
        ])
        if not template_path:
            print("   警告: 找不到合同模板")
            return []

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        invoice_no = _safe_filename(self.invoice_data.get("invoice_no") or "UNKNOWN")
        output_file = output_dir / f"合同_{invoice_no}_{timestamp}.xlsx"

        try:
            filler = ExcelFiller(str(template_path))
            filler.fill_contract_template(
                self.invoice_data,
                self.supplier_info or {},
                self.product_info or {},
                str(output_file)
            )
            print(f"   已生成: {output_file}")
            return [str(output_file)]
        except Exception as e:
            print(f"   填充失败: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _generate_customs_declaration(self, output_dir: Path) -> list:
        template_path = self._find_template(["FedEx报关单模板.xlsx", "报关单模板.xlsx"])
        if not template_path:
            print("   警告: 找不到报关单模板")
            return []

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        invoice_no = _safe_filename(self.invoice_data.get("invoice_no") or "UNKNOWN")
        output_file = output_dir / f"报关单_{invoice_no}_{timestamp}.xlsx"

        try:
            filler = ExcelFiller(str(template_path))
            filler.fill_customs_declaration_template(
                self.invoice_data,
                self.supplier_info or {},
                self.product_info or {},
                str(output_file),
                product_info_map=self._build_product_info_map(),
            )
            print(f"   已生成: {output_file}")
            return [str(output_file)]
        except Exception as e:
            print(f"   填充失败: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _generate_declaration_elements(self, output_dir: Path) -> list:
        """进口：申报要素为空白 Word 正文生成（不使用 export_templates/申报要素总汇.docx）。"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        invoice_no = _safe_filename(self.invoice_data.get("invoice_no") or "UNKNOWN")
        output_file = output_dir / f"申报要素_{invoice_no}_{timestamp}.docx"

        try:
            generator = DeclarationElementsGenerator(self.config_loader.config_path)
            generator.generate(
                self.invoice_data,
                self.product_info or {},
                self.supplier_info or {},
                str(output_file),
                product_info_map=self._build_product_info_map(),
            )
            print(f"   已生成: {output_file}")
            return [str(output_file)]
        except Exception as e:
            print(f"   填充失败: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _find_template(self, names: list) -> Path:
        for name in names:
            path = self.templates_dir / name
            if path.exists():
                return path
        return None

    def _find_export_template(self, names: list) -> Optional[Path]:
        """在出口专用目录中查找模板（不使用进口 templates/）。"""
        for name in names:
            path = self.export_templates_dir / name
            if path.exists():
                return path
        return None

    def _find_export_declaration_elements_template(self) -> Optional[Path]:
        """
        出口申报要素 docx：仅索引 export_templates/（默认 supplier_product_mapping 中
        export_declaration_element_templates.default，一般为 申报要素总汇.docx）。
        """
        primary = self.config_loader.get_export_declaration_template_filename()
        return self._find_export_template(
            [primary, "export_declaration_elements.docx"]
        )

    def _print_parsed_data(self):
        d = self.invoice_data
        print(f"   发票号: {d.get('invoice_no') or 'N/A'}")
        print(f"   发票日期: {d.get('invoice_date', 'N/A')}")
        print(f"   供应商: {d.get('supplier_name', 'N/A')}")
        print(f"   贸易条款: {d.get('trade_term', 'N/A')}")
        print(f"   净重: {d.get('net_weight', 'N/A')} kg")
        print(f"   毛重: {d.get('gross_weight', 'N/A')} kg")
        print(f"   商品数量: {len(d.get('items', []))}")
        print(f"   总金额: {d.get('currency', '')} {d.get('total_amount', 'N/A')}")

        for i, item in enumerate(d.get('items', []), 1):
            print(f"\n   商品 {i}:")
            print(f"      Item Code: {item.get('item_code', 'N/A')}")
            print(f"      前4位: {item.get('item_code_prefix', 'N/A')}")
            print(f"      描述: {item.get('description', 'N/A')}")
            print(f"      数量: {item.get('quantity', 'N/A')} {item.get('unit', '')}")
            print(f"      单价: {item.get('unit_price', 'N/A')}")
            print(f"      金额: {item.get('amount', 'N/A')}")

        customs_items = d.get('customs_items', [])
        if customs_items and customs_items is not d.get('items'):
            print(f"\n   报关合并商品: {len(customs_items)} 项（按 article 汇总）")
            for i, item in enumerate(customs_items, 1):
                print(f"\n   报关商品 {i}:")
                print(f"      品名: {item.get('article_name', item.get('description', 'N/A'))}")
                print(f"      数量: {item.get('quantity', 'N/A')} {item.get('unit', '')}")
                print(f"      单价: {item.get('unit_price', 'N/A')}")
                print(f"      金额: {item.get('amount', 'N/A')}")

    def _print_output_files(self, result: dict):
        print("\n生成的文件:")
        for doc_type, files in result.items():
            if files:
                print(f"  {doc_type}:")
                for f in files:
                    print(f"    - {f}")


def main():
    parser = argparse.ArgumentParser(description="报关资料生成器")
    parser.add_argument("invoice", nargs="?", help="Invoice PDF文件路径")
    parser.add_argument("-c", "--config", default="data/supplier_product_mapping.yaml",
                        help="配置文件路径")
    parser.add_argument("-t", "--templates", default="templates",
                        help="进口模板目录路径（合同/进口报关单等）")
    parser.add_argument(
        "--export-templates",
        default="export_templates",
        help="出口专用模板目录（与 --templates 分开）",
    )
    parser.add_argument("-s", "--supplier", help="供应商代码")
    parser.add_argument("-p", "--product", help="产品代码")
    parser.add_argument("-o", "--output", default="output",
                        help="输出目录")
    parser.add_argument("--ocr", action="store_true",
                        help="若PDF无文本层则自动执行OCR后再解析")
    parser.add_argument("--ocr-lang", default="eng+ita",
                        help="OCR语言，如 eng+ita 或 eng+ita+chi_sim")

    args = parser.parse_args()

    generator = CustomsDocGenerator(
        config_path=args.config,
        templates_dir=args.templates,
        export_templates_dir=args.export_templates,
    )

    if not args.invoice:
        print("用法: python3 main.py <invoice.pdf>")
        print("示例: python3 main.py /path/to/invoice.pdf")
        return

    generator.process_invoice(
        invoice_path=args.invoice,
        supplier_code=args.supplier,
        product_code=args.product,
        output_dir=args.output,
        enable_ocr=args.ocr,
        ocr_lang=args.ocr_lang,
    )


if __name__ == "__main__":
    main()
