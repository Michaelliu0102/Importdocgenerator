"""
报关资料生成器 - 主程序
根据进口Invoice自动生成合同、报关单和申报要素
"""

import os
import re
import sys
import argparse
import subprocess
from pathlib import Path
from datetime import datetime


def _safe_filename(name: str) -> str:
    """Replace characters illegal in file/dir names."""
    return re.sub(r'[\\/:*?"<>|]+', '_', name)

from pdf_parser import InvoiceParser
from docx_generator import DeclarationElementsGenerator
from config_loader import ConfigLoader
from excel_filler import ExcelFiller


class CustomsDocGenerator:
    """报关资料生成器"""

    def __init__(
        self,
        config_path: str = "data/supplier_product_mapping.yaml",
        templates_dir: str = "templates"
    ):
        self.config_loader = ConfigLoader(config_path)
        self.templates_dir = Path(templates_dir)
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
        invoice_no = _safe_filename(self.invoice_data.get("invoice_no", "UNKNOWN"))
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
        invoice_no = _safe_filename(self.invoice_data.get("invoice_no", "UNKNOWN"))
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
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        invoice_no = _safe_filename(self.invoice_data.get("invoice_no", "UNKNOWN"))
        output_file = output_dir / f"申报要素_{invoice_no}_{timestamp}.docx"

        try:
            generator = DeclarationElementsGenerator()
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

    def _print_parsed_data(self):
        d = self.invoice_data
        print(f"   发票号: {d.get('invoice_no', 'N/A')}")
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
                        help="模板目录路径")
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
        templates_dir=args.templates
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
