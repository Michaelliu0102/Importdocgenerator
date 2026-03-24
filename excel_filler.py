"""
Excel模板填充模块
用于填充合同模板(CONTRACT_CAMARI_PRETTY.xlsx / CONTRACT.xlsx)和报关单模板(FedEx报关单模板.xlsx)
"""

from typing import Dict, Any
from openpyxl import load_workbook


def _find_cell_contains(ws, text: str):
    """返回第一个包含 text 的单元格（不区分大小写）。"""
    t = text.lower()
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and t in str(cell.value).lower():
                return cell
    return None


def _unit_to_english(unit: str) -> str:
    mapping = {
        "米": "M",
        "张": "PCS",
        "平方英尺": "SQFT",
        "千克": "KG",
    }
    return mapping.get(unit, unit or "")


def _to_float(value):
    """将字符串/数字安全转换为 float，失败返回 None。"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def _calc_gram_weight(net_weight, quantity):
    """
    计算每平方米克重(g/m2): net_weight / (quantity * 1.42) * 1000
    说明：按用户要求公式计算，net_weight 与 quantity 使用同一票据口径。
    """
    nw = _to_float(net_weight)
    qty = _to_float(quantity)
    if nw is None or qty is None or qty == 0:
        return None
    result = nw / (qty * 1.42) * 1000
    if result <= 0:
        return None
    return round(result, 2)


class ExcelFiller:
    def __init__(self, template_path: str):
        self.template_path = template_path

    def fill_contract_template(
        self,
        invoice_data: Dict[str, Any],
        supplier_info: Dict[str, Any],
        product_info: Dict[str, Any],
        output_path: str
    ):
        wb = load_workbook(self.template_path)
        ws = wb.active

        title = str(ws.cell(row=2, column=2).value or "").strip().upper()
        if title in {"CAMARI IMPORT CONTRACT", "CONTRACT"}:
            self._fill_contract_pretty(ws, invoice_data, supplier_info, product_info)
        else:
            self._fill_contract_legacy(ws, invoice_data, supplier_info, product_info)

        wb.save(output_path)
        return output_path

    def _fill_contract_pretty(
        self,
        ws,
        invoice_data: Dict[str, Any],
        supplier_info: Dict[str, Any],
        product_info: Dict[str, Any],
    ):
        """填充新版排版合同模板 CONTRACT_CAMARI_PRETTY.xlsx。"""
        items = invoice_data.get("items", [])
        supplier_name = supplier_info.get("name", invoice_data.get("supplier_name", ""))
        supplier_addr = supplier_info.get("address", "")
        trade_term = invoice_data.get("trade_term") or supplier_info.get("trade_term", "")
        currency = invoice_data.get("currency", "EUR") or "EUR"
        invoice_no = invoice_data.get("invoice_no", "")
        invoice_date = invoice_data.get("invoice_date", "")
        payment = invoice_data.get("payment_cond") or supplier_info.get("payment_term", "")

        # Header block
        ws.cell(row=4, column=3).value = "CAMARI TRADING (ZHEJIANG) CO., LTD"
        ws.cell(row=4, column=10).value = invoice_no
        ws.cell(row=5, column=10).value = invoice_date
        ws.cell(row=7, column=3).value = supplier_name
        ws.cell(row=8, column=3).value = supplier_addr
        # 第10行已改为只保留 Incoterms / Payment Terms（无 Country / Currency）
        ws.cell(row=10, column=3).value = trade_term
        ws.cell(row=10, column=7).value = payment

        # Item table rows: 13-24
        for i, item in enumerate(items[:12], 1):
            r = 12 + i
            ws.cell(row=r, column=2).value = i
            ws.cell(row=r, column=3).value = item.get("description", "")

            qty_val = item.get("quantity", "")
            try:
                qty_val = float(qty_val)
            except (TypeError, ValueError):
                pass
            ws.cell(row=r, column=7).value = qty_val
            ws.cell(row=r, column=8).value = _unit_to_english(item.get("unit", "M"))

            up_val = item.get("unit_price", "")
            try:
                up_val = float(up_val)
            except (TypeError, ValueError):
                pass
            ws.cell(row=r, column=9).value = up_val

            ws.cell(row=r, column=10).value = currency

            amt_val = item.get("amount", "")
            try:
                amt_val = float(amt_val)
            except (TypeError, ValueError):
                pass
            ws.cell(row=r, column=11).value = amt_val

        # Total
        total = invoice_data.get("total_amount", "")
        try:
            total = float(total)
        except (TypeError, ValueError):
            pass
        ws.cell(row=25, column=11).value = total

    def _fill_contract_legacy(
        self,
        ws,
        invoice_data: Dict[str, Any],
        supplier_info: Dict[str, Any],
        product_info: Dict[str, Any],
    ):
        """兼容旧版 CONTRACT.xlsx 的填充逻辑。"""

        items = invoice_data.get("items", [])
        supplier_name = supplier_info.get("name", invoice_data.get("supplier_name", ""))
        trade_term = invoice_data.get("trade_term") or supplier_info.get("trade_term", "")
        currency = invoice_data.get("currency", "EUR") or "EUR"
        invoice_no = invoice_data.get("invoice_no", "")
        invoice_date = invoice_data.get("invoice_date", "")
        payment = invoice_data.get("payment_cond") or supplier_info.get("payment_term", "")

        # 买方固定（与模板一致）
        ws.cell(row=4, column=3).value = "CAMARI TRADING (ZHEJIANG) CO., LTD"

        # 合同编号 = Invoice number
        ws.cell(row=5, column=11).value = invoice_no

        ws.cell(row=7, column=3).value = supplier_name
        ws.cell(row=7, column=11).value = invoice_date

        # 第12行：货币 / Incoterms / Payment（CAMARI 模板）
        b12 = ws.cell(row=12, column=2).value
        if b12 and "currency" in str(b12).lower():
            ws.cell(row=12, column=3).value = currency
            ws.cell(row=12, column=6).value = trade_term or ""
            ws.cell(row=12, column=9).value = payment or ""
        else:
            # 旧模板：贸易术语写在首行单价列旁
            if trade_term and items:
                ws.cell(row=15, column=10).value = trade_term

        # 表头币种
        hdr = _find_cell_contains(ws, "name of commodity")
        header_row = hdr.row if hdr else 14
        ws.cell(row=header_row, column=10).value = f"Unit Price {currency}"
        ws.cell(row=header_row, column=11).value = f"Amount {currency}"

        first_item_row = header_row + 1

        for i, item in enumerate(items):
            r = first_item_row + i
            ws.cell(row=r, column=3).value = item.get("description", "")

            try:
                ws.cell(row=r, column=8).value = float(item.get("quantity", 0))
            except (ValueError, TypeError):
                ws.cell(row=r, column=8).value = item.get("quantity", "")

            try:
                ws.cell(row=r, column=10).value = float(item.get("unit_price", 0))
            except (ValueError, TypeError):
                ws.cell(row=r, column=10).value = item.get("unit_price", "")

            try:
                ws.cell(row=r, column=11).value = float(item.get("amount", 0))
            except (ValueError, TypeError):
                ws.cell(row=r, column=11).value = item.get("amount", "")

        try:
            total = float(invoice_data.get("total_amount", 0))
        except (ValueError, TypeError):
            total = invoice_data.get("total_amount", "")

        tot_cell = _find_cell_contains(ws, "TOT AMOUNT")
        if tot_cell is not None:
            tot_cell.value = f"TOT AMOUNT: {total} {currency}".strip()

        place = _find_cell_contains(ws, "Place of Shipment")
        if place is not None:
            country = supplier_info.get("country", "")
            if country:
                ws.cell(row=place.row, column=4).value = f"From {country}"

        pay_clause = _find_cell_contains(ws, "(3) Terms of payment")
        if pay_clause is not None and payment:
            ws.cell(row=pay_clause.row, column=5).value = payment

    def fill_customs_declaration_template(
        self,
        invoice_data: Dict[str, Any],
        supplier_info: Dict[str, Any],
        product_info: Dict[str, Any],
        output_path: str,
        product_info_map: Dict[str, Dict] = None,
    ):
        """
        填充FedEx报关单模板 (FedEx报关单模板.xlsx)

        product_info_map: optional {hide_type: product_info} for per-item HS/elements.
        """
        wb = load_workbook(self.template_path)

        items = invoice_data.get("customs_items") or invoice_data.get("items", [])
        currency = invoice_data.get("currency", "EUR")
        trade_term = invoice_data.get("trade_term", supplier_info.get("trade_term", ""))
        supplier_country = supplier_info.get("country", "")
        net_weight = invoice_data.get("net_weight", "")

        if not product_info_map:
            product_info_map = {}

        # === Sheet 1: 企业信息 ===
        if "1.企业信息" in wb.sheetnames:
            ws1 = wb["1.企业信息"]
            ws1.cell(row=9, column=2).value = invoice_data.get("invoice_no", "")

        # === Sheet 2: 商品信息 ===
        if "2.商品信息" in wb.sheetnames:
            ws2 = wb["2.商品信息"]

            item_count = len(items) if items else 1
            per_item_weight = ""
            if net_weight:
                try:
                    per_item_weight = f"{float(net_weight) / item_count:.2f}"
                except (ValueError, TypeError):
                    per_item_weight = net_weight

            for i, item in enumerate(items):
                r = 2 + i
                qty = item.get("quantity", "")

                # Pick product_info for this item: by hide_type, then default
                hide_type = item.get("hide_type", "")
                cur_pi = product_info_map.get(hide_type) or product_info or {}
                hs_code = cur_pi.get("hs_code", "")
                product_name_cn = cur_pi.get("name", "")
                decl_elements = cur_pi.get("declaration_elements", {})

                gram_weight = _calc_gram_weight(net_weight, qty)

                item_decl_elements = dict(decl_elements)
                for k, v in list(item_decl_elements.items()):
                    key_has_gram = ("克重" in str(k)) or ("g/m2" in str(k).lower())
                    if not key_has_gram:
                        continue
                    val = "" if v is None else str(v).strip()
                    if (not val or val == "克/平方米") and gram_weight is not None:
                        item_decl_elements[k] = f"{gram_weight} g/m2"

                decl_parts = []
                for idx, (k, v) in enumerate(item_decl_elements.items(), 1):
                    if v:
                        decl_parts.append(f"{idx}:{k}:{v}")
                decl_text = "; ".join(decl_parts) if decl_parts else ""

                prefix = item.get("item_code_prefix", "")
                if prefix == "5012":
                    cn_name = "汽车装饰用材料"
                    item_hs = "5603149000"
                elif prefix in ("5015", "5030", "5010"):
                    cn_name = "面料"
                    item_hs = "5603149000"
                else:
                    cn_name = product_name_cn
                    item_hs = hs_code

                ws2.cell(row=r, column=1).value = cn_name
                ws2.cell(row=r, column=2).value = item_hs
                ws2.cell(row=r, column=3).value = decl_text
                ws2.cell(row=r, column=4).value = per_item_weight

                unit = item.get("unit", "米")
                ws2.cell(row=r, column=5).value = f"{qty} {unit}"

                origin = _country_code_to_chinese(item.get("country_code", ""))
                if not origin:
                    origin = supplier_country
                ws2.cell(row=r, column=6).value = origin

                amount = item.get("amount", "")
                ws2.cell(row=r, column=7).value = f"{currency} {amount}"

                ws2.cell(row=r, column=8).value = trade_term
                ws2.cell(row=r, column=9).value = supplier_country

        wb.save(output_path)
        return output_path


def _country_code_to_chinese(code: str) -> str:
    mapping = {
        "IT": "意大利",
        "DE": "德国",
        "GB": "英国",
        "US": "美国",
        "NL": "荷兰",
        "FR": "法国",
        "CN": "中国",
    }
    return mapping.get(code, "")
