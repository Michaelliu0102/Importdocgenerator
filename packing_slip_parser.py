"""
从 CAMARI Packing Slip PDF（ItemShip_*.pdf）提取件数、毛重/净重、Ship To 国家等。
"""

import re
from typing import Any, Dict, List, Optional, Tuple

from pypdf import PdfReader


def _last_line_country(address_lines: List[str]) -> str:
    if not address_lines:
        return ""
    last = address_lines[-1].strip()
    return last


def country_name_to_cn(name: str) -> str:
    """报关单常用中文国别（可按需扩展）。"""
    n = (name or "").strip()
    m = {
        "Japan": "日本",
        "CHINA": "中国",
        "China": "中国",
        "Korea": "韩国",
        "USA": "美国",
        "United States": "美国",
        "Germany": "德国",
        "Italy": "意大利",
        "France": "法国",
        "UK": "英国",
        "United Kingdom": "英国",
    }
    return m.get(n, n)


class PackingSlipParser:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.raw_text = ""

    def extract_text(self) -> str:
        reader = PdfReader(self.pdf_path)
        parts = []
        for page in reader.pages:
            parts.append(page.extract_text() or "")
        self.raw_text = "\n".join(parts)
        return self.raw_text

    def parse(self) -> Dict[str, Any]:
        if not self.raw_text:
            self.extract_text()

        t = self.raw_text.replace("\u00a0", " ")
        slip_no = self._extract_slip_number(t)
        invoice_ref = self._extract_invoice_ref(t)
        ship_lines = self._extract_ship_to_lines(t)
        ship_country = _last_line_country(ship_lines)

        pkg_qty, gross_kg, net_kg = self._extract_pkg_weights(t)

        return {
            "format": "camari_packing_slip",
            "slip_no": slip_no,
            "invoice_ref": invoice_ref,
            "ship_to_lines": ship_lines,
            "ship_to_country": ship_country,
            "ship_to_country_cn": country_name_to_cn(ship_country),
            "pkg_qty": pkg_qty,
            "gross_weight_kg": gross_kg,
            "net_weight_kg": net_kg,
        }

    @staticmethod
    def _extract_slip_number(text: str) -> Optional[str]:
        m = re.search(r"Packing\s+Slip\s*\n\s*#(\d+)", text, re.IGNORECASE)
        return m.group(1) if m else None

    @staticmethod
    def _extract_invoice_ref(text: str) -> Optional[str]:
        """装箱单正文中的发票号（如 Invoice #32823）。不得使用 Packing Slip # 当作发票号。"""
        m = re.search(r"Invoice\s*#\s*(\d{4,8})\b", text, re.IGNORECASE)
        if m:
            return m.group(1)
        m = re.search(r"Invoice\s+No\.?\s*(\d{4,8})\b", text, re.IGNORECASE)
        if m:
            return m.group(1)
        return None

    def _extract_ship_to_lines(self, text: str) -> List[str]:
        m = re.search(
            r"Ship\s+To\s*\n(.+?)(?:\n\s*Date\s+SO|\nDate\s+SO|\n#|\Z)",
            text,
            re.IGNORECASE | re.DOTALL,
        )
        if not m:
            return []
        chunk = m.group(1)
        lines = [ln.strip() for ln in chunk.split("\n") if ln.strip()]
        return lines

    @staticmethod
    def _extract_pkg_weights(text: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """
        表头行后第一行数据，例如：
        7/4/2026 SO-1128290 FedEx CN 3 58KG 52KG
        """
        if "Pkg Qty" in text or "Gross Weight" in text:
            m_hdr = re.search(
                r"Net\s+Weight\s*\n(.+)",
                text,
                re.IGNORECASE,
            )
            if m_hdr:
                line = m_hdr.group(1).strip().split("\n")[0]
                parsed = PackingSlipParser._parse_data_line(line)
                if parsed[0]:
                    return parsed

        # 无表头时：匹配 ... 数字 KG 数字 KG 结尾
        for line in text.split("\n"):
            line = line.strip()
            m = re.search(
                r"(\d+)\s+([\d.]+)\s*KG\s+([\d.]+)\s*KG\s*$",
                line,
                re.IGNORECASE,
            )
            if m:
                return m.group(1), m.group(2), m.group(3)
        return None, None, None

    @staticmethod
    def _parse_data_line(line: str) -> tuple[Optional[str], Optional[str], Optional[str]]:
        """
        从一行中提取 Pkg Qty、Gross KG、Net KG。
        例：7/4/2026 SO-1128290 FedEx CN 3 58KG 52KG
        """
        m = re.search(
            r"(\d+)\s+([\d.]+)\s*KG\s+([\d.]+)\s*KG\s*$",
            line.strip(),
            re.IGNORECASE,
        )
        if m:
            return m.group(1), m.group(2), m.group(3)
        return None, None, None
