"""
PDF Invoice 解析模块
从进口Invoice PDF中提取关键信息
支持供应商: Alcantara, DECA GLOBAL, Crest, HDM, Mabo, West Trading, Wipelli, Continental, Mastrotto
"""

import re
from typing import Dict, Any, Optional, List


def _parse_euro_amount(s: str) -> str:
    """Parse European-format amounts like '€8 125,20' or '43,92' to a plain decimal string."""
    s = s.replace("€", "").replace("\u00a0", " ").replace("$", "").strip()
    s = s.replace(" ", "").replace(".", "").replace(",", ".")
    return s


class InvoiceParser:
    """PDF Invoice 解析器"""

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.raw_text = ""
        self.parsed_data = {}
        self._fmt: Optional[str] = None

    def extract_text(self) -> str:
        from pypdf import PdfReader
        reader = PdfReader(self.pdf_path)
        text_parts = []
        for page in reader.pages:
            text_parts.append(page.extract_text())
        self.raw_text = "\n".join(text_parts)
        return self.raw_text

    def _detect_format(self) -> str:
        t = self.raw_text.lower()
        if "sede legale" in t and "alcantara" in t:
            return "alcantara"
        if "deca global" in t:
            return "deca"
        if "crest" in t and "leather" in t:
            return "crest"
        if "higher dimension" in t:
            return "hdm"
        if "mabo" in t:
            return "mabo"
        if "westtrading" in t or "west trading" in t:
            return "west_trading"
        if "wipelli" in t:
            return "wipelli"
        if "hornschuch" in t or "continental" in t:
            return "continental"
        if "mastrotto" in t:
            return "mastrotto"
        return "generic"

    def parse(self) -> Dict[str, Any]:
        if not self.raw_text:
            self.extract_text()

        self._fmt = self._detect_format()
        items = self._extract_items()

        customs_items = self._aggregate_customs_items(items)

        self.parsed_data = {
            "invoice_no": self._extract_invoice_number(),
            "invoice_date": self._extract_date(),
            "supplier_name": self._extract_supplier_name(),
            "buyer_name": self._extract_buyer_name(),
            "trade_term": self._extract_trade_term(),
            "country_of_origin": self._extract_country_of_origin(),
            "items": items,
            "customs_items": customs_items,
            "total_amount": self._compute_total(items),
            "currency": self._extract_currency(),
            "net_weight": self._extract_weight("NET"),
            "gross_weight": self._extract_weight("GROSS"),
            "payment_cond": self._extract_payment_cond(),
            "packages": self._extract_packages(),
        }

        return self.parsed_data

    # =========================================================================
    # Invoice Number
    # =========================================================================
    def _extract_invoice_number(self) -> Optional[str]:
        fmt = self._fmt

        if fmt == "alcantara":
            m = re.search(r"(\d{10})INVOICE:", self.raw_text)
            if m:
                return m.group(1)
            # pypdf concatenated: "FATTURA: dd/mm/yyDEL:2025010310"
            m = re.search(r"DEL:(\d{10})", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "deca":
            m = re.search(r"Invoice\s*\n#(\d+)", self.raw_text)
            if m:
                return m.group(1)
            m = re.search(r"#(\d{4,6})\b", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "crest":
            m = re.search(r"([A-Z]{2}\d{2}/\d{4,6})", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "hdm":
            m = re.search(r"Order\s*#\s*\n?\s*([A-Z0-9]+)", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "mabo":
            m = re.search(r"(RE\d{8})", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "west_trading":
            m = re.search(r"Invoice-number\s*\n?\s*(\d+)", self.raw_text, re.IGNORECASE)
            if m:
                return m.group(1)

        if fmt == "wipelli":
            m = re.search(r"(\d{2,4}/\d{4})\s+\d{2}/", self.raw_text)
            if m:
                return m.group(1)
            m = re.search(r"n\.\s*(\d+/\d{4})", self.raw_text, re.IGNORECASE)
            if m:
                return m.group(1)
            m = re.search(r"^(\d{2,4}/\d{4})\s*$", self.raw_text, re.MULTILINE)
            if m:
                return m.group(1)

        if fmt == "continental":
            # OCR: "93112830 / 11.03.2026" or "11.03.2026 / 93112830"
            m = re.search(r"(\d{8,})\s*/\s*\d{2}\.\d{2}\.\d{4}", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "mastrotto":
            # "Anno/Numero/Data / YearNumber/date\n2024/1553000574 / 22.04.2024"
            m = re.search(r"(\d{4}/\d{10,})\s*/\s*\d{2}\.\d{2}\.\d{4}", self.raw_text)
            if m:
                return m.group(1)

        # Generic fallbacks
        patterns = [
            r"Invoice\s*#\s*(\d+)",
            r"Invoice\s*(?:no|number)[.:\s]*([A-Z0-9\-/]+)",
        ]
        for p in patterns:
            m = re.search(p, self.raw_text, re.IGNORECASE)
            if m:
                return m.group(1).strip()
        return None

    # =========================================================================
    # Date
    # =========================================================================
    def _extract_date(self) -> Optional[str]:
        fmt = self._fmt

        if fmt == "alcantara":
            m = re.search(r"\d{10}\s*-\s*(\d{2}/\d{2}/\d{4})", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "mabo":
            m = re.search(r"Date:\s*\n?.*?\n?\s*(\d{1,2}\.\d{1,2}\.\d{4})", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "west_trading":
            m = re.search(r"(\d{2}-\d{2}-\d{4})", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "crest":
            m = re.search(r"Document Date.*?(\d{1,2}[./]\s*\w+\s+\d{4})", self.raw_text, re.DOTALL)
            if m:
                return m.group(1).strip()
            m = re.search(r"\b(\d{2}/\d{2}/\d{2})\b", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "continental":
            # "93112830 / 11.03.2026"
            m = re.search(r"\d{8,}\s*/\s*(\d{2}\.\d{2}\.\d{4})", self.raw_text)
            if m:
                return m.group(1)

        if fmt == "mastrotto":
            # "2024/1553000574 / 22.04.2024"
            m = re.search(r"\d{4}/\d{10,}\s*/\s*(\d{2}\.\d{2}\.\d{4})", self.raw_text)
            if m:
                return m.group(1)

        # Generic date patterns (with word boundaries to avoid matching bank account numbers)
        patterns = [
            r"\b(\d{1,2}/\d{1,2}/\d{4})\b",
            r"\b(\d{1,2}/\d{1,2}/\d{2})\b",
            r"\b(\d{1,2}\.\d{1,2}\.\d{4})\b",
            r"\b(\d{4}[-/]\d{1,2}[-/]\d{1,2})\b",
        ]
        for p in patterns:
            m = re.search(p, self.raw_text)
            if m:
                val = m.group(1)
                parts = re.split(r"[/.\-]", val)
                if len(parts) == 3:
                    try:
                        nums = [int(x) for x in parts]
                        if nums[0] > 1000:
                            if nums[1] <= 12 and nums[2] <= 31:
                                return val
                        else:
                            if nums[1] <= 12:
                                return val
                    except ValueError:
                        pass
        return None

    # =========================================================================
    # Supplier Name
    # =========================================================================
    def _extract_supplier_name(self) -> Optional[str]:
        text = self.raw_text.lower()
        known = [
            ("alcantara", "s.p.a", "Alcantara S.p.A."),
            ("deca global", None, "DECA GLOBAL S.R.L."),
            ("crest", "leather", "Crest JMT Leather Ltd."),
            ("wipelli", None, "WIPELLI INTERNATIONAL SRL"),
            ("westtrading", None, "West Trading"),
            ("west trading", None, "West Trading"),
            ("mabo", None, "MABO International GmbH"),
            ("higher dimension", None, "Higher Dimension Materials, Inc."),
            ("hornschuch", None, "Konrad Hornschuch GmbH"),
            ("continental", None, "Konrad Hornschuch GmbH"),
            ("mastrotto", None, "GRUPPO MASTROTTO SPA"),
        ]
        for kw1, kw2, name in known:
            if kw1 in text and (kw2 is None or kw2 in text):
                return name
        return None

    # =========================================================================
    # Buyer Name
    # =========================================================================
    def _extract_buyer_name(self) -> Optional[str]:
        m = re.search(r"(CAMARI\s+(?:TRADING|INTERNATIONAL)[^\n]*)", self.raw_text, re.IGNORECASE)
        return m.group(1).strip() if m else None

    # =========================================================================
    # Trade Term
    # =========================================================================
    def _extract_trade_term(self) -> Optional[str]:
        _TERMS = r"FCA\S*|CIF|FOB|CFR|CPT|EXW|DAP|DDP"

        m = re.search(rf"INCOTERMS:\s*({_TERMS})", self.raw_text)
        if m:
            return m.group(1)

        m = re.search(rf"Inco\s*Terms?\s+({_TERMS})", self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1)
        # Mastrotto OCR: "FCA Incoterms® 2020"
        m = re.search(rf"({_TERMS})\s+Incoterms", self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1)

        # DECA: "Incoterms" column header, value within next few lines
        m = re.search(r"Incoterms\b", self.raw_text, re.IGNORECASE)
        if m:
            after = self.raw_text[m.end():m.end() + 300]
            tm = re.search(rf"\b({_TERMS})\b", after)
            if tm:
                return tm.group(1)

        # Alcantara pypdf: "DAP2010INCOTERMS:" or "FCA2010INCOTERMS:"
        m = re.search(rf"({_TERMS})(?:\d{{4}})INCOTERMS:", self.raw_text)
        if m:
            return m.group(1)

        # "Delivery conditions FCA Weifbach (Incoterms 2020)" or "Delivery: DAP (Incoterms 2020)"
        m = re.search(rf"Delivery\s+(?:conditions|terms)?\s*:?\s*({_TERMS})(?:\s+\w+)?(?:\s*\(Incoterms[^)]*\))?", self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1).upper()
        m = re.search(rf"[Dd]elivery\s*:?\s*({_TERMS})", self.raw_text)
        if m:
            return m.group(1).upper()

        # Wipelli / Mabo: "DELIVERY: EX WORKS" or "EX W ex works"
        if re.search(r"EX\s*WORKS|EX\s*W\b", self.raw_text, re.IGNORECASE):
            return "EXW"

        return None

    # =========================================================================
    # Country of Origin
    # =========================================================================
    def _extract_country_of_origin(self) -> Optional[str]:
        m = re.search(r"([A-Z]{2})\s+(\w+)\s*-\s*OR", self.raw_text)
        if m:
            return m.group(2)
        m = re.search(r"COO\s+(\w+)", self.raw_text)
        if m:
            return m.group(1)
        m = re.search(r"OR:\s*(\w+)", self.raw_text)
        if m:
            return m.group(1)
        m = re.search(r"Country of Origin.*?(?:is|:)\s*(.+?)(?:\n|$)", self.raw_text, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if "Germany" in val:
                return "Germany"
            if "Ital" in val:
                return "Italy"
            return val
        # Mastrotto: Italian company, address includes "ITALIA"
        if self._fmt == "mastrotto":
            return "Italy"
        return None

    # =========================================================================
    # Items
    # =========================================================================
    def _extract_items(self) -> list:
        fmt = self._fmt

        if fmt == "alcantara":
            items = self._parse_alcantara_items()
            if items:
                return items

        if fmt == "deca":
            items = self._parse_deca_items()
            if items:
                return items

        if fmt == "crest":
            return self._parse_crest_items()

        if fmt == "hdm":
            return self._parse_hdm_items()

        if fmt == "mabo":
            return self._parse_mabo_items()

        if fmt == "west_trading":
            return self._parse_west_trading_items()

        if fmt == "wipelli":
            return self._parse_wipelli_items()

        if fmt == "continental":
            return self._parse_continental_items()

        if fmt == "mastrotto":
            return self._parse_mastrotto_items()

        return []

    @staticmethod
    def _aggregate_customs_items(items: list) -> list:
        """Aggregate line items for customs declaration.

        Grouping key (in priority order):
        - Wipelli: article_name  (NAPPA SOFT, ADRIA, …)
        - Alcantara / DECA: item_code_prefix  (first 4 digits of item code)
        - Others: no aggregation (each item kept as-is)

        Aggregated row: quantity = sum, amount = sum,
        unit_price = total_amount / total_quantity.
        """
        if not items:
            return items

        has_group_key = any(
            item.get("article_name") or item.get("item_code_prefix")
            for item in items
        )
        if not has_group_key:
            return items

        from collections import OrderedDict
        groups: OrderedDict = OrderedDict()
        solo_idx = 0

        for item in items:
            article = item.get("article_name", "")
            prefix = item.get("item_code_prefix", "")

            if article:
                key = f"art:{article}"
            elif prefix:
                key = f"pfx:{prefix}"
            else:
                key = f"solo:{solo_idx}"
                solo_idx += 1

            if key not in groups:
                groups[key] = {
                    "article_name": article,
                    "item_code_prefix": prefix,
                    "unit": item.get("unit", ""),
                    "hide_type": item.get("hide_type", ""),
                    "country_code": item.get("country_code", ""),
                    "first_desc": item.get("description", ""),
                    "total_qty": 0.0,
                    "total_amount": 0.0,
                }
            try:
                groups[key]["total_qty"] += float(item.get("quantity", 0))
            except (ValueError, TypeError):
                pass
            try:
                groups[key]["total_amount"] += float(item.get("amount", 0))
            except (ValueError, TypeError):
                pass

        result = []
        for key, grp in groups.items():
            qty = round(grp["total_qty"], 2)
            amt = round(grp["total_amount"], 2)
            up = round(amt / qty, 2) if qty > 0 else 0

            agg: dict = {
                "description": grp["first_desc"],
                "quantity": str(qty),
                "unit": grp["unit"],
                "unit_price": str(up),
                "amount": str(amt),
            }
            if grp["article_name"]:
                agg["article_name"] = grp["article_name"]
                agg["description"] = f"LEATHER ART. {grp['article_name']}"
            if grp["item_code_prefix"]:
                agg["item_code_prefix"] = grp["item_code_prefix"]
            if grp["hide_type"]:
                agg["hide_type"] = grp["hide_type"]
            if grp["country_code"]:
                agg["country_code"] = grp["country_code"]

            result.append(agg)
        return result

    def _parse_alcantara_items(self) -> list:
        alcantara_pattern = (
            r"^(\d{8,10}[A-Z]{0,3})\s+"  # item_code (8-10 digits + optional letters)
            r"(.+?)\s{2,}"                # description
            r"(\d+,\d{2})\s+"             # quantity
            r"(\d{2})"                     # VAT code
            r"([\d.]+,\d{2})"             # amount
            r"(\d+,\d{2})"                # unit_price
            r"([A-Z])"                     # unit
            r"([A-Z]{2})\s+"              # country
            r"([A-Z]\d)\s+"               # TD code
            r"(\d+)"                       # composition code
        )
        matches = re.findall(alcantara_pattern, self.raw_text, re.MULTILINE)
        items = []
        for match in matches:
            items.append({
                "item_code": match[0],
                "item_code_prefix": match[0][:4],
                "description": match[1].strip(),
                "quantity": match[2].replace(',', '.'),
                "unit": "米" if match[6] == "M" else match[6],
                "unit_price": match[5].replace(',', '.'),
                "amount": match[4].replace('.', '').replace(',', '.'),
                "country_code": match[7],
                "td_code": match[8],
                "composition_code": match[9],
            })
        return items

    def _parse_deca_items(self) -> list:
        """DECA GLOBAL: line_no+item_code glued, then description, roll numbers,
        then 'qty M €price tax% €amount' on one line."""
        lines = self.raw_text.split("\n")
        items = []
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            m = re.match(r"^(\d{1,2})(\d{8,10}[A-Z]{0,3})$", line)
            if not m:
                i += 1
                continue

            line_no = m.group(1)
            item_code = m.group(2)
            desc = lines[i + 1].strip() if i + 1 < len(lines) else ""

            j = i + 2
            quantity = unit_price = amount = None
            while j < len(lines):
                data_line = lines[j].strip()
                dm = re.match(
                    r"^([\d.]+)\s+M\s+"
                    r"€([\d,]+)\s+"
                    r"\d+%\s+"
                    r"€(.+)$",
                    data_line,
                )
                if dm:
                    quantity = dm.group(1)
                    unit_price = _parse_euro_amount(dm.group(2))
                    amount = _parse_euro_amount(dm.group(3))
                    break
                j += 1

            if quantity is None:
                i += 1
                continue

            prefix_m = re.match(r"^(\d{4})", item_code)
            items.append({
                "line_no": line_no,
                "item_code": item_code,
                "item_code_prefix": prefix_m.group(1) if prefix_m else "",
                "description": desc,
                "quantity": quantity,
                "unit": "米",
                "unit_price": unit_price,
                "amount": amount,
            })
            i = j + 1
        return items

    def _parse_crest_items(self) -> list:
        """Crest: pypdf gives single-line items like:
        'LEGACY-DARK CARAMEL 9 500.75 Sq. Ft. 2.79 VAT20 1,397.09'"""
        items = []
        for m in re.finditer(
            r"([A-Z][A-Z\s\-]+?)\s+(\d+)\s+([\d,.]+)\s+Sq\.\s*Ft\.\s+([\d,.]+)\s+VAT\d+\s+([\d,.]+)",
            self.raw_text,
        ):
            desc, pieces, qty, price, amount = m.group(1), m.group(2), m.group(3), m.group(4), m.group(5)
            items.append({
                "line_no": str(len(items) + 1),
                "description": desc.strip(),
                "quantity": qty.replace(",", ""),
                "unit": "平方英尺",
                "unit_price": price.replace(",", ""),
                "amount": amount.replace(",", ""),
            })
        return items

    def _parse_hdm_items(self) -> list:
        """Higher Dimension Materials: columnar format where headers appear first,
        then all values for each column in sequence."""
        lines = self.raw_text.split("\n")
        stripped = [l.strip() for l in lines]

        try:
            qty_idx = stripped.index("Quantity")
        except ValueError:
            return []

        quantities = []
        j = qty_idx + 1
        while j < len(stripped) and re.match(r"^[\d.]+$", stripped[j]):
            quantities.append(stripped[j])
            j += 1

        n = len(quantities)
        if n == 0:
            return []

        try:
            units_idx = stripped.index("Units")
        except ValueError:
            units_idx = j
        units = []
        j = units_idx + 1
        while j < len(stripped) and len(units) < n:
            units.append(stripped[j])
            j += 1

        try:
            rate_idx = stripped.index("Rate")
        except ValueError:
            rate_idx = j
        rates = []
        j = rate_idx + 1
        while j < len(stripped) and re.match(r"^[\d,.]+$", stripped[j]):
            rates.append(stripped[j])
            j += 1

        try:
            amt_idx = next(k for k in range(rate_idx + 1, len(stripped)) if stripped[k] == "Amount")
        except StopIteration:
            amt_idx = j
        amounts = []
        j = amt_idx + 1
        while j < len(stripped) and re.match(r"^[\d,.]+$", stripped[j]):
            amounts.append(stripped[j])
            j += 1

        # Item descriptions: look for "Item" header followed by description lines
        descs: List[str] = []
        try:
            item_hdr = stripped.index("Item")
        except ValueError:
            item_hdr = -1
        if item_hdr >= 0:
            j = item_hdr + 1
            desc_buf = ""
            collected = 0
            while j < len(stripped) and collected < n:
                if stripped[j] in ("Description", "Comments", ""):
                    j += 1
                    continue
                # Multi-line descriptions: accumulate until we hit next item or keyword
                if desc_buf and (stripped[j].startswith("Wire") or re.match(r"^\d{2}/", stripped[j])):
                    if desc_buf:
                        descs.append(desc_buf.strip())
                        collected += 1
                    desc_buf = stripped[j]
                elif desc_buf:
                    desc_buf += " " + stripped[j]
                else:
                    desc_buf = stripped[j]
                j += 1
            if desc_buf:
                descs.append(desc_buf.strip())

        items = []
        for idx in range(n):
            qty = quantities[idx] if idx < len(quantities) else ""
            unit = units[idx] if idx < len(units) else ""
            rate = rates[idx] if idx < len(rates) else ""
            amt = amounts[idx] if idx < len(amounts) else ""
            desc = descs[idx] if idx < len(descs) else ""

            # Skip zero-amount/free items and wire fees
            try:
                amt_f = float(amt.replace(",", ""))
            except (ValueError, TypeError):
                amt_f = 0
            if amt_f <= 0 and "wire" not in desc.lower():
                continue

            unit_cn = "码" if unit.lower() in ("yd", "yard", "yards") else unit
            items.append({
                "line_no": str(len(items) + 1),
                "description": desc,
                "quantity": qty,
                "unit": unit_cn,
                "unit_price": rate,
                "amount": amt.replace(",", ""),
            })
        return items

    def _parse_mabo_items(self) -> list:
        """MABO: pypdf concatenates item lines like:
        '1 148 m Alcantara 2627 pannelMicrofibre for Automotive Custom Statistic Nr. 56031480'
        followed by price/amount on next line: '35,00 5.180,00'"""
        items = []
        lines = self.raw_text.split("\n")
        for i, line in enumerate(lines):
            m = re.match(
                r"^(\d{1,2})\s+([\d,.]+)\s+m\s+(.+?)(?:Microfibre|Custom Statistic|$)",
                line,
            )
            if not m:
                continue
            line_no = m.group(1)
            qty = m.group(2).replace(",", ".")
            desc = m.group(3).strip()

            price = amount = ""
            if i + 1 < len(lines):
                pm = re.match(r"^\s*([\d.,]+)\s+([\d.,]+)\s*$", lines[i + 1])
                if pm:
                    price = _parse_euro_amount(pm.group(1))
                    amount = _parse_euro_amount(pm.group(2))

            items.append({
                "line_no": line_no,
                "description": desc,
                "quantity": qty,
                "unit": "米",
                "unit_price": price,
                "amount": amount,
            })
        return items

    def _parse_west_trading_items(self) -> list:
        """West Trading: pypdf gives:
        'PORS18118 FLATWOVEN WITH SQUARES STUDIO CHECK'
        '62% wool / 38% pes'
        '10,00 80,00 800,00mtr.'
        """
        items = []
        lines = self.raw_text.split("\n")
        for i, line in enumerate(lines):
            m = re.match(r"^([A-Z]{2,}\d{3,})\s+(.+)$", line)
            if not m:
                continue
            item_code = m.group(1)
            desc = m.group(2).strip()

            composition = ""
            if i + 1 < len(lines) and "%" in lines[i + 1]:
                composition = lines[i + 1].strip()

            data_line_idx = i + 2 if composition else i + 1
            qty = price = amount = ""
            if data_line_idx < len(lines):
                dm = re.match(r"^\s*([\d,]+)\s+([\d,]+)\s+([\d.,]+)\s*(?:mtr\.|pcs)?", lines[data_line_idx])
                if dm:
                    qty = dm.group(1).replace(",", ".")
                    price = _parse_euro_amount(dm.group(2))
                    amount = _parse_euro_amount(dm.group(3))

            full_desc = f"{desc} ({composition})" if composition else desc
            items.append({
                "line_no": str(len(items) + 1),
                "item_code": item_code,
                "description": full_desc,
                "composition": composition,
                "quantity": qty,
                "unit": "米",
                "unit_price": price,
                "amount": amount,
            })
        return items

    def _parse_wipelli_items(self) -> list:
        """Wipelli: LEATHER ART descriptions with (N HIDES), then data lines.

        pypdf may produce two formats:
        A) Inline: "365,5Mq  29,5000  10.782,25" (qty+Mq glued, price, amount on one line)
        B) Columnar: separate "Mq" lines followed by consecutive numbers.
        We try inline first; fall back to columnar if no inline matches found.
        """
        desc_matches = re.findall(
            r"(LEATHER\s+ART\.\s+[^\(]+?)\s*\((\d+)\s*(HIDES|HALF HIDES)\)",
            self.raw_text,
        )
        if not desc_matches:
            return []

        n = len(desc_matches)
        lines = self.raw_text.split("\n")
        stripped = [l.strip() for l in lines]

        # --- Format A: inline "365,5Mq  29,5000  10.782,25" ---
        inline_data: List[dict] = []
        for line in stripped:
            m = re.match(r"([\d.,]+)\s*Mq\s+([\d.,]+)\s+([\d.,]+)", line)
            if m:
                inline_data.append({
                    "qty": m.group(1).replace(",", "."),
                    "price": _parse_euro_amount(m.group(2)),
                    "amount": _parse_euro_amount(m.group(3)),
                })

        if len(inline_data) == n:
            items = []
            for idx, (desc_raw, hide_count, hide_type) in enumerate(desc_matches):
                art_m = re.search(r"ART\.\s+(.+?)\s+COL\.", desc_raw)
                article_name = art_m.group(1).strip() if art_m else None

                desc = f"{desc_raw.strip()} ({hide_count} {hide_type})"
                d = inline_data[idx]

                item = {
                    "line_no": str(idx + 1),
                    "description": desc,
                    "hide_count": hide_count,
                    "hide_type": hide_type,
                    "quantity": d["qty"],
                    "unit": "平方米",
                    "unit_price": d["price"],
                    "amount": d["amount"],
                }
                if article_name:
                    item["article_name"] = article_name
                items.append(item)
            return items

        # --- Format B: columnar (Mq on own line, numbers below) ---
        mq_values: List[str] = []
        prices: List[str] = []

        mq_indices = [k for k, s in enumerate(stripped) if s == "Mq"]
        if mq_indices:
            last_mq = mq_indices[-1]
            j = last_mq + 1
            all_nums: List[str] = []
            while j < len(stripped):
                s = stripped[j]
                if re.match(r"^[\d.,]+$", s):
                    all_nums.append(s)
                    j += 1
                elif re.match(r"^[\d.,]+\s+[\d.,]+$", s):
                    parts = s.split()
                    all_nums.extend(parts)
                    j += 1
                else:
                    break
            mq_values = [v.replace(",", ".") for v in all_nums[:n]]
            prices = [_parse_euro_amount(v) for v in all_nums[n:2 * n]]

        items = []
        for idx, (desc_raw, hide_count, hide_type) in enumerate(desc_matches):
            art_m = re.search(r"ART\.\s+(.+?)\s+COL\.", desc_raw)
            article_name = art_m.group(1).strip() if art_m else None

            desc = f"{desc_raw.strip()} ({hide_count} {hide_type})"
            qty_sqm = mq_values[idx] if idx < len(mq_values) else ""
            unit_price = prices[idx] if idx < len(prices) else ""

            amount = ""
            if qty_sqm and unit_price:
                try:
                    amount = str(round(float(qty_sqm) * float(unit_price), 2))
                except (ValueError, TypeError):
                    pass

            item = {
                "line_no": str(idx + 1),
                "description": desc,
                "hide_count": hide_count,
                "hide_type": hide_type,
                "quantity": qty_sqm or hide_count,
                "unit": "平方米" if qty_sqm else "张",
                "unit_price": unit_price,
                "amount": amount,
            }
            if article_name:
                item["article_name"] = article_name
            items.append(item)
        return items

    def _parse_continental_items(self) -> list:
        """Continental/Hornschuch OCR text:
        '000010*F6473028 skai Torino FLS black'
        '150,000 M 21,65 EUR 1M 3.247,50'
        'surcharge p. unit 15,000 % 487,13'
        '24,90 EUR 1M 3.734,63'
        """
        items = []
        lines = self.raw_text.split("\n")
        for i, line in enumerate(lines):
            # Match position line: "000010*F6473028 skai Torino FLS black"
            m = re.match(r"^0*(\d+)\*?([A-Z0-9]+)\s+(.+)$", line)
            if not m:
                continue
            pos = m.group(1)
            material = m.group(2)
            desc = m.group(3).strip()

            # Scan forward for "150,000 M 21,65 EUR 1M 3.247,50" or "24,90 EUR 1M 3.734,63"
            qty = unit_price = amount = ""
            unit = "米"
            for j in range(i + 1, min(i + 15, len(lines))):
                # Final price line with surcharge: "24,90 EUR 1M 3.734,63"
                fm = re.match(r"^\s*([\d,.]+)\s+EUR\s+1M\s+([\d.,]+)\s*$", lines[j])
                if fm:
                    unit_price = _parse_euro_amount(fm.group(1))
                    amount = _parse_euro_amount(fm.group(2))
                # Base qty line: "150,000 M 21,65 EUR 1M 3.247,50"
                qm = re.match(r"^\s*([\d,.]+)\s+M\s+[\d,.]+\s+EUR\s+1M\s+[\d.,]+\s*$", lines[j])
                if qm:
                    qty = qm.group(1).replace(",", ".").rstrip("0").rstrip(".")
                # "Sum of positions" signals end of this item
                if "Sum of positions" in lines[j]:
                    if not amount:
                        sm = re.search(r"([\d.,]+)\s*$", lines[j])
                        if sm:
                            amount = _parse_euro_amount(sm.group(1))
                    break

            if qty:
                items.append({
                    "line_no": pos,
                    "item_code": material,
                    "description": desc,
                    "quantity": qty,
                    "unit": unit,
                    "unit_price": unit_price,
                    "amount": amount,
                })
        return items

    def _parse_mastrotto_items(self) -> list:
        """Mastrotto OCR text:
        '0010 2508940'
        'COLORE 1621870'
        '463,560 M2'
        'LEMANS 11-13 4210 ROSSO FERRAR'
        'EUR 14.277, 65 YA'
        """
        items = []
        lines = self.raw_text.split("\n")
        for i, line in enumerate(lines):
            # Position + material code: "0010 2508940"
            m = re.match(r"^0*(\d+)\s+(\d{5,})\s*$", line.strip())
            if not m:
                continue
            pos = m.group(1)
            material = m.group(2)

            # Collect description, quantity, and amount from surrounding lines
            desc = ""
            qty = ""
            amount = ""
            unit = "平方米"

            for j in range(i + 1, min(i + 8, len(lines))):
                s = lines[j].strip()
                # Quantity: "463,560 M2"
                qm = re.match(r"^([\d,.]+)\s*M2\b", s)
                if qm:
                    qty = qm.group(1).replace(",", ".")
                    continue
                # Amount: "EUR 14.277, 65" or "EUR 14.277,65"
                am = re.search(r"EUR\s+([\d., ]+)", s)
                if am:
                    amount = _parse_euro_amount(am.group(1).replace(" ", ""))
                    continue
                # Description: "LEMANS 11-13 4210 ROSSO FERRAR" (all caps, not a header)
                if re.match(r"^[A-Z][A-Z0-9\s\-]+$", s) and "COLORE" not in s and len(s) > 5:
                    desc = s

            if qty or amount:
                unit_price = ""
                if qty and amount:
                    try:
                        unit_price = str(round(float(amount) / float(qty), 2))
                    except (ValueError, ZeroDivisionError):
                        pass
                items.append({
                    "line_no": pos,
                    "item_code": material,
                    "description": desc,
                    "quantity": qty,
                    "unit": unit,
                    "unit_price": unit_price,
                    "amount": amount,
                })
        return items

    # =========================================================================
    # Compute Total
    # =========================================================================
    def _compute_total(self, items: list) -> Optional[str]:
        if items:
            total = 0
            for item in items:
                try:
                    total += float(item.get("amount", 0))
                except (ValueError, TypeError):
                    pass
            if total > 0:
                return f"{total:.2f}"

        # Alcantara: "8/A 14.711,20 0,00"
        m = re.search(r"8/A\s+([\d.]+,\d{2})\s+0,00", self.raw_text)
        if m:
            return m.group(1).replace('.', '').replace(',', '.')

        # "Total €33\xa0866,72" or "Total\n€33 866,72"
        m = re.search(r"Total\s*[\n\s]*€([\d\s\xa0,.]+)", self.raw_text, re.IGNORECASE)
        if m:
            return _parse_euro_amount(m.group(1))

        # "Total USD 1,415.19" or "Total USD\n1,415.19"
        m = re.search(r"Total\s+(?:USD|EUR|GBP)\s*\n?\s*([\d,]+\.?\d*)", self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1).replace(',', '')

        # West Trading: "Amount without tax : 1.600,00" or "1.600,00Total amount invoice :"
        m = re.search(r"Amount without tax\s*:\s*([\d.,]+)", self.raw_text, re.IGNORECASE)
        if m:
            return _parse_euro_amount(m.group(1))
        m = re.search(r"([\d.,]+)\s*Total amount invoice", self.raw_text, re.IGNORECASE)
        if m:
            return _parse_euro_amount(m.group(1))

        # "$1,225.00"
        m = re.search(r"\$\s*([\d,]+\.\d{2})", self.raw_text)
        if m:
            return m.group(1).replace(',', '')

        # "Total amount invoice :\n1.600,00" (West Trading)
        m = re.search(r"Total\s+amount\s+invoice\s*:?\s*\n?\s*([\d.,]+)", self.raw_text, re.IGNORECASE)
        if m:
            return _parse_euro_amount(m.group(1))

        # Continental: "Total amount EUR 3.734,63"
        m = re.search(r"Total\s+amount\s+EUR\s+([\d.,]+)", self.raw_text, re.IGNORECASE)
        if m:
            return _parse_euro_amount(m.group(1))

        # Mastrotto: "Totale fattura/Total invoice EUR 14.277,65"
        m = re.search(r"Total\s+invoice\s+EUR\s+([\d.,\s]+)", self.raw_text, re.IGNORECASE)
        if m:
            return _parse_euro_amount(m.group(1))

        # Mabo: "Total*\n10.419,50"
        m = re.search(r"Total\*?\s*\n\s*([\d.,]+)", self.raw_text, re.IGNORECASE)
        if m:
            return _parse_euro_amount(m.group(1))

        return None

    # =========================================================================
    # Currency
    # =========================================================================
    def _extract_currency(self) -> Optional[str]:
        text = self.raw_text
        if "EUR" in text or "€" in text or "Eur" in text:
            return "EUR"
        if "USD" in text or "$" in text:
            return "USD"
        if "GBP" in text or "£" in text:
            return "GBP"
        return None

    # =========================================================================
    # Weight
    # =========================================================================
    def _extract_weight(self, weight_type: str) -> Optional[str]:
        m = re.search(rf"{weight_type}-WEIGHT:\s*([\d.,]+)", self.raw_text)
        if m:
            return m.group(1).replace(',', '.')
        m = re.search(rf"{weight_type}\s*[Ww]eight:?\s*([\d.,]+)\s*(?:KG|kg)?", self.raw_text)
        if m:
            v = m.group(1).replace(',', '.')
            return v.rstrip("0").rstrip(".") if "." in v else v

        # Alcantara pypdf: value before label, e.g. " 289,00PESO NETTO:" / " 353,00PESO LORDO:"
        label_map = {"NET": "PESO NETTO", "GROSS": "PESO LORDO"}
        label = label_map.get(weight_type.upper())
        if label:
            m = re.search(rf"([\d.,]+){label}", self.raw_text)
            if m:
                return _parse_euro_amount(m.group(1))

        def _clean_weight(val: str) -> str:
            v = val.replace(',', '.')
            return v.rstrip("0").rstrip(".") if "." in v else v

        if weight_type.upper() == "GROSS":
            m = re.search(r"GROSS\s+WEIGHT:?\s*\n?\s*(?:KG\s*)?([\d.,]+)", self.raw_text, re.IGNORECASE)
            if m:
                return _clean_weight(m.group(1))
            m = re.search(r"Gross\s+weight:?\s*([\d.,]+)", self.raw_text, re.IGNORECASE)
            if m:
                return _clean_weight(m.group(1))
            m = re.search(r"gross:\s*([\d.,]+)\s*kg", self.raw_text, re.IGNORECASE)
            if m:
                return _clean_weight(m.group(1))
        if weight_type.upper() == "NET":
            m = re.search(r"NET\s+WEIGHT:?\s*\n?\s*(?:KG\s*)?([\d.,]+)", self.raw_text, re.IGNORECASE)
            if m:
                return _clean_weight(m.group(1))
            m = re.search(r"Net\s+weight:?\s*([\d.,]+)", self.raw_text, re.IGNORECASE)
            if m:
                return _clean_weight(m.group(1))
            m = re.search(r"net:\s*([\d.,]+)\s*kg", self.raw_text, re.IGNORECASE)
            if m:
                return _clean_weight(m.group(1))
        return None

    # =========================================================================
    # Payment Condition
    # =========================================================================
    def _extract_payment_cond(self) -> Optional[str]:
        # Alcantara: "PAYMENT COND: 60 DAYS NET  INCOTERMS..."
        m = re.search(r"PAYMENT\s+COND:\s*(.+?)(?:\s+INCOTERMS|\n)", self.raw_text)
        if m:
            return m.group(1).strip()
        # Alcantara pypdf: "90 gg Data FatturaCONDIZIONI DI PAGAMENTO:"
        m = re.search(r"(.+?)CONDIZIONI DI PAGAMENTO:", self.raw_text)
        if m:
            val = m.group(1).strip()
            # The value may start after a number (weight), take only the payment part
            pm = re.search(r"([\d]+\s*gg\s*.+|Anticipato)", val, re.IGNORECASE)
            if pm:
                return pm.group(1).strip()

        # DECA: "100% TT IN ADVANCE" (may span two lines)
        m = re.search(r"(\d+%\s*TT\s+IN\s*\n?\s*ADVANCE)", self.raw_text, re.IGNORECASE)
        if m:
            return re.sub(r"\s+", " ", m.group(1)).strip()

        # Wipelli: "30 DAYS INVOICE DATE"
        m = re.search(r"(\d+\s+DAYS\s+INVOICE\s+DATE)", self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

        # HDM: "Terms\nPrepaid"
        m = re.search(r"Terms\s*\n\s*(Prepaid|Net\s+\d+)", self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

        # Mabo: "Netto payable until dd.mm.yyyy" or "Net 30 days"
        m = re.search(r"(Netto\s+payable\s+until\s+[\d.]+)", self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

        # West Trading: "Terms of payment:: Paid in advance..."
        m = re.search(r"Terms\s+of\s+payment:+\s*\n?\s*(.+?)(?:\n|$)", self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

        # Continental: "Terms of payment in advance"
        m = re.search(r"Terms\s+of\s+payment\s+(.+?)(?:\n|$)", self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

        # Mastrotto: "Scadenze pagamento /Due date for payment\n22.05.2024"
        m = re.search(r"[Dd]ue date.*?payment\s*\n\s*(\d{2}\.\d{2}\.\d{4})", self.raw_text)
        if m:
            return f"Due {m.group(1)}"

        # Generic fallbacks
        m = re.search(r"(Payment\s+Before\s+Delivery|Paid\s+in\s+advance|Net\s+\d+\s+days)",
                       self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        return None

    # =========================================================================
    # Packages
    # =========================================================================
    def _extract_packages(self) -> Optional[str]:
        m = re.search(r"PACKAGES\s+N\.:\s*(\d+)", self.raw_text)
        if m:
            return m.group(1)
        # Alcantara pypdf: "2N. COLLI:"
        m = re.search(r"(\d+)\s*N\.\s*COLLI", self.raw_text)
        if m:
            return m.group(1)
        m = re.search(r"(\d+)\s+(?:PALLETS?|PALETT[AE]|Packages|packages|PACKAGES|pallet)", self.raw_text, re.IGNORECASE)
        if m:
            return m.group(1)
        return None

    def get_parsed_data(self) -> Dict[str, Any]:
        if not self.parsed_data:
            self.parse()
        return self.parsed_data
