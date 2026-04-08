"""
生成与 CAMARI CustInvc 版式相近、可被 pdf_parser（camari_cust）解析的欧元发票 PDF。
使用 ReportLab；若系统存在 Arial Unicode / DejaVu 等则注册以支持中文/日文地址。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

_FONT_REGISTERED: Optional[str] = None


def _register_unicode_font() -> str:
    global _FONT_REGISTERED
    if _FONT_REGISTERED:
        return _FONT_REGISTERED
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/Library/Fonts/Arial Unicode.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "C:/Windows/Fonts/arialuni.ttf",
        "C:/Windows/Fonts/msyh.ttc",
    ]
    for p in candidates:
        path = Path(p)
        if not path.exists():
            continue
        try:
            name = "EurInvFont"
            pdfmetrics.registerFont(TTFont(name, str(path)))
            _FONT_REGISTERED = name
            return name
        except Exception:
            continue
    _FONT_REGISTERED = "Helvetica"
    return _FONT_REGISTERED


def _fmt_money_eur(s: Any) -> str:
    try:
        v = float(str(s).replace(",", "").strip())
    except (ValueError, TypeError):
        return str(s)
    if abs(v - round(v)) < 1e-6:
        return f"{int(round(v)):,}"
    return f"{v:,.2f}"


def _safe_line(s: str, font: str) -> str:
    if font != "Helvetica":
        return s
    return s.encode("latin-1", errors="replace").decode("latin-1")


def build_camari_eur_invoice_lines(invoice_data: Dict[str, Any]) -> List[str]:
    """构造与 CustInvc 解析器兼容的纯文本行（用于写入 PDF）。"""
    inv_no = str(invoice_data.get("invoice_no") or "").strip()
    inv_date = str(invoice_data.get("invoice_date") or "").strip()
    issuer = (invoice_data.get("issuer_name") or "CAMARI TRADING (ZHEJIANG) CO.,LTD").strip()
    iss_addr = (invoice_data.get("issuer_address") or "").strip()
    if not iss_addr:
        iss_addr = "1525 Hexing Rd\nJiaxing Zhejiang 314001\nChina"
    buyer = (invoice_data.get("buyer_name") or "BUYER").strip()
    buyer_addr = (invoice_data.get("buyer_address") or "").strip()
    trade = (invoice_data.get("trade_term") or "Dap").strip()
    pay = (invoice_data.get("payment_cond") or "100% TT IN ADVANCE").strip()
    total_amt = invoice_data.get("total_amount")
    items = invoice_data.get("items") or []

    lines: List[str] = []
    lines.append(issuer)
    for part in iss_addr.split("\n"):
        if part.strip():
            lines.append(part.strip())
    lines.append("Invoice")
    lines.append(f"#{inv_no}" if inv_no else "#0")
    lines.append(inv_date or "")
    lines.append(inv_no or "")
    lines.append("1 of 1")
    lines.append("Bill To Ship To TOTAL")
    lines.append(buyer)
    for part in buyer_addr.split("\n"):
        if part.strip():
            lines.append(part.strip())
    tot_disp = _fmt_money_eur(total_amt) if total_amt is not None else "0"
    lines.append(f"€{tot_disp}")
    lines.append(f"Due Date: {inv_date}" if inv_date else "Due Date:")
    lines.append("Memo:")
    lines.append("")
    lines.append("Terms Due Date PO # Currency Incoterms Shipping Method P.I")
    for part in pay.split("\n"):
        if part.strip():
            lines.append(part.strip())
    # 满足 camari 贸易条款与币种正则：日期 + EUR + 贸易术语（字母）
    lines.append(f"{inv_date} EUR {trade} {tot_disp}")
    lines.append("EUR Sales Order")
    lines.append("# Item Quantity Units Unit Price Tax Rate Amount")

    for it in items:
        ln = str(it.get("line_no") or "").strip()
        desc = (it.get("description") or "").strip()
        qty = str(it.get("quantity") or "").strip()
        unit = (it.get("unit") or "M").strip()
        up = _fmt_money_eur(it.get("unit_price"))
        amt = _fmt_money_eur(it.get("amount"))
        head = f"{ln}{desc}" if desc else ln
        lines.append(head)
        lines.append(f"{qty} {unit} €{up} 0% €{amt}")

    lines.append(" Subtotal€" + str(tot_disp))
    lines.append(" Tax Total€0")
    lines.append(" Total€" + str(tot_disp))
    lines.append("")
    lines.append("Bank Information:")
    lines.append("BENEFICIARY: (see original invoice for bank details)")
    return lines


def write_camari_cust_eur_invoice_pdf(
    output_path: str,
    invoice_data: Dict[str, Any],
) -> None:
    """将 invoice_data（已为 EUR）写成 PDF，文本可被 InvoiceParser(camari_cust) 解析。"""
    font = _register_unicode_font()
    lines = build_camari_eur_invoice_lines(invoice_data)
    c = canvas.Canvas(output_path, pagesize=A4)
    w, h = A4
    margin_x = 40
    y = h - 48
    line_h = 11
    c.setFont(font, 9)
    for line in lines:
        if y < 48:
            c.showPage()
            c.setFont(font, 9)
            y = h - 48
        txt = _safe_line(line[:2000], font)
        c.drawString(margin_x, y, txt)
        y -= line_h
    c.save()
