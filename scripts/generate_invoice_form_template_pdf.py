#!/usr/bin/env python3
"""
Generate an editable invoice PDF template (AcroForm) that can be edited in
macOS Preview.

This template is based on the layout of CustInvc_*.pdf used in your project,
but we generate the PDF from scratch so the resulting form does not contain
barcode/footer.

Editable fields:
- Bill To (multi-line)
- Ship To (multi-line)
- Invoice #, Invoice Date
- Table rows (up to max_rows, max 10)
- Totals: Subtotal, Tax Total, Total, Due Date
"""

from __future__ import annotations

import argparse
from pathlib import Path

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas


def _add_text_field(
    c: canvas.Canvas,
    name: str,
    x: float,
    y: float,
    width: float,
    height: float,
    *,
    multiline: bool = False,
    max_len: int = 200,
    font_size: int = 10,
) -> None:
    # reportlab uses bottom-left coordinates.
    c.acroForm.textfield(
        name=name,
        x=x,
        y=y,
        width=width,
        height=height,
        borderStyle="inset",
        borderColor=colors.black,
        fillColor=colors.white,
        textColor=colors.black,
        fontName="Helvetica",
        fontSize=font_size,
        forceBorder=True,
        # Field flags:
        # - 4096 corresponds to multiline in AcroForm.
        fieldFlags=4096 if multiline else 0,
        maxlen=max_len,
        value="",
    )


def generate_template(output_pdf: str, max_rows: int = 10) -> Path:
    if max_rows < 1 or max_rows > 20:
        raise ValueError("max_rows must be between 1 and 20.")

    out_path = Path(output_pdf).expanduser().resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    w, h = letter  # (612, 792) on US Letter
    c = canvas.Canvas(str(out_path), pagesize=(w, h))

    # --- Style ---
    left = 40
    right = w - 40
    top = h - 40

    c.setFont("Helvetica", 11)
    seller_lines = [
        "CAMARI TRADING (ZHEJIANG) CO.,LTD",
        "1525 Hexing Rd",
        "Jiaxing Zhejiang 314001",
        "China",
    ]
    y = top
    for ln in seller_lines:
        c.drawString(left, y, ln)
        y -= 14

    # --- Invoice meta ---
    c.setFont("Helvetica-Bold", 14)
    c.drawString(360, top - 10, "Invoice")

    # The original screenshot shows Invoice # and Date very large on the right
    _add_text_field(
        c,
        "invoice_no",
        360,
        top - 50,
        width=200,
        height=30,
        multiline=False,
        max_len=40,
        font_size=24,
    )

    _add_text_field(
        c,
        "invoice_date",
        360,
        top - 84,
        width=200,
        height=24,
        multiline=False,
        max_len=20,
        font_size=18,
    )

    # --- Bill To / Ship To ---
    c.setFont("Helvetica-Bold", 10.5)
    bill_x = left
    ship_x = 280
    box_top = top - 110
    box_h = 70
    box_w = 230

    c.drawString(bill_x, box_top + 15, "Bill To")
    c.drawString(ship_x, box_top + 15, "Ship To")

    # The original doesn't have black borders for Bill To / Ship To
    _add_text_field(
        c,
        "bill_to",
        bill_x,
        box_top,
        width=box_w,
        height=box_h,
        multiline=True,
        max_len=500,
        font_size=10,
    )
    _add_text_field(
        c,
        "ship_to",
        ship_x,
        box_top,
        width=box_w,
        height=box_h,
        multiline=True,
        max_len=500,
        font_size=10,
    )

    # --- Memo ---
    memo_y = box_top - 30
    c.setFont("Helvetica-Bold", 11)
    c.drawString(left, memo_y + 4, "Memo:")
    _add_text_field(
        c,
        "memo",
        left + 50,
        memo_y,
        width=400,
        height=16,
        multiline=False,
        max_len=100,
        font_size=10,
    )

    # --- Meta Bar (Grey background) ---
    bar_y = memo_y - 30
    bar_h = 36
    c.setFillColorRGB(0.92, 0.92, 0.92)  # Light grey
    c.setStrokeColor(colors.white)
    c.rect(left, bar_y - bar_h, right - left, bar_h, stroke=0, fill=1)
    
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 9)
    meta_cols = [
        ("Terms", 70),
        ("Due Date", 70),
        ("PO #", 60),
        ("Currency", 60),
        ("Incoterms", 70),
        ("Shipping Method", 100),
        ("P.I", 80),
    ]
    
    mx = left + 4
    for label, w in meta_cols:
        c.drawString(mx, bar_y - 12, label)
        mx += w
        
    # Meta Bar fields
    mx = left + 4
    for i, (label, w) in enumerate(meta_cols):
        _add_text_field(
            c,
            f"meta_{i}",
            mx,
            bar_y - bar_h + 4,
            width=w - 8,
            height=20,
            multiline=False,
            max_len=50,
            font_size=9,
        )
        mx += w

    # --- Table header ---
    table_top = bar_y - bar_h - 20  # start below meta bar
    c.setFont("Helvetica", 10)
    col_defs = [
        ("#", 40),
        ("Item", 160),
        ("Quantity", 75),
        ("Units", 55),
        ("Unit Price", 95),
        ("Tax Rate", 65),
        ("Amount", 95),
    ]

    table_left = left
    table_width = right - left

    # Normalize column widths to fit table_width
    total_w = sum(wi for _, wi in col_defs)
    scale = table_width / total_w
    col_defs = [(lbl, wi * scale) for lbl, wi in col_defs]

    # Compute x positions
    x_positions = [table_left]
    for _, cw in col_defs[:-1]:
        x_positions.append(x_positions[-1] + cw)

    header_h = 18
    row_h = 20

    header_y = table_top - header_h
    c.setStrokeColor(colors.black)
    c.setFillColor(colors.white)

    # Header cells outlines + labels
    for i, (lbl, cw) in enumerate(col_defs):
        x = x_positions[i]
        c.rect(x, header_y, cw, header_h, stroke=1, fill=0)
        c.drawCentredString(x + cw / 2, header_y + 5, lbl)

    # Rows (max_rows)
    for r in range(max_rows):
        y0 = header_y - (r + 1) * row_h
        for i, (lbl, cw) in enumerate(col_defs):
            x = x_positions[i]
            c.rect(x, y0, cw, row_h, stroke=1, fill=0)

        # Column field placements:
        # columns: 0..6 (# Item Quantity Units UnitPrice TaxRate Amount)
        _add_text_field(
            c,
            f"item_{r+1}_no",
            x_positions[0] + 3,
            y0 + 4,
            width=col_defs[0][1] - 6,
            height=row_h - 8,
            multiline=False,
            max_len=6,
            font_size=9,
        )
        _add_text_field(
            c,
            f"item_{r+1}_desc",
            x_positions[1] + 3,
            y0 + 4,
            width=col_defs[1][1] - 6,
            height=row_h - 8,
            multiline=False,
            max_len=60,
            font_size=9,
        )
        _add_text_field(
            c,
            f"item_{r+1}_qty",
            x_positions[2] + 3,
            y0 + 4,
            width=col_defs[2][1] - 6,
            height=row_h - 8,
            multiline=False,
            max_len=20,
            font_size=9,
        )
        _add_text_field(
            c,
            f"item_{r+1}_unit",
            x_positions[3] + 3,
            y0 + 4,
            width=col_defs[3][1] - 6,
            height=row_h - 8,
            multiline=False,
            max_len=20,
            font_size=9,
        )
        _add_text_field(
            c,
            f"item_{r+1}_unit_price",
            x_positions[4] + 3,
            y0 + 4,
            width=col_defs[4][1] - 6,
            height=row_h - 8,
            multiline=False,
            max_len=25,
            font_size=9,
        )
        _add_text_field(
            c,
            f"item_{r+1}_tax_rate",
            x_positions[5] + 3,
            y0 + 4,
            width=col_defs[5][1] - 6,
            height=row_h - 8,
            multiline=False,
            max_len=10,
            font_size=9,
        )
        _add_text_field(
            c,
            f"item_{r+1}_amount",
            x_positions[6] + 3,
            y0 + 4,
            width=col_defs[6][1] - 6,
            height=row_h - 8,
            multiline=False,
            max_len=25,
            font_size=9,
        )

    # --- Totals block (Right-aligned, large fonts) ---
    totals_top = header_y - max_rows * row_h - 20
    
    # Subtotal and Tax
    c.setFont("Helvetica", 10.5)
    tx = w - 240
    ty = totals_top
    
    c.drawString(tx, ty, "Subtotal")
    _add_text_field(c, "subtotal", tx + 70, ty - 4, width=130, height=18, max_len=25)
    
    ty -= 24
    c.drawString(tx, ty, "Tax Total")
    _add_text_field(c, "tax_total", tx + 70, ty - 4, width=130, height=18, max_len=25)
    
    # Grand Total (Grey background)
    ty -= 40
    c.setFillColorRGB(0.92, 0.92, 0.92)
    c.rect(tx - 20, ty - 10, 260, 45, stroke=0, fill=1)
    
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(tx, ty + 10, "Total")
    
    _add_text_field(
        c, "total_amount", tx + 70, ty + 6,
        width=130, height=24, font_size=18, max_len=25
    )
    
    # Due Date below Total
    ty -= 30
    c.setFont("Helvetica-Bold", 12)
    c.drawString(tx, ty, "Due Date:")
    _add_text_field(
        c, "totals_due_date", tx + 70, ty - 2,
        width=130, height=18, font_size=12, max_len=20
    )

    c.showPage()
    c.save()
    return out_path


def main() -> None:
    ap = argparse.ArgumentParser(description="Generate editable invoice template PDF (Preview-compatible).")
    ap.add_argument(
        "--output",
        default="templates/invoice_form_template.pdf",
        help="Output PDF path",
    )
    ap.add_argument("--max-rows", type=int, default=10, help="Max editable item rows")
    args = ap.parse_args()

    out = generate_template(args.output, max_rows=args.max_rows)
    print(f"Generated template: {out}")


if __name__ == "__main__":
    main()

