"""
在「原始 Invoice PDF」上仅替换金额与相关币种字样，保留版式与版面。

策略：
1. 用 page.get_text("dict") 遍历所有 span，找到包含待替换金额的 span
2. 记录其 font/size/color/bbox
3. 检测该位置的背景色（来自绘图矩形），redact 时用同色填充
4. 用系统 Unicode 字体（与原字号一致）在原位插入 EUR 金额

依赖 PyMuPDF (fitz)。
"""

from __future__ import annotations

import copy
import re
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None  # type: ignore


# ---------------------------------------------------------------------------
# Font: use system font that definitely contains € glyph
# ---------------------------------------------------------------------------

_SYSTEM_FONTS: Tuple[str, ...] = (
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    r"C:\Windows\Fonts\arial.ttf",
    r"C:\Windows\Fonts\arialuni.ttf",
)

_cached_font: Optional[str] = None
_font_searched: bool = False


def _system_euro_font() -> Optional[str]:
    global _cached_font, _font_searched
    if _font_searched:
        return _cached_font
    _font_searched = True
    for p in _SYSTEM_FONTS:
        if Path(p).is_file():
            _cached_font = p
            return p
    return None


# ---------------------------------------------------------------------------
# Amount helpers
# ---------------------------------------------------------------------------

def _to_float(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def _fmt_eur(v: float) -> str:
    v = float(v)
    if abs(v - round(v)) < 0.005:
        return f"€{int(round(v)):,}"
    return f"€{v:,.2f}"


# ---------------------------------------------------------------------------
# Build replacement map:  old_text → new_text
# ---------------------------------------------------------------------------

def _build_replacement_map(
    inv_orig: Dict[str, Any],
    inv_eur: Dict[str, Any],
) -> Dict[str, str]:
    rmap: Dict[str, str] = {}

    items_o = inv_orig.get("items") or []
    items_e = inv_eur.get("items") or []
    for io, ie in zip(items_o, items_e):
        for field in ("unit_price", "amount"):
            old_v = _to_float(io.get(field))
            new_v = _to_float(ie.get(field))
            if old_v is not None and new_v is not None:
                _add_amount_variants(rmap, old_v, _fmt_eur(new_v))

    tot_o = _to_float(inv_orig.get("total_amount"))
    tot_e = _to_float(inv_eur.get("total_amount"))
    if tot_o is not None and tot_e is not None:
        _add_amount_variants(rmap, tot_o, _fmt_eur(tot_e))

    rmap["￥0"] = "€0"
    rmap["¥0"] = "€0"
    return rmap


def _add_amount_variants(rmap: Dict[str, str], old_val: float, new_s: str):
    old_val = float(old_val)
    if abs(old_val - round(old_val)) < 0.01:
        n = int(round(old_val))
        for prefix in ("￥", "¥", "ĉ"):
            rmap[f"{prefix}{n:,}"] = new_s
            rmap[f"{prefix}{n}"] = new_s
    else:
        # 生成多种格式：2 位小数、1 位小数、原始精度
        s2 = f"{old_val:,.2f}"
        s_raw = f"{old_val:,}"
        variants = {s2, s_raw}
        if old_val == round(old_val, 1):
            variants.add(f"{old_val:,.1f}")
        for s in variants:
            for prefix in ("￥", "¥", "ĉ"):
                rmap[f"{prefix}{s}"] = new_s
            s_nodec = s.replace(",", "")
            for prefix in ("￥", "¥", "ĉ"):
                rmap[f"{prefix}{s_nodec}"] = new_s


# ---------------------------------------------------------------------------
# Normalize for matching
# ---------------------------------------------------------------------------

def _norm(t: str) -> str:
    return t.replace("ĉ", "¥").replace("￥", "¥").replace("\u00a0", " ")


def _match_span_text(span_text: str, rmap: Dict[str, str]) -> Optional[Tuple[str, str]]:
    """Match a span's text (or the bare number part) against rmap keys."""
    norm = _norm(span_text)
    for old_k in sorted(rmap.keys(), key=len, reverse=True):
        norm_k = _norm(old_k)
        if norm_k == norm:
            return (old_k, rmap[old_k])
        bare = norm_k.lstrip("¥")
        if bare and bare == norm:
            return (old_k, rmap[old_k])
    return None


# ---------------------------------------------------------------------------
# Background color detection
# ---------------------------------------------------------------------------

def _build_bg_rects(page: "fitz.Page") -> List[Tuple[Tuple[float, ...], "fitz.Rect"]]:
    """Collect non-white filled rectangles from page drawings."""
    result = []
    for d in page.get_drawings():
        fill = d.get("fill")
        if fill and fill != (1.0, 1.0, 1.0) and fill != (0.0, 0.0, 0.0):
            r = d.get("rect")
            if r and r.width > 10 and r.height > 5:
                result.append((fill, fitz.Rect(r)))
    return result


def _bg_color_at(bbox: "fitz.Rect", bg_rects: List) -> Tuple[float, float, float]:
    """Find the background color behind a given bbox. Returns white if none."""
    cx = (bbox.x0 + bbox.x1) / 2
    cy = (bbox.y0 + bbox.y1) / 2
    for fill, rect in bg_rects:
        if rect.x0 <= cx <= rect.x1 and rect.y0 <= cy <= rect.y1:
            return tuple(fill)
    return (1.0, 1.0, 1.0)


# ---------------------------------------------------------------------------
# Span color
# ---------------------------------------------------------------------------

def _span_color(span: dict) -> Tuple[float, float, float]:
    c = span.get("color", 0)
    if isinstance(c, int):
        r = ((c >> 16) & 0xFF) / 255.0
        g = ((c >> 8) & 0xFF) / 255.0
        b = (c & 0xFF) / 255.0
        return (r, g, b)
    return (0, 0, 0)


# ---------------------------------------------------------------------------
# Collect span-level replacements
# ---------------------------------------------------------------------------

def _collect_span_replacements(
    page: "fitz.Page",
    rmap: Dict[str, str],
    jpy_block: Optional[str],
    eur_block: Optional[str],
) -> List[Dict[str, Any]]:
    replacements: List[Dict[str, Any]] = []
    processed: Set[Tuple[float, ...]] = set()

    def _bbox_key(rect):
        return (round(rect.x0, 1), round(rect.y0, 1), round(rect.x1, 1), round(rect.y1, 1))

    d = page.get_text("dict")
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            i = 0
            while i < len(spans):
                span = spans[i]
                text = span["text"]
                norm = _norm(text)

                # JPY total block (e.g. "390,960JPY=2129.30EUR")
                if jpy_block and jpy_block in text and eur_block:
                    bk = _bbox_key(fitz.Rect(span["bbox"]))
                    if bk not in processed:
                        replacements.append({
                            "bbox": fitz.Rect(span["bbox"]),
                            "new_text": eur_block,
                            "size": span["size"],
                            "color": _span_color(span),
                        })
                        processed.add(bk)
                    i += 1
                    continue

                # ¥ symbol span + next number span → merge
                if norm.strip() in ("¥",) and i + 1 < len(spans):
                    next_span = spans[i + 1]
                    combined = text + next_span["text"]
                    match = _match_span_text(combined, rmap)
                    if match:
                        merged = fitz.Rect(span["bbox"]) | fitz.Rect(next_span["bbox"])
                        bk = _bbox_key(merged)
                        if bk not in processed:
                            replacements.append({
                                "bbox": merged,
                                "new_text": match[1],
                                "size": next_span["size"],
                                "color": _span_color(next_span),
                            })
                            processed.add(bk)
                        i += 2
                        continue

                # Single span match (number only, or ¥number in one span)
                match = _match_span_text(text, rmap)
                if match:
                    bk = _bbox_key(fitz.Rect(span["bbox"]))
                    if bk not in processed:
                        replacements.append({
                            "bbox": fitz.Rect(span["bbox"]),
                            "new_text": match[1],
                            "size": span["size"],
                            "color": _span_color(span),
                        })
                        processed.add(bk)
                    i += 1
                    continue

                # "JPY" standalone → "EUR"
                if "JPY" in text and span["bbox"][1] < 420:
                    bk = _bbox_key(fitz.Rect(span["bbox"]))
                    if bk not in processed:
                        replacements.append({
                            "bbox": fitz.Rect(span["bbox"]),
                            "new_text": text.replace("JPY", "EUR"),
                            "size": span["size"],
                            "color": _span_color(span),
                        })
                        processed.add(bk)

                i += 1

    return replacements


# ---------------------------------------------------------------------------
# Apply replacements
# ---------------------------------------------------------------------------

def _apply_replacements(
    page: "fitz.Page",
    replacements: List[Dict[str, Any]],
    bg_rects: List,
) -> None:
    if not replacements:
        return

    fontfile = _system_euro_font()
    if not fontfile:
        raise RuntimeError(
            "未找到支持 € 的系统字体。请安装 Arial Unicode 或 DejaVuSans。"
        )

    # Phase 1: redact with matching background color
    for r in replacements:
        rect = r["bbox"]
        bg = _bg_color_at(rect, bg_rects)
        expanded = fitz.Rect(rect.x0 - 0.5, rect.y0 - 0.5, rect.x1 + 0.5, rect.y1 + 0.5)
        page.add_redact_annot(expanded, fill=bg)

    page.apply_redactions()

    # Phase 2: insert with system font at original size
    for r in replacements:
        rect = r["bbox"]
        text = r["new_text"]
        size = r["size"]
        color = r["color"]

        descent_offset = size * 0.18
        baseline_y = rect.y1 - descent_offset

        page.insert_text(
            (rect.x0, baseline_y),
            text,
            fontsize=size,
            fontname="eurofont",
            fontfile=fontfile,
            color=color,
        )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def write_eur_invoice_pdf_from_original(
    original_pdf_path: str,
    output_path: str,
    invoice_data_eur: Dict[str, Any],
    invoice_data_original: Dict[str, Any],
) -> None:
    if fitz is None:
        raise ImportError("需要安装 PyMuPDF: pip install pymupdf")

    orig_path = Path(original_pdf_path)
    if not orig_path.exists():
        raise FileNotFoundError(original_pdf_path)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    inv_o = copy.deepcopy(invoice_data_original)
    inv_e = copy.deepcopy(invoice_data_eur)

    tot_e = _to_float(inv_e.get("total_amount"))
    cur_o = (inv_o.get("currency") or "").strip().upper()
    cur_e = (inv_e.get("currency") or "").strip().upper()

    if cur_o == "EUR" and cur_e == "EUR" and tot_e is not None:
        t_o = _to_float(inv_o.get("total_amount"))
        if t_o is not None and abs(t_o - tot_e) < 0.01:
            shutil.copyfile(str(orig_path), str(out))
            return

    rmap = _build_replacement_map(inv_o, inv_e)
    doc = fitz.open(str(orig_path))
    if len(doc) < 1:
        doc.close()
        raise ValueError("PDF 无页面")

    page0 = doc[0]

    # Detect JPY total block
    page_text = page0.get_text() or ""
    jpy_m = re.search(r"[\d,]+JPY=[\d.]+EUR", page_text)
    if not jpy_m:
        jpy_m = re.search(r"JPY[\d,=]+", page_text)
    jpy_block = jpy_m.group(0) if jpy_m else None

    eur_block = None
    if jpy_block and tot_e is not None:
        te = float(tot_e)
        core = str(int(round(te))) if abs(te - round(te)) < 0.005 else f"{te:.2f}"
        if "=" in jpy_block:
            suffix = "=" + jpy_block.split("=", 1)[1]
            if "EUR" in suffix:
                suffix = suffix.split("EUR")[0]
            eur_block = f"EUR{core}{suffix}EUR"
        else:
            eur_block = f"EUR{core}"

    bg_rects = _build_bg_rects(page0)
    replacements = _collect_span_replacements(page0, rmap, jpy_block, eur_block)
    _apply_replacements(page0, replacements, bg_rects)

    doc.save(str(out))
    doc.close()


def can_use_overlay() -> bool:
    return fitz is not None
