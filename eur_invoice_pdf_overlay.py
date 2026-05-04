"""
在「原始 Invoice PDF」上仅替换金额与相关币种字样，保留版式与版面。

策略：
1. 用 page.get_text("dict") 遍历所有 span，找到包含待替换金额的 span
2. 记录其 font/size/color/bbox
3. 检测该位置的背景色（来自绘图矩形），redact 时用同色填充
4. 用小体积字体 + 子集嵌入（避免 Arial Unicode 等整库嵌入把几十 KB 的 PDF 撑到数 MB）

依赖 PyMuPDF (fitz)。
"""

from __future__ import annotations

import copy
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None  # type: ignore

# EUR 版发票：条款区「Currency」列须显示 EUR；下列 ISO 在 PDF 中一律替换为 EUR（整格匹配）
_CURRENCY_ISO_TO_EUR: Tuple[str, ...] = (
    "AUD",
    "USD",
    "JPY",
    "GBP",
    "CAD",
    "NZD",
    "CHF",
    "CNY",
    "HKD",
    "SGD",
    "SEK",
    "NOK",
    "DKK",
    "PLN",
    "CZK",
    "HUF",
    "MXN",
    "BRL",
    "INR",
    "KRW",
    "THB",
    "MYR",
    "IDR",
    "PHP",
    "TWD",
    "ILS",
    "AED",
    "SAR",
    "QAR",
    "RUB",
    "TRY",
    "ZAR",
    "RON",
    "BGN",
    "HRK",
    "ISK",
)

# span 顶部 y0 大于此值视为页面下半（银行信息等），子串替换 ISO 时仅替换解析到的原币种，避免误改账号中的 USD 等
_TERMS_AREA_MAX_Y0: float = 460.0


# ---------------------------------------------------------------------------
# Font: 必须含 €；优先小文件（整库嵌入）。Arial Unicode 常 >20MB，切勿排在前面。
# ---------------------------------------------------------------------------

_SYSTEM_FONTS: Tuple[str, ...] = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Library/Fonts/Arial.ttf",
    r"C:\Windows\Fonts\arial.ttf",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    r"C:\Windows\Fonts\arialuni.ttf",
)

def _system_euro_font() -> Optional[str]:
    # 不缓存：避免升级字体优先级后仍沿用旧进程里已选的超大字库
    for p in _SYSTEM_FONTS:
        if Path(p).is_file():
            return p
    return None


# ---------------------------------------------------------------------------
# Amount helpers
# ---------------------------------------------------------------------------

def _to_float(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", "").replace(" ", "").strip())
    except (ValueError, TypeError):
        return None


def _thousands_space_int(n: int) -> str:
    """整数部分从右每三位一组，用空格分隔（如 7459 → 7 459）。"""
    n = abs(int(n))
    s = str(n)
    parts: List[str] = []
    while len(s) > 3:
        parts.append(s[-3:])
        s = s[:-3]
    parts.append(s)
    return " ".join(reversed(parts))


def _fmt_eur_number_eu(v: float) -> str:
    """欧式数字：千分位空格，小数 ','，始终两位小数。不含 €。"""
    v = float(v)
    neg = v < 0
    av = abs(v)
    cents_total = int(round(av * 100 + 1e-9))
    whole = cents_total // 100
    cents = cents_total % 100
    ws = _thousands_space_int(whole)
    body = f"{ws},{cents:02d}"
    return f"-{body}" if neg else body


def _fmt_eur(v: float) -> str:
    """EUR 显示用欧式数字，前缀 €。"""
    s = _fmt_eur_number_eu(v)
    if s.startswith("-"):
        return f"-€{s[1:]}"
    return f"€{s}"


def _fmt_eur_legacy_space_no_cents(v: float) -> str:
    """旧版：€ 整数无 ,00、千分位为空格（便于再生成时改为两位小数）。"""
    v = float(v)
    neg = v < 0
    n = int(round(abs(v)))
    body = _thousands_space_int(n)
    if neg:
        return f"-€{body}"
    return f"€{body}"


def _fmt_eur_us(v: float) -> str:
    """旧版/美式 € 文本，用于匹配 PDF 中已存在的 €1,234.56 以便改为欧式。"""
    v = float(v)
    if abs(v - round(v)) < 0.005:
        return f"€{int(round(v)):,}"
    return f"€{v:,.2f}"


def _fmt_eur_legacy_dot_thousands(v: float) -> str:
    """本工具曾用「千分位 .」的完整 € 串（仅作旧 PDF 匹配键；整数可无 ,00）。"""
    v = float(v)
    neg = v < 0
    av = abs(v)
    if abs(av - round(av)) < 0.005:
        n = int(round(av))
        body = f"{n:,}".replace(",", ".")
        s = f"€{body}"
    else:
        cents_total = int(round(av * 100 + 1e-9))
        whole = cents_total // 100
        cents = cents_total % 100
        ws = f"{whole:,}".replace(",", ".")
        s = f"€{ws},{cents:02d}"
    return f"-{s}" if neg else s


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

    # 若 PDF 中已是美式 € 或本工具上一版「千分位 .」输出，统一为「千分位空格」
    for ie in items_e:
        for field in ("unit_price", "amount"):
            nv = _to_float(ie.get(field))
            if nv is not None:
                new_s = _fmt_eur(nv)
                old_us = _fmt_eur_us(nv)
                if old_us != new_s:
                    rmap[old_us] = new_s
                old_dot = _fmt_eur_legacy_dot_thousands(nv)
                if old_dot != new_s:
                    rmap[old_dot] = new_s
                old_spc = _fmt_eur_legacy_space_no_cents(nv)
                if old_spc != new_s:
                    rmap[old_spc] = new_s
    if tot_e is not None:
        te = float(tot_e)
        new_s = _fmt_eur(te)
        old_us = _fmt_eur_us(te)
        if old_us != new_s:
            rmap[old_us] = new_s
        old_dot = _fmt_eur_legacy_dot_thousands(te)
        if old_dot != new_s:
            rmap[old_dot] = new_s
        old_spc = _fmt_eur_legacy_space_no_cents(te)
        if old_spc != new_s:
            rmap[old_spc] = new_s

    z = _fmt_eur(0.0)
    rmap["￥0"] = z
    rmap["¥0"] = z
    rmap["A$0"] = z
    rmap["A$0.00"] = z
    rmap["A$0,00"] = z

    # EUR 版发票：Currency 栏必须为 EUR（与解析器标成何币种无关；PDF 上可能仍印 AUD/USD）
    if (inv_eur.get("currency") or "").strip().upper() == "EUR":
        for code in _CURRENCY_ISO_TO_EUR:
            rmap[code] = "EUR"

    return rmap


def _add_amount_variants(rmap: Dict[str, str], old_val: float, new_s: str):
    old_val = float(old_val)
    if abs(old_val - round(old_val)) < 0.01:
        n = int(round(old_val))
        for prefix in ("￥", "¥", "ĉ", "A$"):
            rmap[f"{prefix}{n:,}"] = new_s
            rmap[f"{prefix}{n}"] = new_s
            # Tax Total 等常为 A$0.00 / ¥0.00
            rmap[f"{prefix}{old_val:,.2f}"] = new_s
            rmap[f"{prefix}{old_val:.2f}"] = new_s
    else:
        # 生成多种格式：2 位小数、1 位小数、原始精度
        s2 = f"{old_val:,.2f}"
        s_raw = f"{old_val:,}"
        variants = {s2, s_raw}
        if old_val == round(old_val, 1):
            variants.add(f"{old_val:,.1f}")
        for s in variants:
            for prefix in ("￥", "¥", "ĉ", "A$"):
                rmap[f"{prefix}{s}"] = new_s
            s_nodec = s.replace(",", "")
            for prefix in ("￥", "¥", "ĉ", "A$"):
                rmap[f"{prefix}{s_nodec}"] = new_s


# ---------------------------------------------------------------------------
# Normalize for matching
# ---------------------------------------------------------------------------

def _norm(t: str) -> str:
    return (
        t.replace("A$", "¥")
        .replace("ĉ", "¥")
        .replace("￥", "¥")
        .replace("\u00a0", " ")
    )


def _replace_rmap_substrings(
    text: str,
    rmap: Dict[str, str],
    *,
    orig_currency_iso: Optional[str] = None,
    span_y0: Optional[float] = None,
) -> str:
    """整行 span（如 Tax Total … A$0.00、条款行 … AUD …）内分段替换。"""
    keys = [
        k
        for k in rmap
        if len(k) > 2
        and (
            k.startswith("A$")
            or k.startswith("¥")
            or k.startswith("￥")
            or k.startswith("ĉ")
            or "JPY" in k
        )
    ]
    oc = (orig_currency_iso or "").strip().upper()
    in_terms = span_y0 is None or span_y0 < _TERMS_AREA_MAX_Y0
    if in_terms:
        # 条款/表体：可将整行中的 AUD/USD 等替换为 EUR；勿替换 JPY（易与 JPY 换算块冲突）
        for k in rmap:
            if (
                len(k) == 3
                and k.isalpha()
                and k.isupper()
                and rmap[k] == "EUR"
                and k != "JPY"
            ):
                keys.append(k)
    else:
        # 页面下方：仅替换解析到的原币种三字母，避免误改银行栏里的 USD 等
        if (
            oc
            and oc != "EUR"
            and oc != "JPY"
            and len(oc) == 3
            and oc.isalpha()
            and rmap.get(oc) == "EUR"
        ):
            keys.append(oc)
    keys = sorted(set(keys), key=len, reverse=True)
    result = text
    for k in keys:
        if k in result:
            result = result.replace(k, rmap[k])
    return result


def _match_span_text(span_text: str, rmap: Dict[str, str]) -> Optional[Tuple[str, str]]:
    """Match a span's text (or the bare number part) against rmap keys."""
    norm = _norm(span_text)
    ns = norm.strip()
    for old_k in sorted(rmap.keys(), key=len, reverse=True):
        norm_k = _norm(old_k)
        if norm_k == norm or norm_k.strip() == ns:
            return (old_k, rmap[old_k])
        bare = norm_k.lstrip("¥")
        if bare and bare == norm:
            return (old_k, rmap[old_k])
        if bare and bare.strip() == ns:
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
    orig_currency_iso: Optional[str] = None,
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

                # ¥ / € 符号与下一数字 span 分开时合并匹配
                if norm.strip() in ("¥", "€") and i + 1 < len(spans):
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

                new_sub = _replace_rmap_substrings(
                    text,
                    rmap,
                    orig_currency_iso=orig_currency_iso,
                    span_y0=span["bbox"][1],
                )
                if new_sub != text:
                    bk = _bbox_key(fitz.Rect(span["bbox"]))
                    if bk not in processed:
                        replacements.append({
                            "bbox": fitz.Rect(span["bbox"]),
                            "new_text": new_sub,
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
            "未找到支持 € 的系统字体。请安装 DejaVu/Liberation/Arial 或 Arial Unicode。"
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


def _save_pdf_compact(doc: "fitz.Document", out_path: str) -> None:
    # 仅保留文中用到的字形，避免嵌入完整 TTF（尤其曾误选 Arial Unicode 时）
    try:
        doc.subset_fonts()
    except Exception:
        pass
    kwargs = {"garbage": 4, "deflate": True, "clean": True}
    try:
        doc.save(out_path, **kwargs, deflate_images=True, deflate_fonts=True)
    except TypeError:
        doc.save(out_path, **kwargs)


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

    # 不直接 copy：原 PDF 的 Currency 栏可能仍印 AUD/USD，与「EUR 版发票」不一致，必须走 overlay

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
        core = _fmt_eur_number_eu(te)
        if "=" in jpy_block:
            suffix = "=" + jpy_block.split("=", 1)[1]
            if "EUR" in suffix:
                suffix = suffix.split("EUR")[0]
            eur_block = f"EUR{core}{suffix}EUR"
        else:
            eur_block = f"EUR{core}"

    bg_rects = _build_bg_rects(page0)
    replacements = _collect_span_replacements(
        page0, rmap, jpy_block, eur_block, orig_currency_iso=cur_o
    )
    _apply_replacements(page0, replacements, bg_rects)

    _save_pdf_compact(doc, str(out))
    doc.close()


def can_use_overlay() -> bool:
    return fitz is not None
