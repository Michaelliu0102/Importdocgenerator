"""
独立将 CustInvc 发票 PDF 转为 EUR 版（不跑出口报关全套流程）。
输出可压缩以尽量满足最大文件大小（默认 1MB）。
"""

from __future__ import annotations

import copy
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

from eur_invoice_converter import apply_eur_conversion
from eur_invoice_pdf_overlay import can_use_overlay, write_eur_invoice_pdf_from_original
from main import CustomsDocGenerator, _safe_filename
from pdf_parser import InvoiceParser

# 默认上限 1MB
DEFAULT_MAX_PDF_BYTES = 1024 * 1024


def _fitz_recompress_file(path: Path) -> bool:
    """用 PyMuPDF 重写保存；若变小则替换。返回是否成功打开并写出。"""
    try:
        import fitz
    except ImportError:
        return False
    tmp = path.with_name(path.stem + "_compact.pdf")
    doc = fitz.open(str(path))
    try:
        try:
            doc.subset_fonts()
        except Exception:
            pass
        kwargs = {"garbage": 4, "deflate": True, "clean": True}
        try:
            doc.save(str(tmp), **kwargs, deflate_images=True, deflate_fonts=True)
        except TypeError:
            doc.save(str(tmp), **kwargs)
    finally:
        doc.close()
    if not tmp.exists():
        return False
    if tmp.stat().st_size < path.stat().st_size:
        tmp.replace(path)
        return True
    tmp.unlink(missing_ok=True)
    return False


def _ghostscript_recompress(path: Path, max_bytes: int) -> bool:
    """若系统有 gs，依次尝试 ebook / screen；变小则替换；达 max_bytes 则提前结束。"""
    gs = shutil.which("gs")
    if not gs:
        return False
    tmp = path.with_name(path.stem + "_gs.pdf")
    changed = False
    for settings in ("/ebook", "/screen"):
        try:
            subprocess.run(
                [
                    gs,
                    "-q",
                    "-sDEVICE=pdfwrite",
                    "-dCompatibilityLevel=1.4",
                    f"-dPDFSETTINGS={settings}",
                    "-dNOPAUSE",
                    "-dBATCH",
                    "-dDetectDuplicateImages=true",
                    f"-sOutputFile={tmp}",
                    str(path),
                ],
                check=True,
                timeout=300,
                capture_output=True,
            )
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError):
            if tmp.exists():
                tmp.unlink(missing_ok=True)
            continue
        if tmp.exists() and tmp.stat().st_size < path.stat().st_size:
            tmp.replace(path)
            changed = True
            if path.stat().st_size <= max_bytes:
                return True
        elif tmp.exists():
            tmp.unlink(missing_ok=True)
    return changed


def ensure_pdf_under_max_bytes(
    path: Path,
    max_bytes: int = DEFAULT_MAX_PDF_BYTES,
) -> Tuple[bool, str]:
    """
    若超过 max_bytes，依次尝试：PyMuPDF 多轮 deflate、Ghostscript（若可用）。
    仍超限则返回 (False, 说明)。
    """
    if path.stat().st_size <= max_bytes:
        return True, ""

    notes: list[str] = []
    pymupdf_noted = False

    # 多轮 PyMuPDF：扫描件大图有时第二轮还能略减
    for _ in range(3):
        if path.stat().st_size <= max_bytes:
            return True, ("；".join(notes) if notes else "")
        prev = path.stat().st_size
        if _fitz_recompress_file(path) and path.stat().st_size < prev:
            if not pymupdf_noted:
                notes.append("已用 PyMuPDF 压缩")
                pymupdf_noted = True
            continue
        break

    if path.stat().st_size <= max_bytes:
        return True, ("；".join(notes) if notes else "")

    if _ghostscript_recompress(path, max_bytes):
        notes.append("已用 Ghostscript 压缩")

    for _ in range(2):
        if path.stat().st_size <= max_bytes:
            return True, ("；".join(notes) if notes else "")
        if not _fitz_recompress_file(path):
            break

    if path.stat().st_size <= max_bytes:
        return True, ("；".join(notes) if notes else "")

    mb = path.stat().st_size / (1024 * 1024)
    tried = "PyMuPDF"
    if any("Ghostscript" in n for n in notes):
        tried += "、Ghostscript"
    tail = f"已尝试压缩（{tried}），当前约 {mb:.2f} MB（目标 ≤ {max_bytes // (1024 * 1024)} MB）"
    if notes:
        return False, "；".join(notes) + "。" + tail
    return False, tail


def convert_custinvc_to_eur_pdf(
    source_pdf: str,
    output_dir: str,
    fx_units_per_eur: Optional[float] = None,
    *,
    config_path: Optional[str] = None,
    templates_dir: Optional[str] = None,
    enable_ocr: bool = False,
    ocr_lang: str = "eng",
    max_pdf_bytes: int = DEFAULT_MAX_PDF_BYTES,
) -> Tuple[str, str]:
    """
    解析 CustInvc PDF，换算 EUR，在原版式上替换金额，写出 PDF 并尽量压到 max_pdf_bytes 以下。

    :returns: (输出 PDF 绝对路径, 提示信息；无问题时提示为空或仅含压缩说明)
    :raises ValueError: 非 camari_cust、无汇率且原币非 EUR、无 PyMuPDF 等
    """
    base = Path(__file__).resolve().parent
    cfg = config_path or str(base / "data" / "supplier_product_mapping_import.yaml")
    tpl = templates_dir or str(base / "templates")

    src = Path(source_pdf).expanduser().resolve()
    if not src.is_file():
        raise FileNotFoundError(str(src))

    out_dir = Path(output_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    working_pdf = str(src)
    if enable_ocr:
        gen = CustomsDocGenerator(config_path=cfg, templates_dir=tpl)
        working_pdf = gen._prepare_invoice_with_ocr(working_pdf, out_dir, ocr_lang)

    parser = InvoiceParser(working_pdf)
    invoice_data = parser.parse()
    fmt = invoice_data.get("format") or ""
    if fmt != "camari_cust":
        raise ValueError(
            f"当前仅支持 CAMARI CustInvc 版式（解析为 {fmt!r}），请确认文件为 CustInvc_*.pdf"
        )

    inv_before = copy.deepcopy(invoice_data)
    cur = (invoice_data.get("currency") or "").strip().upper()

    if cur != "EUR":
        if fx_units_per_eur is None or fx_units_per_eur <= 0:
            raise ValueError("原币非 EUR 时，请填写「1 EUR = 多少发票货币」且为大于 0 的数字")
        invoice_data = apply_eur_conversion(invoice_data, float(fx_units_per_eur))
    else:
        # 已是 EUR：仍可走 overlay（通常直接复制原 PDF）
        pass

    if not can_use_overlay():
        raise ImportError("需要安装 PyMuPDF：pip install pymupdf")

    inv_key = _safe_filename(invoice_data.get("invoice_no") or src.stem)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    outp = out_dir / f"CustInvc_EUR_{inv_key}_{ts}.pdf"

    write_eur_invoice_pdf_from_original(
        working_pdf,
        str(outp),
        invoice_data,
        inv_before,
    )

    ok, note = ensure_pdf_under_max_bytes(outp, max_pdf_bytes)
    if not ok and note:
        return str(outp), note
    return str(outp), note or ""


def convert_custinvc_to_eur_pdf_safe(
    source_pdf: str,
    output_dir: str,
    fx_units_per_eur: Optional[float],
    **kwargs,
) -> Tuple[Optional[str], str]:
    """不抛异常，返回 (路径或 None, 错误/说明文本)。"""
    try:
        if fx_units_per_eur is not None:
            fx = float(fx_units_per_eur)
            if fx <= 0:
                return None, "汇率必须大于 0"
        else:
            fx = None
        return convert_custinvc_to_eur_pdf(
            source_pdf,
            output_dir,
            fx_units_per_eur=fx,
            **kwargs,
        )
    except Exception as e:
        return None, str(e)
