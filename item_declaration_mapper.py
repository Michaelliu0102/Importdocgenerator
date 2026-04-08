"""
Invoice 行 description / item_code 与申报品名、申报要素的映射。
优先：Excel「ITEM / 申报品名」表；其次：item_code 前四位 5012/5015、MF 前缀、NAPPA/VERONA/ROMA 前缀；再其次 hide_type；最后默认 product_info。
"""

from __future__ import annotations

import copy
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook

from config_loader import ConfigLoader

ROOT = Path(__file__).resolve().parent
DEFAULT_EXCEL = ROOT / "data" / "item和品名对应表.xlsx"

# Excel 英文 ITEM 列与 supplier_product_mapping.yaml 中 product_categories 键的对应
EXCEL_ITEM_TO_YAML_KEY: Dict[str, str] = {
    "Phone Case": "export_hs_392690_misc",
    "Airpods Case": "export_hs_392690_misc",
    "Watch Band": "export_hs_911390_watch",
    "Sunvisor Cover": "export_hs_392690_misc",
    "Pouch": "export_hs_392690_misc",
    "Lumbar Cushion": "export_hs_940199_auto",
    "Headrest Pillow": "export_hs_940199_auto",
    "Tray": "export_hs_392690_misc",
    "Bagpack": "export_hs_420212_luggage",
    "Luggage": "export_hs_420212_luggage",
    "Duffle Bag": "export_hs_420212_luggage",
    "Key Fob": "export_hs_392690_misc",
    "Card Holder": "export_hs_392690_misc",
    "Seat Cover": "alcantara_5015_5030_5010",
    "Notebook": "export_hs_482010_notebook",
    "Color Cards": "export_hs_491110_print",
    "Spectacle Case": "export_hs_392690_misc",
}


def _resolve_config_path(config_path: str) -> Path:
    p = Path(config_path)
    if not p.is_absolute():
        p = (Path.cwd() / p).resolve()
    else:
        p = p.resolve()
    return p


def load_excel_mapping_rows(xlsx_path: Path) -> List[Tuple[str, str]]:
    """返回 [(ITEM英文, 申报品名中文), ...]，不含表头。"""
    if not xlsx_path.exists():
        return []
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows: List[Tuple[str, str]] = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue
            if not row or len(row) < 2:
                continue
            a, b = row[0], row[1]
            if a is None or b is None:
                continue
            item_en = str(a).strip()
            name_cn = str(b).strip()
            if not item_en or not name_cn:
                continue
            rows.append((item_en, name_cn))
        return rows
    finally:
        wb.close()


def _match_excel_row(
    description: str, excel_rows: List[Tuple[str, str]]
) -> Optional[Tuple[str, str]]:
    """在 description 中匹配最长的 ITEM 子串（不区分大小写）。"""
    d = (description or "").strip().lower()
    if not d or not excel_rows:
        return None
    best: Optional[Tuple[str, str]] = None
    best_len = -1
    for item_en, name_cn in excel_rows:
        ie = item_en.strip().lower()
        if not ie:
            continue
        if ie in d and len(ie) > best_len:
            best = (item_en, name_cn)
            best_len = len(ie)
    return best


def _first_four_digits(item_code: Any) -> Optional[str]:
    if item_code is None:
        return None
    digits = re.sub(r"\D", "", str(item_code).strip())
    if len(digits) >= 4:
        return digits[:4]
    return None


def _leading_after_line_digits(s: str) -> str:
    """去掉行首数字编号后用于 MF / 皮革前缀判断。"""
    t = (s or "").strip()
    t = re.sub(r"^\d+[\s\-]*", "", t)
    return t.strip()


def _mf_prefix(description: str) -> bool:
    lead = _leading_after_line_digits(description)
    return len(lead) >= 2 and lead[:2].upper() == "MF"


def _leather_keyword_prefix(description: str) -> bool:
    """NAPPA / VERONA / ROMA 为去掉行首数字后的字母前缀（大小写不敏感）。"""
    s = _leading_after_line_digits(description)
    if not s:
        return False
    u = s.upper()
    if len(u) >= 5 and u.startswith("NAPPA"):
        return True
    if len(u) >= 6 and u.startswith("VERONA"):
        return True
    if len(u) >= 4 and u.startswith("ROMA"):
        return True
    return False


def _product_block_from_yaml(
    loader: ConfigLoader, key: str
) -> Dict[str, Any]:
    pi = loader.get_product_info(key) or {}
    return {
        "hs_code": (pi.get("hs_code") or "").strip(),
        "product_name": (pi.get("name") or "").strip(),
        "declaration_elements": copy.deepcopy(pi.get("declaration_elements") or {}),
    }


def _with_cn_name(
    block: Dict[str, Any], name_cn: str
) -> Dict[str, Any]:
    out = {
        "hs_code": block["hs_code"],
        "product_name": name_cn,
        "declaration_elements": copy.deepcopy(block["declaration_elements"]),
    }
    decl = out["declaration_elements"]
    if name_cn:
        decl["品名"] = name_cn
    return out


def resolve_item_declaration(
    item: Dict[str, Any],
    excel_rows: List[Tuple[str, str]],
    loader: ConfigLoader,
    product_info: Optional[Dict[str, Any]],
    product_info_map: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    """
    返回单行的 hs_code, product_name, declaration_elements（已深拷贝）。
    """
    desc = (item.get("description") or "").strip()
    pi_default = product_info or {}

    # 1) Excel ITEM 表
    ex = _match_excel_row(desc, excel_rows)
    if ex:
        item_en, name_cn = ex
        yk = EXCEL_ITEM_TO_YAML_KEY.get(item_en.strip())
        if yk:
            blk = _product_block_from_yaml(loader, yk)
            return _with_cn_name(blk, name_cn)

    # 2) hide_type（与旧逻辑一致，在数字/MF/皮革规则之前便于专用皮料）
    hide_type = (item.get("hide_type") or "").strip()
    if hide_type and hide_type in product_info_map:
        cur = product_info_map[hide_type]
        return {
            "hs_code": (cur.get("hs_code") or "").strip(),
            "product_name": (cur.get("name") or "").strip(),
            "declaration_elements": copy.deepcopy(
                cur.get("declaration_elements") or {}
            ),
        }

    # 3) item 前四位 5012 / 5015 → 面料 5603149000
    d4 = _first_four_digits(item.get("item_code"))
    if d4 == "5012":
        blk = _product_block_from_yaml(loader, "alcantara_5012")
        return _with_cn_name(blk, "面料")
    if d4 == "5015":
        blk = _product_block_from_yaml(loader, "alcantara_5015_5030_5010")
        return _with_cn_name(blk, "面料")

    # 4) MF 开头 → 超细纤维革 5903209000
    if _mf_prefix(desc):
        blk = _product_block_from_yaml(loader, "export_microfibre_hs5903209000")
        return _with_cn_name(blk, blk["product_name"] or "超细纤维革")

    # 5) NAPPA / VERONA / ROMA → 牛皮 4107121090
    if _leather_keyword_prefix(desc):
        blk = _product_block_from_yaml(loader, "mabo_cowhide")
        return _with_cn_name(blk, "牛皮")

    # 6) 默认 invoice 级 product_info
    return {
        "hs_code": (pi_default.get("hs_code") or "").strip(),
        "product_name": (pi_default.get("name") or "").strip(),
        "declaration_elements": copy.deepcopy(
            pi_default.get("declaration_elements") or {}
        ),
    }


def build_declaration_groups(
    items: List[Dict[str, Any]],
    invoice_data: Dict[str, Any],
    product_info: Optional[Dict[str, Any]],
    product_info_map: Optional[Dict[str, Dict[str, Any]]],
    config_path: str,
    mapping_xlsx: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """
    使用 item/品名 对应表与回落规则分组，结构同 DeclarationElementsGenerator._group_items_by_product。
    若未启用或缺少映射表，调用方应改用旧 _group_items_by_product。
    """
    cfg_path = _resolve_config_path(config_path)
    if not cfg_path.exists():
        cfg_path = ROOT / "data" / "supplier_product_mapping.yaml"
    loader = ConfigLoader(str(cfg_path))

    xlsx = _resolve_mapping_xlsx_path(str(cfg_path), mapping_xlsx, invoice_data)
    excel_rows = load_excel_mapping_rows(xlsx)

    pim = product_info_map or {}
    if not items:
        pi = product_info or {}
        return [
            {
                "hs_code": pi.get("hs_code", ""),
                "product_name": pi.get("name", ""),
                "declaration_elements": copy.deepcopy(
                    pi.get("declaration_elements") or {}
                ),
                "items": [],
            }
        ]

    groups: Dict[Tuple[str, str], Dict[str, Any]] = {}

    for item in items:
        resolved = resolve_item_declaration(
            item, excel_rows, loader, product_info, pim
        )
        key = (resolved["hs_code"], resolved["product_name"])
        if key not in groups:
            groups[key] = {
                "hs_code": resolved["hs_code"],
                "product_name": resolved["product_name"],
                "declaration_elements": resolved["declaration_elements"],
                "items": [],
            }
        groups[key]["items"].append(item)

    return list(groups.values())


def _resolve_mapping_xlsx_path(
    config_path: str,
    mapping_xlsx: Optional[str],
    invoice_data: Optional[Dict[str, Any]],
) -> Path:
    ovr = (invoice_data or {}).get("item_mapping_xlsx") or mapping_xlsx
    if ovr:
        p = Path(ovr)
        return p.resolve() if p.is_absolute() else (Path.cwd() / p).resolve()
    return _resolve_config_path(config_path).parent / "item和品名对应表.xlsx"


def use_item_mapping_enabled(invoice_data: Dict[str, Any]) -> bool:
    """默认启用 ITEM/品名映射与回落规则；invoice_data['use_item_mapping']==False 时退回旧分组。"""
    return invoice_data.get("use_item_mapping", True) is not False
