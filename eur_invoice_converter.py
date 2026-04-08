"""
将解析后的 invoice 数据从原币种换算为欧元（EUR）。

汇率含义：1 欧元 = fx_units_per_eur 单位原币种
（例：JPY 时填 165 表示 1 EUR = 165 JPY，金额 EUR = 原金额 / 165）
"""

from __future__ import annotations

import copy
from typing import Any, Dict, Optional


def _to_float(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def _fmt_eur(v: float) -> str:
    return f"{round(v, 2):.2f}"


def apply_eur_conversion(
    invoice_data: Dict[str, Any],
    fx_units_per_eur: float,
) -> Dict[str, Any]:
    """
    深拷贝 invoice_data，将单价、金额、合计换算为 EUR。

    :param fx_units_per_eur: 1 EUR 对应多少「发票货币单位」（须 > 0）
    :returns: 新 dict，currency 为 EUR，并附带 original_currency / fx_units_per_eur
    """
    if fx_units_per_eur <= 0:
        raise ValueError("fx_units_per_eur 必须为正数（表示 1 EUR 等于多少发票货币单位）")

    out: Dict[str, Any] = copy.deepcopy(invoice_data)
    orig_cur = (out.get("currency") or "").strip().upper() or "UNK"

    if orig_cur == "EUR":
        out["original_currency"] = "EUR"
        out["fx_units_per_eur"] = fx_units_per_eur
        out["fx_note"] = "发票已为 EUR，未做换算"
        return out

    def to_eur(val: Any) -> str:
        f = _to_float(val)
        if f is None:
            return str(val) if val is not None else ""
        return _fmt_eur(f / fx_units_per_eur)

    # items 与 customs_items 在无聚合 key 时是同一个 list 引用（共享 dict 对象），
    # 必须用 id() 去重，否则同一个 dict 会被除两次。
    converted_ids: set = set()
    for it in out.get("items") or []:
        if id(it) in converted_ids:
            continue
        converted_ids.add(id(it))
        if "unit_price" in it:
            it["unit_price"] = to_eur(it.get("unit_price"))
        if "amount" in it:
            it["amount"] = to_eur(it.get("amount"))

    for it in out.get("customs_items") or []:
        if id(it) in converted_ids:
            continue
        converted_ids.add(id(it))
        if "unit_price" in it:
            it["unit_price"] = to_eur(it.get("unit_price"))
        if "amount" in it:
            it["amount"] = to_eur(it.get("amount"))

    total_sum = 0.0
    for it in out.get("items") or []:
        f = _to_float(it.get("amount"))
        if f is not None:
            total_sum += f
    if total_sum > 0:
        out["total_amount"] = _fmt_eur(total_sum)
    elif out.get("total_amount"):
        out["total_amount"] = to_eur(out.get("total_amount"))

    out["currency"] = "EUR"
    out["original_currency"] = orig_cur
    out["fx_units_per_eur"] = fx_units_per_eur
    return out
