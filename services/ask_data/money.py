"""Exact money arithmetic for Ask Data (Decimal, cents). No float is ever used for an amount here.

Ledger amounts arrive as floats from the data layer; they are converted once, through their shortest decimal
representation (so 1000.5 is exactly 1000.50), quantised to cents with ROUND_HALF_UP, and from then on only
Decimal is used for sums, differences, comparisons and percentages.
"""
from __future__ import annotations

import re
from decimal import ROUND_HALF_UP, Decimal
from typing import Iterable, Optional

CENT = Decimal("0.01")
ZERO = Decimal("0.00")
_AMOUNT_RE = re.compile(r"(-?)€(-?)([\d,]+(?:\.\d+)?)")


def to_money(value) -> Decimal:
    """Convert a ledger value (float/int/str/Decimal/None) to a Decimal quantised to cents."""
    if value is None or value == "":
        return ZERO
    if isinstance(value, Decimal):
        amount = value
    elif isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            return ZERO
        amount = Decimal(repr(value))
    else:
        try:
            amount = Decimal(str(value))
        except Exception:
            return ZERO
    return amount.quantize(CENT, rounding=ROUND_HALF_UP)


def sum_money(values: Iterable) -> Decimal:
    total = ZERO
    for value in values:
        total += to_money(value)
    return total


def fmt_eur(amount: Decimal) -> str:
    """EUR with cents and a leading minus for negatives: €1,234.50 / -€2,723.00."""
    amount = to_money(amount)
    return f"{'-' if amount < 0 else ''}€{abs(amount):,.2f}"


def percent(part: Decimal, whole: Decimal, places: str = "0.1") -> Optional[Decimal]:
    """part / whole as a percentage (Decimal, half-up). None when whole is zero."""
    if whole == 0:
        return None
    return (to_money(part) / to_money(whole) * 100).quantize(Decimal(places), rounding=ROUND_HALF_UP)


def fmt_percent(value: Optional[Decimal]) -> str:
    return "n/a" if value is None else f"{value}%"


def amounts_in_text(text: str) -> list[Decimal]:
    """Every euro amount that appears in a displayed text, as signed Decimals (used by the grounding check)."""
    found = []
    for sign_a, sign_b, digits in _AMOUNT_RE.findall(text or ""):
        value = Decimal(digits.replace(",", ""))
        found.append(-value if (sign_a or sign_b) else value)
    return found
