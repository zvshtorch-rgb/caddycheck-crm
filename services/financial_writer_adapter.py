"""Compatibility gate only; this release does not install staged financial writers."""
import os

GATE = "CADDYCHECK_COVERAGE_WRITERS"


class FinancialWriteBlocked(ValueError):
    """Invalid financial-writer compatibility configuration."""


def enabled() -> bool:
    value = os.environ.get(GATE, "off").strip().lower()
    if value not in ("off", "on"):
        raise FinancialWriteBlocked(f"{GATE} must be exactly off or on.")
    return value == "on"
