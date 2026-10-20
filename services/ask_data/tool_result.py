"""ToolResult: the single structured result every Ask Data tool produces, plus the drill-down store.

Everything shown to the user comes from here. A financial answer's amounts are Decimal values computed by
the deterministic tools from the ledger rows it retains; the model never supplies a number. `ungrounded_amounts`
re-checks a rendered text against those values as a safety net.

Fields: result value, currency, row count, source rows, calculation / components, applied filters, data
timestamp, warnings/caveats and a result_id so the rows behind a figure can be retrieved (drill-down).
"""
from __future__ import annotations

import datetime
import uuid
from collections import OrderedDict
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Callable, Optional

import pandas as pd

from services.ask_data.data_clock import format_as_of
from services.ask_data.money import amounts_in_text, fmt_eur, sum_money, to_money


def new_result_id() -> str:
    return uuid.uuid4().hex[:12]


@dataclass(frozen=True)
class Component:
    """One labelled part of a calculation (e.g. a country, or 'Charges' / 'Credits'), with the rows that make it up."""
    label: str
    amount: Decimal
    row_count: int = 0
    rows: tuple = ()


@dataclass
class ToolResult:
    tool: str
    text: str                                   # the headline sentence (numbers come from the fields below)
    dataframe: Optional[pd.DataFrame] = None    # table shown to the user
    financial: bool = False
    value: Optional[Decimal] = None
    currency: str = "EUR"
    row_count: int = 0
    rows: list = field(default_factory=list)   # the ledger rows (or bank allocations) behind `value`
    components: list = field(default_factory=list)
    calculation: Optional[str] = None
    derived_amounts: list = field(default_factory=list)   # amounts shown that are computed from the above (e.g. a difference)
    filters: dict = field(default_factory=dict)
    data_as_of: Optional[datetime.datetime] = None
    warnings: list = field(default_factory=list)
    result_id: str = ""
    frame_builder: Optional[Callable] = None    # rows -> DataFrame, used by drill-down

    # ── rendering ───────────────────────────────────────────────────────────────────────────────
    def scope_text(self) -> str:
        parts = []
        for key, value in self.filters.items():
            if value is None:
                continue
            parts.append(f"{key}={value.isoformat() if hasattr(value, 'isoformat') else value}")
        return ", ".join(parts) if parts else "all rows"

    def display_text(self) -> str:
        """Headline + calculation + scope + freshness + caveats. Non-financial results are shown as-is."""
        if not self.financial:
            return self.text
        paragraphs = [self.text]
        if self.calculation:
            paragraphs.append(f"Calculation: {self.calculation}")
        paragraphs.append(f"Scope: {self.scope_text()}.")
        paragraphs.append(f"{format_as_of(self.data_as_of)}.")
        paragraphs.extend(f"Note: {w}" for w in self.warnings)
        return "\n\n".join(paragraphs)

    # ── grounding ───────────────────────────────────────────────────────────────────────────────
    def grounded_amounts(self) -> set:
        allowed: set = set()

        def add(amount):
            amount = to_money(amount)
            allowed.add(amount)
            allowed.add(-amount)

        if self.value is not None:
            add(self.value)
        for amount in self.derived_amounts:
            add(amount)
        for component in self.components:
            add(component.amount)
            if component.rows:
                add(sum_money(_row_amount(r) for r in component.rows))
        if self.rows:
            add(sum_money(_row_amount(r) for r in self.rows))
        for row in self.rows:
            add(_row_amount(row))
        for amount in _frame_amounts(self.dataframe):
            add(amount)
        return allowed

    def ungrounded_amounts(self, text: Optional[str] = None) -> list:
        """Euro amounts in `text` (default: the display text) that do not come from this result. Empty = grounded."""
        allowed = self.grounded_amounts()
        return [fmt_eur(a) for a in amounts_in_text(self.display_text() if text is None else text) if to_money(a) not in allowed]

    def summary(self) -> dict:
        """JSON-safe description (no rows) for logging/tests."""
        return {
            "result_id": self.result_id, "tool": self.tool, "financial": self.financial,
            "value": None if self.value is None else str(self.value), "currency": self.currency,
            "row_count": self.row_count, "filters": {k: str(v) for k, v in self.filters.items() if v is not None},
            "components": [(c.label, str(c.amount), c.row_count) for c in self.components],
            "calculation": self.calculation, "warnings": list(self.warnings),
            "data_as_of": None if self.data_as_of is None else self.data_as_of.isoformat(),
        }


def _row_amount(row) -> Decimal:
    if isinstance(row, dict):
        return to_money(row.get("amount_applied", row.get("payment_amount")))
    return to_money(getattr(row, "payment_amount", None))


def _frame_amounts(frame: Optional[pd.DataFrame]) -> list:
    """Euro columns of the displayed table (they are built from the same rows; included so a shown row amount is grounded)."""
    if frame is None:
        return []
    amounts = []
    for column in frame.columns:
        if "€" in str(column):
            for value in frame[column]:
                if isinstance(value, Decimal):
                    amounts.append(value)
    return amounts


# ── drill-down ───────────────────────────────────────────────────────────────────────────────────
class UnknownResultError(KeyError):
    """The result_id is not (or no longer) in the store."""


@dataclass(frozen=True)
class RowsResult:
    result_id: str
    tool: str
    component: Optional[str]
    row_count: int
    total: Decimal
    rows: tuple
    dataframe: Optional[pd.DataFrame]


class ResultStore:
    """Bounded, in-memory store of recent results so "show me the invoices behind this total" can be answered.

    Keep one store per user session (the router uses st.session_state); results are never shared across sessions.
    """

    def __init__(self, max_results: int = 25):
        self.max_results = max_results
        self._results: "OrderedDict[str, ToolResult]" = OrderedDict()

    def add(self, result: ToolResult) -> str:
        if not result.result_id:
            result.result_id = new_result_id()
        self._results[result.result_id] = result
        self._results.move_to_end(result.result_id)
        while len(self._results) > self.max_results:
            self._results.popitem(last=False)
        return result.result_id

    def get(self, result_id: str) -> Optional[ToolResult]:
        return self._results.get(result_id)

    def latest(self) -> Optional[ToolResult]:
        return next(reversed(self._results.values()), None)

    def get_result_rows(self, result_id: str, component: Optional[str] = None) -> RowsResult:
        """The ledger rows behind a result (optionally one named component, e.g. a country). Read-only."""
        result = self._results.get(result_id)
        if result is None:
            raise UnknownResultError(result_id)
        rows = list(result.rows)
        if component is not None:
            matches = [c for c in result.components if c.label == component]
            if not matches:
                raise KeyError(f"result {result_id} has no component {component!r}; available: {[c.label for c in result.components]}")
            rows = list(matches[0].rows)
        frame = result.frame_builder(rows) if (result.frame_builder is not None and rows) else None
        return RowsResult(result_id, result.tool, component, len(rows), sum_money(_row_amount(r) for r in rows), tuple(rows), frame)


DEFAULT_STORE = ResultStore()


def get_result_rows(result_id: str, component: Optional[str] = None, store: Optional[ResultStore] = None) -> RowsResult:
    return (store or DEFAULT_STORE).get_result_rows(result_id, component)
