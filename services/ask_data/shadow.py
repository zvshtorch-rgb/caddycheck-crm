"""Ask Data v2/v3 shadow comparison (read-only diagnostics for the editor-only v3 rollout).

While an approved editor is piloting Ask Data v3 (see streamlit_app.py's mode switcher, gated by
`get_ask_data_v3_enabled()` + CAN_EDIT), every question is ALSO answered by the other pipeline in the
background ("shadow run"): whichever of v2/v3 is not shown to the user still gets evaluated so the two can
be compared. Nothing here ever changes what the user sees, and nothing here is financial/production data --
it is purely an engineering diagnostic, built once a comparison already exists, from two already-computed
`ToolResult`s (or their absence).

Privacy: the stored comparison never contains ledger rows, invoice numbers, payment amounts/dates, bank
details or Gemini API keys. Project names (potentially customer-identifying) are reduced to a short hash so
"did the two pipelines pick the same project?" can still be answered without storing the name itself. The
free-text question and the *formatted* financial total (e.g. "EUR 19,999.00", already an aggregate CRM
figure, not a raw bank record) are kept, since they are what makes the diagnostic useful at all.
"""
from __future__ import annotations

import datetime
import hashlib
from dataclasses import dataclass, field
from typing import Optional

from services.ask_data.money import fmt_eur
from services.ask_data.tool_result import ToolResult

# The only argument keys ever logged verbatim; everything else (e.g. `project`, `limit`, `invoice_number`)
# is either hashed (`project`) or dropped entirely.
_LOGGED_FILTER_KEYS = ("country", "compare_country", "category", "year", "paid_year", "paid")

MISMATCH_TOTALS_DIFFER = "totals_differ"
MISMATCH_ROW_COUNT_DIFFERS = "row_count_differs"
MISMATCH_ROW_SET_DIFFERS = "row_set_differs"
MISMATCH_SCOPE_DIFFERS = "scope_differs"
MISMATCH_CLARIFICATION_DIFFERS = "clarification_differs"
MISMATCH_OUTCOME_KIND_DIFFERS = "outcome_kind_differs"


def _short_hash(value: object) -> str:
    """A short, non-reversible fingerprint -- enough to test equality, never enough to recover the value."""
    return hashlib.sha256(str(value).encode("utf-8", errors="replace")).hexdigest()[:16]


def _row_set_hash(dataframe) -> Optional[str]:
    """Fingerprint of a result table's exact content (every cell, every row), order-independent. None if there
    is no table. Lets two comparisons detect "same rows" / "different rows" without storing a single cell."""
    if dataframe is None or len(dataframe) == 0:
        return None
    row_tuples = sorted(tuple(str(v) for v in row) for row in dataframe.itertuples(index=False, name=None))
    return _short_hash(row_tuples)


@dataclass
class ShadowSummary:
    """What one pipeline (v2 or v3) did with a single question, reduced to a privacy-safe, comparable shape."""
    outcome_kind: str                                  # "tool" | "clarification" | "drilldown" | "legacy" | "none" | "error"
    tool: Optional[str] = None
    filters: dict = field(default_factory=dict)        # only _LOGGED_FILTER_KEYS, never project/invoice_number/limit
    project_present: bool = False
    project_hash: Optional[str] = None
    financial: bool = False
    value_display: Optional[str] = None                # e.g. "EUR 19,999.00"; already an aggregate, not a raw row
    row_count: Optional[int] = None
    row_set_hash: Optional[str] = None
    clarification_asked: bool = False

    @classmethod
    def from_tool_result(cls, result: Optional[ToolResult], *, outcome_kind: str = "tool") -> "ShadowSummary":
        if result is None:
            return cls(outcome_kind="none")
        filters = {k: v for k, v in result.filters.items() if k in _LOGGED_FILTER_KEYS and v is not None}
        project_value = result.filters.get("project")
        return cls(
            outcome_kind=outcome_kind,
            tool=result.tool,
            filters=filters,
            project_present=bool(project_value),
            project_hash=_short_hash(project_value) if project_value else None,
            financial=bool(result.financial),
            value_display=(fmt_eur(result.value) if result.financial and result.value is not None else None),
            row_count=int(result.row_count) if result.row_count is not None else None,
            row_set_hash=_row_set_hash(result.dataframe),
        )

    @classmethod
    def from_clarification(cls) -> "ShadowSummary":
        return cls(outcome_kind="clarification", clarification_asked=True)

    @classmethod
    def from_legacy(cls, dataframe=None) -> "ShadowSummary":
        """The legacy keyword matcher has no tool/filters concept; only its row table (if any) is comparable."""
        return cls(outcome_kind="legacy", row_count=(len(dataframe) if dataframe is not None else None),
                   row_set_hash=_row_set_hash(dataframe))

    @classmethod
    def from_error(cls) -> "ShadowSummary":
        return cls(outcome_kind="error")

    def as_dict(self) -> dict:
        return {
            "outcome_kind": self.outcome_kind, "tool": self.tool, "filters": dict(self.filters),
            "project_present": self.project_present, "project_hash": self.project_hash,
            "financial": self.financial, "value_display": self.value_display,
            "row_count": self.row_count, "row_set_hash": self.row_set_hash,
            "clarification_asked": self.clarification_asked,
        }


@dataclass
class ShadowComparison:
    timestamp: str
    question: str
    primary_mode: str                  # "v2" or "v3" -- which one the user actually saw
    editor: Optional[str]
    v2: ShadowSummary
    v3: ShadowSummary
    mismatch_flags: list = field(default_factory=list)

    @property
    def has_mismatch(self) -> bool:
        return bool(self.mismatch_flags)

    def to_log_entry(self) -> dict:
        entry = {
            "timestamp": self.timestamp, "question": self.question, "primary_mode": self.primary_mode,
            "editor": self.editor, "mismatch_flags": list(self.mismatch_flags), "has_mismatch": self.has_mismatch,
        }
        for prefix, summary in (("v2", self.v2), ("v3", self.v3)):
            for key, value in summary.as_dict().items():
                entry[f"{prefix}_{key}"] = value
        return entry


def _scope_differs(v2: ShadowSummary, v3: ShadowSummary) -> bool:
    if v2.filters != v3.filters:
        return True
    if v2.project_present != v3.project_present:
        return True
    if v2.project_present and v3.project_present and v2.project_hash != v3.project_hash:
        return True
    return False


def build_comparison(
    question: str, primary_mode: str, v2: ShadowSummary, v3: ShadowSummary, *, editor: Optional[str] = None,
) -> ShadowComparison:
    """Compare a v2 summary and a v3 summary for the SAME question and flag any divergence (goal 6).

    Follow-up questions ("Only Belgium?", "Which invoices?", ...) are expected to diverge when v2 has no
    context to resolve them and v3 does (goal 7) -- that shows up here as `outcome_kind_differs` (and/or
    `scope_differs`), not as a silent false "match"; it is reported, not hidden, so a human can review it.
    """
    flags: list[str] = []

    if v2.outcome_kind != v3.outcome_kind:
        flags.append(MISMATCH_OUTCOME_KIND_DIFFERS)

    if v2.clarification_asked != v3.clarification_asked:
        flags.append(MISMATCH_CLARIFICATION_DIFFERS)

    if v2.financial and v3.financial and v2.value_display is not None and v3.value_display is not None:
        if v2.value_display != v3.value_display:
            flags.append(MISMATCH_TOTALS_DIFFER)

    if v2.row_count is not None and v3.row_count is not None and v2.row_count != v3.row_count:
        flags.append(MISMATCH_ROW_COUNT_DIFFERS)

    if v2.row_set_hash is not None and v3.row_set_hash is not None and v2.row_set_hash != v3.row_set_hash:
        flags.append(MISMATCH_ROW_SET_DIFFERS)

    if v2.outcome_kind == "tool" and v3.outcome_kind == "tool" and _scope_differs(v2, v3):
        flags.append(MISMATCH_SCOPE_DIFFERS)

    return ShadowComparison(
        timestamp=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        question=question, primary_mode=primary_mode, editor=editor, v2=v2, v3=v3, mismatch_flags=flags,
    )
