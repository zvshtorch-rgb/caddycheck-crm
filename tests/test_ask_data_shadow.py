"""Ask Data v2/v3 shadow comparison: summarization, mismatch detection (goal 6), and the privacy guarantee
that no ledger row, invoice number, amount, bank detail, or project name ever ends up in the stored log."""
import unittest
from decimal import Decimal

import pandas as pd

from services.ask_data.shadow import (
    MISMATCH_CLARIFICATION_DIFFERS, MISMATCH_OUTCOME_KIND_DIFFERS, MISMATCH_ROW_COUNT_DIFFERS,
    MISMATCH_ROW_SET_DIFFERS, MISMATCH_SCOPE_DIFFERS, MISMATCH_TOTALS_DIFFER, ShadowSummary, build_comparison,
)
from services.ask_data.tool_result import ToolResult


def _df(rows):
    return pd.DataFrame(rows) if rows else None


def _result(tool="get_debt", *, value=Decimal("100.00"), row_count=2, rows_payload=None, filters=None,
            financial=True) -> ToolResult:
    rows_payload = rows_payload if rows_payload is not None else [{"Invoice #": "1", "Amount (€)": "50.00"},
                                                                   {"Invoice #": "2", "Amount (€)": "50.00"}]
    return ToolResult(
        tool=tool, text="irrelevant", dataframe=_df(rows_payload), financial=financial, value=value,
        row_count=row_count, filters=filters or {"country": "Bel"},
    )


class TestShadowSummary(unittest.TestCase):
    def test_from_tool_result_keeps_only_whitelisted_filter_keys(self):
        result = _result(filters={"country": "Bel", "project": "Edeka Herweg", "limit": 5, "invoice_number": 8001})
        summary = ShadowSummary.from_tool_result(result)
        self.assertEqual(summary.filters, {"country": "Bel"})          # project/limit/invoice_number dropped
        self.assertTrue(summary.project_present)
        self.assertIsNotNone(summary.project_hash)
        self.assertNotIn("Edeka", summary.project_hash)                # hash, never the name itself

    def test_from_tool_result_financial_value_and_row_count(self):
        summary = ShadowSummary.from_tool_result(_result(value=Decimal("19999.00"), row_count=8))
        self.assertTrue(summary.financial)
        self.assertEqual(summary.value_display, "\u20ac19,999.00")
        self.assertEqual(summary.row_count, 8)
        self.assertIsNotNone(summary.row_set_hash)

    def test_from_tool_result_non_financial_has_no_value_display(self):
        summary = ShadowSummary.from_tool_result(_result(financial=False, value=None))
        self.assertFalse(summary.financial)
        self.assertIsNone(summary.value_display)

    def test_from_tool_result_none_is_outcome_none(self):
        summary = ShadowSummary.from_tool_result(None)
        self.assertEqual(summary.outcome_kind, "none")
        self.assertIsNone(summary.row_set_hash)

    def test_from_clarification_legacy_error(self):
        self.assertEqual(ShadowSummary.from_clarification().outcome_kind, "clarification")
        self.assertTrue(ShadowSummary.from_clarification().clarification_asked)
        self.assertEqual(ShadowSummary.from_legacy().outcome_kind, "legacy")
        self.assertEqual(ShadowSummary.from_error().outcome_kind, "error")

    def test_from_historical_unsupported_is_distinct_from_error(self):
        summary = ShadowSummary.from_historical_unsupported()
        self.assertEqual(summary.outcome_kind, "historical_unsupported")
        self.assertNotEqual(summary.outcome_kind, ShadowSummary.from_error().outcome_kind)
        self.assertFalse(summary.financial)
        self.assertIsNone(summary.row_count)

    def test_row_set_hash_is_order_independent_but_content_sensitive(self):
        a = ShadowSummary.from_tool_result(_result(rows_payload=[{"x": 1}, {"x": 2}]))
        b = ShadowSummary.from_tool_result(_result(rows_payload=[{"x": 2}, {"x": 1}]))
        c = ShadowSummary.from_tool_result(_result(rows_payload=[{"x": 1}, {"x": 3}]))
        self.assertEqual(a.row_set_hash, b.row_set_hash)
        self.assertNotEqual(a.row_set_hash, c.row_set_hash)


class TestBuildComparison(unittest.TestCase):
    def test_identical_results_have_no_mismatch(self):
        v2 = ShadowSummary.from_tool_result(_result())
        v3 = ShadowSummary.from_tool_result(_result())
        comparison = build_comparison("How much debt?", "v3", v2, v3)
        self.assertEqual(comparison.mismatch_flags, [])
        self.assertFalse(comparison.has_mismatch)

    def test_totals_differ_is_flagged(self):
        v2 = ShadowSummary.from_tool_result(_result(value=Decimal("100.00")))
        v3 = ShadowSummary.from_tool_result(_result(value=Decimal("150.00")))
        comparison = build_comparison("q", "v3", v2, v3)
        self.assertIn(MISMATCH_TOTALS_DIFFER, comparison.mismatch_flags)

    def test_row_count_differs_is_flagged(self):
        v2 = ShadowSummary.from_tool_result(_result(row_count=2))
        v3 = ShadowSummary.from_tool_result(_result(row_count=3))
        comparison = build_comparison("q", "v3", v2, v3)
        self.assertIn(MISMATCH_ROW_COUNT_DIFFERS, comparison.mismatch_flags)

    def test_row_set_differs_with_same_count_is_flagged(self):
        v2 = ShadowSummary.from_tool_result(_result(rows_payload=[{"x": 1}, {"x": 2}]))
        v3 = ShadowSummary.from_tool_result(_result(rows_payload=[{"x": 1}, {"x": 9}]))
        comparison = build_comparison("q", "v3", v2, v3)
        self.assertIn(MISMATCH_ROW_SET_DIFFERS, comparison.mismatch_flags)
        self.assertNotIn(MISMATCH_ROW_COUNT_DIFFERS, comparison.mismatch_flags)   # same count, only content differs

    def test_scope_differs_by_country_is_flagged(self):
        v2 = ShadowSummary.from_tool_result(_result(filters={"country": "Bel"}))
        v3 = ShadowSummary.from_tool_result(_result(filters={"country": "Ger"}))
        comparison = build_comparison("q", "v3", v2, v3)
        self.assertIn(MISMATCH_SCOPE_DIFFERS, comparison.mismatch_flags)

    def test_scope_differs_by_project_is_flagged_without_revealing_names(self):
        v2 = ShadowSummary.from_tool_result(_result(filters={"project": "AD Alpha"}))
        v3 = ShadowSummary.from_tool_result(_result(filters={"project": "AD Beta"}))
        comparison = build_comparison("q", "v3", v2, v3)
        self.assertIn(MISMATCH_SCOPE_DIFFERS, comparison.mismatch_flags)
        entry = comparison.to_log_entry()
        for value in entry.values():
            self.assertNotIn("AD Alpha", str(value))
            self.assertNotIn("AD Beta", str(value))

    def test_same_project_is_not_a_scope_mismatch(self):
        v2 = ShadowSummary.from_tool_result(_result(filters={"project": "AD Alpha"}))
        v3 = ShadowSummary.from_tool_result(_result(filters={"project": "AD Alpha"}))
        comparison = build_comparison("q", "v3", v2, v3)
        self.assertNotIn(MISMATCH_SCOPE_DIFFERS, comparison.mismatch_flags)

    def test_clarification_vs_tool_is_flagged_as_both_clarification_and_outcome_kind(self):
        v2 = ShadowSummary.from_clarification()
        v3 = ShadowSummary.from_tool_result(_result())
        comparison = build_comparison("Debt for Edeka", "v3", v2, v3)
        self.assertIn(MISMATCH_CLARIFICATION_DIFFERS, comparison.mismatch_flags)
        self.assertIn(MISMATCH_OUTCOME_KIND_DIFFERS, comparison.mismatch_flags)
        # scope/totals are not meaningfully comparable when one side never resolved to a tool
        self.assertNotIn(MISMATCH_SCOPE_DIFFERS, comparison.mismatch_flags)
        self.assertNotIn(MISMATCH_TOTALS_DIFFER, comparison.mismatch_flags)

    def test_followup_where_v2_has_no_context_is_an_outcome_kind_mismatch_not_hidden(self):
        """'Only Belgium?' alone: v2 (no context) falls back to legacy/clarification; v3 (with context)
        resolves it via get_debt. This is EXPECTED (goal 7) but must still be visible in the log, not silently
        treated as a match."""
        v2 = ShadowSummary.from_legacy()
        v3 = ShadowSummary.from_tool_result(_result(filters={"country": "Bel"}))
        comparison = build_comparison("Only Belgium?", "v3", v2, v3)
        self.assertIn(MISMATCH_OUTCOME_KIND_DIFFERS, comparison.mismatch_flags)

    def test_drilldown_outcome_kind_is_v3_only_and_flagged(self):
        v2 = ShadowSummary.from_legacy()
        v3 = ShadowSummary(outcome_kind="drilldown", row_count=5)
        comparison = build_comparison("Which invoices?", "v3", v2, v3)
        self.assertIn(MISMATCH_OUTCOME_KIND_DIFFERS, comparison.mismatch_flags)

    def test_to_log_entry_never_contains_raw_row_or_dataframe_objects(self):
        v2 = ShadowSummary.from_tool_result(_result())
        v3 = ShadowSummary.from_tool_result(_result())
        entry = build_comparison("q", "v2", v2, v3).to_log_entry()
        for forbidden_key in ("rows", "dataframe", "v2_rows", "v3_rows", "api_key", "gemini", "question"):
            self.assertNotIn(forbidden_key, entry)
        self.assertIn("question_hash", entry)
        self.assertIn("question_length", entry)
        self.assertIn("question_word_count", entry)
        self.assertIn("mismatch_flags", entry)
        self.assertIn("primary_mode", entry)

    def test_question_text_is_never_persisted_only_a_hash_and_counts(self):
        v2 = ShadowSummary.from_tool_result(_result())
        v3 = ShadowSummary.from_tool_result(_result())
        secret_question = "What about invoice 8713 for Edeka Herweg, SWIFT ref 8382-CN?"
        entry = build_comparison(secret_question, "v2", v2, v3).to_log_entry()
        for value in entry.values():
            self.assertNotIn("8713", str(value))
            self.assertNotIn("Edeka", str(value))
            self.assertNotIn("8382-CN", str(value))
            self.assertNotIn(secret_question, str(value))
        self.assertEqual(entry["question_length"], len(secret_question))
        self.assertEqual(entry["question_word_count"], len(secret_question.split()))
        self.assertEqual(len(entry["question_hash"]), 16)

    def test_question_hash_is_deterministic_and_distinguishes_different_questions(self):
        v2 = ShadowSummary.from_tool_result(_result())
        v3 = ShadowSummary.from_tool_result(_result())
        e1 = build_comparison("How much debt?", "v2", v2, v3).to_log_entry()
        e2 = build_comparison("How much debt?", "v2", v2, v3).to_log_entry()
        e3 = build_comparison("Only Belgium?", "v2", v2, v3).to_log_entry()
        self.assertEqual(e1["question_hash"], e2["question_hash"])
        self.assertNotEqual(e1["question_hash"], e3["question_hash"])


if __name__ == "__main__":
    unittest.main()
