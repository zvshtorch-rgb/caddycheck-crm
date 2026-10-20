"""Tests for Ask Data Phase 5's historical/trend tools (get_debt_at, compare_debt_dates, get_debt_change,
get_invoice_changes, get_payments_in_period, get_new_invoices_in_period) via execute_tool(), with Supabase
I/O mocked out (services/ask_data/history.py itself is pure and has no I/O to mock).
"""
import unittest
from datetime import date, datetime
from decimal import Decimal
from unittest import mock

from models.invoice import Invoice
from models.project import Project
from services.ask_data import data_clock
from services.ask_data.executor import execute_tool
from services.ask_data.history_dates import resolve_period_phrase

PROJECTS = [
    Project(project_name="AD Anderlecht", country="Belgium"),
    Project(project_name="AD Zwevegem", country="Belgium"),
    Project(project_name="Edeka Herweg", country="Germany"),
]


def make_invoices():
    return [
        Invoice(invoice_number=100, project_name="AD Anderlecht", maintenance_year="Y1", payment_amount=500.0,
                cameras_number=4, payment_date=None, paid="No", year=2026, db_id=1),
        Invoice(invoice_number=101, project_name="AD Zwevegem", maintenance_year="Y1", payment_amount=300.0,
                cameras_number=4, payment_date=date(2026, 10, 10), paid="Yes", year=2026, db_id=2),
        Invoice(invoice_number=102, project_name="AD Anderlecht", maintenance_year="Y2", payment_amount=200.0,
                cameras_number=4, payment_date=None, paid="No", year=2026, db_id=3),
        Invoice(invoice_number=103, project_name="Edeka Herweg", maintenance_year="Y1", payment_amount=150.0,
                cameras_number=4, payment_date=None, paid="No", year=2026, db_id=4),
    ]


RUNS = [
    {"snapshot_date": "2026-10-01", "status": "success", "started_at": "2026-10-01T02:00:00+00:00",
     "completed_at": "2026-10-01T02:01:00+00:00", "row_count": 3},
    {"snapshot_date": "2026-10-15", "status": "success", "started_at": "2026-10-15T02:00:00+00:00",
     "completed_at": "2026-10-15T02:01:00+00:00", "row_count": 4},
    {"snapshot_date": "2026-10-20", "status": "success", "started_at": "2026-10-20T02:00:00+00:00",
     "completed_at": "2026-10-20T02:01:00+00:00", "row_count": 4},
]

SNAPSHOTS = {
    "2026-10-01": [
        {"invoice_row_id": 1, "invoice_number": 100, "project_name": "AD Anderlecht", "payment_amount": 500.0,
         "paid": "No", "year": 2026, "maintenance_year": "Y1"},
        {"invoice_row_id": 2, "invoice_number": 101, "project_name": "AD Zwevegem", "payment_amount": 300.0,
         "paid": "No", "year": 2026, "maintenance_year": "Y1"},
        {"invoice_row_id": 4, "invoice_number": 103, "project_name": "Edeka Herweg", "payment_amount": 150.0,
         "paid": "No", "year": 2026, "maintenance_year": "Y1"},
    ],
    "2026-10-15": [
        {"invoice_row_id": 1, "invoice_number": 100, "project_name": "AD Anderlecht", "payment_amount": 500.0,
         "paid": "No", "year": 2026, "maintenance_year": "Y1"},
        {"invoice_row_id": 2, "invoice_number": 101, "project_name": "AD Zwevegem", "payment_amount": 300.0,
         "paid": "Yes", "year": 2026, "maintenance_year": "Y1"},
        {"invoice_row_id": 3, "invoice_number": 102, "project_name": "AD Anderlecht", "payment_amount": 200.0,
         "paid": "No", "year": 2026, "maintenance_year": "Y2"},
        {"invoice_row_id": 4, "invoice_number": 103, "project_name": "Edeka Herweg", "payment_amount": 150.0,
         "paid": "No", "year": 2026, "maintenance_year": "Y1"},
    ],
    "2026-10-20": [
        {"invoice_row_id": 1, "invoice_number": 100, "project_name": "AD Anderlecht", "payment_amount": 500.0,
         "paid": "No", "year": 2026, "maintenance_year": "Y1"},
        {"invoice_row_id": 2, "invoice_number": 101, "project_name": "AD Zwevegem", "payment_amount": 300.0,
         "paid": "Yes", "year": 2026, "maintenance_year": "Y1"},
        {"invoice_row_id": 3, "invoice_number": 102, "project_name": "AD Anderlecht", "payment_amount": 200.0,
         "paid": "No", "year": 2026, "maintenance_year": "Y2"},
        {"invoice_row_id": 4, "invoice_number": 103, "project_name": "Edeka Herweg", "payment_amount": 150.0,
         "paid": "No", "year": 2026, "maintenance_year": "Y1"},
    ],
}


def _load_day(d):
    key = d.isoformat() if hasattr(d, "isoformat") else str(d)
    return SNAPSHOTS.get(key, [])


def _patches(runs=None, bank_payments=None, allocations=None, change_log=None, sent_invoices=None):
    return [
        mock.patch("services.supabase_service.load_invoice_snapshot_runs", lambda: runs if runs is not None else RUNS),
        mock.patch("services.supabase_service.load_invoice_snapshot", _load_day),
        mock.patch("services.supabase_service.load_invoice_change_log", lambda *a, **k: change_log or []),
        mock.patch("services.supabase_service.load_bank_payments", lambda: bank_payments or []),
        mock.patch("services.supabase_service.load_bank_payment_allocations", lambda *a, **k: allocations or []),
        mock.patch("config.settings.load_sent_invoices_log", lambda: sent_invoices or []),
    ]


class _PatchedTestCase(unittest.TestCase):
    def setUp(self):
        data_clock.mark_loaded("test")
        self.invoices = make_invoices()
        self.projects = PROJECTS

    def _exec(self, tool, args, runs=None, bank_payments=None, allocations=None, change_log=None, sent_invoices=None):
        patches = _patches(runs, bank_payments, allocations, change_log, sent_invoices)
        with patches[0], patches[1], patches[2], patches[3], patches[4], patches[5]:
            return execute_tool(tool, args, self.projects, self.invoices, [])


class TestGetDebtAt(_PatchedTestCase):
    def test_exact_snapshot_date(self):
        result = self._exec("get_debt_at", {"at_date": date(2026, 10, 1)})
        self.assertEqual(result.value, Decimal("950.00"))  # 500 + 300 + 150, all unpaid on 10/1
        self.assertFalse(result.warnings)  # exact match -> no coverage-note warning

    def test_nearest_earlier_snapshot_with_required_wording(self):
        # No snapshot for 10-18 -> falls back to 10-15, with the EXACT required wording.
        result = self._exec("get_debt_at", {"at_date": date(2026, 10, 18)})
        self.assertEqual(result.value, Decimal("850.00"))  # 500 + 200 + 150, unpaid on 10/15 (101 already paid)
        self.assertTrue(any(
            w.startswith("No snapshot exists for 2026-10-18; using the latest available snapshot from")
            for w in result.warnings
        ))

    def test_before_history_start_is_explicit_not_silent(self):
        result = self._exec("get_debt_at", {"at_date": date(2026, 9, 1)})
        self.assertFalse(result.financial)
        self.assertIn("History is available from 2026-10-01", result.text)

    def test_country_filter_reconciles(self):
        belgium = self._exec("get_debt_at", {"at_date": date(2026, 10, 1), "country": "Belgium"})
        germany = self._exec("get_debt_at", {"at_date": date(2026, 10, 1), "country": "Germany"})
        total = self._exec("get_debt_at", {"at_date": date(2026, 10, 1)})
        self.assertEqual(belgium.value + germany.value, total.value)

    def test_network_filter(self):
        result = self._exec("get_debt_at", {"at_date": date(2026, 10, 1), "network": "AD"})
        self.assertEqual(result.value, Decimal("800.00"))  # 500 + 300, both AD-network rows unpaid on 10/1

    def test_value_is_decimal_not_float(self):
        result = self._exec("get_debt_at", {"at_date": date(2026, 10, 1)})
        self.assertIsInstance(result.value, Decimal)

    def test_states_the_actual_snapshot_timestamp_used(self):
        result = self._exec("get_debt_at", {"at_date": date(2026, 10, 1)})
        self.assertIn("captured 2026-10-01 02:01 UTC", result.text)

    def test_grounded(self):
        result = self._exec("get_debt_at", {"at_date": date(2026, 10, 1)})
        self.assertEqual(result.ungrounded_amounts(), [])


class TestCompareDebtDates(_PatchedTestCase):
    def test_compares_two_dates(self):
        result = self._exec("compare_debt_dates", {"date_a": date(2026, 10, 1), "date_b": date(2026, 10, 20)})
        self.assertEqual(result.value, Decimal("-100.00"))  # 850 - 950
        self.assertIn("a decrease of", result.text)

    def test_states_both_snapshot_dates_used(self):
        result = self._exec("compare_debt_dates", {"date_a": date(2026, 10, 1), "date_b": date(2026, 10, 20)})
        self.assertIn("snapshot date 2026-10-01", result.text)
        self.assertIn("snapshot date 2026-10-20", result.text)

    def test_grounded(self):
        result = self._exec("compare_debt_dates", {"date_a": date(2026, 10, 1), "date_b": date(2026, 10, 20)})
        self.assertEqual(result.ungrounded_amounts(), [])


class TestGetDebtChange(_PatchedTestCase):
    def test_bridge_reconciles_exactly(self):
        period = resolve_period_phrase("this month", date(2026, 10, 20))
        result = self._exec("get_debt_change", {"period": period})
        self.assertEqual(result.value, Decimal("-100.00"))  # live closing 850 - opening 950
        self.assertNotIn("Unexplained adjustment", result.text)

    def test_components_sum_to_the_closing_minus_opening_change(self):
        period = resolve_period_phrase("this month", date(2026, 10, 20))
        result = self._exec("get_debt_change", {"period": period})
        opening = next(c for c in result.components if c.label == "Opening").amount
        closing = next(c for c in result.components if c.label == "Closing").amount
        self.assertEqual(closing - opening, result.value)

    def test_non_reconciling_bridge_shows_explicit_unexplained_warning(self):
        """A live ledger edit happens AFTER the last snapshot was captured (so no snapshot-to-snapshot row
        diff can classify it): for a period ending today, `bridge.closing` uses the LIVE total, which can
        then diverge from what the snapshot-based row classification alone would predict. That divergence
        must surface as an explicit 'Unexplained adjustment', never be silently hidden or absorbed."""
        real_today = date.today()
        runs_incl_today = RUNS + [{"snapshot_date": real_today.isoformat(), "status": "success",
                                    "started_at": real_today.isoformat() + "T02:00:00+00:00",
                                    "completed_at": real_today.isoformat() + "T02:01:00+00:00", "row_count": 4}]
        snapshots_incl_today = dict(SNAPSHOTS)
        snapshots_incl_today[real_today.isoformat()] = SNAPSHOTS["2026-10-20"]  # stale: unaware of the live edit below
        live_invoices = make_invoices()
        live_invoices[0].payment_amount = 999.0  # edited live, same day, after the last snapshot was taken
        period = resolve_period_phrase("since October 1", real_today)
        with mock.patch("services.supabase_service.load_invoice_snapshot_runs", lambda: runs_incl_today), \
             mock.patch("services.supabase_service.load_invoice_snapshot", lambda d: snapshots_incl_today.get(d.isoformat(), [])), \
             mock.patch("services.supabase_service.load_invoice_change_log", lambda *a, **k: []), \
             mock.patch("services.supabase_service.load_bank_payments", lambda: []), \
             mock.patch("services.supabase_service.load_bank_payment_allocations", lambda *a, **k: []):
            result = execute_tool("get_debt_change", {"period": period}, self.projects, live_invoices, [])
        self.assertTrue(any("Unexplained adjustment" in w for w in result.warnings))
        self.assertIn("Unexplained adjustment: " + "499.00", result.display_text().replace("€", ""))

    def test_closing_uses_live_current_data_when_period_ends_today(self):
        real_today = date.today()
        runs_incl_today = RUNS + [{"snapshot_date": real_today.isoformat(), "status": "success",
                                    "started_at": real_today.isoformat() + "T02:00:00+00:00",
                                    "completed_at": real_today.isoformat() + "T02:01:00+00:00", "row_count": 4}]
        snapshots_incl_today = dict(SNAPSHOTS)
        snapshots_incl_today[real_today.isoformat()] = SNAPSHOTS["2026-10-20"]
        period = resolve_period_phrase("today", real_today)  # date_from == date_to == real today
        with mock.patch("services.supabase_service.load_invoice_snapshot_runs", lambda: runs_incl_today), \
             mock.patch("services.supabase_service.load_invoice_snapshot", lambda d: snapshots_incl_today.get(d.isoformat(), [])), \
             mock.patch("services.supabase_service.load_invoice_change_log", lambda *a, **k: []), \
             mock.patch("services.supabase_service.load_bank_payments", lambda: []), \
             mock.patch("services.supabase_service.load_bank_payment_allocations", lambda *a, **k: []):
            result = execute_tool("get_debt_change", {"period": period}, self.projects, self.invoices, [])
        self.assertIn("live current data", result.text)
        # Live closing = sum of unpaid `self.invoices` (850.00), independent of whatever the "today" snapshot says.
        closing = next(c for c in result.components if c.label == "Closing").amount
        self.assertEqual(closing, Decimal("850.00"))

    def test_no_gemini_call_needed_value_is_grounded(self):
        period = resolve_period_phrase("this month", date(2026, 10, 20))
        result = self._exec("get_debt_change", {"period": period})
        self.assertEqual(result.ungrounded_amounts(), [])


class TestGetInvoiceChanges(_PatchedTestCase):
    def test_lists_classified_changes(self):
        period = resolve_period_phrase("this month", date(2026, 10, 20))
        result = self._exec("get_invoice_changes", {"period": period})
        actions = set(result.dataframe["Action"])
        self.assertIn("New", actions)
        self.assertIn("Payment", actions)

    def test_no_changes_found_is_explicit_not_hidden(self):
        with mock.patch("services.supabase_service.load_invoice_snapshot_runs", lambda: RUNS), \
             mock.patch("services.supabase_service.load_invoice_snapshot", lambda d: SNAPSHOTS["2026-10-15"]), \
             mock.patch("services.supabase_service.load_invoice_change_log", lambda *a, **k: []):
            period = resolve_period_phrase("since October 15", date(2026, 10, 15))
            result = execute_tool("get_invoice_changes", {"period": period}, self.projects, self.invoices, [])
        self.assertIn("No invoice changes", result.text)


class TestGetPaymentsInPeriod(_PatchedTestCase):
    def test_sums_bank_log_allocations_in_range(self):
        bank_payments = [{"id": 1, "payment_date": "2026-10-10", "applied_amount": 300.0}]
        allocations = [{"payment_id": 1, "invoice_number": 101, "project_name": "AD Zwevegem",
                        "amount_applied": 300.0, "year": 2026}]
        period = resolve_period_phrase("this month", date(2026, 10, 20))
        result = self._exec("get_payments_in_period", {"period": period}, bank_payments=bank_payments, allocations=allocations)
        self.assertEqual(result.value, Decimal("300.00"))

    def test_mismatch_against_snapshot_diff_is_flagged_not_hidden(self):
        """Bank log says 0 but the snapshot diff shows a row flipped to paid -- must be surfaced, not silently picked."""
        period = resolve_period_phrase("this month", date(2026, 10, 20))
        result = self._exec("get_payments_in_period", {"period": period})  # no bank_payments supplied -> 0
        self.assertTrue(any("cross-check" in w for w in result.warnings))

    def test_grounded(self):
        bank_payments = [{"id": 1, "payment_date": "2026-10-10", "applied_amount": 300.0}]
        allocations = [{"payment_id": 1, "invoice_number": 101, "project_name": "AD Zwevegem",
                        "amount_applied": 300.0, "year": 2026}]
        period = resolve_period_phrase("this month", date(2026, 10, 20))
        result = self._exec("get_payments_in_period", {"period": period}, bank_payments=bank_payments, allocations=allocations)
        self.assertEqual(result.ungrounded_amounts(), [])


class TestGetNewInvoicesInPeriod(_PatchedTestCase):
    def test_lists_new_invoice_rows(self):
        period = resolve_period_phrase("this month", date(2026, 10, 20))
        result = self._exec("get_new_invoices_in_period", {"period": period})
        self.assertEqual(result.value, Decimal("200.00"))
        self.assertEqual(list(result.dataframe["Invoice #"]), [102])

    def test_no_new_invoices_is_explicit(self):
        with mock.patch("services.supabase_service.load_invoice_snapshot_runs", lambda: RUNS), \
             mock.patch("services.supabase_service.load_invoice_snapshot", lambda d: SNAPSHOTS["2026-10-15"]), \
             mock.patch("config.settings.load_sent_invoices_log", lambda: []):
            period = resolve_period_phrase("since October 15", date(2026, 10, 15))
            result = execute_tool("get_new_invoices_in_period", {"period": period}, self.projects, self.invoices, [])
        self.assertIn("No new invoices", result.text)


class TestHistoryFlagOffPromptUnchanged(unittest.TestCase):
    """The single most important Phase 5 guarantee: with the flag off (the default), the planner prompt is
    byte-identical to pre-Phase-5 -- structurally guaranteed by `history_enabled` defaulting to False, not
    merely tested, but pinned here as a regression guard."""

    def test_catalog_text_excludes_history_tools_by_default(self):
        from services.ask_data.tool_specs import _ask_data_tool_catalog_text, HISTORY_TOOL_NAMES
        text = _ask_data_tool_catalog_text()
        for tool in HISTORY_TOOL_NAMES:
            self.assertNotIn(tool, text)

    def test_catalog_text_includes_history_tools_when_requested(self):
        from services.ask_data.tool_specs import _ask_data_tool_catalog_text, HISTORY_TOOL_NAMES
        text = _ask_data_tool_catalog_text(include_history=True)
        for tool in HISTORY_TOOL_NAMES:
            self.assertIn(tool, text)


if __name__ == "__main__":
    unittest.main()
