"""Tests for snapshot_invoice_history.py: the standalone, scheduled daily invoice snapshot writer.

Ask Data itself never calls this script; these tests verify its own guarantees in isolation (idempotency,
read-only against invoices/projects, auditable run-log, failure handling) using an in-memory fake instead
of a real Supabase connection.
"""
import datetime
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from models.invoice import Invoice
from models.project import Project
from snapshot_invoice_history import build_snapshot_rows, run


def make_invoices():
    return [
        Invoice(invoice_number=100, project_name="AD Anderlecht", maintenance_year="Y1", payment_amount=500.0,
                cameras_number=4, payment_date=None, paid="No", year=2026, db_id=1),
        Invoice(invoice_number=101, project_name="AD Zwevegem", maintenance_year="Y1", payment_amount=300.0,
                cameras_number=4, payment_date=datetime.datetime(2026, 10, 10), paid="Yes", year=2026, db_id=2),
        # Never-saved row (no db_id) -- must be skipped, it has no stable identity to snapshot.
        Invoice(invoice_number=None, project_name="Unsaved Draft", maintenance_year="Y1", payment_amount=50.0,
                cameras_number=1, payment_date=None, paid="No", year=2026, db_id=None),
    ]


def make_projects():
    return [Project(project_name="AD Anderlecht", country="Belgium"), Project(project_name="AD Zwevegem", country="Belgium")]


class TestBuildSnapshotRows(unittest.TestCase):
    """Pure function: no I/O, directly testable."""

    CAPTURED_AT = "2026-10-20T20:00:00+00:00"

    def test_one_row_per_saved_invoice_with_resolved_country_and_network(self):
        rows = build_snapshot_rows(make_invoices(), make_projects(), self.CAPTURED_AT)
        self.assertEqual(len(rows), 2)  # the unsaved draft (db_id=None) is skipped
        by_id = {r["invoice_row_id"]: r for r in rows}
        self.assertEqual(by_id[1]["country_at_capture"], "Belgium")
        self.assertEqual(by_id[1]["network_at_capture"], "AD")
        self.assertEqual(by_id[1]["payment_amount"], 500.0)
        self.assertEqual(by_id[1]["paid"], "No")
        self.assertEqual(by_id[2]["paid"], "Yes")
        self.assertEqual(by_id[2]["payment_date"], "2026-10-10")

    def test_every_row_gets_the_same_explicit_captured_at(self):
        rows = build_snapshot_rows(make_invoices(), make_projects(), self.CAPTURED_AT)
        self.assertTrue(rows)
        self.assertEqual({r["captured_at"] for r in rows}, {self.CAPTURED_AT})

    def test_never_mutates_the_input_objects(self):
        invoices = make_invoices()
        before = [(i.payment_amount, i.paid, i.project_name) for i in invoices]
        build_snapshot_rows(invoices, make_projects(), self.CAPTURED_AT)
        after = [(i.payment_amount, i.paid, i.project_name) for i in invoices]
        self.assertEqual(before, after)


class _FakeSnapshotStore:
    """In-memory fake standing in for the invoice_snapshots/invoice_snapshot_runs Supabase tables."""

    def __init__(self):
        self.snapshot_rows: dict = {}   # (snapshot_date, invoice_row_id) -> row
        self.runs: dict = {}            # snapshot_date -> run dict
        self.write_calls = 0

    def upsert_rows(self, snapshot_date, rows):
        self.write_calls += 1
        for row in rows:
            key = (snapshot_date.isoformat(), row["invoice_row_id"])
            self.snapshot_rows[key] = dict(row, snapshot_date=snapshot_date.isoformat())
        return len(rows)

    def upsert_run(self, snapshot_date, **kwargs):
        key = snapshot_date.isoformat()
        existing = self.runs.get(key, {})
        existing.update({k: v for k, v in kwargs.items() if v is not None})
        self.runs[key] = existing


class TestSnapshotWriterIdempotency(unittest.TestCase):
    def setUp(self):
        self.store = _FakeSnapshotStore()
        self.invoices = make_invoices()
        self.projects = make_projects()
        self.day = datetime.date(2026, 10, 20)

    def _run(self):
        with mock.patch("services.supabase_service.load_invoices", lambda: self.invoices), \
             mock.patch("services.supabase_service.load_projects", lambda: self.projects), \
             mock.patch("services.supabase_service.upsert_invoice_snapshot_rows", self.store.upsert_rows), \
             mock.patch("services.supabase_service.upsert_invoice_snapshot_run", self.store.upsert_run):
            return run(self.day)

    def test_first_run_writes_exactly_the_saved_invoice_rows(self):
        exit_code = self._run()
        self.assertEqual(exit_code, 0)
        self.assertEqual(len(self.store.snapshot_rows), 2)
        self.assertEqual(self.store.runs[self.day.isoformat()]["status"], "success")
        self.assertEqual(self.store.runs[self.day.isoformat()]["row_count"], 2)

    def test_rerun_same_day_overwrites_rather_than_duplicates(self):
        self._run()
        first_write_calls = self.store.write_calls
        self._run()
        self.assertGreater(self.store.write_calls, first_write_calls)   # it DID write again (reran)...
        self.assertEqual(len(self.store.snapshot_rows), 2)              # ...but row count never duplicated

    def test_rerun_refreshes_captured_at_on_every_row_identically(self):
        """Regression test for the idempotency defect found in production Step 3 verification: a same-day
        rerun must overwrite captured_at on every existing row with the new run's timestamp, never leave
        the original value in place."""
        self._run()
        first_captured = {row["captured_at"] for row in self.store.snapshot_rows.values()}
        self.assertEqual(len(first_captured), 1)   # single run -> single shared timestamp
        first_value = next(iter(first_captured))

        self._run()   # same day, rerun
        second_captured = {row["captured_at"] for row in self.store.snapshot_rows.values()}
        self.assertEqual(len(second_captured), 1)                      # still a single shared timestamp...
        second_value = next(iter(second_captured))
        self.assertNotEqual(second_value, first_value)                  # ...but refreshed, not left stale
        self.assertGreater(second_value, first_value)                   # and strictly newer

        # Everything else about the rerun stays exactly as before: no duplicates, same snapshot_date,
        # unchanged row count, unchanged financial totals.
        self.assertEqual(len(self.store.snapshot_rows), 2)
        self.assertEqual({k[0] for k in self.store.snapshot_rows}, {self.day.isoformat()})
        total = sum(row["payment_amount"] for row in self.store.snapshot_rows.values() if row["paid"] == "No")
        self.assertEqual(total, 500.0)   # only invoice 100 is unpaid in this fixture; unchanged by the rerun

    def test_rerun_after_a_ledger_change_reflects_the_latest_state(self):
        self._run()
        self.invoices[1].paid = "cancelled"   # invoice 101 gets cancelled after the first run
        self._run()
        row = self.store.snapshot_rows[(self.day.isoformat(), 2)]
        self.assertEqual(row["paid"], "cancelled")

    def test_dry_run_writes_nothing(self):
        with mock.patch("services.supabase_service.load_invoices", lambda: self.invoices), \
             mock.patch("services.supabase_service.load_projects", lambda: self.projects), \
             mock.patch("services.supabase_service.upsert_invoice_snapshot_rows", self.store.upsert_rows), \
             mock.patch("services.supabase_service.upsert_invoice_snapshot_run", self.store.upsert_run):
            exit_code = run(self.day, dry_run=True)
        self.assertEqual(exit_code, 0)
        self.assertEqual(self.store.snapshot_rows, {})
        self.assertEqual(self.store.runs, {})

    def test_never_calls_any_write_function_on_invoices_or_projects(self):
        """The job is read-only against the live ledger: load_invoices/load_projects are called, but no
        save/upsert/delete function for `invoices` or `projects` is ever imported or invoked."""
        import ast
        source = Path("snapshot_invoice_history.py").read_text(encoding="utf-8")
        tree = ast.parse(source)
        banned = {"upsert_invoices", "upsert_projects", "delete_invoice", "save_invoices"}
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom) and node.module == "services.supabase_service":
                for alias in node.names:
                    self.assertNotIn(alias.name, banned)

    def test_failure_while_building_rows_marks_the_run_failed_and_returns_nonzero(self):
        def _boom():
            raise RuntimeError("supabase unavailable")
        with mock.patch("services.supabase_service.load_invoices", _boom), \
             mock.patch("services.supabase_service.load_projects", lambda: self.projects), \
             mock.patch("services.supabase_service.upsert_invoice_snapshot_rows", self.store.upsert_rows), \
             mock.patch("services.supabase_service.upsert_invoice_snapshot_run", self.store.upsert_run):
            exit_code = run(self.day)
        self.assertEqual(exit_code, 1)
        self.assertEqual(self.store.runs[self.day.isoformat()]["status"], "failed")
        self.assertIn("supabase unavailable", self.store.runs[self.day.isoformat()]["error_message"])

    def test_failure_while_writing_rows_marks_the_run_failed(self):
        def _boom(*a, **k):
            raise RuntimeError("write failed")
        with mock.patch("services.supabase_service.load_invoices", lambda: self.invoices), \
             mock.patch("services.supabase_service.load_projects", lambda: self.projects), \
             mock.patch("services.supabase_service.upsert_invoice_snapshot_rows", _boom), \
             mock.patch("services.supabase_service.upsert_invoice_snapshot_run", self.store.upsert_run):
            exit_code = run(self.day)
        self.assertEqual(exit_code, 1)
        self.assertEqual(self.store.runs[self.day.isoformat()]["status"], "failed")


if __name__ == "__main__":
    unittest.main()
