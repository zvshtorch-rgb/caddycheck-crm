"""Tests for services/bank_payment_save.py (the Save All Parsed Payments flow).

The flow is exercised against an in-memory fake of the Supabase layer, with the REAL
`_normalize_bank_payment_entry` in the fake save path so the multi-invoice `int()` regression
(2026-09-28 payment lost) would be caught here too.
"""
import datetime
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from services.bank_payment_save import (
    apply_credit_plan,
    apply_partial_payment_plans,
    filter_payable_rows,
    is_row_payable,
    persist_bank_payment,
    plan_credit_application,
    plan_partial_payments,
    save_parsed_bank_payment,
)
from services.supabase_service import _PAYABLE_STATUS_REGEX, _normalize_bank_payment_entry, load_unpaid_credit_rows

PAY_DATE = datetime.date(2026, 9, 28)


class FakeDb:
    """In-memory stand-in for invoices / bank_payments / bank_payment_allocations."""

    def __init__(self, invoice_rows, fail_save=False, fail_mark_ids=(), fail_save_after=None):
        self.invoices = {r["id"]: dict(r) for r in invoice_rows}
        self.payments = {}  # fingerprint -> normalized payment
        self.allocations = {}  # fingerprint -> list of allocations
        self.fail_save = fail_save
        self.fail_mark_ids = set(fail_mark_ids)
        self.fail_save_after = fail_save_after  # raise on the Nth save call (1-based)
        self.save_calls = 0
        self.local_backups = []

    def get_invoice_rows(self, invoice_number):
        return [dict(r) for r in self.invoices.values() if str(r["invoice_number"]) == str(invoice_number)]

    def mark_row_paid(self, db_id, payment_date, payment_amount=None):
        if db_id in self.fail_mark_ids:
            raise RuntimeError(f"boom marking {db_id}")
        self.invoices[db_id]["paid"] = "Yes"
        self.invoices[db_id]["payment_date"] = payment_date.isoformat()

    def save_remote(self, entry, allocations):
        self.save_calls += 1
        if self.fail_save or (self.fail_save_after and self.save_calls == self.fail_save_after):
            raise RuntimeError("supabase insert failed")
        normalized = _normalize_bank_payment_entry(entry)  # real normalizer: must accept "8588,8598,..."
        self.payments[normalized["payment_fingerprint"]] = normalized
        self.allocations[normalized["payment_fingerprint"]] = [dict(a) for a in allocations]

    def save_local(self, entry):
        self.local_backups.append(entry)


def _row(row_id, inv, project, amount, paid="No", year=2025):
    return {
        "id": row_id, "invoice_number": str(inv), "project_name": project, "payment_amount": amount,
        "paid": paid, "maintenance_year": "Y1", "year": year, "cameras_number": 2,
    }


def _entry(fingerprint="fp1"):
    return {
        "payment_date": PAY_DATE.isoformat(), "source_name": "test.pdf", "source_kind": "pdf-batch",
        "payment_fingerprint": fingerprint, "instructed_amount": 100.0, "received_amount": 90.0,
        "fee_amount": 10.0, "currency": "EUR", "raw_text": "x", "parsed_payload": {},
    }


def _run(db, inv_nos, entry=None, on_row_paid=None):
    return save_parsed_bank_payment(
        payment_entry=entry or _entry(),
        invoice_numbers=inv_nos,
        payment_date=PAY_DATE,
        get_invoice_rows=db.get_invoice_rows,
        mark_row_paid=db.mark_row_paid,
        save_remote=db.save_remote,
        save_local_backup=db.save_local,
        on_row_paid=on_row_paid,
    )


class TestSingleInvoicePayment(unittest.TestCase):
    def test_single_invoice_saves_payment_and_marks_row(self):
        db = FakeDb([_row(1, 8678, "AD One", 500.0)])
        out = _run(db, [8678])
        self.assertTrue(out.fully_successful)
        self.assertEqual(db.invoices[1]["paid"], "Yes")
        self.assertEqual(db.payments["fp1"]["invoice_number"], "8678")
        self.assertEqual(db.payments["fp1"]["applied_amount"], 500.0)
        self.assertEqual(len(db.allocations["fp1"]), 1)
        self.assertEqual(out.confirmed_invoice_numbers, {8678})


class TestMultiInvoicePayment(unittest.TestCase):
    def test_multi_invoice_saves_comma_joined_number_and_all_allocations(self):
        db = FakeDb([
            _row(1, 8588, "AD A", 10892.0), _row(2, 8598, "AD B", 6224.0), _row(3, 8697, "AD C", 4668.0),
        ])
        out = _run(db, [8588, 8598, 8697])
        self.assertTrue(out.fully_successful, out.errors)
        self.assertEqual(db.payments["fp1"]["invoice_number"], "8588,8598,8697")
        self.assertEqual(db.payments["fp1"]["applied_amount"], 21784.0)
        self.assertEqual({a["invoice_number"] for a in db.allocations["fp1"]}, {8588, 8598, 8697})
        self.assertTrue(all(r["paid"] == "Yes" for r in db.invoices.values()))
        self.assertEqual(out.marked_rows, 3)

    def test_payment_is_saved_before_any_invoice_row_is_marked(self):
        order = []
        db = FakeDb([_row(1, 1, "AD A", 10.0)])
        original_save, original_mark = db.save_remote, db.mark_row_paid
        db.save_remote = lambda e, a: (order.append("save"), original_save(e, a))[1]
        db.mark_row_paid = lambda **kw: (order.append("mark"), original_mark(**kw))[1]
        _run(db, [1])
        self.assertEqual(order[0], "save")
        self.assertIn("mark", order)


class TestRowsAlreadyPaid(unittest.TestCase):
    def test_already_paid_rows_save_payment_without_allocations(self):
        db = FakeDb([_row(1, 8588, "AD A", 100.0, paid="Yes"), _row(2, 8598, "AD B", 50.0, paid="Yes")])
        out = _run(db, [8588, 8598])
        self.assertTrue(out.fully_successful)
        self.assertEqual(db.payments["fp1"]["invoice_number"], "8588,8598")
        self.assertIsNone(db.payments["fp1"]["applied_amount"])
        self.assertEqual(db.allocations["fp1"], [])
        self.assertEqual(out.marked_rows, 0)
        self.assertTrue(any("already paid" in n for n in out.notes))

    def test_no_invoice_number_and_no_rows_branches(self):
        db = FakeDb([])
        out = _run(db, [])
        self.assertTrue(out.fully_successful)
        self.assertTrue(any("without invoice match" in n for n in out.notes))
        db2 = FakeDb([])
        out2 = _run(db2, [9999], entry=_entry("fp2"))
        self.assertTrue(out2.fully_successful)
        self.assertTrue(any("no invoice rows" in n for n in out2.notes))


class TestFailureHandling(unittest.TestCase):
    def test_failed_payment_save_is_reported_and_leaves_invoices_untouched(self):
        db = FakeDb([_row(1, 8588, "AD A", 100.0), _row(2, 8598, "AD B", 50.0)], fail_save=True)
        with self.assertLogs("services.bank_payment_save", level="ERROR"):
            out = _run(db, [8588, 8598])
        self.assertFalse(out.payment_saved)
        self.assertFalse(out.fully_successful)
        self.assertTrue(any("NOT saved to Supabase" in e for e in out.errors))
        self.assertTrue(all(r["paid"] == "No" for r in db.invoices.values()), "no half-saved state allowed")
        self.assertEqual(out.marked_rows, 0)

    def test_local_backup_never_counts_as_success(self):
        db = FakeDb([_row(1, 1, "AD A", 10.0)], fail_save=True)
        with self.assertLogs("services.bank_payment_save", level="ERROR"):
            out = _run(db, [1])
        self.assertEqual(len(db.local_backups), 1)  # backup written...
        self.assertFalse(out.payment_saved)  # ...but still a failure
        self.assertTrue(any("NOT a durable save" in e for e in out.errors))

    def test_failed_save_in_already_paid_branch_is_reported(self):
        db = FakeDb([_row(1, 1, "AD A", 10.0, paid="Yes")], fail_save=True)
        with self.assertLogs("services.bank_payment_save", level="ERROR"):
            out = _run(db, [1])
        self.assertFalse(out.fully_successful)
        self.assertFalse(any("already paid" in n for n in out.notes), "must not claim success")

    def test_partial_row_failure_is_reported_and_payment_matches_reality(self):
        db = FakeDb(
            [_row(1, 8588, "AD A", 100.0), _row(2, 8598, "AD B", 50.0), _row(3, 8697, "AD C", 25.0)],
            fail_mark_ids={2},
        )
        with self.assertLogs("services.bank_payment_save", level="ERROR"):
            out = _run(db, [8588, 8598, 8697])
        self.assertTrue(out.payment_saved)
        self.assertFalse(out.fully_successful)
        self.assertTrue(any("AD B" in e and "could NOT be marked paid" in e for e in out.errors))
        self.assertEqual(db.invoices[2]["paid"], "No")
        self.assertEqual(db.payments["fp1"]["applied_amount"], 125.0)
        self.assertEqual({a["project_name"] for a in db.allocations["fp1"]}, {"AD A", "AD C"})

    def test_follow_up_step_failure_is_reported_but_row_stays_paid(self):
        db = FakeDb([_row(1, 1, "AD A", 10.0)])

        def boom(_row):
            raise RuntimeError("renewal exploded")

        with self.assertLogs("services.bank_payment_save", level="ERROR"):
            out = _run(db, [1], on_row_paid=boom)
        self.assertEqual(db.invoices[1]["paid"], "Yes")
        self.assertFalse(out.fully_successful)
        self.assertTrue(any("follow-up step failed" in e for e in out.errors))

    def test_persist_returns_none_on_success_and_message_on_failure(self):
        self.assertIsNone(persist_bank_payment({"source_name": "a"}, [], save_remote=lambda e, a: None))
        with self.assertLogs("services.bank_payment_save", level="ERROR"):
            message = persist_bank_payment(
                {"source_name": "a"}, [], save_remote=lambda e, a: (_ for _ in ()).throw(ValueError("x")),
            )
        self.assertIn("NOT saved to Supabase", message)


class TestCancelledRowsAreNeverPaid(unittest.TestCase):
    def test_is_row_payable_rule(self):
        for status, expected in [
            ("No", True), ("no", True), (" NO ", True), ("", True), ("   ", True), (None, True),
            ("Yes", False), ("yes", False), ("cancelled", False), ("Cancelled", False), ("Partial", False),
        ]:
            self.assertEqual(is_row_payable({"paid": status}), expected, repr(status))
        self.assertTrue(is_row_payable({}))  # missing key == "No"

    def test_mixed_cancelled_unpaid_and_paid_rows(self):
        db = FakeDb([
            _row(1, 8588, "AD Unpaid", 100.0, paid="No"),
            _row(2, 8588, "AD Cancelled", 999.0, paid="cancelled"),
            _row(3, 8598, "AD AlreadyPaid", 50.0, paid="Yes"),
            _row(4, 8598, "AD Unpaid2", 25.0, paid="No"),
            _row(5, 8697, "AD Cancelled2", 777.0, paid="Cancelled"),
        ])
        out = _run(db, [8588, 8598, 8697])
        self.assertTrue(out.fully_successful, out.errors)
        # only the two genuinely unpaid rows were marked paid
        self.assertEqual({i for i, r in db.invoices.items() if r["paid"] == "Yes"}, {1, 3, 4})
        self.assertEqual(db.invoices[2]["paid"], "cancelled")
        self.assertEqual(db.invoices[5]["paid"], "Cancelled")
        self.assertIsNone(db.invoices[2].get("payment_date"))
        # allocations and applied_amount exclude cancelled AND already-paid rows
        self.assertEqual({a["project_name"] for a in db.allocations["fp1"]}, {"AD Unpaid", "AD Unpaid2"})
        self.assertEqual(db.payments["fp1"]["applied_amount"], 125.0)
        self.assertEqual(out.marked_rows, 2)
        self.assertEqual(out.confirmed_invoice_numbers, {8588, 8598})
        self.assertTrue(any("2 cancelled invoice row(s) were excluded" in n for n in out.notes))

    def test_cancelled_plus_already_paid_only_marks_nothing(self):
        db = FakeDb([_row(1, 1, "AD Paid", 10.0, paid="Yes"), _row(2, 1, "AD Cancelled", 20.0, paid="cancelled")])
        out = _run(db, [1])
        self.assertTrue(out.fully_successful)
        self.assertEqual(out.marked_rows, 0)
        self.assertEqual(db.invoices[2]["paid"], "cancelled")
        self.assertEqual(db.allocations["fp1"], [])
        self.assertIsNone(db.payments["fp1"]["applied_amount"])
        self.assertTrue(any("no payable rows" in n for n in out.notes))

    def test_all_cancelled_marks_nothing(self):
        db = FakeDb([_row(1, 1, "AD C1", 10.0, paid="cancelled"), _row(2, 1, "AD C2", 20.0, paid="cancelled")])
        out = _run(db, [1])
        self.assertEqual(out.marked_rows, 0)
        self.assertTrue(all(r["paid"] == "cancelled" for r in db.invoices.values()))
        self.assertEqual(db.allocations["fp1"], [])

    def test_manual_lookup_flow_uses_the_same_rule(self):
        # Source-level guard: the Manual Lookup flow must derive payable rows from is_row_payable and
        # must not go back to the old `Paid != "Yes"` check that treated cancelled rows as unpaid.
        source = (Path(__file__).resolve().parent.parent / "streamlit_app.py").read_text(encoding="utf-8")
        self.assertIn('"Payable":      is_row_payable(r)', source)
        self.assertIn('df["Project"].isin(selected) & df["Payable"]', source)
        self.assertNotIn('df[df["Paid"] != "Yes"]', source)


class FakeInvoiceWriter:
    """Records every write made through update_row / insert_remainder_row."""

    def __init__(self):
        self.updates = []
        self.inserts = []

    def update_row(self, db_id, **fields):
        self.updates.append((db_id, fields))

    def insert_remainder_row(self, **fields):
        self.inserts.append(fields)


# One invoice with a payable row, a cancelled row, an already-paid row, an unrecognised-status row and a
# blank-status row -- every flow must treat these identically.
def _mixed_invoice_rows():
    return [
        {"id": 1, "invoice_number": "8700", "project_name": "AD Payable", "payment_amount": 100.0, "paid": "No", "maintenance_year": "Y1", "year": 2026},
        {"id": 2, "invoice_number": "8700", "project_name": "AD Cancelled", "payment_amount": 900.0, "paid": "Cancelled", "maintenance_year": "Y1", "year": 2026},
        {"id": 3, "invoice_number": "8700", "project_name": "AD Paid", "payment_amount": 50.0, "paid": "Yes", "maintenance_year": "Y1", "year": 2026},
        {"id": 4, "invoice_number": "8700", "project_name": "AD Weird", "payment_amount": 70.0, "paid": "Partial", "maintenance_year": "Y1", "year": 2026},
        {"id": 5, "invoice_number": "8700", "project_name": "AD Blank", "payment_amount": 30.0, "paid": "", "maintenance_year": "Y2", "year": 2026},
    ]


class TestFilterPayableRows(unittest.TestCase):
    def test_only_payable_rows_survive(self):
        ids = [r["id"] for r in filter_payable_rows(_mixed_invoice_rows())]
        self.assertEqual(ids, [1, 5])

    def test_positive_amount_only(self):
        rows = _mixed_invoice_rows() + [
            {"id": 6, "project_name": "AD Zero", "payment_amount": 0.0, "paid": "No"},
            {"id": 7, "project_name": "AD Credit", "payment_amount": -40.0, "paid": "No"},
        ]
        self.assertEqual([r["id"] for r in filter_payable_rows(rows, positive_amount_only=True)], [1, 5])
        self.assertEqual([r["id"] for r in filter_payable_rows(rows)], [1, 5, 6, 7])


class TestPartialTab(unittest.TestCase):
    def _editor_records(self, rows):
        # what st.data_editor(...).to_dict("records") returns: hidden id + editable columns
        return [
            {"id": r["id"], "Project": r["project_name"], "Maint. Year": r["maintenance_year"],
             "Original (€)": r["payment_amount"], "Paid Now (€)": r["payment_amount"]}
            for r in rows
        ]

    def test_listing_excludes_non_payable_rows(self):
        payable = filter_payable_rows(_mixed_invoice_rows())
        self.assertEqual({r["project_name"] for r in payable}, {"AD Payable", "AD Blank"})

    def test_cancelled_row_is_never_planned_even_if_present_in_the_editor(self):
        rows = _mixed_invoice_rows()
        payable_ids = {r["id"] for r in filter_payable_rows(rows)}
        # a tampered/stale editor state that still carries every row, including the cancelled one
        plans = plan_partial_payments(self._editor_records(rows), payable_ids)
        self.assertEqual({p["id"] for p in plans}, {1, 5})

    def test_cancelled_row_stays_unchanged_and_is_not_allocated(self):
        rows = _mixed_invoice_rows()
        payable_ids = {r["id"] for r in filter_payable_rows(rows)}
        records = self._editor_records(rows)
        records[0]["Paid Now (€)"] = 40.0  # partial on row 1: 40 of 100
        writer = FakeInvoiceWriter()
        allocations = apply_partial_payment_plans(
            plan_partial_payments(records, payable_ids),
            invoice_number=8700, payment_date=PAY_DATE, year=2026,
            update_row=writer.update_row, insert_remainder_row=writer.insert_remainder_row,
        )
        touched = {db_id for db_id, _ in writer.updates}
        self.assertEqual(touched, {1, 5})
        self.assertFalse(touched & {2, 3, 4}, "cancelled / paid / unrecognised rows must not be written")
        self.assertEqual({a["invoice_row_id"] for a in allocations}, {1, 5})
        self.assertEqual(sum(a["amount_applied"] for a in allocations), 70.0)  # 40 + 30, nothing from row 2
        self.assertEqual(len(writer.inserts), 1)  # remainder (60) only for row 1
        self.assertEqual(writer.inserts[0]["payment_amount"], 60.0)
        self.assertEqual(writer.inserts[0]["project_name"], "AD Payable")

    def test_paid_now_is_clamped_and_zero_is_skipped(self):
        records = [
            {"id": 1, "Project": "A", "Maint. Year": "Y1", "Original (€)": 100.0, "Paid Now (€)": 500.0},
            {"id": 5, "Project": "B", "Maint. Year": "Y1", "Original (€)": 30.0, "Paid Now (€)": 0.0},
        ]
        plans = plan_partial_payments(records, {1, 5})
        self.assertEqual(len(plans), 1)
        self.assertEqual((plans[0]["paid_now"], plans[0]["remaining"]), (100.0, 0.0))


class TestApplyCreditTab(unittest.TestCase):
    def test_cancelled_target_rows_are_not_listed_or_funded(self):
        rows = _mixed_invoice_rows()
        targets = filter_payable_rows(rows, positive_amount_only=True)
        self.assertEqual([r["id"] for r in targets], [1, 5])
        plan = plan_credit_application(rows, 1000.0)  # even if the unfiltered rows are passed in
        self.assertEqual([item["row"]["id"] for item in plan], [1, 5])

    def test_credit_goes_only_to_payable_rows_and_cancelled_row_is_untouched(self):
        rows = _mixed_invoice_rows()
        writer = FakeInvoiceWriter()
        allocations, consumed = apply_credit_plan(
            plan_credit_application(rows, 1000.0),
            invoice_number=8700, apply_date=PAY_DATE, update_row=writer.update_row,
        )
        self.assertEqual({db_id for db_id, _ in writer.updates}, {1, 5})
        self.assertEqual(consumed, 130.0)  # 100 + 30; the 900 cancelled row never absorbs credit
        self.assertEqual({a["invoice_row_id"] for a in allocations}, {1, 5})
        self.assertEqual(sum(a["amount_applied"] for a in allocations), 130.0)
        fields_by_id = dict(writer.updates)
        self.assertEqual(fields_by_id[1]["paid"], "Yes")
        self.assertEqual(fields_by_id[1]["payment_date"], PAY_DATE)

    def test_partial_credit_leaves_row_unpaid_and_stops_when_credit_runs_out(self):
        rows = _mixed_invoice_rows()
        writer = FakeInvoiceWriter()
        allocations, consumed = apply_credit_plan(
            plan_credit_application(rows, 60.0),
            invoice_number=8700, apply_date=PAY_DATE, update_row=writer.update_row,
        )
        self.assertEqual(consumed, 60.0)
        self.assertEqual([db_id for db_id, _ in writer.updates], [1])
        fields = writer.updates[0][1]
        self.assertEqual((fields["payment_amount"], fields["paid"], fields["payment_date"]), (40.0, "No", None))

    def test_cancelled_credit_rows_cannot_fund_a_credit(self):
        credit_rows = [
            {"id": 10, "project_name": "AD C", "payment_amount": -500.0, "paid": "No"},
            {"id": 11, "project_name": "AD C", "payment_amount": -999.0, "paid": "Cancelled"},
            {"id": 12, "project_name": "AD C", "payment_amount": -25.0, "paid": "Partial"},
        ]
        usable = filter_payable_rows(credit_rows)
        self.assertEqual(abs(sum(r["payment_amount"] for r in usable)), 500.0)


class _FakeQuery:
    """Chainable stand-in for a supabase query; ignores filters and returns canned rows (like a DB bug would)."""

    def __init__(self, rows, log):
        self._rows, self._log = rows, log

    def __getattr__(self, name):
        def record(*args, **kwargs):
            self._log.append((name, args, kwargs))
            return self
        return record

    def execute(self):
        return SimpleNamespace(data=list(self._rows))


class _FakeClient:
    def __init__(self, rows):
        self.rows, self.log = rows, []

    def table(self, name):
        self.log.append(("table", (name,), {}))
        return _FakeQuery(self.rows, self.log)


class TestLoadUnpaidCreditRows(unittest.TestCase):
    CREDIT_ROWS = [
        {"id": 10, "project_name": "AD C", "payment_amount": -500.0, "paid": "No"},
        {"id": 11, "project_name": "AD C", "payment_amount": -999.0, "paid": "cancelled"},
        {"id": 12, "project_name": "AD C", "payment_amount": -999.0, "paid": "Cancelled"},
        {"id": 13, "project_name": "AD C", "payment_amount": -25.0, "paid": "Partial"},
        {"id": 14, "project_name": "AD C", "payment_amount": -40.0, "paid": "Yes"},
        {"id": 15, "project_name": "AD C", "payment_amount": -10.0, "paid": None},
    ]

    def _load(self, project=None):
        client = _FakeClient(self.CREDIT_ROWS)
        with patch("services.supabase_service._get_client", return_value=client):
            return load_unpaid_credit_rows(project), client

    def test_cancelled_credit_row_is_never_returned_even_if_the_database_returns_it(self):
        rows, _ = self._load()
        self.assertNotIn(11, {r["id"] for r in rows})
        self.assertNotIn(12, {r["id"] for r in rows})
        self.assertEqual({r["id"] for r in rows}, {10, 15})  # payable "No" and missing status only

    def test_query_filters_in_the_database_and_no_longer_uses_neq_yes(self):
        _, client = self._load("AD C")
        names = [call[0] for call in client.log]
        self.assertNotIn("neq", names)
        or_calls = [call for call in client.log if call[0] == "or_"]
        self.assertEqual(len(or_calls), 1)
        self.assertIn("paid.is.null", or_calls[0][1][0])
        self.assertIn("paid.imatch", or_calls[0][1][0])
        self.assertIn(("eq", ("project_name", "AD C"), {}), client.log)

    def test_database_regex_accepts_exactly_what_is_row_payable_accepts(self):
        import re
        pattern = re.compile(_PAYABLE_STATUS_REGEX, re.IGNORECASE)
        for status in ["No", "no", "NO", " No ", "", "   ", "Yes", "yes", "cancelled", "Cancelled", "Partial", "None", "Nope", "not paid"]:
            self.assertEqual(bool(pattern.match(status)), is_row_payable({"paid": status}), repr(status))
        # NULL is covered by the separate `paid.is.null` branch of the query
        self.assertTrue(is_row_payable({"paid": None}))


class TestOneEligibilityRuleEverywhere(unittest.TestCase):
    """Source-level guards: no payment flow may grow its own status check again."""

    SOURCE = (Path(__file__).resolve().parent.parent / "streamlit_app.py").read_text(encoding="utf-8")

    def test_no_inline_yes_comparison_in_payment_flows(self):
        import re
        self.assertIsNone(re.search(r'lower\(\)\s*!=\s*"yes"', self.SOURCE))
        self.assertNotIn('!= "Yes"', self.SOURCE)

    def test_partial_and_apply_credit_tabs_use_the_shared_helpers(self):
        self.assertIn("partial_rows = filter_payable_rows(get_invoices_by_number(int(partial_inv)))", self.SOURCE)
        self.assertIn("plan_partial_payments(", self.SOURCE)
        self.assertIn("apply_partial_payment_plans(", self.SOURCE)
        self.assertIn("filter_payable_rows(get_invoices_by_number(int(apply_inv)), positive_amount_only=True)", self.SOURCE)
        self.assertIn("filter_payable_rows(load_unpaid_credit_rows(", self.SOURCE)
        self.assertIn("plan_credit_application(", self.SOURCE)
        self.assertIn("apply_credit_plan(", self.SOURCE)

    def test_every_flow_that_marks_rows_paid_is_gated_by_the_rule(self):
        # mark_invoice_row_paid / update_invoice_row(paid="Yes") may only be reached with payable rows:
        # Save All via save_parsed_bank_payment, Manual Lookup via the Payable column.
        self.assertIn("is_row_payable(r)", self.SOURCE)
        self.assertIn('df["Project"].isin(selected) & df["Payable"]', self.SOURCE)
        self.assertIn("save_parsed_bank_payment(", self.SOURCE)


if __name__ == "__main__":
    unittest.main()
