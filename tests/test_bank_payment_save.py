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
    apply_credit,
    apply_partial_payment_plans,
    filter_payable_rows,
    is_row_payable,
    persist_bank_payment,
    plan_credit_application,
    plan_partial_payments,
    save_parsed_bank_payment,
)
from services.supabase_service import (
    _PAYABLE_STATUS_REGEX,
    _normalize_bank_payment_entry,
    consume_credit_row,
    insert_invoice_adjustment_row,
    load_unpaid_credit_rows,
)

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

    def test_cancelled_credit_rows_cannot_fund_a_credit(self):
        credit_rows = [
            {"id": 10, "project_name": "AD C", "payment_amount": -500.0, "paid": "No"},
            {"id": 11, "project_name": "AD C", "payment_amount": -999.0, "paid": "Cancelled"},
            {"id": 12, "project_name": "AD C", "payment_amount": -25.0, "paid": "Partial"},
        ]
        usable = filter_payable_rows(credit_rows)
        self.assertEqual(abs(sum(r["payment_amount"] for r in usable)), 500.0)


class FakeLedger:
    """In-memory invoices table with the same semantics as the Supabase functions apply_credit() calls."""

    def __init__(self, rows):
        self.rows = {r["id"]: dict(r) for r in rows}
        self.next_id = 1000
        self.consume_calls = 0

    # -- what the app loads (mirrors load_unpaid_credit_rows / the target-row query) ------------------
    def credit_rows(self):
        return [dict(r) for r in self.rows.values() if is_row_payable(r) and r["payment_amount"] < 0]

    def target_rows(self, invoice_number):
        return [
            dict(r) for r in self.rows.values()
            if str(r["invoice_number"]) == str(invoice_number) and is_row_payable(r) and r["payment_amount"] > 0
        ]

    # -- writes, same contract as the Supabase service functions ----------------------------------------
    def consume_credit_row(self, row_id, expected_amount, use_amount, apply_date, description):
        self.consume_calls += 1
        row = self.rows.get(row_id)
        if row is None or row["payment_amount"] != expected_amount or not is_row_payable(row):
            return False  # conditional update matched nothing
        remaining = round(abs(expected_amount) - use_amount, 2)
        row["description"] = description
        if remaining <= 0.005:
            row["paid"], row["payment_date"] = "Yes", apply_date
        else:
            row["payment_amount"] = -remaining
        return True

    def update_row(self, row_id, **fields):
        for key, value in fields.items():
            if value is not None:
                self.rows[row_id][key] = value

    def insert_row(self, **kw):
        self.next_id += 1
        self.rows[self.next_id] = {
            "id": self.next_id, "invoice_number": kw.get("invoice_number"), "project_name": kw["project_name"],
            "maintenance_year": kw["maintenance_year"], "payment_amount": kw["payment_amount"],
            "paid": kw.get("paid", "No"), "payment_date": kw.get("payment_date"), "year": kw.get("year"),
            "description": kw.get("description"),
        }
        return self.rows[self.next_id]

    # -- helpers ---------------------------------------------------------------------------------------
    def apply(self, invoice_number, credit_rows=None, target_rows=None, **overrides):
        kwargs = dict(
            target_rows=self.target_rows(invoice_number) if target_rows is None else target_rows,
            credit_rows=self.credit_rows() if credit_rows is None else credit_rows,
            invoice_number=invoice_number, apply_date=PAY_DATE, year=2026,
            consume_credit_row=self.consume_credit_row, update_row=self.update_row, insert_row=self.insert_row,
        )
        kwargs.update(overrides)
        return apply_credit(**kwargs)

    def available(self):
        return round(abs(sum(r["payment_amount"] for r in self.credit_rows())), 2)

    def snapshot(self, ids=None):
        return {i: dict(r) for i, r in self.rows.items() if ids is None or i in ids}

    def total(self, predicate=lambda r: True):
        return round(sum(r["payment_amount"] for r in self.rows.values() if predicate(r)), 2)


def _credit(row_id, amount, paid="No", invoice_number="8683", project="CREDIT NOTE #8683"):
    return {"id": row_id, "invoice_number": invoice_number, "project_name": project, "maintenance_year": "Credit",
            "payment_amount": amount, "paid": paid, "year": 2026, "description": "Duplicate transaction credit",
            "payment_date": None}


def _target(row_id, amount, paid="No", invoice_number="8700", project="AD Target"):
    return {"id": row_id, "invoice_number": invoice_number, "project_name": project, "maintenance_year": "Y1",
            "payment_amount": amount, "paid": paid, "year": 2026, "description": None, "payment_date": None}


class TestApplyCreditAccounting(unittest.TestCase):
    """Apply Credit must consume the credit source rows -- never create new available credit."""

    def test_1000_credit_apply_130_leaves_870(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        result = ledger.apply(8700)
        self.assertEqual(result.errors, [])
        self.assertEqual((result.consumed, result.applied), (130.0, 130.0))
        self.assertEqual(ledger.available(), 870.0)
        self.assertEqual(ledger.rows[1]["payment_amount"], -870.0)  # source row reduced in place
        self.assertEqual(ledger.rows[1]["paid"], "No")  # still the one open credit row
        self.assertEqual(ledger.rows[2]["paid"], "Yes")
        self.assertEqual(ledger.rows[2]["payment_amount"], 130.0)  # target keeps its gross amount

    def test_applying_never_creates_an_unpaid_negative_row(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        before_ids = set(ledger.rows)
        ledger.apply(8700)
        new_rows = [r for i, r in ledger.rows.items() if i not in before_ids]
        self.assertEqual(len(new_rows), 1)  # just the usage record
        self.assertTrue(all(not (r["payment_amount"] < 0 and is_row_payable(r)) for r in new_rows))
        self.assertEqual((new_rows[0]["payment_amount"], new_rows[0]["paid"]), (-130.0, "Yes"))

    def test_applying_the_remaining_870_brings_available_credit_to_zero(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0), _target(3, 870.0, invoice_number="8701")])
        ledger.apply(8700)
        result = ledger.apply(8701)
        self.assertEqual(result.errors, [])
        self.assertEqual(ledger.available(), 0.0)
        self.assertEqual(ledger.credit_rows(), [])  # load_unpaid_credit_rows would return nothing
        self.assertEqual(ledger.rows[1]["paid"], "Yes")  # fully consumed source row is closed...
        self.assertEqual(ledger.rows[1]["payment_amount"], -870.0)  # ...and keeps its amount

    def test_partial_consumption_across_multiple_credit_rows(self):
        ledger = FakeLedger([_credit(1, -100.0), _credit(2, -50.0), _credit(3, -200.0), _target(10, 220.0)])
        result = ledger.apply(8700)
        self.assertEqual(result.errors, [])
        self.assertEqual([u["credit_row_id"] for u in result.credit_usage], [1, 2, 3])  # oldest first
        self.assertEqual([u["used"] for u in result.credit_usage], [100.0, 50.0, 70.0])
        self.assertEqual((ledger.rows[1]["paid"], ledger.rows[2]["paid"], ledger.rows[3]["paid"]), ("Yes", "Yes", "No"))
        self.assertEqual(ledger.rows[3]["payment_amount"], -130.0)
        self.assertEqual(ledger.available(), 130.0)

    def test_credit_source_cannot_be_used_twice_even_with_stale_data(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0), _target(3, 130.0, invoice_number="8701")])
        stale_credit_rows = ledger.credit_rows()  # loaded before the first application
        stale_targets = ledger.target_rows(8701)
        ledger.apply(8700)
        after_first = ledger.snapshot()
        result = ledger.apply(8701, credit_rows=stale_credit_rows, target_rows=stale_targets)
        self.assertEqual((result.consumed, result.applied), (0.0, 0.0))
        self.assertTrue(any("changed since it was loaded" in e for e in result.errors))
        self.assertEqual(ledger.snapshot(), after_first)  # nothing was spent or written the second time
        self.assertEqual(ledger.available(), 870.0)

    def test_cancelled_paid_and_invalid_rows_stay_untouched(self):
        ledger = FakeLedger([
            _credit(1, -300.0),
            _credit(2, -999.0, paid="cancelled"),
            _credit(3, -40.0, paid="Yes"),
            _credit(4, -25.0, paid="Partial"),
            _target(10, 100.0),
            _target(11, 900.0, paid="Cancelled"),
            _target(12, 50.0, paid="Yes"),
            _target(13, 70.0, paid="Partial"),
        ])
        untouched_ids = {2, 3, 4, 11, 12, 13}
        before = ledger.snapshot(untouched_ids)
        result = ledger.apply(8700)
        self.assertEqual(result.applied, 100.0)
        self.assertEqual(ledger.snapshot(untouched_ids), before)
        self.assertEqual(ledger.available(), 200.0)  # only the open credit row (300 - 100)

    def test_rerunning_the_same_action_does_not_duplicate_the_usage(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        ledger.apply(8700)
        after_first = ledger.snapshot()
        second = ledger.apply(8700)  # same click again, data reloaded
        self.assertEqual((second.consumed, second.applied, second.errors), (0.0, 0.0, []))
        self.assertEqual(ledger.snapshot(), after_first)
        self.assertEqual(ledger.available(), 870.0)

    def test_accounting_identity_original_credit_equals_applied_plus_remaining(self):
        ledger = FakeLedger([_credit(1, -600.0), _credit(2, -400.0), _target(10, 130.0), _target(11, 500.0, invoice_number="8701")])
        original = 1000.0
        applied = 0.0
        for invoice in (8700, 8701):
            applied += ledger.apply(invoice).applied
        credit_note_rows = [r for r in ledger.rows.values() if r["maintenance_year"] == "Credit"]
        self.assertEqual(applied, 630.0)
        self.assertEqual(round(applied + ledger.available(), 2), original)
        # the ledger itself still adds up to the original credit: usage rows + source rows
        self.assertEqual(round(abs(sum(r["payment_amount"] for r in credit_note_rows)), 2), original)

    def test_application_does_not_change_ledger_total_debt_or_paid_total(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        total_before = ledger.total()
        debt_before = ledger.total(lambda r: r["paid"] == "No")
        paid_before = ledger.total(lambda r: r["paid"] == "Yes")
        ledger.apply(8700)
        self.assertEqual(ledger.total(), total_before)  # invoiced total unchanged
        self.assertEqual(ledger.total(lambda r: r["paid"] == "No"), debt_before)  # net open debt unchanged
        self.assertEqual(ledger.total(lambda r: r["paid"] == "Yes"), paid_before)  # no cash invented

    def test_allocations_and_applied_amount_describe_only_the_target_rows(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0), _target(3, 20.0)])
        result = ledger.apply(8700)
        self.assertEqual({a["invoice_row_id"] for a in result.allocations}, {2, 3})
        self.assertEqual(sum(a["amount_applied"] for a in result.allocations), result.applied)
        self.assertEqual(result.applied, 150.0)
        self.assertTrue(all(a["amount_applied"] > 0 for a in result.allocations))
        self.assertEqual(result.credit_usage, [
            {"credit_row_id": 1, "used": 150.0, "remaining_after": 850.0, "usage_row_created": True},
        ])

    def test_partially_funded_target_is_split_and_gross_amount_is_preserved(self):
        ledger = FakeLedger([_credit(1, -100.0), _target(2, 130.0)])
        total_before = ledger.total()
        result = ledger.apply(8700)
        self.assertEqual(result.applied, 100.0)
        self.assertEqual((ledger.rows[2]["payment_amount"], ledger.rows[2]["paid"]), (100.0, "Yes"))
        remainder = [r for i, r in ledger.rows.items() if i > 1000 and r["payment_amount"] > 0]
        self.assertEqual([(r["payment_amount"], r["paid"]) for r in remainder], [(30.0, "No")])
        self.assertEqual(ledger.total(), total_before)
        self.assertEqual(ledger.available(), 0.0)

    def test_failures_are_reported_and_credit_is_never_applied_beyond_what_was_consumed(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])

        def exploding_consume(*args, **kwargs):
            raise RuntimeError("db down")

        with self.assertLogs("services.bank_payment_save", level="ERROR"):
            result = ledger.apply(8700, consume_credit_row=exploding_consume)
        self.assertEqual((result.consumed, result.applied), (0.0, 0.0))
        self.assertTrue(any("could not be consumed" in e for e in result.errors))
        self.assertEqual(ledger.rows[2]["paid"], "No")  # target untouched when no credit was obtained
        self.assertEqual(ledger.available(), 1000.0)

    def test_target_write_failure_is_reported_as_consumed_but_not_applied(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])

        def exploding_update(*args, **kwargs):
            raise RuntimeError("update failed")

        with self.assertLogs("services.bank_payment_save", level="ERROR"):
            result = ledger.apply(8700, update_row=exploding_update)
        self.assertEqual((result.consumed, result.applied), (130.0, 0.0))
        self.assertTrue(any("does not match" in e for e in result.errors))


class TestConsumeCreditRowQuery(unittest.TestCase):
    """The real Supabase function: conditional update so a credit row can only be spent once."""

    def _run(self, returned_rows, expected_amount=-1000.0, use=130.0):
        client = _FakeClient(returned_rows)
        with patch("services.supabase_service._get_client", return_value=client):
            ok = consume_credit_row(7, expected_amount, use, PAY_DATE, "note")
        return ok, client

    def test_partial_use_reduces_the_amount_and_is_conditional(self):
        ok, client = self._run([{"id": 7}])
        self.assertTrue(ok)
        update = next(c for c in client.log if c[0] == "update")
        self.assertEqual(update[1][0], {"description": "note", "payment_amount": -870.0})
        self.assertIn(("eq", ("id", 7), {}), client.log)
        self.assertIn(("eq", ("payment_amount", -1000.0), {}), client.log)  # only if still at the loaded amount
        self.assertTrue(any(c[0] == "or_" and "paid.imatch" in c[1][0] for c in client.log))  # and still open

    def test_full_use_marks_the_row_paid_and_keeps_its_amount(self):
        ok, client = self._run([{"id": 7}], expected_amount=-130.0, use=130.0)
        self.assertTrue(ok)
        update = next(c for c in client.log if c[0] == "update")
        self.assertEqual(update[1][0], {"description": "note", "paid": "Yes", "payment_date": PAY_DATE.isoformat()})

    def test_returns_false_when_no_row_matched(self):
        ok, _ = self._run([])
        self.assertFalse(ok)

    def test_insert_helper_defaults_are_unchanged_and_usage_rows_can_be_paid(self):
        client = _FakeClient([{"id": 1}])
        with patch("services.supabase_service._get_client", return_value=client):
            insert_invoice_adjustment_row(invoice_number=8683, project_name="X", maintenance_year="Credit",
                                          payment_amount=-5.0, year=2026)
            insert_invoice_adjustment_row(invoice_number=8683, project_name="X", maintenance_year="Credit",
                                          payment_amount=-5.0, year=2026, paid="Yes", payment_date=PAY_DATE)
        inserts = [c[1][0] for c in client.log if c[0] == "insert"]
        self.assertEqual((inserts[0]["paid"], inserts[0]["payment_date"]), ("No", None))
        self.assertEqual((inserts[1]["paid"], inserts[1]["payment_date"]), ("Yes", PAY_DATE.isoformat()))


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
        self.assertIn("apply_credit(", self.SOURCE)
        self.assertNotIn("apply_credit_plan", self.SOURCE)
        self.assertNotIn('description=f"Credit consumed for INV#', self.SOURCE)  # the old row that re-created credit

    def test_every_flow_that_marks_rows_paid_is_gated_by_the_rule(self):
        # mark_invoice_row_paid / update_invoice_row(paid="Yes") may only be reached with payable rows:
        # Save All via save_parsed_bank_payment, Manual Lookup via the Payable column.
        self.assertIn("is_row_payable(r)", self.SOURCE)
        self.assertIn('df["Project"].isin(selected) & df["Payable"]', self.SOURCE)
        self.assertIn("save_parsed_bank_payment(", self.SOURCE)


if __name__ == "__main__":
    unittest.main()
