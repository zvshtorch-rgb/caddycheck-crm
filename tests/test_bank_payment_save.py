"""Tests for services/bank_payment_save.py (the Save All Parsed Payments flow).

The flow is exercised against an in-memory fake of the Supabase layer, with the REAL
`_normalize_bank_payment_entry` in the fake save path so the multi-invoice `int()` regression
(2026-09-28 payment lost) would be caught here too.
"""
import datetime
import unittest

from services.bank_payment_save import persist_bank_payment, save_parsed_bank_payment
from services.supabase_service import _normalize_bank_payment_entry

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


if __name__ == "__main__":
    unittest.main()
