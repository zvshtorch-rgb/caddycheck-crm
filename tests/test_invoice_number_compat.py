"""Regression tests for the Invoice Details compatibility patch: an existing invoice's number can never change through
the grid save, and a blocked save writes NOTHING (no partial batches).

* logic + service level: pure functions and `upsert_invoices` against a tiny in-memory client
* UI level: the real app script (Streamlit AppTest; skipped when Streamlit is not installed) with `st.data_editor`
  replaced by a stand-in that returns an edited / manipulated grid, a dummy Supabase URL and an in-memory database.
"""
import importlib.util
import os
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from services import supabase_service
from services.invoice_save_guard import (
    InvoiceNumberChangeBlocked,
    assert_no_invoice_number_changes,
    find_invoice_number_changes,
    normalize_invoice_number,
)

HAS_STREAMLIT = importlib.util.find_spec("streamlit") is not None
APP_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "streamlit_app.py")
WRITE_OPS = ("upsert", "insert", "update", "delete")


class _Query:
    def __init__(self, db, table):
        self.db, self.table = db, table
        self.mode, self.payload, self.columns = "select", None, "*"
        self.filters, self.lo, self.hi = [], None, None

    def select(self, columns="*"):
        self.mode, self.columns = "select", columns
        return self

    def upsert(self, payload):
        self.mode, self.payload = "upsert", payload
        return self

    def insert(self, payload):
        self.mode, self.payload = "insert", payload
        return self

    def update(self, payload):
        self.mode, self.payload = "update", payload
        return self

    def delete(self):
        self.mode = "delete"
        return self

    def eq(self, col, val):
        self.filters.append(("eq", col, val))
        return self

    def in_(self, col, vals):
        self.filters.append(("in", col, list(vals)))
        return self

    def order(self, col, desc=False):
        return self

    def range(self, lo, hi):
        self.lo, self.hi = lo, hi
        return self

    def _hit(self, row):
        for kind, col, val in self.filters:
            if kind == "eq" and str(row.get(col)) != str(val):
                return False
            if kind == "in" and row.get(col) not in val:
                return False
        return True

    def execute(self):
        rows = self.db.tables.setdefault(self.table, [])
        if self.mode == "select":
            self.db.log.append(("select", self.table))
            data = [dict(r) for r in rows if self._hit(r)]
            data.sort(key=lambda r: r.get("id", 0))
            if self.lo is not None:
                data = data[self.lo:self.hi + 1]
            if self.columns != "*":
                cols = [c.strip() for c in self.columns.split(",")]
                data = [{c: r.get(c) for c in cols} for r in data]
            return SimpleNamespace(data=data)
        batch = self.payload if isinstance(self.payload, list) else [self.payload]
        self.db.log.append((self.mode, self.table, [dict(b) for b in batch] if self.payload is not None else None))
        if self.mode == "upsert":
            for item in batch:
                for row in rows:
                    if row.get("id") == item.get("id"):
                        row.update(item)
                        break
                else:
                    rows.append(dict(item))
            return SimpleNamespace(data=batch)
        if self.mode == "insert":
            for item in batch:
                rows.append({"id": max([r.get("id", 0) for r in rows] + [0]) + 1, **item})
            return SimpleNamespace(data=batch)
        if self.mode == "update":
            hit = [r for r in rows if self._hit(r)]
            for r in hit:
                r.update(self.payload)
            return SimpleNamespace(data=hit)
        hit = [r for r in rows if self._hit(r)]
        self.db.tables[self.table] = [r for r in rows if not self._hit(r)]
        return SimpleNamespace(data=hit)


class FakeDb:
    def __init__(self, invoices):
        self.tables = {"invoices": [dict(r) for r in invoices]}
        self.log = []

    def table(self, name):
        return _Query(self, name)

    def writes(self):
        return [e for e in self.log if e[0] in WRITE_OPS]


def db_rows(n, number=lambda i: 8000 + i):
    return [{"id": i, "invoice_number": str(number(i)), "project_name": f"P{i}", "maintenance_year": "Y1",
             "payment_amount": 100.0, "paid": "No", "year": 2026} for i in range(1, n + 1)]


def invoice_obj(db_id, number, project="P", amount=100.0):
    return SimpleNamespace(
        invoice_number=float(number) if number is not None else None, project_name=project, maintenance_year="Y1",
        payment_amount=amount, cameras_number=None, payment_date=None, paid="No", year=2026, invoice_type="",
        for_month="", sent_at="", description=None, db_id=db_id)


class TestPureComparison(unittest.TestCase):
    def test_normalisation_treats_equivalent_spellings_as_equal(self):
        for value in (8713, 8713.0, "8713", " 8713 ", "8713.0"):
            self.assertEqual(normalize_invoice_number(value), "8713")
        for value in (None, "", float("nan"), "nan"):
            self.assertIsNone(normalize_invoice_number(value))

    def test_unchanged_grid_has_no_changes_even_for_float_or_nan_values(self):
        originals = [invoice_obj(1, 8700, "A"), invoice_obj(2, None, "B")]
        rows = [{"_invoice_index": 0.0, "Invoice #": 8700.0}, {"_invoice_index": 1, "Invoice #": float("nan")}]
        self.assertEqual(find_invoice_number_changes(rows, originals), [])

    def test_changed_and_cleared_numbers_are_reported(self):
        originals = [invoice_obj(1, 8700, "A"), invoice_obj(2, 8701, "B")]
        rows = [{"_invoice_index": 0, "Invoice #": 8713}, {"_invoice_index": 1, "Invoice #": None}]
        changes = find_invoice_number_changes(rows, originals)
        self.assertEqual([(c["project"], c["old"], c["new"]) for c in changes], [("A", "8700", "8713"), ("B", "8701", None)])

    def test_new_rows_and_unknown_indexes_are_ignored(self):
        originals = [invoice_obj(1, 8700, "A")]
        rows = [{"_invoice_index": None, "Invoice #": 9999}, {"_invoice_index": "", "Invoice #": 9998},
                {"_invoice_index": 7, "Invoice #": 9997}, {"_invoice_index": -1, "Invoice #": 9996}]
        self.assertEqual(find_invoice_number_changes(rows, originals), [])


class TestServiceLevelAtomicity(unittest.TestCase):
    def upsert(self, db, invoices):
        with patch.object(supabase_service, "_get_client", return_value=db):
            supabase_service.upsert_invoices(invoices)

    def test_a_normal_edit_still_saves(self):
        db = FakeDb(db_rows(3))
        invoices = [invoice_obj(i, 8000 + i, f"P{i}", amount=100.0) for i in (1, 2, 3)]
        invoices[1].payment_amount = 555.0
        self.upsert(db, invoices)
        upserts = [e for e in db.log if e[0] == "upsert"]
        self.assertEqual(len(upserts), 1)
        saved = {r["id"]: r for r in upserts[0][2]}
        self.assertEqual(saved[2]["payment_amount"], 555.0)
        self.assertEqual({i: saved[i]["invoice_number"] for i in saved}, {1: "8001", 2: "8002", 3: "8003"})

    def test_a_renumber_in_the_last_batch_of_120_rows_writes_nothing_at_all(self):
        db = FakeDb(db_rows(120))
        invoices = [invoice_obj(i, 8000 + i, f"P{i}") for i in range(1, 121)]
        invoices[-1].invoice_number = 9999.0  # row 120 sits in the THIRD batch of 50
        invoices[0].payment_amount = 777.0     # a legitimate edit in the FIRST batch must not be saved either
        with self.assertRaises(InvoiceNumberChangeBlocked) as ctx:
            self.upsert(db, invoices)
        self.assertIn("#8120 -> #9999", str(ctx.exception))
        self.assertEqual(db.writes(), [], "no batch may be written, not even the clean first ones")
        self.assertEqual(db.tables["invoices"][0]["payment_amount"], 100.0)

    def test_clearing_a_number_is_a_renumber_too(self):
        db = FakeDb(db_rows(2))
        invoices = [invoice_obj(1, 8001, "P1"), invoice_obj(2, None, "P2")]
        with self.assertRaises(InvoiceNumberChangeBlocked):
            self.upsert(db, invoices)
        self.assertEqual(db.writes(), [])

    def test_a_stale_client_writing_back_an_old_number_is_blocked(self):
        # Someone renumbered row 2 in the database; this tab still holds (and re-saves) the OLD number.
        db = FakeDb(db_rows(3))
        db.tables["invoices"][1]["invoice_number"] = "8715"
        invoices = [invoice_obj(i, 8000 + i, f"P{i}") for i in (1, 2, 3)]
        with self.assertRaises(InvoiceNumberChangeBlocked):
            self.upsert(db, invoices)
        self.assertEqual(db.writes(), [])
        self.assertEqual(db.tables["invoices"][1]["invoice_number"], "8715")

    def test_the_check_happens_before_any_write_and_text_formats_are_equal(self):
        db = FakeDb(db_rows(2))
        db.tables["invoices"][0]["invoice_number"] = "8001.0"  # same number, different text format -> allowed
        self.upsert(db, [invoice_obj(1, 8001, "P1"), invoice_obj(2, 8002, "P2")])
        kinds = [e[0] for e in db.log]
        self.assertEqual(kinds[0], "select")
        self.assertLess(kinds.index("select"), kinds.index("upsert"))

    def test_new_rows_can_still_be_added(self):
        db = FakeDb(db_rows(2))
        new = invoice_obj(None, 8100, "P-NEW")
        self.upsert(db, [invoice_obj(1, 8001, "P1"), invoice_obj(2, 8002, "P2"), new])
        self.assertEqual([e[0] for e in db.log if e[0] == "insert"], ["insert"])
        self.assertEqual(len(db.tables["invoices"]), 3)

    def test_a_ledger_larger_than_one_page_is_checked_completely(self):
        db = FakeDb(db_rows(1500))
        rows = [{"id": i, "invoice_number": str(8000 + i)} for i in range(1, 1501)]
        rows[1399]["invoice_number"] = "9999"  # row 1400 is beyond the first 1000-row page
        with self.assertRaises(InvoiceNumberChangeBlocked):
            assert_no_invoice_number_changes(db, rows)

    def test_only_reads_happen_while_checking(self):
        db = FakeDb(db_rows(5))
        assert_no_invoice_number_changes(db, [{"id": i, "invoice_number": str(8000 + i)} for i in range(1, 6)])
        self.assertEqual(db.writes(), [])


@unittest.skipUnless(HAS_STREAMLIT, "streamlit is not installed")
class TestInvoiceDetailsPage(unittest.TestCase):
    def open_page(self, db, manipulate=None, loaded=None):
        from streamlit.testing.v1 import AppTest
        from models.project import Project
        from models.invoice import Invoice

        invoices = loaded or [
            Invoice(invoice_number=8700.0, project_name="AD A", maintenance_year="Y1", payment_amount=100.0, cameras_number=1,
                    payment_date=None, paid="No", year=2026, db_id=1),
            Invoice(invoice_number=8700.0, project_name="AD B", maintenance_year="Y1", payment_amount=200.0, cameras_number=1,
                    payment_date=None, paid="No", year=2026, db_id=2),
            Invoice(invoice_number=8701.0, project_name="AD C", maintenance_year="Y1", payment_amount=300.0, cameras_number=1,
                    payment_date=None, paid="No", year=2026, db_id=3),
        ]
        projects = [Project(project_name=n, country="Bel", num_cams=1, payment_month="June", installation_year=2020, status="Active")
                    for n in ("AD A", "AD B", "AD C")]
        self.captured = {}

        def fake_editor(df, **kwargs):
            self.captured["kwargs"], self.captured["df"] = kwargs, df
            edited = df.copy()
            if manipulate:
                manipulate(edited)
            return edited

        patches = [
            patch.dict(os.environ, {"SUPABASE_URL": "http://127.0.0.1:9", "SUPABASE_SERVICE_ROLE_KEY": "dummy"}),
            patch("services.supabase_service._get_client", return_value=db),
            patch("services.supabase_service.load_projects", return_value=projects),
            patch("services.supabase_service.load_invoices", return_value=invoices),
            patch("streamlit.data_editor", fake_editor),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        app = AppTest.from_file(APP_PATH, default_timeout=120)
        app.session_state["role"] = "admin"
        app.run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        app.sidebar.radio[0].set_value("🧾 Invoice Details").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        return app

    @staticmethod
    def save_button(app):
        return next(b for b in app.button if b.key == "save_invoices_top")

    def test_the_invoice_number_column_is_read_only_in_the_editor(self):
        app = self.open_page(FakeDb(db_rows(3, number=lambda i: 8699 + i)))
        config = self.captured["kwargs"]["column_config"]
        self.assertTrue(config["Invoice #"].get("disabled"), config["Invoice #"])
        for other in ("Project", "Amount (€)", "Paid", "Description"):
            self.assertFalse((config.get(other) or {}).get("disabled", False), other)

    def test_a_normal_edit_still_saves_through_the_real_page(self):
        db = FakeDb([{"id": 1, "invoice_number": "8700", "project_name": "AD A", "maintenance_year": "Y1", "payment_amount": 100.0, "paid": "No", "year": 2026},
                     {"id": 2, "invoice_number": "8700", "project_name": "AD B", "maintenance_year": "Y1", "payment_amount": 200.0, "paid": "No", "year": 2026},
                     {"id": 3, "invoice_number": "8701", "project_name": "AD C", "maintenance_year": "Y1", "payment_amount": 300.0, "paid": "No", "year": 2026}])

        def edit_amount(df):
            df.loc[df["Project"] == "AD B", "Amount (€)"] = 999.0

        app = self.open_page(db, manipulate=edit_amount)
        self.save_button(app).click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual([e.value for e in app.error], [])
        upserts = [e for e in db.log if e[0] == "upsert"]
        self.assertEqual(len(upserts), 1)
        saved = {r["id"]: r for r in upserts[0][2]}
        self.assertEqual(saved[2]["payment_amount"], 999.0)
        self.assertEqual({i: saved[i]["invoice_number"] for i in saved}, {1: "8700", 2: "8700", 3: "8701"})

    def test_a_manipulated_client_renumbering_a_row_causes_zero_database_writes(self):
        db = FakeDb([{"id": 1, "invoice_number": "8700", "project_name": "AD A", "maintenance_year": "Y1", "payment_amount": 100.0, "paid": "No", "year": 2026},
                     {"id": 2, "invoice_number": "8700", "project_name": "AD B", "maintenance_year": "Y1", "payment_amount": 200.0, "paid": "No", "year": 2026},
                     {"id": 3, "invoice_number": "8701", "project_name": "AD C", "maintenance_year": "Y1", "payment_amount": 300.0, "paid": "No", "year": 2026}])

        def renumber_and_edit(df):
            df.loc[df["Project"] == "AD B", "Invoice #"] = 8713      # the disabled cell, changed by a crafted client
            df.loc[df["Project"] == "AD A", "Amount (€)"] = 777.0     # plus a legitimate edit that must not be saved either

        spy_upsert = patch("services.supabase_service.upsert_invoices")
        spy_delete = patch("services.supabase_service.delete_invoice_by_ids")
        upsert_mock, delete_mock = spy_upsert.start(), spy_delete.start()
        self.addCleanup(spy_upsert.stop)
        self.addCleanup(spy_delete.stop)
        app = self.open_page(db, manipulate=renumber_and_edit)
        self.save_button(app).click().run()
        errors = [e.value for e in app.error]
        self.assertTrue(any("Save blocked" in e and "cannot be changed" in e and "#8700 -> #8713" in e for e in errors), errors)
        self.assertEqual(db.writes(), [], "zero database writes")
        upsert_mock.assert_not_called()  # the save routine is never even reached
        delete_mock.assert_not_called()

    def test_a_stale_tab_is_blocked_by_the_database_check_with_zero_writes(self):
        db = FakeDb([{"id": 1, "invoice_number": "8700", "project_name": "AD A", "maintenance_year": "Y1", "payment_amount": 100.0, "paid": "No", "year": 2026},
                     {"id": 2, "invoice_number": "8715", "project_name": "AD B", "maintenance_year": "Y1", "payment_amount": 200.0, "paid": "No", "year": 2026},
                     {"id": 3, "invoice_number": "8701", "project_name": "AD C", "maintenance_year": "Y1", "payment_amount": 300.0, "paid": "No", "year": 2026}])
        app = self.open_page(db)  # the tab loaded 8700 for row 2, but the database now says 8715
        self.save_button(app).click().run()
        errors = [e.value for e in app.error]
        self.assertTrue(any("Save failed" in e and "not allowed" in e and "Nothing was saved" in e for e in errors), errors)
        self.assertEqual(db.writes(), [])
        self.assertEqual(db.tables["invoices"][1]["invoice_number"], "8715")

    def test_the_guard_runs_before_any_loaded_invoice_object_is_mutated(self):
        source = open(APP_PATH, encoding="utf-8").read()
        handler = source.split("if save_invoice_clicked:")[1]
        self.assertLess(handler.index("find_invoice_number_changes("), handler.index("inv.invoice_number = inv_no"))
        self.assertLess(handler.index("find_invoice_number_changes("), handler.index("_save_invoices(invoices, _data_path)"))
        self.assertIn("st.stop()", handler.split("find_invoice_number_changes(")[1].split("deleted_invoice_ids")[0])


if __name__ == "__main__":
    unittest.main()
