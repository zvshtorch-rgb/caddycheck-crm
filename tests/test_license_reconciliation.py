"""Tests for the read-only License EOP reconciliation (services/license_reconciliation.py)
and its registration as an Ask Data tool in streamlit_app.py."""
import datetime
import unittest
from pathlib import Path
from types import SimpleNamespace

from services.license_reconciliation import (
    add_years,
    build_license_reconciliation,
    filter_reconciliation,
    reconcile_project,
)

D = datetime.date


def project(name="P", activation=D(2025, 3, 12), eop=D(2026, 11, 1), status="Active", country="BE"):
    return SimpleNamespace(
        project_name=name, activation_date=activation, license_eop=eop, status=status,
        country=country, installation_year=None,
    )


def inv(label, paid="Yes", number=1, year=None, pay_date=None, inv_type="", name="P", start_year=2025):
    k = int(label[1:])
    return SimpleNamespace(
        project_name=name, maintenance_year=label, paid=paid, invoice_number=number,
        year=year if year is not None else start_year + k - 1,
        payment_date=pay_date, invoice_type=inv_type,
    )


class TestExpectedEop(unittest.TestCase):
    def test_year1_paid_only(self):
        r = reconcile_project(project(eop=D(2025, 12, 1)), [inv("Y1", number=10, pay_date=D(2025, 4, 1))])
        self.assertEqual(r["Highest Paid Maintenance Year"], "Y1")
        self.assertEqual(r["Expected License EOP"], "2026-03-12")
        self.assertEqual(r["Needs EOP Update?"], "Yes")
        self.assertEqual(r["Y1 Invoice Number"], "10")
        self.assertEqual(r["Y1 Payment Date"], "2025-04-01")
        self.assertEqual(r["Y2 Paid?"], "No invoice")

    def test_year1_and_year2_paid(self):
        r = reconcile_project(project(eop=D(2026, 11, 1)), [inv("Y1", number=10), inv("Y2", number=20)])
        self.assertEqual(r["Highest Paid Maintenance Year"], "Y2")
        self.assertEqual(r["Expected License EOP"], "2027-03-12")
        self.assertEqual(r["Needs EOP Update?"], "Yes")
        self.assertEqual(r["Confidence"], "High")

    def test_eop_already_updated_is_not_flagged(self):
        r = reconcile_project(project(eop=D(2027, 3, 1)), [inv("Y1"), inv("Y2")])
        self.assertEqual(r["Needs EOP Update?"], "No")

    def test_later_eop_is_never_flagged(self):
        r = reconcile_project(project(eop=D(2029, 1, 1)), [inv("Y1"), inv("Y2")])
        self.assertEqual(r["Needs EOP Update?"], "No")
        self.assertIn("Later", r["EOP vs Expected"])

    def test_first_of_month_style_eop_counts_as_updated(self):
        r = reconcile_project(project(activation=D(2025, 3, 12), eop=D(2027, 3, 1)), [inv("Y1"), inv("Y2")])
        self.assertEqual(r["Needs EOP Update?"], "No")

    def test_year2_paid_without_year1_does_not_count(self):
        r = reconcile_project(project(), [inv("Y1", paid="No"), inv("Y2")])
        self.assertEqual(r["Highest Paid Maintenance Year"], "None")
        self.assertEqual(r["Needs EOP Update?"], "No")
        self.assertIn("not consecutive", r["Notes"])

    def test_partial_year_is_not_paid(self):
        r = reconcile_project(project(), [inv("Y1"), inv("Y2", number=1), inv("Y2", paid="No", number=2)])
        self.assertEqual(r["Y2 Paid?"], "Partial")
        self.assertEqual(r["Highest Paid Maintenance Year"], "Y1")

    def test_complementary_unpaid_row_is_ignored(self):
        r = reconcile_project(project(), [inv("Y1"), inv("Y1", paid="No", inv_type="Complementary", number=2)])
        self.assertEqual(r["Y1 Paid?"], "Yes")

    def test_cancelled_row_is_ignored(self):
        r = reconcile_project(project(), [inv("Y1"), inv("Y1", paid="Cancelled", number=2)])
        self.assertEqual(r["Y1 Paid?"], "Yes")

    def test_missing_eop_needs_update(self):
        r = reconcile_project(project(eop=None), [inv("Y1")])
        self.assertEqual(r["Needs EOP Update?"], "Yes")
        self.assertEqual(r["EOP vs Expected"], "Missing EOP")

    def test_no_paid_year_means_no_update(self):
        r = reconcile_project(project(), [inv("Y1", paid="No")])
        self.assertEqual(r["Needs EOP Update?"], "No")
        self.assertEqual(r["Confidence"], "n/a")

    def test_inconsistent_history_is_low_confidence(self):
        r = reconcile_project(project(activation=D(2024, 12, 9)), [inv("Y1", year=2020), inv("Y2", year=2021)])
        self.assertEqual(r["Confidence"], "Low")
        self.assertIn("verify manually", r["Notes"])

    def test_missing_activation_date_falls_back_to_installation_year(self):
        p = project(activation=None)
        p.installation_year = 2025
        r = reconcile_project(p, [inv("Y1")])
        self.assertTrue(r["License Start Date"].endswith("(est.)"))

    def test_cancelled_project_is_noted(self):
        r = reconcile_project(project(eop=D(2025, 1, 1), status="cancelled"), [inv("Y1")], lambda s: s.title())
        self.assertEqual(r["Needs EOP Update?"], "Yes")
        self.assertIn("Cancelled", r["Notes"])

    def test_leap_day_start(self):
        self.assertEqual(add_years(D(2024, 2, 29), 1), D(2025, 2, 28))


class TestBuildAndFilter(unittest.TestCase):
    def setUp(self):
        projects = [
            project("A", eop=D(2025, 12, 1)),   # Y1 paid, EOP not updated
            project("B", eop=D(2026, 11, 1)),   # Y1+Y2 paid, EOP not updated
            project("C", eop=D(2027, 6, 1)),    # Y1+Y2 paid, EOP updated
            project("D", eop=D(2026, 1, 1)),    # nothing paid
        ]
        invoices = [
            inv("Y1", name="A", number=1),
            inv("Y1", name="B", number=2), inv("Y2", name="B", number=3),
            inv("Y1", name="C", number=4), inv("Y2", name="C", number=5),
            inv("Y1", name="D", paid="No", number=6),
        ]
        self.rows = build_license_reconciliation(projects, invoices)

    def test_one_row_per_project_needs_update_first(self):
        self.assertEqual(len(self.rows), 4)
        self.assertEqual([r["Project"] for r in self.rows[:2]], ["A", "B"])

    def test_filter_paid_year_1(self):
        names = {r["Project"] for r in filter_reconciliation(self.rows, paid_year=1)}
        self.assertEqual(names, {"A"})  # B's EOP already covers Y1 (only Y2 is missing)

    def test_filter_paid_year_2(self):
        names = {r["Project"] for r in filter_reconciliation(self.rows, paid_year=2)}
        self.assertEqual(names, {"B"})

    def test_filter_needs_update(self):
        yes = {r["Project"] for r in filter_reconciliation(self.rows, needs_update="yes")}
        no = {r["Project"] for r in filter_reconciliation(self.rows, needs_update="no")}
        self.assertEqual(yes, {"A", "B"})
        self.assertEqual(no, {"C", "D"})

    def test_filter_project(self):
        self.assertEqual([r["Project"] for r in filter_reconciliation(self.rows, project="C")], ["C"])

    def test_inputs_are_not_mutated(self):
        p = project("Z", eop=D(2025, 12, 1))
        before = p.license_eop
        build_license_reconciliation([p], [inv("Y1", name="Z")])
        self.assertEqual(p.license_eop, before)


class TestToolRegistration(unittest.TestCase):
    def test_tool_registered_and_read_only(self):
        src = (Path(__file__).resolve().parent.parent / "streamlit_app.py").read_text(encoding="utf-8")
        self.assertIn('"get_license_eop_reconciliation": {', src)
        self.assertIn('if tool_name == "get_license_eop_reconciliation":', src)
        self.assertIn("build_license_reconciliation", src)
        branch = src.split('if tool_name == "get_license_eop_reconciliation":', 1)[1].split('if tool_name == "get_camera_statistics":', 1)[0]
        for forbidden in ("save_project", "update_project", "upsert", ".update(", ".insert(", "save_invoice"):
            self.assertNotIn(forbidden, branch)

    def test_service_module_has_no_db_writes(self):
        src = (Path(__file__).resolve().parent.parent / "services" / "license_reconciliation.py").read_text(encoding="utf-8")
        for forbidden in ("supabase", "save_", "upsert", ".update(", ".insert("):
            self.assertNotIn(forbidden, src)


if __name__ == "__main__":
    unittest.main()
