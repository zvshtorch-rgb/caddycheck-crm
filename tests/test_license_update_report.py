import ast
import datetime
import io
import unittest
from pathlib import Path

import pdfplumber

from models.invoice import Invoice
from models.project import Project
from services.license_update_report import (
    REPORT_COLUMNS,
    build_license_update_pdf,
    build_license_update_rows,
)


def project(name, status="Active", eop=datetime.datetime(2026, 12, 1), **kwargs):
    return Project(
        project_name=name, country="Bel", num_cams=6, status=status,
        license_eop=eop, payment_month="December", installation_year=2024, **kwargs,
    )


def invoice(name, paid, year=2025):
    return Invoice(
        invoice_number=123, project_name=name, maintenance_year="Y2",
        payment_amount=100, cameras_number=6, payment_date=None,
        paid=paid, year=year,
    )


class LicenseUpdateReportTests(unittest.TestCase):
    def test_latest_due_annual_payment_status_and_excluded_projects(self):
        projects = [
            project("Paid"),
            project("Unpaid"),
            project("Missing"),
            project("Cancelled", status="Cancelled"),
            project("Other month", eop=datetime.datetime(2027, 1, 1)),
        ]
        invoices = [
            invoice("Paid", "Yes"), invoice("Paid", "No", year=2024),
            invoice("Paid", "No", year=2026),
            invoice("Unpaid", "No"), invoice("Missing", "cancelled"),
        ]
        rows = build_license_update_rows(projects, invoices, datetime.date(2026, 11, 26))
        self.assertEqual([row["Project"] for row in rows], ["Paid", "Unpaid", "Missing"])
        self.assertEqual(
            [row["Annual Payment Status"] for row in rows],
            ["Paid", "Unpaid", "No Annual Invoice Found"],
        )
        self.assertEqual(tuple(rows[0]), REPORT_COLUMNS)
        self.assertEqual(rows[0]["License EOP"], "2026-12-01")

    def test_december_rollover_and_not_yet_due(self):
        projects = [
            Project(
                project_name="New", status="Active", license_eop=datetime.datetime(2027, 1, 1),
                payment_month="January", installation_year=2027,
            ),
            project("December"),
        ]
        rows = build_license_update_rows(projects, [], datetime.date(2026, 12, 15))
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Project"], "New")
        self.assertEqual(rows[0]["Annual Payment Status"], "Not Yet Due")

    def test_explicit_and_unknown_payment_states(self):
        projects = [project(name) for name in (
            "Paid", "Unpaid", "Blank", "Null", "Missing", "Ambiguous", "No invoice",
        )]
        invoices = [
            invoice("Paid", " Yes "), invoice("Unpaid", "No"),
            invoice("Blank", ""), invoice("Null", None), invoice("Missing", "Yes"),
            invoice("Ambiguous", "pending"),
        ]
        del invoices[4].paid
        rows = build_license_update_rows(projects, invoices, datetime.date(2026, 11, 26))
        self.assertEqual(
            [row["Annual Payment Status"] for row in rows],
            ["Paid", "Unpaid", "Unknown", "Unknown", "Unknown", "Unknown", "No Annual Invoice Found"],
        )

    def test_mixed_invoice_states_never_imply_paid(self):
        projects = [project("Mixed"), project("Unpaid"), project("Cancelled"), project("Unknown and unpaid")]
        invoices = [
            invoice("Mixed", "Yes"), invoice("Mixed", ""),
            invoice("Unpaid", "Yes"), invoice("Unpaid", "No"),
            invoice("Cancelled", "cancelled"), invoice("Cancelled", "Yes"),
            invoice("Unknown and unpaid", "No"), invoice("Unknown and unpaid", None),
        ]
        rows = build_license_update_rows(projects, invoices, datetime.date(2026, 11, 26))
        self.assertEqual([row["Annual Payment Status"] for row in rows], ["Unknown", "Unpaid", "Paid", "Unknown"])

    def test_pdf_retains_unicode_project_name_and_row_status(self):
        name = "Store Łódź"
        rows = build_license_update_rows(
            [project(name)], [invoice(name, "No")], datetime.date(2026, 11, 26),
        )
        with pdfplumber.open(io.BytesIO(build_license_update_pdf(rows, 2026, 12))) as document:
            text = "\n".join(page.extract_text() for page in document.pages)
        row_line = next(line for line in text.splitlines() if name in line)
        self.assertIn(rows[0]["Annual Payment Status"], row_line)

    def test_pdf_rejects_unsupported_characters_instead_of_substituting(self):
        rows = build_license_update_rows(
            [project("Store \U0001f600")], [], datetime.date(2026, 11, 26),
        )
        with self.assertRaisesRegex(ValueError, "U\\+1F600"):
            build_license_update_pdf(rows, 2026, 12)

    def test_pdf_matches_filtered_rows_and_payment_statuses(self):
        projects = [
            project("Alpha"), project("Bravo"), project("Charlie"),
            project("Excluded", status="Cancelled"),
        ]
        invoices = [invoice("Alpha", "Yes"), invoice("Bravo", "No")]
        rows = build_license_update_rows(projects, invoices, datetime.date(2026, 11, 26))
        with pdfplumber.open(io.BytesIO(build_license_update_pdf(rows, 2026, 12))) as document:
            text = "\n".join(page.extract_text() for page in document.pages)
        for row in rows:
            row_line = next(line for line in text.splitlines() if row["Project"] in line)
            self.assertIn(row["Annual Payment Status"], row_line)
        self.assertNotIn("Excluded", text)

    def test_ui_table_and_pdf_use_same_rows(self):
        app = Path(__file__).resolve().parent.parent / "streamlit_app.py"
        tree = ast.parse(app.read_text(encoding="utf-8"))
        calls = [node for node in ast.walk(tree) if isinstance(node, ast.Call)]
        pdf_calls = [
            node for node in calls
            if isinstance(node.func, ast.Name) and node.func.id == "build_license_update_pdf"
        ]
        self.assertEqual(len(pdf_calls), 1)
        self.assertIsInstance(pdf_calls[0].args[0], ast.Name)
        self.assertEqual(pdf_calls[0].args[0].id, "next_month_rows")
        table_calls = [
            node for node in calls
            if isinstance(node.func, ast.Attribute)
            and node.func.attr == "DataFrame"
            and node.args
            and isinstance(node.args[0], ast.Name)
            and node.args[0].id == "next_month_rows"
        ]
        self.assertEqual(len(table_calls), 1)
        self.assertTrue(any(
            keyword.arg == "columns"
            and isinstance(keyword.value, ast.Name)
            and keyword.value.id == "LICENSE_UPDATE_REPORT_COLUMNS"
            for keyword in table_calls[0].keywords
        ))

    def test_pdf_contains_table_data_and_repeats_header_across_pages(self):
        projects = [project(f"Store {i:03d} & <West>") for i in range(90)]
        invoices = [invoice(p.project_name, "Yes") for p in projects]
        rows = build_license_update_rows(projects, invoices, datetime.date(2026, 11, 26))
        pdf = build_license_update_pdf(rows, 2026, 12)
        self.assertTrue(pdf.startswith(b"%PDF"))
        with pdfplumber.open(io.BytesIO(pdf)) as document:
            self.assertGreater(len(document.pages), 1)
            pages = [page.extract_text() for page in document.pages]
        self.assertIn("Projects Needing Update in December 2026", pages[0])
        self.assertIn("latest due annual invoice", pages[0])
        self.assertTrue(all("Annual Payment" in text for text in pages))
        text = "\n".join(pages)
        self.assertIn("Store 000 & <West>", text)
        self.assertIn("Store 089 & <West>", text)
        self.assertEqual(text.count("Paid"), len(rows))


if __name__ == "__main__":
    unittest.main()
