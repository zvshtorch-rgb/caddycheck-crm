"""Offline ZIP regressions; never import or execute the live application."""

import ast
from copy import deepcopy
import io
from pathlib import Path
import unittest
from unittest.mock import Mock, patch
import zipfile
import json
import base64

from models.invoice import Invoice
from services.bank_payment_documents import (
    build_invoice_zip, build_saved_payment_originals_zip, saved_payment_invoice_numbers,
)


def invoice(number, project="AD Test"):
    return Invoice(number, project, "Y1", 778.0, 1, None, "Yes", 2026)


class TestSavedPaymentInvoices(unittest.TestCase):
    def test_all_allocations_and_comma_separated_payment_numbers_are_included(self):
        payment = {"id": 1, "invoice_number": "8596,8597"}
        allocations = [
            {"invoice_number": 8596}, {"invoice_number": "8700"},
            {"invoice_number": 8709}, {"invoice_number": 8709},
        ]
        self.assertEqual(
            saved_payment_invoice_numbers(payment, allocations),
            [8596, 8597, 8700, 8709],
        )

    def test_single_and_manual_payments_need_no_source_pdf(self):
        for kind in ("pdf", "pdf-batch", "manual", "manual-partial"):
            with self.subTest(kind=kind):
                self.assertEqual(saved_payment_invoice_numbers(
                    {"invoice_number": 8687.0, "source_kind": kind}, []
                ), [8687])

    def test_missing_primary_number_uses_saved_allocations(self):
        self.assertEqual(saved_payment_invoice_numbers(
            {"id": 2, "invoice_number": 0}, [{"invoice_number": 8687}]
        ), [8687])
        self.assertEqual(saved_payment_invoice_numbers({}, []), [])

    def test_invalid_reference_is_explicit_not_truncated(self):
        for value in ("not-an-invoice", "8687.5", -1, "nan"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                saved_payment_invoice_numbers({"invoice_number": value}, [])

    def test_zip_contains_one_pdf_per_number_with_all_invoice_rows(self):
        rows = [invoice(8596, "A"), invoice(8596, "B"), invoice(8597), invoice(9999)]
        renderer = Mock(return_value=b"%PDF-test")
        result = build_invoice_zip([8597, 8596, 8596], rows, render_pdf=renderer)
        self.assertEqual(result.included, [8596, 8597])
        self.assertEqual(result.errors, [])
        self.assertEqual(renderer.call_count, 2)
        self.assertEqual(renderer.call_args_list[0].kwargs["invoice_rows"], rows[:2])
        with zipfile.ZipFile(io.BytesIO(result.data)) as archive:
            self.assertEqual(archive.namelist(), ["CC_inv_8596.pdf", "CC_inv_8597.pdf"])
            self.assertEqual(archive.read("CC_inv_8596.pdf"), b"%PDF-test")

    def test_missing_invoice_is_reported_with_partial_zip(self):
        result = build_invoice_zip(
            [8596, 8597], [invoice(8596)], render_pdf=Mock(return_value=b"%PDF-test")
        )
        self.assertEqual(result.included, [8596])
        self.assertEqual(result.errors, ["Invoice #8597: no saved invoice rows found."])
        self.assertTrue(result.data)

    def test_pdf_failure_is_reported_and_other_invoices_remain_downloadable(self):
        renderer = Mock(side_effect=[RuntimeError("renderer failed"), b"%PDF-test"])
        result = build_invoice_zip([1, 2], [invoice(1), invoice(2)], render_pdf=renderer)
        self.assertEqual(result.included, [2])
        self.assertIn("renderer failed", result.errors[0])

    def test_empty_or_failed_zip_is_not_offered_as_success(self):
        for renderer in (Mock(return_value=b""), Mock(side_effect=ValueError("failed"))):
            result = build_invoice_zip([1], [invoice(1)], render_pdf=renderer)
            self.assertEqual(result.data, b"")
            self.assertEqual(result.included, [])
            self.assertEqual(len(result.errors), 1)

    def test_financial_inputs_are_unchanged(self):
        payment = {
            "id": 7, "invoice_number": 8687, "applied_amount": 22572,
            "payment_date": "2026-10-19", "payment_fingerprint": "unchanged",
            "source_kind": "manual", "pdf_storage_path": None,
        }
        allocations = [{"invoice_number": 8687, "amount_applied": 778, "invoice_id": 20}]
        rows = [invoice(8687)]
        before = deepcopy((payment, allocations, rows))
        numbers = saved_payment_invoice_numbers(payment, allocations)
        build_invoice_zip(numbers, rows, render_pdf=Mock(return_value=b"%PDF-test"))
        self.assertEqual((payment, allocations, rows), before)


class TestSavedPaymentZipUI(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Isolate the actual rendering functions from app startup, credentials and data loaders.
        source = Path(__file__).resolve().parents[1] / "streamlit_app.py"
        cls.tree = ast.parse(source.read_text(encoding="utf-8"))
        names = {"_build_confirmed_invoices_zip", "_render_saved_payment_invoice_zip"}
        cls.functions = ast.Module(
            body=[node for node in cls.tree.body
                  if isinstance(node, ast.FunctionDef) and node.name in names],
            type_ignores=[],
        )

    def setUp(self):
        self.st = Mock()
        self.st.session_state = {}
        self.st.button.return_value = True
        self.st.spinner.return_value.__enter__ = Mock()
        self.st.spinner.return_value.__exit__ = Mock(return_value=False)
        namespace = {
            "st": self.st, "invoices": [invoice(8687)],
            "build_invoice_zip": build_invoice_zip,
            "build_saved_payment_originals_zip": build_saved_payment_originals_zip,
        }
        exec(compile(self.functions, "<isolated-bank-payment-zip-ui>", "exec"), namespace)
        self.render = namespace["_render_saved_payment_invoice_zip"]
        self.records = patch(
            "services.supabase_service.load_saved_payment_document_records",
            return_value=(
                {"id": 7, "invoice_number": 8687}, [],
                [{"id": 1, "invoice_number": 8687}],
                [{"id": 4, "invoice_number": 8687,
                  "pdf_storage_bucket": "sent-invoices", "pdf_storage_path": "8687.pdf"}],
            ),
        ).start()
        self.download = patch(
            "services.supabase_service.download_sent_invoice_pdf", return_value=b"%PDF-original"
        ).start()
        self.addCleanup(patch.stopall)

    def test_original_pdf_download_survives_rerun_without_reconfirming(self):
        payment = {"id": 7, "invoice_number": 8687}
        self.render(payment, [])
        download = self.st.download_button.call_args.kwargs
        self.assertEqual(download["mime"], "application/zip")
        with zipfile.ZipFile(io.BytesIO(download["data"])) as archive:
            self.assertEqual(archive.read("originals/CC_inv_8687_1.pdf"), b"%PDF-original")
        self.st.button.return_value = False
        self.st.download_button.reset_mock()
        self.render(payment, [])
        self.st.download_button.assert_called_once()
        self.assertEqual(self.st.download_button.call_args.kwargs["data"], download["data"])
        self.records.assert_called_once_with(7)
        self.download.assert_called_once_with("sent-invoices", "8687.pdf")

    def test_switching_payments_does_not_show_previous_zip(self):
        self.render({"id": 7, "invoice_number": 8687}, [])
        self.st.button.return_value = False
        self.st.download_button.reset_mock()
        self.render({"id": 8, "invoice_number": 8687}, [])
        self.st.download_button.assert_not_called()

    def test_partial_zip_warning_persists_on_rerun(self):
        payment = {"id": 7, "invoice_number": "8687,9999"}
        saved, allocations, rows, sent = self.records.return_value
        self.records.return_value = (payment, allocations, rows, sent)
        self.render(payment, [])
        self.st.button.return_value = False
        self.st.warning.reset_mock()
        self.render(payment, [])
        self.st.warning.assert_called_once_with("ZIP is incomplete. Missing original PDFs: #9999")
        self.assertEqual(self.st.write.call_args.args[0]["Regenerated fallbacks"], 0)

    def test_no_invoices_or_failure_never_creates_download(self):
        self.records.return_value = ({"id": 7}, [], [], [])
        self.render({"id": 7}, [])
        self.st.info.assert_called_once()
        self.records.return_value = ({"id": 7, "invoice_number": 9999}, [], [], [])
        self.render({"id": 7, "invoice_number": 9999}, [])
        self.st.warning.assert_any_call("No original invoice PDFs are available to download.")
        self.st.download_button.assert_not_called()

    def test_database_error_clears_previous_download(self):
        self.render({"id": 7}, [])
        self.st.download_button.reset_mock()
        self.records.side_effect = RuntimeError("database unavailable")
        self.render({"id": 7}, [])
        self.st.error.assert_called_with(
            "Could not retrieve linked invoice originals: database unavailable"
        )
        self.st.download_button.assert_not_called()

    def test_no_primary_key_does_not_query_or_download(self):
        self.render({"invoice_number": 8687}, [])
        self.records.assert_not_called()
        self.st.download_button.assert_not_called()

    def test_saved_payment_details_are_wired_to_download(self):
        calls = [
            node for node in ast.walk(self.tree)
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
            and node.func.id == "_render_saved_payment_invoice_zip"
        ]
        self.assertEqual(len(calls), 1)
        self.assertEqual([arg.id for arg in calls[0].args],
                         ["selected_payment_row", "allocations"])

    def test_saved_payment_section_does_not_use_self_healing_history_loaders(self):
        source = (Path(__file__).resolve().parents[1] / "streamlit_app.py").read_text(encoding="utf-8")
        section = source.split('st.subheader("Saved Bank Payments")', 1)[1]
        self.assertNotIn("load_bank_payments_log(", section)
        self.assertNotIn("load_sent_invoices_log(", section)


class TestArchivedInvoiceZip(unittest.TestCase):
    def setUp(self):
        self.payment = {"id": 7, "invoice_number": 9999, "applied_amount": 100}
        self.allocations = [
            {"id": 1, "payment_id": 7, "invoice_row_id": 11},
            {"id": 2, "payment_id": 7, "invoice_row_id": 12, "invoice_number": 8687},
            {"id": 3, "payment_id": 7, "invoice_number": 8688},
        ]
        self.rows = [
            {"id": 11, "invoice_number": 8687},
            {"id": 12, "invoice_number": 8687},
            {"id": 13, "invoice_number": 8688, "pdf_storage_bucket": "invoice-files",
             "pdf_storage_path": "8688.pdf"},
        ]
        self.sent = [
            {"id": 2, "invoice_number": 8687, "pdf_storage_bucket": "sent-invoices",
             "pdf_storage_path": "8687.pdf"},
            {"id": 3, "invoice_number": 9999, "pdf_storage_bucket": "sent-invoices",
             "pdf_storage_path": "unrelated.pdf"},
        ]
        self.download = Mock(side_effect=lambda bucket, path: b"%PDF-" + path.encode())
        self.local = Mock()

    def build(self):
        return build_saved_payment_originals_zip(
            self.payment, self.allocations, self.rows, self.sent,
            download_pdf=self.download, read_archive=self.local,
        )

    def test_row_id_resolution_all_originals_and_no_financial_changes(self):
        before = deepcopy((self.payment, self.allocations, self.rows, self.sent))
        result = self.build()
        self.assertEqual(result.linked, [8687, 8688])
        self.assertEqual(result.originals, [8687, 8688])
        self.assertEqual(result.missing, [])
        self.assertEqual(result.errors, [])
        self.assertEqual(self.download.call_count, 2)
        with zipfile.ZipFile(io.BytesIO(result.data)) as archive:
            report = json.loads(archive.read("invoice_download_report.json"))
            self.assertEqual(report["regenerated_fallbacks"], 0)
            self.assertEqual(report["invoices"][0]["allocation_invoice_row_ids"], ["11", "12"])
            self.assertEqual(archive.read("originals/CC_inv_8687_1.pdf"), b"%PDF-8687.pdf")
        self.assertEqual((self.payment, self.allocations, self.rows, self.sent), before)

    def test_missing_original_does_not_regenerate_and_partial_zip_has_report(self):
        self.sent = []
        result = self.build()
        self.assertEqual(result.originals, [8688])
        self.assertEqual(result.missing, [8687])
        with zipfile.ZipFile(io.BytesIO(result.data)) as archive:
            report = json.loads(archive.read("invoice_download_report.json"))
            self.assertEqual(report["invoices"][0]["status"], "missing_original_pdf")
            self.assertFalse(any("8687" in name for name in archive.namelist()))

    def test_storage_error_and_invalid_object_are_reported(self):
        self.download.side_effect = [RuntimeError("storage denied"), b"not a PDF"]
        result = self.build()
        self.assertEqual(result.data, b"")
        self.assertEqual(result.missing, [8687, 8688])
        self.assertIn("storage denied", result.errors[0])
        self.assertIn("not a PDF", result.errors[1])

    def test_existing_local_archive_after_storage_failure(self):
        self.sent[0]["pdf_archive_path"] = "existing-archive.pdf"
        self.download.side_effect = RuntimeError("storage unavailable")
        self.local.return_value = b"%PDF-local-original"
        result = self.build()
        self.assertEqual(result.originals, [8687])
        self.local.assert_called_once_with("existing-archive.pdf")
        self.assertTrue(result.errors)

    def test_duplicate_references_and_duplicate_bytes_are_not_duplicated(self):
        self.sent.append(dict(self.sent[0], id=4))
        self.sent.append(dict(self.sent[0], id=5, pdf_storage_path="another-copy.pdf"))
        self.download.side_effect = None
        self.download.return_value = b"%PDF-identical"
        result = self.build()
        with zipfile.ZipFile(io.BytesIO(result.data)) as archive:
            self.assertEqual(len([n for n in archive.namelist() if "8687" in n]), 1)

    def test_distinct_archived_versions_are_preserved_not_arbitrarily_selected(self):
        self.sent.append(dict(self.sent[0], id=5, pdf_storage_path="another-version.pdf"))
        result = self.build()
        self.assertEqual(result.original_pdf_count, 3)
        self.assertEqual(len(result.originals), 2)
        with zipfile.ZipFile(io.BytesIO(result.data)) as archive:
            self.assertEqual(len([n for n in archive.namelist() if "8687" in n]), 2)

    def test_cross_payment_and_conflicting_links_fail_closed(self):
        self.allocations[0]["payment_id"] = 8
        with self.assertRaisesRegex(ValueError, "does not belong"):
            self.build()
        self.allocations[0]["payment_id"] = 7
        self.allocations[0]["invoice_number"] = 9999
        with self.assertRaisesRegex(ValueError, "conflicts"):
            self.build()
        self.download.assert_not_called()

    def test_missing_row_and_number_are_explicit_unresolved_links(self):
        self.allocations = [
            {"id": 1, "payment_id": 7, "invoice_row_id": 999},
            {"id": 2, "payment_id": 7},
        ]
        result = self.build()
        self.assertEqual(result.unresolved, 2)
        self.assertEqual(len(result.errors), 2)
        self.assertEqual(result.linked, [])
        self.download.assert_not_called()

    def test_legacy_payment_number_used_only_without_allocations(self):
        self.allocations = []
        result = self.build()
        self.assertEqual(result.linked, [9999])
        self.assertEqual(result.originals, [9999])

    def test_stored_confirmation_invoice_bytes_are_originals_not_purchase_order_pdfs(self):
        self.sent = [{
            "id": "po-1", "invoice_number": 8687, "source_table": "purchase_orders",
            "archived_pdf_base64": base64.b64encode(b"%PDF-confirmation").decode(),
        }]
        result = self.build()
        with zipfile.ZipFile(io.BytesIO(result.data)) as archive:
            self.assertEqual(archive.read("originals/CC_inv_8687_1.pdf"), b"%PDF-confirmation")
            report = json.loads(archive.read("invoice_download_report.json"))
            self.assertEqual(report["invoices"][0]["documents"][0]["source_table"], "purchase_orders")

    def test_no_original_can_invoke_pdf_regeneration(self):
        with patch("services.pdf_service.generate_invoice_pdf_from_rows") as generate:
            self.sent = []
            self.rows = [{"id": 11, "invoice_number": 8687}, {"id": 12, "invoice_number": 8687}]
            result = self.build()
        generate.assert_not_called()
        self.assertEqual(result.originals, [])
        self.assertEqual(result.missing, [8687, 8688])

    def test_local_archive_read_is_confined_and_read_only(self):
        from tempfile import TemporaryDirectory
        from services.bank_payment_documents import read_archived_invoice_file

        with TemporaryDirectory(prefix="caddycheck-archive-test-") as folder:
            root = Path(folder)
            archive = root / "sent_invoices"
            archive.mkdir()
            existing = archive / "original.pdf"
            existing.write_bytes(b"%PDF-existing")
            with patch("config.settings.get_data_paths", return_value={"output_dir": root}):
                self.assertEqual(read_archived_invoice_file(str(existing)), b"%PDF-existing")
                with self.assertRaisesRegex(ValueError, "outside"):
                    read_archived_invoice_file(str(root / "outside.pdf"))
                with self.assertRaises(FileNotFoundError):
                    read_archived_invoice_file(str(archive / "missing.pdf"))
            self.assertEqual(existing.read_bytes(), b"%PDF-existing")


class TestReadOnlyArchiveQueries(unittest.TestCase):
    def test_loader_resolves_row_ids_and_paginates_without_mutations(self):
        from services.supabase_service import load_saved_payment_document_records

        tables = {
            "bank_payments": [{"id": 7, "invoice_number": 0}],
            "bank_payment_allocations": [
                {"id": 1, "payment_id": 7, "invoice_row_id": 11},
                {"id": 2, "payment_id": 8, "invoice_row_id": 22},
            ],
            "invoices": [
                {"id": 11, "invoice_number": 8687},
                {"id": 12, "invoice_number": 8687},
                {"id": 22, "invoice_number": 9999},
            ],
            "sent_invoices": [
                {"id": 1, "invoice_number": 8687},
                {"id": 2, "invoice_number": 8687},
                {"id": 3, "invoice_number": 9999},
            ],
            "purchase_orders": [
                {"id": "po-1", "confirmation_invoice_number": "8687",
                 "confirmation_pdf_base64": base64.b64encode(b"%PDF-original").decode(),
                 "pdf_storage_bucket": "purchase-order-pdfs", "pdf_storage_path": "NOT-invoice.pdf"},
                {"id": "po-2", "confirmation_invoice_number": "9999"},
            ],
        }
        calls = []

        class Query:
            def __init__(self, table):
                self.table = table
                self.rows = tables[table][:]

            def select(self, columns):
                calls.append((self.table, "select"))
                return self

            def eq(self, column, value):
                self.rows = [r for r in self.rows if str(r[column]) == str(value)]
                return self

            def order(self, column, **kwargs):
                return self

            def range(self, start, end):
                self.rows = self.rows[start:end + 1]
                return self

            def execute(self):
                response = Mock()
                response.data = self.rows
                return response

        client = Mock(spec=["table"])
        client.table.side_effect = Query
        with patch("services.supabase_service._get_client", return_value=client), \
                patch("services.supabase_service.SUPABASE_PAGE_SIZE", 1):
            saved, allocations, rows, sent = load_saved_payment_document_records(7)
        self.assertEqual(saved["id"], 7)
        self.assertEqual(len(allocations), 1)
        self.assertEqual({r["id"] for r in rows}, {11, 12})
        self.assertEqual({r["id"] for r in sent}, {1, 2, "po-1"})
        confirmation = next(r for r in sent if r["id"] == "po-1")
        self.assertEqual(confirmation["source_table"], "purchase_orders")
        self.assertNotIn("pdf_storage_path", confirmation)
        self.assertTrue(all(operation == "select" for _, operation in calls))


if __name__ == "__main__":
    unittest.main()
