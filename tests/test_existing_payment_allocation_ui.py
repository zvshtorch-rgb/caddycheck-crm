"""Run only the extracted renderer with fake loaders; never execute the live app."""
import ast
import os
from pathlib import Path
import unittest
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

from services.existing_payment_allocation import AllocationBlocked, MANUAL_ALLOCATION_GATE, choose_invoice, preview_allocation
from tests.test_existing_payment_allocation import fixture


class TestExistingPaymentAllocationUI(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        source = (Path(__file__).resolve().parents[1] / "streamlit_app.py").read_text(encoding="utf-8")
        tree = ast.parse(source)
        renderer = next(node for node in tree.body
                        if isinstance(node, ast.FunctionDef) and node.name == "_render_existing_payment_allocation")
        cls.renderer_source = ast.get_source_segment(source, renderer)

    def setUp(self):
        release = patch.dict(os.environ, {MANUAL_ALLOCATION_GATE: "on"})
        release.start()
        self.addCleanup(release.stop)
        self.p, self.a, self.i = fixture()
        self.rows = [self.i]
        self.preview_patch = patch(
            "services.supabase_service.preview_existing_payment_allocation", side_effect=self.preview,
        )
        self.lookup_patch = patch("services.supabase_service.get_invoices_by_number", side_effect=lambda _: self.rows)
        self.write_patch = patch("services.supabase_service.add_invoice_to_existing_payment",
                                 return_value={"status": "added"})
        self.gate_patch = patch("services.financial_writer_adapter.enabled", return_value=False)
        self.database_release_patch = patch(
            "services.supabase_service.require_manual_allocation_database_enabled",
        )
        self.preview_mock = self.preview_patch.start()
        self.lookup_mock = self.lookup_patch.start()
        self.write_mock = self.write_patch.start()
        self.gate_mock = self.gate_patch.start()
        self.database_release_mock = self.database_release_patch.start()
        for handle in (self.preview_patch, self.lookup_patch, self.write_patch, self.gate_patch,
                       self.database_release_patch):
            self.addCleanup(handle.stop)

    def preview(self, payment_id, number, row_id, *, role):
        return preview_allocation(self.p, self.a, choose_invoice(self.rows, row_id), [])

    def app(self, role="admin"):
        script = (
            "import logging\nimport pandas as pd\nimport streamlit as st\n"
            "logger = logging.getLogger('allocation-ui-test')\n"
            + self.renderer_source
            + f"\nstore = st.session_state.setdefault('_test_store', {{'payment': {self.p!r}, 'allocations': {self.a!r}}})\n"
            + "payment = dict(store['payment'])\nallocations = [dict(row) for row in store['allocations']]\n"
            + "chosen = st.selectbox('Payment', [14,15])\n"
            + "payment['id'] = chosen\n"
            + "for allocation in allocations: allocation['payment_id'] = chosen\n"
            + f"_render_existing_payment_allocation(payment, allocations, role={role!r})\n"
        )
        at = AppTest.from_string(script, default_timeout=10).run()
        self.assertEqual(list(at.exception), [])
        return at

    def lookup(self, at, text="8707"):
        at.text_input(key="_existing_allocation_number").set_value(text).run()
        at.button(key="_existing_allocation_lookup").click().run()
        self.assertEqual(list(at.exception), [])
        return at

    def prepared(self):
        at = self.lookup(self.app())
        at.button(key="_existing_allocation_preview_button").click().run()
        self.assertEqual(list(at.exception), [])
        return at

    def test_lookup_and_preview_do_not_write(self):
        at = self.prepared()
        self.write_mock.assert_not_called()
        self.assertFalse(at.checkbox(key="_existing_allocation_confirmed").value)
        self.assertTrue(at.button(key="_existing_allocation_add").disabled)
        preview_table = at.dataframe[-1].value
        self.assertEqual(preview_table.iloc[0]["Resulting remaining balance"], "EUR 0.00")
        self.assertEqual(preview_table.iloc[0]["Row ID"], 3087)

    def test_release_off_unset_invalid_hide_actions(self):
        for value in (None, "off", "", "invalid"):
            with self.subTest(value=value), patch.dict(os.environ, {}, clear=True):
                if value is not None:
                    os.environ[MANUAL_ALLOCATION_GATE] = value
                at = self.app()
                self.assertEqual(len(at.text_input), 0)
                self.assertFalse(any(e.label == "Add existing invoice" for e in at.expander))
                self.assertIn("release gate", at.warning[0].value)
        self.lookup_mock.assert_not_called()
        self.write_mock.assert_not_called()

    def test_disabling_release_resets_pending_confirmation(self):
        at = self.prepared()
        at.checkbox(key="_existing_allocation_confirmed").check().run()
        with patch.dict(os.environ, {MANUAL_ALLOCATION_GATE: "off"}):
            at.run()
            self.assertNotIn("_existing_allocation_preview", at.session_state)
            self.assertEqual(len(at.checkbox), 0)
        at.run()
        self.assertNotIn("_existing_allocation_preview", at.session_state)
        self.write_mock.assert_not_called()

    def test_database_off_or_unavailable_hides_actions_even_when_environment_on(self):
        for error in (AllocationBlocked("Database release is OFF"), RuntimeError("unavailable")):
            with self.subTest(error=type(error).__name__):
                self.database_release_mock.side_effect = error
                at = self.app()
                self.assertEqual(len(at.text_input), 0)
                self.assertFalse(any(e.label == "Add existing invoice" for e in at.expander))
                self.assertIn("database release policy", at.warning[0].value)
        self.lookup_mock.assert_not_called()
        self.write_mock.assert_not_called()

    def test_database_disabling_resets_pending_confirmation(self):
        at = self.prepared()
        at.checkbox(key="_existing_allocation_confirmed").check().run()
        self.database_release_mock.side_effect = AllocationBlocked("Database release is OFF")
        at.run()
        self.assertNotIn("_existing_allocation_preview", at.session_state)
        self.assertEqual(len(at.checkbox), 0)
        self.database_release_mock.side_effect = None
        at.run()
        self.assertNotIn("_existing_allocation_preview", at.session_state)
        self.write_mock.assert_not_called()

    def test_invalid_financial_writer_gate_hides_actions(self):
        from services.financial_writer_adapter import FinancialWriteBlocked
        self.gate_mock.side_effect = FinancialWriteBlocked("Invalid compatibility gate")
        at = self.app()
        self.assertEqual(len(at.text_input), 0)
        self.database_release_mock.assert_not_called()
        self.write_mock.assert_not_called()

    def test_explicit_confirmation_required(self):
        at = self.prepared()
        at.checkbox(key="_existing_allocation_confirmed").check().run()
        self.assertFalse(at.button(key="_existing_allocation_add").disabled)
        self.write_mock.assert_not_called()
        at.button(key="_existing_allocation_add").click().run()
        self.assertEqual(list(at.exception), [])
        self.write_mock.assert_called_once()
        self.assertEqual(self.write_mock.call_args.kwargs, {"role": "admin", "confirmed": True})
        self.assertIn("added and verified", at.success[0].value)
        self.assertEqual(len(at.checkbox), 0)

    def test_success_displays_updated_reconciliation(self):
        at = self.prepared()

        def fake_commit(*args, **kwargs):
            import streamlit as st
            store = st.session_state["_test_store"]
            store["payment"]["applied_amount"] = 46680
            store["payment"]["invoice_number"] = "8596,8597,8700,8707,8709"
            store["allocations"].append({
                "id": 77, "payment_id": 14, "invoice_row_id": 3087, "invoice_number": 8707,
                "project_name": "Intermarche Heusy", "amount_applied": 7780,
            })
            return {"status": "added"}

        self.write_mock.side_effect = fake_commit
        at.checkbox(key="_existing_allocation_confirmed").check().run()
        at.button(key="_existing_allocation_add").click().run()
        self.assertEqual(list(at.exception), [])
        self.assertTrue(any(
            "Allocated total: EUR 46,680.00 | Remaining instructed balance: EUR 0.00" in element.value
            for element in at.markdown
        ))

    def test_confirmation_reset_when_invoice_changes(self):
        at = self.prepared()
        at.checkbox(key="_existing_allocation_confirmed").check().run()
        at.text_input(key="_existing_allocation_number").set_value("8708").run()
        self.assertEqual(list(at.exception), [])
        self.assertNotIn("_existing_allocation_preview", at.session_state)
        self.assertFalse(at.session_state["_existing_allocation_confirmed"])
        self.write_mock.assert_not_called()

    def test_confirmation_reset_when_payment_changes(self):
        at = self.prepared()
        at.checkbox(key="_existing_allocation_confirmed").check().run()
        at.selectbox[0].set_value(15).run()
        self.assertEqual(list(at.exception), [])
        self.assertNotIn("_existing_allocation_preview", at.session_state)
        self.assertFalse(at.session_state["_existing_allocation_confirmed"])
        self.write_mock.assert_not_called()

    def test_multiple_rows_require_selection_and_row_change_resets_confirmation(self):
        self.rows = [self.i, {**self.i, "id": 3088, "project_name": "Other row"}]
        at = self.lookup(self.app())
        self.assertIsNone(at.selectbox(key="_existing_allocation_row").value)
        self.assertFalse(any(button.key == "_existing_allocation_preview_button" for button in at.button))
        at.selectbox(key="_existing_allocation_row").set_value(3087).run()
        at.button(key="_existing_allocation_preview_button").click().run()
        at.checkbox(key="_existing_allocation_confirmed").check().run()
        at.selectbox(key="_existing_allocation_row").set_value(3088).run()
        self.assertEqual(list(at.exception), [])
        self.assertNotIn("_existing_allocation_preview", at.session_state)
        self.assertFalse(at.session_state["_existing_allocation_confirmed"])

    def test_admin_only_and_financial_writer_disabled(self):
        at = self.app("viewer")
        self.assertEqual(len(at.text_input), 0)
        self.assertIn("Only an admin", at.warning[0].value)
        self.gate_mock.return_value = True
        at = self.app()
        self.assertEqual(len(at.text_input), 0)
        self.assertIn("financial-writer", at.warning[0].value)
        self.lookup_mock.assert_not_called()
        self.write_mock.assert_not_called()

    def test_missing_and_non_payable_invoices_show_errors_no_confirmation(self):
        self.rows = []
        at = self.lookup(self.app())
        self.assertIn("No existing invoice", at.error[0].value)
        self.rows = [{**self.i, "paid": "cancelled"}]
        at = self.lookup(self.app())
        at.button(key="_existing_allocation_preview_button").click().run()
        self.assertIn("non-payable", at.error[0].value)
        self.assertEqual(len(at.checkbox), 0)
        self.write_mock.assert_not_called()

    def test_stale_or_uncertain_result_requires_a_new_preview(self):
        self.write_mock.side_effect = RuntimeError("Preview is stale")
        at = self.prepared()
        at.checkbox(key="_existing_allocation_confirmed").check().run()
        with self.assertLogs("allocation-ui-test", level="ERROR"):
            at.button(key="_existing_allocation_add").click().run()
        self.assertEqual(list(at.exception), [])
        self.assertIn("reloaded", at.error[0].value)
        self.assertNotIn("_existing_allocation_preview", at.session_state)
        self.assertEqual(len(at.checkbox), 0)

    def test_no_sending_or_renewal_surface(self):
        forbidden = (
            "mark_invoice_row_paid", "append_bank_payment", "create_renewal_link",
            "upsert_subscription", "send_email", "generate_invoice_pdf",
        )
        for name in forbidden:
            self.assertNotIn(name, self.renderer_source)
