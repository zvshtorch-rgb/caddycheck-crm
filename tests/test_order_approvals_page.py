"""Streamlit AppTest of the Order Approvals page: failed-delivery warning and admin Retry button.

Runs the real app script with all network-facing calls mocked. Skipped when Streamlit is not installed.
"""
import importlib.util
import os
import unittest
from unittest.mock import patch

from services import order_approval_service as oas
from services.order_approval_service import _project_name_looks_unreliable

HAS_STREAMLIT = importlib.util.find_spec("streamlit") is not None
APP_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "streamlit_app.py")

ORDER = {
    "id": "po-1", "status": "pending_approval", "project_name": "Test Project",
    "customer_email": "buyer@example.com", "incoming_email_id": "in-1",
    "created_at": "2026-10-15T09:00:00", "pdf_storage_bucket": "", "pdf_storage_path": "",
}
FAILED = {"intake_status": "error", "attempt_count": oas.MAX_INTAKE_ATTEMPTS,
          "error_message": "approval_email: RuntimeError: sendMail 503", "email_status": "failed"}
HEALTHY = {"intake_status": "processed", "attempt_count": 0, "email_status": "sent"}


@unittest.skipUnless(HAS_STREAMLIT, "streamlit is not installed")
class TestOrderApprovalsRetryButton(unittest.TestCase):
    def open_page(self, role, delivery, retry_result=True):
        from streamlit.testing.v1 import AppTest

        patches = [
            patch("services.supabase_service.load_projects", return_value=[]),
            patch("services.supabase_service.load_invoices", return_value=[]),
            patch.object(oas, "list_purchase_orders", return_value=[ORDER]),
            patch.object(oas, "get_order_delivery_state", return_value=delivery),
            patch.object(oas, "create_purchase_order_signed_url", return_value=None),
            patch.object(oas, "count_unread_notifications", return_value=0),
            patch.object(oas, "list_notifications", return_value=[]),
            patch.object(oas, "retry_order_intake", return_value=retry_result),
        ]
        mocks = [p.start() for p in patches]
        self.addCleanup(lambda: [p.stop() for p in patches])
        app = AppTest.from_file(APP_PATH, default_timeout=90)
        app.session_state["role"] = role
        app.run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        app.sidebar.radio[0].set_value("✅ Order Approvals").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        return app, mocks[-1]

    @staticmethod
    def retry_buttons(app):
        return [b for b in app.button if b.key == "retry_intake_po-1"]

    def test_admin_sees_warning_and_retry_queues_the_order(self):
        app, retry = self.open_page("admin", FAILED)
        warnings = [w.value for w in app.warning]
        self.assertTrue(any("has not been delivered" in w and "exhausted" in w and "sendMail 503" in w
                            for w in warnings), warnings)
        self.assertEqual(len(self.retry_buttons(app)), 1)

        self.retry_buttons(app)[0].click().run()
        retry.assert_called_once_with("po-1")
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertTrue(any("Retry queued" in t.value for t in app.toast), [t.value for t in app.toast])

    def test_failed_retry_is_reported_without_error(self):
        app, retry = self.open_page("admin", FAILED, retry_result=False)
        self.retry_buttons(app)[0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertTrue(any("Nothing to retry" in t.value for t in app.toast), [t.value for t in app.toast])

    def test_viewer_sees_the_warning_but_cannot_retry(self):
        app, retry = self.open_page("viewer", FAILED)
        self.assertTrue(any("has not been delivered" in w.value for w in app.warning))
        self.assertEqual(self.retry_buttons(app), [])
        retry.assert_not_called()

    def test_page_survives_a_stale_service_module_without_the_new_function(self):
        saved = oas.get_order_delivery_state
        del oas.get_order_delivery_state
        self.addCleanup(setattr, oas, "get_order_delivery_state", saved)
        from streamlit.testing.v1 import AppTest

        patches = [
            patch("services.supabase_service.load_projects", return_value=[]),
            patch("services.supabase_service.load_invoices", return_value=[]),
            patch.object(oas, "list_purchase_orders", return_value=[ORDER]),
            patch.object(oas, "create_purchase_order_signed_url", return_value=None),
            patch.object(oas, "count_unread_notifications", return_value=0),
            patch.object(oas, "list_notifications", return_value=[]),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        app = AppTest.from_file(APP_PATH, default_timeout=90)
        app.session_state["role"] = "admin"
        app.run()
        app.sidebar.radio[0].set_value("✅ Order Approvals").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertTrue(any("Pending approval" in e.label for e in app.expander))

    def test_healthy_order_shows_no_warning_or_retry(self):
        app, _ = self.open_page("admin", HEALTHY)
        self.assertFalse(any("has not been delivered" in w.value for w in app.warning))
        self.assertEqual(self.retry_buttons(app), [])


class TestProjectNameLooksUnreliable(unittest.TestCase):
    """Pure-function tests for the filename-derived-project-name quality check (regression guard for the
    production incident where a PDF named 'document_iretailche_20260923_2026-0047-2.pdf' -- not the usual
    'YYYY-MM-NNN_ProjectName.pdf' convention -- produced the literal garbage project name
    'document iretailche 20260923 2026-0047-2', which then flowed untouched into the auto-created
    invoice/project/order on approval)."""

    def test_normal_project_names_are_not_flagged(self):
        self.assertFalse(_project_name_looks_unreliable("Edeka Steinebach Hohn"))
        self.assertFalse(_project_name_looks_unreliable("AD Anderlecht"))

    def test_a_store_number_alone_does_not_false_positive(self):
        """A real store/unit number (e.g. 'Hohn 3') must not be confused with an unstripped reference."""
        self.assertFalse(_project_name_looks_unreliable("Edeka Steinebach Hohn 3"))

    def test_generic_leading_word_is_flagged(self):
        self.assertTrue(_project_name_looks_unreliable("document iretailche 20260923 2026-0047-2"))
        self.assertTrue(_project_name_looks_unreliable("Invoice 12345"))
        self.assertTrue(_project_name_looks_unreliable("scan 001"))

    def test_unstripped_long_digit_run_is_flagged_even_without_a_generic_word(self):
        self.assertTrue(_project_name_looks_unreliable("Order 20260923"))

    def test_empty_or_missing_name_is_flagged(self):
        self.assertTrue(_project_name_looks_unreliable(""))
        self.assertTrue(_project_name_looks_unreliable(None))


@unittest.skipUnless(HAS_STREAMLIT, "streamlit is not installed")
class TestOrderApprovalsUnreliableNameWarningAndEdit(unittest.TestCase):
    """Order Approvals page: an admin sees a warning + an editable field for a filename-derived project
    name that looks unreliable, and can correct it (updating only project_name on the purchase order,
    never approving/creating anything) before deciding."""

    UNRELIABLE_ORDER = dict(ORDER, id="po-2", project_name="document iretailche 20260923 2026-0047-2")

    def open_page(self, order, role="admin"):
        from streamlit.testing.v1 import AppTest

        patches = [
            patch("services.supabase_service.load_projects", return_value=[]),
            patch("services.supabase_service.load_invoices", return_value=[]),
            patch.object(oas, "list_purchase_orders", return_value=[order]),
            patch.object(oas, "get_order_delivery_state", return_value=HEALTHY),
            patch.object(oas, "create_purchase_order_signed_url", return_value=None),
            patch.object(oas, "count_unread_notifications", return_value=0),
            patch.object(oas, "list_notifications", return_value=[]),
            patch.object(oas, "update_purchase_order"),
        ]
        mocks = [p.start() for p in patches]
        self.addCleanup(lambda: [p.stop() for p in patches])
        app = AppTest.from_file(APP_PATH, default_timeout=90)
        app.session_state["role"] = role
        app.run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        app.sidebar.radio[0].set_value("✅ Order Approvals").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        return app, mocks[-1]

    def test_unreliable_name_shows_a_warning_and_an_editable_field(self):
        app, _ = self.open_page(self.UNRELIABLE_ORDER)
        self.assertTrue(any("looks unreliable" in w.value for w in app.warning), [w.value for w in app.warning])
        name_inputs = [t for t in app.text_input if t.key == "edit_name_po-2"]
        self.assertEqual(len(name_inputs), 1)
        self.assertEqual(name_inputs[0].value, "document iretailche 20260923 2026-0047-2")

    def test_normal_name_shows_no_warning(self):
        app, _ = self.open_page(ORDER)
        self.assertFalse(any("looks unreliable" in w.value for w in app.warning))

    def test_saving_a_corrected_name_updates_only_project_name_and_does_not_approve(self):
        app, update_mock = self.open_page(self.UNRELIABLE_ORDER)
        app.text_input(key="edit_name_po-2").set_value("Edeka Steinebach Hohn").run()
        [b for b in app.button if b.key == "save_name_po-2"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        update_mock.assert_called_once_with("po-2", project_name="Edeka Steinebach Hohn")

    def test_viewer_does_not_see_the_editable_field(self):
        app, _ = self.open_page(self.UNRELIABLE_ORDER, role="viewer")
        self.assertEqual([t for t in app.text_input if t.key == "edit_name_po-2"], [])


@unittest.skipUnless(HAS_STREAMLIT, "streamlit is not installed")
class TestOrderConfirmationDryRunButton(unittest.TestCase):
    """Order Approvals page: the "Preview order confirmation (dry run)" button renders the preview
    (errors/warnings/sender/To/CC/subject/body/PDF download) without ever creating an invoice, updating
    the purchase order, or sending an email."""

    PREVIEW_ORDER = dict(ORDER, id="po-3")
    OK_PREVIEW = {
        "ok": True, "errors": [], "warnings": ["Country is unknown for this project -- it will be omitted."],
        "sender": "orders.caddycheck@video-inform.com", "to": ["tom@example.com", "karen@example.com"],
        "cc": ["yoram@example.com"], "project_name": "Test Project", "amount": 2334.0,
        "invoice_number_preview": 8719, "cameras": 12,
        "subject": "Order Confirmation \u2013 CaddyCheck \u2013 Test Project \u2013 Invoice 8719",
        "body": "Dear Tom and Karen,\n\nThank you for your order.",
        "pdf_bytes": b"%PDF-1.4 fake",
        "mailbox_check": {"ok": True, "message": "Directory user found; send authorization not verified.",
                          "diagnostics": {"method": "GET", "http_status": 200}},
        "pdf_validation": {"byte_count": 3335, "page_count": 1},
        "config_diagnostics": {"sources": {"cc": "st.secrets['order_confirmation']['cc']"},
                               "secret_values": {"cc": ["yoram@example.com"]}},
    }
    FAILED_PREVIEW = {
        "ok": False, "errors": ["Neither a camera count nor an explicit amount is available on this order."],
        "warnings": [],
    }

    def open_page(self, preview_result, role="admin"):
        from streamlit.testing.v1 import AppTest

        patches = [
            patch("services.supabase_service.load_projects", return_value=[]),
            patch("services.supabase_service.load_invoices", return_value=[]),
            patch.object(oas, "list_purchase_orders", return_value=[self.PREVIEW_ORDER]),
            patch.object(oas, "get_order_delivery_state", return_value=HEALTHY),
            patch.object(oas, "create_purchase_order_signed_url", return_value=None),
            patch.object(oas, "count_unread_notifications", return_value=0),
            patch.object(oas, "list_notifications", return_value=[]),
            patch.object(oas, "update_purchase_order"),
            patch.object(oas, "auto_create_invoice_for_approved_order"),
            patch("services.order_confirmation_service.build_dry_run_preview", return_value=preview_result,
                  side_effect=preview_result if isinstance(preview_result, Exception) else None),
            patch("services.supabase_service.replace_invoice_rows"),   # index -2: must stay un-called
            patch("services.email_service.send_graph_email"),         # index -1: must stay un-called
        ]
        mocks = [p.start() for p in patches]
        self.addCleanup(lambda: [p.stop() for p in patches])
        app = AppTest.from_file(APP_PATH, default_timeout=90)
        app.session_state["role"] = role
        app.run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        app.sidebar.radio[0].set_value("✅ Order Approvals").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        return app, {"replace_rows": mocks[-2], "send_email": mocks[-1],
                     "legacy_auto_invoice": mocks[8], "update_order": mocks[7]}

    def test_preview_button_exists_for_admin(self):
        app, _ = self.open_page(self.OK_PREVIEW)
        self.assertTrue(any(b.key == "preview_confirm_po-3" for b in app.button))

    def test_internal_send_requires_confirmation_and_runs_only_once_per_session(self):
        from services import order_confirmation_service as ocs

        cfg = {"ok": True, "errors": [], "sender": ocs.TEST_CONFIRMATION_SENDER,
               "recipient": "zvika-test@video-inform.com", "subject": ocs.TEST_CONFIRMATION_SUBJECT}
        with patch.object(ocs, "get_confirmation_send_test_config", return_value=cfg), \
             patch.object(ocs, "send_order_confirmation_test",
                          return_value={"accepted": True, "http_status": 202}) as send:
            app, mocks = self.open_page(self.OK_PREVIEW)
            self.assertTrue(app.button(key="send_internal_graph_test").disabled)
            send.assert_not_called()
            app.checkbox(key="confirm_internal_graph_test").check().run()
            app.button(key="send_internal_graph_test").click().run()
            self.assertFalse(app.exception)
            self.assertTrue(app.button(key="send_internal_graph_test").disabled)
            app.run()
            self.assertEqual(send.call_count, 1)
            mocks["replace_rows"].assert_not_called()

    def test_viewer_has_no_internal_send_action(self):
        app, _ = self.open_page(self.OK_PREVIEW, role="viewer")
        self.assertFalse(any(b.key == "send_internal_graph_test" for b in app.button))

    def test_clicking_preview_shows_sender_to_cc_subject_body_and_pdf_download(self):
        app, mocks = self.open_page(self.OK_PREVIEW)
        [b for b in app.button if b.key == "preview_confirm_po-3"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        markdowns = "\n".join(m.value for m in app.markdown)
        self.assertIn("orders.caddycheck@video-inform.com", markdowns)
        self.assertIn("tom@example.com", markdowns)
        self.assertIn("yoram@example.com", markdowns)
        self.assertIn("**Proposed invoice number:** 8719", markdowns)
        self.assertIn("**Number of cameras:** 12", markdowns)
        self.assertTrue(any("not reserved, saved, or consumed" in c.value for c in app.caption))
        self.assertTrue(any("3335 bytes, 1 page(s)" in c.value for c in app.caption))
        self.assertEqual(len(app.json), 2)
        self.assertTrue(any("Dear Tom and Karen" in c.value for c in app.code))
        self.assertTrue(any(b.key == "preview_pdf_po-3" for b in app.download_button))
        mocks["replace_rows"].assert_not_called()
        mocks["send_email"].assert_not_called()

    def test_failed_validation_shows_errors_and_no_sendable_content(self):
        app, mocks = self.open_page(self.FAILED_PREVIEW)
        [b for b in app.button if b.key == "preview_confirm_po-3"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertTrue(any("camera count nor an explicit amount" in e.value for e in app.error))
        self.assertEqual([b for b in app.download_button if b.key == "preview_pdf_po-3"], [])
        mocks["replace_rows"].assert_not_called()
        mocks["send_email"].assert_not_called()

    def test_live_approval_is_disabled_without_explicit_gate(self):
        app, mocks = self.open_page(self.OK_PREVIEW)
        self.assertTrue(app.button(key="appr_po-3").disabled)
        self.assertFalse(app.exception, [e.value for e in app.exception])
        mocks["replace_rows"].assert_not_called()
        mocks["send_email"].assert_not_called()

    def test_only_selected_canary_has_review_and_viewer_has_no_action(self):
        from services import order_confirmation_workflow as workflow

        selected = dict(self.PREVIEW_ORDER, project_name="AD DELHAIZE Roeselare (Mandel)",
                        order_reference="2026-01-005", confirmation_status="none")
        with patch.object(self, "PREVIEW_ORDER", selected), \
             patch.object(workflow, "CANARY_ORDER_ID", "po-3"), \
             patch.object(workflow, "live_approval_enabled", return_value=False), \
             patch.object(workflow, "approve_order") as approve:
            app, _ = self.open_page(self.OK_PREVIEW)
            self.assertTrue(any(b.key == "review_approval_po-3" for b in app.button))
            self.assertTrue(app.button(key="appr_po-3").disabled)
            approve.assert_not_called()
            viewer, _ = self.open_page(self.OK_PREVIEW, role="viewer")
            self.assertFalse(any(b.key == "review_approval_po-3" for b in viewer.button))
            approve.assert_not_called()

    def test_enabled_admin_approve_calls_only_new_orchestrator(self):
        from services import order_confirmation_workflow as workflow

        with patch.object(workflow, "live_approval_enabled", return_value=True), \
             patch.object(workflow, "build_approval_review", return_value={
                 **self.OK_PREVIEW, "customer_display": "AD", "po_reference": "2026-01-005",
                 "fingerprint": "reviewed",
             }), \
             patch.object(workflow, "approve_order",
                          return_value={"status": "sent", "message": "Order approved · Invoice 8719 created · Confirmation email sent"}) as approve, \
             patch.object(oas, "create_invoice_from_purchase_order") as old_create:
            app, boundary_mocks = self.open_page(self.OK_PREVIEW)
            self.assertTrue(app.button(key="appr_po-3").disabled)
            app.button(key="review_approval_po-3").click().run()
            self.assertFalse(app.button(key="appr_po-3").disabled)
            approve.assert_not_called()
            markdowns = "\n".join(m.value for m in app.markdown)
            for label in ("Project/Site", "Network", "PO reference", "Camera count",
                          "Final Y1 amount", "Proposed real invoice number", "Sender", "To", "CC", "Email subject"):
                self.assertIn(label, markdowns)
            app.button(key="appr_po-3").click().run()
            self.assertFalse(app.exception, [e.value for e in app.exception])
            approve.assert_called_once()
            self.assertEqual(approve.call_args.args[0], "po-3")
            self.assertEqual(approve.call_args.kwargs, {"review_fingerprint": "reviewed"})
            old_create.assert_not_called()
            boundary_mocks["legacy_auto_invoice"].assert_not_called()
            boundary_mocks["update_order"].assert_not_called()

    def test_safe_failed_confirmation_exposes_resend_without_legacy_invoice_creation(self):
        from services import order_confirmation_workflow as workflow

        failed_order = dict(self.PREVIEW_ORDER, status="approved", confirmation_status="failed",
                            confirmation_retry_safe=True, confirmation_invoice_number="8719",
                            confirmation_error="Graph rejected")
        with patch.object(self, "PREVIEW_ORDER", failed_order), \
             patch.object(workflow, "live_approval_enabled", return_value=True), \
             patch.object(workflow, "approve_order",
                          return_value={"status": "sent", "message": "Confirmation sent"}) as approve:
            app, _ = self.open_page(self.OK_PREVIEW)
            app.button(key="resend_confirmation_po-3").click().run()
            self.assertFalse(app.exception)
            approve.assert_called_once()
            self.assertEqual(approve.call_args.kwargs, {"resend": True})

    def test_preview_read_failure_is_explicit_and_clears_stale_preview(self):
        app, mocks = self.open_page(RuntimeError("source lookup failed"))
        app.session_state["_confirm_preview_po-3"] = self.OK_PREVIEW
        [b for b in app.button if b.key == "preview_confirm_po-3"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertTrue(any("source lookup failed" in e.value for e in app.error))
        self.assertEqual([b for b in app.download_button if b.key == "preview_pdf_po-3"], [])
        mocks["replace_rows"].assert_not_called()
        mocks["send_email"].assert_not_called()


if __name__ == "__main__":
    unittest.main()
