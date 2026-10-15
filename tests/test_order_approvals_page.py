"""Streamlit AppTest of the Order Approvals page: failed-delivery warning and admin Retry button.

Runs the real app script with all network-facing calls mocked. Skipped when Streamlit is not installed.
"""
import importlib.util
import os
import unittest
from unittest.mock import patch

from services import order_approval_service as oas

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


if __name__ == "__main__":
    unittest.main()
