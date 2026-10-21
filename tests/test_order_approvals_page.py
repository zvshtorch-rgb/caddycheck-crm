"""Streamlit AppTest of the Order Approvals page: failed-delivery warning and admin Retry button.

Runs the real app script with all network-facing calls mocked. Skipped when Streamlit is not installed.
"""
import ast
import importlib.util
import os
import re
import unittest
from pathlib import Path
from unittest.mock import patch

from services import order_approval_service as oas
from services.ask_data.helpers import _safe_str

HAS_STREAMLIT = importlib.util.find_spec("streamlit") is not None
APP_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "streamlit_app.py")


def _load_project_name_looks_unreliable():
    """Pure helper extracted via ast (streamlit_app.py is an entry-point script with top-level UI code,
    so a plain `import streamlit_app` would boot the whole app -- see test_license_update_confirmation.py
    for the same pattern)."""
    tree = ast.parse(Path(APP_PATH).read_text(encoding="utf-8"))
    nodes = [node for node in tree.body if isinstance(node, (ast.FunctionDef, ast.Assign))
             and (getattr(node, "name", None) == "_project_name_looks_unreliable"
                  or any(isinstance(t, ast.Name) and t.id == "_UNRELIABLE_PROJECT_NAME_WORDS" for t in getattr(node, "targets", [])))]
    module = ast.Module(body=nodes, type_ignores=[])
    code = compile(module, "<extracted-project-name-check>", "exec")
    namespace = {"re": re, "_safe_str": _safe_str}
    exec(code, namespace)
    return namespace["_project_name_looks_unreliable"]


_project_name_looks_unreliable = _load_project_name_looks_unreliable()

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


if __name__ == "__main__":
    unittest.main()
