"""Admin workflow unit tests: every external boundary is mocked; no production records or emails."""
import base64
import copy
import threading
import unittest
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import patch

import requests

from services import order_confirmation_workflow as flow
from services import order_confirmation_service as preview


ORDER_ID = "10000000-0000-0000-0000-000000000001"
ORDER = {
    "id": ORDER_ID, "project_name": "AD DELHAIZE Roeselare (Mandel)", "amount": 9336,
    "currency": "EUR", "order_reference": "2026-01-005", "ordered_cameras": 12,
    "customer_email": "contact@example.com", "status": "pending_approval",
    "confirmation_status": "none", "confirmation_invoice_number": None,
}
CONFIG = {
    "live_enabled": True, "sender": "orders.caddycheck@video-inform.com",
    "to": ["tom@example.com", "karen@example.com"],
    "cc": ["yoram@example.com", "zvika@example.com"],
}


class MemoryRPC:
    """Small stateful RPC boundary double; real SQL/concurrency is tested separately with PostgreSQL."""
    def __init__(self):
        self.order = copy.deepcopy(ORDER)
        self.rows = []
        self.registry = []
        self.audit = []
        self.events = []
        self.lock = threading.Lock()
        self.fail = None

    def rpc(self, name, **p):
        with self.lock:
            self.events.append(name)
            if self.fail == name:
                raise RuntimeError(f"Injected failure: {name}")
            if name == "claim_order_confirmation":
                if self.order.get("confirmation_status") == "sent":
                    return {"state": "sent", "order": copy.deepcopy(self.order)}
                if self.order.get("confirmation_lock_token"):
                    return {"state": "busy", "order": copy.deepcopy(self.order)}
                self.order["confirmation_lock_token"] = p["p_token"]
                self.order["confirmation_stage"] = "preparing"
                return {"state": "claimed", "order": copy.deepcopy(self.order)}
            if self.order.get("confirmation_lock_token") != p["p_token"]:
                raise RuntimeError("claim lost")
            if name == "reserve_order_confirmation_invoice":
                if not self.order["confirmation_invoice_number"]:
                    self.order["confirmation_invoice_number"] = "8719"
                    self.registry.append("8719")
                return self.order["confirmation_invoice_number"]
            if name == "save_order_confirmation_invoice":
                if not self.rows:
                    self.rows = [dict(p["p_rows"][0], id=1, source_purchase_order_id=ORDER_ID)]
                    self.order["confirmation_invoice_amount"] = self.rows[0]["payment_amount"]
                return copy.deepcopy(self.rows)
            if name == "start_order_confirmation_send":
                self.order.update({
                    "status": "approved", "confirmation_status": "pending",
                    "confirmation_stage": "sending", "confirmation_retry_safe": False,
                    "confirmation_pdf_base64": p["p_pdf_base64"], "confirmation_details": p["p_details"],
                })
                return copy.deepcopy(self.order)
            if name == "finish_order_confirmation":
                self.order["confirmation_status"] = "sent" if p["p_outcome"] == "sent" else "failed"
                self.order["confirmation_error"] = p["p_error"]
                self.order["confirmation_retry_safe"] = p["p_retry_safe"]
                if p["p_outcome"] != "unknown":
                    self.order["confirmation_lock_token"] = None
                self.audit.append(copy.deepcopy(p))
                return copy.deepcopy(self.order)
            raise AssertionError(name)


class TestLiveOrderConfirmation(unittest.TestCase):
    def setUp(self):
        self.db = MemoryRPC()
        self.pdf_events = []
        self.sends = []
        renderer = preview.generate_year1_invoice_pdf_bytes

        def render(number, rows):
            self.assertEqual(self.db.rows, rows, "PDF must use the committed RPC result, not a draft payload")
            self.pdf_events.append("render")
            return renderer(number, rows)

        def send(**kwargs):
            self.assertEqual(self.db.order["status"], "approved")
            self.assertEqual(self.db.order["confirmation_status"], "pending")
            self.sends.append(copy.deepcopy(kwargs))
            return {"accepted": True, "http_status": 202}

        patches = (
            patch.object(flow, "get_order_confirmation_config", return_value=CONFIG),
            patch.object(flow, "get_purchase_order", side_effect=lambda _: copy.deepcopy(self.db.order)),
            patch.object(flow, "_rpc", side_effect=self.db.rpc),
            patch.object(flow.preview, "resolve_project_country", return_value="Belgium"),
            patch.object(flow.preview, "generate_year1_invoice_pdf_bytes", side_effect=render),
            patch.object(flow, "send_graph_email", side_effect=send),
            patch("services.order_approval_service.auto_create_invoice_for_approved_order",
                  side_effect=AssertionError("Legacy invoice path forbidden")),
            patch("services.order_approval_service.create_invoice_from_purchase_order",
                  side_effect=AssertionError("Legacy invoice path forbidden")),
            patch("services.supabase_service.replace_invoice_rows",
                  side_effect=AssertionError("Destructive replacement forbidden")),
        )
        self.mocks = [p.start() for p in patches]
        for p in patches:
            self.addCleanup(p.stop)

    def approve(self, **kwargs):
        return flow.approve_order(ORDER_ID, "Admin User", **kwargs)

    def test_success_one_committed_invoice_pdf_email_and_audit(self):
        result = self.approve()
        self.assertEqual(result["status"], "sent")
        self.assertEqual(len(self.db.rows), 1)
        self.assertEqual(self.db.registry, ["8719"])
        self.assertEqual(self.pdf_events, ["render"])
        self.assertEqual(len(self.sends), 1)
        message = self.sends[0]
        self.assertEqual(message["sender_mailbox"], CONFIG["sender"])
        self.assertEqual(message["recipients"], CONFIG["to"])
        self.assertEqual(message["cc"], CONFIG["cc"])
        self.assertIn("Customer/Network: AD", message["body"])
        self.assertIn("Number of cameras: 12", message["body"])
        self.assertIn("9,336.00", message["body"])
        self.assertIn("Invoice number: 8719", message["body"])
        self.assertNotIn("Proposed invoice", message["body"])
        self.assertEqual(message["attachment_bytes"], base64.b64decode(self.db.order["confirmation_pdf_base64"]))
        self.assertEqual(self.db.audit[0]["p_outcome"], "sent")
        self.assertEqual(self.db.audit[0]["p_actor"], "Admin User")
        self.assertEqual(self.db.events, [
            "claim_order_confirmation", "reserve_order_confirmation_invoice",
            "save_order_confirmation_invoice", "start_order_confirmation_send", "finish_order_confirmation",
        ])

    def test_repeated_click_or_refresh_never_creates_or_sends_again(self):
        self.approve()
        for _ in range(3):
            self.assertTrue(self.approve()["already_sent"])
        self.assertEqual(len(self.db.rows), 1)
        self.assertEqual(len(self.db.registry), 1)
        self.assertEqual(len(self.sends), 1)
        self.assertEqual(len(self.db.audit), 1)

    def test_concurrent_click_is_blocked_by_durable_claim(self):
        entered = threading.Event()
        release = threading.Event()
        original = self.mocks[4].side_effect

        def pause(number, rows):
            entered.set()
            self.assertTrue(release.wait(5))
            return original(number, rows)

        self.mocks[4].side_effect = pause
        with ThreadPoolExecutor(max_workers=2) as executor:
            first = executor.submit(self.approve)
            self.assertTrue(entered.wait(5))
            second = self.approve()
            release.set()
            self.assertEqual(first.result()["status"], "sent")
        self.assertEqual(second["status"], "busy")
        self.assertEqual(len(self.sends), 1)
        self.assertEqual(self.db.registry, ["8719"])

    def test_validation_failures_have_no_database_writes_or_email(self):
        for fields in ({"project_name": ""}, {"order_reference": ""}, {"amount": float("nan")},
                       {"currency": "USD"}, {"ordered_cameras": 1.5}):
            with self.subTest(fields=fields):
                self.db.order = dict(ORDER, **fields)
                with self.assertRaises(ValueError):
                    self.approve()
        self.assertEqual(self.db.events, [])
        self.assertEqual(self.sends, [])

    def test_invoice_failure_keeps_reservation_and_does_not_approve_or_send(self):
        self.db.fail = "save_order_confirmation_invoice"
        self.assertEqual(self.approve()["status"], "failed")
        self.assertEqual(self.db.order["status"], "pending_approval")
        self.assertEqual(self.sends, [])
        self.assertEqual(self.db.order["confirmation_invoice_number"], "8719")
        self.db.fail = None
        self.assertEqual(self.approve()["status"], "sent")
        self.assertEqual(self.db.registry, ["8719"])

    def test_pdf_failure_keeps_saved_invoice_then_reuses_it_on_retry(self):
        original = self.mocks[4].side_effect
        self.mocks[4].side_effect = RuntimeError("PDF renderer failed")
        self.assertEqual(self.approve()["stage"], "pdf")
        self.assertEqual(self.db.order["status"], "pending_approval")
        original_rows = copy.deepcopy(self.db.rows)
        self.assertEqual(self.sends, [])
        self.mocks[4].side_effect = original
        self.assertEqual(self.approve()["status"], "sent")
        self.assertEqual(self.db.rows, original_rows)
        self.assertEqual(self.db.registry, ["8719"])

    def test_existing_invoice_is_used_even_if_new_payload_has_different_amount(self):
        self.db.order["confirmation_invoice_number"] = "8719"
        self.db.registry = ["8719"]
        self.db.rows = [dict(preview.build_year1_invoice_rows(ORDER, 8719)[0], id=1,
                             source_purchase_order_id=ORDER_ID, payment_amount=9000)]
        self.db.order["confirmation_invoice_amount"] = 9000
        self.approve()
        self.assertIn("9,000.00", self.sends[0]["body"])
        self.assertEqual(self.db.registry, ["8719"])

    def test_graph_rejection_and_resend_reuse_exact_pdf_and_envelope(self):
        response = requests.Response()
        response.status_code = 403
        response._content = b'{"error":{"code":"ErrorAccessDenied","message":"Access denied"}}'
        original = self.mocks[5].side_effect
        self.mocks[5].side_effect = requests.HTTPError(response=response)
        result = self.approve()
        self.assertEqual(result["status"], "failed")
        self.assertEqual(self.db.order["status"], "approved")
        self.assertEqual(self.db.order["confirmation_status"], "failed")
        self.assertIn("ErrorAccessDenied", self.db.order["confirmation_error"])
        stored_pdf = self.db.order["confirmation_pdf_base64"]
        self.mocks[5].side_effect = original
        with patch.object(flow, "get_order_confirmation_config", return_value=dict(CONFIG, to=["different@example.com"])):
            self.assertEqual(self.approve(resend=True)["status"], "sent")
        self.assertEqual(self.pdf_events, ["render"])
        self.assertEqual(self.sends[0]["recipients"], CONFIG["to"])
        self.assertEqual(self.sends[0]["attachment_bytes"], base64.b64decode(stored_pdf))
        self.assertEqual(self.db.registry, ["8719"])
        self.assertEqual(len(self.db.rows), 1)
        self.assertEqual([r["p_outcome"] for r in self.db.audit], ["failed", "sent"])

    def test_timeout_blocks_resend_and_does_not_claim_success(self):
        self.mocks[5].side_effect = requests.Timeout("no response")
        result = self.approve()
        self.assertEqual(result["status"], "unknown")
        self.assertFalse(result["retry_safe"])
        self.assertEqual(self.approve(resend=True)["status"], "busy")
        self.assertEqual(self.mocks[5].call_count, 1)
        self.assertEqual(self.db.audit[0]["p_outcome"], "unknown")

    def test_finalization_failure_after_graph_acceptance_never_resends(self):
        self.db.fail = "finish_order_confirmation"
        with self.assertRaisesRegex(RuntimeError, "Graph accepted.*Do not resend"):
            self.approve()
        self.assertEqual(self.db.order["confirmation_status"], "pending")
        self.assertEqual(self.approve()["status"], "busy")
        self.assertEqual(len(self.sends), 1)
        self.assertEqual(self.db.audit, [])

    def test_disabled_gate_blocks_every_side_effect(self):
        with patch.object(flow, "get_order_confirmation_config", return_value={}):
            with self.assertRaisesRegex(ValueError, "disabled"):
                self.approve()
        self.assertEqual(self.db.events, [])
        self.assertEqual(self.sends, [])

    def test_foreign_invoice_response_is_rejected_before_pdf_or_approval(self):
        self.db.rows = [{"invoice_number": "8719", "source_purchase_order_id": "another-order"}]
        result = self.approve()
        self.assertEqual(result["status"], "failed")
        self.assertEqual(self.pdf_events, [])
        self.assertEqual(self.sends, [])
        self.assertEqual(self.db.order["status"], "pending_approval")

    def test_canary_gate_is_exact_and_stops_after_any_attempt(self):
        with patch.object(flow, "get_order_confirmation_config", return_value=dict(CONFIG, live_enabled=False)):
            selected = dict(ORDER, id=flow.CANARY_ORDER_ID)
            self.assertFalse(flow.live_approval_enabled())
            self.assertTrue(flow.order_approval_enabled(selected))
            for changed in (
                {"id": ORDER_ID}, {"project_name": "Another project"},
                {"order_reference": "another"}, {"status": "approved"},
                {"confirmation_status": "failed"}, {"confirmation_stage": "invoice"},
                {"confirmation_invoice_number": "8719"}, {"is_current": False},
                {"confirmation_lock_token": "claimed"},
            ):
                self.assertFalse(flow.order_approval_enabled(dict(selected, **changed)), changed)
            with self.assertRaisesRegex(ValueError, "disabled"):
                self.approve()
        self.assertEqual(self.db.events, [])
        self.assertEqual(self.sends, [])

    def test_canary_requires_review_and_rechecks_before_any_write(self):
        with patch.object(flow, "CANARY_ORDER_ID", ORDER_ID), \
             patch.object(flow, "get_order_confirmation_config", return_value=dict(CONFIG, live_enabled=False)), \
             patch.object(preview, "_next_invoice_number_preview", return_value=8719):
            with self.assertRaisesRegex(ValueError, "final confirmation panel"):
                self.approve()
            review = flow.build_approval_review(ORDER_ID)
            self.assertEqual(review["customer_display"], "AD")
            self.assertEqual(review["cameras"], 12)
            self.assertEqual(review["amount"], 9336)
            self.assertEqual(self.db.events, [])
            self.db.order["amount"] = 9000
            with self.assertRaisesRegex(ValueError, "changed"):
                self.approve(review_fingerprint=review["fingerprint"])
            self.assertEqual(self.db.events, [])
            self.db.order["amount"] = 9336
            self.assertEqual(self.approve(review_fingerprint=review["fingerprint"])["status"], "sent")
            with self.assertRaisesRegex(ValueError, "disabled"):
                self.approve(review_fingerprint=review["fingerprint"])
        self.assertEqual(len(self.sends), 1)
        self.assertEqual(len(self.db.audit), 1)

    def test_failed_canary_cannot_retry_or_resend(self):
        with patch.object(flow, "CANARY_ORDER_ID", ORDER_ID), \
             patch.object(flow, "get_order_confirmation_config", return_value=dict(CONFIG, live_enabled=False)), \
             patch.object(preview, "_next_invoice_number_preview", return_value=8719):
            review = flow.build_approval_review(ORDER_ID)
            self.db.fail = "save_order_confirmation_invoice"
            self.assertEqual(self.approve(review_fingerprint=review["fingerprint"])["status"], "failed")
            for resend in (False, True):
                with self.assertRaisesRegex(ValueError, "disabled"):
                    self.approve(resend=resend, review_fingerprint=review["fingerprint"])
        self.assertEqual(self.sends, [])
        self.assertEqual(len(self.db.registry), 1)
        self.assertEqual(len(self.db.audit), 1)
