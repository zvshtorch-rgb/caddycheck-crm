"""Tests for the intake claim/lease and approval-email state helpers in order_approval_service."""
import datetime
import unittest
from unittest.mock import patch

from services import order_approval_service as approvals
from tests._intake_fakes import FakeSupabase


class ServiceTestCase(unittest.TestCase):
    def setUp(self):
        self.db = FakeSupabase()
        patcher = patch.object(approvals, "_client", self.db)
        patcher.start()
        self.addCleanup(patcher.stop)

    def new_row(self, message_id="<m1@x>", graph_id="G1", run_id="run-a"):
        return approvals.insert_incoming_email_claimed(
            {"message_id": message_id, "graph_message_id": graph_id}, run_id
        )


class TestSchemaMarker(ServiceTestCase):
    def test_missing_function_means_not_ready(self):
        self.db.schema_version = None
        ready, detail = approvals.schema_ready()
        self.assertFalse(ready)
        self.assertIn("not available", detail)

    def test_old_version_is_not_ready(self):
        self.db.schema_version = 1
        self.assertFalse(approvals.schema_ready()[0])

    def test_current_version_is_ready(self):
        self.assertTrue(approvals.schema_ready()[0])


class TestIncomingRowUniquenessAndLeases(ServiceTestCase):
    def test_duplicate_message_or_graph_id_is_rejected_not_raised(self):
        self.assertIsNotNone(self.new_row())
        self.assertIsNone(self.new_row(message_id="<m1@x>", graph_id="G2"))
        self.assertIsNone(self.new_row(message_id="<m2@x>", graph_id="G1"))
        self.assertEqual(len(self.db.rows("incoming_order_emails")), 1)

    def test_find_matches_either_identifier(self):
        row = self.new_row()
        self.assertEqual(approvals.find_incoming_email("<m1@x>")["id"], row["id"])
        self.assertEqual(approvals.find_incoming_email("<other>", "G1")["id"], row["id"])
        self.assertIsNone(approvals.find_incoming_email("<other>", "G9"))

    def test_only_one_worker_can_claim_a_released_row(self):
        row = self.new_row()
        self.assertIsNone(approvals.claim_incoming_email(row["id"], "run-b"))  # leased by run-a
        approvals.fail_incoming_email(row["id"], "run-a", "boom", 1)
        self.assertIsNotNone(approvals.claim_incoming_email(row["id"], "run-b"))
        self.assertIsNone(approvals.claim_incoming_email(row["id"], "run-c"))

    def test_stale_lease_can_be_taken_over(self):
        row = self.new_row()
        later = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(minutes=approvals.INTAKE_LEASE_MINUTES + 1)
        with patch.object(approvals, "_utcnow", return_value=later):
            claimed = approvals.claim_incoming_email(row["id"], "run-b")
        self.assertEqual(claimed["claimed_by"], "run-b")

    def test_a_worker_that_lost_its_lease_cannot_finish_the_row(self):
        row = self.new_row()
        approvals.fail_incoming_email(row["id"], "run-a", "boom", 1)
        approvals.claim_incoming_email(row["id"], "run-b")
        approvals.finish_incoming_email(row["id"], "run-a")
        self.assertEqual(self.db.rows("incoming_order_emails")[0]["processing_status"], "error")
        approvals.finish_incoming_email(row["id"], "run-b")
        self.assertEqual(self.db.rows("incoming_order_emails")[0]["processing_status"], "processed")

    def test_failures_are_recorded_without_deleting_the_row(self):
        row = self.new_row()
        approvals.fail_incoming_email(row["id"], "run-a", "x" * 900, 2)
        stored = self.db.rows("incoming_order_emails")[0]
        self.assertEqual((stored["processing_status"], stored["attempt_count"], len(stored["error_message"])), ("error", 2, 500))

    def test_processed_rows_are_never_claimable_or_listed(self):
        row = self.new_row()
        approvals.finish_incoming_email(row["id"], "run-a")
        self.assertIsNone(approvals.claim_incoming_email(row["id"], "run-b"))
        floor = datetime.datetime(2020, 1, 1, tzinfo=datetime.timezone.utc)
        self.assertEqual(approvals.list_unfinished_incoming(floor), [])


class TestPurchaseOrderPerEmail(ServiceTestCase):
    def test_at_most_one_order_per_incoming_email(self):
        row = self.new_row()
        first, created_first = approvals.get_or_create_purchase_order_for_incoming(row["id"], {"project_name": "A"})
        second, created_second = approvals.get_or_create_purchase_order_for_incoming(row["id"], {"project_name": "B"})
        self.assertEqual((created_first, created_second, first["id"] == second["id"]), (True, False, True))
        with self.assertRaises(Exception):
            approvals.create_purchase_order(incoming_email_id=row["id"], project_name="C")
        self.assertEqual(len(self.db.rows("purchase_orders")), 1)


class TestApprovalEmailState(ServiceTestCase):
    def setUp(self):
        super().setUp()
        row = self.new_row()
        self.order, _ = approvals.get_or_create_purchase_order_for_incoming(row["id"], {"project_name": "A"})

    def test_state_machine_never_creates_a_second_approval_row(self):
        first = approvals.prepare_approval_email(self.order["id"])
        self.assertEqual((first["state"], first["attempt"]), ("send", 1))
        self.assertEqual(approvals.prepare_approval_email(self.order["id"])["state"], "in_progress")

        approvals.mark_approval_email_failed(first["approval"]["id"], "503")
        second = approvals.prepare_approval_email(self.order["id"])
        self.assertEqual((second["state"], second["attempt"]), ("send", 2))
        self.assertEqual(second["approval"]["id"], first["approval"]["id"])
        self.assertNotEqual(second["token"], first["token"])

        approvals.mark_approval_email_sent(second["approval"]["id"])
        self.assertEqual(approvals.prepare_approval_email(self.order["id"])["state"], "already_sent")
        self.assertEqual(len(self.db.rows("order_approvals")), 1)

    def test_only_the_stored_hash_is_persisted(self):
        prepared = approvals.prepare_approval_email(self.order["id"])
        stored = self.db.rows("order_approvals")[0]
        self.assertEqual(stored["token_hash"], approvals.hash_token(prepared["token"]))
        self.assertNotIn(prepared["token"], str(stored))
        self.assertIsNotNone(approvals.get_approval_by_token(prepared["token"]))

    def test_resend_invalidates_the_previous_link(self):
        first = approvals.prepare_approval_email(self.order["id"])
        approvals.mark_approval_email_failed(first["approval"]["id"], "timeout")
        approvals.prepare_approval_email(self.order["id"])
        self.assertIsNone(approvals.get_approval_by_token(first["token"]))

    def test_decided_or_superseded_orders_need_no_email(self):
        approvals.update_purchase_order(self.order["id"], status="superseded", is_current=False)
        self.assertEqual(approvals.prepare_approval_email(self.order["id"])["state"], "not_needed")

    def test_stale_sending_state_is_retried(self):
        approvals.prepare_approval_email(self.order["id"])
        later = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(minutes=approvals.INTAKE_LEASE_MINUTES + 1)
        with patch.object(approvals, "_utcnow", return_value=later):
            retried = approvals.prepare_approval_email(self.order["id"])
        self.assertEqual((retried["state"], retried["attempt"]), ("send", 2))


if __name__ == "__main__":
    unittest.main()
