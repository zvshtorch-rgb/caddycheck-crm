"""Tests for the order intake poller: auth failures, read state, dedup, retries, concurrency."""
import datetime
import logging
import os
import re
import tempfile
import threading
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import order_intake_poll
from order_intake_poll import main
from services import email_service
from services import order_approval_service as approvals
from services.email_intake_service import IntakeError
from tests._intake_fakes import NOW, FakeMailer, FakeProvider, FakeSupabase, make_message

OLD_FLOOR = "2020-01-01T00:00:00Z"
SECRET = "SUPERSECRET-CLIENT-SECRET-VALUE"


class IntakeTestCase(unittest.TestCase):
    def setUp(self):
        self.db = FakeSupabase()
        self.mailer = FakeMailer()
        self.provider = FakeProvider()
        summary_file = tempfile.NamedTemporaryFile(delete=False, suffix=".md")
        summary_file.close()
        self.summary_path = summary_file.name
        self.addCleanup(os.unlink, self.summary_path)

        env = {
            "ORDER_APPROVAL_CEO_EMAIL": "ceo@example.com",
            "APP_BASE_URL": "https://app.example.com",
            "ORDER_INTAKE_MIN_RECEIVED": OLD_FLOOR,
            "ORDER_INTAKE_GRAPH_CLIENT_SECRET": SECRET,
            "GITHUB_STEP_SUMMARY": self.summary_path,
        }
        for patcher in (
            patch.dict(os.environ, env),
            patch.object(approvals, "_client", self.db),
            patch("order_intake_poll.send_graph_email", self.mailer),
            patch("order_intake_poll._utcnow", return_value=NOW),
            patch("order_intake_poll.extract_pdf_text", return_value=""),
        ):
            patcher.start()
            self.addCleanup(patcher.stop)
        for name in ("ORDER_INTAKE_LOOKBACK_HOURS", "ORDER_INTAKE_MAX_NEW_PER_RUN", "GITHUB_ACTIONS"):
            os.environ.pop(name, None)

    # -- helpers ---------------------------------------------------------------
    def run_poller(self, *argv, provider=None):
        with patch("order_intake_poll.get_email_provider", return_value=provider or self.provider):
            return main(list(argv))

    @property
    def orders(self):
        return self.db.rows("purchase_orders")

    @property
    def incoming(self):
        return self.db.rows("incoming_order_emails")

    @property
    def approval_rows(self):
        return self.db.rows("order_approvals")

    def counters(self):
        with open(self.summary_path, encoding="utf-8") as handle:
            text = handle.read()
        last = text.split("## Order intake")[-1]
        return {k: int(v) for k, v in re.findall(r"\| (\w+) \| (\d+) \|", last)}


class TestAuthAndConfiguration(IntakeTestCase):
    def test_authentication_failure_fails_the_run(self):
        self.provider.preflight_error = IntakeError(
            "auth", "Token request failed: HTTP 401 invalid_client", status=401
        )
        self.assertEqual(self.run_poller(), 1)
        self.assertEqual(self.orders, [])
        with open(self.summary_path, encoding="utf-8") as handle:
            self.assertIn("FAILED", handle.read())

    def test_mailbox_access_denied_fails_the_run(self):
        self.provider.preflight_error = IntakeError("auth", "Mailbox folder check failed: HTTP 403", status=403)
        self.assertEqual(self.run_poller(), 1)

    def test_unconfigured_mailbox_fails_instead_of_skipping(self):
        class Unconfigured(FakeProvider):
            def is_configured(self):
                return False

            def missing_settings(self):
                return ["ORDER_INTAKE_GRAPH_CLIENT_SECRET"]

        self.assertEqual(self.run_poller(provider=Unconfigured()), 1)

    def test_imap_provider_is_rejected(self):
        imap = SimpleNamespace(name="imap", is_configured=lambda: True)
        self.assertEqual(self.run_poller(provider=imap), 1)

    def test_listing_failure_fails_the_run(self):
        with patch.object(self.provider, "list_message_headers",
                          side_effect=IntakeError("transient", "Message listing failed: HTTP 500", status=500)):
            self.assertEqual(self.run_poller(), 1)

    def test_check_mode_only_runs_preflight_and_schema(self):
        self.provider.messages = [make_message(1)]
        self.assertEqual(self.run_poller("--check"), 0)
        self.assertEqual(self.provider.list_calls, [])
        self.assertEqual(self.orders, [])

    def test_live_run_blocked_until_migration_applied(self):
        self.db.schema_version = None
        self.provider.messages = [make_message(1)]
        self.assertEqual(self.run_poller(), 1)
        self.assertEqual(self.provider.list_calls, [])
        self.assertEqual(self.incoming, [])

    def test_dry_run_tolerates_missing_migration_but_warns(self):
        self.db.schema_version = None
        self.provider.messages = [make_message(1)]
        self.assertEqual(self.run_poller("--dry-run"), 0)
        with open(self.summary_path, encoding="utf-8") as handle:
            self.assertIn("migration not applied", handle.read())

    def test_live_requires_go_live_floor(self):
        os.environ.pop("ORDER_INTAKE_MIN_RECEIVED")
        self.assertEqual(self.run_poller(), 1)
        self.assertEqual(self.provider.list_calls, [])

    def test_live_requires_ceo_recipient_and_app_url(self):
        os.environ.pop("ORDER_APPROVAL_CEO_EMAIL")
        self.assertEqual(self.run_poller(), 1)
        os.environ["ORDER_APPROVAL_CEO_EMAIL"] = "ceo@example.com"
        os.environ.pop("APP_BASE_URL")
        self.assertEqual(self.run_poller(), 1)


class TestReadStateAndWindow(IntakeTestCase):
    def test_read_messages_are_processed_and_read_state_untouched(self):
        self.provider.messages = [make_message(1, is_read=True), make_message(2, is_read=False)]
        self.assertEqual(self.run_poller(), 0)  # FakeProvider.mark_processed would raise if called
        self.assertEqual(len(self.orders), 2)
        self.assertEqual(len(self.mailer.sent), 2)

    def test_messages_without_pdf_are_skipped(self):
        self.provider.messages = [make_message(1, has_pdf=False), make_message(2)]
        self.assertEqual(self.run_poller(), 0)
        counters = self.counters()
        self.assertEqual((counters["created"], counters["scanned"]), (1, 2))

    def test_default_window_is_seven_days_rolling(self):
        os.environ["ORDER_INTAKE_MIN_RECEIVED"] = "2026-01-01T00:00:00Z"
        self.run_poller("--dry-run")
        self.assertEqual(self.provider.list_calls[0], NOW - datetime.timedelta(days=7))

    def test_go_live_floor_limits_the_window(self):
        os.environ["ORDER_INTAKE_MIN_RECEIVED"] = "2026-10-15T12:30:00Z"
        self.provider.messages = [make_message(1)]  # received 12:15, before the floor
        self.assertEqual(self.run_poller(), 0)
        self.assertEqual(self.orders, [])
        self.assertEqual(self.provider.list_calls[0].isoformat(), "2026-10-15T12:30:00+00:00")

    def test_explicit_backfill_only_imports_the_targeted_sender(self):
        os.environ["ORDER_INTAKE_MIN_RECEIVED"] = "2026-10-15T08:00:00Z"
        pat = make_message(1, sender="Pat Buyer <pat.buyer@customer-a.example>")
        sam = make_message(
            2, sender="Sam Other <sam.other@customer-b.example>",
            received=datetime.datetime(2026, 10, 14, 16, 40, tzinfo=datetime.timezone.utc),
        )
        self.provider.messages = [pat, sam]
        argv = ("--since", "2026-10-14T00:00:00Z", "--only-from", "pat.buyer@customer-a.example")
        self.assertEqual(self.run_poller(*argv), 0)
        self.assertEqual(len(self.orders), 1)
        self.assertEqual(self.provider.fetch_calls, 1)  # Other sender's PDF was never even downloaded
        self.assertEqual(self.run_poller(*argv), 0)  # repeat: still one order, one email
        self.assertEqual((len(self.orders), len(self.mailer.sent)), (1, 1))
        self.assertEqual(self.run_poller(), 0)  # normal rolling run ignores the pre-floor email
        self.assertEqual(len(self.orders), 1)

    def test_live_backfill_without_filter_is_rejected(self):
        self.assertEqual(self.run_poller("--since", "2026-10-01T00:00:00Z"), 1)
        self.assertEqual(self.provider.list_calls, [])

    def test_backfill_works_without_go_live_floor_and_scheduled_flag(self):
        os.environ.pop("ORDER_INTAKE_MIN_RECEIVED")
        self.provider.messages = [make_message(1), make_message(2, sender="Sam Other <sam.other@customer-b.example>")]
        argv = ("--since", "2026-10-15T08:00:00Z", "--only-from", "pat.buyer@customer-a.example")
        self.assertEqual(self.run_poller(*argv), 0)
        self.assertEqual((len(self.orders), len(self.mailer.sent)), (1, 1))
        self.assertEqual(self.run_poller("--dry-run"), 0)  # rolling runs still need the floor when live
        self.assertEqual(self.run_poller(), 1)

    def test_backfill_filters_must_be_specific(self):
        base = ("--since", "2026-10-15T08:00:00Z")
        for broad in (("--only-from", "com"), ("--only-from", "pat"), ("--only-from", "patbuyer"),
                      ("--subject-contains", "order")):
            self.assertEqual(self.run_poller(*base, *broad), 1, broad)
        self.assertEqual(self.provider.list_calls, [])
        self.assertEqual(self.run_poller(*base, "--subject-contains", "Quokka supermarket"), 0)

    def test_backfill_cannot_reach_further_back_than_thirty_days(self):
        argv = ("--only-from", "pat.buyer@customer-a.example")
        self.assertEqual(self.run_poller("--since", "2026-09-01T00:00:00Z", *argv), 1)
        self.assertEqual(self.provider.list_calls, [])

    def test_safety_limit_prevents_bulk_import(self):
        self.provider.messages = [make_message(n) for n in range(1, 12)]
        self.assertEqual(self.run_poller(), 1)
        self.assertEqual((self.orders, self.mailer.sent), ([], []))
        self.assertEqual(self.run_poller("--max-new", "20"), 0)
        self.assertEqual(len(self.orders), 11)

    def planned_refs(self):
        with open(self.summary_path, encoding="utf-8") as handle:
            last = handle.read().split("## Order intake")[-1]
        return re.findall(r"\| `(\w+)` \| (?:new|resume) \|", last)

    def test_dry_run_reports_every_eligible_message_and_a_filter_isolates_one_sender(self):
        pat = make_message(1, sender="Pat Buyer <pat.buyer@customer-a.example>")
        sam = make_message(2, sender="Sam Other <sam.other@customer-b.example>", received=NOW.replace(hour=11))
        self.provider.messages = [pat, sam, make_message(3, has_pdf=False)]
        self.assertEqual(self.run_poller("--dry-run"), 0)
        everything = self.planned_refs()
        self.assertEqual(len(everything), 2)  # not assumed to be one
        self.assertEqual(self.counters()["skipped"], 1)

        self.assertEqual(self.run_poller("--dry-run", "--only-from", "pat.buyer@customer-a.example"), 0)
        only_tom = self.planned_refs()
        self.assertEqual(len(only_tom), 1)
        self.assertIn(only_tom[0], everything)
        self.assertEqual((self.orders, self.incoming), ([], []))

    def test_dry_run_writes_nothing(self):
        self.provider.messages = [make_message(1)]
        self.assertEqual(self.run_poller("--dry-run"), 0)
        self.assertEqual((self.orders, self.incoming, self.mailer.sent), ([], [], []))
        self.assertEqual(self.db.storage.files, {})
        self.assertEqual(self.counters()["eligible"], 1)


class TestDeduplication(IntakeTestCase):
    def test_second_run_creates_nothing(self):
        self.provider.messages = [make_message(1)]
        self.assertEqual(self.run_poller(), 0)
        self.assertEqual(self.run_poller(), 0)
        self.assertEqual((len(self.orders), len(self.approval_rows), len(self.mailer.sent)), (1, 1, 1))
        self.assertEqual(self.provider.fetch_calls, 1)  # processed mail is not downloaded again
        self.assertEqual(self.counters()["created"], 0)

    def test_same_graph_id_with_changed_internet_id_is_skipped(self):
        self.provider.messages = [make_message(1, message_id="<a@x>", graph_id="G1")]
        self.run_poller()
        self.provider.messages = [make_message(1, message_id="<changed@x>", graph_id="G1")]
        self.run_poller()
        self.assertEqual(len(self.orders), 1)

    def test_same_internet_id_with_new_graph_id_is_skipped(self):
        self.provider.messages = [make_message(1, message_id="<a@x>", graph_id="G1")]
        self.run_poller()
        self.provider.messages = [make_message(1, message_id="<a@x>", graph_id="G-moved")]
        self.run_poller()
        self.assertEqual(len(self.orders), 1)

    def test_identical_pdf_is_flagged_but_not_rejected(self):
        pdf = b"%PDF-1.4 identical content"
        self.provider.messages = [
            make_message(1, pdf=pdf),
            make_message(2, pdf=pdf, received=NOW.replace(hour=12, minute=20)),
        ]
        self.assertEqual(self.run_poller(), 0)
        self.assertEqual(len(self.orders), 2)
        self.assertEqual(len(self.mailer.sent), 2)
        first, second = sorted(self.incoming, key=lambda r: r["created_at"])
        self.assertIsNone(first["duplicate_of_email_id"])
        self.assertEqual(second["duplicate_of_email_id"], first["id"])
        self.assertEqual(self.counters()["possible_duplicates"], 1)
        titles = [n["title"] for n in self.db.rows("crm_notifications")]
        self.assertIn("Possible duplicate purchase-order PDF", titles)

    def test_pdf_storage_path_is_deterministic(self):
        self.provider.messages = [make_message(1)]
        self.run_poller()
        paths = {k[1] for k in self.db.storage.files}
        self.assertEqual(len(paths), 1)
        self.assertIn(self.incoming[0]["pdf_sha256"][:16], paths.pop())


class TestRetriesAreIdempotent(IntakeTestCase):
    def test_failed_email_is_retried_without_duplicates(self):
        self.provider.messages = [make_message(1)]
        self.mailer.fail_next = 1
        self.assertEqual(self.run_poller(), 2)
        self.assertEqual((len(self.orders), len(self.approval_rows), len(self.mailer.sent)), (1, 1, 0))
        self.assertEqual(self.incoming[0]["processing_status"], "error")
        self.assertEqual(self.incoming[0]["attempt_count"], 1)
        self.assertEqual(self.approval_rows[0]["email_status"], "failed")
        first_hash = self.approval_rows[0]["token_hash"]

        self.assertEqual(self.run_poller(), 0)
        self.assertEqual((len(self.incoming), len(self.orders), len(self.approval_rows)), (1, 1, 1))
        self.assertEqual(len(self.mailer.sent), 1)
        self.assertTrue(self.mailer.sent[0]["subject"].startswith("[Resent]"))
        self.assertEqual(self.incoming[0]["processing_status"], "processed")
        self.assertEqual(self.approval_rows[0]["email_status"], "sent")
        self.assertEqual(self.approval_rows[0]["email_attempts"], 2)
        self.assertNotEqual(self.approval_rows[0]["token_hash"], first_hash)  # rotated in place
        self.assertEqual(self.provider.fetch_calls, 1)  # resumed from stored data, not re-downloaded

        self.assertEqual(self.run_poller(), 0)
        self.assertEqual(len(self.mailer.sent), 1)

    def test_crash_after_order_creation_resumes_without_second_order(self):
        self.provider.messages = [make_message(1)]
        real = order_intake_poll._supersede_older_revisions
        calls = {"n": 0}

        def flaky(order, ref):
            calls["n"] += 1
            if calls["n"] == 1:
                raise RuntimeError("database connection lost")
            return real(order, ref)

        with patch("order_intake_poll._supersede_older_revisions", flaky):
            self.assertEqual(self.run_poller(), 2)
            self.assertEqual((len(self.orders), len(self.approval_rows), len(self.mailer.sent)), (1, 0, 0))
            self.assertEqual(self.run_poller(), 0)
        self.assertEqual((len(self.orders), len(self.approval_rows), len(self.mailer.sent)), (1, 1, 1))
        self.assertEqual(self.incoming[0]["processing_status"], "processed")

    def test_email_sent_but_state_not_recorded_is_the_documented_residual_risk(self):
        self.provider.messages = [make_message(1)]
        real = approvals.mark_approval_email_sent
        calls = {"n": 0}

        def flaky(approval_id):
            calls["n"] += 1
            if calls["n"] == 1:
                raise RuntimeError("database connection lost")
            return real(approval_id)

        with patch.object(approvals, "mark_approval_email_sent", flaky):
            self.assertEqual(self.run_poller(), 2)  # email went out, delivery state not saved
            self.assertEqual(len(self.mailer.sent), 1)
            self.assertEqual(self.approval_rows[0]["email_status"], "sending")

            self.assertEqual(self.run_poller(), 2)  # lease still active: no second email yet
            self.assertEqual(len(self.mailer.sent), 1)

            later = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(minutes=25)
            with patch.object(approvals, "_utcnow", return_value=later):
                self.assertEqual(self.run_poller(), 0)  # lease expired: controlled resend

        self.assertEqual(len(self.mailer.sent), 2)
        self.assertTrue(self.mailer.sent[1]["subject"].startswith("[Resent]"))
        self.assertEqual((len(self.orders), len(self.approval_rows), len(self.incoming)), (1, 1, 1))

    def test_gives_up_after_max_attempts_and_alerts(self):
        self.provider.messages = [make_message(1)]
        self.mailer.fail_next = 1000
        for _ in range(approvals.MAX_INTAKE_ATTEMPTS):
            self.assertEqual(self.run_poller(), 2)
        self.assertEqual(self.incoming[0]["attempt_count"], approvals.MAX_INTAKE_ATTEMPTS)
        self.assertEqual(len(self.orders), 1)
        titles = [n["title"] for n in self.db.rows("crm_notifications")]
        self.assertIn("Order intake gave up on an email", titles)

        remaining = self.mailer.fail_next
        self.assertEqual(self.run_poller(), 0)  # no further attempt
        self.assertEqual(self.mailer.fail_next, remaining)
        self.assertEqual(self.counters()["abandoned"], 1)

    def test_exhausted_order_stays_visible_and_can_be_retried_safely(self):
        self.provider.messages = [make_message(1)]
        self.mailer.fail_next = 1000
        for _ in range(approvals.MAX_INTAKE_ATTEMPTS):
            self.run_poller()

        order = approvals.list_purchase_orders()[0]  # still listed on the Order Approvals page
        self.assertEqual(order["status"], "pending_approval")
        state = approvals.get_order_delivery_state(order)
        self.assertEqual((state["intake_status"], state["attempt_count"]), ("error", approvals.MAX_INTAKE_ATTEMPTS))
        self.assertEqual(state["email_status"], "failed")
        self.assertIn("Graph sendMail failed", state["error_message"])
        self.assertEqual(self.mailer.sent, [])

        self.assertTrue(approvals.retry_order_intake(order["id"]))
        self.assertTrue(approvals.retry_order_intake(order["id"]))  # clicking twice is harmless
        self.mailer.fail_next = 0
        self.assertEqual(self.run_poller(), 0)
        self.assertEqual(self.run_poller(), 0)
        self.assertEqual((len(self.orders), len(self.approval_rows), len(self.incoming)), (1, 1, 1))
        self.assertEqual(len(self.mailer.sent), 1)
        self.assertEqual(self.incoming[0]["processing_status"], "processed")
        self.assertFalse(approvals.retry_order_intake(order["id"]))  # nothing left to retry

    def test_failed_ingest_leaves_no_partial_row_and_is_retried(self):
        self.provider.messages = [make_message(1)]
        with patch.object(approvals, "upload_purchase_order_pdf", side_effect=RuntimeError("storage down")):
            self.assertEqual(self.run_poller(), 2)
        self.assertEqual((self.incoming, self.orders), ([], []))
        self.assertEqual(self.run_poller(), 0)
        self.assertEqual((len(self.incoming), len(self.orders), len(self.mailer.sent)), (1, 1, 1))


class TestRevisionHandlingPreserved(IntakeTestCase):
    def _revision_messages(self):
        subject = "Purchase order PO-77881 for Acme Retail"
        return [
            make_message(1, subject=subject, pdf=b"%PDF rev one"),
            make_message(2, subject=subject, pdf=b"%PDF rev two", received=NOW.replace(hour=12, minute=30)),
        ]

    def test_newer_revision_supersedes_the_open_one(self):
        self.provider.messages = self._revision_messages()
        self.assertEqual(self.run_poller(), 0)
        first, second = sorted(self.orders, key=lambda o: o["revision"])
        self.assertEqual((first["status"], first["is_current"]), ("superseded", False))
        self.assertEqual((second["status"], second["revision"]), ("pending_approval", 2))
        self.assertEqual(first["superseded_by"], second["id"])
        old_token = [a for a in self.approval_rows if a["purchase_order_id"] == first["id"]][0]
        self.assertEqual(old_token["status"], "expired")

    def test_interrupted_supersede_is_completed_on_resume(self):
        self.provider.messages = self._revision_messages()
        real = order_intake_poll._supersede_older_revisions
        calls = {"n": 0}

        def flaky(order, ref):
            calls["n"] += 1
            if calls["n"] == 2:
                raise RuntimeError("database connection lost")
            return real(order, ref)

        with patch("order_intake_poll._supersede_older_revisions", flaky):
            self.assertEqual(self.run_poller(), 2)
            self.assertEqual(sorted(o["status"] for o in self.orders), ["pending_approval", "pending_approval"])
            self.assertEqual(self.run_poller(), 0)
        self.assertEqual(sorted(o["status"] for o in self.orders), ["pending_approval", "superseded"])
        self.assertEqual(len(self.orders), 2)


class TestConcurrency(IntakeTestCase):
    def test_two_simultaneous_runs_create_one_order_and_one_email(self):
        self.provider = FakeProvider([make_message(1)], list_barrier=threading.Barrier(2))
        self.mailer.delay = 0.2
        results = []

        def worker():
            results.append(main([]))

        with patch("order_intake_poll.get_email_provider", return_value=self.provider):
            threads = [threading.Thread(target=worker) for _ in range(2)]
            for thread in threads:
                thread.start()
            for thread in threads:
                thread.join(timeout=30)

        self.assertEqual(sorted(results), [0, 0])
        self.assertEqual((len(self.incoming), len(self.orders), len(self.approval_rows)), (1, 1, 1))
        self.assertEqual(len(self.mailer.sent), 1)


class TestLoggingAndSummary(IntakeTestCase):
    def test_summary_reports_counters_and_logs_contain_no_secrets(self):
        self.provider.messages = [make_message(1), make_message(2, has_pdf=False)]
        with self.assertLogs("order_intake_poll", level="INFO") as captured:
            self.assertEqual(self.run_poller(), 0)
        logged = "\n".join(captured.output)
        for forbidden in (SECRET, "PRIVATE BODY TEXT", "%PDF", "Bearer"):
            self.assertNotIn(forbidden, logged)
        self.assertIn("event=config", logged)
        self.assertIn("provider=FakeProvider", logged)
        self.assertIn("graph_client_secret=set", logged)
        counters = self.counters()
        self.assertEqual(
            (counters["scanned"], counters["eligible"], counters["skipped"], counters["created"],
             counters["emailed"], counters["failed"]),
            (2, 1, 1, 1, 1, 0),
        )
        with open(self.summary_path, encoding="utf-8") as handle:
            self.assertNotIn(SECRET, handle.read())

    def test_annotations_are_emitted_on_github_actions(self):
        os.environ["GITHUB_ACTIONS"] = "true"
        self.provider.preflight_error = IntakeError("auth", "Token request failed: HTTP 401", status=401)
        with patch("builtins.print") as printed:
            self.assertEqual(self.run_poller(), 1)
        self.assertTrue(any(str(c.args[0]).startswith("::error::") for c in printed.call_args_list))

    SENSITIVE = ("Acme", "Springfield", "pat.buyer", "customer-a", "PO-77881", "Quokka", "contoso")

    def _customer_message(self):
        return make_message(
            1, sender="Pat Buyer <pat.buyer@customer-a.example>",
            subject="PO-77881 Quokka AD Acme Springfield",
            filename="2026-10-001_AD_Acme_Springfield.pdf",
        )

    def _assert_no_customer_data(self, text):
        for needle in self.SENSITIVE:
            self.assertNotIn(needle, text)

    def test_dry_run_and_live_logs_contain_no_customer_information(self):
        self.provider.messages = [self._customer_message()]
        with self.assertLogs(level="DEBUG") as captured:
            self.assertEqual(self.run_poller("--dry-run"), 0)
            self.assertEqual(self.run_poller(), 0)
        self._assert_no_customer_data("\n".join(captured.output))
        with open(self.summary_path, encoding="utf-8") as handle:
            self._assert_no_customer_data(handle.read())

    def test_failure_details_stay_in_the_database_not_in_logs(self):
        self.provider.messages = [self._customer_message()]
        leaky = RuntimeError("sendMail rejected for AD Acme Springfield <pat.buyer@customer-a.example>")
        with patch("order_intake_poll.send_graph_email", side_effect=leaky), \
                self.assertLogs(level="DEBUG") as captured:
            self.assertEqual(self.run_poller(), 2)
        self._assert_no_customer_data("\n".join(captured.output))
        with open(self.summary_path, encoding="utf-8") as handle:
            self._assert_no_customer_data(handle.read())
        self.assertIn("Springfield", self.incoming[0]["error_message"])  # visible to admins in the CRM

    def test_unexpected_errors_are_logged_by_type_only(self):
        self.provider.messages = [self._customer_message()]
        with patch.object(self.provider, "list_message_headers", side_effect=ValueError("Acme leak")), \
                self.assertLogs(level="DEBUG") as captured:
            self.assertEqual(self.run_poller(), 1)
        self.assertNotIn("Acme", "\n".join(captured.output))
        self.assertIn("ValueError", "\n".join(captured.output))

    def test_http_client_loggers_are_silenced(self):
        for name in ("httpx", "httpcore", "urllib3", "postgrest", "storage3"):
            self.assertGreaterEqual(logging.getLogger(name).getEffectiveLevel(), logging.WARNING)

    def test_graph_send_logs_neither_subject_nor_recipients(self):
        class Ok:
            def raise_for_status(self):
                return None

            def json(self):
                return {"access_token": "tok"}

        env = {"ORDER_INTAKE_GRAPH_TENANT_ID": "t", "ORDER_INTAKE_GRAPH_CLIENT_ID": "c",
               "ORDER_INTAKE_GRAPH_CLIENT_SECRET": SECRET, "ORDER_INTAKE_GRAPH_MAILBOX": "orders@contoso.example"}
        with patch.dict(os.environ, env), patch("requests.post", return_value=Ok()), \
                self.assertLogs(level="DEBUG") as captured:
            email_service.send_graph_email(
                "Purchase order approval needed: AD Acme Springfield", "body", ["ceo@example.com"])
        text = "\n".join(captured.output)
        for needle in ("Acme", "ceo@example.com", "tok", SECRET):
            self.assertNotIn(needle, text)


class TestApprovalAndInvoiceFlowPreserved(IntakeTestCase):
    def test_approval_link_still_approves_and_creates_the_invoice(self):
        self.provider.messages = [make_message(1)]
        self.assertEqual(self.run_poller(), 0)
        body = self.mailer.sent[0]["body"]
        token = re.search(r"\?approval=([A-Za-z0-9_-]+)", body).group(1)
        self.assertEqual(self.mailer.sent[0]["recipients"], ["ceo@example.com"])
        self.assertEqual(self.mailer.sent[0]["attachment_filename"], self.incoming[0]["pdf_filename"])

        result = approvals.apply_decision(token, "approve", comment="ok")
        self.assertTrue(result["success"])
        self.assertEqual(self.orders[0]["status"], "approved")
        invoices = self.db.rows("invoices")
        self.assertEqual(len(invoices), 1)
        self.assertEqual(invoices[0]["source_purchase_order_id"], self.orders[0]["id"])
        self.assertEqual(invoices[0]["send_status"], "ready_to_send")

        again = approvals.apply_decision(token, "approve")
        self.assertFalse(again["success"])
        self.assertEqual(len(self.db.rows("invoices")), 1)


if __name__ == "__main__":
    unittest.main()
