"""Opt-in SQL/concurrency tests. Only an explicitly named loopback confirmation_test DB is accepted.

Set CADDYCHECK_CONFIRMATION_TEST_DB to a fresh local PostgreSQL confirmation_test database URL.
No Supabase credentials, production connections, email, or network services are used.
"""
import concurrent.futures
import json
import os
from pathlib import Path
import shutil
import subprocess
import unittest
from urllib.parse import urlparse
import uuid


DATABASE = os.environ.get("CADDYCHECK_CONFIRMATION_TEST_DB", "")
PSQL = shutil.which("psql")
MIGRATION = Path(__file__).resolve().parents[1] / "migrations" / "create_order_confirmation_tracking.sql"
ORDER_ID = "10000000-0000-0000-0000-000000000001"
OTHER_ID = "10000000-0000-0000-0000-000000000002"

BOOTSTRAP = """
do $$ begin
    if not exists(select 1 from pg_roles where rolname='anon') then create role anon; end if;
    if not exists(select 1 from pg_roles where rolname='authenticated') then create role authenticated; end if;
    if not exists(select 1 from pg_roles where rolname='service_role') then create role service_role bypassrls; end if;
end $$;
create table if not exists public.purchase_orders(
    id uuid primary key, project_name text, amount numeric, currency text default 'EUR',
    order_reference text, customer_name text, customer_email text, incoming_email_id uuid,
    status text default 'pending_approval', is_current boolean default true
);
create table if not exists public.invoices(
    id bigint generated always as identity primary key, invoice_number text, project_name text,
    maintenance_year text, payment_amount numeric, cameras_number integer, payment_date date,
    paid text, year integer, description text, invoice_type text, for_month text, sent_at timestamptz,
    source_type text, source_purchase_order_id uuid, auto_created boolean,
    auto_created_at timestamptz, send_status text
);
create table if not exists public.invoice_number_registry(
    invoice_number bigint primary key, status text default 'active', source text,
    description_key text, created_by text, total_amount numeric, ledger_saved_at timestamptz
);
create table if not exists public.invoice_change_log(
    id bigint generated always as identity primary key, invoice_number text, source text
);
create or replace function public.invoice_number_key(p text) returns text language sql immutable
as $$ select nullif(regexp_replace(trim(coalesce(p,'')), '\\.0+$','','g'),'') $$;
create or replace function public.test_invoice_audit() returns trigger language plpgsql as $$
begin
    insert into public.invoice_change_log(invoice_number,source)
        values(new.invoice_number, current_setting('caddycheck.source',true));
    return new;
end $$;
create or replace trigger invoices_audit after insert on public.invoices
    for each row execute function public.test_invoice_audit();
grant usage on schema public to anon, authenticated, service_role;
grant all on public.purchase_orders, public.invoices, public.invoice_number_registry,
    public.invoice_change_log to service_role;
grant usage, select on all sequences in schema public to service_role;
"""


@unittest.skipUnless(DATABASE and PSQL, "requires explicit local CADDYCHECK_CONFIRMATION_TEST_DB and psql")
class TestOrderConfirmationSQL(unittest.TestCase):
    @classmethod
    def sql(cls, statement, *, fail=False):
        result = subprocess.run(
            [PSQL, DATABASE, "-X", "-A", "-t", "-q", "-v", "ON_ERROR_STOP=1"],
            input=statement, text=True, capture_output=True, encoding="utf-8", timeout=20,
        )
        if fail:
            if result.returncode == 0:
                raise AssertionError("Expected SQL rejection")
            return result.stderr
        if result.returncode:
            raise AssertionError(result.stderr)
        return result.stdout.strip()

    @classmethod
    def setUpClass(cls):
        parsed = urlparse(DATABASE)
        if parsed.scheme != "postgresql" or parsed.hostname != "127.0.0.1" or parsed.path != "/confirmation_test":
            raise RuntimeError("Refusing anything except an explicit loopback confirmation_test database")
        cls.sql(BOOTSTRAP)
        cls.sql(MIGRATION.read_text(encoding="utf-8"))
        # Migration must be re-runnable without resetting tracking or rewriting operational rows.
        cls.sql(MIGRATION.read_text(encoding="utf-8"))

    def setUp(self):
        self.sql("""
            truncate public.order_confirmation_log, public.invoices, public.invoice_change_log,
                     public.invoice_number_registry, public.purchase_orders restart identity cascade;
            insert into public.purchase_orders(id, project_name, amount, order_reference)
                values('10000000-0000-0000-0000-000000000001','AD Roeselare',9336,'2026-01-005'),
                      ('10000000-0000-0000-0000-000000000002','Edeka Site',2334,'PO-2');
            insert into public.invoices(invoice_number,project_name,payment_amount,paid)
                values('8718','Existing untouched invoice',2334,'Yes');
            insert into public.invoice_number_registry(invoice_number,status,source)
                values(8725,'void','earlier-issued-number');
        """)
        self.token = str(uuid.uuid4())
        self.expected = {"id": ORDER_ID, "project_name": "AD Roeselare", "amount": 9336,
                         "status": "pending_approval"}

    @staticmethod
    def literal(value):
        if value is None:
            return "null"
        if isinstance(value, bool):
            return str(value).lower()
        if isinstance(value, (dict, list)):
            return "'" + json.dumps(value).replace("'", "''") + "'::jsonb"
        return "'" + str(value).replace("'", "''") + "'"

    def call(self, name, *args, fail=False):
        text = self.sql(f"select public.{name}({','.join(self.literal(a) for a in args)});", fail=fail)
        return text if fail else json.loads(text)

    def claim(self, order=ORDER_ID, token=None, resend=False, expected=None):
        return self.call("claim_order_confirmation", order, token or self.token,
                         expected if expected is not None else self.expected, "Test Admin", resend)

    def reserve(self):
        return self.call("reserve_order_confirmation_invoice", ORDER_ID, self.token, "Test Admin")

    def payload(self, number=8726):
        return [{"invoice_number": str(number), "project_name": "AD Roeselare", "maintenance_year": "Y1",
                 "payment_amount": 9336, "cameras_number": 12, "paid": "No", "year": 2026,
                 "invoice_type": "New installation", "description": "Order 2026-01-005"}]

    def saved_invoice(self):
        self.claim()
        number = self.reserve()
        rows = self.call("save_order_confirmation_invoice", ORDER_ID, self.token, self.payload(number), "Test Admin")
        return rows

    def start(self, rows):
        details = {"sender": "orders.caddycheck@video-inform.com", "to": ["tom@example.com","karen@example.com"],
                   "cc": ["yoram@example.com","zvika@example.com"], "subject": "Order confirmation",
                   "body": "Saved invoice"}
        self.call("start_order_confirmation_send", ORDER_ID, self.token, rows, "JVBERi0xLjQ=", details, "Test Admin")
        return details

    def finish(self, outcome, details, safe=False):
        return self.call("finish_order_confirmation", ORDER_ID, self.token, outcome, "email",
                         None if outcome == "sent" else "test error", safe, details, "Test Admin")

    def test_additive_migration_does_not_change_existing_order_business_fields(self):
        before = self.sql("select json_agg(row(id,project_name,amount,currency,order_reference,status)) from purchase_orders;")
        self.sql(MIGRATION.read_text(encoding="utf-8"))
        after = self.sql("select json_agg(row(id,project_name,amount,currency,order_reference,status)) from purchase_orders;")
        self.assertEqual(before, after)
        self.assertEqual(self.sql("select count(*) from purchase_orders where confirmation_status='none' and confirmation_invoice_number is null;"), "2")

    def test_concurrent_claims_only_one_winner(self):
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(lambda _: self.claim(token=str(uuid.uuid4())), range(2)))
        self.assertEqual(sorted(r["state"] for r in results), ["busy", "claimed"])

    def test_concurrent_orders_have_distinct_numbers_above_ledger_and_all_registry_statuses(self):
        second_token = str(uuid.uuid4())
        self.claim()
        self.claim(OTHER_ID, second_token, expected={"id": OTHER_ID})
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
            first = pool.submit(self.reserve)
            second = pool.submit(self.call, "reserve_order_confirmation_invoice", OTHER_ID, second_token, "Other Admin")
            self.assertEqual(sorted([first.result(), second.result()]), [8726,8727])
        self.assertEqual(self.reserve(), 8726 if self.sql(
            f"select confirmation_invoice_number from purchase_orders where id='{ORDER_ID}';") == "8726" else 8727)

    def test_save_is_insert_only_idempotent_and_audited_by_existing_trigger(self):
        rows = self.saved_invoice()
        again = self.call("save_order_confirmation_invoice", ORDER_ID, self.token, self.payload(), "Test Admin")
        self.assertEqual(rows, again)
        self.assertEqual(self.sql("select count(*) from invoices where invoice_number='8726';"), "1")
        self.assertEqual(self.sql("select payment_amount||':'||paid from invoices where invoice_number='8718';"), "2334:Yes")
        self.assertEqual(self.sql("select count(*) from invoice_change_log where invoice_number='8726' and source='admin-order-confirmation';"), "1")
        self.assertIn("reserved", self.sql(
            "insert into invoices(invoice_number,project_name,payment_amount) values('8726','Unrelated',999);", fail=True))

    def test_invalid_payload_rolls_back_save_but_preserves_permanent_reservation(self):
        self.claim()
        self.reserve()
        bad = self.payload()
        bad[0]["payment_amount"] = -1
        self.call("save_order_confirmation_invoice", ORDER_ID, self.token, bad, "Test Admin", fail=True)
        self.assertEqual(self.sql("select count(*) from invoices where invoice_number='8726';"), "0")
        self.assertEqual(self.reserve(), 8726)
        self.sql(f"update purchase_orders set confirmation_invoice_number='9000' where id='{ORDER_ID}';", fail=True)

    def test_no_adoption_of_unrelated_existing_invoice_or_legacy_order_invoice(self):
        self.sql(f"insert into invoices(invoice_number,source_purchase_order_id) values('8100','{ORDER_ID}');")
        self.assertIn("Legacy invoice", self.call("claim_order_confirmation", ORDER_ID, self.token,
                                                self.expected, "Admin", False, fail=True))
        self.assertEqual(self.sql("select count(*) from invoice_number_registry where source='order-confirmation';"), "0")

    def test_changed_invoice_blocks_pending_approval_and_send(self):
        rows = self.saved_invoice()
        self.sql("update invoices set payment_amount=1 where invoice_number='8726';")
        self.assertIn("changed while rendering", self.call(
            "start_order_confirmation_send", ORDER_ID, self.token, rows, "PDF", {}, "Admin", fail=True))
        self.assertEqual(self.sql(f"select status from purchase_orders where id='{ORDER_ID}';"), "pending_approval")

    def test_removed_or_financially_changed_saved_invoice_cannot_be_recreated_or_resent(self):
        rows = self.saved_invoice()
        self.sql("update invoices set payment_amount=1 where invoice_number='8726';")
        self.call("save_order_confirmation_invoice", ORDER_ID, self.token, self.payload(), "Admin", fail=True)
        self.sql("delete from invoices where invoice_number='8726';")
        self.assertIn("removed", self.call(
            "save_order_confirmation_invoice", ORDER_ID, self.token, self.payload(), "Admin", fail=True))
        self.assertEqual(self.sql("select count(*) from invoices where invoice_number='8726';"), "0")

    def test_order_changed_since_validation_or_rejected_during_pdf_cannot_be_approved(self):
        self.sql(f"update purchase_orders set amount=1 where id='{ORDER_ID}';")
        self.call("claim_order_confirmation", ORDER_ID, self.token, self.expected, "Admin", False, fail=True)
        self.sql(f"update purchase_orders set amount=9336 where id='{ORDER_ID}';")
        rows = self.saved_invoice()
        self.sql(f"update purchase_orders set status='rejected' where id='{ORDER_ID}';")
        self.assertIn("eligibility changed", self.call(
            "start_order_confirmation_send", ORDER_ID, self.token, rows, "PDF", {}, "Admin", fail=True))
        self.assertEqual(self.sql(f"select status from purchase_orders where id='{ORDER_ID}';"), "rejected")

    def test_success_and_audit_are_atomic_and_finalize_is_idempotent(self):
        rows = self.saved_invoice()
        details = self.start(rows)
        result = self.finish("sent", details)
        self.assertEqual(result["status"], "approved")
        self.assertEqual(result["confirmation_status"], "sent")
        self.assertIsNotNone(result["confirmation_sent_at"])
        self.finish("sent", details)
        self.assertEqual(self.sql("select count(*) from order_confirmation_log;"), "1")
        self.assertEqual(self.claim()["state"], "sent")
        self.assertEqual(self.sql("select invoice_number||':'||invoice_amount||':'||approved_by from order_confirmation_log;"),
                         "8726:9336:Test Admin")

    def test_audit_insert_failure_rolls_back_sent_status_and_keeps_claim(self):
        rows = self.saved_invoice()
        details = self.start(rows)
        self.sql("""
            create or replace function public.test_reject_audit() returns trigger language plpgsql as $$
            begin raise exception 'Injected audit failure'; end $$;
            create trigger reject_audit before insert on order_confirmation_log
                for each row execute function public.test_reject_audit();
        """)
        try:
            self.call("finish_order_confirmation", ORDER_ID, self.token, "sent", "complete",
                      None, False, details, "Admin", fail=True)
            self.assertEqual(self.sql(f"select confirmation_status from purchase_orders where id='{ORDER_ID}';"), "pending")
            self.assertEqual(self.claim()["state"], "busy")
            self.assertEqual(self.sql("select count(*) from order_confirmation_log;"), "0")
        finally:
            self.sql("drop trigger reject_audit on order_confirmation_log;")

    def test_failure_resend_reuses_invoice_and_frozen_pdf(self):
        rows = self.saved_invoice()
        details = self.start(rows)
        self.finish("failed", details, safe=True)
        self.token = str(uuid.uuid4())
        state = self.claim(resend=True, expected={"id": ORDER_ID})
        self.assertEqual(state["state"], "claimed")
        self.assertEqual(self.reserve(), 8726)
        self.assertEqual(self.call("save_order_confirmation_invoice", ORDER_ID, self.token, [], "Admin"), rows)
        self.call("start_order_confirmation_send", ORDER_ID, self.token, rows, "DIFFERENT", details, "Admin", fail=True)
        self.call("start_order_confirmation_send", ORDER_ID, self.token, rows, "JVBERi0xLjQ=", details, "Admin")
        self.finish("sent", details)
        self.assertEqual(self.sql("select count(*) from invoices where invoice_number='8726';"), "1")
        self.assertEqual(self.sql("select count(*) from order_confirmation_log;"), "2")

    def test_unknown_send_blocks_all_new_attempts_and_public_cannot_call_rpc(self):
        rows = self.saved_invoice()
        details = self.start(rows)
        result = self.finish("unknown", details)
        self.assertFalse(result["confirmation_retry_safe"])
        self.assertEqual(self.claim(token=str(uuid.uuid4()), resend=True)["state"], "busy")
        self.assertIn("permission denied", self.sql(
            f"set role authenticated; select public.reserve_order_confirmation_invoice('{ORDER_ID}','{self.token}','Admin');",
            fail=True))
