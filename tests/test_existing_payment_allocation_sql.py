"""Real SQL tests in a newly named disposable Docker container; never use app credentials."""
from concurrent.futures import ThreadPoolExecutor
import json
from pathlib import Path
import subprocess
import time
import unittest
import uuid

try:
    import psycopg
except ImportError:
    psycopg = None

from tests.test_apply_credit_atomic_sql import SCHEMA, _docker_ready, install_production_like_rls

MIGRATION = Path(__file__).resolve().parents[1] / "migrations" / "add_existing_payment_allocation.sql"


@unittest.skipUnless(_docker_ready(), "Disposable Docker PostgreSQL is unavailable")
class TestExistingPaymentAllocationSql(unittest.TestCase):
    container = ""

    @classmethod
    def setUpClass(cls):
        cls.container = f"caddy-allocation-test-{uuid.uuid4().hex[:10]}"
        result = subprocess.run(
            ["docker", "run", "-d", "--rm", "--name", cls.container,
             "-p", "127.0.0.1::5432", "-e", "POSTGRES_PASSWORD=isolated-test-only", "postgres:17"],
            capture_output=True, text=True, timeout=180,
        )
        if result.returncode:
            raise RuntimeError(result.stderr)
        try:
            for _ in range(60):
                probe = subprocess.run(
                    ["docker", "exec", cls.container, "psql", "-h", "127.0.0.1", "-U", "postgres", "-c", "select 1"],
                    capture_output=True,
                )
                if probe.returncode == 0:
                    break
                time.sleep(1)
            else:
                raise RuntimeError("Disposable PostgreSQL did not become ready")
            cls.sql(SCHEMA)
            install_production_like_rls(cls.sql)
            cls.sql("alter table invoices add column currency text default 'EUR';"
                    "grant select, insert, update, delete on invoices, bank_payments, bank_payment_allocations to service_role;"
                    "grant usage, select on all sequences in schema public to service_role;")
            cls.sql("""
                create role production_installer login password 'isolated-test-only'
                    nosuperuser inherit createrole createdb bypassrls;
                grant pg_read_all_data to production_installer;
                alter schema public owner to production_installer;
                grant usage on schema public to anon,authenticated,service_role;
                grant create on database postgres to production_installer;
                do $$ declare r record; begin
                    for r in select c.oid::regclass as name from pg_class c
                             join pg_namespace n on n.oid=c.relnamespace
                             where n.nspname='public' and c.relkind='r' loop
                        execute format('alter table %s owner to production_installer',r.name);
                    end loop;
                    for r in select p.oid::regprocedure as name from pg_proc p
                             join pg_namespace n on n.oid=p.pronamespace
                             where n.nspname='public' loop
                        execute format('alter function %s owner to production_installer',r.name);
                    end loop;
                end $$;
            """)
            cls.sql((MIGRATION.parent / "protect_legacy_payment_allocations.sql").read_text(encoding="utf-8"),
                    user="production_installer")
            cls.sql(MIGRATION.read_text(encoding="utf-8"), user="production_installer")
            cls.default_release_state = cls.sql("select enabled from manual_allocation_release_policy;")
            cls.sql((MIGRATION.parent / "create_apply_credit_atomic.sql").read_text(encoding="utf-8"))
            port = subprocess.run(
                ["docker", "port", cls.container, "5432/tcp"], capture_output=True, text=True, check=True,
            ).stdout.strip()
            cls.port = int(port.rsplit(":", 1)[1])
        except Exception:
            cls.tearDownClass()
            raise

    @classmethod
    def tearDownClass(cls):
        if cls.container:
            subprocess.run(["docker", "rm", "-f", cls.container], capture_output=True, timeout=30)

    @classmethod
    def sql(cls, statement, *, user="postgres"):
        result = subprocess.run(
            ["docker", "exec", "-i", "-e", "PGPASSWORD=isolated-test-only", cls.container,
             "psql", "-h", "127.0.0.1", "-U", user, "-d", "postgres",
             "-v", "ON_ERROR_STOP=1", "-At", "-f", "-"],
            input=statement, capture_output=True, text=True, encoding="utf-8", timeout=30,
        )
        if result.returncode:
            raise RuntimeError(result.stderr)
        return result.stdout.strip()

    def setUp(self):
        self.sql("""
            insert into manual_allocation_release_policy(singleton,enabled) values (true,true)
                on conflict (singleton) do update set enabled=true;
            truncate invoice_change_log;
            truncate legacy_allocation_reversals;
            drop trigger if exists fail_summary on bank_payments;
            drop trigger if exists bank_payments_writer_guard on bank_payments;
            truncate invoices, bank_payments restart identity cascade;
            insert into bank_payments
              (id,payment_date,invoice_number,source_name,source_kind,payment_fingerprint,
               instructed_amount,received_amount,applied_amount,fee_amount,currency,
               raw_text,parsed_payload,notes,pdf_storage_bucket,pdf_storage_path)
            values (14,'2026-10-20','8596,8597,8700,8709','Transaction_21_Oct_2026.pdf','pdf',
                    'original-fingerprint',46680,46616.32,38900,63.68,'EUR',
                    'unchanged raw text','{"original":true}','unchanged notes','bank-payments','original.pdf');
            insert into invoices(id,invoice_number,project_name,maintenance_year,payment_amount,paid,payment_date,year)
            values (500,'8596','AD Wondelgem','Y1',9336,'Yes','2026-10-20',2026),
                   (501,'8597','AD Tienen','Y1',15560,'Yes','2026-10-20',2026),
                   (3080,'8700','AD Pallieter','Y1',6224,'Yes','2026-10-20',2026),
                   (3089,'8709','Proxy Lovenjoel','Y1',7780,'Yes','2026-10-20',2026),
                   (3087,'8707','Intermarche Heusy','Y1',7780,'No',null,2026),
                   (999,'9999','Unrelated','Y2',100,'No',null,2026);
            insert into bank_payment_allocations
              (payment_id,invoice_row_id,invoice_number,project_name,maintenance_year,year,amount_applied)
            select 14,id,invoice_number::bigint,project_name,maintenance_year,year,payment_amount
            from invoices where id in (500,501,3080,3089);
        """)

    def snapshot(self):
        return json.loads(self.sql("""
            select jsonb_build_object(
                'payment',(select to_jsonb(p) from bank_payments p where id=14),
                'invoice',(select to_jsonb(i) from invoices i where id=3087),
                'allocations',(select jsonb_agg(to_jsonb(a) order by id) from bank_payment_allocations a where payment_id=14)
            );
        """))

    def ledger(self):
        return json.loads(self.sql("""
            select jsonb_build_object(
                'payments',(select jsonb_agg(to_jsonb(p) order by id) from bank_payments p),
                'invoices',(select jsonb_agg(to_jsonb(i) order by id) from invoices i),
                'allocations',(select jsonb_agg(to_jsonb(a) order by id) from bank_payment_allocations a)
            );
        """))

    def apply(self, snapshot=None, row_id=3087, role="service_role"):
        snapshot = self.snapshot() if snapshot is None else snapshot
        payload = json.dumps(snapshot).replace("'", "''")
        return json.loads(self.sql(
            f"set role {role}; select public.add_invoice_to_existing_payment(14,{row_id},'{payload}'::jsonb);"
        ).splitlines()[-1])

    def test_database_policy_defaults_off_and_blocks_direct_rpc(self):
        self.assertEqual(self.default_release_state, "f")
        self.sql("update manual_allocation_release_policy set enabled=false;")
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "database release policy"):
            self.apply()
        self.assertEqual(self.ledger(), before)

    def test_missing_release_policy_row_blocks_direct_rpc(self):
        self.sql("delete from manual_allocation_release_policy;")
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "database release policy"):
            self.apply()
        self.assertEqual(self.ledger(), before)

    def test_service_and_browser_roles_cannot_enable_release_or_truncate_policy(self):
        for role in ("service_role", "anon", "authenticated"):
            for operation in (
                "update manual_allocation_release_policy set enabled=true",
                "insert into manual_allocation_release_policy values (true,true)",
                "delete from manual_allocation_release_policy",
                "truncate manual_allocation_release_policy",
                "alter table manual_allocation_release_policy disable row level security",
            ):
                with self.subTest(role=role, operation=operation), self.assertRaisesRegex(
                    RuntimeError, "permission denied|must be owner",
                ):
                    self.sql(f"set role {role}; {operation};")

    def test_caller_guc_and_snapshot_cannot_enable_database_release(self):
        self.sql("update manual_allocation_release_policy set enabled=false;")
        snapshot = self.snapshot()
        snapshot["CADDYCHECK_MANUAL_ALLOCATION_ENABLED"] = "on"
        payload = json.dumps(snapshot).replace("'", "''")
        with self.assertRaisesRegex(RuntimeError, "database release policy"):
            self.sql("set role service_role; set caddycheck.manual_allocation_enabled='on';"
                     f"select public.add_invoice_to_existing_payment(14,3087,'{payload}'::jsonb);")

    def test_reinstall_collision_does_not_reset_release_or_modify_financial_data(self):
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "already exists"):
            self.sql(MIGRATION.read_text(encoding="utf-8"))
        self.assertEqual(self.sql("select enabled from manual_allocation_release_policy;"), "t")
        self.assertEqual(self.ledger(), before)

    def test_release_off_does_not_block_explicit_reversal(self):
        result = self.apply()
        self.sql("update manual_allocation_release_policy set enabled=false;")
        snapshot = self.snapshot()
        snapshot["actor"] = "admin"
        payload = json.dumps(snapshot).replace("'", "''")
        result = json.loads(self.sql(
            f"set role service_role; select public.reverse_manual_payment_allocation("
            f"{result['allocation_id']},'{payload}'::jsonb,'admin','approved correction',true);"
        ).splitlines()[-1])
        self.assertEqual(result["status"], "reversed")

    def test_exact_reconciliation_and_preservation(self):
        before = self.ledger()
        result = self.apply()
        self.assertEqual(result["status"], "added")
        self.assertEqual(result["payment"]["applied_amount"], 46680)
        self.assertEqual(result["payment"]["instructed_amount"], 46680)
        self.assertEqual(result["invoice"]["paid"], "Yes")
        self.assertEqual(result["invoice"]["payment_date"], "2026-10-20")
        after = self.ledger()
        self.assertEqual(after["allocations"][:4], before["allocations"])
        self.assertEqual(len(after["allocations"]), 5)
        self.assertEqual(len(after["payments"]), 1)
        self.assertEqual(len(after["invoices"]), len(before["invoices"]))
        for key, value in before["payments"][0].items():
            if key not in ("invoice_number", "applied_amount", "updated_at"):
                self.assertEqual(after["payments"][0][key], value, key)
        for original in before["invoices"]:
            saved = next(row for row in after["invoices"] if row["id"] == original["id"])
            if original["id"] == 3087:
                for key, value in original.items():
                    if key not in ("paid", "payment_date"):
                        self.assertEqual(saved[key], value, key)
            else:
                self.assertEqual(saved, original)

    def test_success_leaving_remaining_balance(self):
        self.sql("update bank_payments set instructed_amount=50000, received_amount=49936.32 where id=14;")
        result = self.apply()
        self.assertEqual(result["payment"]["instructed_amount"] - result["payment"]["applied_amount"], 3320)

    def test_idempotent_duplicate_click(self):
        snapshot = self.snapshot()
        first = self.apply(snapshot)
        before = self.ledger()
        second = self.apply(snapshot)
        self.assertEqual(second["status"], "already_allocated")
        self.assertEqual(first["allocation_id"], second["allocation_id"])
        self.assertEqual(self.ledger(), before)

    def test_concurrent_confirmations(self):
        snapshot = self.snapshot()
        with ThreadPoolExecutor(max_workers=2) as pool:
            def attempt(_):
                try:
                    return self.apply(snapshot)["status"]
                except RuntimeError as exc:
                    self.assertIn("busy", str(exc))
                    return "busy"
            results = list(pool.map(attempt, range(2)))
        self.assertIn("added", results)
        self.assertTrue(set(results) <= {"added", "already_allocated", "busy"})
        self.assertEqual(len(self.ledger()["allocations"]), 5)

    def test_concurrent_different_payments_cannot_settle_same_invoice(self):
        snapshot = self.snapshot()
        self.sql("""insert into bank_payments(id,payment_date,invoice_number,instructed_amount,applied_amount,currency)
                    values(15,'2026-10-20','',7780,0,'EUR');""")
        other = {**snapshot, "payment": json.loads(self.sql("select to_jsonb(p) from bank_payments p where id=15;")),
                 "allocations": []}
        payload = json.dumps(other).replace("'", "''")

        def attempt(which):
            try:
                if which == 14:
                    return self.apply(snapshot)["status"]
                self.sql(f"set role service_role; select public.add_invoice_to_existing_payment(15,3087,'{payload}'::jsonb);")
                return "added"
            except RuntimeError as exc:
                self.assertTrue("duplicate settlement" in str(exc) or "busy" in str(exc))
                return "blocked"

        with ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(attempt, (14, 15)))
        self.assertEqual(sorted(results), ["added", "blocked"])
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations where invoice_row_id=3087;"), "1")

    def test_status_validation(self):
        for status in ("Yes", "cancelled", "unknown"):
            with self.subTest(status=status):
                self.sql(f"update invoices set paid='{status}' where id=3087;")
                before = self.ledger()
                with self.assertRaisesRegex(RuntimeError, "not unpaid"):
                    self.apply()
                self.assertEqual(self.ledger(), before)

    def test_missing_invoice(self):
        snapshot = self.snapshot()
        self.sql("delete from invoices where id=3087;")
        with self.assertRaisesRegex(RuntimeError, "no longer exists"):
            self.apply(snapshot)

    def test_over_allocation_and_invalid_instructed_amount(self):
        for value in ("46679", "null", "0", "'NaN'"):
            with self.subTest(value=value):
                self.sql(f"update bank_payments set instructed_amount={value} where id=14;")
                before = self.ledger()
                with self.assertRaises(RuntimeError):
                    self.apply()
                self.assertEqual(self.ledger(), before)

    def test_invalid_date_currency_and_summary(self):
        updates = [
            "update bank_payments set payment_date=null where id=14",
            "update invoices set currency='USD' where id=3087",
            "update bank_payments set applied_amount=1 where id=14",
            "update bank_payments set invoice_number='8596' where id=14",
        ]
        for update in updates:
            with self.subTest(update=update):
                self.setUp()
                self.sql(update)
                before = self.ledger()
                with self.assertRaises(RuntimeError):
                    self.apply()
                self.assertEqual(self.ledger(), before)

    def test_stale_preview(self):
        snapshot = self.snapshot()
        self.sql("update invoices set payment_amount=7779 where id=3087;")
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "stale"):
            self.apply(snapshot)
        self.assertEqual(self.ledger(), before)

    def test_stale_payment_preview(self):
        snapshot = self.snapshot()
        self.sql("update bank_payments set notes='changed since preview' where id=14;")
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "stale"):
            self.apply(snapshot)
        self.assertEqual(self.ledger(), before)

    def test_other_payment_allocation_blocks_duplicate_settlement(self):
        self.sql("""
            insert into bank_payments(id,payment_date,invoice_number,instructed_amount,applied_amount,currency)
              values(15,'2026-10-20','8707',7780,7780,'EUR');
            insert into bank_payment_allocations(payment_id,invoice_row_id,invoice_number,project_name,amount_applied)
              values(15,3087,8707,'Intermarche Heusy',7780);
        """)
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "duplicate settlement"):
            self.apply()
        self.assertEqual(self.ledger(), before)

    def test_invalid_invoice_amount_and_number(self):
        for field, value in (
            ("invoice_number", "null"), ("payment_amount", "0"),
            ("payment_amount", "-1"), ("payment_amount", "'NaN'"),
            ("payment_amount", "7780.001"),
        ):
            with self.subTest(field=field, value=value):
                self.setUp()
                self.sql(f"update invoices set {field}={value} where id=3087;")
                before = self.ledger()
                with self.assertRaisesRegex(RuntimeError, "Positive full existing"):
                    self.apply()
                self.assertEqual(self.ledger(), before)

    def test_full_rollback_on_injected_failure(self):
        snapshot = self.snapshot()
        before = self.ledger()
        self.sql("""
            create or replace function public.fail_allocation_test() returns trigger language plpgsql as $$
            begin raise exception 'injected summary failure'; end $$;
            create trigger fail_summary before update on bank_payments
              for each row execute function public.fail_allocation_test();
        """)
        with self.assertRaisesRegex(RuntimeError, "injected summary failure"):
            self.apply(snapshot)
        self.assertEqual(self.ledger(), before)

    def test_non_admin_backend_roles_cannot_execute(self):
        for role in ("anon", "authenticated"):
            with self.subTest(role=role), self.assertRaisesRegex(RuntimeError, "permission denied"):
                self.apply(role=role)

    def test_financial_writer_guard_blocks_before_mutation(self):
        self.sql("""
            create or replace function public.allocation_test_guard() returns trigger language plpgsql as $$
            begin return new; end $$;
            create trigger bank_payments_writer_guard before update on bank_payments
              for each row execute function public.allocation_test_guard();
        """)
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "financial-writer guards"):
            self.apply()
        self.assertEqual(self.ledger(), before)

    def test_legacy_eur_schema_without_currency_column(self):
        self.sql("alter table invoices drop column currency;")
        try:
            self.assertEqual(self.apply()["status"], "added")
        finally:
            self.sql("alter table invoices add column currency text default 'EUR';")

    def connection(self, name="allocation-test"):
        if psycopg is None:
            self.skipTest("psycopg is required for deterministic mixed-writer tests")
        return psycopg.connect(
            host="127.0.0.1", port=self.port, dbname="postgres", user="postgres",
            password="isolated-test-only", application_name=name, connect_timeout=5,
            options="-c statement_timeout=8000 -c lock_timeout=2000",
        )

    def add_on(self, connection, snapshot, payment_id=14):
        with connection.cursor() as cursor:
            cursor.execute("set local role service_role")
            cursor.execute(
                "select public.add_invoice_to_existing_payment(%s,3087,%s::jsonb)",
                (payment_id, json.dumps(snapshot)),
            )
            return cursor.fetchone()[0]

    def credit_payload(self, row_id=3087, payment_id=15):
        number, amount, project = ("8707", 7780, "Intermarche Heusy") if row_id == 3087 else ("9999", 100, "Unrelated")
        return {"ops": [
            {"op": "target_update", "id": row_id, "expected_amount": amount,
             "fields": {"paid": "Yes", "payment_date": "2026-10-20"}},
            {"op": "save_payment", "entry": {
                "payment_fingerprint": "original-fingerprint" if payment_id == 14 else "credit-other",
                "payment_date": "2026-10-20", "invoice_number": number,
                "source_kind": "credit-apply", "source_name": f"credit-apply-{number}",
                "currency": "EUR", "applied_amount": amount,
            }, "allocations": [{
                "invoice_row_id": row_id, "invoice_number": number, "project_name": project,
                "maintenance_year": "Y1", "year": 2026, "amount_applied": amount,
            }]},
        ]}

    def credit_on(self, connection, payload):
        with connection.cursor() as cursor:
            cursor.execute("set local role service_role")
            cursor.execute("select public.apply_credit_atomic(%s::jsonb)", (json.dumps(payload),))
            return cursor.fetchone()[0]

    def wait_for_lock_wait(self, name):
        deadline = time.monotonic() + 4
        with self.connection("mixed-test-observer") as observer:
            observer.autocommit = True
            while time.monotonic() < deadline:
                row = observer.execute(
                    "select exists(select 1 from pg_stat_activity "
                    "where application_name=%s and wait_event_type='Lock')", (name,),
                ).fetchone()
                if row[0]:
                    return
                time.sleep(0.02)
        self.fail(f"Writer {name} did not reach the expected lock wait")

    def test_mixed_credit_first_same_invoice_fails_fast_then_refresh_blocks_duplicate(self):
        snapshot = self.snapshot()
        with self.connection() as credit:
            self.credit_on(credit, self.credit_payload())
            before = self.ledger()
            started = time.monotonic()
            with self.connection() as manual:
                with self.assertRaises(psycopg.errors.LockNotAvailable) as raised:
                    self.add_on(manual, snapshot)
                self.assertIn("fresh preview", str(raised.exception))
                manual.rollback()
            self.assertLess(time.monotonic() - started, 2)
            self.assertEqual(self.ledger(), before)
            credit.commit()
        with self.assertRaisesRegex(RuntimeError, "duplicate settlement"):
            self.apply(snapshot)
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations where invoice_row_id=3087;"), "1")

    def test_mixed_manual_first_same_invoice_credit_waits_then_rolls_back(self):
        snapshot = self.snapshot()
        name = f"credit-wait-{uuid.uuid4().hex[:8]}"

        def credit_worker():
            with self.connection(name) as connection:
                try:
                    self.credit_on(connection, self.credit_payload())
                    connection.commit()
                    return "unexpected success"
                except psycopg.Error as exc:
                    connection.rollback()
                    return exc.sqlstate, str(exc)

        with self.connection() as manual, ThreadPoolExecutor(max_workers=1) as pool:
            self.add_on(manual, snapshot)
            future = pool.submit(credit_worker)
            self.wait_for_lock_wait(name)
            manual.commit()
            state, message = future.result(timeout=6)
        self.assertEqual(state, "P0001")
        self.assertIn("no longer open", message)
        self.assertEqual(self.sql("select count(*) from bank_payments;"), "1")
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations where invoice_row_id=3087;"), "1")

    def test_mixed_same_payment_credit_first_busy_then_stale_no_overwrite(self):
        snapshot = self.snapshot()
        with self.connection() as credit:
            self.credit_on(credit, self.credit_payload(row_id=999, payment_id=14))
            with self.connection() as manual:
                with self.assertRaises(psycopg.errors.LockNotAvailable):
                    self.add_on(manual, snapshot)
                manual.rollback()
            credit.commit()
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "stale"):
            self.apply(snapshot)
        self.assertEqual(self.ledger(), before)

    def test_mixed_same_payment_manual_first_blocks_legacy_replace_after_commit(self):
        snapshot = self.snapshot()
        name = f"legacy-replace-{uuid.uuid4().hex[:8]}"

        def credit_worker():
            with self.connection(name) as connection:
                try:
                    self.credit_on(connection, self.credit_payload(row_id=999, payment_id=14))
                    connection.commit()
                    return "unexpected success"
                except psycopg.Error as exc:
                    connection.rollback()
                    return str(exc)

        with self.connection() as manual, ThreadPoolExecutor(max_workers=1) as pool:
            self.add_on(manual, snapshot)
            future = pool.submit(credit_worker)
            self.wait_for_lock_wait(name)
            manual.commit()
            message = future.result(timeout=6)
        self.assertIn("protected", message)
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations where invoice_row_id=3087;"), "1")
        self.assertEqual(self.sql("select paid from invoices where id=3087;"), "Yes")
        self.assertEqual(self.sql("select paid from invoices where id=999;"), "No")

    def test_mixed_unrelated_writers_complete_in_both_start_orders(self):
        for first in ("manual", "credit"):
            with self.subTest(first=first):
                self.setUp()
                snapshot = self.snapshot()
                with self.connection() as one, self.connection() as two:
                    if first == "manual":
                        self.add_on(one, snapshot)
                        started = time.monotonic()
                        self.credit_on(two, self.credit_payload(row_id=999))
                    else:
                        self.credit_on(one, self.credit_payload(row_id=999))
                        started = time.monotonic()
                        self.add_on(two, snapshot)
                    self.assertLess(time.monotonic() - started, 2)
                    two.commit()
                    one.commit()
                self.assertEqual(self.sql("select count(*) from bank_payments;"), "2")
                self.assertEqual(self.sql("select count(*) from bank_payment_allocations;"), "6")

    def test_opposite_order_payment_holder_does_not_create_a_deadlock(self):
        snapshot = self.snapshot()
        with self.connection() as legacy, self.connection() as manual:
            legacy.execute("select id from bank_payments where id=14 for update")
            with self.assertRaises(psycopg.errors.LockNotAvailable):
                self.add_on(manual, snapshot)
            manual.rollback()
            # Failed manual call released its invoice lock; payment-first writer can finish.
            legacy.execute("update invoices set description='legacy writer' where id=3087")
            legacy.rollback()
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations;"), "4")

    def test_allocation_row_contention_is_scoped_and_fails_fast(self):
        snapshot = self.snapshot()
        before = self.ledger()
        with self.connection() as legacy, self.connection() as manual:
            legacy.execute("select id from bank_payment_allocations where payment_id=14 for update")
            with self.assertRaises(psycopg.errors.LockNotAvailable):
                self.add_on(manual, snapshot)
            manual.rollback()
            legacy.rollback()
        self.assertEqual(self.ledger(), before)

    def test_mixed_credit_timeout_rolls_back_earlier_operations(self):
        snapshot = self.snapshot()
        payload = self.credit_payload()
        payload["ops"].insert(0, {
            "op": "target_update", "id": 999, "expected_amount": 100,
            "fields": {"paid": "Yes", "payment_date": "2026-10-20"},
        })
        before = self.ledger()
        with self.connection() as manual, self.connection() as credit:
            self.add_on(manual, snapshot)
            credit.execute("set local lock_timeout='200ms'")
            with self.assertRaises(psycopg.errors.LockNotAvailable):
                self.credit_on(credit, payload)
            credit.rollback()
            manual.rollback()
        self.assertEqual(self.ledger(), before)

    def test_mixed_credit_injected_failure_rolls_back_manual_can_then_succeed(self):
        self.sql("""
            create or replace function public.fail_credit_test() returns trigger language plpgsql as $$
            begin if new.source_kind='credit-apply' then raise exception 'injected credit failure'; end if;
            return new; end $$;
            create trigger fail_credit before insert on bank_payments
              for each row execute function public.fail_credit_test();
        """)
        before = self.ledger()
        try:
            with self.connection() as credit:
                with self.assertRaisesRegex(psycopg.Error, "injected credit failure"):
                    self.credit_on(credit, self.credit_payload())
                credit.rollback()
            self.assertEqual(self.ledger(), before)
            self.assertEqual(self.apply()["status"], "added")
        finally:
            self.sql("drop trigger fail_credit on bank_payments;")

    def test_implicit_ddl_lock_timeout_fully_rolls_back(self):
        snapshot = self.snapshot()
        before = self.ledger()
        self.assertEqual(self.sql("""
            select proconfig @> array['lock_timeout=5s']
            from pg_proc where oid='public.add_invoice_to_existing_payment(bigint,bigint,jsonb)'::regprocedure;
        """), "t")
        with self.connection() as ddl, self.connection() as manual:
            ddl.execute("lock table bank_payment_allocations in access exclusive mode")
            started = manual.execute("select clock_timestamp()").fetchone()[0]
            with self.assertRaises(psycopg.errors.LockNotAvailable):
                self.add_on(manual, snapshot)
            manual.rollback()
            # Measure PostgreSQL's timeout on its clock, not the Windows/Docker clock boundary.
            finished = manual.execute("select clock_timestamp()").fetchone()[0]
            elapsed = (finished - started).total_seconds()
            self.assertGreaterEqual(elapsed, 4.5)
            self.assertLess(elapsed, 7)
            manual.rollback()
            ddl.rollback()
        self.assertEqual(self.ledger(), before)

    def test_raw_allocation_insert_is_forced_into_invoice_lock_protocol(self):
        snapshot = self.snapshot()
        self.sql("""insert into bank_payments(id,payment_date,invoice_number,instructed_amount,applied_amount,currency)
                    values(15,'2026-10-20','',7780,0,'EUR');""")
        with self.connection() as raw, self.connection() as manual:
            self.add_on(manual, snapshot)
            with self.assertRaises(psycopg.errors.LockNotAvailable):
                raw.execute("""
                    insert into bank_payment_allocations
                      (payment_id,invoice_row_id,invoice_number,project_name,amount_applied)
                    values(15,3087,8707,'Intermarche Heusy',7780)
                """)
            raw.rollback()
            manual.commit()
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations where invoice_row_id=3087;"), "1")

    def test_no_broad_table_or_unused_advisory_lock(self):
        snapshot = self.snapshot()
        with self.connection() as manual:
            self.add_on(manual, snapshot)
            pid = manual.info.backend_pid
            with self.connection() as observer:
                rows = observer.execute(
                    "select mode from pg_locks where pid=%s "
                    "and relation='public.bank_payment_allocations'::regclass and granted", (pid,),
                ).fetchall()
                self.assertNotIn(("ShareRowExclusiveLock",), rows)
                self.assertEqual(observer.execute(
                    "select count(*) from pg_locks where pid=%s and locktype='advisory'", (pid,),
                ).fetchone()[0], 0)
            manual.rollback()

    def seed_other_payment(self):
        self.sql("""insert into bank_payments
          (id,payment_date,invoice_number,instructed_amount,applied_amount,currency)
          values(15,'2026-10-20','',7780,0,'EUR');""")

    def raw_insert(self, connection, row_id=3087, payment_id=15):
        connection.execute("set local role service_role")
        connection.execute("""
            insert into bank_payment_allocations
              (payment_id,invoice_row_id,invoice_number,project_name,amount_applied)
            select %s,id,invoice_number::bigint,project_name,payment_amount from invoices where id=%s
        """, (payment_id, row_id))

    def reverse(self, result, snapshot=None, role="service_role", confirmed=True, actor="test-admin", reason="Wrong row"):
        snapshot = self.snapshot() if snapshot is None else snapshot
        snapshot = {**snapshot, "actor": snapshot.get("actor", "test-admin")}
        payload = json.dumps(snapshot).replace("'", "''")
        return json.loads(self.sql(
            f"set role {role}; select public.reverse_manual_payment_allocation("
            f"{result['allocation_id']},'{payload}'::jsonb,'{actor}','{reason}',"
            f"{'true' if confirmed else 'false'});"
        ).splitlines()[-1])

    def test_legacy_replacement_after_committed_manual_allocation_fully_rolls_back(self):
        self.apply()
        before = self.ledger()
        with self.connection() as credit:
            with self.assertRaisesRegex(psycopg.Error, "protected"):
                self.credit_on(credit, self.credit_payload(row_id=999, payment_id=14))
            credit.rollback()
        self.assertEqual(self.ledger(), before)

    def test_raw_insert_before_manual_commit_and_rollback_recovery(self):
        self.seed_other_payment()
        snapshot = self.snapshot()
        with self.connection() as raw, self.connection() as manual:
            self.raw_insert(raw)
            with self.assertRaises(psycopg.errors.LockNotAvailable):
                self.add_on(manual, snapshot)
            manual.rollback()
            raw.rollback()
        self.assertEqual(self.apply(snapshot)["status"], "added")

    def test_raw_insert_committed_before_manual_blocks_duplicate(self):
        self.seed_other_payment()
        snapshot = self.snapshot()
        with self.connection() as raw:
            self.raw_insert(raw)
            raw.commit()
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "duplicate settlement"):
            self.apply(snapshot)
        self.assertEqual(self.ledger(), before)

    def test_raw_insert_after_manual_blocked_by_native_unique_constraint(self):
        self.seed_other_payment()
        self.apply()
        before = self.ledger()
        with self.connection() as raw:
            with self.assertRaises(psycopg.errors.UniqueViolation) as raised:
                self.raw_insert(raw)
            self.assertEqual(raised.exception.diag.constraint_name, "legacy_invoice_settlement_once")
            raw.rollback()
        self.assertEqual(self.ledger(), before)

    def test_repeatable_read_old_raw_snapshot_cannot_duplicate_committed_manual(self):
        self.seed_other_payment()
        snapshot = self.snapshot()
        with self.connection() as raw:
            raw.execute("set transaction isolation level repeatable read")
            raw.execute("select count(*) from bank_payment_allocations")
            self.apply(snapshot)
            with self.assertRaises((psycopg.errors.SerializationFailure, psycopg.errors.UniqueViolation)):
                self.raw_insert(raw)
            raw.rollback()
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations where invoice_row_id=3087;"), "1")

    def test_two_raw_payments_cannot_duplicate_same_invoice(self):
        self.seed_other_payment()
        with self.connection() as one, self.connection() as two:
            self.raw_insert(one)
            with self.assertRaises(psycopg.errors.LockNotAvailable):
                self.raw_insert(two, payment_id=14)
            two.rollback()
            one.commit()
            with self.assertRaises(psycopg.errors.UniqueViolation):
                self.raw_insert(two, payment_id=14)
            two.rollback()
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations where invoice_row_id=3087;"), "1")

    def test_unrelated_raw_insert_succeeds_while_manual_is_uncommitted(self):
        self.seed_other_payment()
        snapshot = self.snapshot()
        with self.connection() as manual, self.connection() as raw:
            self.add_on(manual, snapshot)
            self.raw_insert(raw, row_id=999)
            raw.commit()
            manual.commit()
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations;"), "6")

    def test_protected_allocation_deletion_mutation_and_payment_cascade_are_blocked(self):
        self.apply()
        before = self.ledger()
        for statement in (
            "delete from bank_payment_allocations where payment_id=14",
            "delete from bank_payment_allocations where invoice_row_id=500",
            "update bank_payment_allocations set manual_invoice_before=null where invoice_row_id=3087",
            "update bank_payment_allocations set amount_applied=1 where invoice_row_id=500",
            "delete from bank_payments where id=14",
        ):
            with self.subTest(statement=statement), self.assertRaisesRegex(RuntimeError, "protected|cannot be deleted"):
                self.sql("set role service_role; " + statement)
            self.assertEqual(self.ledger(), before)

    def test_protected_payment_finance_pdf_and_stale_summary_updates_blocked(self):
        self.apply()
        before = self.ledger()
        for fields in ("applied_amount=0", "instructed_amount=1", "invoice_number='8707'",
                       "pdf_storage_path='changed.pdf'", "payment_fingerprint='changed'", "received_amount=1"):
            with self.subTest(fields=fields), self.assertRaisesRegex(RuntimeError, "protected"):
                self.sql(f"set role service_role; update bank_payments set {fields} where id=14;")
            self.assertEqual(self.ledger(), before)

    def test_protected_invoice_unpay_cancel_amount_and_identity_changes_blocked(self):
        self.apply()
        before = self.ledger()
        for fields in ("paid='No'", "paid='cancelled'", "payment_amount=1",
                       "payment_date='2026-10-21'", "invoice_number='9990'"):
            with self.subTest(fields=fields), self.assertRaisesRegex(RuntimeError, "protected"):
                self.sql(f"set role service_role; update invoices set {fields} where id=3087;")
            self.assertEqual(self.ledger(), before)
        self.sql("set role service_role; update invoices set description='bookkeeping allowed' where id=3087;")

    def test_ordinary_unprotected_replacement_and_deletion_still_work(self):
        with self.connection() as credit:
            self.credit_on(credit, self.credit_payload(row_id=999, payment_id=14))
            credit.commit()
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations where payment_id=14;"), "1")
        self.sql("set role service_role; delete from bank_payments where id=14; delete from invoices where id=999;")
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations;"), "0")

    def test_authorized_reversal_audits_preserves_and_allows_recovery(self):
        original = self.ledger()
        result = self.apply()
        reversed_result = self.reverse(result)
        self.assertEqual(reversed_result["status"], "reversed")
        after = self.ledger()
        self.assertEqual(after["allocations"], original["allocations"])
        self.assertEqual(after["invoices"], original["invoices"])
        for key, value in original["payments"][0].items():
            if key != "updated_at":
                self.assertEqual(after["payments"][0][key], value)
        audit = json.loads(self.sql("select to_jsonb(r) from legacy_allocation_reversals r;"))
        self.assertEqual(audit["actor"], "test-admin")
        self.assertEqual(audit["reason"], "Wrong row")
        self.assertEqual(audit["allocation"]["id"], result["allocation_id"])
        self.assertEqual(self.apply()["status"], "added")

    def test_reversal_requires_admin_confirmation_reason_and_fresh_snapshot(self):
        result = self.apply()
        before = self.ledger()
        for options in ({"role": "anon"}, {"role": "authenticated"}, {"confirmed": False},
                        {"actor": ""}, {"actor": "wrong-admin"}, {"reason": ""}):
            with self.subTest(options=options), self.assertRaises(RuntimeError):
                self.reverse(result, **options)
            self.assertEqual(self.ledger(), before)
        stale = self.snapshot()
        self.sql("update invoices set description='changed' where id=3087;")
        with self.assertRaisesRegex(RuntimeError, "stale"):
            self.reverse(result, snapshot=stale)
        self.assertEqual(self.sql("select count(*) from legacy_allocation_reversals;"), "0")

    def test_correction_owner_cannot_be_assumed_or_audit_changed_by_backend(self):
        self.assertEqual(self.sql(
            "select pg_has_role('service_role','caddycheck_legacy_correction','MEMBER');"), "f")
        with self.assertRaises(RuntimeError):
            self.sql("set session authorization service_role; set role caddycheck_legacy_correction;")
        with self.assertRaises(RuntimeError):
            self.sql("set role service_role; insert into legacy_allocation_reversals(allocation_id) values(1);")
        self.apply()
        with self.assertRaises(RuntimeError):
            self.sql("set role service_role; truncate invoices,bank_payments,bank_payment_allocations cascade;")

    def test_caller_guc_cannot_bypass_protection(self):
        self.apply()
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "protected"):
            self.sql("set role service_role; set caddycheck.allow_correction='on'; "
                     "delete from bank_payment_allocations where payment_id=14;")
        self.assertEqual(self.ledger(), before)

    def test_reversal_failure_in_audit_rolls_back_everything_then_recovers(self):
        result = self.apply()
        before = self.ledger()
        self.sql("""
            create function public.fail_reversal_audit() returns trigger language plpgsql as $$
            begin raise exception 'injected reversal audit failure'; end $$;
            create trigger fail_reversal before insert on legacy_allocation_reversals
              for each row execute function public.fail_reversal_audit();
        """)
        try:
            with self.assertRaisesRegex(RuntimeError, "injected reversal"):
                self.reverse(result)
            self.assertEqual(self.ledger(), before)
            self.assertEqual(self.sql("select count(*) from legacy_allocation_reversals;"), "0")
        finally:
            self.sql("drop trigger fail_reversal on legacy_allocation_reversals;")
        self.assertEqual(self.reverse(result)["status"], "reversed")

    def test_concurrent_reversal_only_one_commits_and_audit_is_once(self):
        result = self.apply()
        snapshot = self.snapshot()
        def attempt(_):
            try:
                return self.reverse(result, snapshot=snapshot)["status"]
            except RuntimeError:
                return "rejected"
        with ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(attempt, range(2)))
        self.assertEqual(sorted(results), ["rejected", "reversed"])
        self.assertEqual(self.sql("select count(*) from legacy_allocation_reversals;"), "1")

    def test_reversal_financial_writer_guard_is_not_bypassed(self):
        result = self.apply()
        self.sql("""
            create or replace function public.test_writer_guard() returns trigger language plpgsql as $$
            begin raise exception 'writer guard'; end $$;
            create trigger bank_payments_writer_guard before update on bank_payments
              for each row execute function public.test_writer_guard();
        """)
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "financial-writer"):
            self.reverse(result)
        self.assertEqual(self.ledger(), before)

    def test_missing_coordinated_guard_disables_manual_feature(self):
        self.sql("alter table bank_payment_allocations disable trigger legacy_allocation_guard;")
        try:
            with self.assertRaisesRegex(RuntimeError, "protections must be installed"):
                self.apply()
        finally:
            self.sql("alter table bank_payment_allocations enable trigger legacy_allocation_guard;")

    def test_protected_payment_can_attach_only_its_missing_pdf_once(self):
        self.sql("update bank_payments set pdf_storage_bucket=null,pdf_storage_path=null where id=14;")
        self.apply()
        self.sql("""set role service_role; update bank_payments
                    set pdf_storage_bucket='bank-payments',pdf_storage_path='first.pdf' where id=14;""")
        before = self.ledger()
        with self.assertRaisesRegex(RuntimeError, "protected"):
            self.sql("set role service_role; update bank_payments set pdf_storage_path='replacement.pdf' where id=14;")
        self.assertEqual(self.ledger(), before)

    def test_production_like_rls_and_correction_role_flags(self):
        self.assertEqual(self.sql("""
            select count(*) from pg_class where oid in (
              'public.invoices'::regclass,'public.bank_payments'::regclass,
              'public.bank_payment_allocations'::regclass,'public.invoice_change_log'::regclass)
              and relrowsecurity and not relforcerowsecurity;
        """), "4")
        self.assertEqual(self.sql("""
            select count(*) from pg_roles where rolname in
              ('caddycheck_legacy_correction','caddycheck_legacy_scope')
              and not rolcanlogin and not rolbypassrls and not rolsuper and not rolinherit;
        """), "2")
        self.assertEqual(self.sql("""
            select count(*) from pg_auth_members
              where member in (select oid from pg_roles where rolname in
                  ('caddycheck_legacy_correction','caddycheck_legacy_scope'))
                 or roleid in (select oid from pg_roles where rolname in
                  ('caddycheck_legacy_correction','caddycheck_legacy_scope'));
        """), "0")
        self.assertEqual(self.sql("""
            select prosecdef from pg_proc where oid='public.invoice_audit_trigger()'::regprocedure;
        """), "f")
        self.assertEqual(self.sql("""
            select md5(pg_get_functiondef('public.invoice_audit_trigger()'::regprocedure));
        """), "bbd826dd9f5c9b8db344238136bcfefb")

    def test_correction_has_no_rows_or_write_access_outside_active_scope(self):
        self.apply()
        before = self.ledger()
        self.assertEqual(self.sql("""
            set role caddycheck_legacy_correction;
            select (select count(*) from invoices),
                   (select count(*) from bank_payments),(select count(*) from bank_payment_allocations);
            update invoices set paid='No' where id=999;
            update bank_payments set applied_amount=0 where id=14;
            delete from bank_payment_allocations;
        """).splitlines(), ["SET", "0|0|0", "UPDATE 0", "UPDATE 0", "DELETE 0"])
        with self.assertRaises(RuntimeError):
            self.sql("set role caddycheck_legacy_correction; select * from invoice_change_log;")
        with self.assertRaises(RuntimeError):
            self.sql("set role caddycheck_legacy_correction; select * from legacy_correction_private.active_scope;")
        self.assertEqual(self.ledger(), before)

    def test_correction_scope_creator_has_no_financial_privileges(self):
        for table in ("invoices", "bank_payments", "bank_payment_allocations", "invoice_change_log"):
            self.assertEqual(self.sql(
                f"select has_table_privilege('caddycheck_legacy_scope','public.{table}','SELECT,INSERT,UPDATE,DELETE');"
            ), "f")
        for role in ("caddycheck_legacy_scope", "caddycheck_legacy_correction"):
            for table in ("invoices", "bank_payments", "bank_payment_allocations", "invoice_change_log"):
                self.assertEqual(self.sql(
                    f"select has_table_privilege('{role}','public.{table}','TRUNCATE,TRIGGER');"
                ), "f")

    def test_no_caller_guc_or_direct_inner_rpc_can_open_a_correction_scope(self):
        result = self.apply()
        before = self.ledger()
        for role in ("service_role", "anon", "authenticated"):
            with self.subTest(role=role), self.assertRaises(RuntimeError):
                self.sql(f"set role {role}; "
                         "select legacy_correction_private.reverse_scoped(1,'{}','actor','reason',true);")
        self.assertEqual(self.sql("""
            set role caddycheck_legacy_correction;
            set caddycheck.payment_id='14';
            set caddycheck.invoice_row_id='3087';
            set caddycheck.actor='test-admin';
            select count(*) from bank_payment_allocations;
        """).splitlines()[-1], "0")
        payload = json.dumps({**self.snapshot(), "actor": "test-admin"}).replace("'", "''")
        with self.assertRaisesRegex(RuntimeError, "missing"):
            self.sql(f"set role caddycheck_legacy_correction; "
                     f"select legacy_correction_private.reverse_scoped({result['allocation_id']},"
                     f"'{payload}'::jsonb,'test-admin','Wrong row',true);")
        self.assertEqual(self.ledger(), before)

    def test_reversal_writes_existing_invoker_invoice_audit_with_reviewed_actor(self):
        result = self.apply()
        before = self.snapshot()
        self.reverse(result)
        audit = json.loads(self.sql("""
            select to_jsonb(l) from invoice_change_log l
            where source='manual-allocation-reversal' order by id desc limit 1;
        """))
        self.assertEqual(audit["action"], "payment_status_change")
        self.assertEqual(audit["invoice_row_id"], 3087)
        self.assertEqual(audit["changed_by"], "test-admin")
        self.assertEqual(audit["reason"], "Wrong row")
        self.assertEqual(audit["old_values"], before["invoice"])
        self.assertEqual(audit["new_values"], self.snapshot()["invoice"])
        self.assertEqual(self.sql("select count(*) from legacy_correction_private.active_scope;"), "0")

    def test_existing_invoice_audit_failure_rolls_back_and_scope_is_removed(self):
        result = self.apply()
        before = self.ledger()
        logs_before = self.sql("select count(*) from invoice_change_log;")
        self.sql("""
            create function public.test_fail_existing_audit() returns trigger language plpgsql as $$
            begin
              if new.source='manual-allocation-reversal' then raise exception 'injected invoice audit failure'; end if;
              return new;
            end $$;
            create trigger fail_existing_audit before insert on invoice_change_log
              for each row execute function public.test_fail_existing_audit();
        """)
        try:
            with self.assertRaisesRegex(RuntimeError, "injected invoice audit failure"):
                self.reverse(result)
            self.assertEqual(self.ledger(), before)
            self.assertEqual(self.sql("select count(*) from invoice_change_log;"), logs_before)
            self.assertEqual(self.sql("select count(*) from legacy_allocation_reversals;"), "0")
            self.assertEqual(self.sql("select count(*) from legacy_correction_private.active_scope;"), "0")
        finally:
            self.sql("drop trigger fail_existing_audit on invoice_change_log;")
        self.assertEqual(self.reverse(result)["status"], "reversed")

    def test_missing_audit_insert_policy_rolls_back_without_success_fallback(self):
        result = self.apply()
        before = self.ledger()
        self.sql("drop policy legacy_correction_invoice_audit_insert on invoice_change_log;")
        try:
            with self.assertRaisesRegex(RuntimeError, "row-level security"):
                self.reverse(result)
            self.assertEqual(self.ledger(), before)
            self.assertEqual(self.sql("select count(*) from legacy_allocation_reversals;"), "0")
            self.assertEqual(self.sql("select count(*) from legacy_correction_private.active_scope;"), "0")
        finally:
            self.sql("""
                create policy legacy_correction_invoice_audit_insert on invoice_change_log
                for insert to caddycheck_legacy_correction with check(
                    legacy_correction_private.permits('audit_insert',invoice_row_id,to_jsonb(invoice_change_log)));
            """)
        self.assertEqual(self.reverse(result)["status"], "reversed")

    def test_identity_audit_id_requires_no_sequence_privilege(self):
        sequence = self.sql("select pg_get_serial_sequence('public.invoice_change_log','id');")
        self.assertEqual(self.sql(
            f"select has_sequence_privilege('caddycheck_legacy_correction','{sequence}','USAGE'),"
            f"has_sequence_privilege('caddycheck_legacy_correction','{sequence}','SELECT'),"
            f"has_sequence_privilege('caddycheck_legacy_correction','{sequence}','UPDATE');"
        ), "f|f|f")
        self.assertEqual(self.sql("""
            select count(*) from pg_class c where relkind='S'
              and c.oid <> pg_get_serial_sequence('public.invoice_change_log','id')::regclass
              and case when c.relkind='S'
                  then has_sequence_privilege('caddycheck_legacy_correction',c.oid,'USAGE,SELECT,UPDATE')
                  else false end;
        """), "0")
        result = self.apply()
        self.assertEqual(self.reverse(result)["status"], "reversed")

    def test_rls_fences_block_unrelated_writes_even_during_active_scope(self):
        self.seed_other_payment()
        result = self.apply()
        self.sql("""
            create function public.test_inspect_correction_scope() returns trigger language plpgsql as $$
            declare n integer;
            begin
                if current_user='caddycheck_legacy_correction' then
                    if (select count(*) from invoices)<>1 or (select count(*) from bank_payments)<>1
                       or (select count(*) from bank_payment_allocations)<>4 then
                        raise exception 'scope leaked unrelated rows';
                    end if;
                    update invoices set paid='Yes' where id=999;
                    get diagnostics n=row_count;
                    if n<>0 then raise exception 'unrelated invoice updated'; end if;
                    update bank_payments set applied_amount=0 where id=15;
                    get diagnostics n=row_count;
                    if n<>0 then raise exception 'unrelated payment updated'; end if;
                    delete from bank_payment_allocations where invoice_row_id=500;
                    get diagnostics n=row_count;
                    if n<>0 then raise exception 'unrelated allocation deleted'; end if;
                    begin
                        update bank_payment_allocations set id=id where invoice_row_id=500;
                        raise exception 'allocation mutation unexpectedly permitted';
                    exception when insufficient_privilege then null;
                    end;
                    begin
                        insert into invoice_change_log(action,invoice_row_id,changed_by,source,reason)
                          values('payment_status_change',999,'test-admin','manual-allocation-reversal','Wrong row');
                        raise exception 'unrelated audit insert unexpectedly permitted';
                    exception when insufficient_privilege then null;
                    end;
                end if;
                return new;
            end $$;
            create trigger inspect_scope before update on invoices
              for each row execute function public.test_inspect_correction_scope();
            create policy test_public_invoices on invoices for all to public using(true) with check(true);
            create policy test_public_payments on bank_payments for all to public using(true) with check(true);
            create policy test_public_allocations on bank_payment_allocations for all to public using(true) with check(true);
            create policy test_public_audit on invoice_change_log for all to public using(true) with check(true);
        """)
        try:
            self.assertEqual(self.reverse(result)["status"], "reversed")
            self.assertEqual(self.sql("select paid from invoices where id=999;"), "No")
            self.assertEqual(self.sql("select applied_amount from bank_payments where id=15;"), "0.00")
            self.assertEqual(self.sql("select count(*) from bank_payment_allocations where invoice_row_id=500;"), "1")
        finally:
            self.sql("""
                drop trigger inspect_scope on invoices;
                drop policy test_public_invoices on invoices;
                drop policy test_public_payments on bank_payments;
                drop policy test_public_allocations on bank_payment_allocations;
                drop policy test_public_audit on invoice_change_log;
            """)

    def test_scope_and_audit_settings_do_not_survive_successful_rpc(self):
        result = self.apply()
        payload = json.dumps({**self.snapshot(), "actor": "test-admin"})
        with self.connection() as connection:
            connection.execute("set local role service_role")
            connection.execute("set local caddycheck.actor='prior-actor'")
            connection.execute("set local caddycheck.reason='prior-reason'")
            connection.execute("set local caddycheck.source='prior-source'")
            connection.execute(
                "select public.reverse_manual_payment_allocation(%s,%s::jsonb,'test-admin','Wrong row',true)",
                (result["allocation_id"], payload),
            )
            self.assertEqual(connection.execute(
                "select current_setting('caddycheck.actor'),current_setting('caddycheck.reason'),"
                "current_setting('caddycheck.source')"
            ).fetchone(), ("prior-actor", "prior-reason", "prior-source"))
            connection.execute("set local role caddycheck_legacy_correction")
            self.assertEqual(connection.execute("select count(*) from invoices").fetchone()[0], 0)
            connection.rollback()
