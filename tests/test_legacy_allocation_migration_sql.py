"""Migration preconditions in an owned disposable PostgreSQL; no app credentials."""
from pathlib import Path
import subprocess
import time
import unittest
import uuid

from tests.test_apply_credit_atomic_sql import SCHEMA, _docker_ready, install_production_like_rls
from tests import test_existing_payment_allocation_sql as harness

MIGRATION = Path(__file__).resolve().parents[1] / "migrations" / "protect_legacy_payment_allocations.sql"


@unittest.skipUnless(_docker_ready(), "Disposable Docker PostgreSQL is unavailable")
class TestLegacyAllocationMigrationSQL(unittest.TestCase):
    sql = harness.TestExistingPaymentAllocationSql.__dict__["sql"]
    container = ""

    @classmethod
    def setUpClass(cls):
        cls.container = "caddy-protection-migration-" + uuid.uuid4().hex[:10]
        result = subprocess.run(
            ["docker", "run", "-d", "--rm", "--name", cls.container,
             "-e", "POSTGRES_PASSWORD=isolated-test-only", "postgres:17"],
            capture_output=True, text=True, timeout=180,
        )
        if result.returncode:
            raise RuntimeError(result.stderr)
        try:
            for _ in range(60):
                result = subprocess.run(
                    ["docker", "exec", cls.container, "pg_isready", "-h", "127.0.0.1", "-U", "postgres"],
                    capture_output=True,
                )
                if result.returncode == 0:
                    return
                time.sleep(1)
            raise RuntimeError("Disposable PostgreSQL did not start")
        except Exception:
            cls.tearDownClass()
            raise

    @classmethod
    def tearDownClass(cls):
        subprocess.run(["docker", "rm", "-f", cls.container], capture_output=True, timeout=30)

    def setUp(self):
        self.sql("""
            drop schema public cascade;
            drop schema if exists legacy_correction_private cascade;
            create schema public;
            do $$ begin
                if exists(select 1 from pg_roles where rolname='caddycheck_legacy_scope') then
                    drop owned by caddycheck_legacy_scope;
                    drop role caddycheck_legacy_scope;
                end if;
                if exists(select 1 from pg_roles where rolname='caddycheck_legacy_correction') then
                    drop owned by caddycheck_legacy_correction;
                    drop role caddycheck_legacy_correction;
                end if;
                if exists(select 1 from pg_roles where rolname='anon') then
                    drop owned by anon,authenticated,service_role;
                    drop role anon,authenticated,service_role;
                end if;
                if exists(select 1 from pg_roles where rolname='production_installer') then
                    drop owned by production_installer;
                    drop role production_installer;
                end if;
            end $$;
        """)
        self.sql(SCHEMA)
        install_production_like_rls(self.sql)
        self.sql("""
            insert into invoices(id,invoice_number,project_name,payment_amount,paid)
                values(1,'8000','Test',100,'Yes');
            insert into bank_payments(id,payment_date,invoice_number,instructed_amount,applied_amount)
                values(1,'2026-10-20','8000',100,100);
            insert into bank_payment_allocations(payment_id,invoice_row_id,invoice_number,amount_applied)
                values(1,1,8000,100);
        """)

    def production_installer(self, *, create_role=True, read_all_data=False):
        self.sql(f"""
            create role production_installer login password 'isolated-test-only' nosuperuser inherit
                {'createrole' if create_role else 'nocreaterole'} createdb bypassrls;
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
        if read_all_data:
            self.sql("grant pg_read_all_data to production_installer;")

    def assert_no_partial_install(self):
        self.assertEqual(self.sql("""
            select to_regnamespace('legacy_correction_private') is null,
                   to_regclass('public.legacy_allocation_reversals') is null,
                   to_regclass('public.legacy_invoice_settlement_once') is null,
                   not exists(select 1 from pg_roles where rolname in
                       ('caddycheck_legacy_scope','caddycheck_legacy_correction','caddycheck_legacy_install')),
                   not exists(select 1 from pg_attribute
                       where attrelid='public.bank_payment_allocations'::regclass
                         and attname='manual_invoice_before' and not attisdropped);
        """), "t|t|t|t|t")

    def test_install_as_production_non_superuser(self):
        self.production_installer()
        audit_before = self.sql("select md5(pg_get_functiondef('public.invoice_audit_trigger()'::regprocedure));")
        self.sql(MIGRATION.read_text(encoding="utf-8"), user="production_installer")
        self.assertEqual(self.sql("select md5(pg_get_functiondef('public.invoice_audit_trigger()'::regprocedure));"),
                         audit_before)
        self.assertEqual(self.sql("""
            select count(*) from pg_class where oid in (
                'public.invoices'::regclass,'public.bank_payments'::regclass,
                'public.bank_payment_allocations'::regclass,'public.invoice_change_log'::regclass)
                and relrowsecurity and not relforcerowsecurity;
        """), "4")
        self.assertEqual(self.sql("""
            select count(*) from pg_auth_members where roleid in (
                select oid from pg_roles where rolname in
                    ('caddycheck_legacy_scope','caddycheck_legacy_correction'));
        """), "0")
        self.assertEqual(self.sql("""
            select count(*) from pg_roles where rolname in
                ('caddycheck_legacy_scope','caddycheck_legacy_correction')
                and not rolcanlogin and not rolbypassrls and not rolsuper;
        """), "2")
        self.assertEqual(self.sql(
            "select count(*) from pg_roles where rolname='caddycheck_legacy_install';"
        ), "0")
        self.assertEqual(self.sql("""
            select has_schema_privilege('caddycheck_legacy_scope','public','CREATE'),
                   has_schema_privilege('caddycheck_legacy_correction','public','CREATE'),
                   has_schema_privilege('caddycheck_legacy_correction','legacy_correction_private','CREATE'),
                   has_schema_privilege('production_installer','legacy_correction_private','CREATE'),
                   has_schema_privilege('production_installer','legacy_correction_private','USAGE');
        """), "f|f|f|f|f")
        for owner in ("caddycheck_legacy_scope", "caddycheck_legacy_correction"):
            with self.subTest(owner=owner), self.assertRaisesRegex(RuntimeError, "permission denied to set role"):
                self.sql(f"set role {owner};", user="production_installer")
        self.sql((MIGRATION.parent / "add_existing_payment_allocation.sql").read_text(encoding="utf-8"),
                 user="production_installer")
        self.assertEqual(self.sql("select enabled from manual_allocation_release_policy;"), "f")
        for role in ("service_role", "anon", "authenticated"):
            with self.subTest(role=role), self.assertRaisesRegex(RuntimeError, "permission denied"):
                self.sql(f"set role {role}; update public.manual_allocation_release_policy set enabled=true;")

    def test_no_createrole_fails_atomically(self):
        self.production_installer(create_role=False)
        with self.assertRaisesRegex(RuntimeError, "permission denied to create role"):
            self.sql(MIGRATION.read_text(encoding="utf-8"), user="production_installer")
        self.assert_no_partial_install()

    def test_production_read_all_data_path_reproduces_cleanup_failure(self):
        self.production_installer(read_all_data=True)
        script = MIGRATION.read_text(encoding="utf-8").replace(
            "    expected_usage := baseline.superuser or baseline.read_all_data or baseline.write_all_data;",
            """    raise notice 'cleanup evidence: role=%, memberships=%, creator=%, owner=%, acl=%, usage=%, create=%, read_all_data=%',
                current_user,
                (select count(*) from pg_auth_members where roleid in (
                    select oid from pg_roles where rolname in ('caddycheck_legacy_scope','caddycheck_legacy_correction'))),
                to_regrole('caddycheck_legacy_install'),
                (select pg_get_userbyid(nspowner) from pg_namespace where nspname='legacy_correction_private'),
                (select nspacl from pg_namespace where nspname='legacy_correction_private'),
                has_schema_privilege(installer,'legacy_correction_private','USAGE'),
                has_schema_privilege(installer,'legacy_correction_private','CREATE'),
                pg_has_role(installer,'pg_read_all_data','USAGE');
            if has_schema_privilege(installer,'legacy_correction_private','USAGE') then
                raise exception 'Installation-only owner permissions were not fully removed';
            end if;
            expected_usage := baseline.superuser or baseline.read_all_data or baseline.write_all_data;""",
            1,
        )
        with self.assertRaisesRegex(RuntimeError, "Installation-only owner permissions") as error:
            self.sql(script, user="production_installer")
        evidence = str(error.exception)
        self.assertIn("role=production_installer, memberships=0, creator=<NULL>", evidence)
        self.assertIn("owner=caddycheck_legacy_scope", evidence)
        self.assertIn("usage=t, create=f, read_all_data=t", evidence)
        self.assert_no_partial_install()

    def test_corrected_assertion_preserves_only_preexisting_global_access(self):
        self.production_installer(read_all_data=True)
        result = self.sql(MIGRATION.read_text(encoding="utf-8"), user="production_installer")
        self.assertIn("COMMIT", result)
        self.assertEqual(self.sql("""
            select pg_has_role('production_installer','pg_read_all_data','USAGE'),
                   has_schema_privilege('production_installer','legacy_correction_private','USAGE'),
                   has_schema_privilege('production_installer','legacy_correction_private','CREATE'),
                   pg_has_role('production_installer','caddycheck_legacy_scope','SET'),
                   pg_has_role('production_installer','caddycheck_legacy_correction','SET'),
                   to_regrole('caddycheck_legacy_install') is null,
                   to_regclass('pg_temp.legacy_install_privilege_baseline') is null;
        """), "t|t|f|f|f|t|t")
        self.assertEqual(self.sql("""
            select count(*) from pg_auth_members where roleid in (
                select oid from pg_roles where rolname in
                    ('caddycheck_legacy_scope','caddycheck_legacy_correction'));
        """), "0")
        for owner in ("caddycheck_legacy_scope", "caddycheck_legacy_correction"):
            with self.subTest(owner=owner), self.assertRaisesRegex(RuntimeError, "permission denied to set role"):
                self.sql(f"set role {owner};", user="production_installer")

    def test_global_access_does_not_hide_residual_direct_schema_grants(self):
        self.production_installer(read_all_data=True)
        script = MIGRATION.read_text(encoding="utf-8").replace(
            "revoke usage, create on schema legacy_correction_private from %I",
            "revoke create on schema legacy_correction_private from %I",
        )
        with self.assertRaisesRegex(RuntimeError, "Installation-only owner permissions"):
            self.sql(script, user="production_installer")
        self.assert_no_partial_install()

    def test_global_access_mismatch_against_baseline_aborts(self):
        self.production_installer()
        self.sql("grant pg_read_all_data to production_installer with admin option;")
        # Deliberate baseline mismatch must not be accepted as an implicit-access exemption.
        script = MIGRATION.read_text(encoding="utf-8").replace(
            "pg_has_role(current_user,'pg_read_all_data','USAGE') as read_all_data",
            "false as read_all_data",
        )
        with self.assertRaisesRegex(RuntimeError, "Installation-only owner permissions"):
            self.sql(script, user="production_installer")
        self.assert_no_partial_install()

    def test_preexisting_write_all_data_access_is_reported_without_owner_memberships(self):
        self.production_installer()
        self.sql("grant pg_write_all_data to production_installer;")
        self.sql(MIGRATION.read_text(encoding="utf-8"), user="production_installer")
        self.assertEqual(self.sql("""
            select has_schema_privilege('production_installer','legacy_correction_private','USAGE'),
                   has_schema_privilege('production_installer','legacy_correction_private','CREATE'),
                   pg_has_role('production_installer','caddycheck_legacy_scope','SET');
        """), "t|f|f")

    def test_no_table_ownership_fails_atomically(self):
        self.production_installer()
        self.sql("alter table public.bank_payment_allocations owner to postgres;")
        with self.assertRaisesRegex(RuntimeError, "must be owner"):
            self.sql(MIGRATION.read_text(encoding="utf-8"), user="production_installer")
        self.assert_no_partial_install()

    def test_missing_temporary_set_permission_reproduces_original_failure(self):
        self.production_installer()
        script = MIGRATION.read_text(encoding="utf-8").replace(
            "with set true, inherit true", "with set false, inherit true",
        )
        with self.assertRaisesRegex(RuntimeError, "must be able to SET ROLE"):
            self.sql(script, user="production_installer")
        self.assert_no_partial_install()

    def test_missing_temporary_target_create_permission_aborts(self):
        self.production_installer()
        script = MIGRATION.read_text(encoding="utf-8").replace(
            "grant create on schema public to caddycheck_legacy_scope, caddycheck_legacy_correction;", "",
        )
        with self.assertRaisesRegex(RuntimeError, "permission denied for schema public"):
            self.sql(script, user="production_installer")
        self.assert_no_partial_install()

    def test_late_failure_rolls_back_objects_and_temporary_memberships(self):
        self.production_installer()
        before = self.sql("select to_jsonb(a) from bank_payment_allocations a;")
        script = MIGRATION.read_text(encoding="utf-8").replace(
            "-- Remove installation grants, then drop the disposable creator and its ADMIN links.",
            "do $$ begin raise exception 'injected installer failure'; end $$;",
        )
        with self.assertRaisesRegex(RuntimeError, "injected installer failure"):
            self.sql(script, user="production_installer")
        self.assert_no_partial_install()
        self.assertEqual(self.sql("select to_jsonb(a) from bank_payment_allocations a;"), before)

    def rejected(self, expected):
        with self.assertRaisesRegex(RuntimeError, expected):
            self.sql(MIGRATION.read_text(encoding="utf-8"))
        self.assertEqual(self.sql("""
            select count(*) from pg_attribute
            where attrelid='bank_payment_allocations'::regclass
              and attname='manual_invoice_before' and not attisdropped;
        """), "0")
        self.assertEqual(self.sql("select count(*) from bank_payment_allocations;"), self.count)

    def test_null_historical_link_blocks_without_rewriting(self):
        self.sql("update bank_payment_allocations set invoice_row_id=null;")
        self.count = "1"
        self.rejected("null values")

    def test_orphan_historical_link_blocks_without_rewriting(self):
        self.sql("update bank_payment_allocations set invoice_row_id=999;")
        self.count = "1"
        self.rejected("foreign key constraint")

    def test_duplicate_historical_settlement_blocks_without_rewriting(self):
        self.sql("""insert into bank_payment_allocations(payment_id,invoice_row_id,invoice_number,amount_applied)
                    values(1,1,8000,100);""")
        self.count = "2"
        self.rejected("could not create unique index")

    def test_financial_writer_schema_blocks_installation(self):
        self.sql("alter table bank_payment_allocations add column writer_operation_id uuid;")
        self.count = "1"
        self.rejected("financial-writer backend")

    def test_financial_writer_trigger_even_disabled_blocks_installation(self):
        self.sql("""
            create function public.dummy_guard() returns trigger language plpgsql as $$ begin return new; end $$;
            create trigger bank_allocations_writer_guard before insert on bank_payment_allocations
              for each row execute function public.dummy_guard();
            alter table bank_payment_allocations disable trigger bank_allocations_writer_guard;
        """)
        self.count = "1"
        self.rejected("financial-writer backend")

    def test_clean_history_installs_without_operational_data_changes(self):
        before = self.sql("select to_jsonb(a) from bank_payment_allocations a;")
        self.sql(MIGRATION.read_text(encoding="utf-8"))
        after = self.sql("select to_jsonb(a)-'manual_invoice_before' from bank_payment_allocations a;")
        self.assertEqual(after, before)
        self.assertEqual(self.sql("select count(*) from legacy_allocation_reversals;"), "0")
        self.assertEqual(self.sql("select count(*) from bank_payments;"), "1")
        self.assertEqual(self.sql("select count(*) from invoices;"), "1")

    def test_serial_audit_default_grants_only_usage_if_needed(self):
        self.sql("""
            alter table invoice_change_log alter column id drop identity;
            create sequence public.audit_serial_test owned by invoice_change_log.id;
            alter table invoice_change_log alter column id set default nextval('public.audit_serial_test');
        """)
        self.sql(MIGRATION.read_text(encoding="utf-8"))
        self.assertEqual(self.sql("""
            select has_sequence_privilege('caddycheck_legacy_correction','public.audit_serial_test','USAGE'),
                   has_sequence_privilege('caddycheck_legacy_correction','public.audit_serial_test','SELECT'),
                   has_sequence_privilege('caddycheck_legacy_correction','public.audit_serial_test','UPDATE');
        """), "t|f|f")
