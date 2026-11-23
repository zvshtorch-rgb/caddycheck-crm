import contextlib
import hashlib
import json
import os
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, Mock, patch

from services import deployment_diagnostics as diag


SECRET = "NEVER_DISPLAY_THIS_SECRET"
DEC_CANDIDATE = "11111111-1111-1111-1111-111111111111"
DEC_EXECUTION = "22222222-2222-2222-2222-222222222222"


class CodedError(Exception):
    def __init__(self, code):
        self.code = self.sqlstate = code
        super().__init__(SECRET)


class SelectOnlyClient:
    def __init__(self, data=None, errors=None):
        self.data = data or {}
        self.errors = errors or {}
        self.calls = []

    def table(self, name):
        client = self

        class Query:
            def select(self, columns):
                client.calls.append((name, columns))
                assert columns == diag.TABLES[name]
                assert "*" not in columns
                return self

            def limit(self, number):
                assert number == diag.LIMIT + 1
                return self

            def execute(self):
                if name in client.errors:
                    raise CodedError(client.errors[name])
                return SimpleNamespace(data=client.data.get(name, []))

        return Query()


class DiagnosticsTests(unittest.TestCase):
    def test_gate_source_identity_and_secrets_fallback_are_sanitized(self):
        from services import existing_payment_allocation as allocation
        name = allocation.MANUAL_ALLOCATION_GATE
        source = Path(allocation.__file__).resolve()
        for value, effective in (("ON", "ON"), ("OFF", "OFF"), (SECRET, "OFF")):
            with self.subTest(value=value), patch.dict(os.environ, {}, clear=True), \
                    patch("streamlit.secrets", {name: value}):
                result = diag._manual_allocation_gate_source()
                self.assertEqual(result, {
                    "resolved_file": str(source),
                    "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
                    "loaded_reader_supports_secrets_fallback": True,
                    "env_key_present": False, "root_secret_key_present": True,
                    "selected_source": "STREAMLIT_SECRET", "effective_state": effective,
                })
                self.assertNotIn(SECRET, json.dumps(result))

    def test_gate_source_environment_precedence_and_no_source(self):
        from services.existing_payment_allocation import MANUAL_ALLOCATION_GATE as name
        for env, secrets, source, effective in (
            ({name: "OFF"}, {name: "ON"}, "ENV", "OFF"),
            ({name: "ON"}, {}, "ENV", "ON"),
            ({name: SECRET}, {name: "ON"}, "ENV", "OFF"),
            ({}, {}, "NONE", "OFF"),
            ({}, {"order_confirmation": {name: "ON"}}, "NONE", "OFF"),
        ):
            with self.subTest(source=source, effective=effective), \
                    patch.dict(os.environ, env, clear=True), patch("streamlit.secrets", secrets):
                result = diag._manual_allocation_gate_source()
                self.assertEqual(result["env_key_present"], name in env)
                self.assertEqual(result["root_secret_key_present"], name in secrets)
                self.assertEqual(result["selected_source"], source)
                self.assertEqual(result["effective_state"], effective)
                self.assertNotIn(SECRET, json.dumps(result))

    def test_gate_source_detects_stale_loaded_reader_despite_current_disk_source(self):
        from services import existing_payment_allocation as allocation

        def legacy_reader():
            value = os.environ.get(allocation.MANUAL_ALLOCATION_GATE)
            return {"effective_state": "ON" if value == "ON" else "OFF"}

        with patch.dict(os.environ, {}, clear=True), \
                patch("streamlit.secrets", {allocation.MANUAL_ALLOCATION_GATE: "ON"}), \
                patch.object(allocation, "manual_allocation_gate_state", legacy_reader):
            result = diag._manual_allocation_gate_source()
        self.assertFalse(result["loaded_reader_supports_secrets_fallback"])
        self.assertTrue(result["root_secret_key_present"])
        self.assertEqual(result["selected_source"], "NONE")
        self.assertEqual(result["effective_state"], "OFF")

    def test_gate_source_accounts_for_streamlit_environment_promotion(self):
        from services.existing_payment_allocation import MANUAL_ALLOCATION_GATE as name

        class PromotingSecrets(dict):
            def __contains__(self, key):
                os.environ[name] = "ON"
                return super().__contains__(key)

        with patch.dict(os.environ, {}, clear=True), \
                patch("streamlit.secrets", PromotingSecrets({name: "ON"})):
            result = diag._manual_allocation_gate_source()
        self.assertTrue(result["env_key_present"])
        self.assertTrue(result["root_secret_key_present"])
        self.assertEqual(result["selected_source"], "ENV")
        self.assertEqual(result["effective_state"], "ON")

    def test_gate_source_collection_is_admin_only_and_errors_are_sanitized(self):
        with patch.object(diag, "_manual_allocation_gate_source", side_effect=OSError(SECRET)) as source, \
                patch.object(diag, "_module_identity", return_value={}), \
                patch.object(diag, "_commit", return_value={}), \
                patch.object(diag, "runtime_snapshot", return_value={}), \
                patch.object(diag, "database_snapshot", return_value={}):
            with self.assertRaises(PermissionError):
                diag.collect_diagnostics("viewer")
            source.assert_not_called()
            with self.assertLogs(diag.logger, level="WARNING") as logs:
                result = diag.collect_diagnostics("admin")
        self.assertEqual(result["manual_allocation_gate_source"]["status"], "UNAVAILABLE")
        self.assertNotIn(SECRET, json.dumps(result) + str(logs.output))

    def test_module_identity_matches_source_and_loaded_flags_without_configuration_reads(self):
        source = Path(diag.__file__).resolve()
        expected_hash = hashlib.sha256(source.read_bytes()).hexdigest()
        with patch.object(diag, "_process_start", return_value={
            "status": "UNAVAILABLE", "started_at_utc": None,
        }), patch.dict(os.environ, {"SENSITIVE_SETTING": SECRET}), \
                patch("config.settings.get_email_config") as configuration, \
                patch.object(diag, "database_snapshot") as database, \
                patch.object(diag.subprocess, "run") as command:
            result = diag._module_identity()
        self.assertEqual(result["status"], "VERIFIED")
        self.assertEqual(result["resolved_file"], str(source))
        self.assertEqual(result["spec_origin"], diag.__spec__.origin)
        self.assertEqual(result["cached"], diag.__cached__)
        self.assertEqual(result["source_sha256"], expected_hash)
        self.assertEqual(result["runtime_snapshot_code_filename"], diag.runtime_snapshot.__code__.co_filename)
        self.assertTrue(result["manual_gate_in_loaded_gates"])
        self.assertTrue(result["release_policy_in_loaded_tables"])
        self.assertNotIn(SECRET, json.dumps(result))
        configuration.assert_not_called()
        database.assert_not_called()
        command.assert_not_called()

    def test_module_identity_reports_actual_loaded_flags(self):
        with patch.object(diag, "GATES", {}), patch.object(diag, "TABLES", {}):
            result = diag._module_identity()
        self.assertFalse(result["manual_gate_in_loaded_gates"])
        self.assertFalse(result["release_policy_in_loaded_tables"])
        self.assertEqual(result["runtime_snapshot_code_filename"], diag.runtime_snapshot.__code__.co_filename)

    def test_module_identity_read_failure_is_explicit_and_sanitized(self):
        with patch.object(Path, "read_bytes", side_effect=OSError(SECRET)), \
                self.assertLogs(diag.logger, level="WARNING") as logs:
            result = diag._module_identity()
        self.assertEqual(result["status"], "UNAVAILABLE")
        self.assertIsNone(result["source_sha256"])
        self.assertNotIn(SECRET, json.dumps(result) + str(logs.output))

    def test_process_start_uses_kernel_start_ticks_not_module_import_time(self):
        fields = ["0"] * 20
        fields[19] = "250"
        with patch.object(diag.sys, "platform", "linux"), \
                patch.object(Path, "read_text", side_effect=[
                    "42 (python name (worker)) " + " ".join(fields),
                    "cpu 1 2 3\nbtime 1700000000\n",
                ]), patch.object(diag.os, "sysconf", return_value=100, create=True):
            result = diag._process_start()
        self.assertEqual(result, {
            "status": "VERIFIED", "started_at_utc": "2023-11-14T22:13:22.500000+00:00",
        })

    def test_process_start_unavailable_is_explicit_and_sanitized(self):
        with patch.object(diag.sys, "platform", "win32"), patch.object(Path, "read_text") as read:
            self.assertEqual(diag._process_start(),
                             {"status": "UNAVAILABLE", "started_at_utc": None})
            read.assert_not_called()
        with patch.object(diag.sys, "platform", "linux"), \
                patch.object(Path, "read_text", side_effect=OSError(SECRET)), \
                self.assertLogs(diag.logger, level="WARNING") as logs:
            result = diag._process_start()
        self.assertEqual(result, {"status": "UNAVAILABLE", "started_at_utc": None})
        self.assertNotIn(SECRET, str(logs.output))

    def test_collection_includes_identity_only_after_admin_authorization(self):
        with patch.object(diag, "_module_identity", return_value={"sanitized": True}) as identity, \
                patch.object(diag, "_commit", return_value={}), \
                patch.object(diag, "runtime_snapshot", return_value={}), \
                patch.object(diag, "database_snapshot", return_value={}):
            with self.assertRaises(PermissionError):
                diag.collect_diagnostics("viewer")
            identity.assert_not_called()
            self.assertEqual(diag.collect_diagnostics("admin")["diagnostics_module_identity"],
                             {"sanitized": True})
            identity.assert_called_once_with()

    def runtime(self, *, graph=False, settings=None, environ=None):
        config = {"sender_email": "sender@example.com", "smtp_username": "fallback@example.com",
                  "smtp_password": SECRET, "default_recipients": ["to@example.com"],
                  "default_cc": []}
        config.update(settings or {})
        values = {"INVOICE_SENDER_MAILBOX": "graph@example.com",
                  "ORDER_INTAKE_GRAPH_MAILBOX": "fallback-graph@example.com"}
        with patch.dict(os.environ, environ or {}, clear=True), \
                patch("config.settings.get_email_config", return_value=config), \
                patch("services.email_service.graph_email_available", return_value=graph), \
                patch("services.email_service._get_graph_setting", side_effect=lambda k: values[k]):
            return diag.runtime_snapshot()

    def test_smtp_resolution_and_no_secret_values(self):
        result = self.runtime()
        self.assertEqual(result["monthly_transport"], "SMTP")
        self.assertEqual(result["effective_sender"], "sender@example.com")
        self.assertNotIn(SECRET, json.dumps(result))
        fallback = self.runtime(settings={"sender_email": ""})
        self.assertEqual(fallback["effective_sender"], "fallback@example.com")

    def test_graph_selection_and_mailbox_fallback(self):
        result = self.runtime(graph=True)
        self.assertTrue(result["graph_email_available"])
        self.assertEqual(result["monthly_transport"], "Graph")
        self.assertEqual(result["effective_sender"], "graph@example.com")
        with patch("config.settings.get_email_config", return_value={}), \
                patch("services.email_service.graph_email_available", return_value=True), \
                patch("services.email_service._get_graph_setting",
                      side_effect=lambda k: "" if k == "INVOICE_SENDER_MAILBOX" else "fallback@example.com"):
            self.assertEqual(diag.runtime_snapshot()["effective_sender"], "fallback@example.com")

    def test_invalid_addresses_are_not_echoed(self):
        result = self.runtime(settings={"sender_email": SECRET,
                                       "default_recipients": [SECRET], "default_cc": SECRET})
        self.assertEqual(result["sender_status"], "UNAVAILABLE")
        self.assertEqual(result["to"]["status"], "UNAVAILABLE")
        self.assertNotIn(SECRET, json.dumps(result))

    @patch.object(Path, "is_file", return_value=True)
    def test_gate_states_and_invalid_production_staging(self, _supported_modules):
        result = self.runtime(environ={"CADDYCHECK_COVERAGE_WRITERS": "on",
                                      "CADDYCHECK_STAGING_FINANCIAL": "off",
                                      "CADDYCHECK_STAGING_SEND": "on"})
        gates = result["gates"]
        self.assertEqual(gates["CADDYCHECK_COVERAGE_WRITERS"]["state"], "ON")
        self.assertEqual(gates["CADDYCHECK_STAGING_FINANCIAL"]["state"], "OFF")
        self.assertEqual(gates["CADDYCHECK_STAGING_MONTHLY"],
                         {"state": "UNSET", "effective_state": "OFF"})
        self.assertEqual(gates["CADDYCHECK_STAGING_SEND"]["effective_state"], "UNSUPPORTED_BY_BUILD")
        self.assertEqual(self.runtime(environ={"CADDYCHECK_COVERAGE_WRITERS": SECRET})["gates"]
                         ["CADDYCHECK_COVERAGE_WRITERS"]["state"], "UNAVAILABLE")
        with patch.object(Path, "is_file", return_value=False):
            self.assertEqual(self.runtime()["gates"]["CADDYCHECK_COVERAGE_WRITERS"]["state"],
                             "UNSUPPORTED_BY_BUILD")

    def test_manual_allocation_gate_diagnostics_fail_closed_and_sanitized(self):
        name = "CADDYCHECK_MANUAL_ALLOCATION_ENABLED"
        for value, expected in (
            (None, {"state": "UNSET", "effective_state": "OFF"}),
            ("off", {"state": "OFF", "effective_state": "OFF"}),
            ("on", {"state": "ON", "effective_state": "ON"}),
            (SECRET, {"state": "INVALID", "effective_state": "OFF"}),
        ):
            with self.subTest(value=value):
                result = self.runtime(environ={} if value is None else {name: value})
                self.assertEqual(result["gates"][name], expected)
                self.assertNotIn(SECRET, json.dumps(result))

    def test_manual_gate_secret_fallback_is_sanitized_and_coverage_reader_is_unchanged(self):
        name = "CADDYCHECK_MANUAL_ALLOCATION_ENABLED"
        coverage = "CADDYCHECK_COVERAGE_WRITERS"
        for value, expected in (("ON", "ON"), ("OFF", "OFF"), (SECRET, "OFF")):
            with self.subTest(value=value), patch("streamlit.secrets", {name: value, coverage: "ON"}):
                result = self.runtime()
                self.assertEqual(result["gates"][name]["effective_state"], expected)
                self.assertEqual(result["gates"][coverage], {"state": "UNSET", "effective_state": "OFF"})
                self.assertNotIn(SECRET, json.dumps(result))

    def test_database_manual_release_diagnostics_are_select_only_and_sanitized(self):
        client = SelectOnlyClient({
            "manual_allocation_release_policy": [
                {"singleton": True, "enabled": False, "unexpected_secret": SECRET}
            ],
        })
        result = diag.rest_inventory(client)
        entry = result["tables"]["public.manual_allocation_release_policy"]
        self.assertEqual(entry["records"], [{"singleton": True, "enabled": False}])
        self.assertNotIn(SECRET, json.dumps(result))
        self.assertIn(("manual_allocation_release_policy", "singleton,enabled"), client.calls)

    def test_admin_guard_precedes_all_configuration_and_database_reads(self):
        with patch.object(diag, "runtime_snapshot") as runtime, \
                patch.object(diag, "database_snapshot") as database, \
                patch.object(diag, "_commit") as build:
            for role in ("viewer", "editor", "finance", "", "Admin"):
                with self.assertRaises(PermissionError):
                    diag.collect_diagnostics(role)
                ui = Mock()
                diag.render_diagnostics(ui, role)
                self.assertEqual(ui.mock_calls, [])
            runtime.assert_not_called()
            database.assert_not_called()
            build.assert_not_called()

    def test_rest_error_classification_and_fail_closed_absence(self):
        client = SelectOnlyClient(errors={"monthly_candidate_executions": "PGRST205",
                                          "monthly_candidate_documents": "42501"})
        result = diag.rest_inventory(client)
        self.assertEqual(result["tables"]["public.monthly_candidate_executions"]["status"], "UNAVAILABLE")
        self.assertEqual(result["tables"]["public.monthly_candidate_documents"]["status"], "PERMISSION_DENIED")
        self.assertEqual(result["tables"]["public.monthly_invoice_candidates"]["status"], "ACCESSIBLE_EMPTY")
        self.assertEqual(result["unresolved_send"], "NOT_VERIFIED")
        self.assertNotIn(SECRET, json.dumps(result))

    def test_december_linkage_unknown_attempt_and_document(self):
        client = SelectOnlyClient({
            "monthly_invoice_candidates": [
                {"id": DEC_CANDIDATE, "billing_period_start": diag.PERIOD, "status": "prepared"},
                {"id": "333", "billing_period_start": "2027-07-01", "status": "prepared"}],
            "monthly_candidate_executions": [
                {"id": DEC_EXECUTION, "candidate_id": DEC_CANDIDATE, "state": "unknown", "invoice_number": 9999},
                {"id": "444", "candidate_id": "333", "state": "unknown"}],
            "monthly_candidate_documents": [{"id": "555", "execution_id": DEC_EXECUTION,
                                              "pdf_bytes": SECRET, "envelope": SECRET}],
            "monthly_candidate_send_attempts": [
                {"id": "666", "execution_id": DEC_EXECUTION, "state": "unknown"},
                {"id": "777", "execution_id": "444", "state": "unknown"}],
        })
        result = diag.rest_inventory(client)
        self.assertEqual(result["unresolved_send"], "FOUND")
        attempts = result["tables"]["public.monthly_candidate_send_attempts"]
        self.assertEqual(len(attempts["records"]), 1)
        self.assertEqual(attempts["unscoped_record_count"], 1)
        self.assertEqual(result["tables"]["public.monthly_candidate_documents"]["status"], "RECORDS_FOUND")
        self.assertNotIn(SECRET, json.dumps(result))

    def test_unlinked_records_and_truncation_cannot_prove_absence(self):
        entries = {"private.monthly_candidate_send_attempts": {
            "status": "ACCESSIBLE", "complete": True,
            "rows": [{"execution_id": "missing", "state": "unknown"}]}}
        result = diag._summarize(entries, inventory_complete=True)
        self.assertFalse(result["inventory_complete"])
        self.assertEqual(result["unresolved_send"], "NOT_VERIFIED")
        result = diag.rest_inventory(SelectOnlyClient({
            "monthly_invoice_candidates": [{"id": str(i), "billing_period_start": "2027-07-01"}
                                            for i in range(diag.LIMIT + 1)]}))
        self.assertFalse(result["tables"]["public.monthly_invoice_candidates"]["complete"])

    def test_build_errors_are_sanitized(self):
        with patch.object(diag.subprocess, "run", return_value=SimpleNamespace(
                returncode=1, stdout=SECRET, stderr=SECRET)):
            self.assertEqual(diag._commit(), {"status": "UNAVAILABLE", "commit_sha": None})

    def test_build_reports_commit_and_tracked_dirty_state(self):
        with patch.object(diag.subprocess, "run", side_effect=[
                SimpleNamespace(returncode=0, stdout="a" * 40 + "\n"),
                SimpleNamespace(returncode=1)]):
            self.assertEqual(diag._commit(), {"status": "VERIFIED", "commit_sha": "a" * 40,
                                             "tracked_worktree_dirty": True})

    def test_render_is_lazy_and_does_not_transport_or_write(self):
        ui = Mock()
        ui.expander.return_value = contextlib.nullcontext()
        ui.button.return_value = False
        with patch.object(diag, "collect_diagnostics") as collect:
            diag.render_diagnostics(ui, "admin")
            collect.assert_not_called()
            ui.json.assert_not_called()
        ui.button.return_value = True
        client = SelectOnlyClient()
        with patch("config.settings.get_email_config", return_value={
            "sender_email": "sender@example.com", "smtp_password": SECRET,
            "default_recipients": ["to@example.com"], "default_cc": []}), \
                patch("services.email_service.graph_email_available", return_value=False), \
                patch("services.supabase_service._get_client", return_value=client), \
                patch.dict(os.environ, {}, clear=True), \
                patch.object(diag, "_commit", return_value={"status": "VERIFIED", "commit_sha": "a" * 40}), \
                patch("requests.sessions.Session.request", side_effect=AssertionError("transport forbidden")), \
                patch("smtplib.SMTP", side_effect=AssertionError("SMTP forbidden")), \
                patch("smtplib.SMTP_SSL", side_effect=AssertionError("SMTP forbidden")), \
                patch("config.settings.load_sent_invoices_log", side_effect=AssertionError("sync forbidden")), \
                patch("services.supabase_service.append_invoice_rows", side_effect=AssertionError("write forbidden")), \
                patch("services.supabase_service.get_next_invoice_number", side_effect=AssertionError("reservation forbidden")):
            diag.render_diagnostics(ui, "admin")
        result = ui.json.call_args.args[0]
        self.assertNotIn(SECRET, json.dumps(result))
        self.assertEqual(result["runtime"]["effective_sender"], "sender@example.com")
        self.assertTrue(client.calls)
        self.assertIn("observed_at_utc", result)

    def test_graph_runtime_collection_never_requests_token_or_sends(self):
        settings = {"default_recipients": ["to@example.com"], "default_cc": []}
        values = {key: SECRET for key in ("ORDER_INTAKE_GRAPH_TENANT_ID",
                                         "ORDER_INTAKE_GRAPH_CLIENT_ID", "ORDER_INTAKE_GRAPH_CLIENT_SECRET")}
        values["ORDER_INTAKE_GRAPH_MAILBOX"] = "graph@example.com"
        with patch("config.settings.get_email_config", return_value=settings), \
                patch("services.email_service._get_graph_setting", side_effect=lambda k: values.get(k, "")), \
                patch("requests.sessions.Session.request", side_effect=AssertionError("network forbidden")), \
                patch("services.email_service.send_graph_email", side_effect=AssertionError("send forbidden")), \
                patch.dict(os.environ, {}, clear=True):
            result = diag.runtime_snapshot()
        self.assertTrue(result["graph_email_available"])
        self.assertEqual(result["effective_sender"], "graph@example.com")
        self.assertNotIn(SECRET, json.dumps(result))

    def test_streamlit_render_admin_only_and_explicit_read(self):
        from streamlit.testing.v1 import AppTest

        script = """
import streamlit as st
from services.deployment_diagnostics import render_diagnostics
render_diagnostics(st, st.session_state.get("test_role", "viewer"))
"""
        with patch.object(diag, "collect_diagnostics", return_value={"sanitized": True}) as collect:
            app = AppTest.from_string(script)
            app.run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.button), 0)
            collect.assert_not_called()
            app.session_state["test_role"] = "admin"
            app.run()
            self.assertFalse(app.exception)
            collect.assert_not_called()
            app.button[0].click().run()
            self.assertFalse(app.exception)
            collect.assert_called_once_with("admin")
            self.assertEqual(len(app.json), 1)


class SQLTests(unittest.TestCase):
    def connection(self, relations, rows=None, failure=None):
        connection = MagicMock()
        cursor = connection.cursor.return_value.__enter__.return_value
        cursor.fetchall.side_effect = [relations] + (rows or [])
        if failure:
            cursor.execute.side_effect = [None, CodedError(failure)]
        return connection, cursor

    def sql_module(self):
        class SQL(str):
            def format(self, *args):
                return SQL(str(self).format(*args))

            def join(self, values):
                return SQL(str(self).join(values))

        return SimpleNamespace(sql=SimpleNamespace(
            SQL=SQL, Identifier=lambda x: '"' + x.replace('"', '""') + '"',
            Literal=lambda x: "'" + x.replace("'", "''") + "'"))

    def test_catalog_absent_and_read_only_selects(self):
        connection, cursor = self.connection([])
        with patch.dict(sys.modules, {"psycopg": self.sql_module()}):
            result = diag.sql_inventory(connection)
        self.assertTrue(connection.read_only)
        self.assertTrue(result["inventory_complete"])
        self.assertEqual(result["unresolved_send"], "NONE_FOUND")
        self.assertEqual(result["tables"]["*.monthly_candidate_send_attempts"]["status"], "ABSENT")
        for call in cursor.execute.call_args_list:
            self.assertTrue(str(call.args[0]).startswith("SELECT "))
        connection.commit.assert_not_called()

    def test_nonpublic_permission_denied_and_rls(self):
        connection, cursor = self.connection([("review_private", "execution_controls", False)], failure="42501")
        with patch.dict(sys.modules, {"psycopg": self.sql_module()}):
            result = diag.sql_inventory(connection)
        self.assertEqual(result["tables"]["review_private.execution_controls"]["status"], "PERMISSION_DENIED")
        self.assertFalse(result["inventory_complete"])
        self.assertNotIn(SECRET, json.dumps(result))
        connection, _ = self.connection([("public", "monthly_candidate_executions", True)], rows=[[]])
        with patch.dict(sys.modules, {"psycopg": self.sql_module()}):
            result = diag.sql_inventory(connection)
        self.assertFalse(result["inventory_complete"])

    def test_unrecognized_claim_table_blocks_clearance(self):
        connection, _ = self.connection([("hidden", "monthly_send_claims", False)])
        with patch.dict(sys.modules, {"psycopg": self.sql_module()}):
            result = diag.sql_inventory(connection)
        self.assertEqual(result["tables"]["hidden.monthly_send_claims"]["status"], "UNAVAILABLE")
        self.assertEqual(result["unresolved_send"], "NOT_VERIFIED")

    def test_missing_driver_is_explicit_and_never_uses_rest_fallback(self):
        with patch.dict(os.environ, {"DIAGNOSTICS_READONLY_DATABASE_URL": SECRET}), \
                patch.object(diag.importlib.util, "find_spec", return_value=None), \
                patch("services.supabase_service._get_client") as client:
            result = diag.database_snapshot()
        self.assertEqual(result["status"], "UNAVAILABLE")
        client.assert_not_called()
        self.assertNotIn(SECRET, json.dumps(result))

    def test_sql_projection_is_allowlisted_and_policy_values_not_payloads(self):
        connection, cursor = self.connection(
            [("review_private", "execution_controls", False)],
            rows=[[({"monthly_enabled": False, "send_enabled": True, "secret": SECRET},)]])
        with patch.dict(sys.modules, {"psycopg": self.sql_module()}):
            result = diag.sql_inventory(connection)
        self.assertTrue(result["inventory_complete"])
        record = result["tables"]["review_private.execution_controls"]["records"][0]
        self.assertEqual(record, {"monthly_enabled": False, "send_enabled": True})
        self.assertNotIn(SECRET, json.dumps(result))
        for call in cursor.execute.call_args_list:
            self.assertTrue(str(call.args[0]).startswith("SELECT "))
            self.assertNotIn("pdf_bytes", str(call.args[0]))
        connection.commit.assert_not_called()


if __name__ == "__main__":
    unittest.main()
