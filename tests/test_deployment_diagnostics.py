import contextlib
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
