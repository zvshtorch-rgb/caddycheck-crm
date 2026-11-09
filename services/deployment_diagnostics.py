"""Admin-only, sanitized runtime and December 2026 metadata inspection."""
from __future__ import annotations

import datetime as dt
import importlib.util
import logging
import os
from pathlib import Path
import re
import subprocess

logger = logging.getLogger(__name__)
ROOT = Path(__file__).resolve().parents[1]
PERIOD = "2026-12-01"
LIMIT = 1000
TABLES = {
    "monthly_invoice_candidates": "id,billing_period_start,status",
    "monthly_candidate_executions": "id,candidate_id,state,invoice_number",
    "monthly_candidate_documents": "id,execution_id,created_at",
    "monthly_candidate_send_attempts": "id,execution_id,state,started_at,finished_at",
    "monthly_candidate_execution_events": "id,execution_id,stage,from_state,to_state",
    "invoice_financial_coverage": "id,project_id,billing_period_start,invoice_number",
    "monthly_candidate_execution_policy": "execution_enabled,coverage_inventory_verified",
    "financial_writer_policy": "enabled,coverage_inventory_verified",
    "monthly_gateway_policy": "monthly_enabled,send_enabled",
    "execution_controls": "monthly_enabled,send_enabled",
    "policy": "domain",
}
GATES = {
    "CADDYCHECK_COVERAGE_WRITERS": "financial_writer_adapter.py",
    "CADDYCHECK_STAGING_FINANCIAL": "staging_controls.py",
    "CADDYCHECK_STAGING_MONTHLY": "staging_controls.py",
    "CADDYCHECK_STAGING_SEND": "staging_controls.py",
}


def _admin(role: str) -> None:
    if role != "admin":
        raise PermissionError("Administrator diagnostics only.")


def _error_code(exc: Exception) -> str:
    value = getattr(exc, "sqlstate", None) or getattr(exc, "code", None)
    known = {"42501", "42P01", "3F000", "42703", "57014", "08001", "08006",
             "28P01", "28000", "PGRST205", "PGRST204", "PGRST301", "PGRST302"}
    return value if isinstance(value, str) and value in known else "UNAVAILABLE"


def _failure(exc: Exception, *, rest: bool = False) -> dict:
    code = _error_code(exc)
    # PGRST205 cannot distinguish absence from an unexposed/stale schema cache.
    status = ("ABSENT" if not rest and code in ("42P01", "3F000") else
              "PERMISSION_DENIED" if code in ("42501", "PGRST301", "PGRST302") else "UNAVAILABLE")
    logger.warning("Read-only deployment diagnostic failed: %s (%s)", status, code)
    return {"status": status, "error_code": code}


def _mailbox(value: object) -> str | None:
    text = str(value or "").strip()
    return text if re.fullmatch(r"[^@\s<>,;]+@[^@\s<>,;]+\.[^@\s<>,;]+", text) else None


def _recipients(value: object) -> dict:
    parts = value.split(",") if isinstance(value, str) else value
    if not isinstance(parts, (list, tuple)):
        return {"status": "UNAVAILABLE", "addresses": []}
    addresses = [_mailbox(part) for part in parts]
    return {"status": "VERIFIED" if all(addresses) else "UNAVAILABLE",
            "addresses": [address for address in addresses if address]}


def runtime_snapshot() -> dict:
    from config.settings import get_email_config
    from services.email_service import _get_graph_setting, graph_email_available

    config = get_email_config()
    graph = graph_email_available()
    sender = ((_get_graph_setting("INVOICE_SENDER_MAILBOX")
               or _get_graph_setting("ORDER_INTAKE_GRAPH_MAILBOX")) if graph else
              (config.get("sender_email") or config.get("smtp_username")))
    gates = {}
    for name, module in GATES.items():
        if not (ROOT / "services" / module).is_file():
            state = "UNSUPPORTED_BY_BUILD"
        elif name not in os.environ:
            state = "UNSET"
        else:
            state = {"on": "ON", "off": "OFF"}.get(os.environ[name].strip().lower(), "UNAVAILABLE")
        if state == "UNAVAILABLE":
            logger.warning("Deployment diagnostic gate has an invalid value: %s", name)
        effective = "OFF" if state == "UNSET" else state
        if (state == "ON" and name.startswith("CADDYCHECK_STAGING_")
                and os.environ.get("CADDYCHECK_IDENTITY_ENVIRONMENT") not in ("local", "staging")):
            effective = "UNSUPPORTED_BY_BUILD"
        gates[name] = {"state": state, "effective_state": effective}
    return {
        "graph_email_available": graph, "monthly_transport": "Graph" if graph else "SMTP",
        "effective_sender": _mailbox(sender),
        "sender_status": "VERIFIED" if _mailbox(sender) else "UNAVAILABLE",
        "to": _recipients(config.get("default_recipients", [])),
        "cc": _recipients(config.get("default_cc", [])),
        "gates": gates,
        "gate_semantics": "Configured environment states, not proof of authorization. UNSET defaults OFF. "
                          "Staging ON is only valid in local/staging, never Production.",
    }


def _commit() -> dict:
    try:
        result = subprocess.run(["git", "rev-parse", "HEAD"], cwd=ROOT, capture_output=True,
                                text=True, timeout=5, check=False)
        sha = result.stdout.strip()
        if result.returncode == 0 and re.fullmatch(r"[0-9a-f]{40,64}", sha):
            dirty = subprocess.run(["git", "diff", "--quiet", "HEAD"], cwd=ROOT, capture_output=True,
                                   timeout=5, check=False)
            return {"status": "VERIFIED", "commit_sha": sha,
                    "tracked_worktree_dirty": {0: False, 1: True}.get(dirty.returncode)}
    except (OSError, subprocess.TimeoutExpired):
        pass
    logger.warning("Deployment diagnostic build identity unavailable.")
    return {"status": "UNAVAILABLE", "commit_sha": None}


def _metadata(row: dict, table: str) -> dict:
    result = {}
    for key in TABLES[table].split(","):
        value = row.get(key)
        if value is None:
            continue
        if key in ("execution_enabled", "coverage_inventory_verified", "monthly_enabled", "send_enabled", "enabled"):
            result[key] = value if isinstance(value, bool) else "UNAVAILABLE"
        elif key in ("status", "state", "stage", "from_state", "to_state", "domain"):
            allowed = {"prepared", "approved", "invoice_committed", "send_pending", "sending", "sent",
                       "accepted", "failed", "unknown", "cancelled", "superseded", "approval",
                       "commit", "document", "claim_send", "finish_send", "reconcile_send",
                       "monthly", "financial", "lifecycle"}
            result[key] = value if isinstance(value, str) and value in allowed else "UNAVAILABLE"
        elif re.fullmatch(r"[0-9a-fA-F-]{1,64}|[0-9T:.+Z -]{1,40}", str(value)):
            result[key] = str(value)
        else:
            result[key] = "UNAVAILABLE"
    return result


def _summarize(entries: dict, *, inventory_complete: bool) -> dict:
    candidates = set()
    executions = set()
    known_candidates = set()
    known_executions = set()
    for name, entry in entries.items():
        if name.endswith(".monthly_invoice_candidates"):
            known_candidates.update(r["id"] for r in entry.get("rows", []) if "id" in r)
            candidates.update(r["id"] for r in entry.get("rows", [])
                              if r.get("billing_period_start") == PERIOD and "id" in r)
    for name, entry in entries.items():
        if name.endswith(".monthly_candidate_executions"):
            known_executions.update(r["id"] for r in entry.get("rows", []) if "id" in r)
            executions.update(r["id"] for r in entry.get("rows", [])
                              if r.get("candidate_id") in candidates and "id" in r)
    found = unresolved = False
    for name, entry in entries.items():
        table = name.split(".")[-1]
        rows = entry.pop("rows", [])
        if table == "monthly_candidate_executions":
            linked = all(r.get("candidate_id") in known_candidates for r in rows)
        elif table in ("monthly_candidate_documents", "monthly_candidate_send_attempts", "monthly_candidate_execution_events"):
            linked = all(r.get("execution_id") in known_executions for r in rows)
        elif table in ("monthly_invoice_candidates", "invoice_financial_coverage"):
            linked = all(re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(r.get("billing_period_start", ""))) for r in rows)
        else:
            linked = True
        if not linked:
            entry["complete"] = False
            entry["scope_warning"] = "Unlinked or ambiguous period metadata; absence cannot be established."
        if table in ("policy", "execution_controls") or table.endswith("_policy"):
            selected = [r for r in rows if table != "policy" or r.get("domain") in ("monthly", "financial")]
            entry["scope"] = "policy"
        else:
            selected = [r for r in rows if r.get("billing_period_start") == PERIOD
                        or (table == "monthly_candidate_executions" and r.get("id") in executions)
                        or r.get("execution_id") in executions]
            entry["scope"] = "December 2026"
            found = found or bool(selected)
            unresolved = unresolved or any(r.get("state") in ("sending", "unknown") for r in selected)
        entry["records"] = selected
        entry["unscoped_record_count"] = len(rows) - len(selected)
        if entry.get("status") == "ACCESSIBLE":
            entry["status"] = "RECORDS_FOUND" if selected else "ACCESSIBLE_EMPTY"
    complete = inventory_complete and all(e.get("complete", False) for e in entries.values())
    return {"period": "2026-12", "inventory_complete": complete,
            "december_records": "RECORDS_FOUND" if found else ("ACCESSIBLE_EMPTY" if complete else "NOT_VERIFIED"),
            "unresolved_send": "FOUND" if unresolved else ("NONE_FOUND" if complete else "NOT_VERIFIED"),
            "tables": entries}


def rest_inventory(client) -> dict:
    entries = {}
    for table, columns in TABLES.items():
        if table in ("execution_controls", "policy"):
            continue
        try:
            rows = client.table(table).select(columns).limit(LIMIT + 1).execute().data
            if not isinstance(rows, list):
                raise ValueError("Invalid metadata response")
            entries["public." + table] = {
                "status": "ACCESSIBLE", "complete": len(rows) <= LIMIT,
                "rows": [_metadata(r, table) for r in rows[:LIMIT]],
            }
        except Exception as exc:
            entries["public." + table] = _failure(exc, rest=True)
    result = _summarize(entries, inventory_complete=False)
    result["access"] = "REST_ONLY: non-public/unexposed schemas cannot be ruled out; "
    result["access"] += "UNAVAILABLE does not mean ABSENT. ACCESSIBLE_EMPTY refers only to the scoped result."
    return result


def sql_inventory(connection) -> dict:
    """Caller supplies a dedicated connection; all queries run in read-only transactions."""
    entries = {}
    try:
        connection.read_only = True
        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT n.nspname, c.relname, pg_catalog.row_security_active(c.oid) FROM pg_catalog.pg_class c "
                "JOIN pg_catalog.pg_namespace n ON n.oid=c.relnamespace "
                "WHERE c.relkind IN ('r','p','v','m') AND "
                "(c.relname = ANY(%s) OR c.relname ~ %s) "
                "AND n.nspname NOT LIKE 'pg_%%' AND n.nspname <> 'information_schema'",
                (list(TABLES), r"monthly.*(execution|document|claim|attempt|policy)"),
            )
            relations = cursor.fetchall()
        connection.rollback()
    except Exception as exc:
        connection.rollback()
        return {"period": "2026-12", "inventory_complete": False,
                "december_records": "NOT_VERIFIED", "unresolved_send": "NOT_VERIFIED",
                "catalog": _failure(exc), "tables": {}}
    from psycopg import sql

    for schema, table, rls_active in relations:
        name = f"{schema}.{table}"
        if table not in TABLES:
            entries[name] = {"status": "UNAVAILABLE", "reason": "Unrecognized metadata shape; operator review required."}
            continue
        try:
            with connection.cursor() as cursor:
                # Only allowlisted scalar metadata leaves the database, never PDF/envelope/payload/secrets.
                pairs = sql.SQL(",").join(sql.SQL("{} , to_jsonb(t)->{}").format(
                    sql.Literal(key), sql.Literal(key)) for key in TABLES[table].split(","))
                cursor.execute(sql.SQL("SELECT jsonb_build_object({}) FROM {}.{} t LIMIT %s").format(
                    pairs, sql.Identifier(schema), sql.Identifier(table)), (LIMIT + 1,))
                rows = [r[0] for r in cursor.fetchall()]
            entries[name] = {"status": "ACCESSIBLE", "complete": len(rows) <= LIMIT and not rls_active,
                             "row_security_active": rls_active,
                             "rows": [_metadata(r, table) for r in rows[:LIMIT]]}
        except Exception as exc:
            entries[name] = _failure(exc)
        finally:
            connection.rollback()
    for table in TABLES:
        if not any(t == table for _, t, _ in relations):
            entries["*." + table] = {"status": "ABSENT", "complete": True,
                                     "reason": "No matching relation in the catalog across inspected schemas."}
    result = _summarize(entries, inventory_complete=True)
    result["access"] = "Read-only SQL catalog inventory across schemas; scoped metadata only."
    return result


def database_snapshot() -> dict:
    # Optional existing operator-provided access; never derive a DSN from REST credentials.
    dsn = os.environ.get("DIAGNOSTICS_READONLY_DATABASE_URL", "").strip()
    if dsn:
        if importlib.util.find_spec("psycopg") is None:
            return {"period": "2026-12", "inventory_complete": False, "december_records": "NOT_VERIFIED",
                    "unresolved_send": "NOT_VERIFIED", "status": "UNAVAILABLE", "reason": "SQL driver not installed."}
        import psycopg
        try:
            with psycopg.connect(dsn, connect_timeout=5,
                                 options="-c default_transaction_read_only=on -c statement_timeout=5000") as connection:
                return sql_inventory(connection)
        except Exception as exc:
            return {"period": "2026-12", "inventory_complete": False, "december_records": "NOT_VERIFIED",
                    "unresolved_send": "NOT_VERIFIED", **_failure(exc)}
    from services.supabase_service import _get_client
    try:
        return rest_inventory(_get_client())
    except Exception as exc:
        return {"period": "2026-12", "inventory_complete": False, "december_records": "NOT_VERIFIED",
                "unresolved_send": "NOT_VERIFIED", **_failure(exc)}


def collect_diagnostics(role: str) -> dict:
    _admin(role)
    result = {"observed_at_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
              "build": _commit()}
    try:
        result["runtime"] = runtime_snapshot()
    except Exception as exc:
        result["runtime"] = _failure(exc)
    result["database"] = database_snapshot()
    return result


def render_diagnostics(ui, role: str) -> None:
    if role != "admin":
        return
    with ui.expander("Deployment diagnostics (sanitized, read-only)", expanded=False):
        ui.caption("December 2026 metadata only. No transport tests, history synchronization or business writes. "
                   "UNAVAILABLE/NOT_VERIFIED is not an empty database or an OFF gate.")
        if ui.button("Read sanitized runtime diagnostics", key="read_deployment_diagnostics"):
            ui.json(collect_diagnostics(role))
