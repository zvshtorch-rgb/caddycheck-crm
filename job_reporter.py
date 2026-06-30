"""
Video Profiler job reporter (runs on each project PC).

Counts the jobs in the local SQL Server ``VideoProfilerDatabase.dbo.Jobs`` table
and pushes the per-machine totals to the central CaddyCheck CRM (Supabase), where
they are compared against the approved camera quantity for the project.

License model: one license per PC, but each PC runs N jobs (≈ cameras). This agent
reports how many jobs are active so the CRM can flag projects that exceed the
quantity approved in their purchase order.

Deploy on each PC (TeamViewer) and run on a schedule (Windows Task Scheduler),
e.g. every 30 minutes:

    py job_reporter.py

Required environment variables (set once per PC, e.g. in a .bat wrapper):
    SUPABASE_URL                 https://<project>.supabase.co
    SUPABASE_SERVICE_ROLE_KEY    <service role key>
Optional:
    PROJECT_NAME                 CRM project name this PC belongs to (recommended)
    SQL_CONNECTION_STRING        full ODBC connection string (overrides defaults)
    SQL_SERVER                   default: localhost\\SQLEXPRESS
    SQL_DATABASE                 default: VideoProfilerDatabase
"""
from __future__ import annotations

import logging
import os
import socket
import sys

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("job_reporter")

AGENT_VERSION = "1.0.0"


def _sql_connection_string() -> str:
    explicit = os.environ.get("SQL_CONNECTION_STRING", "").strip()
    if explicit:
        return explicit
    server = os.environ.get("SQL_SERVER", r"localhost\SQLEXPRESS").strip()
    database = os.environ.get("SQL_DATABASE", "VideoProfilerDatabase").strip()
    # Trusted (Windows) auth — the agent runs as a local user with DB access.
    return (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        f"SERVER={server};DATABASE={database};Trusted_Connection=yes;"
    )


def _count_jobs() -> tuple[int, int, str]:
    """Return (active_jobs, total_jobs, owner) from the local Jobs table.

    Active = jobs that have not completed (CompletedTime IS NULL).
    """
    import pyodbc

    conn = pyodbc.connect(_sql_connection_string(), timeout=15)
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM dbo.Jobs")
        total_jobs = int(cursor.fetchone()[0] or 0)

        cursor.execute("SELECT COUNT(*) FROM dbo.Jobs WHERE CompletedTime IS NULL")
        active_jobs = int(cursor.fetchone()[0] or 0)

        cursor.execute(
            "SELECT TOP 1 Owner FROM dbo.Jobs WHERE Owner IS NOT NULL ORDER BY Id DESC"
        )
        row = cursor.fetchone()
        owner = str(row[0]).strip() if row and row[0] is not None else ""
        return active_jobs, total_jobs, owner
    finally:
        conn.close()


def _supabase_config() -> tuple[str, str]:
    url = os.environ.get("SUPABASE_URL", "").strip()
    key = (
        os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
        or os.environ.get("SUPABASE_KEY", "").strip()
    )
    if not url or not key:
        raise RuntimeError(
            "Supabase credentials not configured. Set SUPABASE_URL and "
            "SUPABASE_SERVICE_ROLE_KEY environment variables."
        )
    return url.rstrip("/"), key


def _report(active_jobs: int, total_jobs: int, owner: str) -> None:
    import datetime

    import requests

    url, key = _supabase_config()
    machine_name = socket.gethostname()
    project_name = os.environ.get("PROJECT_NAME", "").strip() or None

    payload = {
        "machine_name": machine_name,
        "project_name": project_name,
        "owner": owner or None,
        "active_jobs": active_jobs,
        "total_jobs": total_jobs,
        "agent_version": AGENT_VERSION,
        "reported_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    }

    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        # Upsert on the unique machine_name so each PC keeps one current row.
        "Prefer": "resolution=merge-duplicates,return=representation",
    }

    resp = requests.post(
        f"{url}/rest/v1/project_job_status?on_conflict=machine_name",
        json=payload,
        headers=headers,
        timeout=30,
    )
    resp.raise_for_status()
    logger.info(
        "Reported %s: active=%d total=%d owner=%s",
        machine_name, active_jobs, total_jobs, owner or "(none)",
    )


def main() -> int:
    try:
        active_jobs, total_jobs, owner = _count_jobs()
    except Exception as exc:  # noqa: BLE001
        logger.error("Failed to read local SQL Server: %s", exc)
        return 1

    try:
        _report(active_jobs, total_jobs, owner)
    except Exception as exc:  # noqa: BLE001
        logger.error("Failed to push report to Supabase: %s", exc)
        return 2

    logger.info("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
