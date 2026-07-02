"""
Video Profiler job reporter (runs on each project PC).

Counts the jobs in the local SQL Server ``VideoProfilerDatabase.dbo.Jobs`` table
and pushes the per-machine totals to the central CaddyCheck CRM (Supabase), where
they are compared against the approved camera quantity for the project.

License model: one license per PC, but each PC runs N jobs (~= cameras). This agent
reports how many jobs are active so the CRM can flag projects that exceed the
quantity approved in their purchase order.

Deploy on each PC (TeamViewer) and run on a schedule (Windows Task Scheduler),
e.g. every 30 minutes:

    py job_reporter.py

Required environment variables (set once per PC via setx or deploy_reporter.ps1):
    SUPABASE_URL                 https://<project>.supabase.co
    SUPABASE_KEY                 Supabase anon key (read-only for other tables via RLS)
Optional:
    PROJECT_NAME                 Override project name (auto-discovered from machines.csv if absent)
    SQL_CONNECTION_STRING        full ODBC connection string (overrides defaults)
    SQL_SERVER                   default: localhost\\SQLEXPRESS
    SQL_DATABASE                 default: VideoProfilerDatabase

Project name auto-discovery:
    The script fetches machines.csv from GitHub and maps this PC's hostname to a
    CRM project name automatically, so the same script can be deployed to all PCs
    without any per-PC configuration beyond the Supabase credentials.
"""
from __future__ import annotations

import logging
import os
import socket
import sys
import uuid

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("job_reporter")

AGENT_VERSION = "1.0.0"

# machines.csv is fetched from GitHub so the same script runs on every PC.
MACHINES_CSV_URL = (
    "https://raw.githubusercontent.com/zvshtorch-rgb/caddycheck-crm/main/machines.csv"
)

# Cloudflare IPs for Supabase - used as DNS fallback on PCs with broken DNS.
_SUPABASE_FALLBACK_IPS = {
    "rdoxihpmghrvroddnkdi.supabase.co": "104.18.38.10",
}


def _patch_dns_if_broken() -> None:
    """Monkey-patch socket.getaddrinfo for known hosts when OS DNS is broken.

    Some project PCs have no working external DNS (broken Dnscache service,
    missing gateway, Group Policy NRPT overrides). This patches Python's
    resolver at the socket level so HTTPS still works regardless of OS state.
    The TLS SNI/cert verification continues to use the correct hostname.
    """
    for host, ip in _SUPABASE_FALLBACK_IPS.items():
        try:
            socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)
            return  # DNS works - no patch needed
        except socket.gaierror:
            pass

    _orig = socket.getaddrinfo

    def _patched(host, port, *args, **kwargs):
        if host in _SUPABASE_FALLBACK_IPS:
            host = _SUPABASE_FALLBACK_IPS[host]
        return _orig(host, port, *args, **kwargs)

    socket.getaddrinfo = _patched
    logger.info("DNS unavailable - using hardcoded IPs for Supabase endpoints.")


def _apply_system_proxy():
    """Apply system proxy settings for Supabase HTTPS requests.

    Uses .NET GetSystemWebProxy via PowerShell subprocess to resolve proxy
    settings including PAC/WPAD scripts that Python cannot natively execute.
    Falls back to WinHTTP netsh detection if PowerShell call fails.
    """
    if os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY"):
        return
    # Primary: use .NET GetSystemWebProxy (handles PAC/WPAD, IE, WinHTTP)
    try:
        import subprocess
        script = (
            "$proxy=[System.Net.WebRequest]::GetSystemWebProxy();"
            "$uri=[Uri]'https://rdoxihpmghrvroddnkdi.supabase.co';"
            "$p=$proxy.GetProxy($uri);"
            "if($p.Host -ne $uri.Host){$p.AbsoluteUri}else{''}"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True, text=True, timeout=10,
        )
        proxy = result.stdout.strip()
        if proxy and proxy.startswith("http"):
            os.environ["HTTPS_PROXY"] = proxy
            os.environ["HTTP_PROXY"] = proxy
            logger.info("System proxy applied: %s", proxy)
            return
    except Exception as exc:
        logger.debug("GetSystemWebProxy failed: %s", exc)
    # Fallback: WinHTTP netsh
    try:
        import subprocess
        out = subprocess.run(
            ["netsh", "winhttp", "show", "proxy"],
            capture_output=True, text=True, timeout=5,
        ).stdout
        for line in out.splitlines():
            if "proxy server" in line.lower():
                val = line.split(":", 1)[-1].strip()
                if val and "direct" not in val.lower() and "none" not in val.lower():
                    if not val.lower().startswith("http"):
                        val = "http://" + val
                    os.environ["HTTPS_PROXY"] = val
                    os.environ["HTTP_PROXY"] = val
                    logger.info("WinHTTP proxy applied: %s", val)
                    return
    except Exception as exc:
        logger.debug("WinHTTP proxy detection failed: %s", exc)


def _lookup_project_name(machine_name: str) -> str | None:
    """Return the CRM project name for this PC from the hosted machines.csv.

    Falls back to the PROJECT_NAME env var, then returns None (CRM will show
    the PC in the 'unmapped' expander until the CSV is updated).
    """
    # Env var always wins -- useful for overrides / testing.
    override = os.environ.get("PROJECT_NAME", "").strip()
    if override:
        return override

    try:
        import urllib.request

        with urllib.request.urlopen(MACHINES_CSV_URL, timeout=10) as resp:
            lines = resp.read().decode("utf-8").splitlines()
    except Exception as exc:
        logger.warning("Could not fetch machines.csv: %s -- using PROJECT_NAME env var", exc)
        return None

    machine_lower = machine_name.lower()
    for line in lines[1:]:  # skip header
        parts = line.strip().split(",", 1)
        if len(parts) == 2 and parts[0].strip().lower() == machine_lower:
            return parts[1].strip()

    logger.warning("Machine '%s' not found in machines.csv -- add it to the repo.", machine_name)
    return None


def _count_jobs() -> tuple[int, int, str]:
    """Return (active_jobs, total_jobs, owner) from the local Video Profiler DB.

    Auto-detects two known schemas:
      VideoProfilerDatabase / dbo.Jobs      -> Status = 1 (integer)
      VideoInformDB         / dbo.FileJobs  -> Status = 'Running' (text)

    Env var overrides: SQL_CONNECTION_STRING, SQL_SERVER, SQL_DATABASE, SQL_TABLE.
    """
    import pyodbc

    explicit_conn = os.environ.get("SQL_CONNECTION_STRING", "").strip()
    server = os.environ.get("SQL_SERVER", r"localhost\SQLEXPRESS").strip()
    explicit_db = os.environ.get("SQL_DATABASE", "").strip()
    explicit_table = os.environ.get("SQL_TABLE", "").strip()

    DRIVERS = (
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 18 for SQL Server",
        "SQL Server",          # Built-in MDAC driver, always present on Windows
    )

    # (database, table, active_where, owner_col, id_col)
    SCHEMAS = [
        ("VideoProfilerDatabase", "dbo.Jobs",     "Status = 1",          "Owner", "Id"),
        ("VideoInformDB",         "dbo.FileJobs",  "Status = 'Running'",  "Owner", "JobID"),
    ]

    if explicit_conn:
        # User provided a full connection string -- use it with default schema
        conn = pyodbc.connect(explicit_conn, timeout=15)
        db, table, active_where, owner_col, id_col = SCHEMAS[0]
        if explicit_table:
            table = explicit_table
        return _query_jobs(conn, table, active_where, owner_col, id_col)

    last_err = None
    for db, table, active_where, owner_col, id_col in SCHEMAS:
        if explicit_db:
            db = explicit_db
        if explicit_table:
            table = explicit_table

        for driver in DRIVERS:
            conn_str = (
                f"DRIVER={{{driver}}};"
                f"SERVER={server};DATABASE={db};Trusted_Connection=yes;"
            )
            try:
                conn = pyodbc.connect(conn_str, timeout=5)
                result = _query_jobs(conn, table, active_where, owner_col, id_col)
                logger.info("Connected to %s / %s via %s", db, table, driver)
                return result
            except Exception as exc:
                last_err = exc

        if explicit_db:
            break  # Don't iterate schemas when DB is forced

    raise last_err or RuntimeError("Could not connect to any known Video Profiler database.")


def _query_jobs(conn, table, active_where, owner_col, id_col) -> tuple[int, int, str]:
    """Execute job count queries on an open connection and return results."""
    try:
        cursor = conn.cursor()
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        total_jobs = int(cursor.fetchone()[0] or 0)

        cursor.execute(f"SELECT COUNT(*) FROM {table} WHERE {active_where}")
        active_jobs = int(cursor.fetchone()[0] or 0)

        cursor.execute(
            f"SELECT TOP 1 {owner_col} FROM {table} "
            f"WHERE {owner_col} IS NOT NULL ORDER BY {id_col} DESC"
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
    machine_name = "{}-{}".format(socket.gethostname(), format(uuid.getnode(), "012x"))
    project_name = _lookup_project_name(machine_name)

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
    _patch_dns_if_broken()
    _apply_system_proxy()
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
