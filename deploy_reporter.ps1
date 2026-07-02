<#
.SYNOPSIS
    One-shot deployer for the CaddyCheck job reporter agent.

.DESCRIPTION
    Run on each project PC (via TeamViewer remote control) as Administrator.
    Does everything:
      1. Installs Python 3.12 if missing
      2. Installs pip packages (requests, pyodbc)
      3. Creates C:\CaddyCheck and downloads job_reporter.py from GitHub
      4. Writes a minimal run_reporter.bat
      5. Sets system-wide environment variables (no credentials in files)
      6. Creates two Task Scheduler tasks: 1st and 15th of every month at 08:00

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File deploy_reporter.ps1 -ProjectName "AD Aalter"

.PARAMETER ProjectName
    The CRM project name for this PC (must match exactly as shown in CaddyCheck CRM).
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ProjectName = ""   # Optional: auto-discovered from machines.csv in GitHub
)

# Check for admin rights (PS 3.0 compatible - must be AFTER param block)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Error "This script must be run as Administrator. Right-click PowerShell and select 'Run as Administrator'."
    exit 1
}

# -- Configuration (edit once, deploy everywhere) -----------------------------
$SUPABASE_URL  = "https://rdoxihpmghrvroddnkdi.supabase.co"
$SUPABASE_KEY  = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJkb3hpaHBtZ2hydnJvZGRua2RpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzM2NDUxMTcsImV4cCI6MjA4OTIyMTExN30.IFy2YajxTpvwTqFjmDkB6liGQzahCccUsY1Y28LHCvM"   # anon key only - safe to embed
$INSTALL_DIR   = "C:\CaddyCheck"
$TASK_NAME_1   = "CaddyCheck Reporter (1st)"
$TASK_NAME_2   = "CaddyCheck Reporter (15th)"
$RUN_TIME      = "08:00"
# -----------------------------------------------------------------------------

# Force TLS 1.2 - required for GitHub on older Windows (pre-2016) systems
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ErrorActionPreference = "Stop"

function Write-Step([string]$msg) {
    Write-Host "`n==> $msg" -ForegroundColor Cyan
}

# -- Step 1: Python -----------------------------------------------------------
Write-Step "Checking Python..."
$pyExe = $null
foreach ($candidate in @("py", "python", "python3")) {
    try {
        $ver = & $candidate --version 2>&1
        if ($ver -match "Python 3") { $pyExe = $candidate; break }
    } catch {}
}

if (-not $pyExe) {
    Write-Step "Python not found. Downloading Python 3.12.4..."
    $installer = "$env:TEMP\python-3.12.4-amd64.exe"
    Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe" `
                      -OutFile $installer -UseBasicParsing
    Write-Host "Installing Python (silent)..."
    Start-Process -FilePath $installer `
                  -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_launcher=1" `
                  -Wait
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    $pyExe = "py"
    Write-Host "Python installed." -ForegroundColor Green
} else {
    Write-Host "Found: $pyExe" -ForegroundColor Green
}

$pythonFullPath = (& $pyExe -c "import sys; print(sys.executable)" 2>&1).Trim()
if (-not (Test-Path $pythonFullPath)) {
    $pythonFullPath = Get-ChildItem "C:\Python*","C:\Users\*\AppData\Local\Programs\Python\Python*" `
        -Filter python.exe -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
}
Write-Host "Python exe: $pythonFullPath"

# -- Step 2: pip packages -----------------------------------------------------
Write-Step "Installing pip packages..."
& $pyExe -m pip install requests pyodbc --quiet --disable-pip-version-check
Write-Host "packages OK" -ForegroundColor Green

# -- Step 2b: ODBC Driver for SQL Server (required by job_reporter.py) --------
Write-Step "Checking ODBC Driver for SQL Server..."
# Ask Python/pyodbc directly -- more reliable than registry check
$odbcCheck = & $pyExe -c "import pyodbc; ok = any('SQL Server' in d for d in pyodbc.drivers()); exit(0 if ok else 1)" 2>$null
$odbcOk = ($LASTEXITCODE -eq 0)
if ($odbcOk) {
    Write-Host "ODBC driver already installed and visible to Python." -ForegroundColor Green
} else {
    Write-Host "Installing ODBC Driver 17 for SQL Server..." -ForegroundColor Yellow
    $odbcMsi = "$env:TEMP\msodbcsql17.msi"
    try {
        Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?linkid=2168524" `
            -OutFile $odbcMsi -UseBasicParsing
        $result = Start-Process msiexec.exe -Wait -PassThru `
            -ArgumentList "/i `"$odbcMsi`" /quiet /norestart IACCEPTMSODBCSQLLICENSETERMS=YES"
        if ($result.ExitCode -eq 0) {
            Write-Host "ODBC Driver 17 installed." -ForegroundColor Green
        } else {
            Write-Warning "ODBC installer exited with code $($result.ExitCode) - may need manual install."
        }
    } catch {
        Write-Warning "Could not auto-install ODBC driver: $_ - install manually if SQL errors occur."
    }
}
Write-Step "Creating $INSTALL_DIR ..."
New-Item -ItemType Directory -Force -Path $INSTALL_DIR | Out-Null

Write-Step "Writing job_reporter.py (embedded)..."
$jobReporterContent = @'
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

'@
[System.IO.File]::WriteAllText("$INSTALL_DIR\job_reporter.py", $jobReporterContent, [System.Text.Encoding]::ASCII)
Write-Host "job_reporter.py written." -ForegroundColor Green

# -- Step 4: bat file with embedded credentials (anon key is public) ----------
# Embedding the config in the bat avoids all registry/session propagation
# problems - both manual tests and the SYSTEM scheduled task read it reliably.
Write-Step "Writing run_reporter.bat..."
# If -ProjectName wasn't passed, pick up a PROJECT_NAME that was set manually
# in the machine registry before running this script.
if (-not $ProjectName) {
    $ProjectName = [System.Environment]::GetEnvironmentVariable("PROJECT_NAME", "Machine")
}
$projLine = if ($ProjectName) { "set `"PROJECT_NAME=$ProjectName`"`r`n" } else { "" }
$batContent = "@echo off`r`n" +
              "set `"SUPABASE_URL=$SUPABASE_URL`"`r`n" +
              "set `"SUPABASE_KEY=$SUPABASE_KEY`"`r`n" +
              $projLine +
              "`"$pythonFullPath`" `"$INSTALL_DIR\job_reporter.py`"`r`n"
[System.IO.File]::WriteAllText("$INSTALL_DIR\run_reporter.bat", $batContent, [System.Text.Encoding]::ASCII)
Write-Host "bat file written." -ForegroundColor Green

# -- Step 5: System environment variables (registry, best-effort backup) ------
# Credentials are already embedded in the bat above; also mirror them to the
# machine registry so future runs / other tools can read them too.
Write-Step "Setting system environment variables..."
try {
    [System.Environment]::SetEnvironmentVariable("SUPABASE_URL",  $SUPABASE_URL, "Machine")
    [System.Environment]::SetEnvironmentVariable("SUPABASE_KEY",  $SUPABASE_KEY, "Machine")
    if ($ProjectName) {
        [System.Environment]::SetEnvironmentVariable("PROJECT_NAME", $ProjectName, "Machine")
        Write-Host "PROJECT_NAME set to: $ProjectName" -ForegroundColor Green
    } else {
        Write-Host "PROJECT_NAME not set - will be auto-discovered from machines.csv" -ForegroundColor Yellow
    }
    Write-Host "Env vars set in registry." -ForegroundColor Green
} catch {
    Write-Warning "Could not set machine env vars (bat file has them embedded anyway): $_"
}

# -- Step 6: Task Scheduler (1st and 15th of every month) ---------------------
Write-Step "Creating scheduled tasks..."
$batPath = "`"$INSTALL_DIR\run_reporter.bat`""

foreach ($day in @(1, 15)) {
    $taskName = if ($day -eq 1) { $TASK_NAME_1 } else { $TASK_NAME_2 }
    try { schtasks /Delete /TN $taskName /F 2>&1 | Out-Null } catch { }
    $result = schtasks /Create `
        /TN $taskName `
        /TR $batPath `
        /SC MONTHLY /D $day `
        /ST $RUN_TIME `
        /RU "SYSTEM" `
        /RL HIGHEST `
        /F 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Task '$taskName' created." -ForegroundColor Green
    } else {
        Write-Warning "Task '$taskName' may have failed: $result"
    }
}

# -- Done ---------------------------------------------------------------------
Write-Host "`n===========================================" -ForegroundColor Green
Write-Host " CaddyCheck reporter deployed successfully!" -ForegroundColor Green
Write-Host "   PC:      $env:COMPUTERNAME"               -ForegroundColor Green
Write-Host "   Project: $(if ($ProjectName) { $ProjectName } else { '(from machines.csv)' })" -ForegroundColor Green
Write-Host "   Runs:    1st and 15th of every month at $RUN_TIME"  -ForegroundColor Green
Write-Host "===========================================" -ForegroundColor Green
Write-Host "`nRun now to test:"
Write-Host "   & '$INSTALL_DIR\run_reporter.bat'" -ForegroundColor Yellow