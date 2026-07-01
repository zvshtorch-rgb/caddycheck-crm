#Requires -RunAsAdministrator
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

# ── Configuration (edit once, deploy everywhere) ─────────────────────────────
$SUPABASE_URL  = "https://rdoxihpmghrvroddnkdi.supabase.co"
$SUPABASE_KEY  = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJkb3hpaHBtZ2hydnJvZGRua2RpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzM2NDUxMTcsImV4cCI6MjA4OTIyMTExN30.IFy2YajxTpvwTqFjmDkB6liGQzahCccUsY1Y28LHCvM"   # anon key only – safe to embed
$INSTALL_DIR   = "C:\CaddyCheck"
$GITHUB_RAW    = "https://raw.githubusercontent.com/zvshtorch-rgb/caddycheck-crm/main/job_reporter.py"
$TASK_NAME_1   = "CaddyCheck Reporter (1st)"
$TASK_NAME_2   = "CaddyCheck Reporter (15th)"
$RUN_TIME      = "08:00"
# ─────────────────────────────────────────────────────────────────────────────

$ErrorActionPreference = "Stop"

function Write-Step([string]$msg) {
    Write-Host "`n==> $msg" -ForegroundColor Cyan
}

# ── Step 1: Python ────────────────────────────────────────────────────────────
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
    # Refresh PATH in this session
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    $pyExe = "py"
    Write-Host "Python installed." -ForegroundColor Green
} else {
    Write-Host "Found: $pyExe" -ForegroundColor Green
}

# Resolve full path to python.exe for the scheduled task (PATH may differ for SYSTEM)
$pythonFullPath = (& $pyExe -c "import sys; print(sys.executable)" 2>&1).Trim()
if (-not (Test-Path $pythonFullPath)) {
    # Fallback: search common locations
    $pythonFullPath = Get-ChildItem "C:\Python*","C:\Users\*\AppData\Local\Programs\Python\Python*" `
        -Filter python.exe -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
}
Write-Host "Python exe: $pythonFullPath"

# ── Step 2: pip packages ──────────────────────────────────────────────────────
Write-Step "Installing pip packages..."
& $pyExe -m pip install requests pyodbc --quiet --disable-pip-version-check
Write-Host "packages OK" -ForegroundColor Green

# ── Step 3: Install directory + job_reporter.py ───────────────────────────────
Write-Step "Creating $INSTALL_DIR ..."
New-Item -ItemType Directory -Force -Path $INSTALL_DIR | Out-Null

Write-Step "Downloading job_reporter.py from GitHub..."
Invoke-WebRequest -Uri $GITHUB_RAW -OutFile "$INSTALL_DIR\job_reporter.py" -UseBasicParsing
Write-Host "job_reporter.py downloaded." -ForegroundColor Green

# ── Step 4: Minimal bat file (no credentials) ─────────────────────────────────
Write-Step "Writing run_reporter.bat..."
$batContent = "@echo off`r`n`"$pythonFullPath`" `"$INSTALL_DIR\job_reporter.py`"`r`n"
[System.IO.File]::WriteAllText("$INSTALL_DIR\run_reporter.bat", $batContent, [System.Text.Encoding]::ASCII)
Write-Host "bat file written." -ForegroundColor Green

# ── Step 5: System environment variables (registry – not in any file) ─────────
Write-Step "Setting system environment variables..."
[System.Environment]::SetEnvironmentVariable("SUPABASE_URL",  $SUPABASE_URL, "Machine")
[System.Environment]::SetEnvironmentVariable("SUPABASE_KEY",  $SUPABASE_KEY, "Machine")
if ($ProjectName) {
    [System.Environment]::SetEnvironmentVariable("PROJECT_NAME", $ProjectName, "Machine")
    Write-Host "PROJECT_NAME set to: $ProjectName" -ForegroundColor Green
} else {
    Write-Host "PROJECT_NAME not set — will be auto-discovered from machines.csv" -ForegroundColor Yellow
}
Write-Host "Env vars set in registry." -ForegroundColor Green

# ── Step 6: Task Scheduler (1st and 15th of every month) ─────────────────────
Write-Step "Creating scheduled tasks..."
$batPath = "`"$INSTALL_DIR\run_reporter.bat`""

foreach ($day in @(1, 15)) {
    $taskName = if ($day -eq 1) { $TASK_NAME_1 } else { $TASK_NAME_2 }
    # Delete existing task silently before recreating
    schtasks /Delete /TN $taskName /F 2>$null | Out-Null
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

# ── Done ──────────────────────────────────────────────────────────────────────
Write-Host "`n===========================================" -ForegroundColor Green
Write-Host " CaddyCheck reporter deployed successfully!" -ForegroundColor Green
Write-Host "   PC:      $env:COMPUTERNAME"               -ForegroundColor Green
Write-Host "   Project: $(if ($ProjectName) { $ProjectName } else { '(from machines.csv)' })" -ForegroundColor Green
Write-Host "   Runs:    1st and 15th of every month at $RUN_TIME" -ForegroundColor Green
Write-Host "===========================================" -ForegroundColor Green
Write-Host "`nRun now to test:"
Write-Host "   & '$INSTALL_DIR\run_reporter.bat'" -ForegroundColor Yellow
