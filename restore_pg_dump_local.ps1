param(
    [string]$DumpFile = "",
    [string]$DatabaseName = "caddycheck_restore_test",
    [string]$PgBin = "C:\Program Files\PostgreSQL\17\bin",
    [string]$HostName = "localhost",
    [int]$Port = 5432,
    [string]$UserName = "postgres",
    [switch]$FullRestore
)

$ErrorActionPreference = "Stop"

$PgRestore = Join-Path $PgBin "pg_restore.exe"
$Createdb = Join-Path $PgBin "createdb.exe"
$Dropdb = Join-Path $PgBin "dropdb.exe"
$Psql = Join-Path $PgBin "psql.exe"

foreach ($tool in @($PgRestore, $Createdb, $Dropdb, $Psql)) {
    if (-not (Test-Path $tool)) {
        throw "Required PostgreSQL tool not found: $tool"
    }
}

if ([string]::IsNullOrWhiteSpace($DumpFile)) {
    $latestDump = Get-ChildItem -Path "db_backups_pg" -Filter "*.dump" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $latestDump) {
        throw "No .dump files found in db_backups_pg. Pass -DumpFile explicitly."
    }
    $DumpFile = $latestDump.FullName
}

if (-not (Test-Path $DumpFile)) {
    throw "Dump file not found: $DumpFile"
}

Write-Host "This will recreate local database '$DatabaseName' from:" -ForegroundColor Yellow
Write-Host "  $DumpFile" -ForegroundColor Yellow
Write-Host "Target: $UserName@$HostName`:$Port/$DatabaseName" -ForegroundColor Yellow
if (-not $FullRestore) {
    Write-Host "Mode: public schema only (recommended local app-data test). Use -FullRestore for all schemas." -ForegroundColor Yellow
}

$securePassword = Read-Host "Local PostgreSQL password for user '$UserName'" -AsSecureString
$passwordPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
try {
    $env:PGPASSWORD = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPtr)

    & $Dropdb --if-exists --host $HostName --port $Port --username $UserName $DatabaseName
    if ($LASTEXITCODE -ne 0) { throw "dropdb failed with exit code $LASTEXITCODE" }

    & $Createdb --host $HostName --port $Port --username $UserName $DatabaseName
    if ($LASTEXITCODE -ne 0) { throw "createdb failed with exit code $LASTEXITCODE" }

    $restoreArgs = @(
        "--host", $HostName,
        "--port", $Port,
        "--username", $UserName,
        "--dbname", $DatabaseName,
        "--no-owner",
        "--no-privileges",
        "--verbose"
    )

    if (-not $FullRestore) {
        $restoreArgs += @("--schema", "public")
    }

    $restoreArgs += $DumpFile
    & $PgRestore @restoreArgs
    if ($LASTEXITCODE -ne 0) { throw "pg_restore failed with exit code $LASTEXITCODE" }

    Write-Host "\nRestore finished. Verifying public table counts..." -ForegroundColor Cyan
    & $Psql --host $HostName --port $Port --username $UserName --dbname $DatabaseName --command "select schemaname, relname as table_name, n_live_tup::bigint as estimated_rows from pg_stat_user_tables where schemaname = 'public' order by relname;"

    Write-Host "\nLocal restore test complete: $DatabaseName" -ForegroundColor Green
}
finally {
    $env:PGPASSWORD = $null
    if ($passwordPtr -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPtr)
    }
}
