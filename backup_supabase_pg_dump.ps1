param(
    [string]$OutputDir = "db_backups_pg",
    [string]$PgDumpPath = "C:\Program Files\PostgreSQL\17\bin\pg_dump.exe",
    [string]$HostName = "aws-0-ap-southeast-1.pooler.supabase.com",
    [int]$Port = 5432,
    [string]$Database = "postgres",
    [string]$UserName = "postgres.rdoxihpmghrvroddnkdi"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $PgDumpPath)) {
    throw "pg_dump was not found at '$PgDumpPath'. Update -PgDumpPath or reinstall PostgreSQL client tools."
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $OutputDir "caddycheck_supabase_$timestamp.dump"

$securePassword = Read-Host "Supabase database password" -AsSecureString
$passwordPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
try {
    $env:PGPASSWORD = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPtr)
    & $PgDumpPath `
        --host $HostName `
        --port $Port `
        --username $UserName `
        --dbname $Database `
        --format custom `
        --file $outputFile `
        --verbose

    if ($LASTEXITCODE -ne 0) {
        throw "pg_dump failed with exit code $LASTEXITCODE"
    }

    Write-Host "Backup complete: $outputFile" -ForegroundColor Green
}
finally {
    $env:PGPASSWORD = $null
    if ($passwordPtr -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPtr)
    }
}
