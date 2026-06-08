param(
    [string]$OutputDir = "db_backups_pg",
    [string]$PgDumpPath = "C:\Program Files\PostgreSQL\17\bin\pg_dump.exe",
    [string]$HostName = "aws-1-ap-southeast-1.pooler.supabase.com",
    [int]$Port = 5432,
    [string]$Database = "postgres",
    [string]$UserName = "postgres.rdoxihpmghrvroddnkdi",
    [switch]$PromptConnectionDetails
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $PgDumpPath)) {
    throw "pg_dump was not found at '$PgDumpPath'. Update -PgDumpPath or reinstall PostgreSQL client tools."
}

if ($PromptConnectionDetails) {
    Write-Host "Copy these values from Supabase Connect -> Session pooler." -ForegroundColor Cyan
    Write-Host "Leave a field empty to keep the value shown in brackets." -ForegroundColor Cyan

    $hostInput = Read-Host "Pooler host [$HostName]"
    if (-not [string]::IsNullOrWhiteSpace($hostInput)) {
        $HostName = $hostInput.Trim()
    }

    $portInput = Read-Host "Pooler port [$Port]"
    if (-not [string]::IsNullOrWhiteSpace($portInput)) {
        $Port = [int]$portInput.Trim()
    }

    $userInput = Read-Host "Database user [$UserName]"
    if (-not [string]::IsNullOrWhiteSpace($userInput)) {
        $UserName = $userInput.Trim()
    }

    $dbInput = Read-Host "Database name [$Database]"
    if (-not [string]::IsNullOrWhiteSpace($dbInput)) {
        $Database = $dbInput.Trim()
    }
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $OutputDir "caddycheck_supabase_$timestamp.dump"

$securePassword = Read-Host "Supabase database password" -AsSecureString
$passwordPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
try {
    $env:PGPASSWORD = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPtr)
    Write-Host "Running pg_dump against $UserName@$HostName`:$Port/$Database" -ForegroundColor Cyan
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
