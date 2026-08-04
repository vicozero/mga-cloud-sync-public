param(
    [string]$BackupFile = "mga_cloud_sync_render_migration.backup",
    [switch]$SkipRestore,
    [switch]$SkipHealthCheck
)

$ErrorActionPreference = "Stop"

function Require-Command {
    param([string]$Name)
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if (-not $cmd) {
        throw "No se encontro '$Name'. Instala PostgreSQL Tools o agrega pg_dump/pg_restore al PATH."
    }
}

function Require-Env {
    param([string]$Name)
    $value = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "Falta variable de entorno $Name."
    }
    return $value
}

Require-Command "pg_dump"
Require-Command "pg_restore"

$oldDatabaseUrl = Require-Env "OLD_DATABASE_URL"
$newDatabaseUrl = Require-Env "NEW_DATABASE_URL"
$newServiceUrl = [Environment]::GetEnvironmentVariable("NEW_RENDER_URL")
$apiKey = [Environment]::GetEnvironmentVariable("MGA_API_KEY")

$backupPath = [System.IO.Path]::GetFullPath($BackupFile)

Write-Host "Creando respaldo logico desde la base actual..."
pg_dump --format=custom --no-owner --no-privileges --file "$backupPath" "$oldDatabaseUrl"

if (-not (Test-Path $backupPath)) {
    throw "No se genero el respaldo: $backupPath"
}

$sizeMb = [math]::Round((Get-Item $backupPath).Length / 1MB, 2)
Write-Host "Respaldo creado: $backupPath ($sizeMb MB)"

if (-not $SkipRestore) {
    Write-Host "Restaurando respaldo en la base nueva..."
    pg_restore --clean --if-exists --no-owner --no-privileges --dbname "$newDatabaseUrl" "$backupPath"
    Write-Host "Restauracion terminada."
}

if (-not $SkipHealthCheck -and -not [string]::IsNullOrWhiteSpace($newServiceUrl)) {
    $baseUrl = $newServiceUrl.TrimEnd("/")
    Write-Host "Probando health check..."
    $health = Invoke-RestMethod -Uri "$baseUrl/health" -Method Get -TimeoutSec 60
    $health | ConvertTo-Json -Depth 10

    if (-not [string]::IsNullOrWhiteSpace($apiKey)) {
        Write-Host "Probando acceso a portal con API key..."
        $headers = @{"X-MGA-API-Key" = $apiKey}
        $portal = Invoke-RestMethod -Uri "$baseUrl/api/portal" -Headers $headers -Method Get -TimeoutSec 120
        $counts = [ordered]@{
            ok = $true
            equipment = @($portal.equipment).Count
            captures = @($portal.captures).Count
            preventive_records = @($portal.preventive_records).Count
            tire_records = @($portal.tire_records).Count
        }
        $counts | ConvertTo-Json -Depth 5
    }
}

Write-Host "Migracion finalizada."
