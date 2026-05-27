param(
  [string]$ApiKey = "",
  [switch]$SkipLogin,
  [switch]$UseR2
)

$ErrorActionPreference = "Stop"
$project = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $project

function Run-Cmd {
  param([string[]]$ArgsList, [switch]$AllowFail)
  $output = & $ArgsList[0] @($ArgsList[1..($ArgsList.Count - 1)]) 2>&1
  $text = $output | Out-String
  if ($LASTEXITCODE -ne 0 -and -not $AllowFail) {
    throw $text
  }
  return $text
}

if (-not (Test-Path -LiteralPath "node_modules")) {
  Write-Host "Instalando dependencias..."
  Run-Cmd -ArgsList @("npm.cmd", "install") | Write-Host
}

if (-not $SkipLogin) {
  $who = Run-Cmd -ArgsList @("npx.cmd", "wrangler", "whoami") -AllowFail
  if ($who -match "not authenticated") {
    Write-Host "Abriendo login de Cloudflare..."
    Run-Cmd -ArgsList @("npx.cmd", "wrangler", "login") | Write-Host
  }
}

$dbId = ""
$createDb = Run-Cmd -ArgsList @("npx.cmd", "wrangler", "d1", "create", "mga-cloud-db") -AllowFail
if ($createDb -match 'database_id"\s*:\s*"([^"]+)"') {
  $dbId = $Matches[1]
} elseif ($createDb -match 'database_id\s*=\s*"([^"]+)"') {
  $dbId = $Matches[1]
} else {
  $listDb = Run-Cmd -ArgsList @("npx.cmd", "wrangler", "d1", "list") -AllowFail
  if ($listDb -match 'mga-cloud-db[^\r\n]*?([0-9a-fA-F-]{32,36})') {
    $dbId = $Matches[1]
  }
}

if (-not $dbId) {
  $dbId = Read-Host "Pega aqui el database_id de mga-cloud-db"
}

$configPath = Join-Path $project "wrangler.jsonc"
$config = Get-Content -LiteralPath $configPath -Raw
$config = $config -replace '"database_id"\s*:\s*"[^"]+"', ('"database_id": "' + $dbId + '"')
Set-Content -LiteralPath $configPath -Value $config -Encoding UTF8

Run-Cmd -ArgsList @("npx.cmd", "wrangler", "d1", "execute", "mga-cloud-db", "--remote", "--file=./schema.sql") | Write-Host
if ($UseR2) {
  Run-Cmd -ArgsList @("npx.cmd", "wrangler", "r2", "bucket", "create", "mga-evidencias") -AllowFail | Write-Host
}

if (-not $ApiKey) {
  $root = Split-Path -Parent $project
  $configCandidate = Join-Path $root "mobile_app\cloud_config.json"
  if (Test-Path -LiteralPath $configCandidate) {
    $mobileConfig = Get-Content -LiteralPath $configCandidate -Raw | ConvertFrom-Json
    $ApiKey = [string]($mobileConfig.apiKey)
  }
  if (-not $ApiKey) {
    $secure = Read-Host "Pega la API key cloud que usa el APK/programa" -AsSecureString
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
      $ApiKey = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    } finally {
      [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
  }
}

$ApiKey | & npx.cmd wrangler secret put MGA_API_KEY
if ($LASTEXITCODE -ne 0) { throw "No se pudo configurar MGA_API_KEY." }

$deploy = Run-Cmd -ArgsList @("npx.cmd", "wrangler", "deploy")
Write-Host $deploy
if ($deploy -match 'https://[^\s]+' ) {
  $url = $Matches[0].TrimEnd(".")
  Write-Host "Actualizando programa y APK a $url"
  powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "configurar_cloudflare.ps1") -Url $url -ApiKey $ApiKey
} else {
  Write-Host "No pude detectar la URL automaticamente. Ejecuta configurar_cloudflare.ps1 con la URL del deploy."
}
