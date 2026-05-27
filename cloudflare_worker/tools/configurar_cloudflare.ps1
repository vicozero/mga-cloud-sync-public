param(
  [Parameter(Mandatory = $true)]
  [string]$Url,
  [string]$ApiKey = ""
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$cleanUrl = $Url.Trim().TrimEnd("/")
if (-not ($cleanUrl -match "^https://")) {
  throw "La URL debe iniciar con https://"
}

function Update-JsonFile {
  param([string]$Path, [scriptblock]$Mutate)
  if (-not (Test-Path -LiteralPath $Path)) { return }
  $json = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
  & $Mutate $json
  $json | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $Path -Encoding UTF8
}

Update-JsonFile -Path (Join-Path $root "mobile_app\cloud_config.json") -Mutate {
  param($json)
  $json.cloudUrl = $cleanUrl
  if ($ApiKey) { $json.apiKey = $ApiKey }
}

Update-JsonFile -Path (Join-Path $root "android\app\src\main\assets\public\cloud_config.json") -Mutate {
  param($json)
  $json.cloudUrl = $cleanUrl
  if ($ApiKey) { $json.apiKey = $ApiKey }
}

Update-JsonFile -Path (Join-Path $root "mga_config.json") -Mutate {
  param($json)
  $json.cloud_url = $cleanUrl
  if ($ApiKey) { $json.cloud_api_key = $ApiKey }
}

Write-Host "Configuracion actualizada a $cleanUrl"
