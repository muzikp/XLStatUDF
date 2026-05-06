param(
    [string]$ManifestPath = ".\dist\manifest.local.xml",
    [string]$BaseUrl = "https://localhost:3000"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$source = Join-Path $root $ManifestPath
if (-not (Test-Path $source)) {
    & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "generate-local-manifest.ps1") -BaseUrl $BaseUrl
}

if (-not (Test-Path $source)) {
    throw "Manifest not found after generation attempt: $source"
}

$targetDir = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
New-Item -ItemType Directory -Force -Path $targetDir | Out-Null

$target = Join-Path $targetDir "Evalytics.StatLab.local.xml"
Copy-Item -LiteralPath $source -Destination $target -Force

Write-Host "Sideloaded manifest to $target"
Write-Host "Restart Excel if it was already open."
Write-Host "Note: Desktop Excel usually needs a trusted shared-folder catalog. Run npm run desktop:catalog if the add-in does not appear."
