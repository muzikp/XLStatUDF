param(
    [string]$CertDir = ".\.certs"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$certRoot = Join-Path $root $CertDir
New-Item -ItemType Directory -Force -Path $certRoot | Out-Null

$pfxPath = Join-Path $certRoot "localhost-devcert.pfx"
$cerPath = Join-Path $certRoot "localhost-devcert.cer"
$passPath = Join-Path $certRoot "localhost-devcert.pass.txt"

if (-not (Test-Path $passPath)) {
    $password = [Guid]::NewGuid().ToString("N")
    Set-Content -LiteralPath $passPath -Value $password -Encoding ascii
}

$helperProject = Join-Path $PSScriptRoot "devcert\DevCert.csproj"
$thumbprint = (& dotnet run --project $helperProject -- $pfxPath $cerPath $passPath)
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-Host "Local dev certificate ready:"
Write-Host "  Source: generated Evalytics .NET development certificate"
Write-Host "  Thumbprint: $thumbprint"
Write-Host "  PFX: $pfxPath"
Write-Host "  CER: $cerPath"
Write-Host "  Trusted in: Cert:\CurrentUser\Root"
