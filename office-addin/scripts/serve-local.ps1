param(
    [int]$Port = 3000
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot

$existingServer = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if ($existingServer) {
    $process = Get-Process -Id $existingServer.OwningProcess -ErrorAction SilentlyContinue
    if ($process -and $process.ProcessName -eq "node") {
        Write-Host "Stopping existing local Evalytics server on port $Port (PID $($process.Id))."
        Stop-Process -Id $process.Id -Force
        Start-Sleep -Milliseconds 500
    }
    else {
        throw "Port $Port is already in use by PID $($existingServer.OwningProcess). Stop that process or choose another port."
    }
}

& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "ensure-dev-cert.ps1")
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Push-Location $root
try {
    npm run build
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "generate-local-manifest.ps1") -BaseUrl "https://localhost:$Port"
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "install-desktop-catalog.ps1")
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    $env:EVALYTICS_OFFICE_PORT = "$Port"
    node .\scripts\serve-local.mjs
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
