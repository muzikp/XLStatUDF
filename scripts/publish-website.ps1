param(
    [switch]$SkipVersionBump,
    [switch]$SkipInvalidation
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if ($SkipVersionBump) {
    $appVersion = & powershell -ExecutionPolicy Bypass -File ".\scripts\versioning.ps1" -Action GetVersion
}
else {
    $appVersion = & powershell -ExecutionPolicy Bypass -File ".\scripts\versioning.ps1" -Action BumpDeployment
}

Write-Host "Publishing XLStatUDF version $appVersion"

& powershell -ExecutionPolicy Bypass -File ".\build.ps1" -AppVersion $appVersion

$deployArgs = @(
    "-ExecutionPolicy", "Bypass",
    "-File", ".\scripts\deploy-aws.ps1"
)

if ($SkipInvalidation) {
    $deployArgs += "-SkipInvalidation"
}

Push-Location ".\website"
try {
    & powershell @deployArgs
}
finally {
    Pop-Location
}

