param(
    [ValidateSet("GetVersion", "BumpDeployment")]
    [string]$Action = "GetVersion"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$versionFile = Join-Path $root "version.json"

if (-not (Test-Path -LiteralPath $versionFile)) {
    throw "Version file not found: $versionFile"
}

$versionData = Get-Content -LiteralPath $versionFile -Raw | ConvertFrom-Json

function Get-VersionString {
    param([pscustomobject]$VersionObject)
    return "$($VersionObject.major).$($VersionObject.minor).$($VersionObject.patch)"
}

switch ($Action) {
    "GetVersion" {
        Write-Output (Get-VersionString -VersionObject $versionData)
    }
    "BumpDeployment" {
        $versionData.patch = [int]$versionData.patch + 1
        $json = $versionData | ConvertTo-Json
        Set-Content -LiteralPath $versionFile -Value $json -Encoding UTF8
        Write-Output (Get-VersionString -VersionObject $versionData)
    }
}

