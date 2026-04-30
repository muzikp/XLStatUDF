param(
    [string]$ShareName = "EvalyticsOfficeAddin",
    [string]$CatalogId = "{2F27D1E8-915D-4F6E-91CC-62F6BF10CE11}"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$catalogDir = Join-Path $root ".catalog"
$manifestSource = Join-Path $root "dist\manifest.local.xml"
$docsManifestSource = Join-Path $root "dist\manifest.docs.local.xml"
$manifestTarget = Join-Path $catalogDir "Evalytics.Office.local.xml"
$docsManifestTarget = Join-Path $catalogDir "Evalytics.Docs.local.xml"
$catalogUrl = "\\localhost\$ShareName"
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogId"

if (-not (Test-Path $manifestSource) -or -not (Test-Path $docsManifestSource)) {
    & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "generate-local-manifest.ps1")
}

if (-not (Test-Path $manifestSource) -or (Get-Item -LiteralPath $manifestSource).Length -eq 0) {
    throw "Local manifest not found: $manifestSource"
}

if (-not (Test-Path $docsManifestSource) -or (Get-Item -LiteralPath $docsManifestSource).Length -eq 0) {
    throw "Local docs manifest not found: $docsManifestSource"
}

New-Item -ItemType Directory -Force -Path $catalogDir | Out-Null
Copy-Item -LiteralPath $manifestSource -Destination $manifestTarget -Force
Copy-Item -LiteralPath $docsManifestSource -Destination $docsManifestTarget -Force

if ((Get-Item -LiteralPath $manifestTarget).Length -eq 0) {
    throw "Catalog manifest copy failed: $manifestTarget"
}

if ((Get-Item -LiteralPath $docsManifestTarget).Length -eq 0) {
    throw "Catalog docs manifest copy failed: $docsManifestTarget"
}

$shareReady = $false
try {
    $share = Get-SmbShare -Name $ShareName -ErrorAction Stop
    $shareReady = $null -ne $share
}
catch {
    try {
        New-SmbShare -Name $ShareName -Path $catalogDir -ReadAccess $env:USERNAME -ErrorAction Stop | Out-Null
        $shareReady = $true
    }
    catch {
        $shareReady = $false
    }
}

New-Item -Path $registryPath -Force | Out-Null
New-ItemProperty -Path $registryPath -Name "Id" -Value $CatalogId -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registryPath -Name "Url" -Value $catalogUrl -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registryPath -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host "Desktop Excel manifest catalog prepared."
Write-Host "  UDF manifest: $manifestTarget"
Write-Host "  Docs manifest: $docsManifestTarget"
Write-Host "  Catalog URL: $catalogUrl"
Write-Host "  Registry: $registryPath"
Write-Host "  SMB share ready: $shareReady"
Write-Host ""
if (-not $shareReady) {
    Write-Host "Windows did not allow this script to create/read the SMB share."
    Write-Host "Create it manually in File Explorer:"
    Write-Host "  1. Right-click $catalogDir"
    Write-Host "  2. Properties > Sharing > Advanced Sharing"
    Write-Host "  3. Share this folder as: $ShareName"
    Write-Host "  4. Give your Windows user read permission"
    Write-Host ""
}
Write-Host "Then restart Excel, go to Home > Add-ins > Advanced > Shared Folder, and add Evalytics plus Evalytics Docs."
