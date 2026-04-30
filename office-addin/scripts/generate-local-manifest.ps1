param(
    [string]$BaseUrl = "https://localhost:3000"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$sourcePath = Join-Path $root "manifest.xml"
$docsSourcePath = Join-Path $root "manifest.docs.xml"
$distDir = Join-Path $root "dist"
$outputPath = Join-Path $distDir "manifest.local.xml"
$docsOutputPath = Join-Path $distDir "manifest.docs.local.xml"

if (-not (Test-Path $sourcePath)) {
    throw "Manifest not found: $sourcePath"
}

if (-not (Test-Path $docsSourcePath)) {
    throw "Docs manifest not found: $docsSourcePath"
}

New-Item -ItemType Directory -Force -Path $distDir | Out-Null

$manifest = Get-Content -LiteralPath $sourcePath -Raw
$cacheKey = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$manifestVersionRevision = $cacheKey % 65535
$manifest = $manifest.Replace("<Version>1.0.0.0</Version>", "<Version>1.0.0.$manifestVersionRevision</Version>")
$manifest = $manifest.Replace("https://xlstat.evalytics.org/office/functions.json", "$BaseUrl/functions.json?v=$cacheKey")
$manifest = $manifest.Replace("https://xlstat.evalytics.org/office/functions.html", "$BaseUrl/functions.html?v=$cacheKey")
$manifest = $manifest.Replace("https://xlstat.evalytics.org/office/functions.js", "$BaseUrl/functions.js?v=$cacheKey")
$manifest = $manifest.Replace("<AppDomain>https://xlstat.evalytics.org</AppDomain>", "<AppDomain>$BaseUrl</AppDomain>")

Set-Content -LiteralPath $outputPath -Value $manifest -Encoding utf8

$docsManifest = Get-Content -LiteralPath $docsSourcePath -Raw
$docsManifest = $docsManifest.Replace("<Version>1.0.0.0</Version>", "<Version>1.0.0.$manifestVersionRevision</Version>")
$docsManifest = $docsManifest.Replace("https://xlstat.evalytics.org/office/commands.html", "$BaseUrl/commands.html?v=$cacheKey")
$docsManifest = $docsManifest.Replace("https://xlstat.evalytics.org/office/taskpane.html", "$BaseUrl/taskpane.html?v=$cacheKey")
$docsManifest = $docsManifest.Replace("https://xlstat.evalytics.org/office/icon-16.png", "$BaseUrl/icon-16.png?v=$cacheKey")
$docsManifest = $docsManifest.Replace("https://xlstat.evalytics.org/office/icon-32.png", "$BaseUrl/icon-32.png?v=$cacheKey")
$docsManifest = $docsManifest.Replace("https://xlstat.evalytics.org/office/icon-80.png", "$BaseUrl/icon-80.png?v=$cacheKey")
$docsManifest = $docsManifest.Replace("<AppDomain>https://xlstat.evalytics.org</AppDomain>", "<AppDomain>$BaseUrl</AppDomain>")

Set-Content -LiteralPath $docsOutputPath -Value $docsManifest -Encoding utf8

Write-Host "Local manifest generated at $outputPath"
Write-Host "Local docs manifest generated at $docsOutputPath"
