$ErrorActionPreference = "Stop"

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excel) {
    throw "Close Excel before clearing the Office add-in cache."
}

$wefRoot = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"

if (-not (Test-Path $wefRoot)) {
    Write-Host "Office WEF cache folder does not exist: $wefRoot"
    exit 0
}

$resolvedWefRoot = (Resolve-Path -LiteralPath $wefRoot).Path
$expectedPrefix = (Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0")

if (-not $resolvedWefRoot.StartsWith($expectedPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to clear unexpected cache path: $resolvedWefRoot"
}

Get-ChildItem -LiteralPath $resolvedWefRoot -Force | ForEach-Object {
    Remove-Item -LiteralPath $_.FullName -Recurse -Force
    Write-Host "Removed $($_.FullName)"
}

Write-Host "Office WEF add-in cache cleared. Start the local host, reopen Excel, and add Evalytics from the Shared Folder catalog again."
