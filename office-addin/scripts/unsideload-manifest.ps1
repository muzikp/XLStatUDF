$ErrorActionPreference = "Stop"

$targets = @(
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef\Evalytics.Office.local.xml"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef\XLStatUDF.Office.local.xml")
)

$removed = $false
foreach ($target in $targets) {
    if (Test-Path $target) {
        Remove-Item -LiteralPath $target -Force
        Write-Host "Removed sideloaded manifest from $target"
        $removed = $true
    }
}

if (-not $removed) {
    Write-Host "No sideloaded local manifest found."
}
