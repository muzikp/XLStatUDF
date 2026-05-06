param(
    [int]$Port = 3000,
    [switch]$ClearOfficeCache,
    [switch]$NoExcel
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $root
$outLog = Join-Path $root ".local-serve.out.log"
$errLog = Join-Path $root ".local-serve.err.log"
$resetLog = Join-Path $root ".reset-debug-session.log"
$serverScript = Join-Path $root "scripts\serve-local.mjs"

Start-Transcript -LiteralPath $resetLog -Force | Out-Null
Write-Host "Evalytics reset debug session started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "Root: $root"

try {

function Get-ExcelPath {
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
    )

    foreach ($registryPath in $registryPaths) {
        $entry = Get-ItemProperty -LiteralPath $registryPath -ErrorAction SilentlyContinue
        if ($entry -and $entry."(default)" -and (Test-Path -LiteralPath $entry."(default)")) {
            return $entry."(default)"
        }
    }

    $runningExcel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($runningExcel -and $runningExcel.Path -and (Test-Path -LiteralPath $runningExcel.Path)) {
        return $runningExcel.Path
    }

    throw "Could not find EXCEL.EXE. Check the Office installation path."
}

function Stop-Excel {
    $excelProcesses = Get-Process EXCEL -ErrorAction SilentlyContinue
    if (-not $excelProcesses) {
        Write-Host "Excel is not running."
        return
    }

    Write-Host "Closing Excel..."
    foreach ($process in $excelProcesses) {
        try {
            if ($process.MainWindowHandle -ne 0) {
                [void]$process.CloseMainWindow()
            }
        }
        catch {
            Write-Host "Could not request graceful close for Excel PID $($process.Id): $($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 3
    $remaining = Get-Process EXCEL -ErrorAction SilentlyContinue
    if ($remaining) {
        Write-Host "Force-stopping remaining Excel processes..."
        $remaining | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    }

    $remaining = Get-Process EXCEL -ErrorAction SilentlyContinue
    if ($remaining) {
        Write-Host "Stop-Process did not close Excel; using taskkill fallback..."
        & taskkill.exe /IM EXCEL.EXE /F | Out-Host
        Start-Sleep -Seconds 1
    }

    $remaining = Get-Process EXCEL -ErrorAction SilentlyContinue
    if ($remaining) {
        $ids = ($remaining | ForEach-Object { $_.Id }) -join ", "
        throw "Excel is still running after close attempts. Remaining PID(s): $ids"
    }
}

function Stop-LocalServer {
    $listeners = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
    foreach ($listener in $listeners) {
        $process = Get-Process -Id $listener.OwningProcess -ErrorAction SilentlyContinue
        if (-not $process) {
            continue
        }

        if ($process.ProcessName -ne "node") {
            throw "Port $Port is already in use by PID $($process.Id) ($($process.ProcessName)). Stop it or choose another port."
        }

        Write-Host "Stopping existing Evalytics server on port $Port (PID $($process.Id))."
        Stop-Process -Id $process.Id -Force
    }
}

function Clear-OfficeCache {
    $wefRoot = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
    if (-not (Test-Path $wefRoot)) {
        Write-Host "Office WEF cache folder does not exist: $wefRoot"
        return
    }

    $resolvedWefRoot = (Resolve-Path -LiteralPath $wefRoot).Path
    $expectedPrefix = (Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0")
    if (-not $resolvedWefRoot.StartsWith($expectedPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clear unexpected cache path: $resolvedWefRoot"
    }

    Write-Host "Clearing Office WEF cache..."
    Get-ChildItem -LiteralPath $resolvedWefRoot -Force | ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
    }
}

Stop-Excel
Stop-LocalServer
if ($ClearOfficeCache) {
    Clear-OfficeCache
}
else {
    Write-Host "Keeping Office WEF cache. Existing inserted add-ins should stay available."
    Write-Host "Use -ClearOfficeCache only after custom-functions metadata or manifest identity changes."
}

Push-Location $root
try {
    & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "ensure-dev-cert.ps1")
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    npm run build
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "generate-local-manifest.ps1") -BaseUrl "https://localhost:$Port"
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "install-desktop-catalog.ps1")
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
finally {
    Pop-Location
}

if (Test-Path $outLog) { Remove-Item -LiteralPath $outLog -Force }
if (Test-Path $errLog) { Remove-Item -LiteralPath $errLog -Force }

Write-Host "Starting Evalytics server on https://localhost:$Port..."
$serverCommand = "`$env:EVALYTICS_OFFICE_PORT='$Port'; Set-Location '$root'; node '$serverScript'"
$server = Start-Process -FilePath powershell.exe `
    -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", $serverCommand) `
    -WindowStyle Hidden `
    -RedirectStandardOutput $outLog `
    -RedirectStandardError $errLog `
    -PassThru

Start-Sleep -Seconds 2
if ($server.HasExited) {
    Write-Host "Server failed to start. stderr:"
    if (Test-Path $errLog) {
        Get-Content -LiteralPath $errLog
    }
    exit 1
}

Write-Host "Server PID: $($server.Id)"
Write-Host "Logs:"
Write-Host "  $outLog"
Write-Host "  $errLog"

if (-not $NoExcel) {
    $excelPath = Get-ExcelPath
    Write-Host "Launching Excel..."
    Write-Host "  $excelPath"
    Start-Process -FilePath $excelPath
}

Write-Host ""
if ($ClearOfficeCache) {
    Write-Host "Office cache was cleared. In Excel, use Home > Add-ins > Advanced > Shared Folder if the add-ins are not already visible."
    Write-Host "Add Evalytics StatLab for Excel from \\localhost\EvalyticsOfficeAddin."
}
else {
    Write-Host "Office cache was kept. If the add-in was already added before, it should remain available."
    Write-Host "If Excel still shows stale custom-function metadata, rerun with -ClearOfficeCache."
}
}
finally {
    Write-Host "Reset log: $resetLog"
    Stop-Transcript | Out-Null
}
