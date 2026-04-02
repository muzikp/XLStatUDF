param(
    [string]$Configuration = "Release",
    [switch]$CreateInstaller = $true
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if (-not (Test-Path ".dotnet_home")) {
    New-Item -ItemType Directory -Path ".dotnet_home" | Out-Null
}

if (-not (Test-Path ".nuget_packages")) {
    New-Item -ItemType Directory -Path ".nuget_packages" | Out-Null
}

$env:DOTNET_CLI_HOME = (Resolve-Path ".\.dotnet_home")
$env:NUGET_PACKAGES = (Resolve-Path ".\.nuget_packages")
$env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"

$artifactRoot = Join-Path $root "artifacts"
$mainArtifactDir = Join-Path $artifactRoot "main"
$installerArtifactDir = Join-Path $artifactRoot "installer"
$csSetupExe = Join-Path $installerArtifactDir "XLStatUDF_CS_Setup.exe"
$enSetupExe = Join-Path $installerArtifactDir "XLStatUDF_EN_Setup.exe"

if (Test-Path $mainArtifactDir) {
    Remove-Item -LiteralPath $mainArtifactDir -Recurse -Force
}

New-Item -ItemType Directory -Path $mainArtifactDir | Out-Null

if (-not (Test-Path $installerArtifactDir)) {
    New-Item -ItemType Directory -Path $installerArtifactDir | Out-Null
}

dotnet restore .\XLStatUDF.sln
dotnet test .\tests\XLStatUDF.Tests\XLStatUDF.Tests.csproj -c $Configuration /p:RunExcelDnaPack=false
dotnet build .\src\XLStatUDF\XLStatUDF.csproj -c $Configuration --no-restore -o $mainArtifactDir

$packedXll = Join-Path $mainArtifactDir "publish\XLStatUDF-AddIn64-packed.xll"
$installerTarget = Join-Path $root "installer\XLStatUDF-packed.xll"

Copy-Item $packedXll $installerTarget -Force

$canBuildInstaller = $true
foreach ($setupExe in @($csSetupExe, $enSetupExe)) {
    if (Test-Path $setupExe) {
        try {
            Remove-Item -LiteralPath $setupExe -Force
        }
        catch {
            Write-Warning "Existujici instalacni EXE je zamceny jinym procesem a nebyl prepsan: $setupExe"
            Write-Warning "Zavri prosim bezici instalator nebo Explorer preview a spust build znovu pro novou verzi EXE."
            $canBuildInstaller = $false
        }
    }
}

if ($CreateInstaller -and $canBuildInstaller) {
    $isccCandidates = @(
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"),
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe",
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 7\ISCC.exe"),
        "C:\Program Files (x86)\Inno Setup 7\ISCC.exe",
        "C:\Program Files\Inno Setup 7\ISCC.exe"
    )

    $iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $iscc) {
        $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
        if ($command) {
            $iscc = $command.Source
        }
    }

    if ($iscc) {
        $installerScript = Join-Path $root "installer\XLStatUDF.iss"
        & $iscc "/DMyAppVersion=1.0.0" "/DInstallerCulture=CS" $installerScript
        & $iscc "/DMyAppVersion=1.0.0" "/DInstallerCulture=EN" $installerScript
        Write-Host "Instalator CS: $csSetupExe"
        Write-Host "Instalator EN: $enSetupExe"
    }
    else {
        Write-Warning "Inno Setup compiler (ISCC.exe) nebyl nalezen. Instalacni EXE nebyl vygenerovan."
        Write-Warning "Jakmile bude Inno Setup nainstalovany, staci znovu spustit .\build.ps1."
    }
}

Write-Host "Build hotovy."
Write-Host "Hlavni add-in: $packedXll"
Write-Host "Kopie pro installer: $installerTarget"
if ($CreateInstaller -and $canBuildInstaller) {
    Write-Host "CS installer: $csSetupExe"
    Write-Host "EN installer: $enSetupExe"
}
