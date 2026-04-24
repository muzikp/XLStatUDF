param(
    [string]$Configuration = "Release",
    [string]$AppVersion = "",
    [switch]$CreateInstaller = $true
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if ([string]::IsNullOrWhiteSpace($AppVersion)) {
    $versionFile = Join-Path $root "version.json"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version file not found: $versionFile"
    }

    $versionData = Get-Content -LiteralPath $versionFile -Raw | ConvertFrom-Json
    $AppVersion = "$($versionData.major).$($versionData.minor).$($versionData.patch)"
}

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
$csArtifactDir = Join-Path $mainArtifactDir "cs"
$enArtifactDir = Join-Path $mainArtifactDir "en"
$installerArtifactDir = Join-Path $artifactRoot "installer"
$csSetupExe = Join-Path $installerArtifactDir "XLStatUDF_CS_Setup.exe"
$enSetupExe = Join-Path $installerArtifactDir "XLStatUDF_EN_Setup.exe"

if (Test-Path $mainArtifactDir) {
    Remove-Item -LiteralPath $mainArtifactDir -Recurse -Force
}

foreach ($path in @($csArtifactDir, $enArtifactDir, $installerArtifactDir)) {
    if (-not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
}

dotnet restore .\XLStatUDF.sln
dotnet test .\tests\XLStatUDF.Tests\XLStatUDF.Tests.csproj -c $Configuration /p:RunExcelDnaPack=false

dotnet build .\src\XLStatUDF\XLStatUDF.csproj -c $Configuration --no-restore -o $csArtifactDir /p:AddInLanguage=CS /p:Version=$AppVersion /p:FileVersion=$AppVersion /p:InformationalVersion=$AppVersion
dotnet build .\src\XLStatUDF\XLStatUDF.csproj -c $Configuration --no-restore -o $enArtifactDir /p:AddInLanguage=EN /p:Version=$AppVersion /p:FileVersion=$AppVersion /p:InformationalVersion=$AppVersion

$csPackedXll = Join-Path $csArtifactDir "publish\XLStatUDF-AddIn64-packed.xll"
$enPackedXll = Join-Path $enArtifactDir "publish\XLStatUDF-AddIn64-packed.xll"

$canBuildInstaller = $true
foreach ($setupExe in @($csSetupExe, $enSetupExe)) {
    if (Test-Path $setupExe) {
        try {
            Remove-Item -LiteralPath $setupExe -Force
        }
        catch {
            Write-Warning "Existing installer EXE is locked by another process and could not be overwritten: $setupExe"
            Write-Warning "Close any running installer or Explorer preview and run the build again to regenerate the EXE."
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
        & $iscc "/DMyAppVersion=$AppVersion" "/DInstallerCulture=CS" "/DAddInSource=$csPackedXll" $installerScript
        & $iscc "/DMyAppVersion=$AppVersion" "/DInstallerCulture=EN" "/DAddInSource=$enPackedXll" $installerScript
        Write-Host "Installer CS: $csSetupExe"
        Write-Host "Installer EN: $enSetupExe"
    }
    else {
        Write-Warning "Inno Setup compiler (ISCC.exe) was not found. Installer EXE files were not generated."
        Write-Warning "Once Inno Setup is installed, just run .\build.ps1 again."
    }
}

Write-Host "Build completed."
Write-Host "Version: $AppVersion"
Write-Host "Czech add-in: $csPackedXll"
Write-Host "English add-in: $enPackedXll"
if ($CreateInstaller -and $canBuildInstaller) {
    Write-Host "CS installer: $csSetupExe"
    Write-Host "EN installer: $enSetupExe"
}
