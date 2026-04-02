param(
    [string]$Configuration = "Release"
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

if (Test-Path $mainArtifactDir) {
    Remove-Item -LiteralPath $mainArtifactDir -Recurse -Force
}

New-Item -ItemType Directory -Path $mainArtifactDir | Out-Null

dotnet restore .\XLStatUDF.sln
dotnet test .\tests\XLStatUDF.Tests\XLStatUDF.Tests.csproj -c $Configuration /p:RunExcelDnaPack=false
dotnet build .\src\XLStatUDF\XLStatUDF.csproj -c $Configuration --no-restore -o $mainArtifactDir

$packedXll = Join-Path $mainArtifactDir "publish\XLStatUDF-AddIn64-packed.xll"
$installerTarget = Join-Path $root "installer\XLStatUDF-packed.xll"

Copy-Item $packedXll $installerTarget -Force

Write-Host "Build hotovy."
Write-Host "Hlavni add-in: $packedXll"
Write-Host "Kopie pro installer: $installerTarget"
