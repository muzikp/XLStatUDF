param(
    [string]$EnvironmentFile = "..\.env",
    [switch]$SkipBuild,
    [switch]$SkipInvalidation
)

$ErrorActionPreference = "Stop"

function Set-EnvFromFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    foreach ($line in Get-Content -LiteralPath $Path) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith("#")) {
            continue
        }

        $separatorIndex = $trimmed.IndexOf("=")
        if ($separatorIndex -lt 1) {
            continue
        }

        $name = $trimmed.Substring(0, $separatorIndex).Trim()
        $value = $trimmed.Substring($separatorIndex + 1).Trim()

        if (($value.StartsWith('"') -and $value.EndsWith('"')) -or ($value.StartsWith("'") -and $value.EndsWith("'"))) {
            $value = $value.Substring(1, $value.Length - 2)
        }

        [System.Environment]::SetEnvironmentVariable($name, $value, "Process")
    }
}

function Require-Env {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $value = [System.Environment]::GetEnvironmentVariable($Name, "Process")
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "Required environment variable '$Name' is missing."
    }

    return $value
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$websiteRoot = Resolve-Path (Join-Path $scriptRoot "..")
Set-Location $websiteRoot

Set-EnvFromFile -Path (Join-Path $websiteRoot $EnvironmentFile)

$bucket = Require-Env -Name "AWS_S3_BUCKET"
$region = Require-Env -Name "AWS_REGION"
$distributionId = Require-Env -Name "AWS_CLOUDFRONT_DISTRIBUTION_ID"
$null = Require-Env -Name "AWS_ACCESS_KEY_ID"
$null = Require-Env -Name "AWS_SECRET_ACCESS_KEY"

if (-not $SkipBuild) {
    npm run build
}

$buildDir = Resolve-Path ".\build"

Write-Host "Uploading HTML with short cache policy..."
aws s3 sync $buildDir "s3://$bucket" `
    --delete `
    --exclude "*" `
    --include "*.html" `
    --cache-control "public,max-age=300,must-revalidate" `
    --region $region

Write-Host "Uploading static assets with long cache policy..."
aws s3 sync $buildDir "s3://$bucket" `
    --delete `
    --exclude "*.html" `
    --cache-control "public,max-age=31536000,immutable" `
    --region $region

if (-not $SkipInvalidation) {
    Write-Host "Creating CloudFront invalidation..."
    aws cloudfront create-invalidation `
        --distribution-id $distributionId `
        --paths "/*" | Out-Null
}

Write-Host "Deployment finished for bucket $bucket."

