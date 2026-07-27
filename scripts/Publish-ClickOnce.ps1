[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[0-9A-Fa-f]{40}$")]
    [string]$CertificateThumbprint,

    [string]$PublishPath = "artifacts\ClickOnce",

    [string]$InstallUrl = "",

    [ValidatePattern("^\d+\.\d+\.\d+\.\d+$")]
    [string]$Version = "1.0.0.0"
)

$ErrorActionPreference = "Stop"
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repositoryRoot "src\ReportAutomation.Vsto\ReportAutomation.Vsto.csproj"
$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"

if (-not (Test-Path -LiteralPath $vswhere)) {
    throw "Visual Studio Installer could not be found. Install Visual Studio 2022 with Office/SharePoint development tools."
}

$visualStudioPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
if (-not $visualStudioPath) {
    throw "Visual Studio 2022 MSBuild could not be found."
}

$msbuild = Join-Path $visualStudioPath "MSBuild\Current\Bin\MSBuild.exe"
$certificate = Get-Item -LiteralPath "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
if (-not $certificate) {
    throw "Certificate $CertificateThumbprint was not found in Cert:\CurrentUser\My."
}

$resolvedPublishPath = if ([System.IO.Path]::IsPathRooted($PublishPath)) {
    $PublishPath
} else {
    Join-Path $repositoryRoot $PublishPath
}

$installFrom = "Disk"
if ($InstallUrl.StartsWith("\\")) {
    $installFrom = "Unc"
} elseif ($InstallUrl -match "^https://") {
    $installFrom = "Web"
} elseif ($InstallUrl) {
    throw "InstallUrl must be an HTTPS URL, a UNC path beginning with \\, or empty for local/offline installation."
}

$arguments = @(
    $projectPath,
    "/t:Publish",
    "/p:Configuration=Release",
    "/p:SignManifests=true",
    "/p:ManifestCertificateThumbprint=$CertificateThumbprint",
    "/p:ApplicationVersion=$Version",
    "/p:MinimumRequiredVersion=$Version",
    "/p:PublishUrl=$resolvedPublishPath",
    "/p:PublishDir=$resolvedPublishPath",
    "/p:InstallFrom=$installFrom"
)

if ($InstallUrl) {
    $arguments += "/p:InstallUrl=$InstallUrl"
}

& $msbuild @arguments
if ($LASTEXITCODE -ne 0) {
    throw "ClickOnce publishing failed with exit code $LASTEXITCODE."
}

Write-Host "ClickOnce package published to: $resolvedPublishPath"
Write-Host "Distribute setup.exe and the complete folder together."
