[CmdletBinding()]
param(
    [ValidateSet('Trial', 'Standard', 'Professional', 'Automation')]
    [string]$Edition = 'Professional',

    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$ApplicationVersion = '1.1.0.0',

    [ValidateSet('Online')]
    [string]$PrerequisiteMode = 'Online',

    [string]$OutputRoot = (Join-Path $PSScriptRoot '..\artifacts\sellable-release'),

    [switch]$SkipZip
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$projectFile = Join-Path $projectRoot 'AgentForExcel.csproj'
$releaseName = "AgentForExcel-$ApplicationVersion-$Edition"
$outputRootPath = if ([System.IO.Path]::IsPathRooted($OutputRoot))
{
    $OutputRoot
}
else
{
    Join-Path $projectRoot $OutputRoot
}
if (-not (Test-Path -LiteralPath $outputRootPath))
{
    New-Item -ItemType Directory -Path $outputRootPath | Out-Null
}
$outputRootPath = (Resolve-Path $outputRootPath).Path
$releaseRoot = Join-Path $outputRootPath $releaseName

if (Test-Path -LiteralPath $releaseRoot) {
    throw "Release directory already exists: $releaseRoot. Choose a new version or output directory."
}

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$msbuild = if (Test-Path -LiteralPath $vswhere) {
    & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
}
if ([string]::IsNullOrWhiteSpace($msbuild)) {
    $msbuild = (Get-Command msbuild.exe -ErrorAction SilentlyContinue).Source
}
if ([string]::IsNullOrWhiteSpace($msbuild) -or -not (Test-Path -LiteralPath $msbuild)) {
    throw 'MSBuild with the Office/VSTO development components was not found on this build machine. Buyers do not need Visual Studio.'
}

New-Item -ItemType Directory -Path $releaseRoot | Out-Null
$publishPath = Join-Path $releaseRoot 'publish'
New-Item -ItemType Directory -Path $publishPath | Out-Null
$generatedPublishPath = Join-Path $projectRoot 'bin\Release\app.publish'

$publishArgument = '/p:PublishUrl=bin\Release\app.publish\'
Write-Host "Building $Edition package. Prerequisites mode: $PrerequisiteMode (downloaded from Microsoft during installation)."
$buildArgs = @(
    $projectFile,
    '/t:Publish',
    '/p:Configuration=Release',
    '/p:Platform=AnyCPU',
    "/p:AgentEdition=$Edition",
    "/p:ApplicationVersion=$ApplicationVersion",
    $publishArgument,
    '/p:BootstrapperComponentsLocation=HomeSite'
)
& $msbuild @buildArgs
if ($LASTEXITCODE -ne 0) {
    throw "Publish build failed with exit code: $LASTEXITCODE"
}

if (-not (Test-Path -LiteralPath $generatedPublishPath)) {
    throw "VSTO publish output was not created: $generatedPublishPath"
}
Copy-Item -Path (Join-Path $generatedPublishPath '*') -Destination $publishPath -Recurse

$requiredFiles = @(
    (Join-Path $publishPath 'setup.exe'),
    (Join-Path $publishPath 'AgentForExcel.vsto'),
    (Join-Path $publishPath 'Application Files')
)
$missing = $requiredFiles | Where-Object { -not (Test-Path -LiteralPath $_) }
if ($missing) {
    throw "The sellable installer structure is incomplete. Missing: $($missing -join '; '). Do not deliver this package."
}

Copy-Item -LiteralPath (Join-Path $projectRoot 'SELLABLE_RELEASE.md') -Destination $releaseRoot
Copy-Item -LiteralPath (Join-Path $projectRoot 'README.md') -Destination $releaseRoot
$acceptanceTools = Join-Path $releaseRoot 'AcceptanceTools'
New-Item -ItemType Directory -Path $acceptanceTools | Out-Null
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Test-SellableRelease.ps1') -Destination $acceptanceTools
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Test-InstalledAddIn.ps1') -Destination $acceptanceTools
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Sign-SellableRelease.ps1') -Destination $acceptanceTools
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Invoke-CleanMachineAcceptance.ps1') -Destination $acceptanceTools
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Test-ReleaseEvidence.ps1') -Destination $acceptanceTools

if (-not $SkipZip) {
    $zipPath = Join-Path $releaseRoot ($releaseName + '.zip')
    Compress-Archive -Path (Join-Path $publishPath '*'), (Join-Path $releaseRoot 'SELLABLE_RELEASE.md'), (Join-Path $releaseRoot 'README.md'), $acceptanceTools -DestinationPath $zipPath
    Write-Host "Delivery archive created: $zipPath"
}

$verifier = Join-Path $PSScriptRoot 'Test-SellableRelease.ps1'
if (-not (Test-Path -LiteralPath $verifier)) {
    throw "Release verifier is missing: $verifier"
}
$validationTarget = if ($SkipZip) { $releaseRoot } else { $zipPath }
$validationReport = Join-Path $releaseRoot 'AcceptanceReport-build.json'
& $verifier `
    -Package $validationTarget `
    -ExpectedVersion $ApplicationVersion `
    -ExpectedEdition $Edition `
    -ReportPath $validationReport
if (-not $?) {
    throw 'Static release acceptance failed. Do not deliver this package.'
}

Write-Host 'Build complete. Before delivery, blind-install on clean Windows + Excel machines and replace the development certificate with a trusted code-signing certificate.'
