[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ReleaseDirectory,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[A-Fa-f0-9]{40}$')]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Trial', 'Standard', 'Professional', 'Automation')]
    [string]$Edition,

    [ValidateSet('Rehearsal', 'Formal')]
    [string]$Mode = 'Rehearsal',

    [string]$TimestampUri,

    [string]$MagePath,

    [string]$SignToolPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Find-Tool {
    param(
        [string]$ExplicitPath,
        [string]$FileName,
        [string[]]$Roots
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        $resolved = (Resolve-Path -LiteralPath $ExplicitPath -ErrorAction Stop).Path
        if (-not [string]::Equals(
            [IO.Path]::GetFileName($resolved), $FileName, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Expected $FileName but received: $resolved"
        }
        return $resolved
    }

    $command = Get-Command $FileName -ErrorAction SilentlyContinue
    if ($null -ne $command -and (Test-Path -LiteralPath $command.Source)) {
        return $command.Source
    }

    foreach ($root in $Roots) {
        if ([string]::IsNullOrWhiteSpace($root) -or -not (Test-Path -LiteralPath $root)) {
            continue
        }
        $match = Get-ChildItem -LiteralPath $root -Filter $FileName -Recurse -File `
            -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($null -ne $match) {
            return $match.FullName
        }
    }
    return $null
}

function Invoke-ExternalTool {
    param(
        [string]$Path,
        [string[]]$Arguments,
        [string]$Description
    )

    & $Path @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $LASTEXITCODE."
    }
}

function Get-CodeSigningCertificate {
    param([string]$Thumbprint)

    $normalized = $Thumbprint.Replace(' ', '').ToUpperInvariant()
    $certificate = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My `
        -ErrorAction SilentlyContinue |
        Where-Object { $_.Thumbprint -eq $normalized } |
        Select-Object -First 1
    if ($null -eq $certificate) {
        throw "Certificate was not found in CurrentUser/My or LocalMachine/My: $normalized"
    }
    if (-not $certificate.HasPrivateKey) {
        throw 'The selected certificate does not expose a private key.'
    }
    $now = Get-Date
    if ($now -lt $certificate.NotBefore -or $now -gt $certificate.NotAfter) {
        throw 'The selected certificate is outside its validity period.'
    }
    return $certificate
}

function Test-CertificateTrust {
    param([Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    $chain = New-Object Security.Cryptography.X509Certificates.X509Chain
    $chain.ChainPolicy.RevocationMode =
        [Security.Cryptography.X509Certificates.X509RevocationMode]::Online
    $chain.ChainPolicy.RevocationFlag =
        [Security.Cryptography.X509Certificates.X509RevocationFlag]::ExcludeRoot
    $trusted = $chain.Build($Certificate)
    return [pscustomobject]@{
        trusted = $trusted
        statuses = @($chain.ChainStatus | ForEach-Object { $_.Status.ToString() })
    }
}

function Test-CodeSigningUsage {
    param([Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    $usages = @($Certificate.Extensions |
        Where-Object { $_.Oid.Value -eq '2.5.29.37' } |
        ForEach-Object {
            $enhanced = New-Object Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension(
                $_, $_.Critical)
            @($enhanced.EnhancedKeyUsages | ForEach-Object { $_.Value })
        })
    return $usages.Count -gt 0 -and $usages -contains '1.3.6.1.5.5.7.3.3'
}

$sourceRoot = (Resolve-Path -LiteralPath $ReleaseDirectory -ErrorAction Stop).Path
$outputFullPath = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $outputFullPath) {
    throw "Output directory already exists: $outputFullPath"
}

$sourcePublishDirectory = Join-Path $sourceRoot 'publish'
$sourceHasPublishDirectory = Test-Path -LiteralPath $sourcePublishDirectory
$sourcePublish = $sourcePublishDirectory
if (-not $sourceHasPublishDirectory) {
    $sourcePublish = $sourceRoot
}
$sourceDeploymentManifest = Join-Path $sourcePublish 'AgentForExcel.vsto'
$sourceSetup = Join-Path $sourcePublish 'setup.exe'
if (-not (Test-Path -LiteralPath $sourceDeploymentManifest) -or
    -not (Test-Path -LiteralPath $sourceSetup)) {
    throw 'Release directory must contain setup.exe and AgentForExcel.vsto.'
}

$certificate = Get-CodeSigningCertificate $CertificateThumbprint
$certificateInCurrentUserStore = Test-Path -LiteralPath (
    'Cert:\CurrentUser\My\' + $certificate.Thumbprint)
$trust = Test-CertificateTrust $certificate
$isSelfSigned = [string]::Equals(
    $certificate.Subject, $certificate.Issuer, [StringComparison]::OrdinalIgnoreCase)
$hasCodeSigningUsage = Test-CodeSigningUsage $certificate

if ($Mode -eq 'Formal') {
    if ([string]::IsNullOrWhiteSpace($TimestampUri)) {
        throw 'Formal signing requires -TimestampUri.'
    }
    $parsedTimestamp = $null
    if (-not [Uri]::TryCreate($TimestampUri, [UriKind]::Absolute, [ref]$parsedTimestamp) -or
        $parsedTimestamp.Scheme -notin @('http', 'https')) {
        throw 'TimestampUri must be an absolute HTTP or HTTPS URI.'
    }
    if ($isSelfSigned -or -not $trust.trusted) {
        throw 'Formal signing requires a certificate that chains to a trusted root.'
    }
    if (-not $hasCodeSigningUsage) {
        throw 'Formal signing requires the Code Signing enhanced key usage.'
    }
}

$sdkRoots = @(
    (Join-Path ${env:ProgramFiles(x86)} 'Microsoft SDKs\Windows'),
    (Join-Path ${env:ProgramFiles(x86)} 'Windows Kits')
)
$mage = Find-Tool -ExplicitPath $MagePath -FileName 'mage.exe' -Roots $sdkRoots
if ([string]::IsNullOrWhiteSpace($mage)) {
    throw 'Mage.exe was not found. Install the .NET Framework 4.8 SDK.'
}
$signTool = Find-Tool -ExplicitPath $SignToolPath -FileName 'signtool.exe' -Roots $sdkRoots
if ($Mode -eq 'Formal' -and [string]::IsNullOrWhiteSpace($signTool)) {
    throw 'Formal signing requires SignTool.exe from the Windows SDK.'
}

New-Item -ItemType Directory -Path $outputFullPath | Out-Null
$targetPublish = Join-Path $outputFullPath 'publish'
New-Item -ItemType Directory -Path $targetPublish | Out-Null
if ($sourceHasPublishDirectory) {
    Copy-Item -Path (Join-Path $sourcePublish '*') -Destination $targetPublish -Recurse
}
else {
    foreach ($publishItem in @('setup.exe', 'AgentForExcel.vsto', 'Application Files')) {
        $sourceItem = Join-Path $sourcePublish $publishItem
        if (-not (Test-Path -LiteralPath $sourceItem)) {
            throw "Flattened release is missing required publish item: $publishItem"
        }
        Copy-Item -LiteralPath $sourceItem -Destination $targetPublish -Recurse
    }
}

foreach ($fileName in @('README.md', 'SELLABLE_RELEASE.md')) {
    $sourceFile = Join-Path $sourceRoot $fileName
    if (Test-Path -LiteralPath $sourceFile) {
        Copy-Item -LiteralPath $sourceFile -Destination $outputFullPath
    }
}
$sourceTools = Join-Path $sourceRoot 'AcceptanceTools'
if (Test-Path -LiteralPath $sourceTools) {
    Copy-Item -LiteralPath $sourceTools -Destination $outputFullPath -Recurse
}

$deploymentManifest = Join-Path $targetPublish 'AgentForExcel.vsto'
$setup = Join-Path $targetPublish 'setup.exe'
[xml]$deploymentXml = Get-Content -LiteralPath $deploymentManifest -Raw
$identity = $deploymentXml.SelectSingleNode(
    "/*[local-name()='assembly']/*[local-name()='assemblyIdentity']")
$dependentAssembly = $deploymentXml.SelectSingleNode(
    "/*[local-name()='assembly']/*[local-name()='dependency']/*[local-name()='dependentAssembly']")
$version = [string]$identity.version
$applicationCodebase = [string]$dependentAssembly.codebase
$applicationManifest = Join-Path $targetPublish ($applicationCodebase.Replace('/', '\'))
if (-not (Test-Path -LiteralPath $applicationManifest)) {
    throw "Application manifest referenced by deployment manifest was not found: $applicationCodebase"
}

$mageSignArguments = @(
    '-Sign', $applicationManifest,
    '-CertHash', $certificate.Thumbprint,
    '-Algorithm', 'sha256RSA'
)
if (-not [string]::IsNullOrWhiteSpace($TimestampUri)) {
    $mageSignArguments += @('-TimestampUri', $TimestampUri)
}
Invoke-ExternalTool $mage $mageSignArguments 'Application manifest signing'

Invoke-ExternalTool $mage @(
    '-Update', $deploymentManifest,
    '-AppManifest', $applicationManifest
) 'Deployment manifest update'

$mageDeploymentArguments = @(
    '-Sign', $deploymentManifest,
    '-CertHash', $certificate.Thumbprint,
    '-Algorithm', 'sha256RSA'
)
if (-not [string]::IsNullOrWhiteSpace($TimestampUri)) {
    $mageDeploymentArguments += @('-TimestampUri', $TimestampUri)
}
Invoke-ExternalTool $mage $mageDeploymentArguments 'Deployment manifest signing'

if ($Mode -eq 'Formal') {
    $signToolArguments = @(
        'sign',
        '/sha1', $certificate.Thumbprint,
        '/s', 'My',
        '/fd', 'SHA256',
        '/tr', $TimestampUri,
        '/td', 'SHA256',
        '/v',
        $setup
    )
    if (-not $certificateInCurrentUserStore) {
        $signToolArguments = @('sign', '/sm') + $signToolArguments[1..($signToolArguments.Count - 1)]
    }
    Invoke-ExternalTool $signTool $signToolArguments 'Bootstrapper Authenticode signing'
}
else {
    $signature = Get-AuthenticodeSignature -LiteralPath $setup
    $alreadySignedBySelectedCertificate =
        $null -ne $signature.SignerCertificate -and
        $signature.SignerCertificate.Thumbprint -eq $certificate.Thumbprint
    if (-not $alreadySignedBySelectedCertificate) {
        $authenticodeParameters = @{
            FilePath = $setup
            Certificate = $certificate
            HashAlgorithm = 'SHA256'
        }
        if (-not [string]::IsNullOrWhiteSpace($TimestampUri)) {
            $authenticodeParameters.TimestampServer = $TimestampUri
        }
        $signature = Set-AuthenticodeSignature @authenticodeParameters
        if ($null -eq $signature.SignerCertificate -or
            $signature.SignerCertificate.Thumbprint -ne $certificate.Thumbprint) {
            throw 'Bootstrapper rehearsal signing did not apply the selected certificate.'
        }
    }
}

Invoke-ExternalTool $mage @('-Verify', $applicationManifest) 'Application manifest verification'
Invoke-ExternalTool $mage @('-Verify', $deploymentManifest) 'Deployment manifest verification'

$setupSignature = Get-AuthenticodeSignature -LiteralPath $setup
if ($null -eq $setupSignature.SignerCertificate -or
    $setupSignature.SignerCertificate.Thumbprint -ne $certificate.Thumbprint) {
    throw 'Bootstrapper signature verification returned the wrong signer.'
}
if ($Mode -eq 'Formal') {
    Invoke-ExternalTool $signTool @('verify', '/pa', '/all', '/v', $setup) `
        'Bootstrapper Authenticode verification'
}

$archiveName = "AgentForExcel-$version-$Edition-Signed.zip"
$archivePath = Join-Path $outputFullPath $archiveName
$archiveInputs = @(
    (Join-Path $targetPublish '*')
)
foreach ($optional in @('README.md', 'SELLABLE_RELEASE.md', 'AcceptanceTools')) {
    $candidate = Join-Path $outputFullPath $optional
    if (Test-Path -LiteralPath $candidate) {
        $archiveInputs += $candidate
    }
}
Compress-Archive -Path $archiveInputs -DestinationPath $archivePath

$verifier = Join-Path $PSScriptRoot 'Test-SellableRelease.ps1'
if (-not (Test-Path -LiteralPath $verifier)) {
    throw "Release verifier was not found: $verifier"
}
$reportPath = Join-Path $outputFullPath 'AcceptanceReport-signed.json'
$verificationArguments = @{
    Package = $archivePath
    ExpectedVersion = $version
    ExpectedEdition = $Edition
    ReportPath = $reportPath
}
if ($Mode -eq 'Formal') {
    $verificationArguments.RequireTrustedPublisher = $true
}
& $verifier @verificationArguments
if ($LASTEXITCODE -ne 0) {
    throw 'Signed release verification failed.'
}

$summary = [pscustomobject]@{
    mode = $Mode
    version = $version
    edition = $Edition
    certificateSubject = $certificate.Subject
    certificateThumbprint = $certificate.Thumbprint
    certificateTrusted = $trust.trusted
    certificateChainStatus = $trust.statuses
    certificateStore = if ($certificateInCurrentUserStore) { 'CurrentUser/My' } else { 'LocalMachine/My' }
    codeSigningUsage = $hasCodeSigningUsage
    timestampUri = $TimestampUri
    magePath = $mage
    signToolPath = $signTool
    archivePath = $archivePath
    archiveSha256 = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash
    acceptanceReport = $reportPath
}
$summaryPath = Join-Path $outputFullPath 'SigningSummary.json'
$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $summaryPath -Encoding UTF8

Write-Host ''
Write-Host 'Agent for Excel signing completed.'
Write-Host ('Mode: ' + $Mode)
Write-Host ('Archive: ' + $archivePath)
Write-Host ('SHA-256: ' + $summary.archiveSha256)
Write-Host ('Summary: ' + $summaryPath)
