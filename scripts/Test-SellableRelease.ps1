[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Package,

    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$ExpectedVersion,

    [ValidateSet('Trial', 'Standard', 'Professional', 'Automation')]
    [string]$ExpectedEdition,

    [string]$ReportPath,

    [switch]$RequireTrustedPublisher
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Add-Check {
    param(
        [System.Collections.Generic.List[object]]$Checks,
        [string]$Id,
        [string]$Category,
        [ValidateSet('pass', 'warn', 'fail')]
        [string]$Status,
        [string]$Message,
        [object]$Evidence
    )
    $Checks.Add([pscustomobject]@{
        id = $Id
        category = $Category
        status = $Status
        message = $Message
        evidence = $Evidence
    })
}

function Get-ManifestCertificate {
    param([xml]$Manifest)

    $certificateNode = $Manifest.SelectSingleNode("//*[local-name()='X509Certificate']")
    if ($null -eq $certificateNode -or [string]::IsNullOrWhiteSpace($certificateNode.InnerText)) {
        return $null
    }

    $bytes = [Convert]::FromBase64String($certificateNode.InnerText.Trim())
    return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList (,$bytes)
}

function Get-CertificateSummary {
    param([System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    if ($null -eq $Certificate) {
        return $null
    }

    $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
    $chain.ChainPolicy.RevocationMode =
        [System.Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
    $trusted = $chain.Build($Certificate)
    $chainStatus = @($chain.ChainStatus | ForEach-Object { $_.Status.ToString() })
    $now = Get-Date
    return [pscustomobject]@{
        subject = $Certificate.Subject
        issuer = $Certificate.Issuer
        thumbprint = $Certificate.Thumbprint
        notBefore = $Certificate.NotBefore.ToString('o')
        notAfter = $Certificate.NotAfter.ToString('o')
        currentlyValid = ($now -ge $Certificate.NotBefore -and $now -le $Certificate.NotAfter)
        selfSigned = [string]::Equals(
            $Certificate.Subject, $Certificate.Issuer, [StringComparison]::OrdinalIgnoreCase)
        trustedByMachine = $trusted
        chainStatus = $chainStatus
    }
}

function Get-RegistryValue {
    param([string[]]$Paths, [string]$Name)

    foreach ($path in $Paths) {
        try {
            $value = (Get-ItemProperty -LiteralPath $path -ErrorAction Stop).$Name
            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                return $value
            }
        }
        catch { }
    }
    return $null
}

function Get-ExcelEnvironment {
    $clickToRunPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration'
    )
    $platform = Get-RegistryValue -Paths $clickToRunPaths -Name 'Platform'
    $clientFolder = Get-RegistryValue -Paths $clickToRunPaths -Name 'ClientFolder'

    $excelPath = $null
    try {
        $classesRoot = [Microsoft.Win32.Registry]::ClassesRoot
        $applicationKey = $classesRoot.OpenSubKey('Excel.Application\CLSID')
        if ($null -ne $applicationKey) {
            $clsid = [string]$applicationKey.GetValue('')
            $applicationKey.Dispose()
            if (-not [string]::IsNullOrWhiteSpace($clsid)) {
                $serverKey = $classesRoot.OpenSubKey(('CLSID\' + $clsid + '\LocalServer32'))
                if ($null -ne $serverKey) {
                    $rawPath = [string]$serverKey.GetValue('')
                    $serverKey.Dispose()
                    if (-not [string]::IsNullOrWhiteSpace($rawPath)) {
                        $match = [regex]::Match($rawPath, '(?i)(?:^"|^)([^"]*EXCEL\.EXE)')
                        if ($match.Success) {
                            $excelPath = $match.Groups[1].Value.Trim()
                        }
                    }
                }
            }
        }
    }
    catch { }

    if ([string]::IsNullOrWhiteSpace($platform) -and -not [string]::IsNullOrWhiteSpace($excelPath)) {
        $programFilesX86 = ${env:ProgramFiles(x86)}
        if (-not [string]::IsNullOrWhiteSpace($programFilesX86) -and
            $excelPath.StartsWith($programFilesX86, [StringComparison]::OrdinalIgnoreCase)) {
            $platform = 'x86'
        }
        else {
            $platform = 'x64'
        }
    }

    $version = $null
    if (-not [string]::IsNullOrWhiteSpace($excelPath) -and (Test-Path -LiteralPath $excelPath)) {
        $version = (Get-Item -LiteralPath $excelPath).VersionInfo.FileVersion
    }

    return [pscustomobject]@{
        installed = -not [string]::IsNullOrWhiteSpace($excelPath)
        bitness = $platform
        path = $excelPath
        version = $version
        clickToRunClientFolder = $clientFolder
    }
}

function Get-DotNet48Environment {
    $release = Get-RegistryValue -Paths @(
        'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\v4\Full'
    ) -Name 'Release'
    $releaseNumber = if ($null -eq $release) { 0 } else { [int]$release }
    return [pscustomobject]@{
        installed = $releaseNumber -ge 528040
        release = $releaseNumber
    }
}

function Get-VstoEnvironment {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R'
    )
    $installed = $false
    $version = $null
    foreach ($path in $paths) {
        try {
            $item = Get-ItemProperty -LiteralPath $path -ErrorAction Stop
            if ([int]$item.Install -eq 1) {
                $installed = $true
                if ($null -ne $item.Version) { $version = [string]$item.Version }
            }
        }
        catch { }
    }
    return [pscustomobject]@{
        installed = $installed
        version = $version
    }
}

function Find-Mage {
    $sdkRoot = Join-Path ${env:ProgramFiles(x86)} 'Microsoft SDKs\Windows'
    if (-not (Test-Path -LiteralPath $sdkRoot)) { return $null }
    return Get-ChildItem -LiteralPath $sdkRoot -Filter mage.exe -Recurse -ErrorAction SilentlyContinue |
        Select-Object -Last 1 -ExpandProperty FullName
}

$checks = New-Object 'System.Collections.Generic.List[object]'
$resolvedPackage = (Resolve-Path -LiteralPath $Package).Path
$packageItem = Get-Item -LiteralPath $resolvedPackage
$packageHash = if (-not $packageItem.PSIsContainer) {
    (Get-FileHash -LiteralPath $resolvedPackage -Algorithm SHA256).Hash
}
else {
    $null
}

$tempDirectory = $null
$packageRoot = $null
$sourceEntries = @()
try {
    if ($packageItem.PSIsContainer) {
        $packageRoot = $resolvedPackage
        $sourceEntries = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -File |
            ForEach-Object { $_.FullName.Substring($packageRoot.Length).TrimStart('\', '/') })
    }
    elseif ([string]::Equals($packageItem.Extension, '.zip', [StringComparison]::OrdinalIgnoreCase)) {
        $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
        $tempDirectory = Join-Path $tempRoot ('AgentForExcel-ReleaseCheck-' + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $tempDirectory | Out-Null
        Expand-Archive -LiteralPath $resolvedPackage -DestinationPath $tempDirectory
        $packageRoot = $tempDirectory
        $sourceEntries = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -File |
            ForEach-Object { $_.FullName.Substring($packageRoot.Length).TrimStart('\', '/') })
    }
    else {
        throw 'Package must be a release directory or a .zip archive.'
    }

    $publishRoot = if (Test-Path -LiteralPath (Join-Path $packageRoot 'publish')) {
        Join-Path $packageRoot 'publish'
    }
    else {
        $packageRoot
    }

    $required = [ordered]@{
        setup = Join-Path $publishRoot 'setup.exe'
        vsto = Join-Path $publishRoot 'AgentForExcel.vsto'
        applicationFiles = Join-Path $publishRoot 'Application Files'
        readme = Join-Path $packageRoot 'README.md'
        deliveryGuide = Join-Path $packageRoot 'SELLABLE_RELEASE.md'
    }
    foreach ($item in $required.GetEnumerator()) {
        $exists = Test-Path -LiteralPath $item.Value
        Add-Check $checks ('structure.' + $item.Key) 'package' $(if ($exists) { 'pass' } else { 'fail' }) `
            $(if ($exists) { "Found $($item.Key)." } else { "Missing $($item.Key)." }) `
            ([IO.Path]::GetFileName($item.Value))
    }

    $dangerousNames = @($sourceEntries | Where-Object {
        $_ -match '(?i)(^|[\\/])(\.env($|\.)|secrets?|credentials?|tokens?|api[-_]?keys?)([\\/\.]|$)' -or
        $_ -match '(?i)\.(pfx|p12|p8|privatekey|key|pem)$' -or
        $_ -match '(?i)(^|[\\/])(\.git|\.vs|obj|bin)([\\/]|$)'
    })
    Add-Check $checks 'security.sensitive_names' 'security' `
        $(if ($dangerousNames.Count -eq 0) { 'pass' } else { 'fail' }) `
        $(if ($dangerousNames.Count -eq 0) {
            'No secret-bearing or development-only file names were found.'
        } else {
            'Potentially sensitive or development-only files were found.'
        }) $dangerousNames

    $secretPattern = '(?i)(gh[pousr]_[A-Za-z0-9]{20,}|sk-[A-Za-z0-9_-]{20,}|AIza[0-9A-Za-z_-]{20,}|bearer\s+[A-Za-z0-9._-]{30,})'
    $secretFiles = New-Object 'System.Collections.Generic.List[string]'
    Get-ChildItem -LiteralPath $packageRoot -Recurse -File | Where-Object {
        $_.Length -le 2097152 -and $_.Extension -match '(?i)^\.(config|json|xml|md|txt|vsto|manifest)$'
    } | ForEach-Object {
        try {
            if ((Get-Content -LiteralPath $_.FullName -Raw -ErrorAction Stop) -match $secretPattern) {
                $secretFiles.Add($_.FullName.Substring($packageRoot.Length).TrimStart('\', '/'))
            }
        }
        catch { }
    }
    Add-Check $checks 'security.secret_patterns' 'security' `
        $(if ($secretFiles.Count -eq 0) { 'pass' } else { 'fail' }) `
        $(if ($secretFiles.Count -eq 0) {
            'No high-confidence token patterns were found.'
        } else {
            'Potential credentials were found; matched values are intentionally not displayed.'
        }) @($secretFiles)

    $deploymentInfo = $null
    $applicationInfo = $null
    $deploymentCertificate = $null
    $applicationCertificate = $null
    $applicationManifestPath = $null

    if (Test-Path -LiteralPath $required.vsto) {
        [xml]$deployment = Get-Content -LiteralPath $required.vsto -Raw
        $deploymentIdentity = $deployment.SelectSingleNode(
            "/*[local-name()='assembly']/*[local-name()='assemblyIdentity']")
        $description = $deployment.SelectSingleNode(
            "/*[local-name()='assembly']/*[local-name()='description']")
        $publisherIdentity = $deployment.SelectSingleNode(
            "/*[local-name()='assembly']/*[local-name()='publisherIdentity']")
        $framework = $deployment.SelectSingleNode("//*[local-name()='framework']")
        $dependentAssembly = $deployment.SelectSingleNode(
            "/*[local-name()='assembly']/*[local-name()='dependency']/*[local-name()='dependentAssembly']")
        $deploymentInfo = [pscustomobject]@{
            version = [string]$deploymentIdentity.version
            architecture = [string]$deploymentIdentity.processorArchitecture
            publisher = [string]$description.publisher
            product = [string]$description.product
            publisherIdentity = [string]$publisherIdentity.name
            targetFramework = [string]$framework.targetVersion
            applicationCodebase = [string]$dependentAssembly.codebase
        }

        $applicationManifestPath = Join-Path $publishRoot (
            $deploymentInfo.applicationCodebase.Replace('/', '\'))
        $applicationExists = Test-Path -LiteralPath $applicationManifestPath
        Add-Check $checks 'manifest.application_reference' 'manifest' `
            $(if ($applicationExists) { 'pass' } else { 'fail' }) `
            $(if ($applicationExists) {
                'Deployment manifest points to an existing application manifest.'
            } else {
                'Deployment manifest application codebase is missing.'
            }) $deploymentInfo.applicationCodebase

        $deploymentCertificate = Get-ManifestCertificate $deployment
        if ($applicationExists) {
            [xml]$application = Get-Content -LiteralPath $applicationManifestPath -Raw
            $applicationIdentity = $application.SelectSingleNode(
                "/*[local-name()='assembly']/*[local-name()='assemblyIdentity']")
            $entryPoint = $application.SelectSingleNode("//*[local-name()='entryPoint' and @class]")
            $addIn = $application.SelectSingleNode("//*[local-name()='appAddIn' and @application]")
            $applicationInfo = [pscustomobject]@{
                version = [string]$applicationIdentity.version
                architecture = [string]$applicationIdentity.processorArchitecture
                entryPoint = if ($null -eq $entryPoint) { $null } else { $entryPoint.GetAttribute('class') }
                application = if ($null -eq $addIn) { $null } else { $addIn.GetAttribute('application') }
                loadBehavior = if ($null -eq $addIn) { $null } else { $addIn.GetAttribute('loadBehavior') }
                keyName = if ($null -eq $addIn) { $null } else { $addIn.GetAttribute('keyName') }
            }
            $applicationCertificate = Get-ManifestCertificate $application
        }
    }

    $versionMatches = $null -ne $deploymentInfo -and $null -ne $applicationInfo -and
        [string]::Equals($deploymentInfo.version, $applicationInfo.version, [StringComparison]::Ordinal)
    if (-not [string]::IsNullOrWhiteSpace($ExpectedVersion)) {
        $versionMatches = $versionMatches -and
            [string]::Equals($deploymentInfo.version, $ExpectedVersion, [StringComparison]::Ordinal)
    }
    Add-Check $checks 'manifest.version' 'manifest' `
        $(if ($versionMatches) { 'pass' } else { 'fail' }) `
        $(if ($versionMatches) {
            'Deployment and application manifest versions match.'
        } else {
            'Deployment/application/expected versions do not match.'
        }) ([pscustomobject]@{
            deployment = if ($null -eq $deploymentInfo) { $null } else { $deploymentInfo.version }
            application = if ($null -eq $applicationInfo) { $null } else { $applicationInfo.version }
            expected = $ExpectedVersion
        })

    if (-not [string]::IsNullOrWhiteSpace($ExpectedEdition)) {
        $editionMatches = $resolvedPackage.IndexOf(
            $ExpectedEdition, [StringComparison]::OrdinalIgnoreCase) -ge 0
        Add-Check $checks 'manifest.edition_package_name' 'manifest' `
            $(if ($editionMatches) { 'pass' } else { 'fail' }) `
            $(if ($editionMatches) {
                'Package name contains the expected edition.'
            } else {
                'Package name does not contain the expected edition.'
            }) $ExpectedEdition
    }

    $deploymentCertificateSummary = Get-CertificateSummary $deploymentCertificate
    $applicationCertificateSummary = Get-CertificateSummary $applicationCertificate
    $certificatesPresent = $null -ne $deploymentCertificateSummary -and $null -ne $applicationCertificateSummary
    Add-Check $checks 'signature.manifest_certificates' 'signature' `
        $(if ($certificatesPresent) { 'pass' } else { 'fail' }) `
        $(if ($certificatesPresent) {
            'Both ClickOnce manifests contain signing certificates.'
        } else {
            'A ClickOnce manifest signing certificate is missing.'
        }) $null

    $manifestCertificatesValid = $certificatesPresent -and
        $deploymentCertificateSummary.currentlyValid -and $applicationCertificateSummary.currentlyValid
    Add-Check $checks 'signature.certificate_dates' 'signature' `
        $(if ($manifestCertificatesValid) { 'pass' } else { 'fail' }) `
        $(if ($manifestCertificatesValid) {
            'Manifest certificates are within their validity periods.'
        } else {
            'A manifest certificate is expired or not yet valid.'
        }) $null

    $manifestsTrusted = $certificatesPresent -and
        $deploymentCertificateSummary.trustedByMachine -and
        $applicationCertificateSummary.trustedByMachine -and
        -not $deploymentCertificateSummary.selfSigned -and
        -not $applicationCertificateSummary.selfSigned
    $trustStatus = if ($manifestsTrusted) {
        'pass'
    }
    elseif ($RequireTrustedPublisher) {
        'fail'
    }
    else {
        'warn'
    }
    Add-Check $checks 'signature.trusted_publisher' 'signature' $trustStatus `
        $(if ($manifestsTrusted) {
            'Manifest publisher chains to a trusted non-self-signed certificate.'
        } else {
            'Manifest publisher is not trusted for external distribution on this machine.'
        }) $null

    $setupSignature = if (Test-Path -LiteralPath $required.setup) {
        Get-AuthenticodeSignature -LiteralPath $required.setup
    }
    else {
        $null
    }
    $setupSigned = $null -ne $setupSignature -and $setupSignature.Status -eq 'Valid'
    $setupStatus = if ($setupSigned) { 'pass' } elseif ($RequireTrustedPublisher) { 'fail' } else { 'warn' }
    Add-Check $checks 'signature.setup_authenticode' 'signature' $setupStatus `
        $(if ($setupSigned) {
            'setup.exe has a valid Authenticode signature.'
        } else {
            'setup.exe does not have a trusted valid Authenticode signature.'
        }) $(if ($null -eq $setupSignature) { $null } else { [string]$setupSignature.Status })

    $magePath = Find-Mage
    $mageResults = @()
    if (-not [string]::IsNullOrWhiteSpace($magePath)) {
        foreach ($manifestPath in @($required.vsto, $applicationManifestPath)) {
            if ([string]::IsNullOrWhiteSpace($manifestPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
                continue
            }
            $output = & $magePath -Verify $manifestPath 2>&1
            $mageResults += [pscustomobject]@{
                file = [IO.Path]::GetFileName($manifestPath)
                valid = $LASTEXITCODE -eq 0
                message = ($output -join ' ')
            }
        }
        $mageValid = $mageResults.Count -eq 2 -and @($mageResults | Where-Object { -not $_.valid }).Count -eq 0
        Add-Check $checks 'signature.mage_verify' 'signature' `
            $(if ($mageValid) { 'pass' } else { 'fail' }) `
            $(if ($mageValid) {
                'Mage verified both ClickOnce manifest signatures and hashes.'
            } else {
                'Mage could not verify all ClickOnce manifests.'
            }) @($mageResults)
    }
    else {
        Add-Check $checks 'signature.mage_verify' 'signature' 'warn' `
            'Mage is not installed; cryptographic manifest verification was skipped on this machine.' $null
    }

    $excel = Get-ExcelEnvironment
    $dotNet = Get-DotNet48Environment
    $vsto = Get-VstoEnvironment
    Add-Check $checks 'environment.excel' 'environment' `
        $(if ($excel.installed) { 'pass' } else { 'fail' }) `
        $(if ($excel.installed) { 'Desktop Excel was detected.' } else { 'Desktop Excel was not detected.' }) `
        ([pscustomobject]@{ bitness = $excel.bitness; version = $excel.version; path = $excel.path })
    Add-Check $checks 'environment.dotnet48' 'environment' `
        $(if ($dotNet.installed) { 'pass' } else { 'fail' }) `
        $(if ($dotNet.installed) { '.NET Framework 4.8 or later was detected.' } else { '.NET Framework 4.8 was not detected.' }) `
        $dotNet.release
    Add-Check $checks 'environment.vsto_runtime' 'environment' `
        $(if ($vsto.installed) { 'pass' } else { 'warn' }) `
        $(if ($vsto.installed) {
            'VSTO Runtime was detected.'
        } else {
            'VSTO Runtime was not detected; setup.exe must install it before the add-in can run.'
        }) $vsto.version

    $failCount = @($checks | Where-Object { $_.status -eq 'fail' }).Count
    $warnCount = @($checks | Where-Object { $_.status -eq 'warn' }).Count
    $passCount = @($checks | Where-Object { $_.status -eq 'pass' }).Count
    $packageFailures = @($checks | Where-Object {
        $_.status -eq 'fail' -and $_.category -in @('package', 'manifest', 'security')
    }).Count
    $machineFailures = @($checks | Where-Object {
        $_.status -eq 'fail' -and $_.category -eq 'environment'
    }).Count
    $externalReady = $failCount -eq 0 -and $manifestsTrusted -and $setupSigned
    $checkArray = @($checks | ForEach-Object { $_ })

    $report = [pscustomobject]@{
        schemaVersion = 1
        generatedAt = (Get-Date).ToString('o')
        computerName = $env:COMPUTERNAME
        package = [pscustomobject]@{
            path = $resolvedPackage
            sha256 = $packageHash
            entryCount = $sourceEntries.Count
            expectedVersion = $ExpectedVersion
            expectedEdition = $ExpectedEdition
        }
        environment = [pscustomobject]@{
            os = [Environment]::OSVersion.VersionString
            is64BitOperatingSystem = [Environment]::Is64BitOperatingSystem
            excel = $excel
            dotNetFramework = $dotNet
            vstoRuntime = $vsto
        }
        manifests = [pscustomobject]@{
            deployment = $deploymentInfo
            application = $applicationInfo
        }
        signatures = [pscustomobject]@{
            deploymentCertificate = $deploymentCertificateSummary
            applicationCertificate = $applicationCertificateSummary
            setupAuthenticodeStatus = if ($null -eq $setupSignature) { $null } else { [string]$setupSignature.Status }
            mage = $mageResults
        }
        checks = $checkArray
        summary = [pscustomobject]@{
            pass = $passCount
            warn = $warnCount
            fail = $failCount
            packageStructureReady = $packageFailures -eq 0
            machineEnvironmentReady = $machineFailures -eq 0
            externalDeliveryReady = $externalReady
        }
    }

    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        $reportDirectory = if ($packageItem.PSIsContainer) {
            $resolvedPackage
        }
        else {
            $packageItem.DirectoryName
        }
        $ReportPath = Join-Path $reportDirectory (
            'AcceptanceReport-' + $env:COMPUTERNAME + '-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.json')
    }
    elseif (-not [IO.Path]::IsPathRooted($ReportPath)) {
        $ReportPath = Join-Path (Get-Location).Path $ReportPath
    }
    $reportPathParent = Split-Path -Parent $ReportPath
    if (-not (Test-Path -LiteralPath $reportPathParent)) {
        New-Item -ItemType Directory -Path $reportPathParent | Out-Null
    }
    $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ReportPath -Encoding UTF8

    Write-Host ''
    Write-Host 'Agent for Excel release acceptance'
    Write-Host ('Package: ' + $resolvedPackage)
    Write-Host ('SHA-256: ' + $packageHash)
    Write-Host ('Excel: ' + $(if ($excel.installed) {
        ($excel.version + ' / ' + $excel.bitness)
    } else {
        'not detected'
    }))
    $checks | Select-Object status, category, id, message | Format-Table -AutoSize
    Write-Host ("Summary: PASS={0}, WARN={1}, FAIL={2}" -f $passCount, $warnCount, $failCount)
    Write-Host ('Package structure ready: ' + $report.summary.packageStructureReady)
    Write-Host ('Machine environment ready: ' + $report.summary.machineEnvironmentReady)
    Write-Host ('External delivery ready: ' + $report.summary.externalDeliveryReady)
    Write-Host ('JSON report: ' + (Resolve-Path -LiteralPath $ReportPath).Path)

    if ($failCount -gt 0) { exit 2 }
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($tempDirectory) -and (Test-Path -LiteralPath $tempDirectory)) {
        $resolvedTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
        $resolvedTempDirectory = [IO.Path]::GetFullPath($tempDirectory)
        if ($resolvedTempDirectory.StartsWith($resolvedTempRoot, [StringComparison]::OrdinalIgnoreCase) -and
            [IO.Path]::GetFileName($resolvedTempDirectory).StartsWith(
                'AgentForExcel-ReleaseCheck-', [StringComparison]::Ordinal)) {
            Remove-Item -LiteralPath $resolvedTempDirectory -Recurse -Force
        }
    }
}
