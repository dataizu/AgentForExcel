[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FormalGateReport,

    [Parameter(Mandatory = $true)]
    [string]$X86MachineReport,

    [Parameter(Mandatory = $true)]
    [string]$X64MachineReport,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$ExpectedVersion,

    [Parameter(Mandatory = $true)]
    [string]$ReportPath,

    [Parameter(Mandatory = $true)]
    [string]$X86PostUninstallReport,

    [Parameter(Mandatory = $true)]
    [string]$X64PostUninstallReport,

    [ValidateRange(1, 365)]
    [int]$MaxEvidenceAgeDays = 30
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Add-Check {
    param(
        [Collections.Generic.List[object]]$Checks,
        [string]$Id,
        [ValidateSet('pass', 'warn', 'fail')]
        [string]$Status,
        [string]$Message,
        [object]$Evidence
    )
    $Checks.Add([pscustomobject]@{
        id = $Id
        status = $Status
        message = $Message
        evidence = $Evidence
    })
}

function Write-InputFailure {
    param([string[]]$Messages)
    $failureChecks = @($Messages | ForEach-Object {
        [pscustomobject]@{
            id = 'evidence.input'
            status = 'fail'
            message = $_
            evidence = $null
        }
    })
    $failureReport = [pscustomobject]@{
        schemaVersion = 1
        kind = 'AgentForExcelReleaseEvidenceGate'
        generatedAt = (Get-Date).ToString('o')
        expectedVersion = $ExpectedVersion
        evidence = [pscustomobject]@{
            formalGate = $FormalGateReport
            x86PostInstall = $X86MachineReport
            x64PostInstall = $X64MachineReport
            x86PostUninstall = $X86PostUninstallReport
            x64PostUninstall = $X64PostUninstallReport
        }
        checks = $failureChecks
        gaps = @($Messages)
        summary = [pscustomobject]@{
            pass = 0
            warn = 0
            fail = $failureChecks.Count
            releaseReady = $false
        }
    }
    $failureReportPath = [IO.Path]::GetFullPath($ReportPath)
    $failureDirectory = Split-Path -Parent $failureReportPath
    if (-not (Test-Path -LiteralPath $failureDirectory)) {
        New-Item -ItemType Directory -Path $failureDirectory | Out-Null
    }
    $failureReport | ConvertTo-Json -Depth 8 |
        Set-Content -LiteralPath $failureReportPath -Encoding UTF8
    Write-Host ''
    Write-Host 'Agent for Excel release evidence gate'
    Write-Host 'Required evidence is missing or invalid.'
    $Messages | ForEach-Object { Write-Host ('- ' + $_) }
    Write-Host ('Evidence report: ' + $failureReportPath)
    exit 2
}

function Read-JsonReport {
    param([string]$Path, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        throw "$Label report was not found: $Path"
    }
    $resolved = (Resolve-Path -LiteralPath $Path).Path
    try {
        $data = Get-Content -LiteralPath $resolved -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "$Label report is not valid JSON: $resolved"
    }
    return [pscustomobject]@{ path = $resolved; data = $data }
}

function Test-Freshness {
    param([object]$Report, [int]$MaximumDays)
    $generated = [DateTimeOffset]::MinValue
    if ($null -eq $Report -or
        -not [DateTimeOffset]::TryParse([string]$Report.generatedAt, [ref]$generated)) {
        return $false
    }
    $now = [DateTimeOffset]::Now
    return $generated -le $now.AddHours(24) -and $generated -ge $now.AddDays(-$MaximumDays)
}

function Test-MachineReport {
    param(
        [Collections.Generic.List[object]]$Checks,
        [string]$Prefix,
        [object]$Envelope,
        [string]$RequiredBitness,
        [string]$Version,
        [int]$MaximumDays,
        [string]$ExpectedPackageHash
    )
    $report = $Envelope.data
    $schemaValid = [string]$report.kind -eq 'AgentForExcelCleanMachineAcceptance' -and
        [string]$report.phase -eq 'PostInstall'
    Add-Check $Checks ($Prefix + '.schema_and_phase') `
        $(if ($schemaValid) { 'pass' } else { 'fail' }) `
        $(if ($schemaValid) {
            'Machine report schema and PostInstall phase are valid.'
        } else {
            'Machine report schema or phase is invalid.'
        }) $Envelope.path

    $versionValid = [string]$report.expectedVersion -eq $Version
    Add-Check $Checks ($Prefix + '.version') `
        $(if ($versionValid) { 'pass' } else { 'fail' }) `
        $(if ($versionValid) {
            'Machine report version matches.'
        } else {
            'Machine report version does not match.'
        }) ([pscustomobject]@{ expected = $Version; actual = $report.expectedVersion })

    $bitnessValid = [string]$report.machine.excel.bitness -eq $RequiredBitness
    Add-Check $Checks ($Prefix + '.excel_bitness') `
        $(if ($bitnessValid) { 'pass' } else { 'fail' }) `
        $(if ($bitnessValid) {
            'Excel bitness matches the required lane.'
        } else {
            'Excel bitness does not match the required lane.'
        }) ([pscustomobject]@{ expected = $RequiredBitness; actual = $report.machine.excel.bitness })

    $ready = $schemaValid -and [int]$report.summary.fail -eq 0 -and
        [bool]$report.summary.ready
    Add-Check $Checks ($Prefix + '.postinstall_ready') `
        $(if ($ready) { 'pass' } else { 'fail' }) `
        $(if ($ready) {
            'Post-install acceptance is ready.'
        } else {
            'Post-install acceptance is not ready.'
        }) $report.summary

    $fresh = Test-Freshness $report $MaximumDays
    Add-Check $Checks ($Prefix + '.freshness') `
        $(if ($fresh) { 'pass' } else { 'fail' }) `
        $(if ($fresh) {
            'Machine evidence is within the allowed age.'
        } else {
            'Machine evidence is missing a valid timestamp or is stale.'
        }) $report.generatedAt

    $hashValid = -not [string]::IsNullOrWhiteSpace($ExpectedPackageHash) -and
        [string]$report.package.sha256 -eq $ExpectedPackageHash
    Add-Check $Checks ($Prefix + '.package_hash') `
        $(if ($hashValid) { 'pass' } else { 'fail' }) `
        $(if ($hashValid) {
            'Machine report used the signed package hash.'
        } else {
            'Machine report package hash differs from the signed package.'
        }) ([pscustomobject]@{
            expected = $ExpectedPackageHash
            actual = $report.package.sha256
        })
}

function Test-UninstallReport {
    param(
        [Collections.Generic.List[object]]$Checks,
        [string]$Prefix,
        [object]$Envelope,
        [object]$PostInstallReport,
        [string]$RequiredBitness,
        [string]$Version,
        [int]$MaximumDays
    )
    $report = $Envelope.data
    $valid = [string]$report.kind -eq 'AgentForExcelCleanMachineAcceptance' -and
        [string]$report.phase -eq 'PostUninstall' -and
        [string]$report.expectedVersion -eq $Version -and
        [string]$report.machine.excel.bitness -eq $RequiredBitness -and
        [string]$report.machine.computerName -eq [string]$PostInstallReport.machine.computerName -and
        [int]$report.summary.fail -eq 0 -and [bool]$report.summary.ready -and
        (Test-Freshness $report $MaximumDays)
    Add-Check $Checks ($Prefix + '.postuninstall_ready') `
        $(if ($valid) { 'pass' } else { 'fail' }) `
        $(if ($valid) {
            'Post-uninstall evidence is valid.'
        } else {
            'Post-uninstall evidence is invalid or does not match the machine.'
        }) $Envelope.path
}

$requiredInputs = @(
    [pscustomobject]@{ label = 'Formal gate'; path = $FormalGateReport },
    [pscustomobject]@{ label = 'x86 machine'; path = $X86MachineReport },
    [pscustomobject]@{ label = 'x64 machine'; path = $X64MachineReport }
)
$requiredInputs += [pscustomobject]@{
    label = 'x86 post-uninstall'
    path = $X86PostUninstallReport
}
$requiredInputs += [pscustomobject]@{
    label = 'x64 post-uninstall'
    path = $X64PostUninstallReport
}
$missingInputs = @($requiredInputs | Where-Object {
    [string]::IsNullOrWhiteSpace([string]$_.path) -or
    -not (Test-Path -LiteralPath $_.path)
} | ForEach-Object {
    "$($_.label) report was not found: $($_.path)"
})
if ($missingInputs.Count -gt 0) {
    Write-InputFailure $missingInputs
}

$checks = New-Object 'Collections.Generic.List[object]'
$formalEnvelope = Read-JsonReport $FormalGateReport 'Formal gate'
$x86Envelope = Read-JsonReport $X86MachineReport 'x86 machine'
$x64Envelope = Read-JsonReport $X64MachineReport 'x64 machine'
$formal = $formalEnvelope.data

$formalVersionValid = [string]$formal.package.expectedVersion -eq $ExpectedVersion -and
    [string]$formal.manifests.deployment.version -eq $ExpectedVersion -and
    [string]$formal.manifests.application.version -eq $ExpectedVersion
Add-Check $checks 'formal.version' `
    $(if ($formalVersionValid) { 'pass' } else { 'fail' }) `
    $(if ($formalVersionValid) {
        'Formal signed package version matches.'
    } else {
        'Formal signed package version does not match.'
    }) $formal.package.expectedVersion

$formalReady = [int]$formal.summary.fail -eq 0 -and
    [bool]$formal.summary.externalDeliveryReady
Add-Check $checks 'formal.external_delivery_ready' `
    $(if ($formalReady) { 'pass' } else { 'fail' }) `
    $(if ($formalReady) {
        'Formal signature gate passed.'
    } else {
        'Formal signature gate did not pass.'
    }) $formal.summary

$formalFresh = Test-Freshness $formal $MaxEvidenceAgeDays
Add-Check $checks 'formal.freshness' `
    $(if ($formalFresh) { 'pass' } else { 'fail' }) `
    $(if ($formalFresh) {
        'Formal signature evidence is within the allowed age.'
    } else {
        'Formal signature evidence is stale or invalid.'
    }) $formal.generatedAt

$packageHash = [string]$formal.package.sha256
Test-MachineReport $checks 'x86' $x86Envelope 'x86' $ExpectedVersion `
    $MaxEvidenceAgeDays $packageHash
Test-MachineReport $checks 'x64' $x64Envelope 'x64' $ExpectedVersion `
    $MaxEvidenceAgeDays $packageHash

$differentMachines = -not [string]::IsNullOrWhiteSpace(
    [string]$x86Envelope.data.machine.computerName) -and
    -not [string]::IsNullOrWhiteSpace(
        [string]$x64Envelope.data.machine.computerName) -and
    -not [string]::Equals(
        [string]$x86Envelope.data.machine.computerName,
        [string]$x64Envelope.data.machine.computerName,
        [StringComparison]::OrdinalIgnoreCase)
Add-Check $checks 'machines.distinct' `
    $(if ($differentMachines) { 'pass' } else { 'fail' }) `
    $(if ($differentMachines) {
        'x86 and x64 evidence came from different machines.'
    } else {
        'x86 and x64 evidence must come from different machines.'
    }) ([pscustomobject]@{
        x86 = $x86Envelope.data.machine.computerName
        x64 = $x64Envelope.data.machine.computerName
    })

$x86Uninstall = Read-JsonReport $X86PostUninstallReport 'x86 post-uninstall'
Test-UninstallReport $checks 'x86' $x86Uninstall $x86Envelope.data 'x86' `
    $ExpectedVersion $MaxEvidenceAgeDays
$x64Uninstall = Read-JsonReport $X64PostUninstallReport 'x64 post-uninstall'
Test-UninstallReport $checks 'x64' $x64Uninstall $x64Envelope.data 'x64' `
    $ExpectedVersion $MaxEvidenceAgeDays

$checkArray = @($checks | ForEach-Object { $_ })
$failCount = @($checkArray | Where-Object { $_.status -eq 'fail' }).Count
$warnCount = @($checkArray | Where-Object { $_.status -eq 'warn' }).Count
$passCount = @($checkArray | Where-Object { $_.status -eq 'pass' }).Count
$gaps = @($checkArray | Where-Object { $_.status -eq 'fail' } |
    ForEach-Object { $_.id + ': ' + $_.message })
$report = [pscustomobject]@{
    schemaVersion = 1
    kind = 'AgentForExcelReleaseEvidenceGate'
    generatedAt = (Get-Date).ToString('o')
    expectedVersion = $ExpectedVersion
    evidence = [pscustomobject]@{
        formalGate = $formalEnvelope.path
        x86PostInstall = $x86Envelope.path
        x64PostInstall = $x64Envelope.path
        x86PostUninstall = $X86PostUninstallReport
        x64PostUninstall = $X64PostUninstallReport
    }
    checks = $checkArray
    gaps = $gaps
    summary = [pscustomobject]@{
        pass = $passCount
        warn = $warnCount
        fail = $failCount
        releaseReady = $failCount -eq 0
    }
}
$reportFullPath = [IO.Path]::GetFullPath($ReportPath)
$reportDirectory = Split-Path -Parent $reportFullPath
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -ItemType Directory -Path $reportDirectory | Out-Null
}
$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $reportFullPath -Encoding UTF8

Write-Host ''
Write-Host 'Agent for Excel release evidence gate'
$checkArray | Select-Object status, id, message | Format-Table -AutoSize
Write-Host ("Summary: PASS={0}, WARN={1}, FAIL={2}" -f $passCount, $warnCount, $failCount)
Write-Host ('Release ready: ' + $report.summary.releaseReady)
Write-Host ('Evidence report: ' + $reportFullPath)

if ($failCount -gt 0) { exit 2 }
exit 0
