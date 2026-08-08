[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Package,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$ExpectedVersion,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Trial', 'Standard', 'Professional', 'Automation')]
    [string]$ExpectedEdition,

    [Parameter(Mandatory = $true)]
    [ValidateSet('PreInstall', 'PostInstall', 'PostUninstall')]
    [string]$Phase,

    [Parameter(Mandatory = $true)]
    [string]$ReportPath,

    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$ExpectedAssemblyVersion = '1.1.0.0',

    [switch]$RequireTrustedPublisher
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

function Get-RegistryValue {
    param([string[]]$Paths, [string]$Name)
    foreach ($path in $Paths) {
        try {
            $value = (Get-ItemProperty -LiteralPath $path -ErrorAction Stop).$Name
            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                return [string]$value
            }
        }
        catch { }
    }
    return $null
}

function Get-MachineInfo {
    $configurationPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration'
    )
    $excelBitness = Get-RegistryValue $configurationPaths 'Platform'
    $excelVersion = Get-RegistryValue $configurationPaths 'VersionToReport'
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
                    $match = [regex]::Match($rawPath, '(?i)(?:^"|^)([^"]*EXCEL\.EXE)')
                    if ($match.Success) { $excelPath = $match.Groups[1].Value.Trim() }
                }
            }
        }
    }
    catch { }
    if ([string]::IsNullOrWhiteSpace($excelBitness) -and
        -not [string]::IsNullOrWhiteSpace($excelPath)) {
        $programFilesX86 = ${env:ProgramFiles(x86)}
        $excelBitness = if (-not [string]::IsNullOrWhiteSpace($programFilesX86) -and
            $excelPath.StartsWith($programFilesX86, [StringComparison]::OrdinalIgnoreCase)) {
            'x86'
        }
        else {
            'x64'
        }
    }
    if ([string]::IsNullOrWhiteSpace($excelVersion) -and
        -not [string]::IsNullOrWhiteSpace($excelPath) -and
        (Test-Path -LiteralPath $excelPath)) {
        $excelVersion = (Get-Item -LiteralPath $excelPath).VersionInfo.FileVersion
    }
    return [pscustomobject]@{
        computerName = $env:COMPUTERNAME
        osVersion = [Environment]::OSVersion.VersionString
        os64Bit = [Environment]::Is64BitOperatingSystem
        excel = [pscustomobject]@{
            installed = -not [string]::IsNullOrWhiteSpace($excelPath)
            bitness = $excelBitness
            version = $excelVersion
            path = $excelPath
        }
    }
}

function Invoke-ChildPowerShell {
    param([string]$ScriptPath, [string[]]$Arguments)
    $powershell = Join-Path $PSHOME 'powershell.exe'
    if (-not (Test-Path -LiteralPath $powershell)) {
        $powershell = (Get-Command powershell.exe -ErrorAction Stop).Source
    }
    & $powershell -NoProfile -ExecutionPolicy Bypass -File $ScriptPath @Arguments | Out-Host
    $exitCode = $LASTEXITCODE
    return [int]$exitCode
}

function Release-ComObject {
    param([object]$Value)
    try {
        if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
        }
    }
    catch { }
}

$resolvedPackage = (Resolve-Path -LiteralPath $Package -ErrorAction Stop).Path
$packageHash = if ((Get-Item -LiteralPath $resolvedPackage).PSIsContainer) {
    $null
}
else {
    (Get-FileHash -LiteralPath $resolvedPackage -Algorithm SHA256).Hash
}
$reportFullPath = [IO.Path]::GetFullPath($ReportPath)
$reportDirectory = Split-Path -Parent $reportFullPath
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -ItemType Directory -Path $reportDirectory | Out-Null
}
$childReportPath = Join-Path $reportDirectory (
    [IO.Path]::GetFileNameWithoutExtension($reportFullPath) + '-child.json')
$checks = New-Object 'Collections.Generic.List[object]'
$machine = Get-MachineInfo
$childReport = $null
$childExitCode = $null

if ($Phase -eq 'PreInstall') {
    $verifier = Join-Path $PSScriptRoot 'Test-SellableRelease.ps1'
    $arguments = @(
        '-Package', $resolvedPackage,
        '-ExpectedVersion', $ExpectedVersion,
        '-ExpectedEdition', $ExpectedEdition,
        '-ReportPath', $childReportPath
    )
    if ($RequireTrustedPublisher) { $arguments += '-RequireTrustedPublisher' }
    $childExitCode = Invoke-ChildPowerShell $verifier $arguments
    if (Test-Path -LiteralPath $childReportPath) {
        $childReport = Get-Content -LiteralPath $childReportPath -Raw -Encoding UTF8 |
            ConvertFrom-Json
    }
    $ready = $null -ne $childReport -and [int]$childReport.summary.fail -eq 0
    if ($ready -and $RequireTrustedPublisher) {
        $ready = [bool]$childReport.summary.externalDeliveryReady
    }
    elseif ($ready) {
        $ready = [bool]$childReport.summary.packageStructureReady -and
            [bool]$childReport.summary.machineEnvironmentReady
    }
    Add-Check $checks 'preinstall.package_and_environment' `
        $(if ($ready) { 'pass' } else { 'fail' }) `
        $(if ($ready) {
            'Package and machine pre-install checks passed.'
        } else {
            'Package or machine pre-install checks failed.'
        }) ([pscustomobject]@{ childExitCode = $childExitCode; childReport = $childReportPath })
}
elseif ($Phase -eq 'PostInstall') {
    $verifier = Join-Path $PSScriptRoot 'Test-InstalledAddIn.ps1'
    $arguments = @(
        '-ExpectedDeploymentVersion', $ExpectedVersion,
        '-ExpectedAssemblyVersion', $ExpectedAssemblyVersion,
        '-ReportPath', $childReportPath
    )
    $childExitCode = Invoke-ChildPowerShell $verifier $arguments
    if (Test-Path -LiteralPath $childReportPath) {
        $childReport = Get-Content -LiteralPath $childReportPath -Raw -Encoding UTF8 |
            ConvertFrom-Json
    }
    $ready = $null -ne $childReport -and
        [int]$childReport.summary.fail -eq 0 -and
        [bool]$childReport.summary.installedAddInReady
    Add-Check $checks 'postinstall.loaded_addin' `
        $(if ($ready) { 'pass' } else { 'fail' }) `
        $(if ($ready) {
            'Installed add-in acceptance passed.'
        } else {
            'Installed add-in acceptance failed.'
        }) ([pscustomobject]@{ childExitCode = $childExitCode; childReport = $childReportPath })
}
else {
    $existingExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
    if ($existingExcel.Count -gt 0) {
        Add-Check $checks 'postuninstall.excel_closed' 'fail' `
            'Close all Excel windows before post-uninstall verification.' `
            @($existingExcel | Select-Object Id, ProcessName)
    }
    else {
        Add-Check $checks 'postuninstall.excel_closed' 'pass' `
            'No user Excel process was running.' $null
        $registrationPaths = @(
            'HKCU:\Software\Microsoft\Office\Excel\Addins\AgentForExcel',
            'HKLM:\Software\Microsoft\Office\Excel\Addins\AgentForExcel',
            'HKLM:\Software\WOW6432Node\Microsoft\Office\Excel\Addins\AgentForExcel'
        )
        $remainingRegistrations = @($registrationPaths | Where-Object {
            Test-Path -LiteralPath $_
        })
        Add-Check $checks 'postuninstall.registration_removed' `
            $(if ($remainingRegistrations.Count -eq 0) { 'pass' } else { 'fail' }) `
            $(if ($remainingRegistrations.Count -eq 0) {
                'Agent add-in registration is absent.'
            } else {
                'Agent add-in registration still exists.'
            }) $remainingRegistrations

        $excel = $null
        $addIn = $null
        $discovered = $false
        $excelLaunchSucceeded = $false
        try {
            $excel = New-Object -ComObject Excel.Application
            $excelLaunchSucceeded = $true
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            try {
                $addIn = $excel.COMAddIns.Item('AgentForExcel')
                $discovered = $null -ne $addIn
            }
            catch {
                $discovered = $false
            }
        }
        catch {
            Add-Check $checks 'postuninstall.com_discovery' 'fail' `
                'Excel could not be launched for post-uninstall verification.' `
                $_.Exception.Message
        }
        finally {
            try { if ($null -ne $excel) { $excel.Quit() } } catch { }
            Release-ComObject $addIn
            Release-ComObject $excel
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
        if ($excelLaunchSucceeded) {
            Add-Check $checks 'postuninstall.com_addin_removed' `
                $(if (-not $discovered) { 'pass' } else { 'fail' }) `
                $(if (-not $discovered) {
                    'Excel no longer discovers the Agent COM add-in.'
                } else {
                    'Excel still discovers the Agent COM add-in.'
                }) $null
        }
    }
}

$checkArray = @($checks | ForEach-Object { $_ })
$failCount = @($checkArray | Where-Object { $_.status -eq 'fail' }).Count
$warnCount = @($checkArray | Where-Object { $_.status -eq 'warn' }).Count
$passCount = @($checkArray | Where-Object { $_.status -eq 'pass' }).Count
$report = [pscustomobject]@{
    schemaVersion = 1
    kind = 'AgentForExcelCleanMachineAcceptance'
    generatedAt = (Get-Date).ToString('o')
    phase = $Phase
    expectedVersion = $ExpectedVersion
    expectedEdition = $ExpectedEdition
    package = [pscustomobject]@{
        path = $resolvedPackage
        sha256 = $packageHash
    }
    machine = $machine
    childReport = if ($null -eq $childReport) { $null } else {
        [pscustomobject]@{
            path = $childReportPath
            exitCode = $childExitCode
            summary = $childReport.summary
        }
    }
    checks = $checkArray
    summary = [pscustomobject]@{
        pass = $passCount
        warn = $warnCount
        fail = $failCount
        ready = $failCount -eq 0
    }
}
$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $reportFullPath -Encoding UTF8

Write-Host ''
Write-Host 'Agent for Excel clean-machine acceptance'
Write-Host ('Phase: ' + $Phase)
$checkArray | Select-Object status, id, message | Format-Table -AutoSize
Write-Host ("Summary: PASS={0}, WARN={1}, FAIL={2}" -f $passCount, $warnCount, $failCount)
Write-Host ('Machine report: ' + $reportFullPath)

if ($failCount -gt 0) { exit 2 }
exit 0
