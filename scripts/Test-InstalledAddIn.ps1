[CmdletBinding()]
param(
    [string]$AddInProgId = 'AgentForExcel',

    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$ExpectedDeploymentVersion,

    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$ExpectedAssemblyVersion,

    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Add-Check {
    param(
        [System.Collections.Generic.List[object]]$Checks,
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

function Release-ComObject {
    param([object]$Value)
    try {
        if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
        }
    }
    catch { }
}

function Get-InstalledManifestInfo {
    param([string]$ProgId)

    $paths = @(
        "HKCU:\Software\Microsoft\Office\Excel\Addins\$ProgId",
        "HKLM:\Software\Microsoft\Office\Excel\Addins\$ProgId",
        "HKLM:\Software\WOW6432Node\Microsoft\Office\Excel\Addins\$ProgId"
    )
    foreach ($path in $paths) {
        try {
            $manifestValue = [string](Get-ItemProperty -LiteralPath $path -ErrorAction Stop).Manifest
            if ([string]::IsNullOrWhiteSpace($manifestValue)) { continue }
            $manifestLocation = $manifestValue.Split('|')[0]
            $localPath = $manifestLocation
            $uri = $null
            if ([Uri]::TryCreate($manifestLocation, [UriKind]::Absolute, [ref]$uri) -and $uri.IsFile) {
                $localPath = $uri.LocalPath
            }
            $version = $null
            if (Test-Path -LiteralPath $localPath) {
                [xml]$manifest = Get-Content -LiteralPath $localPath -Raw
                $identity = $manifest.SelectSingleNode(
                    "/*[local-name()='assembly']/*[local-name()='assemblyIdentity']")
                if ($null -ne $identity) { $version = [string]$identity.version }
            }
            return [pscustomobject]@{
                registryPath = $path
                manifestValue = $manifestValue
                localPath = $localPath
                version = $version
            }
        }
        catch { }
    }
    return $null
}

$checks = New-Object 'System.Collections.Generic.List[object]'
$excel = $null
$workbook = $null
$sheet = $null
$addIn = $null
$service = $null
$startedAt = Get-Date
$existingExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)

if ($existingExcel.Count -gt 0) {
    throw 'Close all Excel windows before running post-install acceptance. The script never closes user workbooks.'
}

try {
    $manifestInfo = Get-InstalledManifestInfo $AddInProgId
    Add-Check $checks 'installation.registry' `
        $(if ($null -ne $manifestInfo) { 'pass' } else { 'fail' }) `
        $(if ($null -ne $manifestInfo) {
            'Excel add-in registration was found.'
        } else {
            'Excel add-in registration was not found.'
        }) $manifestInfo

    if (-not [string]::IsNullOrWhiteSpace($ExpectedDeploymentVersion)) {
        $deploymentVersionMatches = $null -ne $manifestInfo -and
            [string]::Equals(
                $manifestInfo.version, $ExpectedDeploymentVersion, [StringComparison]::Ordinal)
        Add-Check $checks 'installation.deployment_version' `
            $(if ($deploymentVersionMatches) { 'pass' } else { 'fail' }) `
            $(if ($deploymentVersionMatches) {
                'Installed deployment manifest version matches.'
            } else {
                'Installed deployment manifest version could not be verified or does not match.'
            }) ([pscustomobject]@{
                expected = $ExpectedDeploymentVersion
                actual = if ($null -eq $manifestInfo) { $null } else { $manifestInfo.version }
            })
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.AutomationSecurity = 3
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $excelDirectory = [string]$excel.Path
    $excelPath = if ([string]::IsNullOrWhiteSpace($excelDirectory)) {
        $null
    } else {
        Join-Path $excelDirectory 'EXCEL.EXE'
    }
    $programFilesX86 = ${env:ProgramFiles(x86)}
    $excelBitness = if (-not [string]::IsNullOrWhiteSpace($excelPath) -and
        -not [string]::IsNullOrWhiteSpace($programFilesX86) -and
        $excelPath.StartsWith($programFilesX86, [StringComparison]::OrdinalIgnoreCase)) {
        'x86'
    } else {
        'x64'
    }
    Add-Check $checks 'excel.launch' 'pass' 'An isolated Excel process was started.' `
        ([pscustomobject]@{
            version = [string]$excel.Version
            build = [string]$excel.Build
            bitness = $excelBitness
            path = $excelPath
        })

    try {
        $addIn = $excel.COMAddIns.Item($AddInProgId)
    }
    catch {
        Add-Check $checks 'addin.discover' 'fail' 'Excel could not find the installed COM add-in.' $AddInProgId
        throw
    }
    Add-Check $checks 'addin.discover' 'pass' 'Excel discovered the installed COM add-in.' `
        ([pscustomobject]@{
            progId = [string]$addIn.ProgId
            description = [string]$addIn.Description
        })

    $addIn.Connect = $true
    $deadline = (Get-Date).AddSeconds(20)
    do {
        Start-Sleep -Milliseconds 250
        $connected = [bool]$addIn.Connect
        if ($connected) {
            try { $service = $addIn.Object } catch { }
        }
    } while ((-not $connected -or $null -eq $service) -and (Get-Date) -lt $deadline)

    Add-Check $checks 'addin.connect' `
        $(if ($connected -and $null -ne $service) { 'pass' } else { 'fail' }) `
        $(if ($connected -and $null -ne $service) {
            'The VSTO add-in connected and exposed its automation service.'
        } else {
            'The VSTO add-in did not connect or expose its automation service within 20 seconds.'
        }) $null
    if (-not $connected -or $null -eq $service) {
        throw 'Agent for Excel automation service is unavailable.'
    }

    $assemblyVersion = [string]$service.Version()
    $assemblyVersionMatches = [string]::IsNullOrWhiteSpace($ExpectedAssemblyVersion) -or
        [string]::Equals($assemblyVersion, $ExpectedAssemblyVersion, [StringComparison]::Ordinal)
    Add-Check $checks 'addin.assembly_version' `
        $(if ($assemblyVersionMatches) { 'pass' } else { 'fail' }) `
        $(if ($assemblyVersionMatches) {
            'Loaded add-in assembly version is acceptable.'
        } else {
            'Loaded add-in assembly version does not match.'
        }) ([pscustomobject]@{ expected = $ExpectedAssemblyVersion; actual = $assemblyVersion })

    $capabilityJson = [string]$service.CapabilityCheck()
    $capabilities = $capabilityJson | ConvertFrom-Json
    $capabilitiesValid = [bool]$capabilities.all_registered -and
        [int]$capabilities.required_count -eq 24 -and
        [int]$capabilities.registered_count -eq 24
    Add-Check $checks 'addin.capabilities' `
        $(if ($capabilitiesValid) { 'pass' } else { 'fail' }) `
        $(if ($capabilitiesValid) {
            'All 24 required Agent tools are registered.'
        } else {
            'Required Agent tools are missing.'
        }) $capabilities

    $workbook = $excel.Workbooks.Add()
    $workbook.Activate()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'AcceptanceData'
    $sheet.Range('A1').Value2 = 'Product'
    $sheet.Range('B1').Value2 = 'Sales'
    $sheet.Range('A2').Value2 = 'A'
    $sheet.Range('B2').Value2 = 100
    $sheet.Range('A3').Value2 = 'B'
    $sheet.Range('B3').Value2 = 200

    $readResult = [string]$service.RunReadOnlyTool(
        'cell_read_range', '{"sheet":"AcceptanceData","address":"A1:B3"}')
    $readValid = $readResult.StartsWith('__AGENT_TABLE_PREVIEW__') -and
        $readResult.Contains('Product') -and $readResult.Contains('Sales')
    Add-Check $checks 'addin.read_range' `
        $(if ($readValid) { 'pass' } else { 'fail' }) `
        $(if ($readValid) {
            'The installed add-in read a real workbook range through its dispatcher.'
        } else {
            'The installed add-in could not read the acceptance range.'
        }) ([pscustomobject]@{
            result = if ($readResult.Length -le 1000) {
                $readResult
            } else {
                $readResult.Substring(0, 1000)
            }
        })

    $health = [string]$service.HealthCheck()
    $healthValid = -not [string]::IsNullOrWhiteSpace($health) -and
        $health.IndexOf('Excel', [StringComparison]::OrdinalIgnoreCase) -ge 0
    Add-Check $checks 'addin.health_check' `
        $(if ($healthValid) { 'pass' } else { 'fail' }) `
        $(if ($healthValid) {
            'The installed add-in environment self-check returned a result.'
        } else {
            'The installed add-in environment self-check did not return a valid result.'
        }) $health
}
catch {
    Add-Check $checks 'run.exception' 'fail' 'Installed add-in acceptance stopped with an exception.' `
        ([pscustomobject]@{
            type = $_.Exception.GetType().FullName
            message = $_.Exception.Message
        })
}
finally {
    try { if ($null -ne $workbook) { $workbook.Close($false) } } catch { }
    try { if ($null -ne $excel) { $excel.Quit() } } catch { }
    Release-ComObject $sheet
    Release-ComObject $workbook
    Release-ComObject $service
    Release-ComObject $addIn
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$checkArray = @($checks | ForEach-Object { $_ })
$failCount = @($checkArray | Where-Object { $_.status -eq 'fail' }).Count
$warnCount = @($checkArray | Where-Object { $_.status -eq 'warn' }).Count
$passCount = @($checkArray | Where-Object { $_.status -eq 'pass' }).Count
$report = [pscustomobject]@{
    schemaVersion = 1
    generatedAt = (Get-Date).ToString('o')
    durationSeconds = [Math]::Round(((Get-Date) - $startedAt).TotalSeconds, 2)
    computerName = $env:COMPUTERNAME
    addInProgId = $AddInProgId
    checks = $checkArray
    summary = [pscustomobject]@{
        pass = $passCount
        warn = $warnCount
        fail = $failCount
        installedAddInReady = $failCount -eq 0
    }
}

if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path (Get-Location).Path (
        'InstalledAcceptance-' + $env:COMPUTERNAME + '-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.json')
}
elseif (-not [IO.Path]::IsPathRooted($ReportPath)) {
    $ReportPath = Join-Path (Get-Location).Path $ReportPath
}
$reportDirectory = Split-Path -Parent $ReportPath
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -ItemType Directory -Path $reportDirectory | Out-Null
}
$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ReportPath -Encoding UTF8

Write-Host ''
Write-Host 'Agent for Excel installed add-in acceptance'
$checkArray | Select-Object status, id, message | Format-Table -AutoSize
Write-Host ("Summary: PASS={0}, WARN={1}, FAIL={2}" -f $passCount, $warnCount, $failCount)
Write-Host ('Installed add-in ready: ' + $report.summary.installedAddInReady)
Write-Host ('JSON report: ' + (Resolve-Path -LiteralPath $ReportPath).Path)

if ($failCount -gt 0) { exit 2 }
exit 0
