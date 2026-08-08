$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
$dllPath = Join-Path $root 'bin\Debug\AgentForExcel.dll'
$testPng = Join-Path $scriptDir 'draw-from-image-test.png'
$outPng = Join-Path $scriptDir 'draw-from-image-e2e.png'

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public struct E2ERECT { public int Left; public int Top; public int Right; public int Bottom; }
public class E2EWin32 {
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out E2ERECT rect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@

$dir = Split-Path -Parent $dllPath
$onResolve = [System.ResolveEventHandler]{
    param($s, $e)
    $name = (New-Object System.Reflection.AssemblyName $e.Name).Name
    $candidate = Join-Path $dir ($name + '.dll')
    if (Test-Path $candidate) { return [System.Reflection.Assembly]::LoadFrom($candidate) }
    return $null
}
[System.AppDomain]::CurrentDomain.add_AssemblyResolve($onResolve)
$asm = [System.Reflection.Assembly]::LoadFrom($dllPath)
$ctxType = $asm.GetType('AgentForExcel.AppContext')
$factoryType = $asm.GetType('AgentForExcel.Operations.Cell.DrawFromImageOp+Factory')
$factory = [System.Activator]::CreateInstance($factoryType)
$parse = $factoryType.GetMethod('Parse')

function Invoke-Parse([string]$json) {
    return $parse.Invoke($factory, [object[]]@([string]$json))
}
$flags = [System.Reflection.BindingFlags]'Instance,NonPublic'

function New-LightContext([object]$excelApp) {
    $ctx = [System.Runtime.Serialization.FormatterServices]::GetUninitializedObject($ctxType)
    $ctxType.GetField('<Excel>k__BackingField', $flags).SetValue($ctx, $excelApp)
    return $ctx
}

$excel = $null; $wb = $null; $ctx = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'DrawFromImageE2E'
    $ws.Cells.Clear()
    $ctx = New-LightContext $excel

    # 1) 无调色板：8x8 最近邻
    $jsonA = @{ address='B2'; image_path=$testPng; grid_width=8; grid_height=8; pixel_width=16; pixel_height=16; hide_gridlines=$true } | ConvertTo-Json -Compress
    $opA = Invoke-Parse $jsonA
    $resA = $opA.Execute($ctx)
    "Execute A: $resA"

    # 2) 调色板量化：8x8，画到 K2
    $jsonB = @{ address='K2'; image_path=$testPng; grid_width=8; grid_height=8; pixel_width=16; pixel_height=16; palette=@('#E52521','#00A650','#1A3FA8','#FFD800','#000000') } | ConvertTo-Json -Compress
    $opB = Invoke-Parse $jsonB
    $resB = $opB.Execute($ctx)
    "Execute B: $resB"

    # 3) 逐格校验 B2:I9（无调色板）
    $expectA = @(
      @($null,'#E52521','#E52521','#E52521','#00A650','#00A650','#00A650','#00A650'),
      @('#E52521','#E52521','#E52521','#E52521','#00A650','#00A650','#00A650','#00A650'),
      @('#E52521','#E52521','#E52521','#E52521','#00A650','#00A650','#00A650','#00A650'),
      @('#E52521','#E52521','#E52521','#E52521','#00A650','#00A650','#00A650','#00A650'),
      @('#1A3FA8','#1A3FA8','#1A3FA8','#1A3FA8','#FFD800','#FFD800','#FFD800','#FFD800'),
      @('#1A3FA8','#1A3FA8','#1A3FA8','#1A3FA8','#FFD800','#FFD800','#FFD800','#FFD800'),
      @('#1A3FA8','#1A3FA8','#1A3FA8','#1A3FA8','#FFD800','#FFD800','#FFD800','#FFD800'),
      @('#1A3FA8','#1A3FA8','#1A3FA8','#1A3FA8','#FFD800','#FFD800','#FFD800','#FFD800')
    )
    function HexToOle([string]$hex) {
        $ole = [Convert]::ToInt32($hex.Substring(1), 16)
        return ((($ole -band 0xFF) -shl 16) -bor ($ole -band 0xFF00) -bor (($ole -shr 16) -band 0xFF))
    }
    $okA = $true
    for ($r = 0; $r -lt 8; $r++) {
        for ($c = 0; $c -lt 8; $c++) {
            $cell = $ws.Range([char](66 + $c) + ($r + 2))  # B2..I9
            $exp = $expectA[$r][$c]
            if ($null -eq $exp) {
                $fill = $cell.Interior.ColorIndex
                if ($fill -ne -4142) { $okA = $false; "  mismatch[$r,$c] expected empty got ColorIndex=$fill" }
            } else {
                $ole = [int]$cell.Interior.Color
                $expOle = HexToOle $exp
                if ($ole -ne $expOle) { $okA = $false; "  mismatch[$r,$c] expected $exp (0x$($expOle.ToString('X6'))) got 0x$($ole.ToString('X6'))" }
            }
        }
    }
    "Case A cells (B2:I9): $(if($okA){'PASS'}else{'FAIL'})"

    # 4) 调色板区域 K2:R9 与无调色板逐格一致（测试图颜色本身在调色板内）
    $okB = $true
    for ($r = 0; $r -lt 8; $r++) {
        for ($c = 0; $c -lt 8; $c++) {
            $cell = $ws.Range([char](75 + $c) + ($r + 2))  # K2..R9
            $exp = $expectA[$r][$c]
            if ($null -eq $exp) {
                $fill = $cell.Interior.ColorIndex
                if ($fill -ne -4142) { $okB = $false; "  palette mismatch[$r,$c] expected empty got ColorIndex=$fill" }
            } else {
                $ole = [int]$cell.Interior.Color
                $expOle = HexToOle $exp
                if ($ole -ne $expOle) { $okB = $false; "  palette mismatch[$r,$c] expected $exp got 0x$($ole.ToString('X6'))" }
            }
        }
    }
    "Case B cells (K2:R9): $(if($okB){'PASS'}else{'FAIL'})"

    # 5) 行高/列宽/网格线
    $heights = @(); for ($r = 2; $r -le 9; $r++) { $heights += [double]$ws.Rows.Item($r).Height }
    $hMin = ($heights | Measure-Object -Minimum).Minimum; $hMax = ($heights | Measure-Object -Maximum).Maximum
    "Row heights 2..9: min={0:N2} max={1:N2} (target 16pt)" -f $hMin, $hMax
    $probe = $ws.Range('B2')
    "Probe cell Width={0:N2}pt Height={1:N2}pt" -f [double]$probe.Width, [double]$probe.Height
    try { "DisplayGridlines={0}" -f $excel.ActiveWindow.DisplayGridlines } catch { "DisplayGridlines=unknown" }

    # 6) 截图
    $sel = $ws.Range('B2:R9')
    $ws.Activate() | Out-Null
    $sel.Select() | Out-Null
    try { $excel.ActiveWindow.Zoom = $true } catch {}
    $excel.ActiveWindow.ScrollRow = 1
    $excel.ActiveWindow.ScrollColumn = 1
    Start-Sleep -Milliseconds 800
    $h = [IntPtr]$excel.Hwnd
    [E2EWin32]::SetForegroundWindow($h) | Out-Null
    Start-Sleep -Milliseconds 500
    $rect = New-Object E2ERECT
    [E2EWin32]::GetWindowRect($h, [ref]$rect) | Out-Null
    $bw = $rect.Right - $rect.Left; $bh = $rect.Bottom - $rect.Top
    $bmp = New-Object System.Drawing.Bitmap $bw, $bh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($rect.Left, $rect.Top, 0, 0, (New-Object System.Drawing.Size $bw, $bh))
    $bmp.Save($outPng, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    "Screenshot saved: $outPng"

    $allOk = $okA -and $okB
    if ($allOk) { "E2E RESULT: PASS" } else { "E2E RESULT: FAIL"; exit 1 }
}
catch {
    "ERR: " + $_.Exception.Message
    if ($_.Exception.InnerException) { "INNER: " + $_.Exception.InnerException.Message }
    exit 1
}
finally {
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    if ($excel -ne $null) { try { $excel.Quit() } catch {} }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
}