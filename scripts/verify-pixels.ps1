$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outPng = Join-Path $scriptDir 'pixel-mario-check.png'

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
public class Win32 {
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'PixelCheck'
    $ws.Cells.Clear()

    # Exact replication of the fixed SetSquarePixels (target = whole matrix, probe = single cell)
    function Set-SquarePixels([object]$target, [object]$probe, [double]$widthPt, [double]$heightPt) {
        $target.EntireRow.RowHeight = $heightPt
        $columnWidth = ($widthPt - 5.0) / 7.0
        if ($columnWidth -lt 0.5) { $columnWidth = 0.5 }
        $target.EntireColumn.ColumnWidth = $columnWidth
        for ($i = 0; $i -lt 6; $i++) {
            $actualWidth = [double]$probe.Width
            $err = $actualWidth - $widthPt
            if ([Math]::Abs($err) -lt 0.25) { break }
            $columnWidth -= $err / 7.0
            if ($columnWidth -lt 0.5) { $columnWidth = 0.5 }
            if ($columnWidth -gt 255) { $columnWidth = 255 }
            $target.EntireColumn.ColumnWidth = $columnWidth
        }
        return $columnWidth
    }

    # 1) convergence check per pixel size
    foreach ($size in @(8, 12, 16, 20, 24)) {
        $t = $ws.Range('A1:E5')
        $p = $ws.Range('A1')
        $cw = Set-SquarePixels $t $p $size $size
        $w = [double]$p.Width
        $h = [double]$p.Height
        "{0}pt -> probe Width={1:N2}pt Height={2:N2}pt diff={3:N2} | finalColWidth={4:N3}ch" -f $size, $w, $h, ($w - $h), $cw
    }

    # 2) draw a 12x16 Mario at 12pt and verify EVERY row/col
    $mario = @(
        '..KKKKKKKK..',
        '.KRRRRRRRRK.',
        'KRRRRRRRRRRK',
        'KRRKKRRKKRRK',
        'KRRKKRRKKRRK',
        'KRRRRRRRRRRK',
        'KRRRSSSSRRRK',
        'KRRSSSSSSRRK',
        'KRRSSSSSSRRK',
        'KRRRSSSSRRRK',
        '.KRRRRRRRRK.',
        '.KBBBBBBBBK.',
        'KBBBBBBBBBBK',
        'KBBKBBBBKBBK',
        'KBBKBBBBKBBK',
        '.KKKKKKKKKK.'
    )
    $palette = @{ K = '#1A1A1A'; R = '#E52521'; S = '#FFB366'; B = '#1A3FA8'; Y = '#FFD800'; N = '#6B3306' }
    $rows = $mario.Count
    $cols = $mario[0].Length
    $target2 = $ws.Range('A1:' + [char](64 + $cols) + $rows)
    $probe2 = $ws.Range('A1')
    $null = Set-SquarePixels $target2 $probe2 12 12

    $colorCache = @{}
    for ($r = 0; $r -lt $rows; $r++) {
        for ($c = 0; $c -lt $cols; $c++) {
            $ch = $mario[$r][$c].ToString()
            if ($ch -eq '.') { continue }
            if (-not $palette.ContainsKey($ch)) { continue }
            $hex = $palette[$ch]
            if (-not $colorCache.ContainsKey($hex)) {
                $ole = [Convert]::ToInt32($hex.Substring(1), 16)
                $bgr = (($ole -band 0xFF) -shl 16) -bor ($ole -band 0xFF00) -bor (($ole -shr 16) -band 0xFF)
                $colorCache[$hex] = $bgr
            }
            $ws.Cells.Item($r + 1, $c + 1).Interior.Color = $colorCache[$hex]
        }
    }

    $minH = 999.0; $maxH = 0.0
    for ($r = 1; $r -le $rows; $r++) {
        $hh = [double]$ws.Rows.Item($r).Height
        if ($hh -lt $minH) { $minH = $hh }
        if ($hh -gt $maxH) { $maxH = $hh }
    }
    $minW = 999.0; $maxW = 0.0
    for ($c = 1; $c -le $cols; $c++) {
        $ww = [double]$ws.Columns.Item($c).Width
        if ($ww -lt $minW) { $minW = $ww }
        if ($ww -gt $maxW) { $maxW = $ww }
    }
    "Mario 12x16 @12pt: rowHeight min={0:N2} max={1:N2} | colWidth min={2:N2} max={3:N2} (target 12pt)" -f $minH, $maxH, $minW, $maxW

    # 3) hide gridlines, fit selection, screenshot window
    try { $excel.ActiveWindow.DisplayGridlines = $false } catch {}
    $sel = $ws.Range('A1:' + [char](64 + $cols) + $rows)
    $ws.Activate() | Out-Null
    $sel.Select() | Out-Null
    try { $excel.ActiveWindow.Zoom = $true } catch {}
    $excel.ActiveWindow.ScrollRow = 1
    $excel.ActiveWindow.ScrollColumn = 1
    Start-Sleep -Milliseconds 800

    $h = [IntPtr]$excel.Hwnd
    [Win32]::SetForegroundWindow($h) | Out-Null
    Start-Sleep -Milliseconds 500

    $rect = New-Object RECT
    [Win32]::GetWindowRect($h, [ref]$rect) | Out-Null
    $bw = $rect.Right - $rect.Left
    $bh = $rect.Bottom - $rect.Top
    $bmp = New-Object System.Drawing.Bitmap $bw, $bh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($rect.Left, $rect.Top, 0, 0, (New-Object System.Drawing.Size $bw, $bh))
    $bmp.Save($outPng, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $bmp.Dispose()
    "Screenshot saved: $outPng"
}
catch {
    "ERR: " + $_.Exception.Message
    if ($_.Exception.InnerException) { "INNER: " + $_.Exception.InnerException.Message }
}
finally {
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    if ($excel -ne $null) { try { $excel.Quit() } catch {} }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}