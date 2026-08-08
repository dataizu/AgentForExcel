$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
$dllPath = Join-Path $root 'bin\Debug\AgentForExcel.dll'
$testPng = Join-Path $scriptDir 'draw-from-image-test.png'
$badExtFile = Join-Path $scriptDir 'draw-from-image-bad.svg'

# --- 1) 生成 16x16 测试图：4 象限纯色 + 左上角 2x2 透明 ---
Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap 16, 16
for ($y = 0; $y -lt 16; $y++) {
    for ($x = 0; $x -lt 16; $x++) {
        if ($x -lt 2 -and $y -lt 2) { continue } # transparent
        elseif ($x -lt 8 -and $y -lt 8) { $c = [System.Drawing.ColorTranslator]::FromHtml('#E52521') }
        elseif ($x -ge 8 -and $y -lt 8) { $c = [System.Drawing.ColorTranslator]::FromHtml('#00A650') }
        elseif ($x -lt 8 -and $y -ge 8) { $c = [System.Drawing.ColorTranslator]::FromHtml('#1A3FA8') }
        else { $c = [System.Drawing.ColorTranslator]::FromHtml('#FFD800') }
        $bmp.SetPixel($x, $y, $c)
    }
}
$bmp.Save($testPng, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
"test image saved: $testPng"
[System.IO.File]::WriteAllText($badExtFile, "<svg xmlns='http://www.w3.org/2000/svg'/>", (New-Object System.Text.UTF8Encoding($false)))

# --- 2) 加载 AgentForExcel.dll（含依赖解析） ---
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
$factoryType = $asm.GetType('AgentForExcel.Operations.Cell.DrawFromImageOp+Factory')
if ($null -eq $factoryType) { throw 'Factory type not found' }
$factory = [System.Activator]::CreateInstance($factoryType)
$parse = $factoryType.GetMethod('Parse')

function Invoke-Parse([string]$json) {
    $argsArr = [object[]]::new(1)
    $argsArr[0] = [string]$json
    return $parse.Invoke($factory, $argsArr)
}
$flags = [System.Reflection.BindingFlags]'Instance,NonPublic'

function Get-PixelsFromOp([object]$op) {
    $inner = $op.GetType().GetField('_inner', $flags).GetValue($op)
    $px = $inner.GetType().GetField('_pixels', $flags).GetValue($inner)
    return ,$px
}

# --- 3) 用例 A：8x8 无调色板 ---
$jsonA = @{ address='B2'; image_path=$testPng; grid_width=8; grid_height=8; pixel_width=12; pixel_height=12 } | ConvertTo-Json -Compress
$opA = Invoke-Parse $jsonA
$pxA = Get-PixelsFromOp $opA
"Case A dims: $($pxA.GetLength(0)) x $($pxA.GetLength(1))"
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
$okA = $true
for ($r = 0; $r -lt 8; $r++) { for ($c = 0; $c -lt 8; $c++) {
    $exp = $expectA[$r][$c]; $act = $pxA[$r,$c]
    if (($null -eq $exp) -and ($null -ne $act)) { $okA = $false; "  mismatch[$r,$c] expected null got $act" }
    elseif (($null -ne $exp) -and ($exp -ne $act)) { $okA = $false; "  mismatch[$r,$c] expected $exp got $act" }
}}
"Case A (8x8 no palette): $(if($okA){'PASS'}else{'FAIL'})"

# --- 4) 用例 B：8x8 带调色板量化 ---
$jsonB = @{ address='B2'; image_path=$testPng; grid_width=8; grid_height=8; palette=@('#E52521','#00A650','#1A3FA8','#FFD800','#000000') } | ConvertTo-Json -Compress
$opB = Invoke-Parse $jsonB
$pxB = Get-PixelsFromOp $opB
$okB = $true
for ($r = 0; $r -lt 8; $r++) { for ($c = 0; $c -lt 8; $c++) {
    $act = $pxB[$r,$c]
    if ($null -ne $act -and $expectA[$r][$c] -ne $act) { $okB = $false; "  quantize mismatch[$r,$c] expected $($expectA[$r][$c]) got $act" }
}}
"Case B (8x8 palette): $(if($okB){'PASS'}else{'FAIL'})"

# --- 5) 用例 C：校验失败路径 ---
function Test-Reject([string]$name, [string]$json, [string]$expectMsg) {
    try { $null = Invoke-Parse $json; "Case $name : FAIL (no exception)" }
    catch { $msg = $_.Exception.InnerException.Message; if ($msg -like "*$expectMsg*") { "Case $name : PASS ($msg)" } else { "Case $name : FAIL unexpected ($msg)" } }
}
Test-Reject 'bad-ext'   (@{ address='B2'; image_path=$badExtFile } | ConvertTo-Json -Compress) '仅支持本地图片文件'
Test-Reject 'missing'   (@{ address='B2'; image_path='D:\no-such-file.png'; grid_width=8; grid_height=8 } | ConvertTo-Json -Compress) '找不到图片文件'
Test-Reject 'too-big'   (@{ address='B2'; image_path=$testPng; grid_width=1000; grid_height=1000 } | ConvertTo-Json -Compress) '超过单次上限'
Test-Reject 'empty-addr'(@{ address=''; image_path=$testPng; grid_width=8; grid_height=8 } | ConvertTo-Json -Compress) 'address 不能为空'
