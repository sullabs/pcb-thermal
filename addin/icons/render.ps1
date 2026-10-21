# Renders the icon SVGs in svg\ to transparent PNGs in png\ (20, 32, 40, 64, 96, 128 px) with Edge headless,
# and the VBA toolbar bitmaps ..\..\src\icons\<n>_<Name>_16.bmp and _32.bmp (24 bit, white background).
# Run it after changing an SVG, then commit svg\, png\ and src\icons\. build.ps1 only reads png\.
param(
    [string]$Edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$svgDir = Join-Path $here "svg"
$pngDir = Join-Path $here "png"
$bmpDir = Join-Path (Split-Path -Parent (Split-Path -Parent $here)) "src\icons"
if (-not (Test-Path $Edge)) { throw "Edge not found: $Edge" }

$names = Get-ChildItem $svgDir -Filter *.svg | Sort-Object Name | ForEach-Object { $_.BaseName }
$pngSizes = 20, 32, 40, 64, 96, 128
$allSizes = @(16) + $pngSizes
$gap = 8
$cell = 128 + $gap

# One page with every icon at every size at known positions; one screenshot; then crop.
$tmp = Join-Path $env:TEMP ("ntis_icons_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force $tmp | Out-Null
$imgs = @()
$pos = @{}
$y = 0
foreach ($s in $allSizes) {
    for ($i = 0; $i -lt $names.Count; $i++) {
        $x = $i * $cell
        $uri = ([Uri](Join-Path $svgDir ($names[$i] + ".svg"))).AbsoluteUri
        $imgs += "<img src=""$uri"" style=""position:absolute;left:${x}px;top:${y}px;width:${s}px;height:${s}px"">"
        $pos["$($names[$i])|$s"] = @($x, $y)
    }
    $y += $s + $gap
}
$html = "<!doctype html><html><head><meta charset=""utf-8""><style>html,body{margin:0;background:transparent}</style></head><body>" +
    ($imgs -join "") + "</body></html>"
$page = Join-Path $tmp "sheet.html"
Set-Content -Path $page -Value $html -Encoding UTF8
$shot = Join-Path $tmp "sheet.png"
$w = $names.Count * $cell
$edgeArgs = @("--headless", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
    "--default-background-color=00000000", "--virtual-time-budget=3000", "--user-data-dir=`"$tmp\profile`"",
    "--window-size=$w,$y", "--screenshot=`"$shot`"", "`"$(([Uri]$page).AbsoluteUri)`"")
$p = Start-Process -FilePath $Edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
$p.WaitForExit()
for ($t = 0; $t -lt 50 -and -not (Test-Path $shot); $t++) { Start-Sleep -Milliseconds 200 }
if (-not (Test-Path $shot)) { throw "Edge did not write the screenshot." }

Add-Type -AssemblyName System.Drawing
$sheet = [System.Drawing.Bitmap]::FromFile($shot)
if ($sheet.GetPixel(0, $y - 1).A -ne 0) { throw "Screenshot background is not transparent." }
function Get-Crop([string]$n, [int]$s) {
    $xy = $pos["$n|$s"]
    return $sheet.Clone((New-Object System.Drawing.Rectangle $xy[0], $xy[1], $s, $s), [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
}

New-Item -ItemType Directory -Force $pngDir | Out-Null
Get-ChildItem $pngDir -Filter *.png | Remove-Item
foreach ($n in $names) {
    foreach ($s in $pngSizes) {
        $c = Get-Crop $n $s
        $c.Save((Join-Path $pngDir "$n-$s.png"), [System.Drawing.Imaging.ImageFormat]::Png)
        $c.Dispose()
    }
}

# VBA toolbar bitmaps: the 8 steps, same file names as before (n_Name_16.bmp, n_Name_32.bmp).
$bmpNames = @("1_PrepareWorkbook", "2_ImportStep", "3_BuildBoard", "4_BuildSpreader", "5_SyncCADtoExcel",
    "6_RunThermalStudy", "7_RunFlowSim", "8_ReadResults")
for ($i = 0; $i -lt $bmpNames.Count; $i++) {
    foreach ($s in 16, 32) {
        $c = Get-Crop $names[$i] $s
        $bmp = New-Object System.Drawing.Bitmap $s, $s, ([System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $g.DrawImage($c, 0, 0, $s, $s)
        $g.Dispose()
        $bmp.Save((Join-Path $bmpDir "$($bmpNames[$i])_$s.bmp"), [System.Drawing.Imaging.ImageFormat]::Bmp)
        $bmp.Dispose(); $c.Dispose()
    }
}
$sheet.Dispose()
Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
Write-Host "Rendered $($names.Count) icons to $pngDir and 8 x 2 bitmaps to $bmpDir."
