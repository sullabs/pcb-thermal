# Builds the PCB Thermal add-in with the VB compiler that ships with Windows (.NET Framework 4).
# Output: addin\bin\ (PCBThermal.dll, interop DLLs, icons\). Then run install.ps1.
param(
    [string]$SwDir = ""
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$bin = Join-Path $here "bin"
$vbc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\vbc.exe"
if (-not (Test-Path $vbc)) { throw "VB compiler not found: $vbc" }

# SolidWorks API DLLs (api\redist of the SolidWorks install)
if ($SwDir -eq "") {
    $key = Get-ItemProperty "HKLM:\SOFTWARE\SolidWorks\SOLIDWORKS 2023\Setup" -ErrorAction SilentlyContinue
    if ($key -and $key.SolidWorks_Folder) { $SwDir = $key.SolidWorks_Folder }
}
foreach ($cand in @($SwDir, "E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS", "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS")) {
    if ($cand -and (Test-Path (Join-Path $cand "api\redist\SolidWorks.Interop.sldworks.dll"))) { $SwDir = $cand; break }
}
$redist = Join-Path $SwDir "api\redist"
if (-not (Test-Path (Join-Path $redist "SolidWorks.Interop.sldworks.dll"))) { throw "SolidWorks api\redist not found. Pass -SwDir." }

New-Item -ItemType Directory -Force $bin | Out-Null
New-Item -ItemType Directory -Force (Join-Path $bin "icons") | Out-Null
$interops = @("SolidWorks.Interop.sldworks.dll", "SolidWorks.Interop.swconst.dll", "SolidWorks.Interop.swpublished.dll")
foreach ($d in $interops) { Copy-Item (Join-Path $redist $d) $bin -Force }

# Compile
$sources = Get-ChildItem (Join-Path $here "src") -Filter *.vb | ForEach-Object { $_.FullName }
$refs = ($interops | ForEach-Object { "/reference:" + (Join-Path $bin $_) }) + @(
    "/reference:System.dll", "/reference:System.Drawing.dll", "/reference:System.Windows.Forms.dll")
$outDll = Join-Path $bin "PCBThermal.dll"
$rsp = @("/nologo", "/target:library", "/platform:anycpu", "/optimize+", "/debug:pdbonly",
    "/imports:Microsoft.VisualBasic,System", "/rootnamespace:PCBThermal", "`"/out:$outDll`"") +
    ($refs | ForEach-Object { "`"$_`"" }) + ($sources | ForEach-Object { "`"$_`"" })
$obj = Join-Path $here "obj"
New-Item -ItemType Directory -Force $obj | Out-Null
$rspFile = Join-Path $obj "addin.rsp"
Set-Content -Path $rspFile -Value $rsp -Encoding ASCII
& $vbc "@$rspFile"
if ($LASTEXITCODE -ne 0) { throw "Compile failed." }

# Icons: one strip per size for the ribbon (8 steps + Check Model + Check Workbook), and a main icon.
Add-Type -AssemblyName System.Drawing
$iconDir = Join-Path (Split-Path -Parent $here) "src\icons"
$stepIcons = Get-ChildItem $iconDir -Filter "*_32.bmp" | Sort-Object Name | ForEach-Object { $_.FullName }

function New-GlyphIcon([string]$text, [System.Drawing.Color]$fill) {
    $bmp = New-Object System.Drawing.Bitmap 32, 32
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = "AntiAlias"; $g.TextRenderingHint = "AntiAliasGridFit"
    $g.Clear([System.Drawing.Color]::Transparent)
    $brush = New-Object System.Drawing.SolidBrush $fill
    $g.FillEllipse($brush, 1, 1, 30, 30)
    $font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $sf = New-Object System.Drawing.StringFormat; $sf.Alignment = "Center"; $sf.LineAlignment = "Center"
    $g.DrawString($text, $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF 0, 0, 32, 32), $sf)
    $g.Dispose()
    return $bmp
}

$images = @()
foreach ($f in $stepIcons) { $images += [System.Drawing.Image]::FromFile($f) }
$images += New-GlyphIcon "M" ([System.Drawing.Color]::FromArgb(0, 120, 80))
$images += New-GlyphIcon "W" ([System.Drawing.Color]::FromArgb(30, 110, 60))
$main = New-GlyphIcon "PT" ([System.Drawing.Color]::FromArgb(200, 70, 30))

foreach ($size in 20, 32, 40, 64, 96, 128) {
    $strip = New-Object System.Drawing.Bitmap ($size * $images.Count), $size
    $g = [System.Drawing.Graphics]::FromImage($strip)
    $g.InterpolationMode = "HighQualityBicubic"
    for ($i = 0; $i -lt $images.Count; $i++) { $g.DrawImage($images[$i], $i * $size, 0, $size, $size) }
    $g.Dispose()
    $strip.Save((Join-Path $bin "icons\strip_$size.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $strip.Dispose()
    $one = New-Object System.Drawing.Bitmap $size, $size
    $g = [System.Drawing.Graphics]::FromImage($one)
    $g.InterpolationMode = "HighQualityBicubic"
    $g.DrawImage($main, 0, 0, $size, $size)
    $g.Dispose()
    $one.Save((Join-Path $bin "icons\main_$size.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $one.Dispose()
}
foreach ($im in $images) { $im.Dispose() }
$main.Dispose()

# ---------------- Setup exe: one file with everything inside ----------------
$version = [System.Reflection.AssemblyName]::GetAssemblyName($outDll).Version
$ver3 = "$($version.Major).$($version.Minor).$($version.Build)"
$dist = Join-Path $here "dist"
New-Item -ItemType Directory -Force $dist | Out-Null

# Payload: add-in files, the template workbook, and the macro project if present (until all steps are ported).
$payload = @{}
Get-ChildItem $bin -Recurse -File | Where-Object { $_.Extension -ne ".pdb" } | ForEach-Object {
    $payload[$_.FullName.Substring($bin.Length + 1).Replace("\", "/")] = $_.FullName }
$template = Join-Path (Split-Path -Parent $here) "workbook\PCB_thermal_TEMPLATE.xlsx"
if (Test-Path $template) { $payload["template/PCB_thermal_TEMPLATE.xlsx"] = $template }
$swp = Join-Path $here "macros\PCBThermal.swp"
if (Test-Path $swp) { $payload["PCBThermal.swp"] = $swp } else { Write-Host "Note: addin\macros\PCBThermal.swp not found, setup will not contain the macros." }

$verFile = Join-Path $obj "SetupVersion.vb"
Set-Content -Path $verFile -Encoding ASCII -Value @(
    "<Assembly: System.Reflection.AssemblyVersion(""$version"")>",
    "<Assembly: System.Reflection.AssemblyFileVersion(""$version"")>",
    "<Assembly: System.Reflection.AssemblyProduct(""PCB Thermal"")>",
    "<Assembly: System.Reflection.AssemblyTitle(""PCB Thermal setup"")>")
$setupExe = Join-Path $dist "PCBThermal_Setup_$ver3.exe"
$rsp2 = @("/nologo", "/target:winexe", "/platform:anycpu", "/optimize+", "/main:Setup",
    "/imports:Microsoft.VisualBasic,System", "/reference:System.dll", "/reference:System.Drawing.dll",
    "/reference:System.Windows.Forms.dll", "`"/win32manifest:$(Join-Path $here 'setup\setup.manifest')`"",
    "`"/out:$setupExe`"", "`"$(Join-Path $here 'setup\Setup.vb')`"", "`"$verFile`"") +
    ($payload.Keys | Sort-Object | ForEach-Object { "`"/resource:$($payload[$_]),payload/$_`"" })
$rspFile2 = Join-Path $obj "setup.rsp"
Set-Content -Path $rspFile2 -Value $rsp2 -Encoding ASCII
& $vbc "@$rspFile2"
if ($LASTEXITCODE -ne 0) { throw "Setup compile failed." }

Write-Host "Built $outDll and $setupExe ($($payload.Count) files inside). Share the setup exe."
