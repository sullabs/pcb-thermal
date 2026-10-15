# Builds the NTIS thermal simulation add-in with the VB compiler that ships with Windows (.NET Framework 4).
# Output: addin\bin\ (NTIS_thermal_simulation.dll, interop DLLs, icons\). Then run install.ps1.
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
$interops = @("SolidWorks.Interop.sldworks.dll", "SolidWorks.Interop.swconst.dll", "SolidWorks.Interop.swpublished.dll",
    "SolidWorks.Interop.cosworks.dll")
foreach ($d in $interops) { Copy-Item (Join-Path $redist $d) $bin -Force }

# Compile
$sources = Get-ChildItem (Join-Path $here "src") -Filter *.vb -Recurse | ForEach-Object { $_.FullName }
$refs = ($interops | ForEach-Object { "/reference:" + (Join-Path $bin $_) }) + @(
    "/reference:System.dll", "/reference:System.Drawing.dll", "/reference:System.Windows.Forms.dll")
$outDll = Join-Path $bin "NTIS_thermal_simulation.dll"
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
# Made from the PNGs in icons\png (rendered from icons\svg by icons\render.ps1), each size from its own PNG.
Add-Type -AssemblyName System.Drawing
$pngDir = Join-Path $here "icons\png"
$stripNames = @("01-prepare-workbook", "02-import-step", "03-build-board", "04-build-spreader", "05-sync-cad-to-excel",
    "06-run-thermal-study", "07-run-flow-sim", "08-read-results", "09-check-model", "10-check-workbook")
$mainName = "11-main"
foreach ($size in 20, 32, 40, 64, 96, 128) {
    $strip = New-Object System.Drawing.Bitmap ($size * $stripNames.Count), $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($strip)
    $g.Clear([System.Drawing.Color]::Transparent)
    for ($i = 0; $i -lt $stripNames.Count; $i++) {
        $f = Join-Path $pngDir "$($stripNames[$i])-$size.png"
        if (-not (Test-Path $f)) { throw "Icon missing: $f. Run icons\render.ps1." }
        $im = [System.Drawing.Image]::FromFile($f)
        if ($im.Width -ne $size -or $im.Height -ne $size) { throw "Icon $f is not $size px." }
        $g.DrawImage($im, (New-Object System.Drawing.Rectangle ($i * $size), 0, $size, $size))
        $im.Dispose()
    }
    $g.Dispose()
    $strip.Save((Join-Path $bin "icons\strip_$size.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $strip.Dispose()
    $f = Join-Path $pngDir "$mainName-$size.png"
    if (-not (Test-Path $f)) { throw "Icon missing: $f. Run icons\render.ps1." }
    Copy-Item $f (Join-Path $bin "icons\main_$size.png") -Force
}

# ---------------- Setup exe: one file with everything inside ----------------
$version = [System.Reflection.AssemblyName]::GetAssemblyName($outDll).Version
$ver3 = "$($version.Major).$($version.Minor).$($version.Build)"
$dist = Join-Path $here "dist"
New-Item -ItemType Directory -Force $dist | Out-Null

# Payload: add-in files, the template workbook, and the macro project if present (until all steps are ported).
$payload = @{}
Get-ChildItem $bin -File | Where-Object { $_.BaseName -ne [IO.Path]::GetFileNameWithoutExtension($outDll) -and $_.Name -notlike "SolidWorks.Interop.*" } | Remove-Item
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
    "<Assembly: System.Reflection.AssemblyProduct(""NTIS thermal simulation"")>",
    "<Assembly: System.Reflection.AssemblyTitle(""NTIS thermal simulation setup"")>")
$setupExe = Join-Path $dist "NTIS_thermal_simulation_Setup_$ver3.exe"
$rsp2 = @("/nologo", "/target:winexe", "/platform:anycpu", "/optimize+", "/main:Setup",
    "/imports:Microsoft.VisualBasic,System", "/reference:System.dll", "/reference:System.Drawing.dll",
    "/reference:System.Windows.Forms.dll", "`"/win32manifest:$(Join-Path $here 'setup\setup.manifest')`"",
    "`"/out:$setupExe`"", "`"$(Join-Path $here 'setup\Setup.vb')`"",
    "`"$(Join-Path $here 'setup\Loader.vb')`"", "`"$verFile`"") +
    ($payload.Keys | Sort-Object | ForEach-Object { "`"/resource:$($payload[$_]),payload/$_`"" })
$rspFile2 = Join-Path $obj "setup.rsp"
Set-Content -Path $rspFile2 -Value $rsp2 -Encoding ASCII
& $vbc "@$rspFile2"
if ($LASTEXITCODE -ne 0) { throw "Setup compile failed." }

Write-Host "Built $outDll and $setupExe ($($payload.Count) files inside). Share the setup exe."
