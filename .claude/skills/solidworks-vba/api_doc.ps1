# Prints the SolidWorks API help page (syntax, parameters, remarks) for one member, from the installed help.
# First run unpacks the help to %LOCALAPPDATA%\swapi_help (about 180 MB, one time, needs 7-Zip).
# Usage:  powershell -NoProfile -File api_doc.ps1 -Type IFace2 -Member GetTessTriangles
#         powershell -NoProfile -File api_doc.ps1 -Type IBody2            (member list page)
param([Parameter(Mandatory = $true)][string]$Type, [string]$Member = "", [string]$Ns = "sldworks")
$root = Join-Path $env:LOCALAPPDATA "swapi_help"
$all = Join-Path $root "all"
if (-not (Test-Path $all)) {
    $z = "C:\Program Files\7-Zip\7z.exe"
    $cab = "E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\HelpViewer\sldworksapi\apihelpviewer.cab"
    if (-not (Test-Path $z)) { Write-Output "7-Zip not found at $z"; exit 1 }
    New-Item -ItemType Directory -Force $root | Out-Null
    & $z x $cab "-o$root" -y | Out-Null
    & $z x (Join-Path $root "APIHelpViewer.mshc") "-o$all" -y | Out-Null
}
$pre = "SolidWorks.Interop.$Ns~SolidWorks.Interop.$Ns.$Type"
$file = if ($Member -ne "") { "$pre~$Member.html" } else { "$pre.html" }
$full = Join-Path $all $file
if (-not (Test-Path $full)) { Write-Output "No page: $file"; Get-ChildItem $all -Filter "$pre~*$Member*" | Select-Object -First 15 -ExpandProperty Name; exit 1 }
$t = [IO.File]::ReadAllText($full)
$t = [regex]::Replace($t, "<script.*?</script>", "", "Singleline")
$t = [regex]::Replace($t, "<[^>]+>", " ")
$t = $t -replace "&#160;", " " -replace "&lt;", "<" -replace "&gt;", ">" -replace "&quot;", '"' -replace "&amp;", "&"
$lines = $t -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" -and $_ -notmatch "^Language Filter|^(Visual Basic|C#|C\+\+/CLI)( \(|$)" }
$s = [Array]::IndexOf($lines, "Syntax"); if ($s -lt 0) { $s = 0 }
$e = ($lines | Select-String -Pattern "^Requirements$" | Select-Object -First 1)
$end = if ($e) { $e.LineNumber - 1 } else { $lines.Count }
$lines[$s..($end - 1)] | ForEach-Object { Write-Output $_ }
