# Look up a SolidWorks API member from the installed interop DLLs (works without SolidWorks running).
# Usage (Windows PowerShell or pwsh):
#   powershell -File api_lookup.ps1 -Type IFace2 -Member Tess
#   powershell -File api_lookup.ps1 -Type IBody2            (lists all members)
#   powershell -File api_lookup.ps1 -Enum swEndConditions_e (prints names and values)
#   powershell -File api_lookup.ps1 -Dll cosworks -Type ICWStudy -Member Mesh
#   powershell -File api_lookup.ps1 -Dll cosworks -Enum "sws*Thermal*"   (Simulation enums live in cosworks)
# Dll: sldworks (default), swconst, cosworks, fworks, swpublished, ...
param(
    [string]$Type = "",
    [string]$Member = "",
    [string]$Enum = "",
    [string]$Dll = "sldworks"
)
$dir = "E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist"
if ($Enum -ne "" -and -not $PSBoundParameters.ContainsKey("Dll")) { $Dll = "swconst" }   # Simulation enums: add -Dll cosworks
$path = Join-Path $dir ("SolidWorks.Interop." + $Dll + ".dll")
if (-not (Test-Path $path)) { Write-Output "Not found: $path"; exit 1 }
$asm = [Reflection.Assembly]::LoadFrom($path)
try { $types = $asm.GetTypes() } catch { $types = $_.Exception.Types | Where-Object { $_ } }
if ($Enum -ne "") {
    $t = $types | Where-Object { $_.IsEnum -and $_.Name -like $Enum }
    foreach ($e in $t) {
        Write-Output $e.FullName
        foreach ($n in [Enum]::GetNames($e)) { Write-Output ("  {0} = {1}" -f $n, [int][Enum]::Parse($e, $n)) }
    }
    exit 0
}
$hits = $types | Where-Object { $_.Name -eq $Type -or $_.Name -like $Type }
if (-not $hits) { Write-Output "No type matches '$Type' in $Dll"; exit 1 }
foreach ($t in $hits) {
    Write-Output ("== " + $t.FullName)
    $t.GetMembers() | Where-Object { $Member -eq "" -or $_.Name -match $Member } | ForEach-Object {
        if ($_.MemberType -eq "Method") {
            $ps = ($_.GetParameters() | ForEach-Object { $_.ParameterType.Name + " " + $_.Name }) -join ", "
            Write-Output ("  " + $_.Name + "(" + $ps + ") -> " + $_.ReturnType.Name)
        } elseif ($_.MemberType -eq "Property") {
            Write-Output ("  [prop] " + $_.Name + " : " + $_.PropertyType.Name)
        }
    }
}
