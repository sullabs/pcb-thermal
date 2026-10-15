# Update-TemplateColors.ps1
# Adds the part colours to the PCB thermal workbook and the two-resistor inputs the template lacks.
# Run it from YOUR desktop PowerShell (not from Claude Code: Excel COM is blocked in that session):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\Update-TemplateColors.ps1
# Optional: -Files "path1.xlsx","path2.xlsx". Default: the template and example\RaspberryPi4\RaspberryPi4.xlsx.
# Safe to run twice. Each file is copied to <name>.bak_<time> first. Close the workbooks in Excel before running.
param([string[]]$Files)

$root = Split-Path -Parent $PSScriptRoot
if (-not $Files) { $Files = @("$root\workbook\PCB_thermal_TEMPLATE.xlsx", "$root\example\RaspberryPi4\RaspberryPi4.xlsx") }

# same 12 colours as PT_PartRgb in PTCommon.bas, and PT_PcbRgb for the board
$pal = @(@(66,133,244),@(244,160,0),@(219,68,55),@(171,71,188),@(0,172,193),@(255,202,40),
         @(141,110,99),@(236,64,122),@(92,107,192),@(255,112,67),@(120,144,156),@(38,198,218))
$pcb = @(16,140,39)
function Rgb($c) { return [int]($c[0] + 256 * $c[1] + 65536 * $c[2]) }
function FontFor($c) { if ((0.299*$c[0] + 0.587*$c[1] + 0.114*$c[2]) -lt 140) { return 16777215 } else { return 0 } }

function Find-Sheet($wb, $base) {
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq $base -or $s.Name -match "^[0-9]+ $([regex]::Escape($base))$") { return $s } }
    throw "tab '$base' not found"
}
# removes the rules this script made earlier (they mention Geo_Include or PT_PCB_MARK), so a second run does not stack them
function Drop-OurRules($rng) {
    for ($i = $rng.FormatConditions.Count; $i -ge 1; $i--) {
        $f = $null; try { $f = $rng.FormatConditions.Item($i).Formula1 } catch {}
        if ($f -and ($f -like '*Geo_Include*' -or $f -like '*PT_PCB_MARK*' -or $f -like '*MOD(ROW()-26,12)*')) { $rng.FormatConditions.Item($i).Delete() }
    }
}
function Add-Rule($rng, $formula, $c) {
    $fc = $rng.FormatConditions.Add(2, 0, $formula)        # xlExpression
    $fc.Interior.Color = (Rgb $c)
    $fc.Font.Color = (FontFor $c)
    $fc.StopIfTrue = $true
}
function Add-Note($cell, $txt) { if ($cell.Comment) { $cell.Comment.Delete() }; [void]$cell.AddComment($txt) }
function Has-Name($wb, $nm) { foreach ($n in $wb.Names) { if ($n.Name -eq $nm) { return $true } }; return $false }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    foreach ($path in $Files) {
        if (-not (Test-Path $path)) { Write-Host "SKIP (not found): $path"; continue }
        $bak = "$path.bak_" + (Get-Date -Format 'yyyyMMdd_HHmmss')
        Copy-Item $path $bak
        $wb = $xl.Workbooks.Open($path)
        if ($wb.ReadOnly) { Write-Host "SKIP (read only, close it in Excel): $path"; $wb.Close($false); continue }
        Write-Host "== $path"

        $reg = Find-Sheet $wb 'Component register'
        $stp = Find-Sheet $wb 'STEP import'
        $geo = Find-Sheet $wb 'CAD geometry'
        $sim = Find-Sheet $wb 'Sim setup'

        # names used by the colour rules (body rows of CAD geometry: 26 to 95)
        $g = "'" + $geo.Name + "'!"
        [void]$wb.Names.Add('Geo_Include', "=$g`$G`$26:`$G`$95")
        [void]$wb.Names.Add('Geo_RefDes',  "=$g`$F`$26:`$F`$95")

        # rules. Palette index k = (row on CAD geometry - 26) mod 12, shown only when Include = Yes
        $rGeo = $geo.Range('E26:E95'); $rReg = $reg.Range('B8:B53'); $rStp = $stp.Range('C9:C88'); $rPcb = $stp.Range('C8')
        foreach ($r in @($rGeo, $rReg, $rStp, $rPcb)) { Drop-OurRules $r }
        $fc = $rPcb.FormatConditions.Add(2, 0, '=AND($C8<>"",LEN("PT_PCB_MARK")>0)')
        $fc.Interior.Color = (Rgb $pcb); $fc.Font.Color = (FontFor $pcb)
        for ($k = 0; $k -lt 12; $k++) {
            Add-Rule $rGeo ('=AND($E26<>"",$G26="Yes",MOD(ROW()-26,12)=' + $k + ')') $pal[$k]
            Add-Rule $rStp ('=AND($C9<>"",IFERROR(INDEX(Geo_Include,ROW()-8),"")="Yes",MOD(ROW()-9,12)=' + $k + ')') $pal[$k]
            Add-Rule $rReg ('=AND($B8<>"",IFERROR(INDEX(Geo_Include,MATCH($B8,Geo_RefDes,0)),"")="Yes",MOD(IFERROR(MATCH($B8,Geo_RefDes,0),0)-1,12)=' + $k + ')') $pal[$k]
        }
        Add-Note $stp.Range('C7') 'Name of the body. When its CAD geometry row has Include = Yes it is filled with the colour the part has in the SolidWorks thermal assembly (PCB green). Same colour on CAD geometry (Name used) and the Component register (RefDes).'
        Add-Note $geo.Range('E25') 'Name used. Filled with the colour of the part in the SolidWorks thermal assembly when Include = Yes. 12 colours by row, PCB green.'
        Add-Note $reg.Range('B7')  'RefDes. Filled with the part colour in the SolidWorks thermal assembly when the part with this RefDes has Include = Yes on CAD geometry.'

        # two-resistor inputs the template lacked (SetupSim reads Sim_LayerFrac and Sim_kCore)
        if (-not (Has-Name $wb 'Sim_LayerFrac')) {
            [void]$sim.Range('C11').Copy()
            [void]$sim.Range('C12:C13').PasteSpecial(-4122)          # formats only: blue input style
            $xl.CutCopyMode = 0
            $sim.Range('B12').Value2 = 'Two-resistor: outer layer thickness'
            $sim.Range('C12').Value2 = 0.1
            $sim.Range('D12').Value2 = 'fraction of the chip height, for each of the top and bottom layers'
            $sim.Range('B13').Value2 = 'Two-resistor: k of the junction core'
            $sim.Range('C13').Value2 = 100
            $sim.Range('D13').Value2 = 'W/m K, high so the core is one temperature'
            [void]$wb.Names.Add('Sim_LayerFrac', "='" + $sim.Name + "'!`$C`$12")
            [void]$wb.Names.Add('Sim_kCore',     "='" + $sim.Name + "'!`$C`$13")
            Add-Note $sim.Range('C12') 'Thickness of the top and bottom layer of a two-resistor chip as a fraction of the chip height (above 0, below 0.45). Layer k = layer thickness / (theta x footprint area).'
            Add-Note $sim.Range('C13') 'Thermal conductivity of the middle layer (junction core) of a two-resistor chip, W/m K. High, so the core is one temperature.'
            Write-Host '   added Sim_LayerFrac and Sim_kCore'
        }

        # checks
        $ok = $true
        if ($reg.Range('D8').Validation.Type -ne 3) { Write-Host '   !! Class dropdown on the Component register is not a list'; $ok = $false }
        $missing = @()
        foreach ($nm in 'Geo_Table','Geo_PCB_Used','Reg_RefDes','Reg_QUsed','Reg_TotalQ','Imp_Table','Sim_Chips','Spr_Chips','Log_Header','Set_Prefix','Sim_LayerFrac','Sim_kCore','Geo_Include','Geo_RefDes') {
            if (-not (Has-Name $wb $nm)) { $missing += $nm }
        }
        if ($missing.Count) { Write-Host ("   !! missing names: " + ($missing -join ', ')); $ok = $false }
        Write-Host ("   colour rules: geo {0}, step {1}, register {2}" -f $rGeo.FormatConditions.Count, $rStp.FormatConditions.Count, $rReg.FormatConditions.Count)
        $xl.CalculateFull()
        $wb.Save()
        $wb.Close($true)
        Write-Host ("   saved. backup: $bak" + $(if ($ok) { '  checks OK' } else { '  CHECKS FAILED, see above' }))
    }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
