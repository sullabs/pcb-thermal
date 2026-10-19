# Makes the "PCB Thermal" CommandManager tab in SolidWorks (parts and assemblies): one macro button per step,
# each with its icon, tooltip and prompt, all pointing to PCBThermal.swp in this folder.
#
# Run it with SolidWorks closed (SolidWorks rewrites its settings when it closes):
#   double-click SetupToolbar.cmd, or  powershell -ExecutionPolicy Bypass -File SetupToolbar.ps1
# Options: -Version 2025 (default 2023), -Swp <path to the .swp>, -Remove (takes the tab and buttons out again).
#
# What it writes (HKCU only, found by making one button by hand in SolidWorks 2023, 2026-10-19):
#   SOLIDWORKS <ver>\User Defined Macros\NN - Macro Folder   one key per macro button (macro, method, tooltip,
#                                                            prompt, bitmap, "Command Offset")
#   ...\User Interface\CommandManager\PartContext\TabN       RefName / Tab Props = tab name; GB0\Btn<i> = "2,<id>"
#   ...\User Interface\CommandManager\AssyContext\TabN       the same for assemblies
#   <id> = 33639 + Command Offset.
# Buttons of other macros are kept. A copy of the settings goes to %APPDATA%\PCBThermal first.
param(
    [string]$Version = "2023",
    [string]$Swp = "",
    [switch]$Remove,
    [string]$RootBase = "HKCU:\Software\SolidWorks",   # for tests on a copy of the settings
    [switch]$IgnoreRunning                              # for tests only
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Swp) { $Swp = Join-Path $here "PCBThermal.swp" }
$Swp = [System.IO.Path]::GetFullPath($Swp)
$icons = Join-Path $here "icons"
$tabName = "PCB Thermal"
$firstCommand = 33639

# Tooltip, module (method is main), icon, prompt. Order = order on the tab.
$buttons = @(
    @("Prepare Workbook", "SetupWorkbook", "1_PrepareWorkbook_32.bmp",
      "Old workbook only. Fills empty chip lists on a template copy, or adds missing tabs and named cells to an older workbook."),
    @("Import STEP", "ImportStep", "2_ImportStep_32.bmp",
      "Reads the board STEP (or the open model) and writes the size and position of every body to cad.json (or the workbook)."),
    @("Build Board", "BuildBoard", "3_BuildBoard_32.bmp",
      "Builds or updates the thermal assembly: one box for the PCB and one per included part. Refuses while changes are unticked."),
    @("Build Spreader", "BuildSpreader", "4_BuildSpreader_32.bmp",
      "Builds or updates the heat spreader plate, pedestals and TIM parts from tab 7 of the page (or the Spreader tab). Run Build Board first."),
    @("Sync CAD", "SyncFromCAD", "5_SyncCADtoExcel_32.bmp",
      "Reads the thermal assembly back: sizes, the plate as built, checks. Writes cad.json (or the workbook). Changes no CAD file."),
    @("Run Thermal Study", "SetupSim", "6_RunThermalStudy_32.bmp",
      "Makes a new Simulation thermal study from tab 8 of the page (or Sim setup), meshes and runs it. Then press Read Results."),
    @("Run Flow Simulation", "RunFlow", "7_RunFlowSim_32.bmp",
      "Sets heat sources and ambient on your Flow Simulation project, solves it and writes the goals to results.json (or Sim results)."),
    @("Read Results", "ReadResults", "8_ReadResults_32.bmp",
      "Reads the solved thermal study and writes the temperatures to results.json, tab 9 of the page (or the workbook)."),
    @("Update Macros", "UpdateMacros", "0_UpdateMacros_32.bmp",
      "Imports every .bas file from the src folder into PCBThermal.swp, compiles and saves it. Use it after the code changes.")
)

$root = "$RootBase\SOLIDWORKS $Version"
$macrosKey = "$root\User Defined Macros"
$contexts = "PartContext", "AssyContext"

function Fail([string]$msg) { Write-Host ""; Write-Host $msg -ForegroundColor Red; exit 1 }

if (-not $IgnoreRunning -and (Get-Process sldworks -ErrorAction SilentlyContinue)) {
    Fail "SolidWorks is running. Close it first (check Task Manager for SLDWORKS.exe), then run this again."
}
if (-not (Test-Path "$root\User Interface\CommandManager")) {
    Fail "No SolidWorks $Version settings found. Start and close SolidWorks $Version once, then run this again."
}
if (-not $Remove) {
    if (-not (Test-Path $Swp)) { Write-Host "Note: $Swp does not exist yet. The buttons work once it does." -ForegroundColor Yellow }
    foreach ($b in $buttons) {
        if (-not (Test-Path (Join-Path $icons $b[2]))) { Fail "Icon missing: $(Join-Path $icons $b[2])" }
    }
}

# Copy of the settings before any change.
$backupDir = Join-Path $env:APPDATA "PCBThermal"
New-Item -ItemType Directory -Force $backupDir | Out-Null
$backup = Join-Path $backupDir ("SolidWorks_{0}_settings_{1}.reg" -f $Version, (Get-Date -Format "yyyyMMdd_HHmmss"))
& reg.exe export ($root -replace '^HKCU:', 'HKCU') $backup /y | Out-Null
if ($LASTEXITCODE -ne 0) { Fail "Could not save a copy of the settings to $backup. Nothing was changed." }

# ---------- macro buttons ----------
# Read every existing button; keep the ones that are not ours (same folder order and command offset).
$kept = @()
if (Test-Path $macrosKey) {
    foreach ($k in Get-ChildItem $macrosKey | Sort-Object PSChildName) {
        $p = Get-ItemProperty $k.PSPath
        $values = [ordered]@{}
        foreach ($n in $k.GetValueNames()) { $values[$n] = @($k.GetValue($n), $k.GetValueKind($n)) }
        $ours = ($p."Source Path" -and ([System.IO.Path]::GetFullPath($p."Source Path") -ieq $Swp))
        if (-not $ours) { $kept += , $values }
    }
    Get-ChildItem $macrosKey | Remove-Item -Recurse
} else {
    New-Item -Path $macrosKey | Out-Null
}
$used = @{}
foreach ($v in $kept) { $used[[int]$v["Command Offset"][0]] = $true }

$folder = 0
foreach ($v in $kept) {
    $folder++
    $key = New-Item -Path ("{0}\{1:D2} - Macro Folder" -f $macrosKey, $folder)
    foreach ($n in $v.Keys) { New-ItemProperty -Path $key.PSPath -Name $n -Value $v[$n][0] -PropertyType $v[$n][1] | Out-Null }
}

$commandIds = @()
if (-not $Remove) {
    $offset = 0
    foreach ($b in $buttons) {
        while ($used[$offset]) { $offset++ }
        $used[$offset] = $true
        $folder++
        $key = New-Item -Path ("{0}\{1:D2} - Macro Folder" -f $macrosKey, $folder)
        $s = @{
            "Project" = $b[1]; "MacroMethod" = "main"; "ToolTip" = $b[0]; "Prompt Msg" = $b[3]
            "Source Path" = $Swp; "Bitmap Path" = (Join-Path $icons $b[2]); "Accelerator" = ""
        }
        New-ItemProperty -Path $key.PSPath -Name "Command Offset" -Value $offset -PropertyType DWord | Out-Null
        foreach ($n in "Project", "MacroMethod", "ToolTip", "Prompt Msg", "Source Path", "Bitmap Path") {
            New-ItemProperty -Path $key.PSPath -Name $n -Value $s[$n] -PropertyType String | Out-Null
        }
        foreach ($n in "Mouse Gesture Part", "Mouse Gesture Assembly", "Mouse Gesture Drawing", "Mouse Gesture Sketch") {
            New-ItemProperty -Path $key.PSPath -Name $n -Value 0 -PropertyType DWord | Out-Null
        }
        New-ItemProperty -Path $key.PSPath -Name "Accelerator" -Value "" -PropertyType String | Out-Null
        $commandIds += ($firstCommand + $offset)
    }
}
Set-ItemProperty -Path $macrosKey -Name "Macro Count" -Value $folder -Type DWord

# ---------- the tab ----------
function Get-TabKeys([string]$cm) {
    Get-ChildItem $cm | Where-Object { $_.PSChildName -match '^Tab(\d+)$' } |
        Sort-Object { [int]($_.PSChildName -replace 'Tab', '') }
}
foreach ($ctx in $contexts) {
    $cm = "$root\User Interface\CommandManager\$ctx"
    if (-not (Test-Path $cm)) { continue }
    $tabs = @(Get-TabKeys $cm)
    $mine = $tabs | Where-Object { (Get-ItemProperty $_.PSPath).RefName -eq $tabName } | Select-Object -First 1
    if ($Remove) {
        if ($mine) {
            # Take the tab out and move the tabs after it down by one, so the numbers stay without gaps.
            $n = [int]($mine.PSChildName -replace 'Tab', '')
            Remove-Item $mine.PSPath -Recurse
            foreach ($t in $tabs | Where-Object { [int]($_.PSChildName -replace 'Tab', '') -gt $n }) {
                $i = [int]($t.PSChildName -replace 'Tab', '')
                Copy-Item -Path $t.PSPath -Destination "$cm\Tab$($i - 1)" -Recurse
                Remove-Item $t.PSPath -Recurse
            }
        }
        continue
    }
    if ($mine) {
        $tabPath = $mine.PSPath
        Get-ChildItem $tabPath | Remove-Item -Recurse
    } else {
        $next = 0
        if ($tabs.Count -gt 0) { $next = 1 + [int]($tabs[-1].PSChildName -replace 'Tab', '') }
        $tabPath = (New-Item -Path "$cm\Tab$next").PSPath
    }
    Set-ItemProperty -Path $tabPath -Name "RefName" -Value $tabName -Type String
    Set-ItemProperty -Path $tabPath -Name "Tab Props" -Value "$tabName,0,1,-1" -Type String
    $gb = (New-Item -Path "$tabPath\GB0").PSPath
    for ($i = 0; $i -lt $commandIds.Count; $i++) {
        New-ItemProperty -Path $gb -Name "Btn$i" -Value ("2,{0}" -f $commandIds[$i]) -PropertyType String | Out-Null
    }
}

Write-Host ""
if ($Remove) {
    Write-Host "Removed the $tabName tab and its buttons from SolidWorks $Version." -ForegroundColor Green
} else {
    Write-Host "Made the $tabName tab with $($buttons.Count) buttons in SolidWorks $Version (parts and assemblies)." -ForegroundColor Green
    Write-Host "Buttons run: $Swp"
    Write-Host "Start SolidWorks and open a part or assembly to see the tab."
}
Write-Host "Copy of the old settings: $backup (double-click it to put them back, with SolidWorks closed)."
