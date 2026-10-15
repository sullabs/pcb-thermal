# Installs the PCB Thermal add-in for the current Windows user only (no admin rights).
# Close SolidWorks first. Copies addin\bin to %LOCALAPPDATA%\PCBThermal and registers it under HKCU.
$ErrorActionPreference = "Stop"
$guid = "{F7AF128D-DEE1-4BEB-A464-4FE0B17337DF}"
$progId = "PCBThermal.AddIn"
$className = "PCBThermal.AddIn"
$catSwAddin = "{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $here "bin"
$dest = Join-Path $env:LOCALAPPDATA "PCBThermal"
if (-not (Test-Path (Join-Path $src "PCBThermal.dll"))) { throw "Run build.ps1 first (no bin\PCBThermal.dll)." }
if (Get-Process SLDWORKS -ErrorAction SilentlyContinue) { throw "Close SolidWorks first (it locks the add-in DLL)." }

New-Item -ItemType Directory -Force $dest | Out-Null
Copy-Item (Join-Path $src "*") $dest -Recurse -Force
Get-ChildItem $dest -Recurse -File | Unblock-File
$dll = Join-Path $dest "PCBThermal.dll"
$asm = [System.Reflection.AssemblyName]::GetAssemblyName($dll)
$codeBase = "file:///" + ($dll -replace "\\", "/")

function Set-Key([string]$path, [hashtable]$values) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    foreach ($k in $values.Keys) {
        $v = $values[$k]
        if ($v -is [int]) { New-ItemProperty -Path $path -Name $k -Value $v -PropertyType DWord -Force | Out-Null }
        else { New-ItemProperty -Path $path -Name $k -Value $v -PropertyType String -Force | Out-Null }
    }
}

# COM class, per user (HKCU\Software\Classes merges into HKCR for this user)
$clsid = "HKCU:\Software\Classes\CLSID\$guid"
$inproc = @{ "(default)" = "mscoree.dll"; "ThreadingModel" = "Both"; "Class" = $className;
             "Assembly" = $asm.FullName; "RuntimeVersion" = "v4.0.30319"; "CodeBase" = $codeBase }
Set-Key $clsid @{ "(default)" = $className }
Set-Key "$clsid\InprocServer32" $inproc
Set-Key "$clsid\InprocServer32\$($asm.Version)" @{ "Class" = $className; "Assembly" = $asm.FullName;
    "RuntimeVersion" = "v4.0.30319"; "CodeBase" = $codeBase }
Set-Key "$clsid\ProgId" @{ "(default)" = $progId }
Set-Key "$clsid\Implemented Categories\$catSwAddin" @{}
Set-Key "HKCU:\Software\Classes\$progId" @{ "(default)" = $className }
Set-Key "HKCU:\Software\Classes\$progId\CLSID" @{ "(default)" = $guid }

# SolidWorks add-in list and load at startup, per user
Set-Key "HKCU:\Software\SolidWorks\AddIns\$guid" @{ "(default)" = 1; "Title" = "PCB Thermal";
    "Description" = "PCB thermal workflow: ribbon tab and task pane" }
Set-Key "HKCU:\Software\SolidWorks\AddInsStartup\$guid" @{ "(default)" = 1 }

Write-Host "Installed to $dest for user $env:USERNAME. Start SolidWorks: look for the PCB Thermal tab and Tools > Add-Ins."
