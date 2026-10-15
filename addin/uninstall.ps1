# Removes the PCB Thermal add-in for the current user. Close SolidWorks first.
$guid = "{F7AF128D-DEE1-4BEB-A464-4FE0B17337DF}"
$keys = @("HKCU:\Software\Classes\CLSID\$guid", "HKCU:\Software\Classes\PCBThermal.AddIn",
          "HKCU:\Software\SolidWorks\AddIns\$guid", "HKCU:\Software\SolidWorks\AddInsStartup\$guid")
foreach ($k in $keys) { if (Test-Path $k) { Remove-Item $k -Recurse -Force } }
$dest = Join-Path $env:LOCALAPPDATA "PCBThermal"
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
Write-Host "Removed. Settings and log in %APPDATA%\PCBThermal were kept."
