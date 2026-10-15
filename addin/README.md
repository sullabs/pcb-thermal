# PCB Thermal add-in (MVP)

A SolidWorks add-in that puts the PCB thermal workflow on its own ribbon tab and task pane.
Built with the VB compiler that ships with Windows (.NET Framework 4). No Visual Studio, no Python, no admin rights.

This is an experiment on branch `addin-mvp`. The VBA macros in `../src/` are still the product.

## What the MVP does

- **Ribbon tab "PCB Thermal"** (parts and assemblies) and a **PCB Thermal menu**: 10 buttons.
  - 1 to 8 run the existing VBA macros (`SetupWorkbook.main` ... `ReadResults.main`) from the macro project you pick.
  - **Check Model** counts components, solid bodies, faces and triangles in the open model and times it
    (same kind of work as ImportStep, so it shows in-process API speed).
  - **Check Workbook** finds the open Excel workbook and checks the numbered tabs the macros need.
- **Task pane "PCB Thermal"** (right side panel): active workbook, macro project path (Browse), the 8 steps with
  their last status (double-click to run), and a log.
- Settings and a log file live in `%APPDATA%\PCBThermal\` (`settings.txt`, `log.txt`).

## Build, install, share

Build (needs only Windows and a SolidWorks install for the API files):

    powershell -ExecutionPolicy Bypass -File addin\build.ps1

Output: `addin\dist\PCBThermal_Setup_<version>.exe`, one file with everything inside: the add-in, the SolidWorks
interop DLLs, icons, the Excel template and, if `addin\macros\PCBThermal.swp` exists, the macro project.
That exe is what you share. Raise the version in `src/AssemblyInfo.vb` for every build you hand out.

Install on any PC (close SolidWorks first; no admin rights, installs for the Windows user who runs it):

- Double-click the setup exe, press **Install**. Running a newer setup upgrades in place.
- Silent: `PCBThermal_Setup_<version>.exe /silent`
- Remove: Windows Settings > Apps > Installed apps > **PCB Thermal for SolidWorks**, or `/uninstall` (`/uninstall /silent`).

Files go to `%LOCALAPPDATA%\PCBThermal\`, registry entries to `HKCU` only. Settings and log stay in `%APPDATA%\PCBThermal\`.

Then start SolidWorks, open a part or assembly, and look for the **PCB Thermal** tab.

## Porting the macros

Steps 1 to 8 still run the VBA modules from `PCBThermal.swp`. Plan: port one module at a time to `src/*.vb`
(PTCommon first, then SetupWorkbook, ImportStep, BuildBoard, BuildSpreader, SyncFromCAD, SetupSim, ReadResults, RunFlow),
check each against the VBA result on the reference board, then switch its button to the native code. When all are
ported the setup no longer needs the `.swp`.

Done so far: `src/PTCommon.vb` part 1 (Excel and workbook, values, keys, change log, files; VBA lines up to
`PT_MakeDirs`). Same names as the VBA. Checked on a copy of the reference workbook: the ported `PT_Key` gives the
same body keys as ImportStep wrote (19 of 19), and a change log row lands as the VBA writes it.
.NET gotcha: Excel error cells (`#N/A`) arrive as Int32 codes, so `PT_IsNum`/`PT_Num`/`PT_Str` treat Int32 as an error.
Next: PTCommon part 2 (SolidWorks helpers) together with SetupWorkbook and ImportStep.

## How it loads without admin

SolidWorks 2023 reads its add-in list only from `HKLM` (tested: the `HKCU\Software\SolidWorks\AddIns` key is ignored).
But `ISldWorks.LoadAddIn("<path>\PCBThermal.dll")` works with the per-user COM registration. So setup adds a small
loader, `PCBThermal_Setup.exe /watch`, to `HKCU\...\CurrentVersion\Run`. It starts at Windows logon, waits for each
SolidWorks start (polls every 3 s, waits for `StartupProcessCompleted`) and calls LoadAddIn. Tested: the tab shows
about 15 s after SolidWorks starts. Log: `%APPDATA%\PCBThermal\loader.log`.

- Untick **Start Up** for PCBThermal in Tools > Add-Ins and the loader stops loading it.
- Start menu > **Load PCB Thermal into SolidWorks** (`/load`) loads it once into the running SolidWorks.
- In Tools > Add-Ins it shows under Other Add-ins as "PCBThermal".
- Scripted install: `Start-Process ... -Wait` also waits for the loader and never returns. Use `-PassThru` and
  `$p.WaitForExit()` instead.

## Known limits

- Per-user COM registration is ignored by processes running **as administrator**. Start SolidWorks normally.
- Steps 1 to 8 still run the VBA code. Porting a step means moving its module into a `.vb` file here.

## Files

| File | What it is |
|---|---|
| `src/AddIn.vb` | Add-in entry (ISwAddin): ribbon, menu, task pane, button callbacks, runs the macros |
| `src/TaskPane.vb` | Task pane UI (WinForms, built in code) |
| `src/ExcelLink.vb` | Excel by late binding, tab lookup like `PT_TabBase` |
| `src/ModelCheck.vb` | Body, face, triangle count with timing |
| `src/Support.vb` | Settings file and log |
| `build.ps1` | Compiles the add-in and the setup exe with `vbc.exe`, makes the icon strips |
| `src/AssemblyInfo.vb` | Version number |
| `setup/Setup.vb`, `setup/setup.manifest` | The setup exe: unpacks the embedded files, per-user registration, uninstall entry |
| `setup/Loader.vb` | `/watch` and `/load`: loads the add-in into SolidWorks with LoadAddIn |
