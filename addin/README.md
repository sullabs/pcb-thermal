# NTIS thermal simulation add-in

A SolidWorks add-in that runs the whole PCB thermal workflow from its own ribbon tab and task pane.
Built with the VB compiler that ships with Windows (.NET Framework 4). No Visual Studio, no Python, no admin rights,
no macro project (`.swp`) needed.

Branch `addin-mvp`. The VBA macros in `../src/` stay the reference: every step here is a port of one of them.

## What it does

- **Ribbon tab "NTIS thermal simulation"** (parts and assemblies) and a menu of the same name: 10 buttons.
  - 1 to 8 run the workflow steps (Prepare Workbook, Import STEP, Build Board, Build Spreader, Sync CAD to Excel,
    Run Thermal Study, Run Flow Sim, Read Results). Same behaviour, workbook cells and messages as the VBA macros.
  - **Check Model** counts components, solid bodies, faces and triangles in the open model and times it.
  - **Check Workbook** finds the open Excel workbook and checks the numbered tabs the steps need.
- **Task pane "NTIS thermal simulation"**: active workbook, the 8 steps with their last status (double-click to run),
  what the selected step does and when to use it, a log.
  Tick **Run the VBA macros (.swp) instead** to run the old macros from a macro project (for comparison only).
- Every button has a full tooltip: what it does, then when to use it (`CommandHelp` in `src/AddIn.vb`).
- Every message box text also goes to the log. Settings and log live in `%APPDATA%\PCBThermal\`
  (`settings.txt`, `log.txt`, `loader.log`).

## Build, install, share

Build (needs only Windows and a SolidWorks install for the API files):

    powershell -ExecutionPolicy Bypass -File addin\build.ps1

Output: `addin\dist\NTIS_thermal_simulation_Setup_<version>.exe`, one file with everything inside: the add-in, the SolidWorks
interop DLLs (sldworks, swconst, swpublished, cosworks), icons and the Excel template. That exe is what you share.
Raise the version in `src/AssemblyInfo.vb` for every build you hand out.

Install on any PC (close SolidWorks first; no admin rights, installs for the Windows user who runs it):

- Double-click the setup exe, press **Install**. Running a newer setup upgrades in place.
- Silent: `NTIS_thermal_simulation_Setup_<version>.exe /silent`
- Remove: Windows Settings > Apps > Installed apps > **NTIS thermal simulation for SolidWorks**, or `/uninstall` (`/uninstall /silent`).

Files go to `%LOCALAPPDATA%\PCBThermal\`, registry entries to `HKCU` only. Then start SolidWorks normally and look
for the **NTIS thermal simulation** tab.

## How it loads without admin

SolidWorks 2023 reads its add-in list only from `HKLM` (tested: the `HKCU\Software\SolidWorks\AddIns` key is ignored).
But `ISldWorks.LoadAddIn("<path>\NTIS_thermal_simulation.dll")` works with the per-user COM registration. So setup adds a small
loader, `PCBThermal_Setup.exe /watch`, to `HKCU\...\CurrentVersion\Run`. It starts at Windows logon, waits for each
SolidWorks start (polls every 3 s, waits for `StartupProcessCompleted`) and calls LoadAddIn. The tab shows about
15 s after SolidWorks starts.

- Untick **Start Up** for the add-in in Tools > Add-Ins and the loader stops loading it.
- Start menu > **Load NTIS thermal simulation into SolidWorks** (`/load`) loads it once into the running SolidWorks.
- In Tools > Add-Ins it shows under Other Add-ins.
- Scripted install: `Start-Process ... -Wait` also waits for the loader and never returns. Use `-PassThru` and
  `$p.WaitForExit()` instead.

## The port

`src/Macros/*.vb` holds one module per VBA macro, `src/PTCommon.vb` and `src/PTCommon2.vb` hold PTCommon. Same
procedure names, same order, same comments, so a change in a `.bas` file can be carried over line by line.
`src/VbaCompat.vb` keeps the VBA idioms working (`MsgBox`, `Arr` for `Array`, `IsEmpty`, `Round`, `Dict` for
`Scripting.Dictionary`, `NewTable` for 1-based arrays).

Rules learned while porting (apply them to any new code):

- SolidWorks rejects late-bound `ByRef` Integer, String and Boolean arguments from .NET ("Type mismatch"; they arrive
  as `VARIANT*`). Every call with such an argument is typed: `SwOpenDoc6`, `SwSaveAs`, `SwSave3`, `SwGetOpenFileName`,
  `SwLoadFile4`, `ActivateDoc3`, `CustomPropertyManager.Get6`, and every Simulation call with an `ErrorCode` (cosworks interop).
- `obj.GetType` on an `Object` is .NET's `Object.GetType`. Use `SwDocType(md)` for the SolidWorks document type.
- Excel error cells (`#N/A`) arrive as Int32 codes; `PT_IsNum`, `PT_Num`, `PT_Str` treat them as errors.
- SolidWorks returns VT_NULL (DBNull) where VBA saw Empty; `IsEmpty` treats both as empty.
- Pass `NullDisp` where the VBA passed `Nothing` to a late-bound SolidWorks call.

## Test result (2026-10-15, Raspberry Pi 4 reference)

Each step was run with `AddIn.TestRun` (quiet, scripted answers) on copies of the reference workbooks and a Pack and Go
copy of the VBA-built assemblies, and compared with what the VBA macros had produced:

| Step | Check | Result |
|---|---|---|
| ImportStep | re-import of the same STEP into the VBA-imported workbook | 109 rows, all values equal by key, 0 change log rows |
| BuildBoard | update on the VBA assembly; create in a new folder | nothing changed; all 15 PCB, part and TIM boxes equal |
| BuildSpreader | update on the VBA assembly (top and bottom plate); create | 0 changes, report equal to VBA; rectangle plate made |
| SyncFromCAD | on the VBA assembly | 0 overrides, 0 differences, 0 failed checks |
| SetupSim + ReadResults | new study on the VBA assembly | report equal; every temperature equal to VBA to 0.01 C |
| RunFlow | read the last Flow results | every value equal to the VBA row |
| SetupWorkbook | template path; full refresh path | both run without errors |

Times on this PC: ImportStep 274 s (109 bodies, almost all inside SolidWorks' `GetTessTriangles`), BuildBoard
update 6 s, create 57 s, BuildSpreader 32 s, SetupSim 194 s (mesh and solve), ReadResults 3 s.

## Known limits

- Per-user COM registration is ignored by processes running **as administrator**. Start SolidWorks normally.
- ImportStep is slow on boards with many small multi-face bodies (about 38 ms per face in SolidWorks). The VBA macro
  makes the same calls.
- Row order on STEP import follows the order SolidWorks lists the components, which can differ between sessions.
  Inputs follow their body by key, so nothing is lost.

## Files

| File | What it is |
|---|---|
| `src/AddIn.vb` | Add-in entry (ISwAddin): ribbon, menu, task pane, runs the steps, `TestRun` for scripted tests |
| `src/TaskPane.vb` | Task pane UI (WinForms, built in code) |
| `src/PTCommon.vb`, `src/PTCommon2.vb` | Port of PTCommon.bas (workbook helpers; SolidWorks helpers and typed calls) |
| `src/VbaCompat.vb` | VBA idioms for the port, message boxes with a quiet mode |
| `src/Macros/*.vb` | Port of each macro: SetupWorkbook, ImportStep, BuildBoard, BuildSpreader, SyncFromCAD, SetupSim, ReadResults, RunFlow |
| `src/ExcelLink.vb` | Check Workbook |
| `src/ModelCheck.vb` | Check Model |
| `src/Support.vb` | Settings file and log |
| `build.ps1` | Compiles the add-in and the setup exe with `vbc.exe`, makes the icon strips from `icons/png/` |
| `icons/svg/` | Icon sources, one per command in ribbon order (`11-main.svg` is the add-in icon) |
| `icons/render.ps1` | Renders `icons/svg/` with Edge headless to `icons/png/` (20 to 128 px) and to the VBA toolbar bitmaps in `../src/icons/`. Run it after changing an SVG and commit the output |
| `src/AssemblyInfo.vb` | Version number |
| `setup/Setup.vb`, `setup/setup.manifest` | The setup exe: unpacks the embedded files, per-user registration, uninstall entry |
| `setup/Loader.vb` | `/watch` and `/load`: loads the add-in into SolidWorks with LoadAddIn |

## Names

Shown to the user: **NTIS thermal simulation** (ribbon tab, menu, task pane, message boxes, setup, Installed apps).
Internal names stay as before so upgrades and settings carry over: install folder `%LOCALAPPDATA%\PCBThermal`,
settings `%APPDATA%\PCBThermal`, COM class `PCBThermal.AddIn`, same GUID. The DLL is `NTIS_thermal_simulation.dll`:
no spaces, because COM could not load a DLL with spaces in its name (file not found). Setup removes the old
`PCBThermal.dll` and `NTIS_thermal_similation.dll`, the old Start menu shortcuts and the old COM entries.
