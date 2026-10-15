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

## Build and install

Close SolidWorks, then in a normal (not admin) PowerShell window:

    cd addin
    powershell -ExecutionPolicy Bypass -File .\build.ps1
    powershell -ExecutionPolicy Bypass -File .\install.ps1

`install.ps1` copies `bin\` to `%LOCALAPPDATA%\PCBThermal\` and registers the add-in for the current user only
(`HKCU\Software\Classes\CLSID\...` and `HKCU\Software\SolidWorks\AddIns\...`). Remove it with `uninstall.ps1`.

Then start SolidWorks, open the board assembly, and look for the **PCB Thermal** tab. In the task pane press
**Browse** and pick the `.swp` that holds the macro modules.

To ship to another PC: copy `bin\` and `install.ps1` / `uninstall.ps1` (the compiler is not needed there).

## Known limits

- Per-user COM registration is ignored by processes running **as administrator**. Start SolidWorks normally.
- Not yet confirmed: that SolidWorks 2023 reads the per-user add-in list (`HKCU\Software\SolidWorks\AddIns`).
  If the tab does not appear, check Tools > Add-Ins. If PCB Thermal is not listed, that key is not read and
  the add-in needs the machine-wide key once (admin), or a macro stub that calls `swApp.LoadAddIn`.
- Steps 1 to 8 still run the VBA code. Porting a step means moving its module into a `.vb` file here.

## Files

| File | What it is |
|---|---|
| `src/AddIn.vb` | Add-in entry (ISwAddin): ribbon, menu, task pane, button callbacks, runs the macros |
| `src/TaskPane.vb` | Task pane UI (WinForms, built in code) |
| `src/ExcelLink.vb` | Excel by late binding, tab lookup like `PT_TabBase` |
| `src/ModelCheck.vb` | Body, face, triangle count with timing |
| `src/Support.vb` | Settings file and log |
| `build.ps1` | Compiles with `vbc.exe`, copies the SolidWorks interop DLLs, makes the icon strips |
| `install.ps1`, `uninstall.ps1` | Per-user registration |
