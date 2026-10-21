# PCB Thermal tab in SolidWorks

Result: a CommandManager tab named **PCB Thermal** in parts and assemblies with 9 buttons: the 8 steps and
**Update Macros**. Each button has its icon, a tooltip (the name) and a prompt (what it does, shown in the status bar).
All buttons run modules of one macro project, `src\PCBThermal.swp`. No add-in is needed.

Two files do the work:

- `src\UpdateMacros.bas`: imports every `.bas` file from `src\` into the `.swp` it sits in, compiles and saves it.
- `src\SetupToolbar.cmd` (runs `SetupToolbar.ps1`): writes the tab and the buttons into the SolidWorks settings.
  SolidWorks has no API for this from a macro, so it writes the same registry entries that Tools > Customize writes.

## Once

1. SolidWorks: Tools > Macro > New, save as `src\PCBThermal.swp`.
2. In the VBA editor: File > Import File, pick `src\UpdateMacros.bas`. Delete the empty `Module1`.
3. Run `UpdateMacros.main` (F5). It imports the other modules, compiles and saves. The message lists them.
   If it says "Not saved yet", press Debug > Compile VBAProject, then Ctrl+S.
4. Close SolidWorks. Check Task Manager: no `SLDWORKS.exe` may be left.
5. Double-click `src\SetupToolbar.cmd`. It says "Made the PCB Thermal tab with 9 buttons".
6. Start SolidWorks and open a part or assembly. The **PCB Thermal** tab is there and stays after a restart.

Other SolidWorks version: `SetupToolbar.cmd -Version 2025`. Take it out again: `SetupToolbar.cmd -Remove`.
Before each change the script saves the old settings to `%APPDATA%\PCBThermal\SolidWorks_<ver>_settings_<date>.reg`;
double-click that file (SolidWorks closed) to put them back.

## After a code change

Press **Update Macros** on the tab. It replaces every module from `src\` and saves the `.swp`. The other buttons keep
working. Only `UpdateMacros.bas` itself is not replaced: after a change to it, remove it in the VBA editor and import
it again.

Run `SetupToolbar.cmd` again only when the buttons change (new step, new icon, new text) or the project folder moves.
It keeps buttons of other macros and replaces only its own.

## The buttons

| # | Tooltip | Module | Icon |
|---|---|---|---|
| 1 | Prepare Workbook | `SetupWorkbook` | `1_PrepareWorkbook_32.bmp` |
| 2 | Import STEP | `ImportStep` | `2_ImportStep_32.bmp` |
| 3 | Build Board | `BuildBoard` | `3_BuildBoard_32.bmp` |
| 4 | Build Spreader | `BuildSpreader` | `4_BuildSpreader_32.bmp` |
| 5 | Sync CAD to Excel | `SyncFromCAD` | `5_SyncCADtoExcel_32.bmp` |
| 6 | Run Thermal Study | `SetupSim` | `6_RunThermalStudy_32.bmp` |
| 7 | Run Flow Simulation | `RunFlow` | `7_RunFlowSim_32.bmp` |
| 8 | Read Results | `ReadResults` | `8_ReadResults_32.bmp` |
| 9 | Update Macros | `UpdateMacros` | `0_UpdateMacros_32.bmp` |

Method is `main` for all. Prompts are in `SetupToolbar.ps1`. Icons are in `src\icons\`; sources in `addin\icons\svg\`,
`addin\icons\render.ps1` remakes the bitmaps (`0_UpdateMacros` is made from `11-main`).

## How SolidWorks stores the buttons (found 2026-10-19, SolidWorks 2023)

- `HKCU\Software\SolidWorks\SOLIDWORKS 2023\User Defined Macros`: `Macro Count`, and one key `NN - Macro Folder` per
  button with `Project` (module), `MacroMethod`, `ToolTip`, `Prompt Msg`, `Source Path` (.swp), `Bitmap Path`,
  `Command Offset`, `Accelerator` and four `Mouse Gesture ...` values. Written as soon as Customize closes with OK.
- `...\User Interface\CommandManager\PartContext\TabN` (and `AssyContext`): `RefName` and `Tab Props` = `<name>,0,1,-1`,
  buttons in `GB0` as `Btn0`, `Btn1`, ... = `2,<command id>`, command id = 33639 + `Command Offset`. Written when
  SolidWorks closes, so a running SolidWorks overwrites any change made meanwhile.
- A macro can reach its own VBA project (`Application.VBE`), import and remove modules: tested with
  `tests\swp\VbeProbe.bas`.

## Notes

- If you move the project folder, run `SetupToolbar.cmd` again from the new `src\` folder.
- `tests\swp\PTTest.swp` stays the test harness. The buttons use `PCBThermal.swp`, so tests do not touch them.
- Every macro reuses the workbook or project of the last run, keeps a daily backup before it writes, and logs to
  `%APPDATA%\PCBThermal\log.txt`.
