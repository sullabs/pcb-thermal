# PCB Thermal tab in SolidWorks

Result: a CommandManager tab named **PCB Thermal** in parts and assemblies. It has one button per macro.
Each button runs that macro on its own. Hover a button to see its name. The description shows in the
status bar at the bottom of SolidWorks.

No add-in is needed. These are standard SolidWorks macro buttons.

## 1. One macro project for the buttons

All buttons point to one file: `src\PCBThermal.swp`. It holds `PTCommon` plus the 8 macro modules, so
every macro finds the shared code.

1. Tools > Macro > New, save as `src\PCBThermal.swp`.
2. In the VBA editor: File > Import File (Ctrl+M), import `PTCommon.bas` and the 8 macro `.bas` files from `src\`.
   Delete the empty `Module1`.
3. Debug > Compile VBAProject, then Ctrl+S.

When a `.bas` changes later: remove that module, import it again, compile, save. The buttons keep working.

## 2. Make the tab

1. Open any part (do it again later for an assembly, or use step 4).
2. Right-click any CommandManager tab > Add Tab > Empty Tab. Right-click the new tab > Rename Tab > `PCB Thermal`.
3. Tools > Customize > Commands tab > Category **Macro**. Drag **New Macro Button** onto the PCB Thermal tab.
   The Customize Macro Button box opens. Fill it from the table below and press OK. Repeat for all 8.
4. Right-click the tab > Copy Tab to Assembly (or repeat steps 2 and 3 with an assembly open).
5. Close Customize.

Fields in the Customize Macro Button box:

- **Image**: the `_32.bmp` file in `src\icons\` (use `_16.bmp` if SolidWorks refuses the 32 px one)
- **Tooltip**: the button name
- **Prompt**: the description
- **Macro**: `...\thermal_similation\src\PCBThermal.swp`
- **Method**: the module, for example `ImportStep.main`

## 3. The 8 buttons (in workflow order)

| # | Tooltip (name) | Method | Image | Prompt (description) |
|---|---|---|---|---|
| 1 | Prepare Workbook | `SetupWorkbook.main` | `1_PrepareWorkbook_32.bmp` | Not needed for a copy of PCB_thermal_TEMPLATE. On a template copy it fills empty chip lists. On an older workbook it adds the missing tabs and named cells. |
| 2 | Import STEP | `ImportStep.main` | `2_ImportStep_32.bmp` | Reads a PCB STEP file (or the open model) and writes the size and position of every body to the STEP import tab. |
| 3 | Build Board | `BuildBoard.main` | `3_BuildBoard_32.bmp` | Builds or updates the thermal assembly: one box for the PCB and one per included part, from the CAD geometry tab. |
| 4 | Build Spreader | `BuildSpreader.main` | `4_BuildSpreader_32.bmp` | Builds or updates the heat spreader plate, pedestals and TIM parts from the Spreader tab. Run Build Board first. |
| 5 | Sync CAD to Excel | `SyncFromCAD.main` | `5_SyncCADtoExcel_32.bmp` | Reads what you changed in the thermal assembly and writes it back to the workbook. Does not change any CAD file. |
| 6 | Run Thermal Study | `SetupSim.main` | `6_RunThermalStudy_32.bmp` | Makes a new SolidWorks Simulation thermal study from the Sim setup tab, meshes it and runs it. Bare board study when there is no spreader. Then press Read Results. |
| 7 | Run Flow Simulation | `RunFlow.main` | `7_RunFlowSim_32.bmp` | Sets heat sources and ambient on your Flow Simulation project, solves it and writes the goals to Sim results. |
| 8 | Read Results | `ReadResults.main` | `8_ReadResults_32.bmp` | Reads the solved thermal study and writes temperatures to Sim results and Thermal solution column AF. |

Icon colours: green = writes to Excel, blue = builds CAD, orange = runs a simulation.

## Notes

- `Launcher.bas` (the number menu) is retired. The buttons replace it.
- `tests\swp\PTTest.swp` stays as the test harness. The buttons use `PCBThermal.swp`, so tests do not touch them.
- If you move the project folder, edit each button (right-click > Customize, then click the button) and pick the new macro path.

## What changed on 2026-10-14 (re-import these three modules)
Remove the old module in `PCBThermal.swp`, import the new `.bas`, compile, save: `PTCommon`, `BuildBoard`, `SyncFromCAD`, `SetupSim`, `ReadResults`.
- Every macro reuses the workbook of the last run when it is still open in Excel, without the file dialog. To use another
  workbook, close the old one in Excel first.
- Before the first macro of the day writes to a workbook, a copy goes to `backup\<name>_yyyymmdd.xlsx` next to it.
- Build Board colours the parts (see HOWTO_COLOURS.md) and warns about RefDes that have heat but no built part.
- Sync CAD to Excel logs a failed check once; a check that is still unchecked on the Change log is not added again.
- Run Thermal Study shows a progress bar in the SolidWorks status bar.
