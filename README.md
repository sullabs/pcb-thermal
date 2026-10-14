# PCB Thermal macros for SolidWorks 2023

Turns a PCB STEP file into a thermal model, with an optional heat spreader on one or both sides, runs the
SolidWorks thermal simulation and writes the results back to Excel. You need SolidWorks 2023 (with Simulation
for the thermal study, Flow Simulation for the flow study) and Excel.

## What is in this zip

- `macros\` : 9 `.bas` files (one shared module, 8 macros)
- `icons\` : button icons (16 and 32 px)
- `PCB_thermal_TEMPLATE.xlsx` : the workbook. Make one copy per board and keep the template clean.
- `docs\` : the workflow and the spreader notes

## 1. Install (once, about 10 minutes)

1. Unzip to a folder with a short path that will not move, for example `C:\PCBThermal\`.
2. SolidWorks: Tools > Macro > New. Save it as `C:\PCBThermal\PCBThermal.swp`.
3. In the VBA editor: File > Import File (Ctrl+M). Import every file in `macros\`
   (`PTCommon.bas` plus the 8 others). Delete the empty `Module1`.
4. Debug > Compile VBAProject. Nothing should appear. Press Ctrl+S, then close the editor.

If a compile error appears, send me the highlighted line and the message.

## 2. Make the toolbar tab (optional but easier)

1. Open any part. Right-click a CommandManager tab > Add Tab > Empty Tab. Rename it `PCB Thermal`.
2. Tools > Customize > Commands > category Macro. Drag **New Macro Button** onto the tab, fill the box, press OK.
   Repeat for each row. Macro = `C:\PCBThermal\PCBThermal.swp` for all of them.
3. Right-click the tab > Copy Tab to Assembly, so it also shows in assemblies.

| # | Tooltip | Method | Icon (in `icons\`) |
|---|---|---|---|
| 1 | Prepare Workbook | `SetupWorkbook.main` | `1_PrepareWorkbook_32.bmp` |
| 2 | Import STEP | `ImportStep.main` | `2_ImportStep_32.bmp` |
| 3 | Build Board | `BuildBoard.main` | `3_BuildBoard_32.bmp` |
| 4 | Build Spreader | `BuildSpreader.main` | `4_BuildSpreader_32.bmp` |
| 5 | Sync CAD to Excel | `SyncFromCAD.main` | `5_SyncCADtoExcel_32.bmp` |
| 6 | Run Thermal Study | `SetupSim.main` | `6_RunThermalStudy_32.bmp` |
| 7 | Run Flow Simulation | `RunFlow.main` | `7_RunFlowSim_32.bmp` |
| 8 | Read Results | `ReadResults.main` | `8_ReadResults_32.bmp` |

No tab? Run any macro with Tools > Macro > Run, pick `PCBThermal.swp`, then the module and `main`.

## 3. Use

1. Copy `PCB_thermal_TEMPLATE.xlsx` to a new name, for example `MyBoard.xlsx`, and open it in Excel.
2. Fill the `1 Component register` tab (RefDes, heat, theta values). Blue cells are inputs.
3. Run the macros in order. Each one asks for the workbook the first time.
   - **Import STEP**: open the board STEP (or the open model). Writes the bodies to the CAD geometry tab.
   - Set Include = Yes for the PCB and every part with heat. Tick the Change log rows.
   - **Build Board**: builds the thermal assembly. Read `CAD\<prefix>BuildBoard_report.txt`.
   - **Run Thermal Study**, then **Read Results**: bare board temperatures go to the Sim results tab.
4. Spreader (optional, one plate per side of the board):
   - In the thermal assembly, add a part with ONE sketch named `Outline` on its plane. That is the plate shape.
     Name the part with `Spreader` in the file name. For the plate under the board also put `Bottom` in the name.
   - In the `7 Spreader` tab (and `7b Spreader bottom` for the bottom plate) mark the chips it touches and enter
     plate thickness and TIM thickness.
   - Run **Build Spreader**, then **Sync CAD to Excel** (checks the CAD against the sheet), then the study again.

Full steps: `docs\WORKFLOW_BARE_BOARD.md`, `docs\HOWTO_M3_M4.md`, `docs\HOWTO_M5.md`.

## Rules to know

- Excel units are mm. Do not edit grey cells.
- Keep the workbook open in Excel while the macros run.
- If you move `PCBThermal.swp`, edit each toolbar button and pick the new path.
- Accuracy of the numbers is not validated. Check them against your own hand calculation before relying on them.

## Updating a macro later

Open the project (Tools > Macro > Edit), remove the old module (answer No to export), import the new `.bas`,
compile, save.
