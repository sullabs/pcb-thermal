# PCB Thermal

Takes a PCB from its STEP file to a thermal model in SolidWorks. You fill in the parts and their power on a web page;
SolidWorks macros import the STEP, build the thermal assembly and the spreader, and run the thermal study.

You need: Chrome or Edge, SolidWorks 2023 (with Simulation for the thermal study, Flow Simulation only for flow
studies). Nothing to install besides the toolbar step below.

## What is here

| File | What it is |
|---|---|
| `PCB_Thermal.html` | The page. Open it from disk in Chrome or Edge. Everything is inside this one file. |
| `macros\PCBThermal.swp` | The SolidWorks macros: Import STEP, Build Board, Build Spreader, Sync CAD, Run Thermal Study, Run Flow Simulation, Read Results, Build Solution and Send Solution (for the Ansys page, `ansys\README.md`), Update Macros. |
| `macros\SetupToolbar.cmd` | Makes the **PCB Thermal** tab in SolidWorks with one button per macro. |
| `macros\*.bas`, `macros\icons\` | The macro source and the button icons (used by the two files above). |
| `macros\PCBThermal_Flow.fwp` | The Flow Simulation template Run Flow Simulation builds its project from. |

## Install (once)

1. Put this folder somewhere it will stay, with a short path, for example `C:\PCBThermal\`.
2. Close SolidWorks (check Task Manager: no `SLDWORKS.exe`).
3. Double-click `macros\SetupToolbar.cmd`. It says "Made the PCB Thermal tab with 8 buttons".
4. Start SolidWorks and open a part or assembly. The **PCB Thermal** tab is there.

If you move the folder later, run `SetupToolbar.cmd` again.

## Start a new board

1. Make a folder for the board and put its STEP file in it.
2. Open `PCB_Thermal.html` in Chrome or Edge. On tab 0 press **Choose the project folder** and pick that folder.
3. Type your prefix (your initials, for example `AB`) and the board name. Press **Start a new project here**.
   The page writes `AB_project.json`, `AB_cad.json`, `AB_results.json` and a `CAD` folder. The prefix keeps file
   names unique.
4. Follow the tabs in order. A grey tab says what it waits for. **Learn** in the top bar explains each step.

Have an old workbook (`.xlsx`) for this board? Use **Make it from an old workbook** in step 3 instead.

Want to look around first? Tab 0: **Open the example study** (Raspberry Pi 4, read only).

## In SolidWorks

Each SolidWorks step has a card on its page tab: the checks, which button to press, and what to look at after.
Press **I pressed it** on the card and the page waits until the macro has written its file.

| Button | Page tab | What it does |
|---|---|---|
| Import STEP | 2 | Reads the STEP and writes the bodies to `AB_cad.json`. |
| Build Board | 6 | Builds the thermal assembly in the `CAD` folder. |
| Build Spreader | 7 | Builds the spreader plate, pedestals and TIM parts. |
| Sync CAD | 7, 10 | Reads the assembly back; tab 10 offers sizes you changed in SolidWorks. |
| Run Thermal Study | 8 | Makes and solves a new thermal study; the run goes to `AB_results.json`. |
| Run Flow Simulation | 8 | Builds the Flow project from `PCBThermal_Flow.fwp` with the mesh, convergence, environment, materials, surfaces and contacts of tab 8, solves it, reads its goals and the run health (tab 9). |
| Read Results | 9 | Reads the solved study: temperatures per part, energy balance, plot picture. |
| Build Solution | Ansys page, step 4 | Makes the solution assembly from the board STEP and builds or updates the spreader plates and TIM parts. |
| Send Solution | Ansys page, step 4 | Sends the finished plates (STEP and measures) to the board folder, for the hand check and AEDT. |

Flow materials: Flow takes a k or an emissivity only as a named item of its Engineering Database. Tab 8 (study type
Flow) lists the user materials (`k ...`) and radiative surfaces (`e ...`) a project needs; make each once under
Solids (or Radiative Surfaces) > User Defined > PCBThermal. A missing one falls back to Flow's generic item and the
report says so.

Each macro asks which project the first time (pick `AB_project.json`), then offers the last one. The macros write
only `AB_cad.json` and `AB_results.json`; the page reloads them on its own.

## Help and exports

**Help** in the top bar: heat basics, reading a datasheet, power methods, the hand checks, spreaders, simulation,
every tab and every button, troubleshooting and a searchable glossary.

**Export** in the top bar: a report (print preview, then Save as PDF), a workbook (`.xlsx`, values with a note on
every header and every formula on the Method tab) and a project archive (`.zip`). Pick the last run or all runs.

## Saving

The page saves `AB_project.json` on Save (Ctrl+S) and on its own every 30 s. Before the first save of each day it
copies the old file to `backup\`. The macros write only `AB_cad.json` and `AB_results.json`.

## Updating

New `PCB_Thermal.html`: replace the file. New `.bas` files: put them in `macros\`, then press **Update Macros** on the
PCB Thermal tab, or run `macros\BuildMacros.cmd` (no clicks: it starts SolidWorks if needed, imports, compiles,
saves, and prints the result; exit code 0 = OK). New buttons: close SolidWorks and run `SetupToolbar.cmd` again.

## For developers

This branch is the user package. Development happens on `addin-mvp`. Every change there to the page or to a macro
in this package must also land here, in the same session:

- Page: `npm run build` in `webapp\`, copy `webapp\dist\PCB_Thermal.html` here.
- Macros: copy the changed `.bas` files to `macros\`, run `macros\BuildMacros.cmd -nopause` (or press **Update Macros**),
  commit the `.swp` too. BuildMacros.cmd reaches SolidWorks through COM (`BuildMacros.vbs`), so it needs the desktop
  session and a SolidWorks started normally (one started by COM is not found by the next script). A compile error
  shows the VBA editor's message; after OK the log `%TEMP%\PCBThermal_build.log` names the module and line.
- A new macro that works with the project files: add its `.bas`, button and icons here and to `README.md`.
