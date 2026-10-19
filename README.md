# PCB Thermal

Takes a PCB from its STEP file to a thermal model in SolidWorks. You fill in the parts and their power on a web page;
SolidWorks macros import the STEP and build the thermal assembly.

You need: Chrome or Edge, SolidWorks 2023. Nothing to install besides the toolbar step below.

## What is here

| File | What it is |
|---|---|
| `PCB_Thermal.html` | The page. Open it from disk in Chrome or Edge. Everything is inside this one file. |
| `macros\PCBThermal.swp` | The SolidWorks macros, ready to run: Import STEP, Build Board, Update Macros. |
| `macros\SetupToolbar.cmd` | Makes the **PCB Thermal** tab in SolidWorks with one button per macro. |
| `macros\*.bas`, `macros\icons\` | The macro source and the button icons (used by the two files above). |

## Install (once)

1. Put this folder somewhere it will stay, with a short path, for example `C:\PCBThermal\`.
2. Close SolidWorks (check Task Manager: no `SLDWORKS.exe`).
3. Double-click `macros\SetupToolbar.cmd`. It says "Made the PCB Thermal tab with 3 buttons".
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

- **Import STEP** (tab 2 on the page): pick `AB_project.json` in the board folder. It reads the STEP and writes
  `AB_cad.json`. The page reloads it on its own.
- **Build Board** (tab 6): builds the thermal assembly in the `CAD` folder and writes the build to `AB_cad.json`.

Each macro asks which project the first time, then offers the last one.

## Saving

The page saves `AB_project.json` on Save (Ctrl+S) and on its own every 30 s. Before the first save of each day it
copies the old file to `backup\`. The macros write only `AB_cad.json` and `AB_results.json`.

## Not in this package yet

Spreader, simulation and results macros (tabs 7 to 9) still work from the old workbook only. They come here once they
read the project files.

## Updating

New `PCB_Thermal.html`: replace the file. New `.bas` files: put them in `macros\`, then press **Update Macros** on the
PCB Thermal tab.

## For developers

This branch is the user package. Development happens on `addin-mvp`. Every change there to the page or to a macro
in this package must also land here, in the same session:

- Page: `npm run build` in `webapp\`, copy `webapp\dist\PCB_Thermal.html` here.
- Macros: copy the changed `.bas` files to `macros\`, press **Update Macros** (or rebuild `PCBThermal.swp`), commit
  the `.swp` too.
- A new macro that works with the project files: add its `.bas`, button and icons here and to `README.md`.
