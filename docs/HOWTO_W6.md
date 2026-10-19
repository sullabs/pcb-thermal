# W6: Import STEP and Build Board with the page

Import STEP and Build Board now work on a project folder made by the page (`webapp/dist/PCB_Thermal.html`).
They read `AB_project.json` and write `AB_cad.json` (AB is the prefix the page gave the project; old folders without a prefix work too). They still work on an old workbook too: you pick which at the start.
The other macros (Build Spreader, Sync CAD, Setup Sim, Read Results, Run Flow) still use the workbook until W7.

## Make the two macros (once)

Each one is now three modules. SolidWorks: Tools > Macro > Edit (or New), then in the VBA editor remove the old
modules and File > Import File for each file from `src\`:

| Macro file | Import these files from `src\` |
|---|---|
| ImportStep.swp | PTCommon.bas, PTJson.bas, ImportStep.bas |
| BuildBoard.swp | PTCommon.bas, PTJson.bas, BuildBoard.bas |

Delete any empty `Module1`, then Debug > Compile VBAProject and save. The other macros can keep their PTCommon
(the changes are additions only).

## Test with the page

1. Make an empty folder, for example `Desktop\RPi4_page\`, and copy `example\RaspberryPi4\Raspberry Pi 4 Model B.STEP` into it.
2. Open `webapp\dist\PCB_Thermal.html` in Chrome or Edge. Tab 0: choose that folder, then "Start a new project here".
   Type the board name and the STEP file name (`Raspberry Pi 4 Model B.STEP`).
3. Tab 1: add the parts (Paste a parts list, or type them). Save.
4. SolidWorks: run **ImportStep**.
   - "Work on this project?" No, then pick `AB_project.json` in the folder. Next time Yes is enough.
   - "Use the model open in SolidWorks?" No, then pick the STEP file (the picker starts there).
   - Expect: "Written to ...\cad.json", 18 bodies plus the PCB (85 x 56 x 1.6 mm). Within 2 s the page says
     "cad.json reloaded" and tab 2 shows the board drawing.
5. Page, tab 3: include and link the bodies (the page suggests links by name). Save.
6. SolidWorks: run **BuildBoard**, Yes to the same project.
   - Expect: assembly `CAD\<board name>_Thermal.sldasm` in the folder, one box per included body, PCB green.
   - The page shows the build on tab 6 within 2 s.
7. Change report: move one component 1 mm in the STEP model (or in the open model), run ImportStep again.
   - Expect one "moved" record on tab 10 with delta 1.
   - BuildBoard now refuses and lists it. Tick it on tab 10, save, run BuildBoard: that part moves.

Faster start: on tab 0 use "Make it from an old workbook" with `example\RaspberryPi4\RaspberryPi4.xlsx`. It fills
every tab and writes a `cad.json` from the workbook. ImportStep then compares the new import with that one.

## What the macros check first

- Build Board refuses when `cad.json` has change records not ticked on tab 10.
- Build Board refuses when the page has not saved since the last Import STEP (it builds from what the page worked
  out on Save, `forSolidWorks`). Press Save in the page and run again.
- Neither macro writes the project file. New bodies start not included; the page keeps your choices by body key.
  A body whose name changed keeps its old key when its position and size are the same.

## Files

| File | Written by | Holds |
|---|---|---|
| `cad.json` | ImportStep, BuildBoard | import, importPrevious, build, changes |
| `backup\cad_yyyymmdd.json` | the macros | the file as it was before the first write of the day |
| `CAD\<board>_BuildBoard_report.txt` | BuildBoard | failures, skipped bodies, cuts at the board, overlaps |
| `%APPDATA%\PCBThermal\log.txt` | every macro | which folder was used, every message |

## If something fails

Send the macro name, the highlighted line and the message (compile errors), or the text of the message box.
