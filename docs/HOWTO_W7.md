# W7: the whole study from the page

All seven macros now work on a project folder made by the page. They read `AB_project.json` and write
`AB_cad.json` or `AB_results.json` (AB is the project's prefix). Old workbooks still work: pick the `.xlsx` when a
macro asks.

| Button | Reads | Writes | Page tab |
|---|---|---|---|
| 2 Import STEP | settings, bodies | cad `import`, `changes` | 2 |
| 3 Build Board | `forSolidWorks` bodies and PCB | cad `build` (now with `inputsHash`), `changes` | 6 |
| 4 Build Spreader | `forSolidWorks.plates` | cad `spreaderBuild`, `changes` | 7 |
| 5 Sync CAD | `forSolidWorks` | cad `asBuilt` (boxes that differ, plates as built, failed checks), `changes` | 7, 10 |
| 6 Run Thermal Study | `forSolidWorks`, `simSetup` | results: a new run with a copy of the inputs, status `solved` or `set up` | 8 |
| 7 Run Flow Simulation | `forSolidWorks` heat and ambient, `simSetup.studyType` | results: a new `FL_` run | 8, 9 |
| 8 Read Results | the newest thermal run not read yet | the same run: temperatures, energy balance, plot, status `read` | 9 |

Sync CAD never writes the project file. Tab 10 lists the bodies whose size in SolidWorks differs and offers
**Take these** (they become typed values on tab 3).

## Update the macros (once)

1. Copy the new `.bas` files into the `macros\` (or `src\`) folder.
2. In SolidWorks press **Update Macros** on the PCB Thermal tab (or rebuild `PCBThermal.swp` by hand:
   PTCommon, PTJson and all macro modules, then Debug > Compile VBAProject).
3. User package only: run `macros\SetupToolbar.cmd` again with SolidWorks closed, so the tab gets the five new buttons.

## Test with the page (RPi4)

1. Start from the W6 test folder (Import STEP and Build Board done, `HOWTO_W6.md`). Page tab 6: the Build Board card
   should show every check ticked.
2. Tab 7: the example already has a top plate; for a new board press **Add a top plate**, tick Covered for U1 to U6,
   L1, L2, save. Every Check should be OK. Press **I pressed it** on the Build Spreader card.
3. SolidWorks: **4 Build Spreader**, Yes to the project. Yes to the rectangle the size of the PCB.
   - Expect: the plate with 6 pedestals, TIM parts, 5 clearance cuts. Within 2 s the card stops waiting, the page
     says "Build Spreader finished", and tab 7 shows the plate as built.
4. Move one part 0.5 mm by hand in the assembly. **5 Sync CAD**. Expect on tab 10 one row in Sizes changed in
   SolidWorks and a "differs in SolidWorks" change record. Press Take these, tick the change, save.
5. Tab 8: study type Thermal, Add the parts with heat, save. **6 Run Thermal Study**.
   - Expect: a study PT_date_time solved, and a new run on tab 9 with status solved.
6. **8 Read Results**. Expect on tab 9: peak per part, PASS or FAIL, energy balance within 2 %, the plot picture,
   and Tj hand vs simulation within a few degrees for the big parts.
7. Change the ambient on tab 0 and save: tab 8 and tab 9 say the run is older than the inputs (ambient 40 C, now ...).
8. Flow (optional, the macro builds the Flow project, `HOWTO_M7.md`): tab 8 study type Flow, save,
   **7 Run Flow Simulation**.

## Exports (W9)

Export in the top bar: **Report (PDF)** opens a print preview; Print, then Save as PDF. **Workbook (.xlsx)**
downloads a workbook made by the page (values only, a note on every header, the Method tab lists every formula).
**Project archive (.zip)** downloads the three project files, the macro reports and the plot pictures.
Pick the last run or all runs in the menu.

## If something fails

Send the macro name, the highlighted line and the message (compile errors), or the text of the message box, and
`%APPDATA%\PCBThermal\log.txt`.
