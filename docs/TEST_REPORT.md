# Test report: M1 to M7 on the Raspberry Pi 4 reference case

Run on 2026-09-30 / 2026-10-01 (M1 to M6 overnight, M7 in the morning), on this PC (SolidWorks 2023 SP5, Excel 16), with Claude driving SolidWorks and Excel.
Workbook used: `example/RaspberryPi4/old/pcb_thermal_survey_RPi4_test.xlsx` (a copy; the original is untouched).
All macros were imported into one project, `tests/swp/PTTest.swp`, and run from there.

## Result in one line
All 11 plan steps and M7 (RunFlow) now run. Five real bugs were found and fixed (BuildBoard, SetupSim, ReadResults, two in RunFlow).
Two plan expectations are not met because of the RPi4 data, not the code (see "Open items").
M7 ran on a test Flow project I built (assumed numbers), so its temperatures show the macro works, not that the model is validated.

## Results per step

| # | Step | Result | Notes |
|---|---|---|---|
| 0 | Import 8 modules, Debug > Compile | PASS | No compile errors on first import (all 8 modules together). |
| 1 | SetupWorkbook | PASS | 5 new tabs, 80 names, "Unchecked Change log rows: 0". Class dropdown still works (list validation present on the Class column). |
| 2 | ImportStep | PASS | 18 bodies plus PCB (85 x 56 x 1.6 mm), "Changes since the last import: none". |
| 3 | Move a part 1 mm, ImportStep, move back, ImportStep | PASS | One `moved` row each time: U1 X min 21.75 to 22.75 (delta 1), then 22.75 to 21.75 (delta -1). CAD geometry column AF showed "moved X min". I did not look at the red colour, only that the flag and the conditional format exist. |
| 4 | BuildBoard | PASS after fix 1 | PCB plus 20 parts, 0 failed, 11 connectors cut at the board surface. 1 overlap left (see open items). |
| 5 | BuildBoard again | PASS | "Nothing changed", 0 moved. This also proves parts were placed where the workbook says. |
| 6 | BuildSpreader | PASS | 8 covered chips, 18 cuts and pedestals, 8 TIM parts, 0 problems. Underside 4.1 mm = chip top + TIM 0.2 + pedestal (checked by arithmetic for every chip). I did not measure touching faces in SolidWorks. |
| 7 | Add a hole, BuildSpreader, answer Yes | PASS | Hole kept, nothing rebuilt. (My first hole landed in a cut-out and removed only 1 mm3, so I redid it inside the plate.) |
| 8 | SyncFromCAD | PARTIAL | As built block filled, Holes = 1 after step 7, volume 4744.3 mm3, mass 12.81 g. But "0 checks failed" is not met: 3 checks failed, all from the same overlap (open items). |
| 9 | Move a part box 2 mm, SyncFromCAD | PASS | 1 override read from CAD, Change log row `override from CAD`, SD card slot X min 1.5 to 3.5 (delta 2). |
| 10 | SetupSim | PASS after fix 2 | Study PT_20261001_0001 made, meshed and solved (about 6 min). Heat applied 6.728 W = register total 6.728 W. |
| 11 | ReadResults | PASS after fix 3 | Sim results row written, Thermal solution column AF filled. Energy balance 0 % (was -70.87 % before the fix). Spreader 64.3 to 72.8 C, PCB 49.9 to 79.7 C, U1 peak 76.0 C, U1 hand minus sim = -15.5 C. |
| 12 | RunFlow (M7) | PASS after fixes 4 and 5 | Connects to the Flow API from the macro host, sets 8 volume sources = 6.674 W of 6.728 W (J1, 0.055 W, has no source and is listed as a note), sets 3 temperatures, solves (Yes, about 8 min at mesh level 3, under 1 min of reporting at level 2) and reads goals (No path also tested). Writes a Sim results row `FL_<date>_<time>`, Thermal solution column AF and `CAD\RaspberryPi4_RunFlow_FL_*.txt`. Realistic run (external air 40 C, gravity, radiation on, FR4 PCB, silicon chips, mesh level 2): U1 98.8 C, U3 99.2 C, U4 to U6 97 to 98 C, spreader 51.9 to 99.1 C, PCB 47.1 to 99.2 C. L1 and L2 returned 0 K at level 2 and were skipped with a note. |

## Bugs found and fixed

1. **BuildBoard: every box failed with "the sketch was not made"** (`src/PTCommon.bas`, `PT_ToSketch`).
   Late-bound `MathUtility.CreatePoint` returned a zero point, so the box sketch was empty and SolidWorks threw it away.
   Fix: typed locals (`SldWorks.MathUtility`, `MathPoint`, `MathTransform`) in `PT_ToSketch`. `PT_PlaceAndFix` already used typed objects and works.
2. **SetupSim crashed SolidWorks** (heap corruption, exit code 0xc0000374), 4 times out of 4, right after the study was created
   (`src/SetupSim.bas`). The crashing line is `study.ThermalStudyOptions.SolverType = swsFFEPlus`.
   Fix: that line is removed; the study uses the default solver (Automatic). The steady state setting is still applied.
   If you want FFEPlus forced, that needs another way; I did not find one.
3. **ReadResults energy balance was -70.87 %** (`src/ReadResults.bas`). It compared the model's gross in and out flows, which include
   exchange with ambient (9.49 W in, 16.22 W out). The model net flow is -6.728 W, which matches the 6.728 W applied, so the physics was fine.
   Fix: balance = (register total heat - model net flow) / register total heat.
4. **RunFlow stopped with "Flow API could not load"** (`src/RunFlow.bas`). The Flow API returns Boolean results as a raw 1, not -1,
   and VBA `Not` on a raw 1 is truthy, so a successful `LoadProductAPI2` looked like a failure. It is the same for `Solve2` and `Rebuild`.
   Fix: new `IsTrue()` helper (non-zero test) used for all three; the failure message now carries the error number and text.
5. **RunFlow wrote -273.15 C to Excel for a goal that came back as 0 K** (L2, body not meshed), and **did not set the initial solid temperature**
   (only the fluid one). Fix: goals under -270 C are skipped and listed in the notes; initial conditions now set fluid, solid and radiation
   environment temperature (parameter ids 10, 26, 30).

## Open items for you

- **Overlap in the reference data.** Display Connector and the Cypress Wi-Fi cover overlap by 0.12 x 3.2 x 1.6 mm. BuildBoard lists it,
  and SyncFromCAD reports 3 failed interference checks (Wi-Fi cover, spreader, TIM U6). The plan expects 0. Fix it with an override on CAD geometry or
  Include = No on one body. I did not change your data. On the test copy I ticked Yes under Checked for those rows so SetupSim and BuildSpreader would run.
- **SyncFromCAD logs the same failed checks again on every run** (3 new rows each time), so Change log fills with duplicates. SetupSim and BuildSpreader
  stop until all are ticked. Probably worth logging a failed check only once.
- **Thermal solution tab** shows "not under plate" for chips the Spreader tab counts as covered (U2 to U6, L1, L2). That tab uses its own plate size inputs
  (100 x 80 x 2 mm), so this may just be its inputs, but I did not check.
- L1 (the small inductor) peaks at 133.5 C in the simulation. Plausible for a 0.233 W part in 2.5 x 2 mm, but worth a glance.
- The Sim results tab has 2 rows for the same study: the first is from the run before fix 3 (energy balance -70.87 %). Delete it if you like.
- **M7 setup needed once per PC and project.** Flow Simulation must be loaded in SolidWorks (Tools > Add-Ins) or the API cannot attach. The Flow project
  is made by hand (Wizard). The test project is `Project(1)PT_Flow` in `RaspberryPi4_Thermal.SLDASM`; its sources, goals and materials were added through the API.
- **M7 not covered:** L1 and L2 need a higher mesh level or a local mesh (0 K at level 2). No radiative surfaces were defined, so radiation did nothing.
  J1 has no source. Flow results were not compared with the ReadResults (Simulation) results, because the Flow test model has assumed boundary conditions.
- On the test copy I also ticked Yes on 7 Change log rows (the known overlap rows) so RunFlow would pass its check.

## Things that cost time, for next time
- From a shell, SolidWorks is reachable from VBScript (`cscript`), not PowerShell. Macros are started with `RunMacro`. See `tests/`.
- In the STEP model, moving a component with `Transform2` from the top level does not move only that part. Move it inside the
  sub-assembly's own document (`TestOps.StepMoveSub` in `tests/TestOps.bas`). Plan step 3 done by hand in SolidWorks is fine.
- SolidWorks reopens on the second monitor and the VBA editor on the first; file dialogs follow SolidWorks.

## Files changed or added
- Changed: `src/PTCommon.bas`, `src/SetupSim.bas`, `src/ReadResults.bas`, `src/RunFlow.bas`. All ASCII, compile clean, no debug code left.
  Also `docs/HOWTO_M7.md` and the M7 row in `PLAN.md`. `RunFlow` is now a module in `tests/swp/PTTest.swp`.
- Added: `tests/swp/PTTest.swp` (test project, has extra test modules inside), `tests/` (test helpers: `TestOps.bas`, VBScript runners),
  `example/RaspberryPi4/old/pcb_thermal_survey_RPi4_test.xlsx`, `example/RaspberryPi4/CAD/*` (built parts, assembly, reports, study results).
