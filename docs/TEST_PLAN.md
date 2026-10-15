# Test plan: all macros in one session (M1 to M6)

Nothing here is specific to the RPi4; the expected numbers are for the RPi4 reference case.
Work on a COPY of the workbook: `pcb_thermal_survey_RPi4_test.xlsx`.

## 0. Before you start
- Excel installed on this PC.
- SolidWorks: Tools > Add-Ins > SOLIDWORKS Simulation ticked.
- The STEP is extracted: `example\RaspberryPi4\Raspberry Pi 4 Model B.STEP`.

## 1. Make the seven macros (once)
For each line: SolidWorks > Tools > Macro > New, save as the name shown, then in the VBA editor
File > Import File for both files, delete any empty `Module1`, Debug > Compile VBAProject, save.

| Macro file | Import these two files from `src\` |
|---|---|
| SetupWorkbook.swp | PTCommon.bas, SetupWorkbook.bas |
| ImportStep.swp | PTCommon.bas, ImportStep.bas |
| BuildBoard.swp | PTCommon.bas, BuildBoard.bas |
| BuildSpreader.swp | PTCommon.bas, BuildSpreader.bas |
| SyncFromCAD.swp | PTCommon.bas, SyncFromCAD.bas |
| SetupSim.swp | PTCommon.bas, SetupSim.bas |
| ReadResults.swp | PTCommon.bas, ReadResults.bas |

Your old ImportStep.swp: remove its old module first, then import both files.
Compile error "User-defined type not defined" on `SldWorks.Component2` or `SldWorks.MathTransform`:
Tools > References, tick "SldWorks 2023 Type Library".
If a compile error stops you: send the macro name, the highlighted line and the message.

## 2. Run in this order

Order note (2026-10-14): do steps 1 to 5, then steps 10 and 11 as a **bare board** run (no spreader in the assembly), and compare
with the hand calculation (section B). Steps 6 to 9 (BuildSpreader, hole, SyncFromCAD, move a part) come after that.

| # | Run | You should see |
|---|---|---|
| 1 | SetupWorkbook | New tabs Settings, Change log, Spreader, Sim setup, Sim results. "Unchecked Change log rows: 0". Class dropdown on the register still works. |
| 2 | ImportStep (the RPi4 model) | "Changes since the last import: none." |
| 3 | Move one part 1 mm in SolidWorks, ImportStep again | One `moved` row on Change log, delta 1. CAD geometry column AF red on that row. Type Yes under Checked. Move the part back, run ImportStep, check that row too. |
| 4 | BuildBoard | `CAD\<prefix>Thermal.sldasm`: PCB + 20 parts. Report lists connectors cut at the board surface. Overlaps left: ideally 0 (send the report if not). |
| 5 | BuildBoard again | "Nothing changed." |
| 6 | Spreader tab: check column Covered and Check (all OK). BuildSpreader | `Spreader_r01` + TIM parts. Measure chip top / TIM / pedestal: touching. |
| 7 | Add a hole to the spreader in SolidWorks, BuildSpreader, answer Yes | Hole kept. |
| 8 | SyncFromCAD | As built block filled (As_Holes = 1), 0 checks failed. |
| 9 | Move one part box 2 mm (unfix, move, fix), SyncFromCAD | One override on CAD geometry, `override from CAD` on Change log. Check the rows. |
| 10 | SetupSim | New study PT_yyyymmdd_hhmm, "Meshed and solved", heat applied = register total (6.73 W). |
| 11 | ReadResults | Sim results row, Thermal solution AF filled, AG = hand minus sim, energy balance within 2 %. |

## 3. What to send back
For every step that does not match: the message box text, and the report file named in it
(`CAD\*_report.txt`, `CAD\*SetupSim_*.txt`). For a VBA runtime error: the line the debugger stops on.

## R. Every body, rotation, Include switch (written 2026-10-14, not yet run)

Work on a COPY of `example\RaspberryPi4\RaspberryPi4.xlsx`. Re-import `ImportStep.bas` and `PTCommon.bas` (and `BuildBoard.bas` +
`PTCommon.bas` for BuildBoard) into their .swp projects, Debug > Compile first. Send any compile error line.

| # | Run | You should see |
|---|---|---|
| R1 | ImportStep on the RPi4 STEP | Message "Workbook prepared once" (Rot columns, setting changed to No). About 108 body rows (the 108 solids), not 19. Message that the change report was skipped (level changed). Status bar shows "measuring body n of N". |
| R2 | Look at STEP import N:P and CAD geometry AI:AK | Rot 0 for most rows. Any part turned in the board plane shows Rot, Own L, Own W; L and W stay the envelope. |
| R3 | CAD geometry column G | New rows = No. Rows you had set to Yes before (single-body parts) keep Yes. Small parts (under 1 mm) are all listed. |
| R4 | Set a few small parts and one turned part to Yes, BuildBoard | Only Yes rows are built. The turned part is a turned rectangle in SolidWorks (check Rot against a measure). Overlap list has no false pairs from touching turned parts. |
| R5 | ImportStep again, same model | "Changes since the last import: none." Same row count. |
| R6 | Turn the PCB part 30 degrees in the STEP (or open a rotated copy), ImportStep | Message "board is turned ... degrees"; board X/Y follow the board edges; L x W of the PCB unchanged. |

If R1 shows "gave no surface points" for many bodies, send the message: the tessellation call needs another route.

## F. Fresh run, 2026-10-14 (RPi4 STEP, `example\RaspberryPi4\RPi4_fresh.xlsx`, macros in `tests\swp\PTFresh3.swp`)

| Step | Result |
|---|---|
| Register filled through Excel COM (9 made-up rows, 5.92 W; L1, L2 removed later: the STEP has no inductor bodies) | pass |
| ImportStep (R1) | pass: 108 body rows plus PCB, 2 turned bodies |
| Re-import (R5) | pass after fix: "Changes since the last import: none" (first try gave 108 false "renamed", see PLAN.md section 9) |
| Include = Yes on 7 bodies, RefDes linked | pass: all Pd placed, 0 rows without a body |
| BuildBoard, twice | pass: PCB + 7 parts, 0 overlaps, second run "Nothing changed", parts coloured (PCB green) after fix |
| R2, R3 (small parts), R4 (turned part), R6 (turned board) | not run |

## B. Next step: bare board simulation against the hand calculation (before any spreader)

Use the fresh copy. Follow `WORKFLOW_BARE_BOARD.md`, stages A, D to G. The assembly must have no spreader and no TIM.

| # | Run | You should see |
|---|---|---|
| B1 | `6 Bare board network` and `1 Component register` | Hand Tj per RefDes with no heatsink. Write them down. |
| B2 | `9 Sim setup`: Sim_Type Thermal, ambient = register D5, h and emissivity same as the hand network, interface resistance 0, chip model Block | Values set, nothing pink. |
| B3 | Tick every row on `11 Change log` (unchecked count 0). SyncFromCAD if you moved anything. | "Unchecked Change log rows: 0". |
| B4 | SetupSim | Report says "Bare board study", heat applied = register total (5.70 W with L1 and L2 removed), nothing missing, study solved. |
| B5 | ReadResults | New row on `10 Sim results`, Thermal solution column AF filled, energy balance within 2 %. If a healthy run shows a large imbalance, go back to net flow against `Reg_TotalQ` (WORKFLOW_BARE_BOARD.md, "Needs a test run"). |
| B6 | `6 Bare board network` (leave R7 blank for the newest run) | Hand minus sim per chip. Accept when within 10 % of the rise above ambient or 5 C. Record the accepted run name on `10 Sim results`. |

Only after B6 is accepted: steps 6 to 9 of section 2 (BuildSpreader and later).

## B result, 2026-10-14 (fresh RPi4 copy, study PT_20261014_1528, macros in `tests\swp\PTFresh4.swp` and `PTFresh5.swp`)

| # | Result |
|---|---|
| B4 SetupSim | pass: "Bare board study (no spreader)", meshed and solved in about 2 minutes, heat applied 5.696 W = register total, 0 problems. Note: U6 has no Rjc, so the generic k is used |
| B5 ReadResults | pass after fix: energy balance 0 % (net -5.695 W against 5.696 W). The old formula (gross in against gross out) gave 49.16 %. PCB 62.4 to 94.0 C. Sim results has two rows for this study (row 8 is from before the fix, 49.16 %) |
| B6 hand against sim | Hand is cooler than sim on every chip. Hand minus sim: U1 -10.6, U2 -3.8, U3 -10.5, U4 -9.4, U5 -6.6, U6 -3.3 C (mean -7.4). Tab 6 flags U1 FAIL, U2 and U6 Tight, rest OK on margin. Tolerance (10 % of rise or 5 C) is met by U2 and U6 only. Not accepted yet |

Reading the gap: sign is the same on all chips, so look at the board side first (PCB k, convection and radiation over the whole faces, board temperature). The sim PCB peak is 94 C against a hand board average of 79.6 C.

## S. Spreader drawn by the user (written 2026-10-14, after the bare board baseline)

Re-import `PTCommon.bas` and `BuildSpreader.bas` (project with PTCommon + BuildSpreader, Debug > Compile first). Work on the fresh copy.

| # | Run | You should see |
|---|---|---|
| S1 | In the thermal assembly: Insert > Component > New Part, name it `RPi4_Spreader`, pick the Front Plane, sketch an outline with a chamfered or round corner that covers the chips, exit, name it `Outline`, save as a file. In Excel: Covered = Yes for the chips, TIM data | Nothing is built yet. |
| S2 | BuildSpreader | "Spreader completed from your Outline sketch". Plate in the shape of your sketch at the underside height (Spreader tab, 4.1 mm for all six chips), pedestals and TIM parts for each chip, 0 problems. |
| S3 | Measure chip top, TIM, pedestal bottom | Touching (distance 0). Interference Detection: none. |
| S4 | Edit the Outline (move a corner or change an arc), BuildSpreader again | "Spreader updated". Plate follows the new outline. |
| S5 | Make the outline miss one chip | Notes say that chip has no pedestal because it is outside the outline. |
| S6 | Add a hole in SolidWorks, BuildSpreader | The hole is kept. |
| S7 | Make the part on the PCB top face instead of the Front Plane (part sits 1.6 mm above the origin) | Same result: heights and positions still match the Spreader tab (the macro reads the part position). |
| S8 | Delete the spreader part, run BuildSpreader | Offers a PCB size rectangle (old behaviour). |

Then SetupSim (study with the spreader) and ReadResults; compare with the bare board baseline.

## S2. Two spreaders (written 2026-10-14, not run)
The RPi4 has no heat chip on the bottom side, so use a stand-in: in CAD geometry give one bottom-side body (Side = Bottom, for example the SD card slot,
Zmin -1.45) a RefDes such as TST with Include = Yes, type TST in the top tab chip list (B26), set Covered = Yes on `7b Spreader bottom`, TIM 0.2. Undo it afterwards.

| # | Run | You should see |
|---|---|---|
| T1 | Excel only | Bottom tab: face Z = chip Zmin, face height used = Zmin - TIM, Check OK. Top tab Check for TST: "chip is on the bottom side". Done 2026-10-14 on RPi4_fresh: works. |
| T2 | Make a part `RPi4_Spreader_Bottom` (Front Plane, outline sketch over the chip, named Outline) next to the top part. BuildSpreader | Both parts completed. Bottom plate below the board: Z from face height down by the thickness. Pedestal up to the chip bottom, TIM box between chip bottom and plate. 0 problems. |
| T3 | Measure chip bottom, TIM, pedestal end | Touching. Interference Detection: none. |
| T4 | Change the bottom plate thickness on its tab, run again | Only the bottom plate changes. The top plate and its TIMs are untouched. |

## S and S2 results, 2026-10-14 (RPi4_fresh, macros in `tests\swp\PTFresh13.swp` / `PTFresh14.swp`, helpers in `tests\`)

| # | Result |
|---|---|
| S1, S2 | pass: a 5-line chamfered outline became the plate (Z 4.1 to 5.6), 5 pedestals, 6 TIM parts, 0 problems, 1 note (U1 has no pedestal) |
| S3 | pass by measurement (`tests\GeoProbe.bas`): every TIM touches the chip top and the pedestal bottom or plate underside, for all six chips (for example TIM U1 Z 3.9 to 4.1, TIM U2 Z 2.4 to 2.6). Interference Detection not run |
| S4 | pass: plate thickness 1.5 to 2.0 in Excel, run again: plate Z 4.1 to 6.1, 1 Change log row; back to 1.5 later: 1 rebuilt |
| S5 | pass (seen by accident): chips outside the outline got no pedestal and a note |
| S6 | pass after two fixes (below): a hole cut by hand is kept |
| S7 | pass: part placed 1.6 mm above the origin, 5 pedestals and the plate rebuilt in the new place, all absolute Z values unchanged |
| S8 | not run (rectangle fallback is the old tested path) |
| T2, T3 | pass: bottom plate (outline over two stand-in chips): face Z -1.65, plate Z -3.15 to -1.65, pedestal for the second chip up to its TIM, TIM Z -1.65 to -1.45 touching plate and chip. TIM for a 0.34 mm wide stand-in could not be sketched (SolidWorks, not the bottom logic) |
| T4 | pass in effect: with no bottom chip covered the bottom part is left alone (note), the stand-in TIM was suppressed, the top plate was untouched |

Bugs found by running it (all fixed in `src\BuildSpreader.bas`):
- `PlateStartsByOffset` called `AccessSelections` on the plate feature and never released it. That left the spreader part rolled back to before the plate,
  so on every update run the part had no body, face naming failed (7 problems) and the saved file kept the rollback bar. It now only reads the definition.
  The macro also moves the rollback bar to the end before it edits.
- A part reached by `GetModelDoc2` only as a component can return no bodies: the macro opens the part in its own window if it has none after activation.
- `AddComponent5` puts a new part with its bounding box centred near the origin, not its own origin. The macro already reads the part position, so it worked (one
  pedestal only, the others fell outside the shifted outline), and `CheckOrigin` now rebuilds pedestals and cuts when the part moves.
- Bottom side with no covered chip no longer aborts the macro.

## M6 with the spreader, 2026-10-14 (RPi4_fresh, top plate from the user outline, `tests\swp\PTFresh15.swp`)

| Step | Result |
|---|---|
| SetupSim | pass: study PT_20261014_1724 made, meshed and solved (about 3 minutes). Heat applied 5.696 W = register total. 0 problems. No "bare board" line, so the spreader and its 6 TIM parts were seen. Compile and first run clean after the SetupSim change for two plates |
| ReadResults | pass: Sim results row 10, energy balance 0 % (net flow), spreader 62.3 to 65.3 C, PCB 52.7 to 69.6 C |
| Spreader against bare board (same heat, same h) | U1 102.9 to 70.2 C, U3 95.6 to 71.2 C, U4 94.3 to 71.1 C, U5 90.9 to 68.3 C, U6 86.3 to 66.9 C, U2 84.7 to 65.4 C. PCB peak 94.0 to 69.6 C. Accuracy is not judged, all inputs are assumed |
| SetupSim with a bottom plate | NOT run. The code picks the bottom spreader by name, uses the bottom tab's k and TIM data and the chip Bottom face, compiles, but there was no bottom heat chip to test it with |
| ReadResults with two plates | not changed: spreader max and min are over both plates together |

## Two-sided board, 2026-10-14 (RPi4_fresh plus a manual bottom chip U7, `tests\swp\PTFresh16.swp`)

Setup (all in the workbook, no new STEP): U7 = LPDDR4 12 x 12 x 1.2 mm on a CAD geometry manual row (176) with Side typed as Bottom, X 40, Y 8, Z -1.2, 0.99 W
(register row 15, Sim setup row 35, tab 7 chip list B26, tab 7b Covered = Yes with TIM 0.2 mm, k 3).

| Step | Result |
|---|---|
| BuildBoard | pass: U7 part added, 0 overlaps, 0 failed |
| Bottom spreader part (outline over U7, saved as `RPi4_fresh_Spreader_Bottom2`), BuildSpreader | pass: top plate updated, bottom plate completed at face Z -1.4, plate Z -2.9 to -1.4, TIM U7 Z -1.4 to -1.2 touching chip (Z -1.2 to 0) and plate, 0 problems. No pedestal (single chip, face equals chip face) |
| SetupSim | pass: study PT_20261014_1750, solved, heat 6.686 W = register total, 0 problems. First run of the bottom-plate contact and material code |
| ReadResults | pass: energy balance 0 %, U7 peak 65.5 C, spreader (both plates together) 60.9 to 66.2 C, PCB 54.4 to 70.7 C. Top chips about 1 C warmer than the top-plate-only run because of the extra 0.99 W |

Fixed later the same day: ReadResults lists both spreader files and reports each plate separately (top 63.1 to 66.2 C, bottom 60.9 to 61.3 C on this run).
`tests\GeoProbe.bas` fails on a suppressed component (it has no box); module `Module1.Probe4` inside `PTFresh16.swp` is the fixed version.

### SyncFromCAD and hand tab, two plates, 2026-10-14

SyncFromCAD (reworked to loop over the top and the bottom spreader part) on the two-sided RPi4_fresh: pass, 0 overrides, 0 differences, 0 failed checks.
Tab 7: outline 4378.4 mm2, volume 7105.3 mm3, mass 19.18 g, 1 hole. Tab 7b: outline 888 mm2, volume 1332 mm3, mass 3.6 g, thickness 1.5, 0 holes, U7 pedestal 0, TIM 0.2 as built.
Bottom plate checks used: face height 1.4 below the chip face, TIM top on the chip and TIM bottom on the plate, outer face facing down.
Tab 8b Thermal solution bottom (copy of tab 8, Sim Tj linked to tab 8): plate 30 x 24 x 1.5, U7 marked Yes, hand Tj 99.8 C vs sim 65.5 C. The gap is the hand model having no heat path into the PCB; not judged.
`GeoProbe.bas` was removed from `PTFresh16.swp` because it no longer compiled.

### PTCommon tab matching, 2026-10-15
`PT_Sheet` now ignores number style, spaces and case (new helper `PT_TabBase`). First version broke tabs like "7b Spreader bottom" and "8b Thermal solution bottom"; fixed by allowing one letter after the step number. Test module `tests\TabTests.bas`: 25 checks (14 on `PT_TabBase`, 11 on `PT_Sheet` in a scratch workbook with the real tab names), 0 failures. SyncFromCAD re-run on the two-plate board: 0 overrides, 0 differences, 0 failed checks.
Same day, with the fixed PTCommon on the two-plate RPi4_fresh board: BuildBoard "Nothing changed", 0 overlaps. BuildSpreader both plates unchanged (top Z 4.1 mm, bottom Z -1.4 mm, 0 added). SetupSim study PT_20261015_1707 meshed and solved, heat applied 6.686 W = register 6.686 W. ReadResults: top 63.1 to 66.2 C, bottom 60.9 to 61.3 C, imbalance 0 %. All match the 2026-10-14 run. ImportStep, SetupWorkbook and RunFlow are not in the test project and were not re-run.

## Notes moved from WORKFLOW_BARE_BOARD.md (2026-10-15)

## Risks to keep in mind

- Import and build were run on the Raspberry Pi 4 on 2026-10-01 (TEST_REPORT.md) and again on 2026-10-14 on a fresh copy (`TEST_PLAN.md` section F):
  compiled, run, colours working, re-import and re-build stable. Rotation checks R2, R4, R6 are still open.
- Convection on PCB top and bottom covers the whole face, also under parts (small error).
- A heat part left at Include = No gets no heat. The total applied heat check catches it.
- SetupSim and ReadResults passed on 2026-10-01. The M8 two-resistor layers have never been run.

## Needs a test run (2026-10-14)
- Energy balance: resolved 2026-10-14. The in-against-out formula gave 49 % on a healthy run. `ReadResults` now compares the net flow
  out of the model with the register total (5.695 W against 5.696 W, 0 %) and lists the raw flows (5.89 W in, 11.59 W out) in the message.
- Colours (HOWTO_COLOURS.md): confirmed 2026-10-14 after the `PT_ApplyColor` fix (the old call stored nothing).
- The Raspberry Pi 4 overlap (Display Connector and Wi-Fi cover, 0.12 mm) still makes SyncFromCAD report failed checks. Decide on an
  override or Include = No for one body in `RaspberryPi4.xlsx`; it is not changed here.
- Include on CAD geometry applies to parts. The PCB size comes from the Used column (Length X, Width Y, Thickness Z) and has no Include.
