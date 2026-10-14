# M3 BuildBoard and M4 BuildSpreader: how to install and run

Both work on any board. Run them after SetupWorkbook and ImportStep (see HOWTO_M1_M2.md).

## Make the macros (once)
Same as before: a new .swp, then File > Import File in the VBA editor.
- `BuildBoard.swp`    = `PTCommon.bas` + `BuildBoard.bas`
- `BuildSpreader.swp` = `PTCommon.bas` + `BuildSpreader.bas`

PTCommon changed in this step (new SolidWorks helpers). Re-import it into the older macros too
(ImportStep, SetupWorkbook): remove the old PTCommon module first, then import the new file.

The two build macros use the SolidWorks type library (always there in a new SolidWorks macro).
If Compile says "User-defined type not defined" on `SldWorks.Component2`, open Tools > References
and tick "SldWorks 2023 Type Library".

## Before the first run
- Change log: every row must say Yes under Checked, or the build macros stop and tell you which rows.
- Settings tab: part and assembly templates (blank = SolidWorks defaults), prefix, output folder
  (blank = folder `CAD` next to the workbook).

## BuildBoard
Run it, pick the workbook.
- First time: creates `<prefix>Thermal.sldasm`, `<prefix>PCB.sldprt` and one `<prefix>P_<key>.sldprt`
  per included body, all fixed at their coordinates.
- Later: updates the same assembly and logs every change on the Change log tab.
- Writes `<prefix>BuildBoard_report.txt` next to the assembly: failures, skipped rows, overlapping boxes.

Acceptance (RPi4): PCB + 20 parts (18 imported + L1, L2). Run it again straight away: "Nothing changed."
Overlaps listed are only connectors going into the PCB (and anything the report explains).

## BuildSpreader (2026-10-14: you draw the outline, the macro completes the part)
Run it after BuildBoard, and after the bare board simulation has been compared with the hand calculation.

1. In SolidWorks, in the thermal assembly: Insert > Component > New Part. Put `Spreader` in its name
   (for example `RPi4_Spreader`). Pick the **Front Plane** of the assembly (or a face parallel to it, such as the PCB top face).
2. Sketch the plate outline: a rectangle or any closed shape, arcs included. Exit the sketch and **name it `Outline`**
   (a part with only one sketch works without renaming). Leave edit mode, then save the part as a file
   (right-click it in the tree, Save Part (in External File)) next to the assembly.
3. In Excel, tab `7 Spreader`: set Covered = Yes for the chips the plate must touch, Height mode, TIM mode, TIM thickness and TIM k per chip,
   plate thickness and material, clearances. Underside height mode Auto sets the plate underside from the tallest covered chip plus its TIM.
4. Run **BuildSpreader**, pick the workbook. It finds the part with `Spreader` in its name (or the one you selected in the tree) and adds:
   `PT_Under` (plane), `PT_Plate` (your Outline extruded, starting at the underside height), `Clear_*` cuts, `Ped_*` pedestals down to the TIM,
   named faces `Top`, `Underside`, `Contact_<RefDes>`, and one `<prefix>TIM_<RefDes>.sldprt` per covered chip with TIM mode Body.
   It saves the part and the assembly and writes `<prefix>BuildSpreader_report.txt`.
5. Change the outline, or anything on the Spreader tab, and run it again: the plate, pedestals and cuts follow; holes and fillets you added are kept.

Rules and messages:
- A chip that lies outside the outline gets no pedestal; the report and the Notes say so. Redraw the Outline or set Covered = No.
- A tall part outside the outline gets no clearance cut (noted).
- "virtual part": right-click it, Save Part (in External File), run again. "turned": make the part again on the Front Plane.
- No Spreader part in the assembly: the macro offers a rectangle the size of the PCB (the old behaviour); its `Outline` can be edited.
- Two spreader parts that are not suppressed: select the one to complete in the tree and run again.
- `PT_Under`, `PT_Plate`, `Clear_*`, `Ped_*` belong to the macro. Change their numbers on the Spreader tab, not in SolidWorks.

Acceptance (RPi4, all U chips covered, outline of any shape that covers them):
- The plate follows your outline, its underside is at the Spreader tab value, thickness as set.
- Every covered chip: chip top, TIM and pedestal bottom touch (Tools > Evaluate > Measure: distance 0).
- Tools > Evaluate > Interference Detection on the assembly: nothing between the spreader and any part.
- Add a hole to the spreader in SolidWorks, run BuildSpreader again: the hole is still there.

## If something fails
Send the highlighted line and message (compile errors), the text of the message box, and the report file.

## Added 2026-10-14
- BuildBoard colours the PCB green and each part with its own colour (HOWTO_COLOURS.md).
- If no row has Include = Yes, BuildBoard asks before building only the PCB.
- The summary box lists RefDes that have heat in the register but no built part (RefDes and Include = Yes on CAD geometry).
  SetupSim would give them no heat source.
- A box part whose top or bottom face cannot be named is now reported as FAILED, instead of being built without the face
  SetupSim needs (Chip_<RefDes>_Top, Bottom, PCB_Top, PCB_Bottom).
- If adding a part fails, the half-made part is closed, so a rerun is not blocked by "already open".
- The macros use one project `PCBThermal.swp` (HOWTO_TOOLBAR.md); the per-macro `.swp` files above are for testing only.

## Which plane is the board in?
The thermal assembly is built by BuildBoard in the board frame: board X and Y in the **Front Plane**, Z up through the board, whatever
way the STEP had the board. The RPi4 STEP itself lies parallel to the Top Plane (`2 STEP import`, Frame: Board Z = model +Y);
ImportStep maps it, so in the thermal assembly it stands up in the default isometric view. Use the Front view to see the board from the
component side, and draw the spreader Outline on the Front Plane (or on a face parallel to it, such as the PCB top face).
An Outline on the Top or Right Plane is refused with a message.

## Two spreaders: one on each side of the PCB
For chips on both sides (CAD geometry column Side = Top or Bottom):
- Top plate: part with `Spreader` in its name, tab `7 Spreader`. Bottom plate: part with `Spreader` and `Bottom` in its name
  (for example `RPi4_Spreader_Bottom`), tab `7b Spreader bottom`. Draw each Outline on the Front Plane (or a face parallel to it).
- On each tab set Covered = Yes only for chips of that side. The Check column says "chip is on the bottom side" or "chip is on the top side"
  when a chip is on the wrong tab. The bottom tab lists the same RefDes as the top tab.
- Heights are board Z in mm. Top plate: face toward the board = tallest covered chip top + TIM. Bottom plate: face toward the board = lowest covered chip
  bottom minus TIM, so it is below the PCB bottom face (Z below 0). Bottom plate thickness goes further down; pedestals go up to the chips.
- Clearance cuts of the bottom plate come from CAD geometry column AL (parts whose bottom reaches below the plate face plus the Z clearance).
- BuildSpreader completes both parts in one run (top first), then builds the TIM parts for both sides. TIM face names: `_Top` is the face toward the plate and
  `_Bottom` the face on the chip, on both sides.
- SyncFromCAD reads both plates (as built block of tab 7 and tab 7b). The hand network for the bottom plate is tab `8b Thermal solution bottom`, a copy of tab 8 (mark the bottom chips Yes in its column H).
- ReadResults lists both spreader files in Sim results and, with two plates, adds a top and a bottom min to max line to its notes (the Sim results spreader columns stay the overall range).
