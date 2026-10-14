# M5 SyncFromCAD: how to install and run

`SyncFromCAD.swp` = `PTCommon.bas` + `SyncFromCAD.bas`. Works for any board. It never changes SolidWorks files.

Run it after you changed something in SolidWorks (moved or resized a part box, edited the spreader).
It writes to the workbook only:
- CAD geometry override columns O..T (and PCB overrides D7:D9) where the real box differs. CAD wins.
  If you had typed a different override, the row on the Change log says `conflict` and keeps your old value.
- Spreader tab: AS BUILT block (areas, volume, mass, thickness, holes, last sync) and columns U, V
  (pedestal height and TIM thickness as found). Column W flags "differs".
- Change log rows: `override from CAD`, `conflict`, `excel-owned value differs`, `check failed`.

Acceptance (RPi4):
- Move one part box 2 mm in SolidWorks (unfix, move, fix) and run: one override written, one log row.
- Delete or suppress a pedestal and run: `check failed` for that chip.
- Add two holes to the spreader and run: As_Holes = 2, As_Area smaller.
- Right after BuildBoard and BuildSpreader with no edits: nothing to report, 0 checks failed.

## Repeated failed checks (2026-10-14)
A `check failed` row is added once. If the same check (same key and text) is still unchecked on the Change log, the next run
counts it ("N already unchecked, not added again") and adds no new row. Tick it with Yes once the cause is fixed or accepted;
if it fails again later it is logged again.
