# M1 and M2: how to install and run

Works on any board STEP. Nothing in the macros names a part, a RefDes or a size.

## Make the two macros (once)

Every macro is two modules: the shared `PTCommon.bas` plus its own module.

1. SolidWorks: Tools > Macro > New, save as `SetupWorkbook.swp`.
2. In the VBA editor: File > Import File > `src\PTCommon.bas`, then again for `src\SetupWorkbook.bas`.
   Delete the empty `Module1` if SolidWorks made one.
3. Debug > Compile VBAProject. It should compile with no message. Save.
4. Same again for `ImportStep.swp` with `PTCommon.bas` + `ImportStep.bas`
   (if you reuse your old ImportStep.swp: remove the old ImportStep module first, then import both files).

## Run

1. **SetupWorkbook** (once per workbook, safe to run again). Pick the workbook.
   You should see: new tabs Settings, Change log, Spreader, Sim setup, Sim results (and the hidden tab
   STEP import (previous)), and "Unchecked Change log rows: 0".
   Then look at: Settings (templates, prefix) and the Spreader column "Covered" (default: low parts on the top side).
2. **ImportStep** on the same model as before. Expect: "Changes since the last import: none."
3. Test the change report: move one component 1 mm in SolidWorks, run ImportStep again.
   Expect exactly one `moved` row on the Change log tab, with Delta 1 (or -1). CAD geometry column AF
   shows it in red until you type Yes under Checked.

## Acceptance (M1)
- Workbook opens with no errors, the Class dropdown on the Component register still works.
- The new tabs calculate. On the RPi4 data: Spreader underside height about 4.1 mm, clearance cuts listed for the
  tall connectors, Log_Unchecked = 0.

## Acceptance (M2)
- Same model twice: no Change log rows.
- One part moved 1 mm: one `moved` row, right delta.

## If something fails
Send the highlighted line and the message (compile errors), or the text of the message box.

## Notes (2026-10-14)
- You are asked for the workbook only the first time, or when the last one is no longer open in Excel.
- A copy of the workbook is saved in `backup\` next to it before the first macro of each day.
- The template needs `tools\Update-TemplateColors.ps1` once for the colours and the two-resistor inputs (HOWTO_COLOURS.md).
