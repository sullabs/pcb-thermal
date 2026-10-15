# Part colours (SolidWorks and Excel)

Result: every part of the thermal assembly has its own colour, the PCB is green, and the same colour shows in the workbook.

## SolidWorks
BuildBoard colours the PCB green and every built part with one of 12 colours. The colour comes from the part's row on
CAD geometry (row minus the first body row, mod 12), so a part keeps its colour on every run. The summary box says how many
parts were coloured. TIMs and the spreader keep the default colour.
The palette is `PT_PartRgb` in `PTCommon.bas`; the PCB colour is `PT_PcbRgb`.

## Excel
Conditional formats in the workbook, not cell fills, so nothing is written by a macro:
- `3 CAD geometry`, Name used (E26:E95)
- `2 STEP import`, name (C9:C88), and the PCB row C8 in green
- `1 Component register`, RefDes (B8:B53), matched to CAD geometry by RefDes

A cell is coloured only when Include = Yes on CAD geometry. Names used by the rules: `Geo_Include`, `Geo_RefDes`.

## Install in the template and the reference workbook (once)
Excel COM is blocked inside Claude Code, so run this from your own PowerShell window with the workbooks closed in Excel:

    powershell -NoProfile -ExecutionPolicy Bypass -File tools\Update-TemplateColors.ps1

It copies each file to `<name>.bak_<time>`, adds the rules, adds the two-resistor inputs (`Sim_LayerFrac`, `Sim_kCore`) if the
workbook lacks them, checks that the Class dropdown and the required names survived, and saves. Safe to run twice.
Copies of the template made before this change do not have the colours: run the script on them too.
