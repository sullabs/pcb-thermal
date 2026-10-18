# PCB thermal automation: rules for working in this repo

Read `PLAN.md` first. It is the spec. This file is the short list of rules that apply to every change.

## Stack (fixed, do not change)

- SolidWorks 2023 VBA macros, delivered as `.bas` files in `src/`. The user imports each into a `.swp`
  (Tools > Macro > New, then File > Import File in the VBA editor). A `.swp` is binary; never try to write one.
- Excel template `workbook/PCB_thermal_TEMPLATE.xlsx` (plain `.xlsx`, no macros inside it). One copy per board.
  Reference case: `example/RaspberryPi4/RaspberryPi4.xlsx`. Tabs are numbered ("7 Spreader"); macros find them with `PT_Sheet`
  by the words after the number, never by the full name.
- No Python, no build123d, no external tools on the user's PC. Only SolidWorks, SolidWorks Simulation,
  SolidWorks Flow Simulation and Excel.
- Macros talk to Excel by COM late binding (`GetObject(, "Excel.Application")`, `As Object`). No project references
  beyond the SolidWorks defaults.
- **Direction since 2026-10-18: a web page replaces the workbook as the user interface.** Spec: `docs/WEBAPP_DESIGN.md`
  (milestones W1 to W9). The page is one HTML file, `webapp/dist/PCB_Thermal.html`, opened from disk in Chrome or Edge,
  with every library inside it. It writes only `project.json`; the macros write only `cad.json` and `results.json`. Schemas:
  `webapp/schema/`. The workbook and the current macros keep working until W7.
- Page source is `webapp/src/page/` (Preact and htm, no JSX). `npm run build` in `webapp/` makes `dist/PCB_Thermal.html`;
  commit the rebuilt file with every page change. `npm run serve` serves the repo on localhost for browser checks.
- Node tools in `webapp/` (`npm test`, converter, validator) are for development here only. The user's PC never needs
  Node, npm or Python: only the built HTML file ships.

## Web app rules

- Every field in the three JSON files is in `webapp/schema/`. Change the schema first, then the code, then rerun
  `npm test` in `webapp/` (it also checks that `example/RaspberryPi4/*.json` are up to date; `node tools/test_w1.mjs --write`
  refreshes them).
- Fixed units in the files: mm, C, W, W/(m K), C/W. Blank values are left out, never `null`.
- The page reads the RPi4 example as its reference case. Keep board-specific data out of `webapp/src/`, as for `.bas` files.
- Every formula lives once, in `webapp/src/calc/` (pure functions on the three files). Port a template formula exactly;
  `npm test` compares it with `tests/webapp/expected_*.json` (Excel's saved values). After a workbook formula change, open
  and save the example workbook in Excel, then `npm run expected`.
- Dark theme by default, light on request. Every value has a hover explanation from the term catalogue
  (`webapp/src/terms/catalogue.js`). A new value on the page needs: a catalogue term, a getter in `VALUES` and, if
  calculated, a formula in `FORMULAS` (`webapp/src/calc/formulas.js`); show it with `<V>` and label it with `<L>`.
  `npm test` (W4) fails on a number without a term or a formula that does not give the calculated value.
- Learn mode texts live in `webapp/src/terms/learn.js` (one entry per step tab). Values inside a sentence use `<V nofx>`.
- The example study inside the page comes from `example/RaspberryPi4/*.json` at build time (`virtual:example`), so
  board data stays out of `webapp/src/`. Rebuild after those files change.

## VBA rules (each one cost a failed run already)

- Never use a VBA reserved word or built-in function name as a variable: `Imp`, `Eqv`, `Mid`, `Len`, `Name`,
  `Date`, `Error`, `Left`, `Right`, `Val`, `Str`, `Type`, `Get`, `Set`, `Open`, `Close`. Run a check over every
  `Dim`, `Const` and parameter name before handing a file over.
- Pass objects and scalars `ByVal` in every Sub/Function signature. Array elements of a Variant array cannot be
  passed to a `ByRef ... As Object` parameter ("ByRef argument type mismatch"). Arrays stay ByRef (VBA requires it).
- No duplicate `Dim` names inside one procedure.
- ASCII only in `.bas` files (the user's Windows locale is Arabic; non-ASCII gets mangled on import).
- First line of each `.bas` is `Attribute VB_Name = "<Module>"`. Entry point is `Sub main()`.
- Wrap every block that writes to Excel in an error handler that restores `ScreenUpdating = True`.
- The SolidWorks API works in metres; the workbook works in mm. Convert at the boundary only.
- Read numbers from cells with `IsNumeric` + `CDbl`, never `Val(CStr(...))` (locale).
- The user compiles with Debug > Compile VBAProject before running. Aim for zero compile errors on first import:
  re-read every file for the rules above before delivering it.

## Workbook rules

- The workbook uses Excel-only features (x14 data validation on the Class column, conditional formats).
  If you ever edit the `.xlsx` outside Excel, do it at the XML level. **Never re-save it with openpyxl**: that
  silently deletes the Class dropdown. Preferred: change the workbook through Excel itself (COM from PowerShell
  on the user's PC, or ask the user), or by direct XML edits with `fullCalcOnLoad="1"` set.
- Every header cell has a note (cell comment). Add one for every new header.
- Every cell a macro reads or writes gets a workbook **defined name** (see PLAN.md, "Named cells"). Macros use
  names, not addresses, so layout changes do not break them. Migrate `ImportStep` to names when you add them.
- Input cells are blue (style of `'1 Component register'!B8`), calculated cells grey. Part sizes live only on CAD geometry. Macros write only to cells
  the plan says they own.
- Keep formulas to Excel 2007-era functions (`INDEX`, `MATCH`, `SUMIFS`, `COUNTIFS`, `SUMPRODUCT`, `IFERROR`).
  No `XLOOKUP`, `FILTER`, `UNIQUE`, `LET`, `LAMBDA`.
- Do not change the existing register, Triage, Rail budget, Bare board network, Thermal solution, worked tabs,
  Q methods, Datasheet checklist or Materials formulas unless PLAN.md says so.

## Testing

- Keep the example out of the template and the `.bas` files: no board names, RefDes or numbers from it. Everything
  board-specific lives in `example/RaspberryPi4/` (STEP, workbooks, `CAD/` outputs).
- The reference case is the Raspberry Pi 4 STEP (`example/RaspberryPi4/Raspberry Pi 4 Model B.STEP`, SolidWorks 2018 export,
  19 bodies) and the data already in the workbook. Every milestone in PLAN.md has acceptance checks against it.
- If Claude Code runs on the user's Windows PC, SolidWorks and Excel can be driven from PowerShell via COM
  (`New-Object -ComObject SldWorks.Application`, `Excel.Application`) to test macro logic. Ask before doing so.
- Otherwise the user runs the macro and reports: compile errors (highlighted line + message) and the summary box.

## Writing style for anything the user reads (messages, sheet notes, MsgBox text)

Short and plain. No em or en dashes. Lead with the result. Say where data was written.
