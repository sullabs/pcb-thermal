# PCB thermal automation: plan

Goal: learn and run PCB thermal analysis end to end on one workbook and one SolidWorks model:
component heat (Excel) -> geometry from STEP -> hand resistance network (Excel) -> thermal model built
by macros (SolidWorks) -> simulation set up and read back by macros -> compare, edit, repeat.
Data accuracy is not the point; correct files and a correct process are.

Reference board: Raspberry Pi 4 Model B (`Raspberry Pi 4 Model B.STEP`).

---

## 1. Status

| Item | State |
|---|---|
| Component register filled (12 heat rows, 6.73 W) | done |
| `CAD geometry` tab (imported / override / used, checks, PCB stack k) | done |
| `STEP import` tab (settings M3:M5, messages column L) | done |
| `ImportStep` macro (SolidWorks -> Excel) | done, tested by user on SW 2023 |
| M1 `SetupWorkbook` macro (adds tabs 4.1 to 4.8 and names) | written, not yet run |
| M2 `ImportStep` change report, names, keys | written, not yet run |
| M3 `BuildBoard` | written, not yet run |
| M4 `BuildSpreader` (plate, cuts, pedestals, TIM parts) | written, not yet run |
| M5 `SyncFromCAD` | written; two-plate version run on RPi4_fresh 2026-10-14, 0 failed checks |
| M6 `SetupSim` + `ReadResults` (Simulation thermal) | written, not yet run |
| M7 Flow Simulation spike + `RunFlow` | spike done (`docs/flow_api_notes.md`), `RunFlow.bas` written, run and fixed 2026-10-01 (API Boolean quirk, 0 K goals) |
| M8 two-resistor chip model | written in SetupSim, not yet run (docs/HOWTO_M8.md) |
| Test of M1 to M6 in one session | `docs/TEST_PLAN.md` |
| Register rebuilt with Vin/Iin/Vout/Iout/Pin/Pout/Pd, Ta, Tj, Rja/Rjc/Rjb; L and W moved to CAD geometry | done 2026-10-01, RPi4 refilled: same results |
| ImportStep: every solid body, rotation of the board and of each body, Include = No by default, tables grow on demand (decisions 29 to 31) | compiled and run 2026-10-14 on a fresh copy (`example/RaspberryPi4/RPi4_fresh.xlsx`): 108 bodies, re-import "none", BuildBoard "Nothing changed", 0 overlaps, parts coloured. Three bugs fixed (see section 9). Rotation checks R2, R4, R6 not done |
| Bare board simulation compared with the hand calculation (tab 6) | done 2026-10-14 (study run, energy balance 0 %). The hand and sim gap is open and not judged. Then the spreader: user-drawn outline, BuildSpreader, SetupSim and ReadResults with the spreader all run end to end on RPi4_fresh (`docs/TEST_PLAN.md`). Next: a board with heat chips on both sides to test the bottom plate in SetupSim, ReadResults per plate. SyncFromCAD (as built per plate) and tab 8b (hand check, bottom plate) done 2026-10-14 |
| Workbook template `PCB_thermal_TEMPLATE.xlsx` (numbered tabs in workflow order, notes on every header, new `6 Bare board network`, How to, Read me) | done 2026-10-01, RPi4 data refilled into `RaspberryPi4.xlsx`: 0 differences against the old workbook |
| Web app (page as the calculator, three JSON files, macros on JSON; milestones W1 to W9) | W1 done 2026-10-18: schemas in `webapp/schema/`, RPi4 example converted (`example/RaspberryPi4/project.json`, `cad.json`, `results.json`), `npm test` in `webapp/`. W2 done 2026-10-18: calculation module `webapp/src/calc/`, matches every calculated value of `RaspberryPi4.xlsx` and `RPi4_fresh.xlsx` (`tests/webapp/expected_*.json`). W3 done 2026-10-18: page shell `webapp/dist/PCB_Thermal.html` (`npm run build`): folder picking, save with daily backup and autosave, reload of `cad.json` and `results.json`, tabs with unlock rules and reasons, undo. W4 done 2026-10-18: term catalogue `webapp/src/terms/catalogue.js`, formulas with numbers `webapp/src/calc/formulas.js` (each checked against the calculation), hover texts, Explain panel with trace, Show formulas. W5 done 2026-10-18: tabs 1 to 5 complete (paste a parts list, body suggestions, bodies by hand, size overrides, PCB size and stack, import settings, flag limits, air and surfaces, heat paths drawing per part), Learn mode with plain readings, the RPi4 example inside the page (read only). Next: W6 macros on JSON. The workbook and current macros stay in use until W7 |

Files in this folder:

- `workbook/PCB_thermal_TEMPLATE.xlsx` the template. One copy per board, renamed to the board name. Keep it empty.
- `example/RaspberryPi4/RaspberryPi4.xlsx` the reference case: a template copy filled with the RPi4 data (worked example).
- `src/*.bas` the macros (the only files that ship, with the template, `docs/` user guides and `src/icons/`).
- `example/RaspberryPi4/` everything specific to the reference board: the STEP, `RaspberryPi4.xlsx`, the test copy
  `RPi4_fresh.xlsx`, and `CAD/` (built parts, studies, reports; not in git).
- `tests/` test macros and launch scripts; `tests/swp/` the compiled test projects PTTest and PTFresh16 (not in git).
- Old versions live in git history, not as copies. `backup/` folders next to a workbook (daily copy by PT_DailyBackup) are not in git.
- `docs/cad_geometry_topview.png` top view of the imported layout.
- `docs/plate_concept.png` concept of an auto-derived spreader outline (board minus clearances).

---

## 2. Decisions already made (do not reopen)

1. Tools: SolidWorks 2023 VBA + Excel only. No Python on the user's PC.
2. Excel owns numbers, SolidWorks owns shape (details in section 5).
3. Files are created from the user's SolidWorks templates (paths in Settings; blank = SolidWorks defaults).
4. Change reports: every macro that changes geometry data logs to a `Change log` tab. The user ticks each row.
5. Conflicts between an Excel override and a CAD edit: **CAD wins**. The old Excel value is kept in the log.
6. Build macros **refuse to run** while the Change log has unchecked rows.
7. No swap detection for now. A swapped part is logged as one removed plus one added.
8. BuildSpreader does **not** add fins.
9. Licences: user has Simulation (thermal) and Flow Simulation. Stage 1 = Simulation Thermal. Stage 2 = Flow.
10. Chip model: start with a calibrated block (k through thickness from theta_jc). Two-resistor model
    (junction, case, board; needs theta_jb) comes later as an option in the same column.
11. TIM settings are per chip, stored once (Spreader tab), read by both BuildSpreader and SetupSim.
12. Spreader parts are fixed in the assembly at coordinates. No mates.
13. Every macro must work on ANY PCB STEP, not only the RPi4. No part names, RefDes or sizes in code;
    defaults come from the geometry.
14. Body key = NormKey of the imported name plus `_<n>` when the instance number is above 1 (two identical
    connectors get different keys), deduplicated with `_d2`. Stored in `STEP import` column K; CAD geometry AG
    mirrors it (manual rows: `m_<refdes>`). The key follows the imported name, not the name override.
15. M1 is done by a macro (`SetupWorkbook`), not by editing the file, so it works on any copy of the workbook.
    Spreader and Sim setup chip rows are keyed by a RefDes typed in column B (prefilled on first run), so inputs
    never slide to another chip when the body list changes.
16. Each macro project = `PTCommon.bas` (shared code) + the macro's own module.
17. BuildSpreader does not use Equation-manager global variables. Cuts and pedestals are macro-owned features
    rebuilt from the Spreader tab when their numbers change (parameters kept in part custom properties
    `PT_<feature>` in micrometres); the underside plane height and plate thickness are changed in place through
    the feature definitions. Reason: equation text depends on the Windows decimal separator (Arabic locale) and
    dimension-driven sketches are fragile through the API. Excel owns these numbers anyway (section 5).
18. BuildBoard has one path: with no assembly it creates everything, otherwise it updates. Create logs nothing;
    update logs every change. Geometry is checked after each feature (body box) and flipped if SolidWorks put it
    on the wrong side, so plane orientation in the user's templates does not matter.
19. Build and sync macros compare the REAL component box in the assembly (GetBox), not stored numbers, so a
    part edited inside its own file or moved by hand is seen correctly by BuildBoard and SyncFromCAD.
20. SyncFromCAD also defines As_Area = net area of the Top face, As_TopArea = all up-facing faces,
    As_ExposedArea = total surface minus pedestal contact faces minus the chip footprints of pedestal-less chips.
21. BuildBoard cuts every part box at the PCB surface (through-hole and edge connectors), so the simulation has no
    overlapping bodies. CAD geometry keeps the full size; SyncFromCAD compares with the same cut (PT_ClipToBoard).
22. SetupSim: steady state is SolutionType 1 (the API help example uses 0 = transient). All Simulation objects are
    late bound, no extra reference. Units SI: temperatures to the solver in kelvin, conductivity W/m K.
23. "Include parts without heat" = No is done with configuration PT_HeatOnly (the API has no exclude-from-analysis
    call). Build and sync macros leave that configuration before they work (PT_LeaveHeatOnly).
24. The workbook is a template (`PCB_thermal_TEMPLATE.xlsx`). Tabs carry a step number in workflow order
    ("1 Component register" ... "19 Materials"). Macros find tabs by the words after the number (`PT_Sheet`), so old
    unnumbered workbooks still work. Every header has a cell note.
25. Blank `Set_Prefix` = workbook file name + "_" (`PT_Prefix`). Renaming the project file therefore changes CAD file names.
26. Chip lists of Spreader and Sim setup are filled by ImportStep and BuildBoard when empty (`PT_FillChipLists`).
    SetupWorkbook is only for workbooks older than the template (name `Tpl_Version` marks a template copy).
27. `6 Bare board network`: hand network per RefDes with no heatsink (top path through the case, board path through
    PCB spreading, one shared board node solved in closed form). Surface h and emissivity come from Sim setup.
    It reads Sim Tj from `10 Sim results` by run name (blank = newest). SetupSim makes a bare board study when the
    assembly has no unsuppressed spreader.
28. ImportStep links bodies to the register by name: an empty CAD geometry RefDes cell gets the register RefDes whose
    NormKey equals the body NormKey (ECAD exports name bodies by designator). Typed RefDes are never changed and a
    RefDes is never given to two bodies. Vendor models (like the RPi4) keep the manual link.
29. ImportStep imports EVERY solid body (setting `Imp_TopOnly` = No, the default; Yes = old one row per top-level
    component). A component with one body keeps the component name (keys and inputs carry over); a component with several
    bodies gets `<comp name without instance>_<body name><instance>`. New rows get Include = No: the user switches on the
    parts to build, BuildBoard builds only Include = Yes. Hidden components are read too (Include = No keeps them out).
    Surface bodies are counted and reported, not imported.
30. Rotation. The shape of each body is measured from its surface points (face triangles, else `Body2.GetTessellation`;
    component transform applied after a self-check against `Component2.GetBox`). The board frame is turned with the PCB when
    the board is not square to the model axes. A body turned in the board plane gets `Rot` (deg, -45..45), `Own L`, `Own W`
    (STEP import N:P, CAD geometry AI:AK). Xmin, Ymin, L, W stay the axis-aligned ENVELOPE of the turned rectangle, so
    SyncFromCAD, BuildSpreader, SetupSim and every box comparison are unchanged. BuildBoard builds the turned rectangle (the part
    footprint is a 4-line sketch inside the envelope; part properties `Rot_mdeg`, `OwnL_um`, `OwnW_um`). It does so only when
    the envelope of Rot/Own L/Own W still equals the used L and W, so an L or W override gives a plain box. Overlap check
    uses the turned shapes. Rotation about axes other than the board normal (tilted parts) is not modelled.
31. Capacity. The body tables are not fixed at 60 rows. ImportStep inserts rows into CAD geometry (before its manual rows) and
    lengthens `Imp_Table`, `ImpPrev_Table`, `Imp_Rot`, so Excel moves every formula and name. The first run on a workbook adds
    the Rot columns and sets `Imp_TopOnly` = No (`UpgradeWorkbook`, name `Imp_Rot` marks it done). Name `Imp_Level` stores
    "bodies" or "components"; when it changes the change report is skipped once.

---

## 3. Current workbook layout (exact)

Tab names below are the base names. In the template each has its step number in front (see decision 24);
`6 Bare board network`, `13 How to` and `14 Read me` are described on the tabs themselves.

Frame used everywhere: origin at the PCB corner on its bottom face, X along the board length, Y along the width,
Z up through the component side, mm.

### Component register (rebuilt 2026-10-01, electrical columns)
- Inputs B8:S53: RefDes B, Part C, Class D, Package E, Rail F, Vin G, Iin H, Vout I, Iout J, eta K, R L,
  method override M, Pd meas N, Tj max O, Rja P, Rjc Q, Rjb R, Ta override S. No L or W: part size comes from CAD geometry.
- Calculated T:AJ: method T, Pin U, Pout V, area W (CAD geometry used box by RefDes), Pd calc X, Pd used Y (`Reg_QUsed`),
  Ta Z, rise AA, Tj est AB, margin AC, W/cm2 AD, share AE, flag AF ("No Pd" when empty), why AG, ranks AH AI, rail # AJ.
- Pd calc: LDO and switcher use Pin - Pout when both are known, else (Vin - Vout) Iout or Pout (1/eta - 1);
  digital Pin - Pout; resistive I^2 R with I = Iin (else Iout); manual = Pd meas.
- `D4` total Pd. `D5` ambient. Flag limits AA3:AD5. Rjb feeds `6 Bare board network` column L and `9 Sim setup` column E.
- U3 (PMIC) appears on 4 rows (one per rail). Anything per part must sum by RefDes (SUMIFS on Y by B).

### Thermal solution (existing)
- Plate inputs D6:D12 (L, W, t, k, h, faces, R override). Calcs J5:J12.
- Rows 16..68 mirror register rows 8..53 (row r <-> register row r-8). Inputs H..R per row.
- `AF` = Sim Tj input, `AG` = hand minus sim. ReadResults writes AF.
- Plate is modelled as L x W. It must read area/thickness from the Spreader tab "as built" later (section 7.4).

### CAD geometry (built this session)
- PCB outline: rows 7..9 (Length X, Width Y, Thickness Z); C imported, D override (input), E used.
- PCB stack: inputs D12:D16 (layers, Cu um, coverage, k Cu, k dielectric); results D17:D20
  (D19 = k in-plane, D20 = k through-plane).
- Checks M6:M12 (bodies imported, included, register Q, Q placed, Q not placed, register rows with no body,
  rows with a problem).
- Body table header row 25, rows 26..85 mirror `STEP import` rows 9..68, rows 86..95 manual parts.
  - B #, C imported name (formula), D name override (input), E name used, F RefDes (input), G Include (input, Yes/No)
  - H..M imported Xmin Ymin Zmin L W H (formulas), N side
  - O..T override Xmin Ymin Zmin L W H (input)
  - U..Z used (formula: override else imported), AA X centre, AB Y centre
  - AC Q from register (SUMIFS by RefDes), AD W/cm2, AE check text
- Manual rows 86, 87 hold L1 and L2 (inductors not in the STEP).

### STEP import (built this session)
- D4 model name/path, D5 import time, D6 frame description.
- Settings (inputs): M3 PCB name contains, M4 skip bodies with L and W both under (mm), M5 top-level only (Yes/No).
- Header row 7: B #, C body name, D Xmin, E Ymin, F Zmin, G L, H W, I H, J side. Row 8 = PCB, rows 9.. bodies.
- L8:L200 messages from the last import.

### Names in the SolidWorks model
SolidWorks names imported parts like `Raspberry Pi 4 Model B.STEP-1/PCB, RPi4ModelB.STEP-1`
(sub-assembly path, `.STEP` in the name, instance suffix). Matching keys strip path, instance suffix and
file extension, then keep only lowercase letters and digits (`NormKey` in ImportStep.bas).

---

## 4. Workbook additions (milestone M1)

### 4.1 Settings block (new tab `Settings`)
| Name | Meaning |
|---|---|
| `Set_PartTemplate` | path to .prtdot (blank = SolidWorks default part template) |
| `Set_AsmTemplate` | path to .asmdot (blank = default) |
| `Set_OutputFolder` | output folder (blank = `CAD\` next to the workbook) |
| `Set_Prefix` | file name prefix, e.g. `RPi4_` |
| `Set_TolPos` | position change tolerance for change reports, mm (default 0.05) |
| `Set_TolSize` | size change tolerance, mm (default 0.05) |
| `Set_AsmName` | thermal assembly file name (default `<prefix>Thermal.sldasm`) |

### 4.2 Change log tab (`Change log`)
Append-only, newest row on top (insert at row 8). Columns:
`Date | Macro | Key | Name | RefDes | Change | Field | Old | New | Delta | Checked`

- `Change` values: `added`, `removed`, `moved`, `resized`, `renamed`, `override from CAD`, `conflict`,
  `excel-owned value differs`, `check failed`.
- `Checked`: the user types `Yes`. Data validation list `Yes`.
- Summary cell `Log_Unchecked` = number of rows with a Change but Checked <> Yes. Build macros read it and refuse
  when > 0, listing the first rows to check.

### 4.3 `STEP import (previous)` tab (hidden)
Same layout as `STEP import`. ImportStep copies the current import here before writing the new one.

### 4.4 CAD geometry: new column `AF` "Last change"
Formula: latest unchecked Change log entry for this row's key (COUNTIFS/INDEX on the log). Conditional format
highlights non-empty cells. Also add a hidden column with the row's `Key` (NormKey of the name) so formulas and
macros can join on it.

### 4.5 Spreader tab (`Spreader`)
Section A, global inputs:
`Spr_Material`, `Spr_k`, `Spr_Density`, `Spr_Thickness` (t_plate), `Spr_ClearXY`, `Spr_ClearZ`,
`Spr_MinPedestal`, `Spr_UnderMode` (Auto/Manual), `Spr_UnderManual` (mm).

Section B, per chip rows (one row per included body with a RefDes). Filled by formulas from CAD geometry, inputs in blue:
`RefDes | Key | X min | Y min | L | W | Top Z | Covered (Yes/No) | Height mode (Pedestal/Thick TIM) |
TIM mode (Body/Contact R) | TIM t nominal | TIM k | -> calc: gap | TIM t used | Pedestal h | TIM R (K/W)`

Calculations:
- `z_under` = Auto: max(Top Z of covered chips) + that chip's TIM t nominal. Manual: `Spr_UnderManual`.
- gap_i = z_under - TopZ_i
- Pedestal mode: TIM t used = TIM t nominal, pedestal = gap_i - TIM t used; if pedestal < Spr_MinPedestal then
  pedestal = 0 and TIM t used = gap_i.
- Thick TIM mode: pedestal = 0, TIM t used = gap_i.
- TIM R = TIM t used / (TIM k x L x W) (units handled).

Section C, clearance cuts (formula list): every included body whose top Z > z_under - Spr_ClearZ gets
`Key | X min - ClearXY | Y min - ClearXY | L + 2 ClearXY | W + 2 ClearXY`.

Section D, as built (written by SyncFromCAD only):
`As_Area` (outline area), `As_TopArea`, `As_ExposedArea`, `As_Volume`, `As_Mass`, `As_Thickness`,
`As_Holes` (count of circular through edges), `As_LastSync`, and per chip: pedestal h and TIM t as found,
plus a flag column "differs from design".

### 4.6 Sim setup tab (`Sim setup`)
Global: `Sim_Type` (Thermal / Flow), `Sim_RunName` (auto stamped), `Sim_Ambient` (= register D5),
`Sim_MeshGlobal`, `Sim_MeshFine` (chips and TIMs), `Sim_IncludeNonHeat` (Yes/No).
Surface groups (h, emissivity): Spreader top, Spreader other, PCB top, PCB bottom, Parts.
Interface contact resistance (K*m2/W): part-to-PCB, chip-to-TIM, TIM-to-spreader.
Per chip: `RefDes | Chip model (Block / Two-resistor) | theta_jc (from register) | theta_jb (input, unused for now) |
k_z calibrated = H / (theta_jc x A) | k_xy`.

### 4.7 Sim results tab (`Sim results`)
One row per run, appended: run name, date, study name, snapshot of inputs (h values, TIM modes, mesh sizes,
spreader file), outputs per RefDes (peak, average), spreader max/min, PCB max/min, energy balance %, plot path.

### 4.8 Named cells
Give every cell or range a macro reads or writes a workbook name. Existing cells too, e.g.
`Reg_Table` ('Component register'!B8:AF53), `Reg_Ambient` (D5), `Geo_Table` (CAD geometry B26:AF95),
`Geo_PCB_Used` (E7:E9), `Geo_kXY` (D19), `Geo_kZ` (D20), `Imp_Settings`, `Imp_Table`, `Imp_Messages`,
`TS_SimTj` (Thermal solution AF16:AF68). Update ImportStep to use them.

Acceptance M1: workbook opens in Excel with no errors, Class dropdown on the register still works,
all new tabs calculate on the RPi4 data, `Log_Unchecked` = 0.

---

## 5. Ownership and identity

| Data | Owner | Written by |
|---|---|---|
| Imported body boxes | STEP | ImportStep |
| Component overrides (position, size) | user in Excel **or** CAD edits (CAD wins) | user, SyncFromCAD |
| PCB outline override | same | user, SyncFromCAD |
| Covered chips, TIM settings, plate thickness, clearances, materials | Excel | user |
| Spreader outline, holes, pockets, user features | CAD | user in SolidWorks |
| Spreader as-built numbers | CAD | SyncFromCAD |
| Pedestal and TIM sizes/positions, clearance cut sizes/positions | derived from Excel | BuildSpreader |

Identity tags. Every component a macro creates gets custom properties:
- `Role` = `PCB` | `Part` | `Spreader` | `TIM`
- `Key` = NormKey of the CAD geometry name used
- `RefDes` = register RefDes (blank for non-heat parts)
- `SourceRow` = CAD geometry row at build time (informational only)
Macros find components by these properties, never by file or instance name.
Named faces (IPartDoc/IEntity names): spreader pedestal bottoms `Contact_<RefDes>`, spreader plate top `Top`,
TIM faces `TIM_<RefDes>_Top` / `_Bottom`, chip block top `Chip_<RefDes>_Top`.

---

## 6. Macro specs

All macros: `Sub main()`, pick the workbook first (file dialog starting at the last used path, stored with
`SaveSetting "PCBThermal", ...`), check it has the required tabs, write, `wb.Save`, summary MsgBox that starts with
where it wrote.

### 6.1 ImportStep changes (M2)
Existing behaviour stays. Add:
1. Copy `STEP import` values to `STEP import (previous)` before writing.
2. After writing, compare previous vs new by Key (then by box within tolerance for renamed):
   `added`, `removed`, `moved` (any of Xmin/Ymin/Zmin changed > Set_TolPos), `resized` (L/W/H > Set_TolSize),
   `renamed` (same box, different Key). One log row per field changed.
3. PCB outline changes logged too.
4. Summary lists counts per change type and says "check them on the Change log tab".
Acceptance: re-running on the same model logs nothing. Moving one component 1 mm in SolidWorks and re-running
logs exactly one `moved` row with the right delta.

### 6.2 BuildBoard (M3), Create / Update
Reads CAD geometry used values (Include = Yes), PCB used outline, Settings.
Refuse if `Log_Unchecked` > 0.
Create:
1. New assembly from `Set_AsmTemplate`, saved as `Set_AsmName` in the output folder.
2. PCB part: box L x W x T at origin, material = custom orthotropic (kxy = Geo_kXY, kz = Geo_kZ) or store k in
   custom properties if material assignment belongs to SetupSim (preferred: SetupSim assigns all materials).
3. One part per included body: box at used Xmin/Ymin/Zmin with used L/W/H. File name `<prefix>P_<Key>.sldprt`.
4. Insert each at fixed position (component transform), fix it. Tag properties (section 5).
Update:
1. For each tagged Part component, compare its box with the used values; move/resize (edit sketch dims or
   rebuild the feature) where they differ; add parts for new Keys; suppress components whose Key is gone.
2. Log every change made (Macro = BuildBoard).
Acceptance: RPi4 builds 20 parts + PCB (18 imported + L1, L2); interference check between parts only shows the
known overlaps (connectors with PCB through pins are acceptable; list them); re-running Update without changes
does nothing.

### 6.3 BuildSpreader (M4), Create / Update
**Changed 2026-10-14 (user decision, supersedes steps 2 and 3 below): the user draws the plate outline.** The user inserts a new part in the
thermal assembly (name contains `Spreader`), sketches any closed outline on the assembly Front Plane (or a face parallel to it), names the sketch
`Outline` and saves the part. The user marks the chips to touch (Covered) and enters the TIM data on the Spreader tab. BuildSpreader completes
the part: plane `PT_Under` at z_under, `PT_Plate` = the user's `Outline` extruded `t_plate` with start offset z_under (so any shape works and
later outline edits flow through), then the cuts, pedestals, face names and TIM parts as below. The part position in the assembly is read from
the component transform (translation allowed, rotation not); all coordinates are shifted by it. A chip or tall part outside the outline is
reported in the notes (no pedestal, no cut). If there is no spreader part, the macro offers the old behaviour (rectangle = PCB outline).
Update: run again; `PT_Plate` start and thickness, cuts and pedestals follow Excel; the user's own features are kept.
**Two plates (2026-10-14):** a second plate under the PCB for chips on the bottom side. Tab `7b Spreader bottom` is a copy of `7 Spreader`
with tab-level names (`Spr_*`), face height = lowest covered chip bottom minus TIM, gap = chip bottom minus face, cuts from CAD geometry
column AL. The bottom part has `Spreader` and `Bottom` in its name. BuildSpreader runs `DoSide` for the top, then the bottom; the plate
spans `zu - t` to `zu`, pedestals go up, TIM boxes hang under the chip (z0 = chip bottom - TIM), TIM face names `_Top` = toward the plate.
SetupSim, ReadResults, SyncFromCAD handle both plates (SyncFromCAD picks each part by "bottom" in its name and writes the bottom as-built block on tab 7b). Hand check: tab `8b Thermal solution bottom` is a copy of tab 8 for the bottom plate (mark the bottom chips Yes there); its Sim Tj column links to tab 8, which ReadResults fills.
Refuse if `Log_Unchecked` > 0.
Spreader part (from `Set_PartTemplate`), file `<prefix>Spreader_rNN.sldprt` (NN increments on each Create):
1. Global variables (Equation manager): `t_plate`, `z_under`; per covered chip `x_U1`, `y_U1`, `L_U1`, `W_U1`,
   `ped_U1`, `tim_U1`; per clearance cut `cx_<Key>`, `cy_<Key>`, `cL_<Key>`, `cW_<Key>`.
2. Plane at z_under. Sketch `Outline` = PCB used outline rectangle (user may edit freely afterwards).
3. Extrude `Plate` up by t_plate.
4. One cut per clearance entry `Clear_<Key>`, rectangle driven by its variables, through all.
5. One boss per covered chip with pedestal > 0: `Ped_<RefDes>`, rectangle driven by x/y/L/W, extruded down
   from the underside by `ped_<RefDes>`. Name its bottom face `Contact_<RefDes>`. Where pedestal = 0 the contact
   is the plate underside: record that in the Spreader tab row, and SetupSim selects the underside face region
   by the chip footprint.
6. Name plate top face `Top`. Custom properties Role=Spreader.
TIM parts `<prefix>TIM_<RefDes>.sldprt`: box L x W x tim, placed on the chip top. Role=TIM, RefDes. Only for
TIM mode = Body. For Contact R mode no body is made (SetupSim applies a contact resistance instead).
Update:
- Rewrite global variables from the Spreader tab (moves and resizes pedestals, cuts, TIMs).
- New covered chip or new tall part: add feature at the end of its group. Removed: suppress, do not delete.
- Never edit `Outline` or any feature not created by the macro (track macro features by name prefix).
- Log every change (Macro = BuildSpreader).
Acceptance: on RPi4 with all U-chips covered, every covered chip has chip top + TIM + pedestal bottom touching
(distance 0 within 0.001 mm); no interference between spreader and any part; adding a mounting hole in
SolidWorks then running Update keeps the hole.

### 6.4 SyncFromCAD (M5)
Reads the thermal assembly:
1. PCB box vs CAD geometry used outline; differences -> write to override D7:D9, log `override from CAD`.
2. Each Role=Part component box vs used values; differences -> write overrides O..T, log. If the Excel override
   already had a different value, log `conflict` with the old Excel value (CAD wins).
3. Spreader: measure area, top area, exposed area, volume, mass, thickness, holes -> Spreader as-built. Compare
   pedestal heights and TIM thicknesses with design -> `excel-owned value differs` rows (not accepted).
4. Checks, each failure logged as `check failed`:
   - every covered chip still has a contact face over it and its TIM touches both sides;
   - no clearance cut suppressed or removed while its tall part is included;
   - interference check across the assembly;
   - Excel-owned thicknesses match.
Acceptance: moving a part box 2 mm in SolidWorks then Sync writes the override and logs one row; deleting a
pedestal logs a failed check; adding holes updates As_Area and As_Holes.

### 6.5 SetupSim (M6), Simulation Thermal
Refuse if `Log_Unchecked` > 0.
Using the SolidWorks Simulation API (CosmosWorks add-in; verify exact calls in the SW 2023 API help:
study manager CreateNewStudy, thermal study options steady state, solid body materials, heat power,
convection, radiation, contact sets with thermal resistance, mesh, run):
1. New steady-state thermal study named `Sim_RunName` (never overwrite older studies).
2. Materials: PCB orthotropic (Geo_kXY, Geo_kZ); chips calibrated block (k_z = H / (theta_jc x A),
   k_xy from Sim setup); parts per Sim_IncludeNonHeat (exclude from analysis or give a generic k);
   spreader from Spreader tab; TIM bodies with TIM k.
3. Heat power per RefDes body = SUMIFS register V by RefDes.
4. Contacts: global bonded; thermal contact resistance on interfaces set in Sim setup; Contact R TIMs as a
   resistance between chip top and contact face.
5. Exposed-face finder: all faces of each group minus faces touching another body (coincident within tolerance).
   Apply convection (h, ambient) and radiation (emissivity) per surface group.
6. Mesh (global + fine on chips/TIMs), run. Stop before meshing if anything is missing, and say what.
Acceptance: study runs on RPi4; total applied heat = register total; energy balance within 2 %.

### 6.6 ReadResults (M6)
1. Per RefDes: peak and average temperature of the chip body -> `Sim results` row, peak -> Thermal solution AF.
2. Spreader and PCB max/min; energy balance.
3. Save plot image to `CAD\results\<run>.png`.
Acceptance: AF column fills, AG shows hand minus sim, results row appended.

### 6.7 Flow Simulation (M7), stage 2
Start with a spike: what the SW 2023 Flow Simulation API can set (heat sources, materials, ambient, goals,
run, read goals). Expected route: user builds one Flow project once from a template; macro updates heat source
values and materials by name, runs, reads goals into `Sim results`. Output of the spike: a short note in
`docs/flow_api_notes.md` and a revised M7 spec.

Revised spec (spike done): `RunFlow` = `PTCommon` + `RunFlow.bas`, see `docs/HOWTO_M7.md`. Sim_Type must be Flow.
Volume sources whose name contains a RefDes get the register heat (heat generation rate, W). Initial temperature =
Sim_Ambient. Optional solve (`Solve2`). Goals whose name contains a RefDes plus max/avg/min (and Spreader/PCB goals)
go to `Sim results` (new row `FL_<stamp>`) and to Thermal solution AF. Report `CAD\<prefix>RunFlow_<run>.txt`.
Acceptance: total heat set = register total; AF fills for every chip with a max goal; a RefDes with heat and no source is listed.

### 6.8 Two-resistor chip model (M8)
Per chip where Chip model = Two-resistor: split the block into junction node with theta_jc to top and theta_jb to
board (e.g. thin top and bottom layers with calibrated k, heat applied to the core). Needs theta_jb in Sim setup.

---

## 7. Hand network changes (with M4 or M6)
1. Thermal solution plate area and thickness from Spreader as-built instead of L x W.
2. Per-chip TIM thickness and k from Spreader rows.
3. Add a board path (junction to board to air) so the hand model and the simulation share the same physics.
4. Keep the template's existing formulas working; add columns rather than rewrite.

---

## 8. Milestones and order

| # | Milestone | Depends on |
|---|---|---|
| M1 | Workbook additions (4.1 to 4.8), named cells | - |
| M2 | ImportStep change report | M1 |
| M3 | BuildBoard | M1 |
| M4 | BuildSpreader | M3 |
| M5 | SyncFromCAD | M3, M4 |
| M6 | SetupSim + ReadResults (Simulation Thermal, calibrated blocks) | M4 |
| M7 | Flow Simulation spike, then macros | M6 |
| M8 | Two-resistor chip model | M6 |
| - | Hand network changes (section 7) | M4 |
| W1 to W9 | Web app replacing the workbook as the user interface (`docs/WEBAPP_DESIGN.md`, section 12) | M6 |

Run order for a new board: M2 ImportStep, M3 BuildBoard, then the **bare board** part of M6 (SetupSim, ReadResults) and the
comparison with `6 Bare board network`, and only then M4 BuildSpreader and a second M6 run with the spreader. The bare board
run is the baseline that shows how much the spreader helps.

Deliver each milestone as: updated workbook (if changed), `.bas` files, a short note of how to import and run,
and the acceptance checks the user should see.

---

## 9. Known pitfalls from this session
- `Imp` as a variable name: compile error. `mid` shadows `Mid()`.
- Passing `bodies(i)` from a Variant array into `ByRef o As Object`: "ByRef argument type mismatch".
- SolidWorks component names include the sub-assembly path, `.STEP` and `-1`. Match with NormKey.
- SolidWorks `GetBox` is slightly loose on curved parts (up to about 0.13 mm on connectors). Chips were exact.
- The STEP opened inside another assembly: top level had one sub-assembly. ImportStep steps into single
  wrappers.
- The first import cleared the RefDes column because names did not match; the fix was NormKey plus a
  box-match fallback. Never clear user inputs before the new values are matched.
- openpyxl drops the x14 Class dropdown on save. See CLAUDE.md.
- Macros must tell the user where they wrote. The user could not find the output on the first run.
- `Component2.SetMaterialPropertyValues2` returns without error and stores nothing here, so parts stayed uncoloured while the macro
  reported success. Use `Component2.MaterialPropertyValues = arr` and check `HasMaterialPropertyValues` (fixed 2026-10-14 in `PT_ApplyColor`).
- SolidWorks names a STEP `RPi4ModelB_1.STEP` instead of `RPi4ModelB.STEP` when the same file is already open in the session. That `_1`
  changed every body key (108 false "renamed" rows). `NormKey` now drops `_<n>` right before `.STEP`.
- A formula whose range ends exactly on the last row of a table is not stretched when rows are added after that row (the CAD geometry
  body count `M6`). Point such formulas at a defined name (`INDEX(Geo_Table,0,2)`). `EnsureCapacity` repairs `M6` on old copies.
- `ReadResults` energy balance: use the net flow of the whole model against the register total. Gross in and out include exchange with the
  ambient (5.89 W in, 11.59 W out for 5.70 W applied) and never match.
- From a shell, `GetOpenDocSpec/OpenDoc7/OpenDoc6` fail on a STEP (error 2097152); open it with File > Open. Do not start SolidWorks again with
  Computer Use `open_application`: it launches a second instance.
- Never call `IExtrudeFeatureData2.AccessSelections` (or any `GetDefinition` + `AccessSelections`) just to read a value: without `ModifyDefinition` or
  `ReleaseSelectionAccess` the part stays rolled back to before that feature, GetBodies2 returns nothing and the saved file keeps the rollback bar.
  Read `FromType` and `GetDepth` from `GetDefinition` directly. `FeatureManager.EditRollback 1, ""` moves the bar to the end.
- `AddComponent5(path, ..., x, y, z)` centres the part's bounding box at x, y, z instead of placing its origin there. Place new components with `PT_PlaceAndFix`.
