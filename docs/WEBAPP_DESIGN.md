# PCB thermal web app: design

Status: design, agreed in outline on 2026-10-18. No code yet. This replaces the Excel workbook as the place where the
user works. The SolidWorks macros stay and keep doing the CAD and simulation work.

## 1. Goal

A fresh graduate engineer can take a PCB from STEP file to a checked thermal result, and learn the thermal engineering
while doing it. The tool does the manual work (building the model, setting up and running the study). The user does
the engineering: power numbers, datasheet values, judging the results.

What that means for the design:

1. **Every value explains itself.** No user should have to ask "what is this number". Each value, unit, abbreviation
   and status word has a short explanation on hover and a longer one on request.
2. **Every calculated value can show its formula,** with the user's own numbers put in.
3. **Steps unlock in order.** A step that cannot run yet is greyed out and says why, and what to do first.
4. **The structure cannot be broken.** Parts are records the user adds and removes. There are no rows or formulas to
   delete by mistake.
5. **Nothing is hidden.** Explanations stay out of the way until asked for, but they are always one hover away.

The current workbook is a reference for the numbers and formulas only. The page does not copy its layout.

## 2. Pieces

| Piece | What it is | Writes |
|---|---|---|
| **The page** | One HTML file, `PCB_Thermal.html`. Opens in Chrome or Edge from disk. No install, no server, no internet. | `project.json` only |
| **The macros** | The 8 SolidWorks VBA macros (and the add-in port, where IT allows it). They read the project and do the CAD and simulation steps. | `cad.json`, `results.json`, the log |
| **The project folder** | One folder per board. Holds the three JSON files, the STEP file, `CAD\` (built parts, studies, pictures) and `backup\`. | |

Excel is no longer needed to run a study. It becomes an export format (section 9).

### 2.1 How the page reaches the files

- The user picks the project folder once (Chrome and Edge "File System Access"). The page remembers it for next time.
- The page saves `project.json` when the user presses Save, and on its own every 30 s while there are unsaved edits.
- While the page is open it checks the modified time of `cad.json` and `results.json` every 2 s. When a macro has
  written one, the page reloads it and shows a short message ("Import STEP finished: 108 bodies").
- Other browsers (Firefox): open and download files by hand. The page says so on start.

### 2.2 How the macros reach the files

- Each macro asks for the project folder the first time, like the workbook today, and remembers it. Every run logs
  which folder it used, and the summary box names it.
- Macros read `project.json` and write their own file. They write to a temporary file and then rename it, so the page
  never reads a half-written file.
- The page cannot start SolidWorks. When the next step is a SolidWorks step, the page shows a card: which button to
  press, what it will do, and what the page will show when it is done.

## 3. The three files

One writer per file. The page and a macro never save the same file, so neither can wipe the other's work.

All lengths in **mm**, temperatures in **degrees C**, power in **W**, conductivity in **W/(m K)**, resistance in
**C/W**. The macros convert to metres at the SolidWorks boundary, as today. Every file starts with:

```json
{ "schema": "pcb-thermal/project", "schemaVersion": 1, "savedBy": "page 0.1.0", "savedAt": "2026-10-18T11:20:00" }
```

### 3.1 `project.json` (the page writes it)

Everything the user decides, plus the numbers the macros need, already worked out.

| Section | Holds | Replaces |
|---|---|---|
| `board` | name, STEP file name, ambient, measured board input power, notes (enclosure, load case) | register header, rail budget E35, Read me |
| `parts[]` | one per RefDes: id, part number, class, package, Tj max, Rja, Rjc, Rjb, Ta override, notes, and `loads[]` | 1 Component register (inputs) |
| `parts[].loads[]` | one per heat source in the part: label, rail, method override, Vin, Iin, Vout, Iout, efficiency, R, measured Pd. Part Pd = sum of its loads. | register rows (a PMIC with four bucks was four rows) |
| `bodies{}` | user choices per body, by body key: include, RefDes link, name override, box override. A body typed in by hand (key `m_...`) carries its own name and box. | 3 CAD geometry (blue columns) |
| `pcbOverride` | PCB L, W, thickness that replace the imported ones | 3 CAD geometry D7:D9 |
| `pcbStack` | copper layers, copper thickness (mm), coverage, copper k, dielectric k | 3 CAD geometry PCB stack |
| `spreader` | `top` and `bottom` plate: enabled, material, k, density, thickness, clearances, smallest pedestal, height mode and manual height, per chip: covered, height mode, TIM mode, TIM t, TIM k | 7 and 7b Spreader (inputs) |
| `handCheck` | surface rise for radiation, run to compare, and the thermal solution inputs (top and bottom: plate, h, faces, per part TIM1, local spreader, TIM2, plate share) | 6 Bare board network, 8 and 8b Thermal solution (inputs) |
| `simSetup` | study type, mesh sizes, include parts without heat, k of parts without heat, h and emissivity per surface group, contact R per interface, per chip model and k xy | 9 Sim setup (inputs) |
| `settings` | templates, output folder, file prefix, assembly name, tolerances, Import STEP settings, flag limits | 12 Settings, 2 STEP import M3:M5, register flag limits |
| `forSolidWorks` | written on every save: ambient, PCB size and k in-plane and through, bodies to build with their used box, heat per RefDes, per chip model, k z, k xy, Rjc, Rjb, per plate face height, material, chips (TIM t, pedestal, TIM mode, k, R) and clearance cuts | grey columns the macros read today |
| `checks` | the user's ticks on change records, by change id | Change log "Checked" column |

Rules for all three files (fixed in W1, `webapp/schema/`):

- A blank value is left out. `null` is never written. Unknown fields are refused, so a typo cannot hide.
- Yes/No cells become `true`/`false`. Dates are local time, `2026-10-18T11:20:00`.
- Change ids are `yyyymmddThhmmss-n` (time of the change, then a counter), so the page can store a tick by id.
- **Parts by RefDes.** The workbook worked out Tj est per register row, so each buck of a PMIC got its own Tj from
  its own share of the heat. The page works per part: Tj est uses the part's total Pd. This changes the RPi4 U3
  numbers on purpose; the W2 tests compare per RefDes.

`forSolidWorks` is the only place the macros read numbers from. The page computes it, so the macros never repeat a
calculation.

### 3.2 `cad.json` (the macros write it)

What SolidWorks knows.

| Section | Written by | Holds |
|---|---|---|
| `import` | Import STEP | model file, date, board frame, PCB size, settings used, `bodies[]` (key, name, X Y Z min, L W H, side, rotation, own L and W), messages |
| `importPrevious` | Import STEP | the import before this one, for the change report |
| `build` | Build Board | assembly path, date, per body: built or failed (with reason), cut at board surface, colour, overlaps |
| `spreaderBuild` | Build Spreader | plate as built, pedestals, TIM parts, failures |
| `asBuilt` | Sync CAD | sizes and positions read back from the assembly, spreader area, volume, mass, failed checks |
| `changes[]` | Import STEP, Build, Sync | one record per change: id, date, macro, key, field, old, new. The user ticks them in the page. |

### 3.3 `results.json` (the macros write it)

| Section | Written by | Holds |
|---|---|---|
| `runs[]` | Run Thermal Study, Read Results, Run Flow | run name, date, study type, settings used (a copy of `forSolidWorks` at run time), mesh, per RefDes peak and average, PCB and spreader max and min, energy balance, plot picture path, notes |

Keeping a copy of the inputs with each run means an old result always shows the inputs it was made with, even after
the user changed them.

### 3.4 Backups

- Before the first save of the day, the page copies `project.json` to `backup\project_yyyymmdd.json`.
- The macros do the same for their file before writing.
- The page has "Open a backup" in the Project menu.

## 4. The page

### 4.1 Layout

- **Top bar:** project name, save state ("Saved 11:20" or "Unsaved changes"), Show formulas switch, Learn mode switch,
  Dark/light switch, Help, Export.
- **Colour theme:** dark by default, light on request. The choice is remembered in the browser, not in the project.
  Status colours (CRITICAL, Watch, OK, locked) are chosen to read the same in both themes. Print and PDF export always
  use the light theme.
- **Step tabs** across the top, in workflow order. Each tab shows its state: locked, ready, done, or needs attention.
- **Main area:** the tab content.
- **Explain panel** on the right, closed by default. It opens when the user clicks a value's info icon (section 5).

### 4.2 Tabs and when they unlock

| # | Tab | What the user does | Unlocks when | Greyed-out message |
|---|---|---|---|---|
| 0 | **Project** | Pick the folder, name the board, set ambient. See the overall status and the next step. | always | |
| 1 | **Parts and power** | Add parts, pick a power method, enter datasheet values. Pd, Tj estimate, margin and flag are calculated as they type. | project folder chosen | "Choose the project folder first, so your work can be saved." |
| 2 | **Import geometry** | SolidWorks card: run Import STEP. Then see the board and its bodies on a top-view drawing. | parts tab has at least one part with Pd | "Add at least one part with power first. The import links bodies to your parts by RefDes." |
| 3 | **Link bodies** | For each part, pick its body (suggested by name). Include or skip bodies. Override sizes if the STEP is wrong. | `cad.json` has an import | "Run Import STEP in SolidWorks first (button 2). This tab lists the bodies it finds." |
| 4 | **Triage** | Read only: ranked by Pd and by W/cm2, rail budget, what to worry about first. | parts linked to bodies | "Link at least one part to a body on tab 3. Power density needs the body footprint." |
| 5 | **Hand check: bare board** | Read only: the resistance network per part, board temperature, pass or fail. Learn the physics here. | triage has no blocking problem | "Fix the problems listed on tab 4 first: {list}." |
| 6 | **Build model** | SolidWorks card: run Build Board. See what was built, what failed and why, overlaps. Tick changes. | at least one body included | "Include at least one body on tab 3. Build Board builds a box for each included body." |
| 7 | **Spreader** (optional) | Choose a plate, cover chips, TIM. Hand check with spreader. SolidWorks card: Build Spreader, then Sync. | model built, no unticked changes | "Run Build Board first (tab 6)." or "Tick the {n} changes on tab 6 first. Building on top of unchecked changes can hide a mistake." |
| 8 | **Simulation setup** | Surface h and emissivity, contacts, mesh, chip models. Shows the k z each chip will get and why. SolidWorks card: Run Thermal Study. | model built, no unticked changes | same as above |
| 9 | **Results** | Temperatures per part, hand vs simulation, energy balance, pass or fail, run history, plot picture. | `results.json` has a run | "Run the study first (tab 8). Results appear here when Read Results has finished." |
| 10 | **Changes** | All change records, tick them. | always (badge shows unticked count) | |
| 11 | **Help** | Section 7. | always | |

Rules:

- A locked tab can still be opened to read. Its inputs are greyed and its SolidWorks card shows the reason.
- Every reason names the fix and links to it ("Go to tab 3").
- A tab that is done but whose inputs changed later shows "out of date" with what changed ("Pd of U5 changed after the
  last build"). The page compares `forSolidWorks` with the copy kept in `build` and `runs[]`.

### 4.3 SolidWorks cards

Each SolidWorks step is a card on its tab:

- **What it does:** "Builds one box per included body and a PCB plate, in a new assembly."
- **Before you press it:** the checks, each with a tick or a cross ("Model saved", "No unticked changes").
- **Press:** "In SolidWorks: PCB Thermal tab > 3 Build Board." with the button's icon.
- **Waiting:** after the user says they pressed it, the card shows "Waiting for Build Board..." until `cad.json`
  changes, then the outcome.
- **Outcome:** the summary in plain words, with failures listed first.

## 5. Explaining every value

### 5.1 The term catalogue

Every value on every tab has an entry in one catalogue, `terms` (built into the page). Each entry has:

| Field | Example (Rjc) |
|---|---|
| `id` | `part.rjc` |
| `label` | Junction to case resistance |
| `symbol` | R<sub>jc</sub> (also written theta jc) |
| `unit` | C/W |
| `kind` | input, calculated, from SolidWorks, or from a run |
| `short` | "How many degrees the chip junction is above its case top for every watt it makes." |
| `long` | what it is, where to find it in a datasheet, typical values, common mistakes (Rjc from the top vs from the board side, Rja is not Rjc) |
| `formula` | for calculated values (section 6) |
| `usedBy` | the values that depend on it ("k z of the chip block", "R junction to top") |
| `range` | a typical range; the page warns outside it ("0.1 to 50 C/W for ICs") |
| `help` | link to a Help section |

The first version of the catalogue starts from the 729 header notes of the current template, already extracted with
their formulas into `docs/webapp/workbook_terms.json`. Each note gets rewritten for a graduate reader and split into
`short` and `long`.

### 5.2 How it shows

- **Hover** (or keyboard focus, or tap on a touch screen) on a label or value shows `short`, the unit, and the kind
  ("calculated" or "you type this").
- **Click the small info icon** next to a value: the Explain panel opens with `long`, the formula with numbers,
  where the inputs come from (clickable), what uses it, and the help link.
- **Abbreviations** in any text (Pd, Tj, Ta, Rja, Rjc, Rjb, TIM, RefDes, h, k, epsilon, Bi, tau, lambda, phi, W/cm2)
  are marked with a dotted underline and show their meaning on hover. They come from a glossary inside the catalogue.
- **Status words** (CRITICAL, Watch, Minor, Tight, FAIL, OK) show what the word means and the limit that set it
  ("Margin 4.2 C is below the Watch limit of 10 C").
- **Every warning** says what is wrong, why it matters, and how to fix it, with a link to the field.

### 5.3 Learn mode

A switch in the top bar. Off by default after the first project.

- Each tab starts with a short "What this step is for" box and "What to look at" hints.
- Results carry a plain reading: "U5 is 12 C below its limit. Most of its heat (70 %) goes into the board, so a
  bigger copper area under it would help more than a heatsink."
- An **example project** (the Raspberry Pi 4, read only) can be opened from Help to see a finished study.

## 6. Formulas

- Every calculation is written once, in the page's calculation module, as a named formula: inputs, output, the
  formula as text, and a reference where one exists (for example Lee et al. 1995 for spreading).
- **Show formulas** (top bar switch) puts the formula under every calculated value on the screen, for example:

  ```
  Tj est = Ta + Pd x Rja = 40 + 1.20 x 35.0 = 82.0 C
  ```

- With the switch off, the same line shows in the Explain panel for the clicked value.
- **Trace:** in the Explain panel each input of the formula is a link. Clicking it shows that value's formula, so the
  user can walk back to the typed inputs.
- The formulas are the same as the template's. The port is checked against the RPi4 workbook (section 11).

## 7. Help page

| Section | Content |
|---|---|
| Start here | What the tool does, the steps in one picture, what you do and what the tool does |
| Heat basics | Conduction, convection, radiation; thermal resistance as Ohm's law for heat; junction, case, board, air |
| Reading a datasheet | Where to find Pd, Tj max, Rja, Rjc, Rjb; which one to trust; traps |
| Power methods | The template's Q methods tab, rewritten: LDO, switching regulator, digital load, resistive, measured |
| The bare board network | The two heat paths, spreading resistance, the shared board node; with a drawing |
| Spreaders and TIMs | When a spreader helps, pedestals, TIM thickness and k |
| Simulation | What the study does, mesh, contacts, energy balance, why hand and simulation differ |
| Each tab | One section per tab: purpose, inputs, outputs, typical mistakes |
| SolidWorks steps | Each button, what it changes, what to check after |
| Troubleshooting | Messages and fixes (sketch was not made, wrong workbook, Trend Micro, missing material) |
| Glossary | Every term in the catalogue, searchable |

Help text uses the same catalogue, so a term's hover text and its glossary entry never disagree.

## 8. Inputs that cannot break

- Parts are records: Add part, Duplicate, Delete (with undo). No rows to delete by mistake.
- Every input has a unit, a type and a range. Wrong entries are marked at once with the reason.
- Required inputs for the current step are marked; optional ones say so ("optional: only for the two-resistor model").
- Undo and redo for every edit in the session.

## 9. Export

| Export | Content | How |
|---|---|---|
| **Report (PDF)** | Cover (board, date, run), inputs, assumptions, hand check, simulation results, hand vs simulation, pass or fail per part, plot picture, open issues | A print layout of the page; the browser's Save as PDF |
| **Workbook (`.xlsx`)** | A new workbook made by the page: Summary, Parts, Bodies, Triage, Rail budget, Bare board network, Spreader, Sim setup, Results, Changes, Method. Values only, units in headers, a note on every header (from the catalogue). The Method tab writes out every formula. | ExcelJS, built into the page |
| **Project archive (`.zip`)** | The three JSON files, the log, build reports, result pictures | JSZip, built into the page |

The user picks one run or all runs for the report and the workbook.

## 10. Macro changes

| Change | Where |
|---|---|
| JSON reader and writer in pure VBA (no install) | new module `PTJson.bas` |
| Pick and remember the project folder; log it | `PTCommon.bas` (replaces `PT_PickWorkbook`) |
| Read inputs from `project.json` `forSolidWorks` instead of named cells | every macro |
| Write their outputs to `cad.json` or `results.json` (temp file, then rename) | every macro |
| Change records go to `cad.json` `changes[]`; the "refuse to build while unchecked" rule reads the ticks from `project.json` `checks` | Import STEP, Build Board, Build Spreader, Sync CAD |
| SetupWorkbook is retired; the page creates new projects | |
| The add-in port does the same with .NET's built-in JSON serializer | `addin/src` |

Per macro:

| Macro | Reads | Writes |
|---|---|---|
| Import STEP | `settings`, `bodies` (to keep keys) | `cad.import`, `importPrevious`, `changes` |
| Build Board | `forSolidWorks.bodies`, PCB size, `settings` | `cad.build`, `changes` |
| Build Spreader | `forSolidWorks` spreader and covered chips | `cad.spreaderBuild`, `changes` |
| Sync CAD | `forSolidWorks` | `cad.asBuilt`, `changes` |
| Run Thermal Study | `forSolidWorks`, `simSetup` | `results.runs[]` (new run, inputs copy) |
| Read Results | the run name | `results.runs[]` (temperatures, picture) |
| Run Flow | `forSolidWorks` heat, ambient | `results.runs[]` |

## 11. Building it

- Source in `webapp/src/` (HTML, CSS, JavaScript modules, the catalogue). A small build script joins them into one
  file, `webapp/dist/PCB_Thermal.html`, because Chrome does not load separate script files from disk. Only the dist
  file ships.
- Libraries are put inside the file, at pinned versions: Preact and htm (interface, no build step), ExcelJS (export),
  JSZip (archive), SheetJS read only (importing an old workbook). Drawings and charts are plain SVG.
- **Old workbooks:** the page can read a filled workbook (read only) and make a `project.json` from it, so the RPi4
  workbook becomes the example project. The converter is `webapp/src/io/workbook_import.js` (done in W1); the page
  uses the same module.
- **Tests:** `webapp/tools/expected_from_workbook.mjs` reads every calculated value of `RaspberryPi4.xlsx` and
  `RPi4_fresh.xlsx` (bottom plate) as Excel last saved them and writes `tests/webapp/expected_rpi4.json` and
  `expected_rpi4_fresh.json`. Excel itself is not needed: the workbooks were last saved by Excel, so their cached values
  are Excel's results. The calculation module (`webapp/src/calc/`) must match them to 6 significant digits, texts exactly
  (`npm test`). This runs for every change to a formula.

## 12. Milestones

| # | Milestone | Done when |
|---|---|---|
| W1 | File formats: JSON schemas for the three files, with the RPi4 example converted from the workbook | `project.json`, `cad.json`, `results.json` for RPi4 exist and validate. **Done 2026-10-18** (`webapp/schema/`, `example/RaspberryPi4/*.json`, `npm test` in `webapp/`) |
| W2 | Calculation module: register, triage, rail budget, PCB stack, bare board network, spreader checks, thermal solution top and bottom, sim setup k values | matches `expected_rpi4.json`. **Done 2026-10-18** (`webapp/src/calc/`; about 4000 values of both RPi4 workbooks match; `forSolidWorks` is now computed by it) |
| W3 | Page shell: folder picking, save and reload, tabs with unlock rules and reasons, undo | RPi4 example opens, edits save, a locked tab names its reason |
| W4 | Term catalogue and Explain panel: hover text, abbreviations, formulas with numbers, trace | every value on every tab has a hover text and, if calculated, a formula |
| W5 | Tabs 1 to 5 complete (parts to bare board hand check), Learn mode texts | a new user can fill a board from a datasheet list without the workbook |
| W6 | Macros on JSON: `PTJson.bas`, project folder, Import STEP and Build Board read and write JSON | RPi4: import and build from the page give the same model as from the workbook |
| W7 | Remaining macros on JSON, tabs 6 to 10, SolidWorks cards and file watching | full RPi4 run from the page: build, spreader, study, results |
| W8 | Help page and glossary | every catalogue term in the glossary; each tab has its help section |
| W9 | Exports: PDF report, workbook, archive | RPi4 report and workbook made and checked by the user |

The workbook and the current macros keep working until W7 is done.

## 13. Decisions (answered 2026-10-18)

1. **File access:** Chrome or Edge on the work PC can save local files. The page saves straight into the project
   folder; open and download stays only as a fallback for other browsers.
2. **Add-in:** the VBA macros move to JSON first. The add-in follows only if it is needed later.
3. **One board per project folder.** Each board folder stays self-contained, so a later multi-board assembly file can
   point to several board folders without changing the board format.
4. **Units: fixed, with input helpers.** Values are stored and shown in mm, C, W, W/(m K) and C/W. Some fields accept
   other units as typed text and convert on entry: copper thickness in oz (1 oz = 0.035 mm) or um, lengths in mil,
   power in mW, temperature differences in K, resistance in K/W. The hover text of each field lists what it accepts.
5. **Theme:** dark by default, with a switch to light (section 4.1).
