# Flow Simulation: plan to make the Flow run trustworthy

Goal: a Flow run that the page can check and believe on its own. Today RunFlow builds and solves the project, but
nothing says whether the result is meshed, converged, energy balanced, or comparable with the hand check and the
Thermal study. This plan fixes that in phases F0 to F8. Each phase works through the three layers the same way:

1. **Schema** (`webapp/schema/`): new fields first.
2. **Page** (`webapp/src/calc/`, `webapp/src/page/`, `webapp/src/terms/`): inputs on tab 8, checks on tab 9.
3. **Macro** (`src/RunFlow.bas`, shared code in `PTCommon.bas` / `PTJson.bas`): reads `project.json`, writes `results.json`.
4. **Tests**: `npm test` in `webapp/`, then a live RPi4 run with the acceptance checks listed in each phase.
5. **Package**: copy the rebuilt `PCB_Thermal.html`, the changed `.bas` files and `PCBThermal_Flow.fwp` to branch
   `webpage`, rebuild `PCBThermal.swp` (Update Macros), commit both branches (CLAUDE.md rule).

Work happens on `addin-mvp` (worktree `thermal_wt_addin`). This file goes to `docs/FLOW_PLAN.md` there and gets a row
in `PLAN.md` section 1. The add-in port `addin/src/Macros/RunFlow.vb` is already behind `RunFlow.bas` (it still expects a
hand-made project); it is out of scope here and gets ported once F1 to F5 are done.

Rule for all phases: the page owns the inputs (`project.json`), the macro only writes `results.json` and the report.
A Flow input the macro needs goes into `forSolidWorks` (computed by the page) or `simSetup.flow` (typed by the user).

---

## What we already know (from the RPi4 runs)

- `CAD\<n>\<n>.info.json` (written by the Flow solver) has everything about the solve: `finished`, cell counts
  (`cells_total`, `cells_solid`, ...), `iteration`, `travel`, and per goal `value`, `unit`, `progress`, `criteria`,
  `delta`. `calculation_warnings.txt` next to it lists unresolved sources and invalid goals. The macro can read both
  with `PTJson` and plain text.
- The service goals give the energy balance. On run FL_20261001_0753: `Serv Heatf` = 2.817 W (convection) and
  `Radiation Energy Delta` = 3.604 W (radiation), together 6.42 W. Heat on the meshed sources was 6.42 W (6.67 W minus
  L1 and L2, which were not meshed). So the balance closes, and it also shows the missing 0.26 W. To be confirmed on the
  next run that `Serv Heatf` is always the convective heat leaving the solids.
- Bug: RunFlow subtracts 273.15 from every goal, also from W, Pa and kg/s goals (`Serv Heatf = -270.333` in the reports).
- Default goal criteria came out at about 2.3 K with 1 travel. That is too loose to compare runs within 1 to 2 C.
- Mesh study 2026-10-19 (HOWTO_M7): local mesh on chips, TIMs and spreader moved chips by 0.6 C; level 5 moved them by
  2.8 C. Small parts (L1, L2) are not meshed at level 2 or 3.
- Known API pieces (flow_api_notes.md): local mesh = feature 14 with parameters 68 and 69; `GlobalMeshSettings`;
  boundary conditions = feature 2. Everything else below marked **spike** must be checked in `cfwapi.chm` first.

---

## F0. API spike (one session, no user-facing change)

**Status: done 2026-10-19.** Answers in docs/flow_api_notes.md (section "F0 spike"). No API for user materials, emissivity values or goal criteria: fallbacks chosen there.

Check in `cfwapi.chm` and on a scratch project, write the answers into `docs/flow_api_notes.md`:

| # | Question | Needed by |
|---|---|---|
| S1 | How to get the project folder (`CAD\<n>\`) from the `FDAProject` object. Fallback: newest folder under the assembly folder whose `<n>.info.json` has our `project_name`. | F1 |
| S2 | Global mesh level and minimum gap / wall thickness parameters in `GlobalMeshSettings`; local mesh feature 14 parameters (solid refinement, refining cells, channels). | F2 |
| S3 | Goal criteria and finish conditions (min travels, max iterations, criteria per goal). | F3 |
| S4 | Can the API make a user material (k in-plane, k through, density, cp), or only pick an existing database path? | F4, F5 |
| S5 | Radiative surface feature (type id, emissivity parameter) on faces or components. | F5 |
| S6 | Contact resistance feature (type id, value parameter, faces). | F5 |
| S7 | Gravity vector and ambient pressure in general settings / initial conditions; external flow velocity. | F6 |

Done when each row has a type id or parameter id, or a "not possible" with the fallback chosen.

---

## F1. Run health: converged, meshed, energy balanced (do first)

**Status: done 2026-10-19 (schema, page, macro, tests in webapp/tools/test_flow.mjs). Live acceptance on RPi4 not run yet: needs SolidWorks.**

**Problem.** A run is stored as "read" even when sources were not meshed, goals came back 0 K, or the solve stopped
after 1 travel. The 484 C runs went into the results unflagged. Flow runs have no `energyBalancePct`.

**results.json** (schema `run`):
- `energyBalancePct`: already in the schema, now also written by Flow runs:
  `|heat applied - (Serv Heatf + Radiation Energy Delta)| / heat applied x 100`.
- New `solver` object: `finished` (bool), `iterations`, `travels`, `cellsTotal`, `cellsSolid`, `cellsFluid`.
- New `goalsHealth` per goal name: `progressPct`, `deltaC`, `criteriaC` (converted from K only for temperature goals).
- New `warnings` (array of text, from `calculation_warnings.txt`) and `unresolved` (RefDes whose source or goal failed).
- New `heatAppliedW` (sum of meshed sources) next to `heatRequestedW` (sum of `forSolidWorks.heat`).

**Macro (`RunFlow.bas`).**
- After solve (or on "read last run"), find the project folder (S1), read `<n>.info.json` and `calculation_warnings.txt`.
- Map "Unresolved volume source: <name>" back to its RefDes; drop that RefDes's goals and list it in `unresolved`.
- Convert only goals with `unit` = temperature to C. Fix the report lines for the other goals (W, Pa, kg/s).
- Status: `read` when finished, no warnings and balance within the limit; else still write the run, with `status`
  `read` and the reasons in `notes`, so the page can show why. (No new status value: the page judges health.)

**Page.**
- `calc/runs.js`: `runHealth(run)` gives `ok` / `check` / `bad` with reasons: not finished, balance above
  `BALANCE_LIMIT_PCT` (2 %), any `unresolved`, any goal progress below 100 %, travels below 1.
- Tab 9: a health badge per run, the reasons in plain words ("L1 and L2 were not meshed: 0.26 W left out"), and the
  PASS/FAIL verdict greyed out while health is `bad`.
- Terms in `catalogue.js` for every new value; Learn text on tab 9 says check the badge before reading temperatures.

**Acceptance (RPi4).** Re-read FL_20261001_0753: badge `bad`, reasons name L1 and L2, balance about 3.8 % against the
requested 6.67 W. Report shows `Serv Heatf = 2.817 W`, not -270.333.

---

## F2. Mesh control and mesh check

**Status: done 2026-10-19 (schema simSetup.flow, local mesh features, global level, mesh check pair, tab 8 card, tab 9 comparison, tests). Change from the plan: a part thinner than the estimated smallest cell gives a warning on tab 8 and in the report, not a refusal, because the cell estimate (calibrated on two RPi4 runs) is too rough to block a run; F1 health shows what was really meshed. Live acceptance not run yet.**

**Problem.** The automatic mesh leaves small parts and thin TIMs without solid cells, and nothing tells the user how
much the answer depends on the mesh.

**project.json** (`simSetup.flow`, new object, page owned):
- `meshLevel` (1 to 7, default 4), `localMesh` (bool, default true), `solidRefine` (default 3),
  `meshCheckLimitC` (default 1.0).

**Macro.**
- Set the global level (S2). With `localMesh`, own one feature `Local mesh <RefDes>` per part with heat, one
  `Local mesh TIM` and one `Local mesh Spreader`, same fixed-name pattern as the goals (`OwnFeature`), deleted when the
  component is gone. A local mesh the user made by hand is kept.
- Before solving, refuse with a clear message if a part with heat is thinner than the smallest cell the settings can
  give (estimate from domain size and level; exact rule from S2).
- New button choice "Mesh check": solve twice (level N and N+1, local mesh kept), store both runs with a shared
  `meshCheck: { pairWith, level }` field.

**results.json.** `mesh` object (already there for Thermal) gets `level` and `localMesh` for Flow runs; `meshCheck` as above.

**Page.** Tab 8 Flow card: mesh level and local mesh inputs, with Learn text (cells vs time, from the 2026-10-19 study).
Tab 9: for a mesh check pair, the change per chip; `ok` when every chip moved less than `meshCheckLimitC`.

**Acceptance.** RPi4 with local mesh: L1 and L2 resolved (no warnings), F1 badge `ok`. Mesh check pair at levels 4/5
moves every chip by less than 1 C, or the page says which chip did not.

---

## F3. Convergence settings

**Status: done 2026-10-19 (minTravels, goalCriteriaC; finish conditions set as all satisfied + goals convergence + manual travels; goals above the criterion flagged in the report and on tab 9). The API cannot set a goal's criterion (F0, S3). Live acceptance not run yet.**

**Problem.** Goals stop at about 2.3 K criteria after 1 travel.

**project.json** (`simSetup.flow`): `minTravels` (default 2), `goalCriteriaC` (default 0.2).

**Macro.** Set finish conditions and the criteria of the goals it owns (S3). If the API cannot set criteria, set min
travels only and rely on F1 to flag `delta` above `goalCriteriaC`.

**Page.** Inputs on tab 8 with Learn text; F1 health uses `goalCriteriaC` for the "not converged" reason.

**Acceptance.** RPi4 run reaches at least 2 travels, every chip goal has delta below 0.2 K.

---

## F4. Junction temperature on Flow runs

**Status: done 2026-10-19. The API cannot make a material (F0, S4), so the page names one user material per k step (10 %, 24 per decade; Rjc within 5 %) in forSolidWorks.flow and lists them on tab 8; the user makes each once in the Engineering Database (no shipped database file: its format is not documented). Two-resistor chips use the three layer bodies Run Thermal Study makes, with source and goals on the core body. Every part gets tjBasis; tab 9 marks package parts. Live acceptance not run yet.**

**Problem.** RunFlow stores the peak of a generic plastic block as Sim Tj. The Thermal study uses a block calibrated to
Rjc (`calc/sim.js`, k z = H / (Rjc x A)) or the M8 two-resistor body, so its peak means Tj and the Flow peak does not.

**Plan.** Make Flow use the same chip model as Thermal, so the numbers mean the same thing:
- **Block chips**: one material per chip with the k z and k xy that `forSolidWorks.chips` already carries (needs S4).
  Fallback if S4 says no: ship a user material database `PCBThermal_Materials` with anisotropic materials in 10 % steps
  of k (0.2 to 400 W/m K), and the macro picks the nearest step; the page shows the step used and the Rjc it gives.
- **Two-resistor chips (M8)**: the part already has three bodies. Put the source and the goals on the core body;
  its peak is Tj.
- **Chips with no Rjc**: keep the generic material, and mark the value as package temperature, not Tj.

**results.json.** Per part `tjBasis`: `calibrated block`, `two-resistor`, or `package` (no Rjc).

**Page.** Tab 9 shows "Tj" only for the first two; "package, no Rjc" otherwise, and the Tj max check uses it with a
warning. Same rule for Thermal runs, so both tables read alike.

**Acceptance.** RPi4: U1 to U6 have Rjc, so `tjBasis` = calibrated block; Flow and Thermal U1 within a few C on the
same board and spreader.

---

## F5. Materials, radiation and contacts from the page

**Status: done 2026-10-19. PCB, TIM and spreader get user materials from the page k; radiation is one owned radiative surface per component group (PCB, parts, spreader) named after the emissivity; contacts are integral values on the PCB (parts to PCB) and the spreader (TIM to spreader); chip to TIM is folded into the TIM material. Face-level surfaces and contacts were not done: they need face selection the API help does not cover well enough to write offline. Tab 8 lists every value and simplification; docs/HOWTO_FLOW_MATERIALS.md says how to make the materials once. Live acceptance not run yet.**

**Problem.** PCB, TIM and spreader materials are fixed library picks; emissivity is 0.9 everywhere; no contact resistance.
The page already has all of these values.

**From `forSolidWorks` and `simSetup` (no new inputs):**
- PCB: `forSolidWorks.pcb.kxy` / `kz` from the stack, instead of `PCB 4/8/12-layers` (S4 or the material steps of F4).
- TIM: the TIM k per chip from `forSolidWorks.plates.*.chips` instead of GR25A.
- Spreader: the plate k instead of aluminium/copper by threshold.
- Radiation: one owned radiative surface feature per group in `simSetup.surfaces` (spreader top, spreader other,
  PCB top, PCB bottom, parts) with its emissivity (S5).
- Contacts: `simSetup.contacts` (part to PCB, chip to TIM, TIM to spreader) as owned contact resistance features (S6).

**Macro.** Same owned-feature pattern; each feature it cannot make is listed in `notes` and the report.

**Page.** Tab 8 Flow card lists, per item, the value that will go to Flow and where it came from. Tab 9 `inputChanges`
already covers these fields, so an old run says what changed.

**Acceptance.** RPi4 report shows PCB k from the stack, the spreader emissivity from tab 8 (not 0.9), and the contact
features. Flow and Thermal differ by less than with the old fixed materials (compare with the 2026-10-19 numbers).

---

## F6. Environment: orientation, pressure, forced air

**Problem.** Only a flat board in still air at sea level.

**project.json** (`simSetup.flow.environment`):
- `orientation`: `flat, parts up` (default) / `flat, parts down` / `vertical, long side up` / `vertical, short side up`.
- `pressureKPa` (default 101.325) or `altitudeM`.
- `airSpeedMs` and `airDirection` (+X, -X, +Y, -Y, +Z); 0 = natural convection.

**Macro.** Gravity vector from orientation, pressure and velocity in the initial and ambient conditions (S7). Domain
margins follow gravity: 1.5 x the board size on the side the warm air goes, 0.5 x elsewhere; with forced air, 2 x
downstream. User features (fan, lid) stay untouched as today.

**Page.** Tab 8 inputs with a small drawing of the board and the gravity arrow. The hand check gets a note when the
orientation or air speed does not match the h it assumes (the page's h values are for natural convection).

**Acceptance.** RPi4 vertical, long side up: chips cooler than flat by a plausible few C, domain grows upward along +Y.

---

## F7. Compare and validate (hand, Thermal, Flow, measured)

**Problem.** No single view of all estimates, and no place for a real measurement.

**project.json** (page owned, new top-level `measurements` array):
`{ name, date, ambientC, note, parts: { <RefDes>: { tC, method: "thermocouple" | "IR" | "internal sensor", where } } }`.

**Page.**
- Tab 9 comparison table per RefDes: hand check, newest healthy Thermal run, newest healthy Flow run, measurement,
  and the gaps. Only runs with F1 health `ok` fill the columns by default.
- Tab 9 "Add a measurement" form; Learn text on where to place thermocouples and how to read an IR camera on a glossy chip.
- When Flow and measurement differ by more than a limit, a list of what usually explains it (emissivity, contact R,
  Rjc, missing heat), ranked from the F5 inputs.
- Report and workbook export get the table.

**Macro.** None.

**Acceptance.** RPi4: enter the `RPi4_manual_test` values as a measurement; the table fills all four columns.

---

## F8. Small fixes

- **Heat with no body** (J1, 0.05 W): page check on tab 6 before Build Board: "J1 has heat but no body in the model.
  Add a body, or set its heat to 0 on purpose." Stored in `project.json checks`, and the macros already refuse while
  checks are unticked.
- Goal-name matching: `FindRef` takes the first word that is a RefDes; a goal "U1 to U2 delta" would map to U1. Match
  only the owned names (`<RefDes> max/avg`) plus user goals that start with a RefDes.
- Report: write the cell counts and travels from F1 at the top.

---

## Order and size

| Phase | Depends on | Size | Why this order | Status |
|---|---|---|---|---|
| F0 spike | none | 1 session | every later phase needs the ids | done 2026-10-19 |
| F1 health | S1 | medium | without it no Flow number can be trusted | done 2026-10-19 |
| F2 mesh | S2, F1 | medium | removes the unresolved parts F1 will show | done 2026-10-19 |
| F3 convergence | S3, F1 | small | | done 2026-10-19 |
| F4 Tj | S4 | medium | changes the number compared with Tj max | done 2026-10-19 |
| F5 materials | S4 to S6 | medium | | done 2026-10-19 |
| F8 fixes | none | small | can go any time, good first warm-up | open |
| F6 environment | S7 | medium | | open |
| F7 compare | F1, F4 | medium | page only | open |

Not in this plan: transient runs and power profiles, enclosures and fans built by the macro, the add-in port.
