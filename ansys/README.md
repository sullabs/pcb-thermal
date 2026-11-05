# Ansys Icepak route (branch `ansys-aedt`)

Same goal as the SolidWorks route: a PCB to a thermal model. Here the board comes from Altium as an Ansys EDB
(`.aedb`), and a script builds the Icepak project through PyAEDT. Scope for now: **everything up to a ready study.
No solving.** Owners of every piece of information are in [SOURCES.md](SOURCES.md).

## What it is for

Get from an Altium board to a first Icepak result set-up fast, then change the few things a mechanical engineer
changes (power, ambient, board position, air, mesh) from the page, and do everything else in AEDT. The page is not
meant to replace AEDT. The PCB designer owns power and rails; here you only type the parts that matter.
A spreader plate with TIMs (step 4, Thermal solution) is drawn in SolidWorks from the page and comes back to it; see below.

## Steps

| # | Who | What |
|---|---|---|
| 1 | You | Altium: File > Export > Ansys EDB. Put the exported `<board>.edb` (or `<board>.aedb`) in a board folder anywhere; a `.edb` is copied to `.aedb` on read. Also File > Export > STEP 3D into the same folder: part heights and body sizes then come from it |
| 2 | You | Double-click `ansys\start.bat` (or drag a board folder onto it). It starts AEDT if needed and the page server; the browser opens http://localhost:8765 on the folder opened last. **Open folder…** under the steps picks any folder on the PC (Windows folder dialog); the list above it holds the folders opened before. Press **Read the board** (once) |
| 3 | You | "What the export is missing" lists the gaps. Fill them: heights and body sizes of the parts that matter, tall parts without power, material values. Then ambient, board position, air, power (or the designer's CSV: RefDes, W, height, body X, body Y), mesh. Press **Build** |
| 4 | You | Look at the project in AEDT. Edit there if you want. **Read AEDT** shows what is there |
| 5 | You | Change values on the page, press **Apply changes** (about 40 s), **Generate mesh** to see the mesh without solving |
| 6 | You | Optional, a spreader: step 4 **Thermal solution** on, tick the parts, **Build Solution** in SolidWorks, finish the plate there, **Send Solution**, then **Apply changes** (below) |

**AEDT is the source of truth.** The page asks the server every 15 s to read AEDT (the `read` mode: it never starts
AEDT and never opens a project). Whatever changes in AEDT (ambient, board position, air, mesh, power, height, Rjc,
Rjb, parts added or removed) shows up in the page fields by itself. Fields you edit on the page turn yellow and stay
yours until you press **Apply** (it says how many changes). Apply only touches what the page owns: heat-source
blocks and their power, the air box and its opening, ambient and gravity, the setup values shown, the global mesh and
`Mesh_<RefDes>` regions. Everything else you made in AEDT is left alone and listed under "Other thermal items".
Build replaces the whole project and asks first. If AEDT is closed, the page shows the last values it read and says so.

Rja and Tj max are page-only (kept in `study.json`): the page shows Tj by Rja = ambient + power x Rja as a hand check.
Rjc and Rjb together make a two-resistor junction model in AEDT (`P_<RefDes>`, type Network), the block then sits on
the board surface so it can touch the PCB.

### Thermal solution (spreader plate)

Step 4 of the page, with its own On switch. You tick the parts the plate touches (only parts in Parts). Top-side parts
make the top plate, bottom-side parts the bottom plate, so one or two plates follow. Per plate: material (aluminium 167,
copper 390, or a typed k), thickness, margin, **Lift** and minimum pedestal; per part: TIM thickness and k. The page
works out the first plate from the STEP bodies: underside = highest ticked part top + its TIM + Lift, a rectangle around
the ticked parts plus the margin (inside the board outline), and per part a pedestal down to its TIM (none under the
minimum pedestal: the TIM fills that gap). It warns about parts not ticked that reach above the underside (with the Lift
that clears them by 0.5 mm), ticked parts with no STEP body, and typed heights or body sizes that differ from the STEP
(SolidWorks builds on the STEP bodies, AEDT on the typed values, so the TIM may then not touch the block).
It saves all of it in `study.json` `solution`, with `forSolidWorks`: the plates in the board STEP's frame.

In SolidWorks (PCB Thermal tab, after **Update Macros** and `SetupToolbar.cmd` once):
- **Build Solution** (`macros\BuildSolution.bas`, the work in `BuildSpreader.bas` `RunAnsys`): pick `study.json` of the
  board folder. It opens the board STEP and saves it as `solidworks\<board>_solution.sldasm` in the board folder (first run
  only), then makes `Spreader_Top` / `Spreader_Bottom` (an `Outline` sketch of the page's rectangle on the Front Plane,
  the plate, `Ped_<RefDes>` pedestals; faces named Top, Underside, `Contact_<RefDes>`) and one `TIM_<RefDes>` part per
  ticked part. Run it again after a change on the page: the plate moves, pedestals and TIMs follow, and the Outline sketch
  and every feature you added (holes, cuts, fillets) stay. It writes `solution_cad.json` `build`.
- Finish the plates: mounting holes, shape, cuts around parts in the way. Only move a plate, never turn it.
- **Send Solution** (`macros\SendSolution.bas`) changes no CAD file. It writes each plate as `solution_<side>.step`
  (part frame, with `offsetMm`) and `solution_cad.json` `sentAt`, `plates[]` (volume, area of all faces, box, underside
  and top Z, thickness, area of the Top, Underside and `Contact_<RefDes>` faces) and `tims[]` (RefDes, side, box).

Back on the page, step 4 shows what SolidWorks sent, and says when the page now wants another underside or thickness
(Build and Send again). The hand check puts each plate in its network as one more node: a ticked part's top path is then
R junction to top + TIM + pedestal + Lee spreading in its share of the plate (by power), and the plate loses heat from all
its faces with h plate and the plate emissivity (Settings, default 8 W/m²K and 0.85). It uses SolidWorks' areas once
sent, else the page's plan. **Apply** imports the plates into the same Icepak design as `SOL_Spreader_<Side>` (again
only when the STEP or its place changed) and the TIMs as boxes `SOL_TIM_<RefDes>`, with their own materials
(`sol_plate_<side>`, `sol_tim_<refdes>`: the page's k), placed by `board.json` `stepFrame`. It warns when a TIM does not
sit on its part's block (over 0.02 mm). The switches: off takes the objects out of the model (out of the mesh and the
solve), on puts them back; they are never deleted by the switch. Solve adds the plates' hottest and average temperature
to `results.json` `plates`. A thin TIM may need a finer mesh: use **Check the mesh**.

### When something is missing

| Missing | What happens |
|---|---|
| Power | No heat source for that part; warning if none has power |
| Height (EDB gives a placeholder under 0.2 mm; Altium never exports heights) | From the board's STEP if the folder has one, else 0.2 mm and a warning to type the real one |
| Body size (the EDB has only the pad outline) | From the STEP if there is one, else the pad outline, both along the part's own axes (a 15 x 15 part turned 45 deg stays 15 x 15; the angle comes from the EDB pads, Altium turns pins not parts); overlapping blocks are cut back to meet. Type Body X, Y (board axes) |
| Part not in the STEP (no 3D model, e.g. test points) | As if there were no STEP for that part. Typed values always win over the STEP |
| Material k, density, heat capacity (Altium exports 0) | Copper, solder or FR-4 stand-ins, shown grey on the page; typed values are sent on Apply |
| Height but no power | A block without heat: it shapes the air flow and spreads heat |
| Only one of Rjc, Rjb | Plain block, warning; no junction temperature |
| Rjc and Rjb both | Junction model; if AEDT refuses it, plain block and a warning |
| Local mesh on a junction model | Not sent, warning: the block is not meshed inside, and a mesh region on it stops the solver. The value stays on the page |
| Pd from the circuit gives below 0 W, or efficiency over 1 (85 typed for 0.85) | Pd not calculated, warning; the typed Pd is used |
| Tj max or Rja | No hand check or margin line, nothing else changes |
| Part not in the EDB | Error for that part, the others are still built |
| Part without position or size in the EDB | Listed as skipped in `board.json`, warning, no heat source possible |
| `study.json` | Defaults, no heat sources |
| AEDT closed (read) | Last known values kept, page says AEDT is not running |
| AEDT running without the page's gRPC port (started outside `start.bat`) | The read leaves it alone and says so; it never starts a second AEDT |
| An edit made in the AEDT window and not saved | Blocks, power, ambient and air show at the next read; mesh and setup settings only once saved in AEDT (PyAEDT reads them from the saved file). Apply, Generate mesh and Solve save the project first, so they never undo such an edit |
| A typed material value cleared | Apply sets the export's value again, or the stand-in |
| `.busy` mark of another script (the SO-8 benchmark) | The job stops with an error instead of using AEDT at the same time |
| A project with the same name open from another folder | Error, nothing changed: AEDT finds a project by name, so the job would change that copy. Close it in AEDT |

**Do not run another PyAEDT script against the same AEDT while the page is open**: two clients attaching at the same
time made one of them hang. The page's own jobs never overlap, and a hung job is killed after 60 s (read), 4 h (mesh),
12 h (solve), 24 h (mesh check) or 30 min (any other job).

## Files

| File | Role |
|---|---|
| `page.html` | The page. No build step |
| `serve.py` | Local server for the page (stdlib only). Reads/writes the folder files, starts the scripts one at a time. Folders opened before: `%APPDATA%\pcb-thermal\recent.json` |
| `read_board.py` | EDB -> `board.json` (outline, stackup, parts with position, size, height). PyEDB only, about 7 s. `--step` reads only the STEP again into `board.json` (the page's **Read heights from the STEP**) |
| `step_parts.py` | Board STEP -> height and body size per part (`stepHMm`, `stepXMm`, `stepYMm`, `stepAngleDeg` in `board.json`; the body measured along the part's own axes: the pads' `angleDeg`, or the body's own angle when it sits clearly turned from that. `setup_study.py` builds the block turned by that angle). Standard library only, under a second. Works on whole components, never on one sub-part: the products on the first level under the STEP's top assembly, each with everything under it, matched by name (U1, C7, VR8; SolidWorks' `U1-1` occurrence names too) or else as the component centred on the pads. Height is from the board face the part sits on (standoff included) |
| `setup_study.py` | `build`, `update`, `read`, `mesh`, `solve`, `meshcheck` modes. Writes `state.json` (what AEDT holds) and `checks.json`; `solve` writes `results.json`; `meshcheck` solves with a finer mesh (one automatic level up, or every largest element and local mesh x 2/3), then with the page's mesh again, and writes both to `meshcheck.json` |
| `run.ps1` | Runs a script with Ansys' Python and `ansys\.venv` packages (the venv launcher itself is broken) |
| `start.bat` | Starts AEDT (gRPC port 50051) if it is not running, then `serve.py`. If the page already runs, it just opens it |
| `make_offline.ps1` | Writes `dist\pcb-thermal` (+ zip) to install on a PC without internet |
| `..\macros\BuildSolution.bas`, `SendSolution.bas` | SolidWorks side of the thermal solution (above). Need `PTCommon`, `PTJson` and `BuildSpreader` in `PCBThermal.swp` |

Folder contents: `<board>.aedb` (you), `board.json`, `study.json` (page), `state.json`, `checks.json`, `<board>_icepak.aedt`,
`results.json` (Solve), `meshcheck.json` (Check the mesh), `built.json` (when and with which PCB mapping Build ran),
`meshed.json` (the mesh signature of the last Generate mesh or solve), `applied.json` (the body sizes the last Build or
Apply made the blocks from: the page compares Body X and Y with it), `solution_cad.json` (Build and Send Solution),
`solution_<side>.step` (Send Solution), `solidworks\` (the SolidWorks assembly and parts of the solution),
`solution_applied.json` (which SolidWorks export the `SOL_` objects in AEDT were made from),
`runs\<yyyymmdd-hhmmss>.json` (every solve, mesh check solves too: `at`, `label`, `note` typed on the page, `settings` it ran with
(mesh, local mesh, iterations, flow and energy limits and radiation as AEDT held them, ambient, gravity, total Pd) and `result`
as in `results.json`; the page's **Runs** table compares them, and Remove deletes the file).
Solve health in `results.json` (and so in each run): `converged {ok, items[{name, value, limit, ok, history}], iterations,
totalIterations, continued, maxIterations, hitMax}` (`iterations` is this solve's; a solve with nothing changed goes on from
the last solution, then `continued` is true and `totalIterations` counts from its start; with no monitor history this solve's
share is not known and `iterations` and `continued` are null), `balance {inW, outW, convectionW, radiationW, errorPct, rawOutW, netMassKgS, refFixW}` (the solver counts air heat from 25 C, so at another ambient the leftover net mass flow x cp x (ambient - 25) is taken off convection; `rawOutW` is AEDT's own number), `monitors[{refDes, driftK, ok, read, history}]` (how far each heat source temperature still moved over
the last 10 % of this solve's iterations; over 0.1 K is not settled; `read` false: no history found, `ok` null with `read` true:
one iteration only), `profile {cells, nodes, memory, elapsed, cores, status}` (read from
the newest `*.profile` AEDT wrote in the project's `.aedtresults`). Histories are `[[iteration, value]]`, at most 200 points;
the page plots them and a report can reuse them. They come from the solver's own files (`DV*_S*_MON*.sd` named in
`.aedtresults\Icepak.asol`, read after a save, each monitor's file found by the monitor's `ID=` in the saved `.aedt`),
because AEDT's report API gives only the final value. A failed read is a warning only.
Solve also repairs the mesh of an older project before it starts: it removes a local mesh on a junction model and any
mesh sub-region box no region uses (both stop the solver), says so in the checks, and rereads `state.json`.
`study.json`: `edb`, `ambientC`, `orientation` (the direction gravity points in the board's axes: -Z flat parts up, +Z flat parts down, ±X or ±Y standing on that edge; the page sets it by turning the board about its X or Y axis in 90° steps; the old names flat-up, flat-down, vertical-long, vertical-short still load), `marginMm`,
`marginAboveMm`, `radiation`, `pcbResolution`, `iterations`, `flowLimit?`, `energyLimit?` (solver limits; left out keeps AEDT's), `mesh {mode auto|manual, level, maxMm (largest element X), maxYMm, maxZMm, gapElems, edgeElems, ratio}` (manual only; an older file with one `maxMm` uses it for X, Y and Z),
`parts[{refDes, powerW?, heightMm?, bodyXMm?, bodyYMm?, meshMm?, rjcCW?, rjbCW?, rjaCW?, tjMaxC?, cls?, vinV?, iinA?, voutV?, ioutA?, efficiency?, rOhm?, pdMeasuredW?}]`
(`powerW` is Pd, the heat given off; `cls` to `pdMeasuredW` are page only and give it as the SolidWorks register),
`materials {name: {kWmK?, rhoKgM3?, cpJkgK?}}`, `handCheck {hTop?, hBottom?, hParts?, emissivity?, surfaceRiseK?, coverage?, kxy?, kParts?, critMargin?, critDensity?, critShare?, watchMargin?, watchDensity?, watchShare?, matched? {hPcb, run, powerW, ambientC, orientation, boardMeanC}}`
(page only; empty means the SolidWorks template defaults 8, 5, 8 W/m²K, 0.9, 20 K, 0.6, from the stackup, 0.5, 15 K; `hPlate?`, `plateEps?` 8 W/m²K, 0.85),
`solution {on, plates {top|bottom: {on?, material?, kWmK?, tMm?, marginMm?, liftMm?, minPedMm?}}, parts {RefDes: {on?, timTMm?, timK?}},
forSolidWorks {stepFile, frame, computedAt, plates[{side, on, material, kWmK, tMm, dir, undersideZMm, outlineMm, chips[{refDes, boxMm, faceZMm, timTMm, timK, pedestalMm}]}]}}`
(`forSolidWorks` is worked out by the page on every save, in the STEP's frame). `board.json` `stepFrame {offXMm, offYMm, topZMm}`:
STEP point + offset = EDB point; topZMm is the board's top face in the STEP.

Timings on the sample: build about 1.5 min, apply with a change 40 to 80 s, generate mesh 15 s.

## Environment

- AEDT 2026 R1 (v261) here, 2025 R2 (v252) or later elsewhere; PyAEDT 1.9.0, Python 3.10 (AEDT ships its own with pip).
  Install root here: `E:\Program Files\ANSYS Inc\v261\AnsysEM`.
- `start.bat` and `run.ps1` use the newest AEDT installed (the highest `ANSYSEM_ROOTnnn` variable; `AEDT_RELEASE=252`
  picks one) and pass its version to the scripts as `AEDT_VERSION` (`2025.2`). Only tested on 2026 R1 so far.
- They install `requirements.txt` into `ansys\.venv\Lib\site-packages` (`pip install --target`) when PyAEDT is not
  there: from `offline\wheels` if that folder exists, else from the internet. They check only that PyAEDT is
  present, not its version: after changing the pin, delete `ansys\.venv` to reinstall.
- New PC without internet: `make_offline.ps1` writes `dist\pcb-thermal\` and `dist\pcb-thermal.zip` (about 60 MB):
  these files, the sample board, `offline\wheels`, the Python 3.10.11 installer (only for an AEDT without
  `commonfiles\CPython`) and `INSTALL.txt`. Copy it over, then run `start.bat`.
- `requirements.txt` alone gives the same `board.json` as the fuller `.venv` here (PyEDB over gRPC, no pythonnet).
- Same on the work PC: AEDT installed, gRPC allowed on localhost (port 50051 here), then `start.bat`.

## Status

- [x] AEDT reachable from Python (attach and start both work)
- [x] Source of truth table
- [x] Sample EDB in `ansys/example/`
- [x] PyEDB reads the sample (stackup, components, nets, vias, copper area)
- [x] EDB -> AEDT -> Icepak PCB works from Python (about 45 s); see [SESSION_LOG.md](SESSION_LOG.md) for gotchas
- [x] `setup_study.py` build / update / read / mesh verified on the sample (power, orientation, air, mesh global and local)
- [x] Page for the mechanical engineer; "In AEDT now" card with differences and Take AEDT values
- [x] Real Altium exports: Bluetooth Sentinel and Spartan-IIE (AD examples), built and solved
- [x] Missing heights, body sizes and material values filled from the page (Altium exports none of them)
- [x] Rjc/Rjb junction model, Rja/Tj max hand check, AEDT as source of truth (auto read), missing-data fallbacks
- [x] Rjc without Rjb: a solid block of its own material `rjc_<refdes>`, k = H / (Rjc x L x W), so it conducts Rjc from
  top to bottom as the hand check's "Block, Rjc"; read back from AEDT as H / (k x volume / H). Rjb without Rjc stays on
  the page only (a junction model needs both); the hand check still uses it as the SolidWorks route does
- [x] Solve button and results card (temperatures per part, board hot spot, convergence, energy balance)
- [x] Thermal solution: spreader plates and TIMs from the page, drawn and finished in SolidWorks, back to the hand check
  and into AEDT. Run end to end on a copy of the Bluetooth sample (2026-11-05): Build Solution (plate, 1 pedestal, 2 TIMs),
  Send Solution, Apply, off/on, Solve: U4 49.8 to 44.3 C, U1 48.7 to 44.4 C, plate 44.0 C. The hand check gives the
  plate 32.6 C: it cools the whole plate to air and ignores the board heating the underside across the small gap
- [ ] Radiation model name for S2S (the page uses discrete ordinates); per-surface emissivity; a heat sink
- [ ] Mesh cell count in the page (Generate mesh reports time only)
- [x] Minimal page: Settings holds the hand check values, flag limits, air box (air around, extra above), cells in gaps
  and on edges, PCB mapping and hover help; the Settings button turns yellow when one of them is not yet in AEDT. The
  explanations are hover help on the step and table headings, not text on the page
- [x] Questions before a job: Build on a built project (Build again and overwrite, Open the built one, Cancel); Apply
  lists what cannot reach AEDT (Rjb without Rjc, Rjc/Rjb without Pd, parts with nothing to make, PCB mapping); Generate mesh
  with changes not applied, or when the mesh of these settings is already there. **Open in AEDT** opens the built project
  (mode `open`, starts AEDT if needed, changes nothing). Not yet run in AEDT
- [x] Check the mesh button (Mesh step): fine when the finer mesh moves no temperature (parts, board hottest, board average)
  by more than 1 K or 5 % of its rise, fair up to 3 K, else go finer. Warns when a run is not converged or the energy
  balance is off by more than 2 %, and when the mesh or power on the page changed since the check. Not yet run in AEDT
- [x] Hand check from the SolidWorks route (bare board network of `addin-mvp:webapp/src/calc/network.js`, Lee spreading,
  network diagram, formula per value) on the page; same numbers as the SolidWorks calc for the same inputs. Its inputs
  (h, emissivity, copper coverage, board k, part k) are page-only, saved in `study.json` under `handCheck`
- [x] Hand check scope: one board node gives a fair board **average**, never hot spots (Spartan-IIE: hand 30.4 C average
  against Icepak 33.9 C average and 49.4 C hottest). Every part with power gets a hand Tj, small ones too. A small part
  near hot neighbours can be warmer in Icepak than by hand, because one board node does not see its neighbours
- [x] Pd from the circuit as the SolidWorks register: class, Vin, Iin, Vout, Iout, efficiency, R, Pd measured per part
  (one load per part). The Pd it gives fills Power and goes to AEDT and the hand check; without it Pd is typed
- [x] Flag limits as the SolidWorks register: CRITICAL / Watch / Minor from the hand margin, power density over the body
  footprint and share of the total power. Shown in the hand check and on the Parts table; the CRITICAL margin sets Tight
- [x] "Match to the last Icepak run": fits the h of both PCB faces so the hand board average equals Icepak's. Allowed
  only from a converged run with a closed energy balance and the same power and ambient as the page. While matched the
  board average no longer checks that run (compare with textbook h first); a warning shows when power moves over 30 %,
  ambient over 5 K or the board position changes from the matched run. Saved as `handCheck.matched`
