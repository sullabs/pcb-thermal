# Ansys Icepak route (branch `ansys-aedt`)

Same goal as the SolidWorks route: a PCB to a thermal model. Here the board comes from Altium as an Ansys EDB
(`.aedb`), and a script builds the Icepak project through PyAEDT. Scope for now: **everything up to a ready study.
No solving.** Owners of every piece of information are in [SOURCES.md](SOURCES.md).

## What it is for

Get from an Altium board to a first Icepak result set-up fast, then change the few things a mechanical engineer
changes (power, ambient, board position, air, mesh) from the page, and do everything else in AEDT. The page is not
meant to replace AEDT. The PCB designer owns power and rails; here you only type the parts that matter.
No solve, no spreader, no other cooling solution for now.

## Steps

| # | Who | What |
|---|---|---|
| 1 | You | Altium: File > Export > Ansys EDB. Put the exported `<board>.edb` (or `<board>.aedb`) in a board folder anywhere; a `.edb` is copied to `.aedb` on read. Also File > Export > STEP 3D into the same folder: part heights and body sizes then come from it |
| 2 | You | Double-click `ansys\start.bat` (or drag a board folder onto it). It starts AEDT if needed and the page server; the browser opens http://localhost:8765 on the folder opened last. **Open folder…** at the top picks any folder on the PC (Windows folder dialog); the list above it holds the folders opened before. Press **Read the board** (once) |
| 3 | You | "What the export is missing" lists the gaps. Fill them: heights and body sizes of the parts that matter, tall parts without power, material values. Then ambient, board position, air, power (or the designer's CSV: RefDes, W, height, body X, body Y), mesh. Press **Build** |
| 4 | You | Look at the project in AEDT. Edit there if you want. **Read AEDT** shows what is there |
| 5 | You | Change values on the page, press **Apply changes** (about 40 s), **Generate mesh** to see the mesh without solving |

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

### When something is missing

| Missing | What happens |
|---|---|
| Power | No heat source for that part; warning if none has power |
| Height (EDB gives a placeholder under 0.2 mm; Altium never exports heights) | From the board's STEP if the folder has one, else 0.2 mm and a warning to type the real one |
| Body size (the EDB has only the pad outline) | From the STEP if there is one, else the pad outline; overlapping blocks are cut back to meet. Type Body X, Y (board axes) |
| Part not in the STEP (no 3D model, e.g. test points) | As if there were no STEP for that part. Typed values always win over the STEP |
| Material k, density, heat capacity (Altium exports 0) | Copper, solder or FR-4 stand-ins, shown grey on the page; typed values are sent on Apply |
| Height but no power | A block without heat: it shapes the air flow and spreads heat |
| Only one of Rjc, Rjb | Plain block, warning; no junction temperature |
| Rjc and Rjb both | Junction model; if AEDT refuses it, plain block and a warning |
| Tj max or Rja | No hand check or margin line, nothing else changes |
| Part not in the EDB | Error for that part, the others are still built |
| Part without position or size in the EDB | Listed as skipped in `board.json`, warning, no heat source possible |
| `study.json` | Defaults, no heat sources |
| AEDT closed (read) | Last known values kept, page says AEDT is not running |

**Do not run another PyAEDT script against the same AEDT while the page is open**: two clients attaching at the same
time made one of them hang. The page's own jobs never overlap, and a hung job is killed after 60 s (read) or 15 min.

## Files

| File | Role |
|---|---|
| `page.html` | The page. No build step |
| `serve.py` | Local server for the page (stdlib only). Reads/writes the folder files, starts the scripts one at a time. Folders opened before: `%APPDATA%\pcb-thermal\recent.json` |
| `read_board.py` | EDB -> `board.json` (outline, stackup, parts with position, size, height). PyEDB only, about 7 s |
| `step_parts.py` | Board STEP -> height and body size per part (`stepHMm`, `stepXMm`, `stepYMm` in `board.json`). Standard library only, under a second. Works on whole components (the children of the board assembly, with all their sub-parts), never on one sub-part: matches a component named like the RefDes, else the component centred on the pads. Height is from the board face the part sits on (standoff included) |
| `setup_study.py` | `build`, `update`, `read`, `mesh` modes. Writes `state.json` (what AEDT holds) and `checks.json` |
| `run.ps1` | Runs a script with Ansys' Python and `ansys\.venv` packages (the venv launcher itself is broken) |
| `start.bat` | Starts AEDT (gRPC port 50051) if it is not running, then `serve.py`. If the page already runs, it just opens it |
| `make_offline.ps1` | Writes `dist\pcb-thermal` (+ zip) to install on a PC without internet |

Folder contents: `<board>.aedb` (you), `board.json`, `study.json` (page), `state.json`, `checks.json`, `<board>_icepak.aedt`.
`study.json`: `edb`, `ambientC`, `orientation` (flat-up, flat-down, vertical-long, vertical-short), `marginMm`,
`marginAboveMm`, `radiation`, `pcbResolution`, `mesh {mode auto|manual, level, maxMm, gapElems, edgeElems}`,
`parts[{refDes, powerW?, heightMm?, bodyXMm?, bodyYMm?, meshMm?, rjcCW?, rjbCW?, rjaCW?, tjMaxC?}]`,
`materials {name: {kWmK?, rhoKgM3?, cpJkgK?}}`.

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
- [x] Solve button and results card (temperatures per part, board hot spot, convergence, energy balance)
- [ ] Radiation model name for S2S (the page uses discrete ordinates); TIM and spreader (with a cooling solution), per-surface emissivity
- [ ] Mesh cell count in the page (Generate mesh reports time only)
- [ ] Bring over the hand calculations, thermal network analysis and network diagram from the SolidWorks route (`../PCB_Thermal.html`): the user wants them on this page too
