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
| 1 | You | Altium: File > Export > Ansys EDB. Put `<board>.aedb` in the folder `ansys\example` (or any folder) |
| 2 | You | `.nsysun.ps1 serve.py [folder]`. The browser opens http://localhost:8765. Press **Read the board** (once) |
| 3 | You | Ambient, board position, air, heat sources (add parts, or load the designer's CSV: RefDes, W), mesh. Press **Build** |
| 4 | You | Look at the project in AEDT. Edit there if you want. **Read AEDT** shows what is there |
| 5 | You | Change values on the page, press **Apply changes** (about 40 s), **Generate mesh** to see the mesh without solving |

The page and AEDT can drift apart: the "In AEDT now" card marks the differences and offers **Take AEDT values**
(copy AEDT into the page) or **Apply changes** (copy the page into AEDT). Apply only touches what the page owns:
heat-source blocks and their power, the air box and its opening, ambient and gravity, the setup values shown,
the global mesh and `Mesh_<RefDes>` regions. Everything else you made in AEDT is left alone. Build replaces the
whole project and asks first.

## Files

| File | Role |
|---|---|
| `page.html` | The page. No build step |
| `serve.py` | Local server for the page (stdlib only). Reads/writes the folder files, starts the scripts one at a time |
| `read_board.py` | EDB -> `board.json` (outline, stackup, parts with position, size, height). PyEDB only, about 7 s |
| `setup_study.py` | `build`, `update`, `read`, `mesh` modes. Writes `state.json` (what AEDT holds) and `checks.json` |
| `run.ps1` | Runs a script with Ansys' Python and `ansys\.venv` packages (the venv launcher itself is broken) |

Folder contents: `<board>.aedb` (you), `board.json`, `study.json` (page), `state.json`, `checks.json`, `<board>_icepak.aedt`.
`study.json`: `edb`, `ambientC`, `orientation` (flat-up, flat-down, vertical-long, vertical-short), `marginMm`,
`marginAboveMm`, `radiation`, `pcbResolution`, `mesh {mode auto|manual, level, maxMm, gapElems, edgeElems}`,
`parts[{refDes, powerW, heightMm?, meshMm?}]`.

Timings on the sample: build about 1.5 min, apply with a change 40 to 80 s, generate mesh 15 s.

## Environment

- AEDT 2026 R1 (v261), PyAEDT 1.9.0, Python 3.10 (Ansys ships 3.10.19; use a venv from it, then `pip install pyaedt`).
  Install root here: `E:\Program Files\ANSYS Inc\v261\AnsysEM`.
- Same on the work PC: AEDT installed, gRPC allowed on localhost (port 50051 here), then the pip line above.

## Status

- [x] AEDT reachable from Python (attach and start both work)
- [x] Source of truth table
- [x] Sample EDB in `ansys/example/`
- [x] PyEDB reads the sample (stackup, components, nets, vias, copper area)
- [x] EDB -> AEDT -> Icepak PCB works from Python (about 45 s); see [SESSION_LOG.md](SESSION_LOG.md) for gotchas
- [x] `setup_study.py` build / update / read / mesh verified on the sample (power, orientation, air, mesh global and local)
- [x] Page for the mechanical engineer; "In AEDT now" card with differences and Take AEDT values
- [ ] Real Altium export (only the GitHub sample board was used)
- [ ] Real part heights and materials (sample EDB heights are 0.04 mm placeholders)
- [ ] Radiation model name for S2S (the page uses discrete ordinates); Rjc/Rjb, TIM, per-surface h; solving
- [ ] Mesh cell count in the page (Generate mesh reports time only)
