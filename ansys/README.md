# Ansys Icepak route (branch `ansys-aedt`)

Same goal as the SolidWorks route: a PCB to a thermal model. Here the board comes from Altium as an Ansys EDB
(`.aedb`), and a script builds the Icepak project through PyAEDT. Scope for now: **everything up to a ready study.
No solving.** Owners of every piece of information are in [SOURCES.md](SOURCES.md).

## Steps (4 for the user)

| # | Who | What |
|---|---|---|
| 1 | You | Altium: File > Export > Ansys EDB. Put `<board>.aedb` in a project folder |
| 2 | You | `.nsysun.ps1 serve.py <folder>`, open http://localhost:8765, press **Read the board** (once) |
| 3 | You | On the page: ambient, air, mesh, power per RefDes (and height when the EDB value is a placeholder). Press **Set up** |
| 4 | You | Read the checks list; fix what it names; press **Set up** again. Open `<board>_icepak.aedt` in AEDT |

No solve, no spreader, no other cooling solution for now.

## Files

| File | Role |
|---|---|
| `page.html` | The whole page: environment, power table, Set up button, checks list. No build step |
| `serve.py` | Local server for the page (stdlib only). Reads/writes files in the folder, starts the scripts |
| `read_board.py` | EDB -> `board.json` (outline, stackup, parts with position, size, height). PyEDB only, about 7 s |
| `setup_study.py` | `study.json` + EDB -> Icepak project. About 35 s to a minute. Writes `checks.json` |
| `run.ps1` | Runs a script with Ansys' Python and `ansys\.venv` packages (the venv launcher itself is broken) |

Folder contents: `<board>.aedb` (you), `board.json`, `study.json` (page), `checks.json`, `<board>_icepak.aedt`.
`study.json`: `edb`, `ambientC`, `regionPaddingMm`, `meshLevel`, `packageK`, `parts[{refDes, powerW, heightMm?}]`.

What `setup_study.py` does: reads the EDB, connects to AEDT (attach, else start), imports the EDB, creates the Icepak
design and the PCB, one box and one power source per part with power, an air region open on all faces, ambient and
gravity (-Z), a setup (laminar, flow and gravity on) and the mesh level, validates, saves. Everything is rebuilt each
run; it is not incremental yet.

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
- [x] Power blocks per RefDes, air region, ambient, gravity, setup, checks: `setup_study.py` verified on the sample
- [x] Simple page (`page.html` + `serve.py`): read board, type power, Set up, checks list
- [ ] Real Altium export (only the GitHub sample board was used)
- [ ] Real part heights and materials (sample EDB heights are 0.04 mm placeholders)
- [ ] Rjc/Rjb, TIM, emissivity, per-surface h, board orientation; incremental updates; solving
