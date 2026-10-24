# Session log: Ansys AEDT / EDB tests

Branch `ansys-aedt` (from `webpage`), folder `ansys/`. Nothing is committed yet.

## Where it stands (2026-10-24, second session)

The whole chain **EDB -> AEDT -> Icepak PCB** works from Python on a fresh AEDT, about 45 s in total:

| Step | Result | Time |
|---|---|---|
| Start AEDT (`Desktop(new_desktop=True, port=50051)`) or attach (`new_desktop=False`) | works | 14.5 s / 4 s |
| PyEDB opens `power_board.aedb` | works, no licence error | 6.7 s |
| `Hfss3dLayout(...)` then `import_edb(<board>.aedb)` | works | 7 s + 7.5 s |
| `Icepak(project=, design=)` | works | 4 s |
| `Icepak.create_pcb_from_3dlayout(name, project, design)` | works, PCB1 bbox equals the EDB outline | 10.6 s |
| Icepak basics: box, `assign_source`, `create_setup`, mesh region, save | work | under 1 s each |

What PyEDB reads fine: outline bbox, stackup (layers, thickness, total 1.232 mm), materials, components (RefDes,
part name, type, placement layer, centre), nets, power nets, via count, polygons/paths, copper area per layer.

### Gotchas that cost time

- **`import_edb` puts the board in a new project named after the file** (`power_board`, design `siw_dc`), not in
  the design you created. Read `h.project_name` / `h.design_name` after the import and pass those to
  `create_pcb_from_3dlayout`. Passing the empty `lay1` design froze AEDT's gRPC for good (the first session's
  "hang" was most likely the same kind of wrong-name call, not a licence problem).
- **`save_project(path)` is a save-as**: the project is renamed to the file name. Look the project up by its new
  name afterwards.
- **A frozen AEDT does not recover** (window "responding", CPU flat, every gRPC call blocks). Only killing
  `ansysedt.exe` and starting a new one helped. Always run test scripts with a hard timeout.
- **Opening AEDT through the desktop-control tool starts a second, empty instance.** Do not.
- PyAEDT calls that look like API errors were wrong arguments on my side: `Desktop.project_list` is a property;
  `assign_source` takes `thermal_condition=`, not `assignment_type=`; Icepak already creates a `Region`, so
  `create_region` returns `False`. After one failed call, later calls in the same run fail too (cascade), so
  isolate tests.
- `Edb(edbversion=...)` is deprecated; use `version="2026.1"`. The server warns that shared-memory IPC and fast
  mode need AEDT 2026.1.2 or newer; it falls back to normal gRPC and works.

### Still open

- **Components in Icepak.** The PCB arrived as 15 stackup-layer bodies (`PCB_001_TOP_BBox` ...) plus `Region`.
  No per-component bodies (C1 ...) were seen. Not tested: `create_ipk_3dcomponent_pcb` options for components,
  and how per-RefDes power is applied.
- **`validate_simple()` returned `False`** on the bare Icepak design. Reason not checked (probably no setup or
  boundaries yet).
- Component height is not at `component_property.height` (PyEDB); find the right property. `components.others`
  does not exist (use `instances` and filter by `type`).
- `FR4` is not in the Icepak material list; materials must come from the library by exact name or be created.
- Altium export of a real board (File > Export > Ansys EDB) is untested; only the GitHub sample was used.

## Machine facts (checked, corrected)

- AEDT 2026 R1 (v261), `ansysedt.exe`, gRPC on port 50051.
- Install: `E:\Program Files\ANSYS Inc\v261\AnsysEM` (**no** `ansys\` folder; the first log had this wrong).
- Env vars: `ANSYSEM_ROOT261`, `AWP_ROOT261`, `ANSYS261_DIR`, `ANSYSLMD_LICENSE_FILE=1055@localhost`.
  Licence server `ansyslmd` runs on this PC; `lmutil lmstat -a -c 1055@localhost` showed it UP with 0 licences
  in use. `lmutil` is at `E:\Program Files\ANSYS Inc\Shared Files\fnp\winx64\lmutil.exe`.
- Ansys Python 3.10.19:
  `E:\Program Files\ANSYS Inc\v261\AnsysEM\commonfiles\CPython\3_10\winx64\Release\python\python.exe`.
- **`ansys/.venv` launcher is broken** (`python.exe` fails with `No Python at '"E:\...`), because it was created
  with the wrong install path. Workaround used: run the Ansys Python directly with
  `PYTHONPATH=ansys\.venv\Lib\site-packages`. Proper fix: recreate the venv from the correct path, then
  `pip install -r ansys/requirements.txt`.

## First session (2026-10-24, earlier)

1. PyAEDT 1.9.0 attached to the running AEDT over gRPC (read-only).
2. `pip install pyaedt` needs `pip install -U pip` first (about a minute after that). Send pip output to a file,
   not into `head`/`Select -First`.
3. `ansys/.venv` built from the Ansys Python with `pyaedt==1.9.0` (PyEDB 0.85.0 included); `requirements.txt` pins
   it; `.gitignore` has `ansys/.venv/` and `ansys/**/__pycache__/`.
4. Format decision: Ansys EDB (`.aedb`) first, then ODB++, then IPC-2581.
5. Sample board `power_board.aedb/edb.def` (136,117 bytes, MIT) downloaded from github.com/ansys/example-data
   (`pyaedt/edb/power_board.aedb`). Other candidates there: `icepak_pcb_building/EDB.aedb` (made for Icepak),
   `ANSYS-HSD_V1.aedb`, `BGA_Package.aedb`.
6. `SOURCES.md` and `README.md` written (single owner per piece of information; 4 user steps).
7. Failures then: PyEDB reported "No valid license detected" and looped 13 minutes; Hfss3dLayout/Icepak design
   insertion hung. Neither reproduced in the second session (see Gotchas for the likely cause of the hang).

## Third session: automation built

`read_board.py`, `setup_study.py`, `serve.py`, `page.html`, `run.ps1` written and run on the sample board.
Page flow tested in the built-in browser: Read the board -> power on V1, L1, C1 -> Set up -> "Ready". Read back in
AEDT: blocks V1, L1, C1 with 1, 0.5, 0.1 W, opening `Open_air`, setup `Study` (laminar, flow, gravity, 300 iters),
`validate_simple()` True, one PCB component. Times: connect 1 s, import EDB 11 s, Icepak design 4 s, PCB 10.5 s.
One run took longer than the others (waited on AEDT for about a minute) but finished.
Not read back: gravity direction and ambient value (set through `edit_design_settings`, no error).
PyEDB: component `bounding_box` (m), `lower_elevation`/`upper_elevation` (m), `placement_layer`.
Icepak: `materials.add_material(name)` returns the Material; `modeler.create_region(pad_value=[mm]*6,
pad_type="Absolute Offset")` after deleting `Region`; `assign_free_opening(face ids)`.

## Third session, part 2: engineer page, edit mode, mesh

- Gravity code of `edit_design_settings` is 0 -X, 1 -Y, 2 -Z, 3 +X, 4 +Y, 5 +Z (read back from `GravityVec` and
  `GravityDir`). My first guess was wrong; fixed.
- **Slow call**: `ic.boundaries` (5 to 10 s) and `boundary.props` (about 10 s each). Use the model tree instead:
  `ic.get_oo_name(ic.odesign, "Thermal")` and `ic.get_oo_property_value(node, name, "Total Power")` take 0.1 s.
  Delete with `ic.odesign.GetModule("BoundarySetup").DeleteBoundaries([name])`, set with `SetPropValue`.
  Local mesh values: `get_oo_object(ic.odesign, "Mesh")`, property `Maximum Element Size/X`.
- Local mesh regions create a `Mesh_<Ref>_SubRegion` object; delete stale ones and `modeler.refresh_all_ids()` before
  recreating, else `KeyError: 'Mesh_V1_SubRegion'`.
- Closing a project by name fails when two have the same name (a saved-as project and a re-import). Build closes
  same-name projects first. Failed attaches create `ProjectN` junk projects.
- Setup property `Radiation Model` accepts any string without error; only "Discrete Ordinates Model" is a known
  valid value (library template). Surface to surface name unverified.
- Generate mesh on the sample: 13 s. Update after changes: 40 to 80 s.

## Third session, part 3: junction model, truth from AEDT

- `create_two_resistor_network_block(object, pcb, power, rjb, rjc)`: `power` must be a **string** ("3W"), a float fails
  with a generic "Failed to create boundary". The block must touch the PCB, so blocks now sit on the board surface
  (z of the PCB1 top). Rename the result to `P_<RefDes>` with `.name =`. Read back from the tree:
  `Internal: Properties/Power`, `Rjc: Properties/ThermalResistance`, `Rjb: Properties/ThermalResistance`.
  Power change only: `SetPropValue` on the same property works.
- Two PyAEDT clients attaching at the same moment hung one of them (the page's background read vs my test script).
  Server jobs are serialised and killed after a timeout.
- `read` mode: `tasklist` check first, so it never starts AEDT; does not open a project that is not open.
- Page: form values follow AEDT unless edited (yellow); `derive()` maps state to form; air margins come from the
  region against the union of PCB and blocks.

## Not started

Rjc/Rjb, TIM, emissivity, per-surface h, orientation, incremental updates, real Altium board, real heights.
The old SolidWorks page lives on branch `addin-mvp` (worktree `thermal_wt_addin`); this page is separate and much smaller.

## Repo state

Untracked, uncommitted: `ansys/`, the `.gitignore` edit, `FLOW_PLAN.md`, `.claude/`. Test scripts are in the
session scratchpad only.

## Fourth session: Solve and results

- New mode `solve` (page button Solve): puts a point temperature monitor `T_<ref>` on every heat source, runs `analyze_setup("Study")`, reads the monitors, writes `results.json` (temperature per heat source, hottest heat source, Tj max margin) and the page shows a Result card. The kill timer for a solve job is 1 hour.
- Network (Rjc/Rjb) blocks need `solve_inside = False`. Without it the solver input fails with "Face 'Case_side' of Network has mesh on both sides". Fixed in `_blocks`.
- Network blocks are hollow, so their monitor sits on the case-side face (top centre). Plain blocks get it in the body.
- The field calculator (`get_scalar_field_value`) fails in this setup. Monitors plus `get_solution_data(report_category="Monitor")` work.
- Not available yet: board hot spot, true junction temperature of a network. The sample board gives huge temperatures (7 W on a tiny part): the numbers prove the pipeline, not the physics.
- A killed AEDT leaves a `.lock` file next to the .aedt. Delete it before reopening.
- `design_settings` can be None right after a solve. Do not read it there.
