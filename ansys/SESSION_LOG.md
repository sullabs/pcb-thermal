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
- Field calculator fixed: `get_scalar_field_value` needs `solution="Study : SteadyState"` given explicitly. Never call it on a hollow network block: the failed call breaks the AEDT gRPC session (Save and GetObjectsInGroup fail afterwards; kill AEDT and delete the .lock). Board temperature = maximum over the PCB1 layer objects (`get_3d_component_object_list`), hottest anywhere = `AllObjects`. Heat source: field maximum for plain blocks, case-side monitor for network blocks.
- Convergence check: after a solve the final residuals (continuity, three velocities, energy) are compared with the setup limits. results.json has "converged"; the page shows a green or red line on top of the Result card. The velocity residuals sit right at the limit (0.00085 against 0.001), so the sample only just converged.
- Energy balance: HeatFlowRate face monitors Q_<n> on the six faces of Region, summed (AEDT reports milliwatts, the unit is read from the data) and compared with the total power in. Sample: 7.5 W in, 7.4991 W out. Page shows it on the Result card; over 1 percent is flagged red.
- Radiation residual: `Residual.DOIntensity`, limit `Convergence Criteria - Discrete Ordinates` (1e-6). It exists only when the setup radiation is on. Now part of the convergence check.
- Radiation on: the HeatFlowRate face monitors count only convection (0.63 W of 7.5 W on the sample), so the energy balance is shown as "not checked" instead of a false failure. A radiation-inclusive balance is still open. (Done later: see "Energy balance with radiation".)
- With radiation on, `get_scalar_field_value("AllObjects")` fails and breaks the session. Dropped. "Hottest measured" = max of the board layers and the heat source readings; junction temperature is not read.
- Projects built before the hollow fix keep solve_inside True after AEDT is killed: `_hollow` enforces it before each solve.
- After a solve, reading the design (read_state) fails in the same session, so solve mode skips the read back.
- Open: V1 case temperature rose from 809 to 1072 C with radiation on while the board fell from 748 to 354 C. Unexplained, check on a real board.

Fix round: the whole-model maximum now asks every non-hollow object one by one (no AllObjects call), so it works with radiation on and the session survives. Icepak monitors have no radiation heat flow quantity (only HeatFlowRate on faces, convection only), so the energy balance with radiation stays "not checked"; the Energy residual is the check there. (Superseded: the Fields Summary has the radiation, see "Energy balance with radiation".) The read back right after a solve still fails, so solve mode skips it; the page reads AEDT in its next poll.

## More boards (example-data, pyaedt/edb)
Downloaded to ansys/boards_src (git-excluded); working copies in ansys/boards (excluded).
- Read OK: ANSYS-HSD_V1 (509 parts), ANSYS_SVP_V1_1 (544), BGA_Package (2), cpwg (0), siwave_multi_zones (2).
- Powerboard_SiC_MOSFET.tgz is ODB++, not an aedb. Converts with PyEDB (56 parts) but layer thickness and part heights read 1000x too big (35 mm copper). Not usable without a unit fix.
- Fix added: EDB materials with no density/heat capacity (megtron, solder resist) stop the solver. _thermal_defaults gives them FR-4 values and warns.
- siwave_multi_zones: build OK, solve fails: "PCB layer property out of sync with geometry / Unable to obtain PCB component data". Open.
- ANSYS_SVP_V1_1: build OK, solve OK (170 s), but convergence, monitors and temperatures could not be read afterwards. Open.

## SVP board: unreadable results after solve
Not reproducible. The failing solve came right after AEDT had been killed in a broken session; residuals and monitors read "No trace data" even in a fresh session. A later solve (with and without T_ and Q_ monitors, and through the real script, also after a fresh AEDT start) converged and read everything: board 97.6 C, hottest measured 98.0 C at 10 W, Energy residual 2e-9. Likely cause: stale result files from the killed run. Warning text now says to press Solve again.

## siwave_multi_zones: cause of the failed solve
The EDB has a multi-zone stackup. Icepak's PCB import makes 15 layer objects with D4, D5, L5, L6 twice and no L1, D1, L2, D2. The solver then says "layer property out of sync with geometry / unable to obtain PCB component data". Not fixable from the script. Build now stops with an error naming the duplicated layers (checked right after the import).

## Energy balance with radiation
The Fields Summary (`ic.post.create_field_summary`, entity "Boundary", surface "Open_air") has HeatFlowRate, RadiationFlow and ConvectiveHeatFlow in W. HeatFlowRate there is the total, radiation included. The Q_ face monitors are gone (deleted from older projects at solve). Test board 3.3 W: radiation on 1.49 W convection + 1.81 W radiation = 3.3006 W (+0.02 %); radiation off 3.3012 W, all convection (+0.04 %).
- Page checked in a browser: the Result card shows "Energy balance: 3.3 W in, 3.3006 W out (1.4903 W convection, 1.8103 W radiation, +0.02 %)" in green. An unbalanced result shows red with the split, a results.json from before the split shows the line without it, and an old radiation result still shows the "not checked" warning.
- Fixed: changing the radiation model and pressing Solve restarted from the previous case (`rc previous.nc_cas`), which kept the old radiation model: the residuals still showed do-intensity and the result did not change. Only after deleting the solution did radiation off take effect (board 210 to 325 C). Now `_setup` calls `cleanup_solution` when the radiation model changes; tested off to on (fresh solve, 210 C again) and an unchanged Apply (solution kept).

## Real Altium board: Bluetooth Sentinel (AD22 example)
Altium AD22 File > Export > Ansys EDB works on this PC (the exporter is there after all). It writes `<board>.edb`, not `.aedb`: copied to `example/bluetooth_sentinel/Bluetooth_Sentinel.aedb` (example/ is git-excluded) and it reads and builds.
- Read: 72 parts, none skipped (LEDs, J1 and 6 resistors on the bottom), 80 x 55 mm, 4 copper layers, 0.82 mm. "Dielectric 3" is only 0.0125 mm (from the Altium stackup).
- No part heights in the EDB: every part is 0.036 mm (the copper), `package_def` is null. The 0.2 mm placeholder and its warning cover it; type heights for powered parts.
- No thermal data for any material, copper included (k, density, cp all 0). `_thermal_defaults` fixed only the Icepak copies, so at the end of every solve the PCB component copied the layout's copper back ("Material 'copper': Thermal properties have been updated"), the solution became invalid (`is_solved` False) and nothing could be read ("No trace data"). Fixed: `_thermal_defaults` now also runs on the layout design before the PCB import, and gives copper/conductors copper values, solder solder values (it got FR-4 before) and dielectrics FR-4.
- Test powers (not the board's real ones): U4 0.3 W (Rjc 10, Rjb 15), U1 0.2, U2 0.05, LED1/LED2 0.06 on the bottom. Result: board 60.4 C, hottest 81.6 C (LED1); U4 case 39.8, U1 47.7, U2 44.9, LED2 79.6. Balance 0.67 W in, 0.6701 out (0.19 convection, 0.48 radiation, +0.01 %). Z-velocity residual 1.6e-3 against 1e-3 after 200 iterations: not converged.
- AEDT hung and later crashed (access violation) while builds ran from the command line and a page server polled `read` every 15 s on another folder. With the page server stopped, the same build worked. Likely cause: the poll connecting and releasing in the middle of a build. Not proven. (Now prevented: a run writes a lock file per AEDT port in the temp folder and `read` skips while it exists, see below.)

Follow-up fixes: a bare `<board>.edb` in the folder is copied to `<board>.aedb` on read (again when re-exported); a new build no longer says "Radiation changed" (only an existing setup can change); build/update/mesh/solve write `%TEMP%\pcb_thermal_aedt_50051.busy` and `read` leaves AEDT alone while that process lives. The lock is per AEDT port, not per folder, because the page polls one folder while a command-line run can use another.

## Second Altium board: SL1 Xilinx Spartan-IIE PQ208 (AD example)
`example/spartan_iie`, read straight from the `.edb`: 114 parts, none skipped, 180 x 50.5 mm, 2 copper layers, 1.67 mm.
- Same as Sentinel: every height 0.036 mm, and every material has electrical values (copper 5.8e7 S/m, FR-4 Dk 4.8) but k, density and cp 0. Two of two exports, so this is the Altium exporter, not the board.
- First solve failed: "Parts U5 and U1 intersect", "Parts Y1 and U1 intersect". The EDB box is the pad extent (U1 PQ208 35.7 x 34.5 mm, body 28 x 28), and close neighbours overlap it by 0.37 and 1.73 mm. Fixed: `_footprints` cuts two overlapping same-side heat-source blocks back to meet halfway across the smaller overlap, with a warning.
- Test powers (not real): U1 0.6 W (Rjc 8, Rjb 15), U3 0.5, U4 0.4 (TO-263), U5 0.05, U2 0.05, Y1 0.03; 500 iterations. Converged in 132 s (Z-velocity 9.99e-4, just under 1e-3). Board 45.6 C, hottest U3 53.8 C, U4 52.2, U5 42.6, U1 case 41.6, Y1 36.0, U2 34.2. Balance 1.63 W in, 1.6301 out (0.49 convection, 1.14 radiation, +0.01 %).
- Open: radiation takes 70 % here and on Sentinel. Convection 0.49 W over about 0.018 m2 of board at about 20 K rise is about 1.4 W/m2K, where natural convection is usually 5 to 10. Check with a finer mesh (level 3 to 4) whether convection is under-resolved at mesh level 2.

## Filling what the export lacks from the page, start.bat
Altium exports no heights, only pad outlines and no thermal material values (two of two boards). The page now fills them:
- "What the export is missing" card: how many parts have placeholder heights, that sizes are pad outlines, which board materials have no k/density/cp. A materials table (from `board.json`, which now lists the stackup materials with their layers) shows the stand-in in grey; typed values go to `study.json` `materials` and are set on Build and on Apply. Apply works because materials belong to the AEDT project, which holds the layout design the PCB copies from. Checked from the page: FR-4 k 0.4 typed, Apply, read back 0.4.
- Parts: Body X / Body Y (board axes, centred on the footprint) replace the pad outline. A part with a height and no power becomes a block without heat (object named after the RefDes, no `P_` boundary); `read` reports it as kind `Passive`. Checked on Spartan: HDR2, HDR3 8.5 mm and J1 11 mm, U1 body 28 x 28: 9 of 9 parts, 3 without power, no overlaps left.
- CSV columns: RefDes, W, height, body X, body Y (the last three optional). "Read the export again" after a new export.
- `start.bat [folder]`: starts AEDT with `-grpcsrv 50051` if it is not running, then `serve.py`; if the page already runs it only opens it. The page has a board folder picker (the start folder and its sub-folders with an export).
