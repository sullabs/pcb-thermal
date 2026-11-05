"""Build or edit the Icepak study from <folder>/study.json. No solve.

    python ansys/setup_study.py <folder> [build|update|read|open|mesh|solve|meshcheck]

build  : new AEDT project from the EDB: PCB, power blocks, air, environment, setup, mesh (about 40 s)
update : the same settings applied to the project that is already built (fast); AEDT edits to the things the
         page owns (power, ambient, air, mesh) are overwritten, everything else is left alone
read   : only read the project into state.json, as it is in AEDT now (never starts AEDT or opens a project)
open   : start AEDT if needed, open the built project and read it into state.json. Changes nothing
mesh   : generate the mesh and report how long it took
solve  : solve the setup (opens the project, never changes it) and write results.json: temperature per heat source
meshcheck: solve with a finer mesh, then with the page's mesh again, and write meshcheck.json with both results.
         The project ends as it was, with the page's mesh and its solution in results.json

Input  : <folder>/<board>.aedb (or Altium's <board>.edb), <folder>/study.json (written by the page)
Output : <folder>/board.json, state.json, checks.json, <board>_icepak.aedt; built.json (what Build used that only Build
         can change), meshed.json (the meshSig of state.json when the mesh was last generated, by mesh or a solve),
         applied.json (the body sizes build or update made the blocks from), solution_applied.json (which SolidWorks
         export the SOL_ objects of the thermal solution were made from)
"""
import hashlib
import json
import math
import os
import re
import subprocess
import sys
import tempfile
import time
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).parent))
import read_board  # noqa: E402

VERSION = read_board.VERSION
PORT = 50051
MIN_BLOCK_H_MM = 0.2   # EDB heights below this are placeholders; the page value or this minimum is used
DESIGN = "Icepak"
SETUP = "Study"
PKG = "pkg_material"
RJC_MAT = "rjc_"         # rjc_<refdes>: the material of a block with Rjc and no Rjb, its k gives Rjc through the height
SOL = "SOL_"             # objects of the thermal solution: SOL_Spreader_Top, SOL_Spreader_Bottom, SOL_TIM_<RefDes>
SOL_MAT = "sol_"         # their materials: sol_plate_top, sol_plate_bottom, sol_tim_<refdes>
PLATE_MAT = {"Aluminium": (167, 2700, 896), "Copper": (390, 8933, 385)}    # k W/m.K, density kg/m3, heat capacity J/kg.K
TIM_DENSITY_CP = (2500, 1000)
# exists while a build/update/mesh/solve talks to the AEDT on PORT, for any board folder; mode read leaves AEDT alone then
BUSY = Path(tempfile.gettempdir()) / f"pcb_thermal_aedt_{PORT}.busy"

# study["orientation"] is the direction gravity points in the board's axes: the page turns the board from flat, parts up,
# about its X or Y axis in 90 degree steps, which always puts gravity on one of the six axis directions.
# edit_design_settings code: 0 -X, 1 -Y, 2 -Z, 3 +X, 4 +Y, 5 +Z (read back from AEDT)
# fields: gravity code, label, text, index of the air face opposite to gravity in [+X, -X, +Y, -Y, +Z, -Z]
ORIENTATION = {
    "-Z": (2, "-Z", "flat, parts face up", 4),
    "+Z": (5, "+Z", "flat, parts face down", 5),
    "-Y": (1, "-Y", "standing on its -Y edge", 2),
    "+Y": (4, "+Y", "standing on its +Y edge", 3),
    "-X": (0, "-X", "standing on its -X edge", 0),
    "+X": (3, "+X", "standing on its +X edge", 1),
}
OLD_ORIENTATION = {"flat-up": "-Z", "flat-down": "+Z", "vertical-long": "-Y", "vertical-short": "-X"}   # study.json before the turn control

checks = []


def check(level, text):
    """level: ok | warn | error"""
    checks.append({"level": level, "text": text})
    print(f"[{level}] {text}", flush=True)


def step(label, fn, fatal=False):
    t = time.time()
    try:
        r = fn()
        print(f"  {label}: {time.time() - t:.1f}s", flush=True)
        return r
    except BaseException as ex:  # noqa: BLE001  keep going, report
        check("error", f"{label} failed: {type(ex).__name__}: {str(ex).splitlines()[0][:160] if str(ex) else ''}")
        if fatal:
            raise
        return None


def load_study(folder: Path) -> dict:
    f = folder / "study.json"
    s = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    if not f.exists():
        check("warn", "No study.json yet: defaults used, no heat sources")
    s.setdefault("ambientC", 25)
    s["orientation"] = OLD_ORIENTATION.get(s.get("orientation"), s.get("orientation"))
    if s["orientation"] not in ORIENTATION:
        s["orientation"] = "-Z"
    s.setdefault("marginMm", 50)
    s.setdefault("marginAboveMm", 50)
    s.setdefault("radiation", True)
    s.setdefault("pcbResolution", 2)
    s.setdefault("iterations", 300)
    s.setdefault("packageK", 1.0)
    s.setdefault("mesh", {"mode": "auto", "level": 3})
    s.setdefault("parts", [])
    return s


_TO_PAGE = {"w": 1.0, "mw": 1e-3, "uw": 1e-6, "kw": 1e3,                                  # power: W
            "mm": 1.0, "um": 1e-3, "nm": 1e-6, "cm": 10.0, "m": 1e3, "meter": 1e3, "mil": 0.0254, "in": 25.4}   # length: mm
_TEMP = {"cel": lambda t: t, "kel": lambda t: t - 273.15, "fah": lambda t: (t - 32) / 1.8}   # temperature: C


def num(text, default=None):
    """a value as AEDT holds it, in the page's units: '1.5W' or '500mW' -> W, '30cel' or '303.15kel' -> C, '2mm' or
    '0.002meter' -> mm. Any other unit ('5cel_per_w') or none: the number as it is. Values typed in AEDT may carry
    any unit, and the page writes back in its own"""
    m = re.match(r"\s*(-?[\d.]+(?:[eE][-+]?\d+)?)\s*([A-Za-z_]*)", str(text))
    if not m:
        return default
    v, unit = float(m.group(1)), m.group(2).lower()
    return _TEMP[unit](v) if unit in _TEMP else v * _TO_PAGE.get(unit, 1.0)


def connect():
    from ansys.aedt.core import Desktop
    try:
        return Desktop(version=VERSION, new_desktop=False, port=PORT, non_graphical=False)
    except BaseException:  # noqa: BLE001
        return Desktop(version=VERSION, new_desktop=True, port=PORT, non_graphical=False)


def _quiet_history():
    """PyAEDT 1.9 reads an object's history (ic.mesh does, for the air region and each mesh region) and copies its
    "Command" onto a child node with SetPropValue. AEDT refuses, Command is read only, and logs "Script macro error:
    Command is read only" each time: every 15 s from the page's read. The node's properties are read fresh from AEDT
    on every access, so the copy never did anything: skip only that write."""
    from ansys.aedt.core.modeler.cad.elements_3d import BinaryTreeNode
    if getattr(BinaryTreeNode.update_property, "_quiet", False):
        return
    orig = BinaryTreeNode.update_property

    def update_property(self, prop_name, prop_value):
        return False if prop_name == "Command" else orig(self, prop_name, prop_value)
    update_property._quiet = True
    BinaryTreeNode.update_property = update_property


def main(folder: Path, mode: str = "build") -> int:
    from ansys.aedt.core import Hfss3dLayout, Icepak
    from ansys.aedt.core.generic.settings import settings
    settings.enable_screen_logs = False
    _quiet_history()
    fresh = mode == "solvefresh"        # a solve that deletes the old solution first, so it starts from nothing
    if fresh:
        mode = "solve"

    study = load_study(folder)
    if mode == "build" or not (folder / "board.json").exists():
        try:
            read_board.main(folder)
        except BaseException as ex:  # noqa: BLE001  no .aedb (SystemExit) or PyEDB failed: say so in checks.json
            check("error", f"Reading the EDB failed: {str(ex).splitlines()[0][:160] if str(ex) else type(ex).__name__}")
            return _finish(folder, 0)
    board = json.loads((folder / "board.json").read_text(encoding="utf-8"))
    by_ref = {p["refDes"]: p for p in board["parts"]}
    project_file = folder / f"{board['name']}_icepak.aedt"

    if mode == "read":
        return read_only(folder, board, project_file)
    for ref in board.get("skipped", []):
        check("warn", f"{ref}: could not be read from the EDB (no position or size), it cannot get a heat source")
    # a block per part with power (heat source) or with only a height, typed or from the STEP (no heat, it shapes the air)
    parts = [p for p in study["parts"] if (p.get("powerW") or 0) > 0 or (p.get("heightMm") or 0) > 0
             or by_ref.get(p["refDes"], {}).get("stepHMm")]
    for p in parts:
        if p["refDes"] not in by_ref:
            check("error", f"{p['refDes']} has a power or height but is not in the EDB")
    parts = [p for p in parts if p["refDes"] in by_ref]
    powered = [p for p in parts if (p.get("powerW") or 0) > 0]
    total = sum(p["powerW"] for p in powered)
    if mode in ("build", "update"):
        check("ok" if powered else "warn",
              f"{len(powered)} heat sources, total {total:.3f} W" if powered else "No part has power: no heat source")

    try:
        desktop = step("connect to AEDT", connect, fatal=True)
    except BaseException:  # noqa: BLE001
        return _finish(folder, total)
    try:
        if mode == "build":
            ic = step("build project", lambda: _build(Hfss3dLayout, Icepak, desktop, folder, board, study), fatal=True)
        else:
            if not project_file.exists():
                raise RuntimeError("No built project yet. Press Build first.")
            ic = step("open project", lambda: _open(Icepak, project_file), fatal=True)
        if mode in ("update", "mesh", "solve", "meshcheck"):
            # PyAEDT reads the setup and mesh settings from the saved .aedt, not from the AEDT window: save first, so an
            # edit made in AEDT and not saved yet is seen (and kept where the page does not own it), not undone
            step("save the edits made in AEDT", lambda: ic.save_project())
        if mode == "open":
            check("ok", f"Opened {project_file.name} in AEDT")
        if mode == "update":
            step("materials", lambda: _thermal_defaults(ic, study, board))
        if mode in ("build", "update"):
            if _apply(ic, folder, board, by_ref, parts, study):
                _applied(folder, parts)
            step("save project", lambda: ic.save_project())
            check("ok", f"Saved {project_file.name}")
        if mode == "mesh":
            meshed = _mesh_now(ic)
        if mode == "solve":
            if fresh and step("clear old solution", lambda: ic.cleanup_solution()):
                check("ok", "Old solution deleted: this solve starts fresh")
            step("solve", lambda: _solve(ic, folder, board, powered, study, parts), fatal=True)
        if mode == "meshcheck":
            step("mesh check", lambda: _mesh_check(ic, folder, board, parts, powered, study), fatal=True)
        state = None if mode in ("solve", "meshcheck") else step("read back", lambda: read_state(ic, board))   # right after a solve AEDT cannot be read
        if state:
            (folder / "state.json").write_text(json.dumps(state, indent=1), encoding="utf-8")
        if mode == "mesh" and meshed and state:
            _meshed(folder, state["meshSig"], "generated")
        elif mode in ("solve", "meshcheck") and not any(c["level"] == "error" for c in checks):
            old = json.loads((folder / "state.json").read_text(encoding="utf-8")) if (folder / "state.json").exists() else {}
            if old.get("meshSig"):      # AEDT cannot be read right after a solve: the last read, as Solve needs Apply first
                _meshed(folder, old["meshSig"], "solved")
    except BaseException as ex:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        check("error", f"Stopped: {type(ex).__name__}: {str(ex).splitlines()[0][:160] if str(ex) else ''}")
    finally:
        desktop.release_desktop(close_projects=False, close_on_exit=False)
    return _finish(folder, total)


def _finish(folder, total):
    n_err = sum(c["level"] == "error" for c in checks)
    (folder / "checks.json").write_text(json.dumps({
        "schema": "pcb-thermal-ansys/checks",
        "ready": n_err == 0,
        "totalPowerW": round(total, 4),
        "checks": checks,
    }, indent=1), encoding="utf-8")
    print("READY" if n_err == 0 else f"NOT READY: {n_err} error(s)")
    return 0 if n_err == 0 else 1


# ---------------------------------------------------------------- build

def _close_same_name(desktop, names):
    for n in names:
        for _ in range(3):   # earlier runs can leave duplicates
            if n in list(desktop.odesktop.GetProjectList()):
                try:
                    desktop.odesktop.CloseProject(n)
                except BaseException:  # noqa: BLE001
                    break


THERMAL = read_board.THERMAL
MAT_F = (("kWmK", "thermal_conductivity", 0.01), ("rhoKgM3", "mass_density", 0), ("cpJkgK", "specific_heat", 0))


def _thermal_defaults(app, study, board):
    """EDB materials from an electrical export (Altium, Megtron, Rogers...) can lack density, heat capacity and
    conductivity: the solver then refuses to start or, for copper, solves with no conduction. Values typed on the page
    (study "materials") are used first, a missing one gets copper, solder or FR-4 values, and the check says which.
    A board material the page lists (board.json) and nobody typed gets the export's value, else the stand-in, as the
    page shows: a value typed before and cleared goes back too.
    Run on the layout design before the PCB import too: the PCB component copies the layout's materials again at every
    solve, and a changed material marks the solution invalid. On update it runs on the Icepak design: materials belong
    to the project, which holds both designs."""
    typed = {k.lower(): v for k, v in study.get("materials", {}).items()}
    export = {b["name"].lower(): b for b in board.get("materials", [])}
    fixed, own = {}, set()
    for name, m in app.materials.material_keys.items():
        if name == PKG or name.startswith(RJC_MAT) or name.startswith(SOL_MAT):     # the blocks' and plates' own materials
            continue
        try:
            t, e = typed.get(name.lower(), {}), export.get(name.lower())
            kind = e["kind"] if e else read_board.thermal_kind(name, not m.is_dielectric())
            for i, (f, attr, low) in enumerate(MAT_F):
                have = float(getattr(m, attr).value)
                if t.get(f):
                    if abs(have - t[f]) > 1e-9:
                        setattr(m, attr, t[f])
                        own.add(getattr(m, "name", name))
                elif e:
                    want = e[f] if (e.get(f) or 0) > low else THERMAL[kind][i]
                    if abs(have - want) > 1e-9 * max(1, want):
                        setattr(m, attr, want)
                        if want == THERMAL[kind][i]:
                            fixed.setdefault(kind, set()).add(name)
                elif have <= low:
                    setattr(m, attr, THERMAL[kind][i])
                    fixed.setdefault(kind, set()).add(name)
        except BaseException:  # noqa: BLE001
            continue
    if own:
        check("ok", f"Material values from the page: {', '.join(sorted(own))}")
    if fixed:
        used = "; ".join(f"{kind} values ({THERMAL[kind][0]} W/mK) for {', '.join(sorted(names))}" for kind, names in fixed.items())
        check("warn", f"No thermal data in the EDB: {used}. Type the real values on the page (Board materials) if they matter.")


def _build(Hfss3dLayout, Icepak, desktop, folder, board, study):
    _close_same_name(desktop, [board["name"], f"{board['name']}_icepak"])
    # the EDB becomes a new project of its own. Not Hfss3dLayout() then import_edb: without a project, Hfss3dLayout()
    # puts an empty layout design into whatever project is active in AEDT, or leaves a new empty project open
    desktop.odesktop.GetTool("ImportExport").ImportEDB(str(folder / board["edb"] / "edb.def"))
    proj = desktop.odesktop.GetActiveProject()
    pn, dn = proj.GetName(), proj.GetActiveDesign().GetName().split(";")[-1]
    lay = Hfss3dLayout(project=pn, design=dn, version=VERSION, new_desktop=False, port=PORT)
    _thermal_defaults(lay, study, board)
    ic = Icepak(project=pn, design=DESIGN, version=VERSION, new_desktop=False, port=PORT)
    ok = ic.create_pcb_from_3dlayout("PCB", pn, dn, resolution=int(study["pcbResolution"]), extent_type="Bounding Box",
                                     close_linked_project_after_import=False,
                                     rad="Both" if study["radiation"] else "Nothing")
    check("ok" if ok else "error", f"PCB from the EDB, mapping resolution {study['pcbResolution']}"
          if ok else "PCB import returned False")
    _thermal_defaults(ic, study, board)
    names = [re.sub(r"^PCB_\d+_|_BBox$", "", o) for o in ic.modeler.get_3d_component_object_list(name=PCB)]
    dup = sorted({n for n in names if names.count(n) > 1})
    if dup:
        check("error", f"The stackup has more than one zone (rigid-flex or stepped board): layers {', '.join(dup)} are "
                       f"imported twice, so the solver will refuse to start. Icepak's PCB import cannot take this board. "
                       f"Export one zone only, or model it in AEDT.")
    ic.save_project(str(folder / f"{board['name']}_icepak.aedt"))
    (folder / "meshed.json").unlink(missing_ok=True)
    (folder / "built.json").write_text(json.dumps({"schema": "pcb-thermal-ansys/built", "at": time.strftime("%Y-%m-%d %H:%M"),
                                                   "edb": board["edb"], "pcbResolution": int(study["pcbResolution"])}, indent=1),
                                       encoding="utf-8")
    return ic


# ---------------------------------------------------------------- apply (same for build and update)

def _apply(ic, folder, board, by_ref, parts, study):
    """True when the part blocks were made or fixed"""
    rebuilt = step("parts", lambda: _blocks(ic, board, by_ref, parts, study))
    step("thermal solution", lambda: _solution(ic, folder, board, study))
    step("air", lambda: _air(ic, study, board))
    step("environment", lambda: _environment(ic, study))
    step("setup", lambda: _setup(ic, study))
    step("mesh", lambda: _mesh(ic, parts, study, rebuilt or set()))
    return rebuilt is not None


def _applied(folder, parts):
    """applied.json: the body sizes the blocks were made from. AEDT cannot give them back (a block cut back for an
    overlap is smaller), and study.json holds the page's values whether applied or not, so the page compares with this"""
    (folder / "applied.json").write_text(json.dumps({
        "schema": "pcb-thermal-ansys/applied", "at": time.strftime("%Y-%m-%d %H:%M"),
        "bodies": {p["refDes"]: {k: p[k] for k in ("bodyXMm", "bodyYMm") if p.get(k)} for p in parts}}, indent=1),
        encoding="utf-8")


# ---------------------------------------------------------------- thermal solution (SolidWorks plates)

def _sol_material(ic, name, k, rho, cp):
    """an own material of the solution; only what differs is written (a changed material voids the solution)"""
    m = ic.materials.material_keys.get(name) or ic.materials.add_material(name)
    if not m:
        return None
    for attr, v in (("thermal_conductivity", k), ("mass_density", rho), ("specific_heat", cp)):
        try:
            same = abs(float(getattr(m, attr).value) - v) <= 1e-9 * max(1, v)
        except (TypeError, ValueError):
            same = False
        if not same:
            setattr(m, attr, v)
    return name


def _solution(ic, folder, board, study):
    """The plates SolidWorks sent (Send Solution: solution_cad.json, solution_<side>.step) and their TIMs as objects of
    this design. A plate is imported again only when its STEP or its place changed; a TIM is a box of SolidWorks' size.
    Turned off on the page (the whole solution, or one plate): its objects stay but are not in the model, so they are
    out of the mesh and the solve. Coordinates: STEP point + (offXMm, offYMm), and the STEP's board top face on the PCB's"""
    sol = study.get("solution") or {}
    names = set(ic.modeler.object_names)
    have = sorted(n for n in names if n.startswith(SOL))
    cad_f = folder / "solution_cad.json"
    cad = json.loads(cad_f.read_text(encoding="utf-8")) if cad_f.exists() else {}
    if not cad.get("plates"):
        for n in have:
            if ic.modeler[n].is_model:
                ic.modeler[n].is_model = False
        if sol.get("on"):
            check("warn", "Thermal solution: SolidWorks has not sent the plates yet (Build Solution, finish them, Send "
                          "Solution). Nothing of it is in AEDT" + (", the old objects are out of the model." if have else "."))
        return
    fr = board.get("stepFrame") or cad.get("frame")
    if not fr:
        check("error", "Thermal solution: board.json has no STEP frame, so the plates cannot be placed. Press Read heights from the STEP.")
        return
    zlo, zhi = board_z(ic, board)
    dx, dy, dz = fr["offXMm"], fr["offYMm"], zhi - fr["topZMm"]
    plates_in = sol.get("plates") or {}
    plan = {x["side"]: x for x in (sol.get("forSolidWorks") or {}).get("plates", [])}
    on = {side: bool(sol.get("on")) and (plates_in.get(side) or {}).get("on", True) for side in ("top", "bottom")}
    rec_f = folder / "solution_applied.json"
    rec = json.loads(rec_f.read_text(encoding="utf-8")) if rec_f.exists() else {}
    done = {}
    for pl in cad["plates"]:
        side = pl["side"]
        stem = f"{SOL}Spreader_{side.title()}"
        f = folder / pl.get("file", "")
        if not pl.get("file") or not f.exists():
            check("error", f"Thermal solution: {pl.get('file') or 'the plate STEP'} is not in the board folder. Press Send Solution again.")
            continue
        sig = hashlib.sha1(f.read_bytes() + json.dumps([pl.get("offsetMm"), dx, dy, dz]).encode()).hexdigest()[:16]
        mine = [n for n in names if n == stem or n.startswith(stem + "_")]
        if not mine or (rec.get("plates") or {}).get(side) != sig:
            for n in mine:
                ic.modeler.delete(n)
            before = set(ic.modeler.object_names)
            ic.modeler.import_3d_cad(str(f), create_group=False)
            new = sorted(set(ic.modeler.object_names) - before)
            if not new:
                check("error", f"Thermal solution: {f.name} imported no solid. Open it in AEDT to see why.")
                continue
            mine = []
            for i, n in enumerate(new):
                o = ic.modeler[n]
                o.name = stem if len(new) == 1 else f"{stem}_{i + 1}"
                ox, oy, oz = pl.get("offsetMm") or (0, 0, 0)
                o.move([ox + dx, oy + dy, oz + dz])
                mine.append(o.name)
            check("ok", f"Thermal solution: {side} plate imported from {f.name}" + (f" ({len(new)} bodies)" if len(new) > 1 else ""))
        p_in = plates_in.get(side) or {}
        mat = p_in.get("material") or "Aluminium"
        k = (plan.get(side) or {}).get("kWmK") or p_in.get("kWmK") or PLATE_MAT.get(mat, PLATE_MAT["Aluminium"])[0]
        _, rho, cp = PLATE_MAT.get(mat, PLATE_MAT["Aluminium"])
        mname = _sol_material(ic, f"{SOL_MAT}plate_{side}", k, rho, cp)
        for n in mine:
            o = ic.modeler[n]
            if mname and o.material_name.lower() != mname:
                o.material_name = mname
            if o.is_model != on[side]:
                o.is_model = on[side]
        done[side] = sig
    # TIMs: a box each, SolidWorks' size, the page's k
    parts_in = sol.get("parts") or {}
    want = {}
    for t in cad.get("tims", []):
        ref, b = t.get("refDes"), t.get("boxMm")
        if not ref or not b:
            continue
        box = [b[0] + dx, b[1] + dy, b[2] + dz, b[3] + dx, b[4] + dy, b[5] + dz]
        want[f"{SOL}TIM_{ref}"] = (ref, t.get("side") or ("top" if box[2] >= zhi - 1e-3 else "bottom"), box)
    for n in [n for n in names if n.startswith(f"{SOL}TIM_") and n not in want]:
        ic.modeler.delete(n)
    gaps = []
    for n, (ref, side, box) in want.items():
        if n in names and max(abs(a - c) for a, c in zip(ic.modeler[n].bounding_box, box)) > 1e-3:
            ic.modeler.delete(n)
        if n not in set(ic.modeler.object_names):
            ic.modeler.create_box(box[:3], [box[3] - box[0], box[4] - box[1], box[5] - box[2]], name=n)
        o = ic.modeler[n]
        mname = _sol_material(ic, f"{SOL_MAT}tim_{ref}".lower(), (parts_in.get(ref) or {}).get("timK") or 3, *TIM_DENSITY_CP)
        if mname and o.material_name.lower() != mname:
            o.material_name = mname
        if o.is_model != on[side]:
            o.is_model = on[side]
        if ref in names and on[side]:       # the TIM must sit on the part's block: a typed height moves the block, not the STEP
            bb = ic.modeler[ref].bounding_box
            g = box[2] - bb[5] if side == "top" else bb[2] - box[5]
            if abs(g) > 0.02:
                gaps.append(f"{ref} {g:+.2f} mm")
    if gaps:
        check("warn", f"Thermal solution: TIMs not on their part's block ({', '.join(gaps)}; + is a gap, - an overlap). "
                      f"The block height on the page differs from the STEP: clear the typed height, or Build Solution again.")
    rec_f.write_text(json.dumps({"schema": "pcb-thermal-ansys/solution-applied", "at": time.strftime("%Y-%m-%d %H:%M"),
                                 "sentAt": cad.get("sentAt"), "plates": done}, indent=1), encoding="utf-8")
    off = [s for s in done if not on[s]]
    check("ok", f"Thermal solution: {len(done)} plate{'s' if len(done) != 1 else ''}, {len(want)} TIMs in AEDT"
                + (f"; {' and '.join(off)} out of the model (off on the page)" if off else ""))


def sol_objects(ic):
    """the thermal solution in AEDT: [{name, model, box}] of the SOL_ objects"""
    out = []
    for n in sorted(o for o in ic.modeler.object_names if o.startswith(SOL)):
        o = ic.modeler[n]
        out.append({"name": n, "model": bool(o.is_model), "box": [round(v, 3) for v in o.bounding_box]})
    return out


def thermal(ic):
    """{boundary name: {type, power W, rjc, rjb}} from the model tree: about 0.1 s each, ic.boundaries takes 5 to 10 s."""
    node = ic.get_oo_object(ic.odesign, "Thermal")
    val = lambda n, prop: ic.get_oo_property_value(node, n, prop)   # noqa: E731
    out = {}
    for n in ic.get_oo_name(ic.odesign, "Thermal"):
        typ = val(n, "Type")
        d = {"type": typ, "power": None, "rjc": None, "rjb": None}
        if typ == "Solid Block":
            d["power"] = num(val(n, "Total Power"))
        elif typ == "Network":
            d["power"] = num(val(n, "Internal: Properties/Power"))
            d["rjc"] = num(val(n, "Rjc: Properties/ThermalResistance"))
            d["rjb"] = num(val(n, "Rjb: Properties/ThermalResistance"))
        out[n] = d
    return node, out


def _delete_boundary(ic, name):
    ic.odesign.GetModule("BoundarySetup").DeleteBoundaries([name])


def _delete_block(ic, ref):
    if f"P_{ref}" in ic.get_oo_name(ic.odesign, "Thermal"):     # a block without power has none
        _delete_boundary(ic, f"P_{ref}")
    if ref in ic.modeler.object_names:
        ic.modeler.delete(ref)


PCB = "PCB1"


def board_z(ic, board):
    """lowest and highest z of the imported PCB in mm"""
    try:
        bb = ic.modeler.user_defined_components[PCB].bounding_box
        return bb[2], bb[5]
    except BaseException:  # noqa: BLE001
        return 0.0, board["thicknessMm"]


def _kind(p):
    rjc, rjb = p.get("rjcCW") or 0, p.get("rjbCW") or 0
    if rjc > 0 and rjb > 0:
        return "Network"
    if rjb > 0:
        check("warn", f"{p['refDes']}: Rjb without Rjc. A junction model needs both, so a plain block of the package "
                      f"material is used and Rjb is not in AEDT.")
    return "Solid Block"


def _rjc_k(rjc, l, w, h):
    """k (W/m.K) of a block l x w x h mm whose resistance from the top face to the bottom face is rjc C/W"""
    return h * 1000 / (rjc * l * w)


def _block_rjc(ic, obj, h):
    """Rjc of a block of its own rjc_ material, back from AEDT: H / (k x footprint), the footprint as volume / H
    (a turned block's bounding box is not its footprint)"""
    try:
        k = float(ic.materials.material_keys[obj.material_name.lower()].thermal_conductivity.value)
        return round(h * h * 1000 / (k * obj.volume), 6) if k > 0 and obj.volume > 0 else None
    except BaseException:  # noqa: BLE001  material or volume not readable
        return None


def _rjc_material(ic, ref, k):
    """the block's own material rjc_<ref> with conductivity k; returns its name, or None when it cannot be made"""
    name = f"{RJC_MAT}{ref}".lower()
    m = ic.materials.material_keys.get(name)
    if not m:
        m = ic.materials.add_material(name)
        if not m:
            return None
        m.mass_density = 2000
        m.specific_heat = 1000
    if abs(float(m.thermal_conductivity.value) - k) > 1e-9 * max(1, k):     # a changed material voids the solution
        m.thermal_conductivity = k
    return name


def _same(a, b):
    return a is not None and b is not None and abs(a - b) < 1e-9


def _passive(ic, tree, board):
    """blocks without power: objects named after a board part that have no P_ boundary"""
    refs = {p["refDes"] for p in board["parts"]}
    return [o for o in ic.modeler.object_names if o in refs and f"P_{o}" not in tree]


def _footprints(parts, by_ref):
    """(cx, cy, l, w, angle) per block: the body size typed on the page, else the STEP's, else the EDB's, which is the
    pad extents, not the package body; l and w along the part's own axes, turned angle degrees from the board's (the
    STEP body's angle, else the pads'). Close parts can overlap, and Icepak refuses to solve intersecting blocks: two
    turned the same way that overlap are cut back to meet halfway; two turned differently only get a warning."""
    rect, turn = {}, {}
    for p in parts:
        b = by_ref[p["refDes"]]
        t = b.get("stepAngleDeg", b.get("angleDeg")) or 0.0
        cx, cy = b["xMinMm"] + b["lMm"] / 2, b["yMinMm"] + b["wMm"] / 2
        lx, ly = p.get("bodyXMm") or b.get("stepXMm") or b["lMm"], p.get("bodyYMm") or b.get("stepYMm") or b["wMm"]
        u, v = _turn(cx, cy, -t)                     # the centre in the part's own frame
        rect[p["refDes"]], turn[p["refDes"]] = [u - lx / 2, v - ly / 2, u + lx / 2, v + ly / 2], t
    refs = list(rect)
    for i, a in enumerate(refs):
        for c in refs[i + 1:]:
            if by_ref[a]["side"] != by_ref[c]["side"]:
                continue
            if abs(turn[a] - turn[c]) > 0.5:
                if _crossing(_corners(rect[a], turn[a]), _corners(rect[c], turn[c])):
                    check("warn", f"{a} and {c}: bodies overlap and are turned differently ({turn[a]:g} and {turn[c]:g} deg). "
                                  f"Type smaller Body X or Y for one of them, or Icepak may refuse to solve")
                continue
            A, C = rect[a], rect[c]
            if turn[c] != turn[a]:                   # within half a degree: C seen in A's frame
                u, v = _turn(*_turn((C[0] + C[2]) / 2, (C[1] + C[3]) / 2, turn[c]), -turn[a])
                C[:] = [u - (C[2] - C[0]) / 2, v - (C[3] - C[1]) / 2, u + (C[2] - C[0]) / 2, v + (C[3] - C[1]) / 2]
                turn[c] = turn[a]
            ox, oy = min(A[2], C[2]) - max(A[0], C[0]), min(A[3], C[3]) - max(A[1], C[1])
            if ox <= 0 or oy <= 0:
                continue
            k = 0 if ox <= oy else 1                       # cut across the smaller overlap
            lo, hi = (A, C) if A[k] + A[k + 2] <= C[k] + C[k + 2] else (C, A)
            lo[k + 2] = hi[k] = (max(A[k], C[k]) + min(A[k + 2], C[k + 2])) / 2
            check("warn", f"{a} and {c}: footprints overlap by {min(ox, oy):.2f} mm, both blocks cut back to meet")
    out = {}
    for r, (u0, v0, u1, v1) in rect.items():
        cx, cy = _turn((u0 + u1) / 2, (v0 + v1) / 2, turn[r])
        out[r] = (cx, cy, u1 - u0, v1 - v0, turn[r])
    return out


def _turn(x, y, deg):
    c, s = math.cos(math.radians(deg)), math.sin(math.radians(deg))
    return x * c - y * s, x * s + y * c


def _corners(r, deg):
    """the four corners on the board of a rectangle [u0, v0, u1, v1] in a frame turned deg"""
    return [_turn(u, v, deg) for u, v in ((r[0], r[1]), (r[2], r[1]), (r[2], r[3]), (r[0], r[3]))]


def _crossing(P, Q):
    """two convex outlines overlap: no side of either separates them"""
    for poly in (P, Q):
        for (x0, y0), (x1, y1) in zip(poly, poly[1:] + poly[:1]):
            nx, ny = y0 - y1, x1 - x0
            a, b = [nx * x + ny * y for x, y in P], [nx * x + ny * y for x, y in Q]
            if max(a) <= min(b) + 1e-9 or max(b) <= min(a) + 1e-9:
                return False
    return True


def _blocks(ic, board, by_ref, parts, study):
    """make or fix one block per part. Returns the RefDes of the blocks made new (a local mesh on the old block does
    not follow the new one)"""
    rebuilt = set()
    m = ic.materials.add_material(PKG)
    if m:       # only what differs: every write is an EditMaterial, and a changed material voids the solution
        for attr, v in (("thermal_conductivity", study["packageK"]), ("mass_density", 2000), ("specific_heat", 1000)):
            try:
                same = abs(float(getattr(m, attr).value) - v) <= 1e-9 * max(1, v)
            except (TypeError, ValueError):     # not a single number (anisotropic, or not set yet)
                same = False
            if not same:
                setattr(m, attr, v)
    want = {p["refDes"] for p in parts}
    node, tree = thermal(ic)
    for name in list(tree):
        if name.startswith("P_") and name[2:] not in want:
            _delete_block(ic, name[2:])
            tree.pop(name)
            check("ok", f"{name[2:]}: heat source removed")
    for ref in _passive(ic, tree, board):
        if ref not in want:
            ic.modeler.delete(ref)
            check("ok", f"{ref}: block without power removed")
    objects = set(ic.modeler.object_names)
    zlo, zhi = board_z(ic, board)
    xy = _footprints(parts, by_ref)
    made = nets = passive = 0
    for p in parts:
        ref, b = p["refDes"], by_ref[p["refDes"]]
        h = p.get("heightMm") or b.get("stepHMm") or b["hMm"]    # typed, else the STEP's, else the EDB placeholder
        if not p.get("heightMm") and not b.get("stepHMm") and h < MIN_BLOCK_H_MM:
            check("warn", f"{ref}: EDB height {b['hMm']} mm is a placeholder, using {MIN_BLOCK_H_MM} mm. "
                          f"Type the real height.")
            h = MIN_BLOCK_H_MM
        z = zhi if b["side"] == "top" else zlo - h      # the part sits on the board surface, so a network can touch it
        cx, cy, l, w, t = xy[ref]
        xs, ys = zip(*_corners([-l / 2, -w / 2, l / 2, w / 2], t))
        box = [cx + min(xs), cy + min(ys), z, cx + max(xs), cy + max(ys), z + h]     # what AEDT reports as its bounding box
        kind = _kind(p) if (p.get("powerW") or 0) > 0 else "Passive"
        # Rjc without Rjb: the block conducts Rjc from top to bottom, as the hand check's "Block, Rjc"
        own = _rjc_material(ic, ref, _rjc_k(p["rjcCW"], l, w, h)) if kind == "Solid Block" and (p.get("rjcCW") or 0) > 0 else None
        mat = own or (PKG if m else "Ceramic_material")
        cur = tree.get(f"P_{ref}")
        ok = ref in objects and (cur is None if kind == "Passive" else cur is not None and cur["type"] == kind)
        if ok:
            ok = max(abs(a - c) for a, c in zip(ic.modeler[ref].bounding_box, box)) <= 1e-3
        if ok and kind == "Network":
            ok = _same(cur["rjc"], p["rjcCW"]) and _same(cur["rjb"], p["rjbCW"])
        if ok and kind == "Passive":
            made += 1
            passive += 1
            continue
        if ok and kind != "Network" and ic.modeler[ref].material_name.lower() != mat.lower():
            ic.modeler[ref].material_name = mat
        if ok:
            if not _same(cur["power"], p["powerW"]):
                prop = "Total Power" if kind == "Solid Block" else "Internal: Properties/Power"
                ic.get_oo_object(node, f"P_{ref}").SetPropValue(prop, f"{p['powerW']}W")
            made += 1
            nets += kind == "Network"
            continue
        if cur is not None or ref in objects:
            _delete_block(ic, ref)
        obj = ic.modeler.create_box([-l / 2, -w / 2, z], [l, w, h], name=ref, material=mat)
        rebuilt.add(ref)
        if not obj:
            check("error", f"{ref}: block not created")
            continue
        if abs(t) > 0.05 and not obj.rotate("Z", t):          # made at the origin, turned there, then moved onto its pads
            check("warn", f"{ref}: block not turned by {t:g} deg, left square to the board")
        obj.move([cx, cy, 0])
        if kind == "Passive":
            made += 1
            passive += 1
            continue
        if kind == "Network":
            try:
                net = ic.create_two_resistor_network_block(ref, PCB, f"{p['powerW']}W", p["rjbCW"], p["rjcCW"])
                net.name = f"P_{ref}"
                obj.solve_inside = False        # a network block must not be meshed inside, or the solver input fails
                nets += 1
                made += 1
                continue
            except BaseException as ex:  # noqa: BLE001
                check("warn", f"{ref}: junction model not created ({type(ex).__name__}), plain block used")
        if ic.assign_solid_block(obj.name, f"{p['powerW']}W", boundary_name=f"P_{ref}"):
            made += 1
        else:
            check("error", f"{ref}: power not assigned")
    if parts:
        check("ok" if made == len(parts) else "error",
              f"{made} of {len(parts)} parts in AEDT: {nets} with a junction model (Rjc, Rjb), {passive} without power")
    return rebuilt


def _air(ic, study, board):
    side, above = study["marginMm"], study["marginAboveMm"]
    pad = [side] * 6                           # +X -X +Y -Y +Z -Z
    pad[ORIENTATION[study["orientation"]][3]] += above   # extra room on the face opposite to gravity
    bb, have = air_pad(ic, board)
    if have and "Open_air" in ic.get_oo_name(ic.odesign, "Thermal") and \
            max(abs(a - b) for a, b in zip(have, pad)) <= 0.06:   # rebuilding would throw away the mesh and solution
        check("ok", f"Air {bb[3]-bb[0]:.0f} x {bb[4]-bb[1]:.0f} x {bb[5]-bb[2]:.0f} mm, open on all faces (unchanged)")
        return
    ic.modeler.delete("Region")
    reg = ic.modeler.create_region(pad_value=pad, pad_type="Absolute Offset", name="Region")
    if not reg:
        check("error", "Air region not created")
        return
    if "Open_air" in ic.get_oo_name(ic.odesign, "Thermal"):
        _delete_boundary(ic, "Open_air")
    ic.assign_free_opening([f.id for f in reg.faces], boundary_name="Open_air")
    bb = reg.bounding_box
    check("ok", f"Air {bb[3]-bb[0]:.0f} x {bb[4]-bb[1]:.0f} x {bb[5]-bb[2]:.0f} mm, open on all faces")


def _environment(ic, study):
    gdir, label, what, _ = ORIENTATION[study["orientation"]]
    # radiation_temperature must be given: PyAEDT sets 20 C otherwise, and a colder surround over-states radiation
    ic.edit_design_settings(gravity_dir=gdir, ambient_temperature=study["ambientC"], radiation_temperature=study["ambientC"])
    check("ok", f"Ambient {study['ambientC']} C (air and radiation surround), gravity {label} ({what})")


def _setup(ic, study):
    s = next((x for x in ic.setups if x.name == SETUP), None)
    existed = bool(s)
    s = s or ic.create_setup(SETUP)
    if not s:
        check("error", "Setup not created")
        return
    want = [("Flow Regime", "Laminar"), ("Include Flow", True), ("Include Gravity", True),
            ("Convergence Criteria - Max Iterations", int(study["iterations"]))]
    for k, f in (("Convergence Criteria - Flow", "flowLimit"), ("Convergence Criteria - Energy", "energyLimit")):
        if (study.get(f) or 0) > 0:      # left out: AEDT keeps its own limit
            want.append((k, f"{study[f]:g}"))
    for k, v in want:
        if k in s.props:
            s.props[k] = v
    was = str(s.props.get("Radiation Model", "Off"))
    s.props["Radiation Model"] = "Discrete Ordinates Model" if study["radiation"] else "Off"
    s.update()
    check("ok", f"Setup '{SETUP}': laminar natural convection, radiation "
                f"{'on (discrete ordinates)' if study['radiation'] else 'off'}, {study['iterations']} iterations max, "
                f"limits flow {s.props.get('Convergence Criteria - Flow')}, energy {s.props.get('Convergence Criteria - Energy')}")
    if existed and was != s.props["Radiation Model"]:
        # the next solve would restart from the old case file, which keeps the old radiation model
        if step("clear old solution", lambda: ic.cleanup_solution()):
            check("ok","Radiation changed: old solution deleted, the next solve starts fresh")


def _mesh(ic, parts, study, rebuilt=()):
    """global mesh and one local mesh region per part with meshMm. `rebuilt`: blocks _blocks made new, whose old
    region is made again"""
    m = study["mesh"]
    gm = ic.mesh.global_mesh_region
    manual = m.get("mode") == "manual"
    want = ({**{f"MaxElementSize{ax}": v for ax, v in zip("XYZ", _max_xyz(m))}, "MinElementsInGap": m.get("gapElems", 3),
             "MinElementsOnEdge": m.get("edgeElems", 2), "MaxSizeRatio": m.get("ratio", 2)}
            if manual else {"MeshRegionResolution": int(m.get("level", 3))})
    same = bool(gm.manual_settings) == manual and all(num(_get(gm.settings, k)) == float(v) for k, v in want.items())
    if not same:
        gm.manual_settings = manual
        for k, v in want.items():
            gm.settings[k] = f"{v}mm" if k.startswith("MaxElementSize") else int(v) if k == "MeshRegionResolution" else str(v)
        gm.update()
    # a junction model block is hollow (not meshed inside): a mesh region on it leaves it with no mesh at all and the
    # solver stops with an execution error (seen 2026-11-02). Taken from AEDT's block type, not the page's Rjc and Rjb:
    # a junction model AEDT refused is a plain block, which can take a local mesh
    nets = _networks(ic)
    hollow = sorted(p["refDes"] for p in parts if p.get("meshMm") and p["refDes"] in nets)
    if hollow:
        check("warn", f"{', '.join(hollow)}: no local mesh. A junction model (Rjc and Rjb) is not meshed inside, and a "
                      f"local mesh on it stops the solver. The global mesh is used there.")
    wanted = {p["refDes"]: p["meshMm"] for p in parts if p.get("meshMm") and p["refDes"] not in nets}
    have = local_mesh(ic)
    stale = [ref for ref, size in have.items() if wanted.get(ref) != size or ref in rebuilt]
    _drop_local(ic, stale)
    for ref in stale:
        have.pop(ref)
    for ref, size in wanted.items():
        if ref in have:
            continue
        r = ic.mesh.assign_mesh_region([ref], name=f"Mesh_{ref}")
        if r:
            r.manual_settings = True
            for ax in "XYZ":
                r.settings[f"MaxElementSize{ax}"] = f"{size}mm"
            r.update()
        else:
            check("error", f"{ref}: local mesh not created")
    check("ok", f"Mesh: global {_mesh_text(m)}, {len(wanted)} local refinements")


def _max_xyz(m):
    """largest element X, Y, Z of a manual mesh in mm. An older study.json has one maxMm for all three"""
    return [m.get(k) or m["maxMm"] for k in ("maxMm", "maxYMm", "maxZMm")]


def _manual_read(gm):
    """the manual settings of the global mesh region, in the keys of study.json mesh"""
    g = lambda k: num(_get(gm.settings, k))
    return {"maxMm": g("MaxElementSizeX"), "maxYMm": g("MaxElementSizeY"), "maxZMm": g("MaxElementSizeZ"),
            "gapElems": g("MinElementsInGap"), "edgeElems": g("MinElementsOnEdge"), "ratio": g("MaxSizeRatio")}


def local_mesh(ic):
    """{RefDes: max element mm} of the local mesh regions, from the model tree"""
    node = ic.get_oo_object(ic.odesign, "Mesh")
    return {n[5:]: num(ic.get_oo_property_value(node, n, "Maximum Element Size/X"))
            for n in ic.get_oo_name(ic.odesign, "Mesh") if n.startswith("Mesh_") and not n.endswith("_SubRegion")}


def _networks(ic, tree=None):
    """RefDes of the junction model (Network) blocks in AEDT"""
    if tree is None:
        _, tree = thermal(ic)
    names = set(ic.modeler.object_names)
    return {n[2:] for n, d in tree.items() if n.startswith("P_") and d["type"] == "Network" and n[2:] in names}


def _drop_local(ic, refs):
    """delete the local mesh regions of `refs`, then every sub-region box no mesh region uses. AEDT keeps a deleted
    region's box, renamed SubRegion<n>: left there it overlaps the part's new box, the part gets no mesh and the solve
    fails. Boxes left by earlier runs go too. Returns how many boxes were deleted"""
    refs = set(refs)
    if refs:
        for r in list(ic.mesh.meshregions):
            if r.name.startswith("Mesh_") and not r.name.endswith("_SubRegion") and r.name[5:] in refs:
                try:
                    r.delete()
                except BaseException:  # noqa: BLE001  AEDT removes the sub-region together with the region
                    pass
    ic.modeler.refresh_all_ids()
    node = ic.get_oo_object(ic.odesign, "Mesh")
    used = set()
    for n in ic.get_oo_name(ic.odesign, "Mesh"):
        try:
            a = ic.get_oo_property_value(node, n, "Assignment")
        except BaseException:  # noqa: BLE001
            if n.startswith("Mesh_"):
                return 0        # a region whose box is not known: delete nothing rather than its box
            continue            # the global region has no sub-region
        used.update(a if isinstance(a, (list, tuple)) else [s.strip() for s in str(a).split(",")])
    orphans = [o for o in ic.modeler.object_names if "subregion" in o.lower() and o not in used]
    for o in orphans:
        ic.modeler.delete(o)
    if orphans:
        ic.modeler.refresh_all_ids()
    return len(orphans)


def _mesh_now(ic):
    t = time.time()
    ok = ic.mesh.generate_mesh(SETUP)
    check("ok" if ok else "error", f"Mesh generated in {time.time() - t:.0f} s" if ok else "Mesh generation failed")
    return ok


def _meshed(folder, sig, how):
    (folder / "meshed.json").write_text(json.dumps({"schema": "pcb-thermal-ansys/meshed", "sig": sig, "how": how,
                                                    "at": time.strftime("%H:%M")}, indent=1), encoding="utf-8")


def _mesh_sig(geo, mesh, pad):
    """what the mesh depends on that the page can change: block boxes, air box, mesh settings"""
    return hashlib.sha1(json.dumps([geo, mesh, pad], sort_keys=True).encode()).hexdigest()[:16]


# ---------------------------------------------------------------- solve

def _hollow(ic):
    """network blocks must not be meshed inside and must have no local mesh (it stops the solver), also in projects
    built before this was set; sub-region boxes no region uses go too. Returns True if the mesh changed"""
    nets = _networks(ic)
    for ref in nets:
        if ic.modeler[ref].solve_inside:
            ic.modeler[ref].solve_inside = False
    bad = sorted(set(local_mesh(ic)) & nets)
    boxes = _drop_local(ic, bad)
    if bad:
        check("warn", f"{', '.join(bad)}: local mesh removed. A junction model is not meshed inside, and a local mesh on "
                      f"it stops the solver. The global mesh is used there.")
    if boxes:
        check("warn", f"{boxes} leftover mesh sub-region box{'es' if boxes > 1 else ''} removed (left by a deleted local "
                      f"mesh; it would leave its part with no mesh)")
    return bool(bad or boxes)


def _sources(ic, powered):
    """heat sources as they are in AEDT, which is what gets solved: [{refDes, powerW, network, rjaCW, tjMaxC}].
    Rja and Tj max stay on the page, so they come from study.json."""
    _, tree = thermal(ic)
    names = set(ic.modeler.object_names)
    page = {p["refDes"]: p for p in powered}
    out = []
    for n, d in tree.items():
        if n.startswith("P_") and n[2:] in names and d["type"] in ("Solid Block", "Network"):
            p = page.get(n[2:], {})
            out.append({"refDes": n[2:], "powerW": d["power"] or 0, "network": d["type"] == "Network",
                        "rjaCW": p.get("rjaCW"), "tjMaxC": p.get("tjMaxC")})
    missing = sorted(set(page) - {s["refDes"] for s in out})
    if missing:
        check("warn", f"{', '.join(missing)}: on the page but not in AEDT, so not solved. Press Apply first.")
    return out


def _monitors(ic, sources, board):
    """one point monitor T_<ref> per heat source, made fresh for each solve. A network block is hollow, so its
    monitor sits on the case side face (the face away from the board), a plain block gets it at its centre.
    The Q_ heat flow face monitors older versions made go too."""
    for n in [n for n in ic.get_oo_name(ic.odesign, "Monitor") if n.startswith(("T_", "Q_"))]:
        ic.monitor.delete_monitor(n)
    _, zhi = board_z(ic, board)
    for s in sources:
        ref = s["refDes"]
        bb = ic.modeler[ref].bounding_box
        if s["network"]:
            pos = [(bb[0] + bb[3]) / 2, (bb[1] + bb[4]) / 2, bb[5] if bb[2] >= zhi - 1e-6 else bb[2]]
            ic.monitor.assign_point_monitor(pos, monitor_quantity="Temperature", monitor_name=f"T_{ref}")
        else:
            ic.monitor.assign_point_monitor_in_object(ref, monitor_quantity="Temperature", monitor_name=f"T_{ref}")


def _fmax(ic, obj, op="Maximum"):
    """highest (op "Mean": volume average) temperature in one object (field calculator). The solution name must be
    given, or AEDT refuses. Not post.get_scalar_field_value: it tries EnterQty("Temperature") first, which in Icepak
    always fails (Temperature is only a named expression) and puts "Abnormal script termination" in the message manager
    for every object. A hollow object (network block) has no volume and fails; the stack is cleared so the next call
    still works."""
    fr = ic.post.ofieldsreporter
    f = Path(ic.working_directory) / f"fmax_{obj}.fld"
    try:
        fr.CalcStack("clear")
        fr.CopyNamedExprToStack("Temperature")
        fr.EnterVol(obj)
        fr.CalcOp(op)
        fr.CalculatorWrite(str(f), ["Solution:=", f"{SETUP} : SteadyState"], [])
        return round(float(f.read_text().split()[-1]), 2)
    except BaseException:  # noqa: BLE001
        return None
    finally:
        try:
            fr.CalcStack("clear")
            f.unlink(missing_ok=True)
        except BaseException:  # noqa: BLE001
            pass


def _temp(ic, ref, hollow, monitor=None):
    """temperature of a heat source: field maximum for a block, case-side monitor for a hollow network block.
    The field calculator must never see a hollow block: the failed call breaks the whole AEDT session.
    `monitor`: the monitor's final value when it was already read"""
    v = None if hollow else _fmax(ic, ref)
    if v is not None:
        return v
    if monitor is not None:
        return round(monitor, 2)
    try:
        return round(_final(ic, f"T_{ref}.Temperature", "Monitor"), 2)
    except BaseException:  # noqa: BLE001  one missing value must not stop the others
        return None


ENTHALPY_REF_C = 25.0   # the solver counts the heat the air carries from 298.15 K, whatever the ambient
AIR_CP = 1006.43        # J/(kg K), Icepak's air


def _balance(ic, total_w, ambient):
    """heat out through the Open_air faces against the power put in, from AEDT's Fields Summary.
    Face monitors count only convection; the summary also has the radiation that leaves through the opening.
    The convective flow is counted from 25 C, not from the ambient: air that comes in and goes out in slightly
    different amounts (the leftover mass flow of a converged run) then adds mass flow x cp x (ambient - 25 C), several %
    at 55 C ambient and nothing at 25 C. That part is taken off, so the balance checks the physics."""
    fs = ic.post.create_field_summary()
    qs = {"outW": "HeatFlowRate", "radiationW": "RadiationFlow", "convectionW": "ConvectiveHeatFlow"}
    for q in [*qs.values(), "MassFlowRate"]:
        fs.add_calculation("Boundary", "Surface", "Open_air", q)
    data = fs.get_field_summary_data(setup=f"{SETUP} : SteadyState")
    got, mass = {}, None
    for name, total in zip(data["Quantity"], data["Total"]):    # "HeatFlowRate[W]" -> "3.30061", signed along the face normals
        q, unit = name.rstrip("]").split("[")
        if q == "MassFlowRate":
            u = unit.lower().replace("_per_", "/")
            scale = 1.0 if u.startswith("kg") else 1e-3 if u.startswith("g") else None
            mass = float(total) * scale if scale else None
        else:
            got[q] = float(total) / (1000.0 if unit == "mW" else 1.0)
    # MassFlowRate is signed opposite to the heat flows on the same faces: with the same sign the 55 C run of
    # 2026-11-02 went from AEDT's +6.5 % to +12.6 %; with the opposite sign to +0.33 %, as the same board at 25 C (+0.35 %)
    fix = -mass * AIR_CP * (ambient - ENTHALPY_REF_C) if mass is not None else 0.0
    raw_conv, raw_out = got["ConvectiveHeatFlow"], got["HeatFlowRate"]
    bal = {"outW": round(abs(raw_out - fix), 4), "radiationW": round(abs(got["RadiationFlow"]), 4),
           "convectionW": round(abs(raw_conv - fix), 4)}
    return {"inW": round(total_w, 4), **bal, "errorPct": round(100 * (bal["outW"] - total_w) / total_w, 2),
            "rawOutW": round(abs(raw_out), 4), "netMassKgS": mass, "refFixW": round(abs(fix), 4) if mass is not None else None}


def _convergence(ic, hist):
    """final residuals against the setup's convergence criteria: {"ok": bool, "items": [{name, value, limit, ok, history}]}.
    `hist` is _solver_files(ic): the residual files count every iteration since the solution started, a restart
    included; the monitors cover this solve only, from its first new iteration, so where they start tells a continued
    solve from a fresh one. With no monitor history this solve's share is not known: `iterations` and `continued`
    are None then, `totalIterations` is still given"""
    setup = next((x for x in ic.setups if x.name == SETUP), None)
    limit = {"Continuity": "Flow", "XVelocity": "Flow", "YVelocity": "Flow", "ZVelocity": "Flow", "Energy": "Energy"}
    if setup and str(setup.props.get("Radiation Model", "Off")) != "Off":
        limit["DOIntensity"] = "Discrete Ordinates"
    res, mons = hist
    items = []
    for q, crit in limit.items():
        try:
            val = _final(ic, "Residual." + q, "Residual")
            lim = num(setup.props.get(f"Convergence Criteria - {crit}")) if setup else None
            h = res.get(q, [])
            items.append({"name": q, "value": val, "limit": lim, "ok": lim is None or val <= lim,
                          **({"history": _thin(h)} if len(h) > 1 else {})})
        except BaseException:  # noqa: BLE001
            items.append({"name": q, "value": None, "limit": None, "ok": None})
    known = [i for i in items if i["ok"] is not None]
    most = num(setup.props.get("Convergence Criteria - Max Iterations")) if setup else None
    total = int(max(h[-1][0] for h in res.values() if h)) if any(res.values()) else None
    start = min((h[0][0] for h in mons.values() if h), default=None)
    # a fresh solve's monitors start at iteration 1; a continued one's at the first iteration after the one it went on
    # from (kept run 20261102-145928: residuals to 424, monitors 423 and 424, so 2 iterations in that solve)
    run = total - int(start) + 1 if total and start else None
    return {"ok": bool(known) and all(i["ok"] for i in known), "read": len(known), "items": items,
            "iterations": run, "totalIterations": total, "continued": None if run is None else run != total,
            "maxIterations": int(most) if most else None, "hitMax": bool(run and most and run >= most)}


def _final(ic, expr, category):
    """the final value of a residual or monitor, as AEDT reports it"""
    sd = ic.post.get_solution_data(expressions=[expr], setup_sweep_name=f"{SETUP} : SteadyState", report_category=category)
    return float(sd.get_expression_data(sd.expressions[0])[1][-1])


def _solver_files(ic):
    """residuals and temperature monitors per iteration, read from the files the solver wrote: ({residual: [[it, v]]},
    {monitor name: [[it, v]]}). AEDT's report API gives only the final value however the sweep is asked for (one
    point at X = 0, 2026-11-02), so read the files `<design>.asol` names for the setup's solution: DV*_S*_MON0_*.sd
    holds every residual of every iteration since the solution started (a restart adds to it), DV*_S*_MON<id>_*.sd
    one monitor over this solve only (a continued solve starts at its first new iteration). The monitor ids are the
    ID= of each monitor in the saved .aedt, so each history is matched to its monitor by name.
    Read after the project is saved: before that the .asol names the last solve's files"""
    root = Path(ic.results_directory)
    asol = root / f"{DESIGN}.asol"
    if not asol.exists():
        asol = next(root.rglob(f"{DESIGN}.asol"), None)
        if not asol:
            return {}, {}
    sd_dir = asol.parent / f"{DESIGN}.results"
    names = re.findall(r"File='(DV\d+_S\d+_MON(\d+)_V\d+\.sd)'", asol.read_text(encoding="utf-8", errors="replace"))
    ids = {i: n for n, i in _monitor_ids(ic.project_file).items()}
    res, mons = {}, {}
    for name, mid in names:
        f = sd_dir / name
        if not f.exists():
            continue
        rows = {}
        for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
            head, _, rest = line.strip().partition(" ")
            try:
                it = float(head)
                if it < 1:      # the solution block's own file holds only the final value, at iteration -1
                    continue
                for q, v in re.findall(r"(\w+)\(([^)]*)\)", rest):
                    rows.setdefault(q, []).append([it, float(v)])
            except ValueError:
                continue
        if mid == "0":
            res = rows
        elif rows.get("Temperature") and ids.get(int(mid), "").startswith("T_"):
            mons[ids[int(mid)]] = rows["Temperature"]
    return res, mons


def _monitor_ids(project_file):
    """{monitor name: id} from the IcepakMonitors block of the saved project file"""
    txt = Path(project_file).read_bytes().decode("latin-1")
    i = txt.find("$begin 'IcepakMonitors'")
    j = txt.find("$end 'IcepakMonitors'", i)
    if i < 0 or j < 0:
        return {}
    out, stack = {}, []
    for line in txt[i:j].splitlines()[1:]:
        s = line.strip()
        m = re.fullmatch(r"\$begin '(.*)'", s)
        if m:
            stack.append(m.group(1))
        elif s.startswith("$end "):
            if stack:
                stack.pop()
        elif len(stack) == 1 and (m := re.fullmatch(r"ID=(\d+)", s)):
            out[stack[0]] = int(m.group(1))
    return out


def _thin(h, most=200):
    """at most `most` points of a history [[x, y]], the last one always kept: enough for a plot in the page or a report"""
    k = max(1, -(-len(h) // most))
    idx = list(range(0, len(h), k))
    if idx[-1] != len(h) - 1:
        idx.append(len(h) - 1)
    return [h[i] for i in idx]


DRIFT_K = 0.1     # a monitor that still moves more than this over the last 10 % of the iterations is not settled


def _monitor_health(ic, powered, hist):
    """each heat source's temperature monitor per iteration, and how far it still moved over the last 10 % of this
    solve's iterations (at least its last step). Residuals under their limits do not prove the temperatures stopped
    moving; this does. Each history (matched by monitor name) is checked against AEDT's own final value before it is
    used. Returns (monitors, {RefDes: final value}); `read` False on a monitor: no history found or it did not match;
    True with ok None: one iteration only, nothing to compare"""
    out, finals = [], {}
    for p in powered:
        ref = p["refDes"]
        try:
            final = finals[ref] = _final(ic, f"T_{ref}.Temperature", "Monitor")
        except BaseException:  # noqa: BLE001  one missing monitor must not stop the others
            out.append({"refDes": ref, "driftK": None, "ok": None, "read": False})
            continue
        h = hist[1].get(f"T_{ref}", [])
        if not h or abs(h[-1][1] - final) > 1e-6 * max(1.0, abs(final)):
            out.append({"refDes": ref, "driftK": None, "ok": None, "read": False})
        elif len(h) < 2:
            out.append({"refDes": ref, "driftK": None, "ok": None, "read": True})
        else:
            tail = h[min(int(len(h) * 0.9), len(h) - 2)][1]
            drift = round(abs(h[-1][1] - tail), 3)
            out.append({"refDes": ref, "driftK": drift, "ok": drift <= DRIFT_K, "read": True, "history": _thin(h)})
    return out, finals


def _profile(ic, since):
    """mesh size, memory, cores and end status from the solver profile AEDT writes next to the project.
    Only a profile written by this solve (after `since`) counts"""
    files = [f for f in Path(ic.results_directory).rglob("*.profile") if f.stat().st_mtime >= since - 5]
    if not files:
        return None
    txt = max(files, key=lambda f: f.stat().st_mtime).read_text(encoding="utf-8", errors="replace")
    grab = lambda pat: (re.search(pat, txt) or [None, None])[1]  # noqa: E731
    cells, nodes = grab(r"Total Cells\\?',\s*(\d+)"), grab(r"Total Nodes\\?',\s*(\d+)")
    return {"cells": int(cells) if cells else None, "nodes": int(nodes) if nodes else None,
            "memory": grab(r"'ComEngine Memory',\s*'([^']+)'"), "elapsed": grab(r"'Elapsed Time',\s*'([^']+)'"),
            "cores": int(grab(r"Cores\\?',\s*(\d+)") or 0) or None, "status": grab(r"Status\\?',\s*\\?'([^'\\]+)")}


def _dur(secs):
    """a solve time as people read it: '45 s' under a minute, '5 min', '1 h 12 min' from an hour (as the page)"""
    if secs < 60:
        return f"{secs:.0f} s"
    m = round(secs / 60)
    return f"{m} min" if m < 60 else f"{m // 60} h {m % 60} min"


def _setup_now(ic):
    """the values the solve will use, as AEDT holds them: setup, ambient, gravity, mesh. Not study.json, which holds
    the page's values whether applied or not. Read before the solve: after it AEDT cannot be read"""
    out = {}
    s = next((x for x in ic.setups if x.name == SETUP), None)
    if s:
        p = s.props
        out.update(iterations=num(p.get("Convergence Criteria - Max Iterations")), flowLimit=num(p.get("Convergence Criteria - Flow")),
                   energyLimit=num(p.get("Convergence Criteria - Energy")), radiation=str(p.get("Radiation Model", "Off")))
    ds = ic.design_settings
    out.update(ambientC=num(ds["AmbTemp"]), orientation=_gravity(ds))
    gm = ic.mesh.global_mesh_region
    out["mesh"] = ({"mode": "manual", **_manual_read(gm)}
                   if gm.manual_settings else {"mode": "auto", "level": int(num(_get(gm.settings, "MeshRegionResolution"), 3))})
    out["localMm"] = local_mesh(ic)
    return out


def _keep_run(folder, res, study, parts, powered, setup, label):
    """every solve also goes to runs/<date-time>.json with the settings it ran with, so the page can compare runs.
    results.json stays the last solve only. `setup` (_setup_now) holds them as AEDT had them; the page's values are
    only the fallback when AEDT could not be read. The PCB mapping is the one Build used"""
    runs = folder / "runs"
    runs.mkdir(exist_ok=True)
    rid = time.strftime("%Y%m%d-%H%M%S")
    built = json.loads((folder / "built.json").read_text(encoding="utf-8")) if (folder / "built.json").exists() else {}
    (runs / f"{rid}.json").write_text(json.dumps({
        "schema": "pcb-thermal-ansys/run", "at": time.strftime("%Y-%m-%d %H:%M"), "label": label, "note": "",
        "settings": {"ambientC": study["ambientC"], "orientation": study["orientation"],
                     "pcbResolution": built.get("pcbResolution", study["pcbResolution"]),
                     "mesh": study["mesh"], "localMm": {p["refDes"]: p["meshMm"] for p in parts if p.get("meshMm")},
                     **setup, "totalW": round(sum(p["powerW"] for p in powered), 4)},
        "result": res}, indent=1), encoding="utf-8")
    return rid


def _solve(ic, folder, board, powered, study, parts, write=True, label="Solve"):
    """solve and read the results; written to results.json unless `write` is False, and always kept in runs/.
    Returns them, None if it failed"""
    if step("junction blocks hollow", lambda: _hollow(ic)):
        # the mesh changed: read it now (AEDT cannot be read after the solve), so meshed.json gets the new meshSig
        state = step("read back", lambda: read_state(ic, board))
        if state:
            (folder / "state.json").write_text(json.dumps(state, indent=1), encoding="utf-8")
    powered = _sources(ic, powered)     # from here on the heat sources are the ones in AEDT, with AEDT's power
    step("temperature monitors", lambda: _monitors(ic, powered, board))
    setup = step("setup values", lambda: _setup_now(ic)) or {}
    ambient = setup["ambientC"] if setup.get("ambientC") is not None else study["ambientC"]    # what AEDT solves with
    t = time.time()
    ok = ic.analyze_setup(SETUP)
    secs = time.time() - t
    if not ok:
        check("error", "Solve failed or was stopped. Open AEDT and read the message window.")
        return None
    check("ok", f"Solved in {_dur(secs)}")
    step("save project", lambda: ic.save_project())     # the .asol names this solve's files only once saved
    try:    # only for the plots and the iteration count: a failed read is a warning, the result stands
        hist = _solver_files(ic)
    except BaseException as ex:  # noqa: BLE001
        hist = ({}, {})
        check("warn", f"Solver histories not read ({type(ex).__name__}): no plots, settled check or iteration count")
    conv = _convergence(ic, hist)
    n, total = conv["iterations"], conv["totalIterations"]
    used = (f", after {n} more iteration{'s' if n != 1 else ''} ({total} in all; it went on from the last solution, "
            f"limit {conv['maxIterations']} per solve)" if n and conv["continued"]
            else f", after {n} of {conv['maxIterations']} iterations" if n
            else f", after {total} iterations in all (how many in this solve is not known: no monitor history)" if total
            else "")
    if not conv["read"]:
        check("warn", "Convergence could not be read. The result files may be stale (AEDT was killed or crashed earlier). "
                      "Press Solve again; if it stays, check the residuals in AEDT before trusting the numbers.")
    elif conv["ok"]:
        check("ok", "Converged: all residuals are below the setup limits" + used)
    else:
        bad = ", ".join(i["name"] for i in conv["items"] if i["ok"] is False)
        full = f"; it used all {n} iterations: raise Iterations max" if conv["hitMax"] else used
        check("warn", f"NOT converged ({bad}){full}. Raise the iterations or refine the mesh before trusting the numbers.")
    mon, finals = step("monitor histories", lambda: _monitor_health(ic, powered, hist)) or ([], {})
    span = "the last 10 % of this solve's iterations"
    moving = [f"{m['refDes']} {m['driftK']} K" for m in mon if m["ok"] is False]
    if moving:
        check("warn", f"Temperatures still moving over {span}: {', '.join(moving)} "
                      f"(more than {DRIFT_K} K). Tighten the limits or raise the iterations.")
    elif any(m["ok"] for m in mon):
        most = max(m["driftK"] for m in mon if m["ok"])
        check("ok", f"Temperatures settled: every monitor moved {most} K or less over {span} (limit {DRIFT_K} K)")
    if not any(i.get("history") for i in conv["items"]) or any(m["read"] is False for m in mon):
        check("warn", "History per iteration not read for every item: the solver files were not found or did not match "
                      "AEDT's final values, so the plots and the settled check may be missing.")
    elif any(m["ok"] is None for m in mon):
        check("ok", "Temperatures not checked for settling: this solve ran one iteration only (it went on from a "
                    "solution that had already converged)")
    prof = step("solver profile", lambda: _profile(ic, t))
    if prof:
        check("ok" if (prof["status"] or "").startswith("Normal") else "warn",
              f"Mesh {prof['cells']:,} cells, {prof['nodes']:,} nodes; solver memory {prof['memory']}, {prof['cores']} cores, "
              f"{prof['status'] or 'status not read'}" if prof["cells"] and prof["nodes"] else
              f"Solver profile read, but no mesh size in it; status {prof['status'] or 'not read'}")
    total_w = sum(p["powerW"] for p in powered)
    bal = step("energy balance", lambda: _balance(ic, total_w, ambient)) if total_w > 0 else None
    if bal:
        e = abs(bal["errorPct"])
        check("ok" if e <= 1 else "warn", f"Energy balance: {bal['inW']} W in, {bal['outW']} W out "
              f"({bal['convectionW']} W convection, {bal['radiationW']} W radiation, {bal['errorPct']:+.2f} %)"
              + (f"; {bal['refFixW']} W of the {ambient - ENTHALPY_REF_C:g} K air reference taken off ({bal['rawOutW']} W as AEDT reports it)"
                 if bal["refFixW"] else "")
              + ("" if e <= 1 else ". More than 1 % is not balanced: converge further or refine the mesh."))
        if bal["netMassKgS"] is None and abs(ambient - ENTHALPY_REF_C) > 1:
            check("warn", f"Net air mass flow not read: the balance at {ambient} C ambient is not corrected for the solver's "
                          f"25 C air reference and may be off by a few % for that reason alone.")
    rows = [{"refDes": p["refDes"], "powerW": p["powerW"], "tempC": _temp(ic, p["refDes"], p["network"], finals.get(p["refDes"])),
             "rjaCW": p["rjaCW"], "tjMaxC": p["tjMaxC"]} for p in powered]
    if any(r["tempC"] is None for r in rows):
        check("warn", "Some temperatures could not be read. Look at the result in AEDT.")
    layers = ic.modeler.get_3d_component_object_list(name=PCB)
    on_board = [v for v in (_fmax(ic, o) for o in layers) if v is not None]
    board_max = max(on_board) if on_board else None     # not `board`: that is the board.json dict here
    if board_max is None:
        check("warn", "Board temperature could not be read. Look at the result in AEDT.")
    board_mean = step("board average", lambda: _board_mean(ic, layers))
    # not "AllObjects": hollow network blocks in it break the AEDT session. Ask every other object one by one.
    _, tree = thermal(ic)
    hollow = {n[2:] for n, d in tree.items() if d["type"] == "Network" and n.startswith("P_")}
    rest = [o for o in ic.modeler.object_names if o not in hollow and o not in layers and o != "Region"
            and not (o.startswith(SOL) and not ic.modeler[o].is_model)]
    plates = [{"name": o, "maxC": _fmax(ic, o), "meanC": _fmax(ic, o, "Mean")} for o in rest if o.startswith(f"{SOL}Spreader_")]
    others = [v for v in (_fmax(ic, o) for o in rest) if v is not None]
    seen = [v for v in [board_max] + others + [r["tempC"] for r in rows] if v is not None]
    anywhere = max(seen) if seen else None
    res = {"schema": "pcb-thermal-ansys/results", "solvedAt": time.strftime("%H:%M:%S"), "seconds": round(secs),
           "ambientC": ambient, "converged": conv, "balance": bal, "monitors": mon, "profile": prof,
           "boardC": board_max, "boardMeanC": board_mean, "hottestC": anywhere, "parts": rows, "plates": plates}
    if write:
        (folder / "results.json").write_text(json.dumps(res, indent=1), encoding="utf-8")
    rid = step("keep the run", lambda: _keep_run(folder, res, study, parts, powered, setup, label))
    check("ok", f"Board hottest {board_max} C, board average {board_mean} C, hottest measured {anywhere} C"
                + ("; results.json written" if write else "") + (f"; kept as runs\\{rid}.json" if rid else ""))
    return res


# ---------------------------------------------------------------- mesh check

FINER = 2 / 3     # a finer mesh: every largest element times this, or one automatic level up


def _mesh_text(m):
    if m.get("mode") != "manual":
        return f"automatic level {int(m.get('level', 3))}"
    return f"manual, largest element {' x '.join(f'{v:g}' for v in _max_xyz(m))} mm"


def _finer(study, parts):
    """the study and parts with a finer global mesh and finer local meshes; None at automatic level 5, the finest"""
    m = dict(study["mesh"])
    if m.get("mode") == "manual":
        m["maxMm"], m["maxYMm"], m["maxZMm"] = (round(v * FINER, 3) for v in _max_xyz(m))
    elif int(m.get("level", 3)) >= 5:
        return None, None
    else:
        m["level"] = int(m.get("level", 3)) + 1
    fine_parts = [{**p, "meshMm": round(p["meshMm"] * FINER, 3)} if p.get("meshMm") else p for p in parts]
    return {**study, "mesh": m}, fine_parts


def _mesh_check(ic, folder, board, parts, powered, study):
    """the finer mesh first, the page's mesh last: AEDT and results.json end with the page's mesh and its solution"""
    fine, fine_parts = _finer(study, parts)
    if not fine:
        check("error", "The mesh is already at automatic level 5, the finest. Switch to Manual to check a finer one.")
        return
    check("ok", f"Mesh check 1 of 2: the finer mesh, {_mesh_text(fine['mesh'])}")
    _mesh(ic, fine_parts, fine)
    local_b = local_mesh(ic)       # as AEDT holds them: a junction model gets none, whatever the page has
    b = _solve(ic, folder, board, powered, fine, fine_parts, write=False, label="Mesh check, finer")
    check("ok", f"Mesh check 2 of 2: your mesh, {_mesh_text(study['mesh'])}")
    _mesh(ic, parts, study)
    local_a = local_mesh(ic)
    a = _solve(ic, folder, board, powered, study, parts, label="Mesh check, yours") if b else None
    if not (a and b):
        check("error", "Mesh check stopped: a solve failed")
        return
    (folder / "meshcheck.json").write_text(json.dumps({
        "schema": "pcb-thermal-ansys/meshcheck", "at": time.strftime("%Y-%m-%d %H:%M"), "ambientC": a["ambientC"],
        "yours": {"mesh": study["mesh"], "local": local_a, "text": _mesh_text(study["mesh"]), "result": a},
        "finer": {"mesh": fine["mesh"], "local": local_b, "text": _mesh_text(fine["mesh"]), "result": b},
    }, indent=1), encoding="utf-8")
    check("ok", "Mesh check done: meshcheck.json written")


def _board_mean(ic, layers):
    """volume average temperature of the PCB, from each layer's average weighted by its volume. Tells how much of the
    board is near the hot spot; heat leaves through the whole surface, so this is the number to compare with a
    hand estimate of convection and radiation."""
    got = [(m, ic.modeler[o].volume) for o in layers for m in [_fmax(ic, o, "Mean")] if m is not None]
    total = sum(v for _, v in got)
    return round(sum(m * v for m, v in got) / total, 2) if total else None


# ---------------------------------------------------------------- read back

def _get(settings, key, default=None):
    try:
        return settings[key]
    except BaseException:  # noqa: BLE001  key absent in this mode
        return default


def _gravity(ds):
    """the gravity direction of the design settings: "-Z" and so on, as study["orientation"]"""
    return f"{'-' if str(ds['GravityDir']) == 'Negative' else '+'}{str(ds['GravityVec'])[-1]}"


def air_pad(ic, board, tree=None):
    """air box bounding box and its room around the board and part blocks in mm, [+X -X +Y -Y +Z -Z]"""
    names = set(ic.modeler.object_names)
    if "Region" not in names:
        return None, None
    if tree is None:
        _, tree = thermal(ic)
    ext = []
    try:
        ext.append(ic.modeler.user_defined_components[PCB].bounding_box)
    except BaseException:  # noqa: BLE001
        pass
    ext += [ic.modeler[n[2:]].bounding_box for n, d in tree.items()
            if n.startswith("P_") and n[2:] in names and d["type"] in ("Solid Block", "Network")]
    ext += [ic.modeler[o].bounding_box for o in _passive(ic, tree, board)]
    reg = ic.modeler["Region"].bounding_box
    if not ext:
        return reg, None
    lo = [min(e[i] for e in ext) for i in range(3)]
    hi = [max(e[i + 3] for e in ext) for i in range(3)]
    return reg, [round(reg[3] - hi[0], 1), round(lo[0] - reg[0], 1), round(reg[4] - hi[1], 1),
                 round(lo[1] - reg[1], 1), round(reg[5] - hi[2], 1), round(lo[2] - reg[2], 1)]


def read_state(ic, board):
    ds = ic.design_settings
    names = set(ic.modeler.object_names)
    setup = next((x for x in ic.setups if x.name == SETUP), None)
    _, tree = thermal(ic)
    blocks, other, geo = [], [], []
    for n, d in tree.items():
        if n.startswith("P_") and n[2:] in names and d["type"] in ("Solid Block", "Network"):
            o = ic.modeler[n[2:]]
            bb = o.bounding_box
            geo.append([n[2:], [round(v, 3) for v in bb]])
            rjc = d["rjc"]
            if d["type"] == "Solid Block" and o.material_name.lower().startswith(RJC_MAT):
                rjc = _block_rjc(ic, o, bb[5] - bb[2])
            blocks.append({"refDes": n[2:], "kind": d["type"], "powerW": d["power"] or 0,
                           "heightMm": round(bb[5] - bb[2], 4), "rjcCW": rjc, "rjbCW": d["rjb"]})
        elif d["type"] != "Opening":
            other.append({"name": n, "type": d["type"], "powerW": d["power"]})
    for o in _passive(ic, tree, board):
        bb = ic.modeler[o].bounding_box
        geo.append([o, [round(v, 3) for v in bb]])
        blocks.append({"refDes": o, "kind": "Passive", "powerW": 0, "heightMm": round(bb[5] - bb[2], 4), "rjcCW": None, "rjbCW": None})
    sol = sol_objects(ic)
    geo += [[o["name"], o["box"], o["model"]] for o in sol]
    reg, pad = air_pad(ic, board, tree)
    keys = ic.materials.material_keys
    mats = [{"name": b["name"], **{f: num(getattr(m, attr).value) for f, attr, _ in MAT_F}}
            for b in board.get("materials", []) for m in [keys.get(b["name"].lower())] if m]
    gm = ic.mesh.global_mesh_region
    local = [{"refDes": r, "maxMm": v} for r, v in local_mesh(ic).items()]
    mesh = {"manual": bool(gm.manual_settings),
            "level": _get(gm.settings, "MeshRegionResolution"),
            **_manual_read(gm),
            "local": local}
    return {
        "schema": "pcb-thermal-ansys/state",
        "open": True,
        "readAt": time.strftime("%H:%M:%S"),
        "project": ic.project_name, "path": str(ic.project_file),
        "ambientC": num(ds["AmbTemp"]),
        "radiationC": num(_get(ds, "AmbRadTemp")),
        "gravity": _gravity(ds),
        "airMm": [round(reg[3] - reg[0], 1), round(reg[4] - reg[1], 1), round(reg[5] - reg[2], 1)] if reg else None,
        "padMm": pad,
        "setup": {"name": setup.name, "iterations": setup.props.get("Convergence Criteria - Max Iterations"),
                  "flowLimit": num(setup.props.get("Convergence Criteria - Flow")),
                  "energyLimit": num(setup.props.get("Convergence Criteria - Energy")),
                  "flow": setup.props.get("Flow Regime"),
                  "radiation": setup.props.get("Radiation Model")} if setup else None,
        "blocks": blocks,
        "materials": mats,
        "other": other,
        "mesh": mesh,
        "meshSig": _mesh_sig(sorted(geo), {**mesh, "level": str(mesh["level"])}, pad),
        "objects": len(names),
        "openings": [n for n, d in tree.items() if d["type"] == "Opening"],
        "solution": sol,
    }


def _open(Icepak, project_file):
    """the folder's project in AEDT. AEDT finds an open project by name: a copy with the same name from another
    folder would be the one changed (seen 2026-11-03), so that stops here"""
    ic = Icepak(project=str(project_file), design=DESIGN, version=VERSION, new_desktop=False, port=PORT)
    if Path(ic.project_file).resolve() != Path(project_file).resolve():
        raise RuntimeError(f"close the other {project_file.name} in AEDT and try again; it is open from "
                           f"{Path(ic.project_file).parent}")
    return ic


def read_only(folder, board, project_file):
    """mode read: look at AEDT as it is. Never starts AEDT, never opens a project, never writes checks.json."""
    from ansys.aedt.core import Desktop, Icepak
    out = folder / "state.json"
    old = json.loads(out.read_text(encoding="utf-8")) if out.exists() else {}
    name = f"{board['name']}_icepak"

    def closed(why):
        out.write_text(json.dumps({**old, "open": False, "why": why, "checkedAt": time.strftime("%H:%M:%S")}, indent=1),
                       encoding="utf-8")
        print(why)
        return 0

    running = subprocess.run(["tasklist", "/FI", "IMAGENAME eq ansysedt.exe"], capture_output=True, text=True).stdout
    if "ansysedt.exe" not in running.lower():
        return closed("AEDT is not running")
    pid = _busy_pid()
    if pid:   # a second gRPC client during a build made AEDT fail and once crash: keep the last state
        print(f"setup_study is running (process {pid}), AEDT not read")
        return 0
    from ansys.aedt.core.generic.general_methods import is_grpc_session_active
    if not is_grpc_session_active(PORT, version=VERSION):    # Desktop() would start a new AEDT on the port
        return closed(f"AEDT {VERSION} runs without the page's port {PORT} (started outside start.bat?), so it cannot be read")
    desktop = Desktop(version=VERSION, new_desktop=False, port=PORT, non_graphical=False)
    try:
        if name not in list(desktop.odesktop.GetProjectList()):
            return closed(f"Project {name} is not open in AEDT")
        try:
            ic = _open(Icepak, project_file)
        except RuntimeError as ex:
            return closed(str(ex))
        state = read_state(ic, board)
        out.write_text(json.dumps(state, indent=1), encoding="utf-8")
        print("read", state["readAt"])
        return 0
    finally:
        desktop.release_desktop(close_projects=False, close_on_exit=False)


def _busy_pid():
    """process id of a running build/update/mesh/solve (or the SO-8 benchmark), else None. A file left by a killed run
    is ignored."""
    try:
        pid = int(BUSY.read_text())
    except (OSError, ValueError):
        return None
    tasks = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True).stdout
    return pid if str(pid) in tasks else None


def take_busy():
    """mark AEDT as used by this process; returns the process id of another living user instead (another job, or the
    SO-8 benchmark, which takes it too), and then marks nothing"""
    pid = _busy_pid()
    if pid and pid != os.getpid():
        return pid
    BUSY.write_text(str(os.getpid()))
    return None


def drop_busy():
    """unmark, only when this process marked it: never another user's mark"""
    try:
        if int(BUSY.read_text()) == os.getpid():
            BUSY.unlink()
    except (OSError, ValueError):
        pass


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        raise SystemExit(__doc__)
    folder, mode = Path(sys.argv[1]).resolve(), sys.argv[2] if len(sys.argv) == 3 else "build"
    if mode == "read":
        sys.exit(main(folder, mode))
    other = take_busy()
    if other:       # two gRPC clients at once made AEDT hang, and once crash
        check("error", f"Another script (process {other}) is using AEDT, as the SO-8 benchmark does. Try again when it has ended.")
        sys.exit(_finish(folder, 0))
    try:
        code = main(folder, mode)
    finally:
        drop_busy()
    sys.exit(code)
