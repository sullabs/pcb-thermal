"""Build or edit the Icepak study from <folder>/study.json. No solve.

    python ansys/setup_study.py <folder> [build|update|read|mesh|solve]

build  : new AEDT project from the EDB: PCB, power blocks, air, environment, setup, mesh (about 40 s)
update : the same settings applied to the project that is already built (fast); AEDT edits to the things the
         page owns (power, ambient, air, mesh) are overwritten, everything else is left alone
read   : only read the project into state.json, as it is in AEDT now (never starts AEDT or opens a project)
mesh   : generate the mesh and report how long it took
solve  : solve the setup (opens the project, never changes it) and write results.json: temperature per heat source

Input  : <folder>/<board>.aedb (or Altium's <board>.edb), <folder>/study.json (written by the page)
Output : <folder>/board.json, state.json, checks.json, <board>_icepak.aedt
"""
import json
import os
import re
import subprocess
import sys
import time
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).parent))
import read_board  # noqa: E402

VERSION = "2026.1"
PORT = 50051
MIN_BLOCK_H_MM = 0.2   # EDB heights below this are placeholders; the page value or this minimum is used
DESIGN = "Icepak"
SETUP = "Study"
PKG = "pkg_material"
BUSY = ".setup_running"   # in <folder> while a build/update/mesh/solve runs; mode read leaves AEDT alone then

# gravity points down. edit_design_settings code: 0 -X, 1 -Y, 2 -Z, 3 +X, 4 +Y, 5 +Z (read back from AEDT)
# fields: gravity code, label, text, index of the air face opposite to gravity in [+X, -X, +Y, -Y, +Z, -Z]
ORIENTATION = {
    "flat-up": (2, "-Z", "parts face up", 4),
    "flat-down": (5, "+Z", "parts face down", 5),
    "vertical-long": (1, "-Y", "board standing, long edge down", 2),
    "vertical-short": (0, "-X", "board standing, short edge down", 0),
}

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
    s.setdefault("orientation", "flat-up")
    s.setdefault("marginMm", 50)
    s.setdefault("marginAboveMm", 50)
    s.setdefault("radiation", True)
    s.setdefault("pcbResolution", 2)
    s.setdefault("iterations", 300)
    s.setdefault("packageK", 1.0)
    s.setdefault("mesh", {"mode": "auto", "level": 3})
    s.setdefault("parts", [])
    return s


def num(text, default=None):
    """'1.5W' or '30cel' or '2mm' -> 1.5, 30, 2"""
    m = re.match(r"\s*(-?[\d.]+(?:e-?\d+)?)", str(text))
    return float(m.group(1)) if m else default


def connect():
    from ansys.aedt.core import Desktop
    try:
        return Desktop(version=VERSION, new_desktop=False, port=PORT, non_graphical=False)
    except BaseException:  # noqa: BLE001
        return Desktop(version=VERSION, new_desktop=True, port=PORT, non_graphical=False)


def main(folder: Path, mode: str = "build") -> int:
    from ansys.aedt.core import Hfss3dLayout, Icepak
    from ansys.aedt.core.generic.settings import settings
    settings.enable_screen_logs = False

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
    powered = [p for p in study["parts"] if (p.get("powerW") or 0) > 0]
    for p in powered:
        if p["refDes"] not in by_ref:
            check("error", f"{p['refDes']} has power but is not in the EDB")
    powered = [p for p in powered if p["refDes"] in by_ref]
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
            ic = step("open project", lambda: Icepak(project=str(project_file), design=DESIGN, version=VERSION,
                                                    new_desktop=False, port=PORT), fatal=True)
        if mode in ("build", "update"):
            _apply(ic, board, by_ref, powered, study)
            step("save project", lambda: ic.save_project())
            check("ok", f"Saved {project_file.name}")
        if mode == "mesh":
            _mesh_now(ic)
        if mode == "solve":
            step("solve", lambda: _solve(ic, folder, board, powered, study["ambientC"]), fatal=True)
        state = None if mode == "solve" else step("read back", lambda: read_state(ic, board))   # right after a solve AEDT cannot be read
        if state:
            (folder / "state.json").write_text(json.dumps(state, indent=1), encoding="utf-8")
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


THERMAL = {"copper": (400, 8933, 385), "solder": (48, 8000, 167), "FR-4": (0.3, 1900, 1100)}   # W/mK, kg/m3, J/kgK


def _thermal_kind(name, m):
    n = name.lower()
    if "solder" in n and "resist" not in n and "mask" not in n:
        return "solder"
    if "copper" in n or n.startswith("cu") or not m.is_dielectric():
        return "copper"
    return "FR-4"


def _thermal_defaults(app):
    """EDB materials from an electrical export (Altium, Megtron, Rogers...) can lack density, heat capacity and
    conductivity: the solver then refuses to start or, for copper, solves with no conduction. Give them copper, solder
    or FR-4 values and say so. Run on the layout design before the PCB import too: the PCB component copies the
    layout's materials again at every solve, and a changed material marks the solution invalid."""
    fixed = {}
    for name, m in app.materials.material_keys.items():
        try:
            if float(m.mass_density.value) > 0 and float(m.thermal_conductivity.value) > 0.01:
                continue
            kind = _thermal_kind(name, m)
            k, rho, cp = THERMAL[kind]
            if float(m.thermal_conductivity.value) <= 0.01:
                m.thermal_conductivity = k
            m.mass_density = rho
            m.specific_heat = cp
            fixed.setdefault(kind, []).append(name)
        except BaseException:  # noqa: BLE001
            continue
    if fixed:
        used = "; ".join(f"{kind} values ({THERMAL[kind][0]} W/mK) for {', '.join(names)}" for kind, names in fixed.items())
        check("warn", f"No thermal data in the EDB: {used}. Type the real values in AEDT if they matter.")


def _build(Hfss3dLayout, Icepak, desktop, folder, board, study):
    _close_same_name(desktop, [board["name"], f"{board['name']}_icepak"])
    lay = Hfss3dLayout(version=VERSION, new_desktop=False, port=PORT)
    lay.import_edb(str(folder / board["edb"]))
    pn, dn = lay.project_name, lay.design_name
    _thermal_defaults(lay)
    ic = Icepak(project=pn, design=DESIGN, version=VERSION, new_desktop=False, port=PORT)
    ok = ic.create_pcb_from_3dlayout("PCB", pn, dn, resolution=int(study["pcbResolution"]), extent_type="Bounding Box",
                                     close_linked_project_after_import=False,
                                     rad="Both" if study["radiation"] else "Nothing")
    check("ok" if ok else "error", f"PCB from the EDB, mapping resolution {study['pcbResolution']}"
          if ok else "PCB import returned False")
    _thermal_defaults(ic)
    names = [re.sub(r"^PCB_\d+_|_BBox$", "", o) for o in ic.modeler.get_3d_component_object_list(name=PCB)]
    dup = sorted({n for n in names if names.count(n) > 1})
    if dup:
        check("error", f"The stackup has more than one zone (rigid-flex or stepped board): layers {', '.join(dup)} are "
                       f"imported twice, so the solver will refuse to start. Icepak's PCB import cannot take this board. "
                       f"Export one zone only, or model it in AEDT.")
    ic.save_project(str(folder / f"{board['name']}_icepak.aedt"))
    return ic


# ---------------------------------------------------------------- apply (same for build and update)

def _apply(ic, board, by_ref, powered, study):
    step("heat sources", lambda: _blocks(ic, board, by_ref, powered, study))
    step("air", lambda: _air(ic, study))
    step("environment", lambda: _environment(ic, study))
    step("setup", lambda: _setup(ic, study))
    step("mesh", lambda: _mesh(ic, powered, study))


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
    if rjc > 0 or rjb > 0:
        check("warn", f"{p['refDes']}: only one of Rjc and Rjb is given. A junction model needs both, so a plain block "
                      f"is used and no junction temperature is computed.")
    return "Solid Block"


def _same(a, b):
    return a is not None and b is not None and abs(a - b) < 1e-9


def _blocks(ic, board, by_ref, powered, study):
    m = ic.materials.add_material(PKG)
    if m:
        m.thermal_conductivity = study["packageK"]
        m.mass_density = 2000
    want = {p["refDes"] for p in powered}
    node, tree = thermal(ic)
    objects = set(ic.modeler.object_names)
    for name in list(tree):
        if name.startswith("P_") and name[2:] not in want:
            _delete_block(ic, name[2:])
            tree.pop(name)
            check("ok", f"{name[2:]}: heat source removed")
    zlo, zhi = board_z(ic, board)
    made = nets = 0
    for p in powered:
        ref, b = p["refDes"], by_ref[p["refDes"]]
        h = p.get("heightMm") or b["hMm"]     # a typed height is used as it is, only the EDB value gets the minimum
        if not p.get("heightMm") and h < MIN_BLOCK_H_MM:
            check("warn", f"{ref}: EDB height {b['hMm']} mm is a placeholder, using {MIN_BLOCK_H_MM} mm. "
                          f"Type the real height.")
            h = MIN_BLOCK_H_MM
        z = zhi if b["side"] == "top" else zlo - h      # the part sits on the board surface, so a network can touch it
        box = [b["xMinMm"], b["yMinMm"], z, b["xMinMm"] + b["lMm"], b["yMinMm"] + b["wMm"], z + h]
        kind = _kind(p)
        cur = tree.get(f"P_{ref}")
        ok = cur is not None and ref in objects and cur["type"] == kind
        if ok:
            ok = max(abs(a - c) for a, c in zip(ic.modeler[ref].bounding_box, box)) <= 1e-3
        if ok and kind == "Network":
            ok = _same(cur["rjc"], p["rjcCW"]) and _same(cur["rjb"], p["rjbCW"])
        if ok:
            if not _same(cur["power"], p["powerW"]):
                prop = "Total Power" if kind == "Solid Block" else "Internal: Properties/Power"
                ic.get_oo_object(node, f"P_{ref}").SetPropValue(prop, f"{p['powerW']}W")
            made += 1
            nets += kind == "Network"
            continue
        if cur is not None or ref in objects:
            _delete_block(ic, ref)
        obj = ic.modeler.create_box(box[:3], [box[3] - box[0], box[4] - box[1], box[5] - box[2]],
                                    name=ref, material=PKG if m else "Ceramic_material")
        if not obj:
            check("error", f"{ref}: block not created")
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
    if powered:
        check("ok" if made == len(powered) else "error",
              f"{made} of {len(powered)} heat sources in AEDT, {nets} with a junction model (Rjc, Rjb)")


def _air(ic, study):
    side, above = study["marginMm"], study["marginAboveMm"]
    pad = [side] * 6                           # +X -X +Y -Y +Z -Z
    pad[ORIENTATION[study["orientation"]][3]] += above   # extra room on the face opposite to gravity
    bb, have = air_pad(ic)
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
    ic.edit_design_settings(gravity_dir=gdir, ambient_temperature=study["ambientC"])
    check("ok", f"Ambient {study['ambientC']} C, gravity {label} ({what})")


def _setup(ic, study):
    s = next((x for x in ic.setups if x.name == SETUP), None)
    existed = bool(s)
    s = s or ic.create_setup(SETUP)
    if not s:
        check("error", "Setup not created")
        return
    for k, v in (("Flow Regime", "Laminar"), ("Include Flow", True), ("Include Gravity", True),
                 ("Convergence Criteria - Max Iterations", int(study["iterations"]))):
        if k in s.props:
            s.props[k] = v
    was = str(s.props.get("Radiation Model", "Off"))
    s.props["Radiation Model"] = "Discrete Ordinates Model" if study["radiation"] else "Off"
    s.update()
    check("ok", f"Setup '{SETUP}': laminar natural convection, radiation "
                f"{'on (discrete ordinates)' if study['radiation'] else 'off'}, {study['iterations']} iterations max")
    if existed and was != s.props["Radiation Model"]:
        # the next solve would restart from the old case file, which keeps the old radiation model
        if step("clear old solution", lambda: ic.cleanup_solution()):
            check("ok","Radiation changed: old solution deleted, the next solve starts fresh")


def _mesh(ic, powered, study):
    m = study["mesh"]
    gm = ic.mesh.global_mesh_region
    manual = m.get("mode") == "manual"
    want = ({"MaxElementSizeX": m["maxMm"], "MinElementsInGap": m.get("gapElems", 3), "MinElementsOnEdge": m.get("edgeElems", 2)}
            if manual else {"MeshRegionResolution": int(m.get("level", 3))})
    same = bool(gm.manual_settings) == manual and all(num(_get(gm.settings, k)) == float(v) for k, v in want.items())
    if not same:
        gm.manual_settings = manual
        if manual:
            for ax in "XYZ":
                gm.settings[f"MaxElementSize{ax}"] = f"{m['maxMm']}mm"
            gm.settings["MinElementsInGap"] = str(m.get("gapElems", 3))
            gm.settings["MinElementsOnEdge"] = str(m.get("edgeElems", 2))
        else:
            gm.settings["MeshRegionResolution"] = int(m.get("level", 3))
        gm.update()
    wanted = {p["refDes"]: p["meshMm"] for p in powered if p.get("meshMm")}
    have = local_mesh(ic)
    stale = [ref for ref, size in have.items() if wanted.get(ref) != size]
    if stale:
        for r in list(ic.mesh.meshregions):
            if r.name[5:] in stale and r.name.startswith("Mesh_") and not r.name.endswith("_SubRegion"):
                try:
                    r.delete()
                except BaseException:  # noqa: BLE001  AEDT removes the sub-region together with the region
                    pass
        for ref in stale:
            have.pop(ref)
    for o in [o for o in ic.modeler.object_names if o.startswith("Mesh_") and o.endswith("_SubRegion")
              and o[5:-10] not in wanted]:
        ic.modeler.delete(o)
    ic.modeler.refresh_all_ids()
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
    what = f"global max element {m['maxMm']} mm" if manual else f"global automatic level {m.get('level', 3)}"
    check("ok", f"Mesh: {what}, {len(wanted)} local refinements")


def local_mesh(ic):
    """{RefDes: max element mm} of the local mesh regions, from the model tree"""
    node = ic.get_oo_object(ic.odesign, "Mesh")
    return {n[5:]: num(ic.get_oo_property_value(node, n, "Maximum Element Size/X"))
            for n in ic.get_oo_name(ic.odesign, "Mesh") if n.startswith("Mesh_") and not n.endswith("_SubRegion")}


def _mesh_now(ic):
    t = time.time()
    ok = ic.mesh.generate_mesh(SETUP)
    check("ok" if ok else "error", f"Mesh generated in {time.time() - t:.0f} s" if ok else "Mesh generation failed")


# ---------------------------------------------------------------- solve

def _hollow(ic):
    """network blocks must not be meshed inside, also in projects built before this was set"""
    _, tree = thermal(ic)
    for n, d in tree.items():
        if d["type"] == "Network" and n.startswith("P_") and ic.modeler[n[2:]].solve_inside:
            ic.modeler[n[2:]].solve_inside = False


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
    monitor sits on the case side face (the face away from the board), a plain block gets it at its centre."""
    for n in [n for n in ic.get_oo_name(ic.odesign, "Monitor") if n.startswith("T_")]:
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


def _fmax(ic, obj):
    """highest temperature in one object (field calculator). The solution name must be given, or AEDT refuses.
    Not post.get_scalar_field_value: it tries EnterQty("Temperature") first, which in Icepak always fails (Temperature
    is only a named expression) and puts "Abnormal script termination" in the message manager for every object.
    A hollow object (network block) has no volume and fails; the stack is cleared so the next call still works."""
    fr = ic.post.ofieldsreporter
    f = Path(ic.working_directory) / f"fmax_{obj}.fld"
    try:
        fr.CalcStack("clear")
        fr.CopyNamedExprToStack("Temperature")
        fr.EnterVol(obj)
        fr.CalcOp("Maximum")
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


def _temp(ic, ref, hollow):
    """temperature of a heat source: field maximum for a block, case-side monitor for a hollow network block.
    The field calculator must never see a hollow block: the failed call breaks the whole AEDT session."""
    v = None if hollow else _fmax(ic, ref)
    if v is not None:
        return v
    try:
        sd = ic.post.get_solution_data(expressions=[f"T_{ref}.Temperature"], setup_sweep_name=f"{SETUP} : SteadyState",
                                       report_category="Monitor")
        return round(float(sd.get_expression_data(sd.expressions[0])[1][-1]), 2)
    except BaseException:  # noqa: BLE001  one missing value must not stop the others
        return None


def _balance(ic, total_w):
    """heat out through the Open_air faces against the power put in, from AEDT's Fields Summary.
    Face monitors count only convection; the summary also has the radiation that leaves through the opening."""
    for n in [n for n in ic.get_oo_name(ic.odesign, "Monitor") if n.startswith("Q_")]:
        ic.monitor.delete_monitor(n)    # the heat flow face monitors older versions made
    fs = ic.post.create_field_summary()
    qs = {"outW": "HeatFlowRate", "radiationW": "RadiationFlow", "convectionW": "ConvectiveHeatFlow"}
    for q in qs.values():
        fs.add_calculation("Boundary", "Surface", "Open_air", q)
    data = fs.get_field_summary_data(setup=f"{SETUP} : SteadyState")
    got = {}
    for name, total in zip(data["Quantity"], data["Total"]):    # "HeatFlowRate[W]" -> "3.30061"
        q, unit = name.rstrip("]").split("[")
        got[q] = abs(float(total)) / (1000.0 if unit == "mW" else 1.0)
    bal = {k: round(got[q], 4) for k, q in qs.items()}
    return {"inW": round(total_w, 4), **bal, "errorPct": round(100 * (bal["outW"] - total_w) / total_w, 2)}


def _convergence(ic):
    """final residuals against the setup's convergence criteria: {"ok": bool, "items": [{name, value, limit, ok}]}"""
    setup = next((x for x in ic.setups if x.name == SETUP), None)
    limit = {"Continuity": "Flow", "XVelocity": "Flow", "YVelocity": "Flow", "ZVelocity": "Flow", "Energy": "Energy"}
    if setup and str(setup.props.get("Radiation Model", "Off")) != "Off":
        limit["DOIntensity"] = "Discrete Ordinates"
    items = []
    for q, crit in limit.items():
        try:
            sd = ic.post.get_solution_data(expressions=["Residual." + q], setup_sweep_name=f"{SETUP} : SteadyState",
                                           report_category="Residual")
            val = float(sd.get_expression_data(sd.expressions[0])[1][-1])
            lim = num(setup.props.get(f"Convergence Criteria - {crit}")) if setup else None
            items.append({"name": q, "value": val, "limit": lim, "ok": lim is None or val <= lim})
        except BaseException:  # noqa: BLE001
            items.append({"name": q, "value": None, "limit": None, "ok": None})
    known = [i for i in items if i["ok"] is not None]
    return {"ok": bool(known) and all(i["ok"] for i in known), "read": len(known), "items": items}


def _solve(ic, folder, board, powered, ambient):
    step("junction blocks hollow", lambda: _hollow(ic))
    powered = _sources(ic, powered)     # from here on the heat sources are the ones in AEDT, with AEDT's power
    step("temperature monitors", lambda: _monitors(ic, powered, board))
    t = time.time()
    ok = ic.analyze_setup(SETUP)
    secs = time.time() - t
    if not ok:
        check("error", "Solve failed or was stopped. Open AEDT and read the message window.")
        return
    check("ok", f"Solved in {secs:.0f} s")
    conv = _convergence(ic)
    if not conv["read"]:
        check("warn", "Convergence could not be read. The result files may be stale (AEDT was killed or crashed earlier). "
                      "Press Solve again; if it stays, check the residuals in AEDT before trusting the numbers.")
    elif conv["ok"]:
        check("ok", "Converged: all residuals are below the setup limits")
    else:
        bad = ", ".join(i["name"] for i in conv["items"] if i["ok"] is False)
        check("warn", f"NOT converged ({bad}). Raise the iterations or refine the mesh before trusting the numbers.")
    total_w = sum(p["powerW"] for p in powered)
    bal = step("energy balance", lambda: _balance(ic, total_w)) if total_w > 0 else None
    if bal:
        e = abs(bal["errorPct"])
        check("ok" if e <= 1 else "warn", f"Energy balance: {bal['inW']} W in, {bal['outW']} W out "
              f"({bal['convectionW']} W convection, {bal['radiationW']} W radiation, {bal['errorPct']:+.2f} %)"
              + ("" if e <= 1 else ". More than 1 % is not balanced: converge further or refine the mesh."))
    rows = [{"refDes": p["refDes"], "powerW": p["powerW"], "tempC": _temp(ic, p["refDes"], p["network"]),
             "rjaCW": p["rjaCW"], "tjMaxC": p["tjMaxC"]} for p in powered]
    if any(r["tempC"] is None for r in rows):
        check("warn", "Some temperatures could not be read. Look at the result in AEDT.")
    layers = ic.modeler.get_3d_component_object_list(name=PCB)
    on_board = [v for v in (_fmax(ic, o) for o in layers) if v is not None]
    board = max(on_board) if on_board else None
    if board is None:
        check("warn", "Board temperature could not be read. Look at the result in AEDT.")
    # not "AllObjects": hollow network blocks in it break the AEDT session. Ask every other object one by one.
    _, tree = thermal(ic)
    hollow = {n[2:] for n, d in tree.items() if d["type"] == "Network" and n.startswith("P_")}
    rest = [o for o in ic.modeler.object_names if o not in hollow and o not in layers and o != "Region"]
    others = [v for v in (_fmax(ic, o) for o in rest) if v is not None]
    seen = [v for v in [board] + others + [r["tempC"] for r in rows] if v is not None]
    anywhere = max(seen) if seen else None
    (folder / "results.json").write_text(json.dumps({
        "schema": "pcb-thermal-ansys/results", "solvedAt": time.strftime("%H:%M:%S"), "seconds": round(secs),
        "ambientC": ambient, "converged": conv, "balance": bal, "boardC": board, "hottestC": anywhere, "parts": rows}, indent=1), encoding="utf-8")
    step("save project", lambda: ic.save_project())
    check("ok", f"Board hottest {board} C, hottest measured {anywhere} C; results.json written")


# ---------------------------------------------------------------- read back

def _get(settings, key, default=None):
    try:
        return settings[key]
    except BaseException:  # noqa: BLE001  key absent in this mode
        return default


def air_pad(ic, tree=None):
    """air box bounding box and its room around the board and heat sources in mm, [+X -X +Y -Y +Z -Z]"""
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
    reg = ic.modeler["Region"].bounding_box
    if not ext:
        return reg, None
    lo = [min(e[i] for e in ext) for i in range(3)]
    hi = [max(e[i + 3] for e in ext) for i in range(3)]
    return reg, [round(reg[3] - hi[0], 1), round(lo[0] - reg[0], 1), round(reg[4] - hi[1], 1),
                 round(lo[1] - reg[1], 1), round(reg[5] - hi[2], 1), round(lo[2] - reg[2], 1)]


def read_state(ic, board=None):
    ds = ic.design_settings
    names = set(ic.modeler.object_names)
    setup = next((x for x in ic.setups if x.name == SETUP), None)
    _, tree = thermal(ic)
    blocks, other = [], []
    for n, d in tree.items():
        if n.startswith("P_") and n[2:] in names and d["type"] in ("Solid Block", "Network"):
            bb = ic.modeler[n[2:]].bounding_box
            blocks.append({"refDes": n[2:], "kind": d["type"], "powerW": d["power"] or 0,
                           "heightMm": round(bb[5] - bb[2], 4), "rjcCW": d["rjc"], "rjbCW": d["rjb"]})
        elif d["type"] != "Opening":
            other.append({"name": n, "type": d["type"], "powerW": d["power"]})
    reg, pad = air_pad(ic, tree)
    gm = ic.mesh.global_mesh_region
    local = [{"refDes": r, "maxMm": v} for r, v in local_mesh(ic).items()]
    return {
        "schema": "pcb-thermal-ansys/state",
        "open": True,
        "readAt": time.strftime("%H:%M:%S"),
        "project": ic.project_name, "path": str(ic.project_file),
        "ambientC": num(ds["AmbTemp"]),
        "gravity": f"{'-' if str(ds['GravityDir']) == 'Negative' else '+'}{str(ds['GravityVec'])[-1]}",
        "airMm": [round(reg[3] - reg[0], 1), round(reg[4] - reg[1], 1), round(reg[5] - reg[2], 1)] if reg else None,
        "padMm": pad,
        "setup": {"name": setup.name, "iterations": setup.props.get("Convergence Criteria - Max Iterations"),
                  "flow": setup.props.get("Flow Regime"),
                  "radiation": setup.props.get("Radiation Model")} if setup else None,
        "blocks": blocks,
        "other": other,
        "mesh": {"manual": bool(gm.manual_settings),
                 "level": _get(gm.settings, "MeshRegionResolution"),
                 "maxMm": num(_get(gm.settings, "MaxElementSizeX")),
                 "gapElems": num(_get(gm.settings, "MinElementsInGap")),
                 "edgeElems": num(_get(gm.settings, "MinElementsOnEdge")),
                 "local": local},
        "objects": len(names),
        "openings": [n for n, d in tree.items() if d["type"] == "Opening"],
    }


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
    pid = _busy_pid(folder)
    if pid:   # a second gRPC client during a build made AEDT fail and once crash: keep the last state
        print(f"setup_study is running (process {pid}), AEDT not read")
        return 0
    desktop = Desktop(version=VERSION, new_desktop=False, port=PORT, non_graphical=False)
    try:
        if name not in list(desktop.odesktop.GetProjectList()):
            return closed(f"Project {name} is not open in AEDT")
        ic = Icepak(project=str(project_file), design=DESIGN, version=VERSION, new_desktop=False, port=PORT)
        state = read_state(ic, board)
        out.write_text(json.dumps(state, indent=1), encoding="utf-8")
        print("read", state["readAt"])
        return 0
    finally:
        desktop.release_desktop(close_projects=False, close_on_exit=False)


def _busy_pid(folder):
    """process id of a running build/update/mesh/solve on this folder, else None. A file left by a killed run is ignored."""
    try:
        pid = int((folder / BUSY).read_text())
    except (OSError, ValueError):
        return None
    tasks = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True).stdout
    return pid if str(pid) in tasks else None


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        raise SystemExit(__doc__)
    folder, mode = Path(sys.argv[1]).resolve(), sys.argv[2] if len(sys.argv) == 3 else "build"
    if mode == "read":
        sys.exit(main(folder, mode))
    busy = folder / BUSY
    busy.write_text(str(os.getpid()))
    try:
        code = main(folder, mode)
    finally:
        busy.unlink(missing_ok=True)
    sys.exit(code)
