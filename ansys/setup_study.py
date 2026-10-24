"""Build or edit the Icepak study from <folder>/study.json. No solve.

    python ansys/setup_study.py <folder> [build|update|read|mesh]

build  : new AEDT project from the EDB: PCB, power blocks, air, environment, setup, mesh (about 40 s)
update : the same settings applied to the project that is already built (fast); AEDT edits to the things the
         page owns (power, ambient, air, mesh) are overwritten, everything else is left alone
read   : only read the project into state.json, as it is in AEDT now (never starts AEDT or opens a project)
mesh   : generate the mesh and report how long it took

Input  : <folder>/<board>.aedb, <folder>/study.json (written by the page)
Output : <folder>/board.json, state.json, checks.json, <board>_icepak.aedt
"""
import json
import re
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
        read_board.main(folder)
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
        state = step("read back", lambda: read_state(ic, board))
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


def _build(Hfss3dLayout, Icepak, desktop, folder, board, study):
    _close_same_name(desktop, [board["name"], f"{board['name']}_icepak"])
    lay = Hfss3dLayout(version=VERSION, new_desktop=False, port=PORT)
    lay.import_edb(str(folder / board["edb"]))
    pn, dn = lay.project_name, lay.design_name
    ic = Icepak(project=pn, design=DESIGN, version=VERSION, new_desktop=False, port=PORT)
    ok = ic.create_pcb_from_3dlayout("PCB", pn, dn, resolution=int(study["pcbResolution"]), extent_type="Bounding Box",
                                     close_linked_project_after_import=False,
                                     rad="Both" if study["radiation"] else "Nothing")
    check("ok" if ok else "error", f"PCB from the EDB, mapping resolution {study['pcbResolution']}"
          if ok else "PCB import returned False")
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
        h = p.get("heightMm") or b["hMm"]
        if h < MIN_BLOCK_H_MM:
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
    s = next((x for x in ic.setups if x.name == SETUP), None) or ic.create_setup(SETUP)
    if not s:
        check("error", "Setup not created")
        return
    for k, v in (("Flow Regime", "Laminar"), ("Include Flow", True), ("Include Gravity", True),
                 ("Convergence Criteria - Max Iterations", int(study["iterations"]))):
        if k in s.props:
            s.props[k] = v
    s.props["Radiation Model"] = "Discrete Ordinates Model" if study["radiation"] else "Off"
    s.update()
    check("ok", f"Setup '{SETUP}': laminar natural convection, radiation "
                f"{'on (discrete ordinates)' if study['radiation'] else 'off'}, {study['iterations']} iterations max")


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


# ---------------------------------------------------------------- read back

def _get(settings, key, default=None):
    try:
        return settings[key]
    except BaseException:  # noqa: BLE001  key absent in this mode
        return default


def read_state(ic, board=None):
    ds = ic.design_settings
    names = set(ic.modeler.object_names)
    setup = next((x for x in ic.setups if x.name == SETUP), None)
    _, tree = thermal(ic)
    blocks, other = [], []
    ext = []
    try:
        ext.append(ic.modeler.user_defined_components[PCB].bounding_box)
    except BaseException:  # noqa: BLE001
        pass
    for n, d in tree.items():
        if n.startswith("P_") and n[2:] in names and d["type"] in ("Solid Block", "Network"):
            bb = ic.modeler[n[2:]].bounding_box
            ext.append(bb)
            blocks.append({"refDes": n[2:], "kind": d["type"], "powerW": d["power"] or 0,
                           "heightMm": round(bb[5] - bb[2], 4), "rjcCW": d["rjc"], "rjbCW": d["rjb"]})
        elif d["type"] != "Opening":
            other.append({"name": n, "type": d["type"], "powerW": d["power"]})
    reg = ic.modeler["Region"].bounding_box if "Region" in names else None
    pad = None
    if reg and ext:
        lo = [min(e[i] for e in ext) for i in range(3)]
        hi = [max(e[i + 3] for e in ext) for i in range(3)]
        pad = [round(reg[3] - hi[0], 1), round(lo[0] - reg[0], 1), round(reg[4] - hi[1], 1),
               round(lo[1] - reg[1], 1), round(reg[5] - hi[2], 1), round(lo[2] - reg[2], 1)]   # +X -X +Y -Y +Z -Z
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
    import subprocess
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


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        raise SystemExit(__doc__)
    sys.exit(main(Path(sys.argv[1]).resolve(), sys.argv[2] if len(sys.argv) == 3 else "build"))
