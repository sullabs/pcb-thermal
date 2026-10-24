"""Build or edit the Icepak study from <folder>/study.json. No solve.

    python ansys/setup_study.py <folder> [build|update|read|mesh]

build  : new AEDT project from the EDB: PCB, power blocks, air, environment, setup, mesh (about 40 s)
update : the same settings applied to the project that is already built (fast); AEDT edits to the things the
         page owns (power, ambient, air, mesh) are overwritten, everything else is left alone
read   : only read the project back into state.json (the page shows it)
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
    s = json.loads((folder / "study.json").read_text(encoding="utf-8"))
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
            _apply(ic, by_ref, powered, study)
            step("save project", lambda: ic.save_project())
            check("ok", f"Saved {project_file.name}")
        if mode == "mesh":
            _mesh_now(ic)
        state = step("read back", lambda: read_state(ic))
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

def _apply(ic, by_ref, powered, study):
    step("heat sources", lambda: _blocks(ic, by_ref, powered, study))
    step("air", lambda: _air(ic, study))
    step("environment", lambda: _environment(ic, study))
    step("setup", lambda: _setup(ic, study))
    step("mesh", lambda: _mesh(ic, powered, study))


def thermal(ic):
    """{boundary name: (type, power string)} from the model tree: about 0.1 s each, ic.boundaries takes 5 to 10 s."""
    node = ic.get_oo_object(ic.odesign, "Thermal")
    out = {}
    for n in ic.get_oo_name(ic.odesign, "Thermal"):
        out[n] = (ic.get_oo_property_value(node, n, "Type"), ic.get_oo_property_value(node, n, "Total Power"))
    return node, out


def _delete_boundary(ic, name):
    ic.odesign.GetModule("BoundarySetup").DeleteBoundaries([name])


def _delete_block(ic, ref):
    _delete_boundary(ic, f"P_{ref}")
    if ref in ic.modeler.object_names:
        ic.modeler.delete(ref)


def _blocks(ic, by_ref, powered, study):
    m = ic.materials.add_material(PKG)
    if m:
        m.thermal_conductivity = study["packageK"]
        m.mass_density = 2000
    want = {p["refDes"] for p in powered}
    node, bmap = thermal(ic)
    objects = set(ic.modeler.object_names)
    for name in list(bmap):
        if name.startswith("P_") and name[2:] not in want:
            _delete_block(ic, name[2:])
            bmap.pop(name)
            check("ok", f"{name[2:]}: heat source removed")
    made = 0
    for p in powered:
        ref, b = p["refDes"], by_ref[p["refDes"]]
        h = p.get("heightMm") or b["hMm"]
        if h < MIN_BLOCK_H_MM:
            check("warn", f"{ref}: EDB height {b['hMm']} mm is a placeholder, using {MIN_BLOCK_H_MM} mm. "
                          f"Type the real height.")
            h = MIN_BLOCK_H_MM
        z = b["zMinMm"] if b["side"] == "top" else b["zMinMm"] + b["hMm"] - h
        box = [b["xMinMm"], b["yMinMm"], z, b["xMinMm"] + b["lMm"], b["yMinMm"] + b["wMm"], z + h]
        bound = bmap.get(f"P_{ref}")
        if bound and ref in objects:
            cur = ic.modeler[ref].bounding_box
            if max(abs(a - c) for a, c in zip(cur, box)) > 1e-3:
                _delete_block(ic, ref)
                bound = None
        if bound:
            if num(bound[1]) != p["powerW"]:
                ic.get_oo_object(node, f"P_{ref}").SetPropValue("Total Power", f"{p['powerW']}W")
            made += 1
            continue
        obj = ic.modeler.create_box(box[:3], [box[3] - box[0], box[4] - box[1], box[5] - box[2]],
                                    name=ref, material=PKG if m else "Ceramic_material")
        if obj and ic.assign_solid_block(obj.name, f"{p['powerW']}W", boundary_name=f"P_{ref}"):
            made += 1
        else:
            check("error", f"{ref}: block or power not created")
    if powered:
        check("ok" if made == len(powered) else "error", f"{made} of {len(powered)} heat sources in AEDT")


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


def read_state(ic):
    ds = ic.design_settings
    reg = ic.modeler["Region"].bounding_box if "Region" in ic.modeler.object_names else None
    setup = next((x for x in ic.setups if x.name == SETUP), None)
    _, tree = thermal(ic)
    blocks = []
    for n, (typ, power) in tree.items():
        if n.startswith("P_") and n[2:] in ic.modeler.object_names:
            bb = ic.modeler[n[2:]].bounding_box
            blocks.append({"refDes": n[2:], "powerW": num(power, 0), "heightMm": round(bb[5] - bb[2], 4)})
    gm = ic.mesh.global_mesh_region
    local = [{"refDes": r, "maxMm": v} for r, v in local_mesh(ic).items()]
    return {
        "schema": "pcb-thermal-ansys/state",
        "readAt": time.strftime("%H:%M:%S"),
        "project": ic.project_name, "path": str(ic.project_file),
        "ambientC": num(ds["AmbTemp"]),
        "gravity": f"{'-' if str(ds['GravityDir']) == 'Negative' else '+'}{str(ds['GravityVec'])[-1]}",
        "airMm": [round(reg[3] - reg[0], 1), round(reg[4] - reg[1], 1), round(reg[5] - reg[2], 1)] if reg else None,
        "setup": {"name": setup.name, "iterations": setup.props.get("Convergence Criteria - Max Iterations"),
                  "flow": setup.props.get("Flow Regime"),
                  "radiation": setup.props.get("Radiation Model")} if setup else None,
        "blocks": blocks,
        "mesh": {"manual": bool(gm.manual_settings),
                 "level": _get(gm.settings, "MeshRegionResolution"),
                 "maxMm": num(_get(gm.settings, "MaxElementSizeX")),
                 "gapElems": num(_get(gm.settings, "MinElementsInGap")),
                 "edgeElems": num(_get(gm.settings, "MinElementsOnEdge")),
                 "local": local},
        "objects": len(ic.modeler.object_names),
        "openings": [n for n, (typ, _) in tree.items() if typ == "Opening"],
    }


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        raise SystemExit(__doc__)
    sys.exit(main(Path(sys.argv[1]).resolve(), sys.argv[2] if len(sys.argv) == 3 else "build"))
