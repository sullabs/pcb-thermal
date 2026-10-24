"""Build the Icepak study from <folder>/study.json. No solve.

    python ansys/setup_study.py <folder>

Input  : <folder>/<board>.aedb, <folder>/study.json (written by the page)
Output : <folder>/board.json, <folder>/checks.json, <folder>/<board>_icepak.aedt

Steps: read the EDB, connect to AEDT (attach, else start), import the EDB, make the Icepak design with the PCB,
one power block per part with power, air region with open faces, ambient and gravity, setup and mesh, checks.
Everything is rebuilt each run (about a minute); the old design is replaced.
"""
import json
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
    s.setdefault("regionPaddingMm", 50)
    s.setdefault("meshLevel", 2)
    s.setdefault("packageK", 1.0)
    s.setdefault("parts", [])
    return s


def connect():
    from ansys.aedt.core import Desktop
    try:
        return Desktop(version=VERSION, new_desktop=False, port=PORT, non_graphical=False)
    except BaseException:  # noqa: BLE001
        return Desktop(version=VERSION, new_desktop=True, port=PORT, non_graphical=False)


def main(folder: Path) -> int:
    from ansys.aedt.core import Hfss3dLayout, Icepak
    from ansys.aedt.core.generic.settings import settings
    settings.enable_screen_logs = False

    study = load_study(folder)
    read_board.main(folder)
    board = json.loads((folder / "board.json").read_text(encoding="utf-8"))
    by_ref = {p["refDes"]: p for p in board["parts"]}
    check("ok", f"Board {board['name']}: {len(board['parts'])} parts, {board['thicknessMm']} mm thick")

    powered = [p for p in study["parts"] if (p.get("powerW") or 0) > 0]
    for p in powered:
        if p["refDes"] not in by_ref:
            check("error", f"{p['refDes']} has power but is not in the EDB")
    powered = [p for p in powered if p["refDes"] in by_ref]
    if not powered:
        check("warn", "No part has power. The study will have no heat source.")
    total = sum(p["powerW"] for p in powered)
    check("ok", f"{len(powered)} parts with power, total {total:.3f} W")

    try:
        desktop = step("connect to AEDT", connect, fatal=True)
    except BaseException:  # noqa: BLE001
        return _finish(folder, total)
    try:
        lay = step("import EDB", lambda: _import(Hfss3dLayout, desktop, folder / board["edb"], board["name"]), fatal=True)
        pn, dn = lay.project_name, lay.design_name
        ic = step("Icepak design", lambda: Icepak(project=pn, design="Icepak", version=VERSION,
                                                 new_desktop=False, port=PORT), fatal=True)
        pcb_ok = step("PCB from EDB", lambda: ic.create_pcb_from_3dlayout(
            "PCB", pn, dn, resolution=2, extent_type="Bounding Box",
            close_linked_project_after_import=False, rad="Both"))
        check("ok" if pcb_ok else "error", "PCB imported from the EDB" if pcb_ok else "PCB import returned False")

        _blocks(ic, board, by_ref, powered, study)
        _air(ic, board, study)
        _setup(ic, study)
        _validate(ic)

        out = folder / f"{board['name']}_icepak.aedt"
        step("save project", lambda: ic.save_project(str(out)))
        check("ok", f"Saved {out.name}")
    except BaseException as ex:  # noqa: BLE001
        check("error", f"Stopped early: {type(ex).__name__}: {str(ex).splitlines()[0][:160] if str(ex) else ''}")
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


def _import(Hfss3dLayout, desktop, edb_path, name):
    # a project from an earlier run has the board's name; AEDT cannot hold two
    if name in desktop.project_list:
        desktop.odesktop.CloseProject(name)
    lay = Hfss3dLayout(version=VERSION, new_desktop=False, port=PORT)
    lay.import_edb(str(edb_path))
    return lay


def _blocks(ic, board, by_ref, powered, study):
    pkg = "pkg_material"
    m = step("package material", lambda: ic.materials.add_material(pkg))
    if m:
        m.thermal_conductivity = study["packageK"]
        m.mass_density = 2000
    top_z = board["thicknessMm"]
    made = 0
    for p in powered:
        b = by_ref[p["refDes"]]
        h = p.get("heightMm") or b["hMm"]
        if h < MIN_BLOCK_H_MM:
            check("warn", f"{p['refDes']}: EDB height {b['hMm']} mm is a placeholder, using {MIN_BLOCK_H_MM} mm. "
                          f"Type the real height on the page.")
            h = MIN_BLOCK_H_MM
        z = b["zMinMm"] if b["side"] == "top" else b["zMinMm"] + b["hMm"] - h
        obj = ic.modeler.create_box([b["xMinMm"], b["yMinMm"], z], [b["lMm"], b["wMm"], h],
                                    name=p["refDes"], material=pkg if m else "Ceramic_material")
        src = ic.assign_solid_block(obj.name, f"{p['powerW']}W", boundary_name=f"P_{p['refDes']}") if obj else None
        if obj and src:
            made += 1
        else:
            check("error", f"{p['refDes']}: block or power not created")
    if powered:
        check("ok" if made == len(powered) else "error", f"{made} of {len(powered)} power blocks created")


def _air(ic, board, study):
    pad = study["regionPaddingMm"]
    ic.modeler.delete("Region")
    reg = ic.modeler.create_region(pad_value=[pad] * 6, pad_type="Absolute Offset", name="Region")
    if not reg:
        check("error", "Air region not created")
        return
    ic.assign_free_opening([f.id for f in reg.faces], boundary_name="Open_air")
    ic.edit_design_settings(gravity_dir=4, ambient_temperature=study["ambientC"])   # 4 = -Z
    bb = reg.bounding_box
    check("ok", f"Air region {bb[3]-bb[0]:.0f} x {bb[4]-bb[1]:.0f} x {bb[5]-bb[2]:.0f} mm, open on all faces, "
                f"{study['ambientC']} C ambient, gravity -Z")


def _setup(ic, study):
    s = ic.create_setup("Study")
    if not s:
        check("error", "Setup not created")
        return
    for k, v in (("Flow Regime", "Laminar"), ("Include Flow", True), ("Include Gravity", True),
                 ("Convergence Criteria - Max Iterations", 300)):
        if k in s.props:
            s.props[k] = v
    s.update()
    mesh_ok = step("global mesh", lambda: ic.globalMeshSettings(int(study["meshLevel"])))
    check("ok" if mesh_ok else "warn", f"Setup 'Study' created, mesh level {study['meshLevel']}")


def _validate(ic):
    ok = step("validate", lambda: ic.validate_simple())
    check("ok" if ok else "warn", "AEDT validation passed" if ok else "AEDT validation reported problems")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    sys.exit(main(Path(sys.argv[1]).resolve()))
