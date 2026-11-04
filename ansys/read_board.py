"""Read an Ansys EDB (.aedb) and write board.json for the page.

    python ansys/read_board.py <folder>
    python ansys/read_board.py <folder> --step     only the heights and bodies again, from the STEP into board.json

<folder> holds one <board>.aedb (or study.json names it). Altium writes <board>.edb: it is copied to <board>.aedb,
because PyEDB and AEDT only open the .aedb name. Writes <folder>/board.json. Units in the file: mm.
A .step or .stp of the board in <folder> gives each part it holds stepHMm, stepXMm, stepYMm, stepAngleDeg (step_parts.py),
the body measured along the part's own axes: angleDeg, or the body's own angle when it sits clearly turned from that. lMm, wMm: the pad outline along those axes (footprint()).
Needs no AEDT session, only PyEDB (about 7 s).
"""
import json
import math
import os
import re
import shutil
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
VERSION = os.environ.get("AEDT_VERSION", "2026.1")   # start.bat / run.ps1 set it to the newest AEDT installed
M2MM = 1000.0
# W/mK, kg/m3, J/kgK. air: Icepak's own, because Icepak's air box uses the material "air" too, and an EDB names its
# fill "AIR" in the same project: FR-4 values there would make the air a solid
THERMAL = {"copper": (400, 8933, 385), "solder": (48, 8000, 167), "FR-4": (0.3, 1900, 1100), "air": (0.0261, 1.1614, 1006.43)}


def thermal_kind(name, conductor):
    """which THERMAL values stand in for a material that has none: air, copper, solder or FR-4"""
    n = name.lower()
    if n == "air" or n.startswith("air_"):
        return "air"
    if "solder" in n and "resist" not in n and "mask" not in n:
        return "solder"
    if "copper" in n or n.startswith("cu") or conductor:
        return "copper"
    return "FR-4"


def find_edb(folder: Path) -> Path:
    # an Altium .edb exported again is copied first, also when study.json names its .aedb (the page always does)
    for src in sorted(folder.glob("*.edb")):
        if src.is_dir() and (src / "edb.def").exists():
            _copy_edb(src, src.with_suffix(".aedb"))
    study = folder / "study.json"
    if study.exists():
        name = json.loads(study.read_text(encoding="utf-8")).get("edb")
        if name and (folder / name).exists():
            return folder / name
    found = sorted(folder.glob("*.aedb"))
    if not found:
        raise SystemExit(f"No .aedb or .edb folder in {folder}")
    return found[0]


def _copy_edb(src: Path, dst: Path):
    """the .aedb copy is refreshed when the .edb is exported again"""
    if (dst / "edb.def").exists() and (dst / "edb.def").stat().st_mtime >= (src / "edb.def").stat().st_mtime:
        return
    shutil.copytree(src, dst, dirs_exist_ok=True)
    print(f"Copied {src.name} to {dst.name} (AEDT only opens .aedb)", flush=True)


def footprint(c, pads):
    """(xMin, yMin, l, w, angle) in m and degrees: the pad outline of part c along its own axes, the axes nearest the
    board's, so a 15 x 15 part turned 45 deg is 15 x 15, not the 21.2 x 21.2 of its box on the board. Altium leaves
    the part's rotation at 0 and turns each pin, so the angle is the commonest pad angle (mod 90 deg, within +-45).
    pads caches the outline of each pad definition per layer"""
    rot = float(c.rotation or 0.0)
    corners, angles = [], {}
    for pin in c.pins.values():
        key = (pin.padstack_definition, c.placement_layer)
        if key not in pads:
            pad = pin.definition.pad_by_layer.get(c.placement_layer)
            pads[key] = _pad_outline(pad) if pad else None
        outline = pads[key]
        x, y = (float(v) for v in pin.position)
        if outline is None:          # a shape not read: its box on the board
            (bx0, by0), (bx1, by1) = pin.bounding_box
            corners += [(bx0, by0), (bx1, by1), (bx0, by1), (bx1, by0)]
            continue
        pts, round_ = outline
        a = rot + float(pin.rotation or 0.0)
        ca, sa = math.cos(a), math.sin(a)
        corners += [(x + px * ca - py * sa, y + px * sa + py * ca) for px, py in pts]
        if not round_:
            k = round(math.degrees(a) % 90.0, 1) % 90.0
            angles[k] = angles.get(k, 0) + 1
    if not corners:
        return None
    t = max(angles, key=angles.get, default=0.0)
    t = t - 90.0 if t > 45.0 else t
    ct, st = math.cos(math.radians(t)), math.sin(math.radians(t))
    local = [(px * ct + py * st, -px * st + py * ct) for px, py in corners]      # turned back by t
    u0, u1 = min(p[0] for p in local), max(p[0] for p in local)
    v0, v1 = min(p[1] for p in local), max(p[1] for p in local)
    um, vm = (u0 + u1) / 2, (v0 + v1) / 2
    cx, cy = um * ct - vm * st, um * st + vm * ct
    return cx - (u1 - u0) / 2, cy - (v1 - v0) / 2, u1 - u0, v1 - v0, t


_UNIT = {"": 1.0, "m": 1.0, "cm": 1e-2, "mm": 1e-3, "um": 1e-6, "nm": 1e-9, "mil": 25.4e-6, "in": 0.0254,
         "rad": 1.0, "deg": math.pi / 180}


def _si(text):
    """an EDB value as PyEDB gives it, a string such as '0', '1.5708', '90deg' or '0.1mm', in m or radians (a number
    without a unit is in those already). ValueError for anything else"""
    m = re.fullmatch(r"\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)\s*([A-Za-z]*)\s*", str(text or 0))
    if not m or m.group(2).lower() not in _UNIT:
        raise ValueError(f"EDB value not read: {text!r}")
    return float(m.group(1)) * _UNIT[m.group(2).lower()]


def _f(v):
    """a coordinate: a number, or an EDB Value (PyEDB over gRPC gives those for polygon points)"""
    return float(v) if isinstance(v, (int, float, str)) else float(getattr(v, "value", None) if hasattr(v, "value") else str(v))


def _pad_outline(pad):
    """(points in the pin's frame, round) of a pad, or None for a shape not read"""
    shape, v = pad.shape, [float(x) for x in (pad.parameters_values or [])]
    if shape == "polygon" and pad.polygon_data is not None:
        # an arc is stored as a point with x near 1e308 and y its height: not a corner, left out
        pts = [(x, y) for x, y in ((_f(p.x), _f(p.y)) for p in pad.polygon_data.points) if abs(x) < 1e10 and abs(y) < 1e10]
        if not pts:
            return None
    elif shape == "circle" and v:
        return [(v[0] / 2 * math.cos(i * math.pi / 4), v[0] / 2 * math.sin(i * math.pi / 4)) for i in range(8)], True
    elif shape == "square" and v:
        pts = [(sx * v[0] / 2, sy * v[0] / 2) for sx in (-1, 1) for sy in (-1, 1)]
    elif shape in ("rectangle", "oval", "bullet") and len(v) >= 2:
        pts = [(sx * v[0] / 2, sy * v[1] / 2) for sx in (-1, 1) for sy in (-1, 1)]
    else:
        return None
    try:
        r, ox, oy = _si(pad.rotation), _si(pad.offset_x), _si(pad.offset_y)
    except ValueError:      # an expression: the pin's box on the board stands in, as for a shape not read
        return None
    cr, sr = math.cos(r), math.sin(r)
    return [(ox + px * cr - py * sr, oy + px * sr + py * cr) for px, py in pts], False


def read_board(edb_path: Path) -> dict:
    from pyedb import Edb

    edb = Edb(str(edb_path), version=VERSION)
    try:
        (x0, y0), (x1, y1) = edb.get_bounding_box()
        layers, used = [], {}
        for name, lay in edb.stackup.layers.items():
            try:
                layers.append({"name": name, "type": str(lay.type), "thicknessMm": round(lay.thickness * M2MM, 4)})
                for mat in {lay.material, getattr(lay, "fill_material", None)} - {None, ""}:
                    used.setdefault(mat, []).append(name)
            except Exception:
                pass
        materials = []      # an electrical export (Altium) leaves the thermal values at 0; the page shows what stands in
        for name, m in edb.materials.materials.items():
            if name in used:
                materials.append({"name": name, "layers": used[name], "kind": thermal_kind(name, (m.conductivity or 0) > 1e5),
                                  "kWmK": m.thermal_conductivity or 0, "rhoKgM3": m.mass_density or 0, "cpJkgK": m.specific_heat or 0})
        parts, skipped, pads = [], [], {}
        for ref, c in edb.components.instances.items():
            try:
                fp = footprint(c, pads)
                if fp is None:
                    bx0, by0, bx1, by1 = c.bounding_box
                    fp = bx0, by0, bx1 - bx0, by1 - by0, 0.0
                z0, z1 = sorted((c.lower_elevation, c.upper_elevation))
                parts.append({
                    "refDes": ref,
                    "part": str(c.part_name),
                    "type": str(c.type),
                    "side": "top" if c.is_top_mounted else "bottom",   # from the stackup order, not the layer name
                    "xMinMm": round(fp[0] * M2MM, 4), "yMinMm": round(fp[1] * M2MM, 4),
                    "lMm": round(fp[2] * M2MM, 4), "wMm": round(fp[3] * M2MM, 4), "angleDeg": round(fp[4], 1),
                    "zMinMm": round(z0 * M2MM, 4), "hMm": round((z1 - z0) * M2MM, 4),
                })
            except Exception:
                skipped.append(ref)
        parts.sort(key=lambda p: -(p["lMm"] * p["wMm"]))
        return {
            "schema": "pcb-thermal-ansys/board",
            "edb": edb_path.name,
            "name": edb_path.stem,
            "outlineMm": {"xMin": round(x0 * M2MM, 4), "yMin": round(y0 * M2MM, 4),
                          "xMax": round(x1 * M2MM, 4), "yMax": round(y1 * M2MM, 4)},
            "thicknessMm": round(sum(l["thicknessMm"] for l in layers), 4),
            "layers": layers,
            "materials": materials,
            "parts": parts,
            "skipped": skipped,
        }
    finally:
        edb.close_edb()


def add_step(folder: Path, board: dict):
    """heights and body sizes from the board's STEP (Altium: File > Export > STEP 3D), the newest .step/.stp in the folder"""
    steps = sorted((f for f in folder.iterdir() if f.suffix.lower() in (".step", ".stp")), key=lambda f: f.stat().st_mtime)
    if not steps:
        print("No .step or .stp in the folder: heights stay the EDB's placeholders", flush=True)
        return
    import step_parts
    try:
        rep = {}
        found = step_parts.for_board(steps[-1], board, rep)
    except Exception as ex:  # noqa: BLE001  a STEP the reader does not follow must not stop the board read
        print(f"[warn] {steps[-1].name} could not be read ({type(ex).__name__}: {ex}); heights stay empty", flush=True)
        return
    for p in board["parts"]:
        if p["refDes"] in found:
            f = found[p["refDes"]]
            p.update(stepHMm=f["hMm"], stepXMm=f["xMm"], stepYMm=f["yMm"], stepAngleDeg=f["aDeg"])
    board["step"] = {"file": steps[-1].name, "parts": len(found), "byPosition": sum(f["by"] == "position" for f in found.values()), **rep}
    print(f"{steps[-1].name}: {rep['components']} components besides the board, {len(found)} of them matched to a part"
          f" (the EDB has {len(board['parts'])} parts, many without a 3D body: test points, fiducials, holes)", flush=True)
    for a in rep["away"]:
        print(f"  away from its pads: {a['refDes']} body at {a['atMm']}, pads at {a['padsMm']} ({a['distMm']} mm)", flush=True)
    for u in rep["unmatched"]:
        print(f"  not used: {u['name']} ({u['xMm']} x {u['yMm']} x {u['hMm']} mm at {u['atMm']}): {u['why']}", flush=True)


def step_only(folder: Path) -> Path:
    """the STEP again, without PyEDB: a new STEP export, or a fix in this reader, without reading the EDB again"""
    out = folder / "board.json"
    if not out.exists():
        raise SystemExit("No board.json yet: read the board first")
    board = json.loads(out.read_text(encoding="utf-8"))
    board.pop("step", None)
    for p in board["parts"]:
        for k in ("stepHMm", "stepXMm", "stepYMm", "stepAngleDeg"):
            p.pop(k, None)
    add_step(folder, board)
    out.write_text(json.dumps(board, indent=1), encoding="utf-8")
    return out


def main(folder: Path) -> Path:
    edb = find_edb(folder)
    board = read_board(edb)
    add_step(folder, board)
    out = folder / "board.json"
    out.write_text(json.dumps(board, indent=1), encoding="utf-8")
    return out


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3) or sys.argv[2:] not in ([], ["--step"]):
        raise SystemExit(__doc__)
    print((step_only if sys.argv[2:] else main)(Path(sys.argv[1]).resolve()))
