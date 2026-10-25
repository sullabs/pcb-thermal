"""Read an Ansys EDB (.aedb) and write board.json for the page.

    python ansys/read_board.py <folder>

<folder> holds one <board>.aedb (or study.json names it). Altium writes <board>.edb: it is copied to <board>.aedb,
because PyEDB and AEDT only open the .aedb name. Writes <folder>/board.json. Units in the file: mm.
Needs no AEDT session, only PyEDB (about 7 s).
"""
import json
import shutil
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
VERSION = "2026.1"
M2MM = 1000.0
THERMAL = {"copper": (400, 8933, 385), "solder": (48, 8000, 167), "FR-4": (0.3, 1900, 1100)}   # W/mK, kg/m3, J/kgK


def thermal_kind(name, conductor):
    """which THERMAL values stand in for a material that has none: copper, solder or FR-4"""
    n = name.lower()
    if "solder" in n and "resist" not in n and "mask" not in n:
        return "solder"
    if "copper" in n or n.startswith("cu") or conductor:
        return "copper"
    return "FR-4"


def find_edb(folder: Path) -> Path:
    study = folder / "study.json"
    if study.exists():
        name = json.loads(study.read_text(encoding="utf-8")).get("edb")
        if name and (folder / name).exists():
            return folder / name
    for src in sorted(folder.glob("*.edb")):
        if src.is_dir() and (src / "edb.def").exists():
            _copy_edb(src, src.with_suffix(".aedb"))
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
        parts, skipped = [], []
        for ref, c in edb.components.instances.items():
            try:
                bx0, by0, bx1, by1 = c.bounding_box
                z0, z1 = sorted((c.lower_elevation, c.upper_elevation))
                parts.append({
                    "refDes": ref,
                    "part": str(c.part_name),
                    "type": str(c.type),
                    "side": "top" if c.is_top_mounted else "bottom",   # from the stackup order, not the layer name
                    "xMinMm": round(bx0 * M2MM, 4), "yMinMm": round(by0 * M2MM, 4),
                    "lMm": round((bx1 - bx0) * M2MM, 4), "wMm": round((by1 - by0) * M2MM, 4),
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


def main(folder: Path) -> Path:
    edb = find_edb(folder)
    board = read_board(edb)
    out = folder / "board.json"
    out.write_text(json.dumps(board, indent=1), encoding="utf-8")
    return out


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    print(main(Path(sys.argv[1]).resolve()))
