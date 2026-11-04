"""Benchmark B0: the SO-8 package alone on a cold plate (BENCHMARK_SO8.md, sections 1.4, 4 and 6).

    run.ps1 benchmark/so8_coldplate.py [mesh level, default 4] [case ...]

Builds one Icepak project, benchmark/runs/SO8_B0_L<level>.aedt, with one design per case, solves them and writes
benchmark/runs/b0_L<level>.json plus a table on the screen. Conduction only: the paper's cold-plate model has
insulation all round, so the air here is still (flow off) and the outer walls are adiabatic. The cold plate is the
bottom face of the region, held at 20 C; the paste sits between it and the package.

Cases (all four live-bug ones and both dead-bug ones are run when none are named):
  dead_10, dead_5           package top on the plate, die attach 10 or 5 um
  live_10, live_5           leads and base on the plate, standoff gap under the body filled with paste
  liveair_10, liveair_5     the same with air in the standoff gap (the paste only under the feet and body edge)
  <case>_k<value>           any of these with another mould conductivity, e.g. dead_10_k0.9 (sensitivity check)

Axes: X along the package length (lead rows parallel to X), Y across (leads point to +Y and -Y), Z up.
"""
import json
import sys
import time
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))
import setup_study  # noqa: E402  the AEDT version (AEDT_VERSION, set by run.ps1), its port and the busy mark

VERSION = setup_study.VERSION
PORT = setup_study.PORT
SETUP = "Study"
RUNS = HERE / "runs"

PLATE_C = 20.0          # paper
PASTE_K = 3.0           # paper
PASTE_MM = 0.025        # ours: not given; the paper found the paste thickness did not matter
STANDOFF_MM = 0.10      # ours: JEDEC MS-012 A1 minimum
MEASURED = {"dead": 82.3, "live": 51.4}       # paper, theta_jc in C/W
POWER_W = {"dead": 0.6, "live": 1.0}          # about 50 C rise, the middle of the paper's three points

# W/mK, paper Table 2. Silicon is 117.5 - 0.42 (T - 100); 130 is its value near 70 C, the die temperature here.
MATERIALS = {
    "so8_silicon": 130.0,
    "so8_attach": 1.25,
    "so8_paddle": 260.0,
    "so8_mould": 0.62,
    "so8_inner_lf": 208.0,      # inner leadframe strips, volume averaged
    "so8_outer_lf": 96.0,       # leads outside the body, one block per lead row, volume averaged
    "so8_paste": PASTE_K,
}


def package(attach_mm):
    """[(name, material, (x0, x1), (y0, y1), (z0, z1))] with the body bottom at z = 0 and the lead feet below it."""
    die_z = 0.5 + attach_mm
    parts = [
        ("Body", "so8_mould", (-2.5, 2.5), (-2.0, 2.0), (0.0, 1.45)),
        ("Paddle", "so8_paddle", (-2.15, 2.15), (-1.115, 1.115), (0.3, 0.5)),
        ("Attach", "so8_attach", (-0.315, 0.315), (-0.315, 0.315), (0.5, die_z)),
        ("Die", "so8_silicon", (-0.315, 0.315), (-0.315, 0.315), (die_z, die_z + 0.4)),
    ]
    for s, side in ((1, "N"), (-1, "S")):
        def y(a, b):
            return (min(s * a, s * b), max(s * a, s * b))
        lx = (-2.11, 2.11)          # 4 leads at 1.27 pitch, 0.41 wide: 3 x 1.27 + 0.41 (ours, MS-012)
        parts += [
            (f"Strip{side}", "so8_inner_lf", (-2.15, 2.15), y(1.315, 2.0), (0.55, 0.75)),
            # gull wing (ours, MS-012: 6.0 over the leads): shoulder at the strip height, leg down, foot on the board
            (f"Shoulder{side}", "so8_outer_lf", lx, y(2.0, 2.4), (0.55, 0.75)),
            (f"Leg{side}", "so8_outer_lf", lx, y(2.4, 2.6), (-STANDOFF_MM, 0.75)),
            (f"Foot{side}", "so8_outer_lf", lx, y(2.6, 3.0), (-STANDOFF_MM, -STANDOFF_MM + 0.2)),
        ]
    return parts


def case_geometry(case):
    """global boxes for one case: package moved onto the paste, plus the paste itself"""
    kind, attach_um = case.split("_")[:2]
    attach = int(attach_um) / 1000.0
    boxes = []
    for name, mat, xs, ys, (z0, z1) in package(attach):
        if kind == "dead":      # turned over: the body top lies on the paste
            z0, z1 = PASTE_MM + 1.45 - z1, PASTE_MM + 1.45 - z0
        else:                   # the lead feet lie on the paste
            z0, z1 = z0 + PASTE_MM + STANDOFF_MM, z1 + PASTE_MM + STANDOFF_MM
        boxes.append((name, mould_name(case) if mat == "so8_mould" else mat, xs, ys, (z0, z1)))
    if kind == "dead":
        boxes.append(("Paste", "so8_paste", (-2.5, 2.5), (-2.0, 2.0), (0.0, PASTE_MM)))
    else:
        boxes.append(("Paste", "so8_paste", (-2.5, 2.5), (-3.0, 3.0), (0.0, PASTE_MM)))
        if kind == "live":
            boxes.append(("Gap", "so8_paste", (-2.5, 2.5), (-2.0, 2.0), (PASTE_MM, PASTE_MM + STANDOFF_MM)))
    return boxes


def mould_name(case):
    """material of the mould for this case: its own material when the case changes k (materials are per project)"""
    parts = case.split("_")
    return f"so8_mould_{parts[2]}" if len(parts) > 2 else "so8_mould"


def build(ic, case, level):
    ic.modeler.model_units = "mm"
    mats = dict(MATERIALS)
    if mould_name(case) != "so8_mould":
        mats[mould_name(case)] = float(case.split("_")[2][1:])
    for name, k in mats.items():
        if name not in ic.materials.material_keys:
            m = ic.materials.add_material(name)
            m.thermal_conductivity = k
            m.mass_density = 2000
            m.specific_heat = 800
    names = []
    for name, mat, xs, ys, zs in case_geometry(case):
        ic.modeler.create_box([xs[0], ys[0], zs[0]], [xs[1] - xs[0], ys[1] - ys[0], zs[1] - zs[0]],
                              name=name, material=mat)
        names.append(name)
    # the mould is cut around what sits inside it, so no two solids overlap
    ic.modeler.subtract("Body", ["Paddle", "Attach", "Die", "StripN", "StripS"], keep_originals=True)
    # around the package: 15 mm of still air to the sides and above, none below (the plate is there)
    ic.modeler.delete("Region")     # a new design comes with a default one
    reg =ic.modeler.create_region(pad_value=[15, 15, 15, 15, 15, 0], pad_type="Absolute Offset", name="Region")
    # the plate: the region's bottom face and the paste's own bottom face. On the region face alone the paste stayed
    # hot (the wall held only the air at 20 C), so a solid touching the region needs its face named too
    faces = [min(o.faces, key=lambda f: f.center[2]).id for o in (reg, ic.modeler["Paste"])]
    ic.assign_stationary_wall_with_temperature(faces, name="ColdPlate", temperature=f"{PLATE_C}cel")
    kind = case.split("_")[0].replace("liveair", "live")
    ic.assign_solid_block("Die", f"{POWER_W[kind]}W", boundary_name="Junction")
    ic.mesh.assign_mesh_region(names, level=level, name="Mesh_package")
    s = ic.create_setup(SETUP)
    for k, v in (("Include Flow", False), ("Include Gravity", False), ("Radiation Model", "Off"),
                 ("Convergence Criteria - Max Iterations", 500)):
        s.props[k] = v
    s.update()


def die_temp(ic, op):
    fr = ic.post.ofieldsreporter
    f = Path(ic.working_directory) / "die_t.fld"
    fr.CalcStack("clear")
    fr.CopyNamedExprToStack("Temperature")
    fr.EnterVol("Die")
    fr.CalcOp(op)
    fr.CalculatorWrite(str(f), ["Solution:=", f"{SETUP} : SteadyState"], [])
    fr.CalcStack("clear")
    v = float(f.read_text().split()[-1])
    f.unlink(missing_ok=True)
    return v


def main(level, cases):
    from ansys.aedt.core import Desktop, Icepak
    RUNS.mkdir(exist_ok=True)
    tag = f"L{level}" + ("_k" if any(mould_name(c) != "so8_mould" for c in cases) else "")   # a check keeps the main runs
    project = RUNS / f"SO8_B0_{tag}.aedt"
    desktop = Desktop(version=VERSION, new_desktop=False, port=PORT, non_graphical=False)
    try:
        if project.stem in list(desktop.odesktop.GetProjectList()):
            desktop.odesktop.CloseProject(project.stem)
        for p in [project, project.with_suffix(".aedt.lock")]:
            p.unlink(missing_ok=True)
        rows = []
        for case in cases:
            t = time.time()
            ic = Icepak(project=str(project), design=case, version=VERSION, new_desktop=False, port=PORT)
            build(ic, case, level)
            ic.save_project()
            ok = ic.analyze_setup(SETUP)
            kind = case.split("_")[0].replace("liveair", "live")
            p = POWER_W[kind]
            row = {"case": case, "powerW": p, "solved": bool(ok), "seconds": round(time.time() - t)}
            if ok:
                tmax, tmean = die_temp(ic, "Maximum"), die_temp(ic, "Mean")
                row.update(tjMaxC=round(tmax, 2), tjMeanC=round(tmean, 2),
                           thetaMaxCW=round((tmax - PLATE_C) / p, 1), thetaMeanCW=round((tmean - PLATE_C) / p, 1),
                           measuredCW=MEASURED[kind],
                           errorPct=round(100 * ((tmax - PLATE_C) / p - MEASURED[kind]) / MEASURED[kind], 1))
            rows.append(row)
            print(json.dumps(row), flush=True)
            ic.save_project()
        out = RUNS / f"b0_{tag}.json"
        out.write_text(json.dumps({"meshLevel": level, "plateC": PLATE_C, "pasteMm": PASTE_MM,
                                   "standoffMm": STANDOFF_MM, "rows": rows}, indent=1), encoding="utf-8")
        print(f"\n{'case':<12}{'P W':>6}{'Tj max':>9}{'theta max':>11}{'theta mean':>12}{'measured':>10}{'error':>8}")
        for r in rows:
            if r["solved"]:
                print(f"{r['case']:<12}{r['powerW']:>6}{r['tjMaxC']:>9}{r['thetaMaxCW']:>11}{r['thetaMeanCW']:>12}"
                      f"{r['measuredCW']:>10}{r['errorPct']:>7}%")
            else:
                print(f"{r['case']:<12} not solved")
        print(f"written {out}")
    finally:
        desktop.release_desktop(close_projects=False, close_on_exit=False)


if __name__ == "__main__":
    lvl = int(sys.argv[1]) if len(sys.argv) > 1 else 4
    todo = sys.argv[2:] or ["dead_10", "dead_5", "live_10", "live_5", "liveair_10", "liveair_5"]
    # while this runs, the page's jobs and its background read leave AEDT alone; it does not start during a page job
    other = setup_study.take_busy()
    if other:
        raise SystemExit(f"setup_study.py (process {other}) is using AEDT. Try again when the page's job has ended.")
    try:
        main(lvl, todo)
    finally:
        setup_study.drop_busy()
