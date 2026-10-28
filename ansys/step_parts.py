"""Body boxes of the parts in a board's STEP file (Altium: File > Export > STEP 3D). Python standard library only.

    python ansys/step_parts.py <board.step>      prints {name: [xMin, yMin, zMin, xMax, yMax, zMax]} in mm
    read_board.py calls for_board() when the board folder holds a .step or .stp next to the EDB.

Walks the assembly tree, applies each placement, and boxes the solids under every product: the box of a product
holds its own solids and those of its children. The box comes from the vertices, the arcs of circles and ellipses
(sampled) and the control points of splines, so a dome that bulges past its edges is boxed a little short.
"""
import json
import math
import re
import sys
from pathlib import Path
from statistics import median

_TOKEN = re.compile(r"'(?:[^']|'')*'|#\d+|[A-Za-z_][A-Za-z0-9_]*|\.[A-Za-z_0-9]+\.|[-+0-9.Ee]+|[(),$*]")
SOLIDS = {"MANIFOLD_SOLID_BREP", "BREP_WITH_VOIDS", "SHELL_BASED_SURFACE_MODEL", "FACETED_BREP"}


class Ref(int):
    pass


def _args(tokens, i):
    """the list that starts at tokens[i] == '(' ; returns (list, index after ')')"""
    out, i = [], i + 1
    while tokens[i] != ")":
        t = tokens[i]
        if t == "(":
            v, i = _args(tokens, i)
        elif t[0] == "'":
            v, i = t[1:-1].replace("''", "'"), i + 1
        elif t[0] == "#":
            v, i = Ref(t[1:]), i + 1
        elif t[0].isalpha() and tokens[i + 1] == "(":     # a typed value: LENGTH_MEASURE(1.)
            v, i = _args(tokens, i + 1)
            v = v[0] if len(v) == 1 else v
        elif t in ("$", "*") or t[0] == ".":
            v, i = t, i + 1
        else:
            v, i = float(t), i + 1
        out.append(v)
        if tokens[i] == ",":
            i += 1
    return out, i + 1


def read(path):
    """{id: {TYPE: args}}; a complex entity has several TYPEs"""
    text = Path(path).read_text(encoding="latin-1")
    data = text[text.index("DATA;") + 5:text.rindex("ENDSEC;")]
    ents = {}
    for m in re.finditer(r"#(\d+)\s*=\s*((?:'(?:[^']|'')*'|[^;'])*);", data):
        tokens = _TOKEN.findall(m.group(2))
        parts, i = {}, 0
        if tokens[0] == "(":           # ( A(..) B(..) )
            i = 1
            while tokens[i] != ")":
                parts[tokens[i]], i = _args(tokens, i + 1)
        else:
            parts[tokens[0]], _ = _args(tokens, 1)
        ents[int(m.group(1))] = parts
    return ents


def _one(e):
    return next(iter(e.items()))


# ---- 3D maths on 3x4 matrices [[r00, r01, r02, tx], ...]
def _placement(ents, ref):
    """AXIS2_PLACEMENT_3D as a matrix"""
    _, a = _one(ents[ref])
    o = ents[a[1]]["CARTESIAN_POINT"][1]
    z = _unit(ents[a[2]]["DIRECTION"][1]) if isinstance(a[2], Ref) else [0.0, 0.0, 1.0]
    x = ents[a[3]]["DIRECTION"][1] if len(a) > 3 and isinstance(a[3], Ref) else [1.0, 0.0, 0.0]
    d = sum(x[k] * z[k] for k in range(3))
    x = _unit([x[k] - d * z[k] for k in range(3)])
    if not any(x):
        x = _unit([1.0, 0.0, 0.0] if abs(z[0]) < 0.9 else [0.0, 1.0, 0.0])
    y = [z[1] * x[2] - z[2] * x[1], z[2] * x[0] - z[0] * x[2], z[0] * x[1] - z[1] * x[0]]
    return [[x[r], y[r], z[r], o[r]] for r in range(3)]


def _unit(v):
    n = math.sqrt(sum(c * c for c in v)) or 1.0
    return [c / n for c in v]


def _mul(a, b):
    return [[sum(a[r][k] * b[k][c] for k in range(3)) + (a[r][3] if c == 3 else 0.0) for c in range(4)] for r in range(3)]


def _inv(a):
    rt = [[a[c][r] for c in range(3)] for r in range(3)]
    t = [-sum(rt[r][k] * a[k][3] for k in range(3)) for r in range(3)]
    return [rt[r] + [t[r]] for r in range(3)]


def _apply(m, p):
    return [m[r][0] * p[0] + m[r][1] * p[1] + m[r][2] * p[2] + m[r][3] for r in range(3)]


I = [[1.0, 0.0, 0.0, 0.0], [0.0, 1.0, 0.0, 0.0], [0.0, 0.0, 1.0, 0.0]]


def _arc(ents, curve, p0, p1, sense):
    """points along a circle or ellipse from p0 to p1 (the whole curve when they are the same point)"""
    kind, a = _one(ents[curve])
    m = _placement(ents, a[1])
    rx, ry = (a[2], a[2]) if kind == "CIRCLE" else (a[2], a[3])
    loc = lambda p: [sum((p[k] - m[k][3]) * m[k][c] for k in range(3)) for c in range(2)]  # noqa: E731
    t0, t1 = (math.atan2(q[1] / ry, q[0] / rx) for q in (loc(p0), loc(p1)))
    if not sense:
        t0, t1 = t1, t0
    span = (t1 - t0) % (2 * math.pi) or 2 * math.pi
    return [_apply(m, [rx * math.cos(t0 + span * i / 32), ry * math.sin(t0 + span * i / 32), 0.0]) for i in range(33)]


def _solid_points(ents, ref, seen):
    """the box points of one solid, in its representation's frame"""
    pts, stack = [], [ref]
    while stack:
        r = stack.pop()
        if r in seen:
            continue
        seen.add(r)
        kind, a = _one(ents[r])
        if kind == "VERTEX_POINT":
            pts.append(ents[a[1]]["CARTESIAN_POINT"][1])
        elif kind == "EDGE_CURVE":
            ck, ca = _one(ents[a[3]])
            if ck in ("CIRCLE", "ELLIPSE"):
                p0, p1 = (ents[ents[v]["VERTEX_POINT"][1]]["CARTESIAN_POINT"][1] for v in a[1:3])
                pts += _arc(ents, a[3], p0, p1, a[4] != ".F.")
            elif "B_SPLINE_CURVE" in ents[a[3]]:        # rational: a complex entity, no name in this part
                pts += [ents[p]["CARTESIAN_POINT"][1] for p in ents[a[3]]["B_SPLINE_CURVE"][1]]
            elif ck.startswith("B_SPLINE_CURVE"):
                pts += [ents[p]["CARTESIAN_POINT"][1] for p in ca[2]]
            stack += [a[1], a[2]]
        elif kind == "POLY_LOOP":
            pts += [ents[p]["CARTESIAN_POINT"][1] for p in a[1]]
        else:   # shells, faces, loops: walk their references, never into surfaces and placements
            if kind in ("ADVANCED_FACE", "FACE_SURFACE"):
                sk, sa = _one(ents[a[2]])
                if sk == "SPHERICAL_SURFACE":       # a dome rises above its edges: box the whole sphere
                    c, r = _placement(ents, sa[1]), sa[2]
                    pts += [[c[0][3] + sx * r, c[1][3] + sy * r, c[2][3] + sz * r] for sx in (-1, 1) for sy in (-1, 1) for sz in (-1, 1)]
                a = a[:2]
            stack += [x for x in _refs(a)]
    return pts


def _refs(v):
    if isinstance(v, Ref):
        yield v
    elif isinstance(v, list):
        for x in v:
            yield from _refs(x)


def boxes(path):
    """{product name: [xMin, yMin, zMin, xMax, yMax, zMax]} in mm, in the frame of the top assembly"""
    ents = read(path)
    scale = 1.0
    for e in ents.values():
        if "LENGTH_UNIT" in e and "SI_UNIT" in e:
            scale = {".MILLI.": 1.0, ".CENTI.": 10.0, "$": 1000.0}.get(e["SI_UNIT"][0], 1.0)
            break
    by_type = {}
    for i, e in ents.items():
        for k in e:
            by_type.setdefault(k, []).append(i)
    name_of_pd = {}         # PRODUCT_DEFINITION -> product name
    for i in by_type.get("PRODUCT_DEFINITION", []):
        pdf = ents[i]["PRODUCT_DEFINITION"][2]
        name_of_pd[i] = ents[ents[pdf]["PRODUCT_DEFINITION_FORMATION"][2]]["PRODUCT"][0]
    rep_of_pd = {}          # PRODUCT_DEFINITION -> its shape representation
    for i in by_type.get("SHAPE_DEFINITION_REPRESENTATION", []):
        pds, rep = ents[i]["SHAPE_DEFINITION_REPRESENTATION"]
        d = ents[pds]["PRODUCT_DEFINITION_SHAPE"][2]
        if d in name_of_pd:
            rep_of_pd[d] = rep
    pd_of_rep = {r: d for d, r in rep_of_pd.items()}
    children = {}           # parent rep -> [(child rep, transform child -> parent, child product def)]
    for i in by_type.get("CONTEXT_DEPENDENT_SHAPE_REPRESENTATION", []):
        rel, pds = ents[i]["CONTEXT_DEPENDENT_SHAPE_REPRESENTATION"]
        rr = ents[rel]["REPRESENTATION_RELATIONSHIP"]
        idt = ents[ents[rel]["REPRESENTATION_RELATIONSHIP_WITH_TRANSFORMATION"][0]]["ITEM_DEFINED_TRANSFORMATION"]
        child, parent = rr[2], rr[3]
        t = _mul(_placement(ents, idt[3]), _inv(_placement(ents, idt[2])))
        nauo = ents[ents[pds]["PRODUCT_DEFINITION_SHAPE"][2]]
        child_pd = nauo["NEXT_ASSEMBLY_USAGE_OCCURRENCE"][4] if "NEXT_ASSEMBLY_USAGE_OCCURRENCE" in nauo else pd_of_rep.get(child)
        children.setdefault(parent, []).append((child, t, child_pd))
    # a representation may also hold its solids through a plain SHAPE_REPRESENTATION_RELATIONSHIP (no transform)
    linked = {}
    for i in by_type.get("SHAPE_REPRESENTATION_RELATIONSHIP", []):
        if "REPRESENTATION_RELATIONSHIP_WITH_TRANSFORMATION" not in ents[i]:
            rr = ents[i].get("SHAPE_REPRESENTATION_RELATIONSHIP") or ents[i]["REPRESENTATION_RELATIONSHIP"]
            linked.setdefault(rr[2], []).append(rr[3])
            linked.setdefault(rr[3], []).append(rr[2])
    local = {}

    def own_points(rep):
        if rep not in local:
            pts, seen, todo, done = [], set(), [rep], set()
            while todo:
                r = todo.pop()
                if r in done:
                    continue
                done.add(r)
                todo += linked.get(r, [])
                for item in _one(ents[r])[1][1]:
                    if _one(ents[item])[0] in SOLIDS:
                        pts += _solid_points(ents, item, seen)
            local[rep] = pts
        return local[rep]

    child_reps = {c for v in children.values() for c, _, _ in v}
    roots = [d for d, r in rep_of_pd.items() if r not in child_reps]
    out = {}

    def walk(rep, m, names):
        pts = [_apply(m, p) for p in own_points(rep)]
        for child, t, pd in children.get(rep, []):
            pts += walk(child, _mul(m, t), names + [name_of_pd.get(pd, "")])
        if pts and names:
            box = [min(p[k] for p in pts) * scale for k in range(3)] + [max(p[k] for p in pts) * scale for k in range(3)]
            out.setdefault(names[-1], []).append(box)
        return pts

    for d in roots:
        walk(rep_of_pd[d], I, [name_of_pd[d]])
    return out


def for_board(path, board):
    """{RefDes: {hMm, xMm, yMm, by}} for the parts of board.json found in the STEP: height above the board face the part
    sits on, body size along the board axes. by: "name" (a product named like the RefDes) or "position" (the biggest body
    centred on the part's pads, for exporters that name products after the model)"""
    found = boxes(path)
    centre = lambda b: ((b[0] + b[3]) / 2, (b[1] + b[4]) / 2)  # noqa: E731
    edb = {p["refDes"]: p for p in board["parts"]}
    named = {r: found[r][0] for r in edb if len(found.get(r, [])) == 1}
    # the STEP origin may differ from the EDB one: shift by the median offset of the parts found by name
    off = [median(edb[r]["xMinMm"] + edb[r]["lMm"] / 2 - centre(b)[0] for r, b in named.items()),
           median(edb[r]["yMinMm"] + edb[r]["wMm"] / 2 - centre(b)[1] for r, b in named.items())] if named else [0.0, 0.0]
    near = lambda p, b, grow: (p["xMinMm"] - grow <= centre(b)[0] + off[0] <= p["xMinMm"] + p["lMm"] + grow and  # noqa: E731
                               p["yMinMm"] - grow <= centre(b)[1] + off[1] <= p["yMinMm"] + p["wMm"] + grow)
    hits = {r: (b, "name") for r, b in named.items() if near(edb[r], b, 2.0)}    # a name far from its pads is someone else's
    every = [b for v in found.values() for b in v]
    for r, p in edb.items():
        if r not in hits:
            area = (p["lMm"] + 1) * (p["wMm"] + 1) * 4         # bigger than that is the board or a whole sub-assembly
            fit = [b for b in every if near(p, b, 0.5) and (b[3] - b[0]) * (b[4] - b[1]) <= area]
            if fit:
                hits[r] = (max(fit, key=lambda b: (b[3] - b[0]) * (b[4] - b[1]) * (b[5] - b[2])), "position")
    # the face a part sits on: the commonest body bottom of the top-side parts; the bottom face is a board thickness lower
    tops = [round(b[2], 2) for r, (b, _) in hits.items() if edb[r]["side"] == "top"]
    top = max(set(tops), key=tops.count) if tops else max(b[2] for b, _ in hits.values()) if hits else 0.0
    bottom = top - board["thicknessMm"]
    out = {}
    for r, (b, by) in hits.items():
        h = b[5] - b[2]                     # a part off the board faces (a rigid-flex strip): its own height
        if edb[r]["side"] == "top" and -0.05 <= b[2] - top <= 0.5:
            h = b[5] - top                   # standoff included: the block stands on the board
        elif edb[r]["side"] == "bottom" and -0.05 <= bottom - b[5] <= 0.5:
            h = bottom - b[2]
        out[r] = {"hMm": round(h, 3), "xMm": round(b[3] - b[0], 3), "yMm": round(b[4] - b[1], 3), "by": by}
    return out


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        raise SystemExit(__doc__ + "\n    python ansys/step_parts.py <board.step> <board.json>   heights and bodies per RefDes")
    if len(sys.argv) == 3:
        print(json.dumps(for_board(sys.argv[1], json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))), indent=1))
    else:
        print(json.dumps({k: [[round(x, 3) for x in b] for b in v] for k, v in boxes(sys.argv[1]).items()}, indent=1))
