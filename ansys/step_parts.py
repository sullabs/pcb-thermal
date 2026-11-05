"""Body boxes of the parts in a board's STEP file (Altium: File > Export > STEP 3D). Python standard library only.

    python ansys/step_parts.py <board.step>      prints {name: [xMin, yMin, zMin, xMax, yMax, zMax]} in mm
    read_board.py calls for_board() when the board folder holds a .step or .stp next to the EDB.

Walks the assembly tree, applies each placement, and boxes the solids under every product: the box of a product
holds its own solids and those of its children. The components are the products on the first level under the top
assembly, each boxed whole with its sub-parts, never a product below them; they match a RefDes by name (U1, C7, VR8;
SolidWorks' "U1-1" too) or by position. The box comes from the vertices, the arcs of circles and ellipses
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
    out = {}
    for name, box, *_ in tree(path):
        out.setdefault(name, []).append(box)
    return out


def refdes(name, refs):
    """the RefDes a product is named after ("U1", also "U1-1" or "U1<1>" from SolidWorks), or None"""
    up = {r.upper(): r for r in refs}
    name = name.split("/")[-1].strip().upper()
    return up.get(name) or up.get(re.sub(r"(-\d+|<\d+>)$", "", name))


def components(path, refs=()):
    """[(name, box + [outline])] of the components: the products on the first level under the top assembly, each boxed whole with
    its sub-parts, never a product below them. Named by the RefDes in refs when the occurrence (SolidWorks: "U1-1") or
    its product (Altium: "U1") is named like one, else by the product name (the board itself, a model name)"""
    nodes = tree(path)
    roots = {i for i, n in enumerate(nodes) if n[2] is None}
    return [(refdes(inst, refs) or refdes(name, refs) or name, box + [hull]) for name, box, parent, inst, hull in nodes if parent in roots]


def tree(path):
    """[(product name, box, index of the parent in this list or None, occurrence name, outline)], one per occurrence,
    parents after children; outline: the convex hull of the body seen from above, [(x, y)] in mm"""
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
        # PRODUCT_DEFINITION_FORMATION, or a subtype (SolidWorks and Creo: ..._WITH_SPECIFIED_SOURCE): its product is the 3rd
        formation = next(v for k, v in ents[pdf].items() if k.startswith("PRODUCT_DEFINITION_FORMATION"))
        name_of_pd[i] = ents[formation[2]]["PRODUCT"][0]
    rep_of_pd = {}          # PRODUCT_DEFINITION -> its shape representation
    for i in by_type.get("SHAPE_DEFINITION_REPRESENTATION", []):
        pds, rep = ents[i]["SHAPE_DEFINITION_REPRESENTATION"]
        d = ents[pds]["PRODUCT_DEFINITION_SHAPE"][2]
        if d in name_of_pd:
            rep_of_pd[d] = rep
    pd_of_rep = {r: d for d, r in rep_of_pd.items()}
    children = {}           # parent rep -> [(child rep, transform child -> parent, child product def, occurrence name)]
    for i in by_type.get("CONTEXT_DEPENDENT_SHAPE_REPRESENTATION", []):
        rel, pds = ents[i]["CONTEXT_DEPENDENT_SHAPE_REPRESENTATION"]
        rr = ents[rel]["REPRESENTATION_RELATIONSHIP"]
        idt = ents[ents[rel]["REPRESENTATION_RELATIONSHIP_WITH_TRANSFORMATION"][0]]["ITEM_DEFINED_TRANSFORMATION"]
        child, parent = rr[2], rr[3]
        t = _mul(_placement(ents, idt[3]), _inv(_placement(ents, idt[2])))
        nauo = ents[ents[pds]["PRODUCT_DEFINITION_SHAPE"][2]]
        nu = nauo.get("NEXT_ASSEMBLY_USAGE_OCCURRENCE")
        child_pd = nu[4] if nu else pd_of_rep.get(child)
        inst = (nu[1] or nu[0]) if nu and isinstance(nu[1], str) and isinstance(nu[0], str) else ""
        children.setdefault(parent, []).append((child, t, child_pd, inst))
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

    child_reps = {c for v in children.values() for c, *_ in v}
    roots = [d for d, r in rep_of_pd.items() if r not in child_reps]
    out = []

    def walk(rep, m, name, inst=""):
        """the points under this occurrence; adds it to out and returns its index there (None without solids)"""
        pts, kids = [_apply(m, p) for p in own_points(rep)], []
        for child, t, pd, ci in children.get(rep, []):
            p, k = walk(child, _mul(m, t), name_of_pd.get(pd, ""), ci)
            pts += p
            kids.append(k)
        if not pts:
            return pts, None
        out.append([name, [min(p[k] for p in pts) * scale for k in range(3)] + [max(p[k] for p in pts) * scale for k in range(3)], None, inst,
                    [(x * scale, y * scale) for x, y in _hull(pts)]])
        for k in kids:
            if k is not None:
                out[k][2] = len(out) - 1
        return pts, len(out) - 1

    for d in roots:
        walk(rep_of_pd[d], I, name_of_pd[d])
    return [tuple(n) for n in out]


def _hull(pts):
    """the convex hull of the points seen from above (monotone chain), [(x, y)]"""
    ps = sorted({(round(p[0], 6), round(p[1], 6)) for p in pts})
    if len(ps) < 3:
        return ps
    cross = lambda o, a, b: (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])  # noqa: E731
    lo, hi = [], []
    for p in ps:
        while len(lo) >= 2 and cross(lo[-2], lo[-1], p) <= 0:
            lo.pop()
        lo.append(p)
    for p in reversed(ps):
        while len(hi) >= 2 and cross(hi[-2], hi[-1], p) <= 0:
            hi.pop()
        hi.append(p)
    return lo[:-1] + hi[:-1]


def size_at(outline, deg):
    """(x, y) size of an outline along axes turned deg from the board's"""
    c, s = math.cos(math.radians(deg)), math.sin(math.radians(deg))
    u = [x * c + y * s for x, y in outline]
    v = [-x * s + y * c for x, y in outline]
    return max(u) - min(u), max(v) - min(v)


def body_angle(outline, hint):
    """angle of the part's own axes in degrees, within +-45: hint (from its pads) unless the body's smallest enclosing
    rectangle is clearly turned from it, as when the EDB carries no rotation for the part"""
    area = lambda a: math.prod(size_at(outline, a))  # noqa: E731
    sides = [((math.degrees(math.atan2(q[1] - p[1], q[0] - p[0])) + 45.0) % 90.0) - 45.0
             for p, q in zip(outline, outline[1:] + outline[:1]) if p != q]
    best = min(sides, key=area, default=hint)
    return round(best, 1) if area(best) < 0.95 * area(hint) else hint


def for_board(path, board, report=None):
    """{RefDes: {hMm, xMm, yMm, aDeg, by}} for the parts of board.json found in the STEP: height above the board face the
    part sits on, body size along the part's own axes turned aDeg (angleDeg of board.json, from its pads, unless the body
    itself sits clearly turned from that: body_angle()). Only whole components count (components()), never one of their sub-parts.
    by: "name" (the one component named like the RefDes, wherever it sits) or "position" (the component centred on the
    part's pads, for exporters that name components after the model). report, a dict, gets
      components: how many components the STEP holds, the board itself left out
      unmatched: the components that are no part's, [{name, xMm, yMm, hMm, atMm (centre, EDB frame), why}]
      away: the parts matched by name whose body is not over their pads (a STEP from another revision of the layout,
            or a 3D model off its footprint), [{refDes, distMm, atMm (body centre), padsMm (pad centre)}]"""
    comps = components(path, [p["refDes"] for p in board["parts"]])
    found = {}
    for name, box in comps:
        found.setdefault(name, []).append(box)
    centre = lambda b: ((b[0] + b[3]) / 2, (b[1] + b[4]) / 2)  # noqa: E731
    edb = {p["refDes"]: p for p in board["parts"]}
    named = {r: found[r][0] for r in edb if len(found.get(r, [])) == 1}
    # the STEP origin may differ from the EDB one: shift by the median offset of the parts found by name
    off = [median(edb[r]["xMinMm"] + edb[r]["lMm"] / 2 - centre(b)[0] for r, b in named.items()),
           median(edb[r]["yMinMm"] + edb[r]["wMm"] / 2 - centre(b)[1] for r, b in named.items())] if named else [0.0, 0.0]
    near = lambda p, b, grow: (p["xMinMm"] - grow <= centre(b)[0] + off[0] <= p["xMinMm"] + p["lMm"] + grow and  # noqa: E731
                               p["yMinMm"] - grow <= centre(b)[1] + off[1] <= p["yMinMm"] + p["wMm"] + grow)
    over = lambda p, b: (b[0] + off[0] <= p["xMinMm"] + p["lMm"] + 1.0 and p["xMinMm"] - 1.0 <= b[3] + off[0] and  # noqa: E731
                         b[1] + off[1] <= p["yMinMm"] + p["wMm"] + 1.0 and p["yMinMm"] - 1.0 <= b[4] + off[1])
    # a RefDes is unique: its body counts wherever it sits, only its height and size are used, the place is the EDB's
    hits = {r: (b, "name") for r, b in named.items()}
    unmatched, away = [], []
    if report is not None:
        for r, b in named.items():
            p = edb[r]
            if not over(p, b) and not near(p, b, 2.0):
                c, pc = (centre(b)[0] + off[0], centre(b)[1] + off[1]), (p["xMinMm"] + p["lMm"] / 2, p["yMinMm"] + p["wMm"] / 2)
                away.append({"refDes": r, "distMm": round(math.dist(c, pc), 1), "atMm": [round(v, 2) for v in c],
                             "padsMm": [round(v, 2) for v in pc]})
    o = board["outlineMm"]
    half = (o["xMax"] - o["xMin"]) * (o["yMax"] - o["yMin"]) / 2      # bigger than that is the board itself
    free = [b for _, b in comps if (b[3] - b[0]) * (b[4] - b[1]) < half and not any(b is h for h, _ in hits.values())]
    # the board's mid-plane, so a body under the board is not taken for a top part above it: from the board body, else
    # from the top-side parts found by name, else not known
    slab = max((b for _, b in comps if (b[3] - b[0]) * (b[4] - b[1]) >= half), key=lambda b: (b[3] - b[0]) * (b[4] - b[1]), default=None)
    named_tops = [round(b[2], 2) for r, b in named.items() if edb[r]["side"] == "top"]
    if slab:
        mid = (slab[2] + slab[5]) / 2
    elif named_tops:
        mid = max(set(named_tops), key=named_tops.count) - board["thicknessMm"] / 2
    else:
        mid = None
    side_ok = lambda p, b: mid is None or ((b[2] + b[5]) / 2 >= mid) == (p["side"] == "top")  # noqa: E731
    for r, p in sorted(edb.items(), key=lambda rp: rp[1]["lMm"] * rp[1]["wMm"], reverse=True):
        if r not in hits:
            fit = [b for b in free if near(p, b, 0.5) and side_ok(p, b)]    # centred on the pads; a connector may overhang them
            if fit:
                pc = (p["xMinMm"] + p["lMm"] / 2 - off[0], p["yMinMm"] + p["wMm"] / 2 - off[1])
                b = min(fit, key=lambda b: math.dist(centre(b), pc))
                hits[r] = (b, "position")
                free.remove(b)
    if report is not None:
        report.update(components=sum((b[3] - b[0]) * (b[4] - b[1]) < half for _, b in comps), unmatched=unmatched, away=away)
        for name, b in comps:
            if (b[3] - b[0]) * (b[4] - b[1]) >= half or any(b is h for h, _ in hits.values()):
                continue
            c = (centre(b)[0] + off[0], centre(b)[1] + off[1])
            if found.get(name) and len(found[name]) > 1:
                why = f"{len(found[name])} components are named {name}"
            else:
                inside = [r for r, p in edb.items() if near(p, b, 0.5)]
                why = (f"centred on the pads of {', '.join(inside)}, which already have a body" if inside else
                       f"no part named {name}, and no part's pads under its centre")
            unmatched.append({"name": name, "xMm": round(b[3] - b[0], 2), "yMm": round(b[4] - b[1], 2),
                              "hMm": round(b[5] - b[2], 2), "atMm": [round(c[0], 2), round(c[1], 2)], "why": why})
    # the face a part sits on: the commonest body bottom of the top-side parts; the bottom face is a board thickness lower
    tops = [round(b[2], 2) for r, (b, _) in hits.items() if edb[r]["side"] == "top"]
    top = max(set(tops), key=tops.count) if tops else max(b[2] for b, _ in hits.values()) if hits else 0.0
    bottom = top - board["thicknessMm"]
    if report is not None:      # STEP point + (offXMm, offYMm) = EDB point; topZMm is the board's top face in the STEP
        report["frame"] = {"offXMm": round(off[0], 4) + 0.0, "offYMm": round(off[1], 4) + 0.0, "topZMm": round(top, 4) + 0.0}
    out = {}
    for r, (b, by) in hits.items():
        h = b[5] - b[2]                     # a part off the board faces (a rigid-flex strip): its own height
        if edb[r]["side"] == "top" and -0.05 <= b[2] - top <= 0.5:
            h = b[5] - top                   # standoff included: the block stands on the board
        elif edb[r]["side"] == "bottom" and -0.05 <= bottom - b[5] <= 0.5:
            h = bottom - b[2]
        a = body_angle(b[6], edb[r].get("angleDeg") or 0.0)
        x, y = size_at(b[6], a)
        out[r] = {"hMm": round(h, 3), "xMm": round(x, 3), "yMm": round(y, 3), "aDeg": a, "by": by}
    return out


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        raise SystemExit(__doc__ + "\n    python ansys/step_parts.py <board.step> <board.json>   heights and bodies per RefDes")
    if len(sys.argv) == 3:
        rep = {}
        print(json.dumps(for_board(sys.argv[1], json.loads(Path(sys.argv[2]).read_text(encoding="utf-8")), rep), indent=1))
        print(f"{rep['components']} components in the STEP", file=sys.stderr)
        print(f"frame: {rep['frame']}", file=sys.stderr)
        for a in rep["away"]:
            print(f"away from its pads: {a['refDes']} body at {a['atMm']}, pads at {a['padsMm']} ({a['distMm']} mm)", file=sys.stderr)
        for u in rep["unmatched"]:
            print(f"not used: {u['name']} ({u['xMm']} x {u['yMm']} x {u['hMm']} mm at {u['atMm']}): {u['why']}", file=sys.stderr)
    else:
        print(json.dumps({k: [[round(x, 3) for x in b] for b in v] for k, v in boxes(sys.argv[1]).items()}, indent=1))
