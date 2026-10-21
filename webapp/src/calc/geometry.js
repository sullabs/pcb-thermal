// PCB size, PCB stack k and the body table (3 CAD geometry).
// Body order is the workbook's row order: imported bodies in import order, then bodies typed in by hand.

import { isNum, N } from "./util.js";

const BOX = ["xMin", "yMin", "zMin", "l", "w", "h"];

export function bodyTable(project, cad) {
  const rows = [];
  for (const b of cad?.import?.bodies || []) {
    const ch = project.bodies[b.key] || {};
    const row = { key: b.key, name: ch.nameOverride || b.name, importedName: b.name, side: b.side,
      refDes: ch.refDes, include: ch.include === true, box: usedBox(b.box, ch.override) };
    if (b.rotDeg) { row.rotDeg = b.rotDeg; row.ownL = b.ownL; row.ownW = b.ownW; }
    row.size = ownSize(row);
    rows.push(row);
  }
  for (const [key, ch] of Object.entries(project.bodies)) {
    if (!ch.manual) continue;
    const row = { key, name: ch.nameOverride || ch.manual.name, side: ch.manual.side, refDes: ch.refDes,
      include: ch.include === true, manual: true, box: usedBox(ch.manual.box, ch.override) };
    row.size = ownSize(row);
    rows.push(row);
  }
  return rows;
}

// Footprint L x W of a body. box.l and box.w are the axis-aligned envelope (position, edges, drawing); a body
// turned in the board plane (rotDeg, ownL, ownW from the import) has its own L x W inside it, and every area
// uses that. Same test as PT_ShapeIsRotated in the macros: the turned shape counts only while it still fits the
// box used, so an L or W override makes it a plain box again.
export function ownSize(b) {
  const x = b.box || {};
  const a = (b.rotDeg || 0) * Math.PI / 180, c = Math.abs(Math.cos(a)), s = Math.abs(Math.sin(a));
  const turned = Math.abs(b.rotDeg || 0) >= 0.05 && b.ownL > 0 && b.ownW > 0 && isNum(x.l) && isNum(x.w)
    && Math.abs(b.ownL * c + b.ownW * s - x.l) <= 0.02 && Math.abs(b.ownL * s + b.ownW * c - x.w) <= 0.02;
  return turned ? { l: b.ownL, w: b.ownW, turned } : { l: x.l, w: x.w, turned };
}

function usedBox(imported = {}, override = {}) {
  const b = {};
  for (const k of BOX) {
    const v = isNum(override[k]) ? override[k] : imported[k];
    if (isNum(v)) b[k] = v;
  }
  return b;
}
export const fullBox = b => BOX.every(k => isNum(b[k]));

// First body linked to this RefDes, included or not (the workbook's MATCH on the RefDes column).
export const bodyOf = (rows, refDes) => rows.find(r => r.refDes === refDes);

export function pcb(project, cad) {
  const imp = cad?.import?.pcb?.box || {};
  const ov = project.pcbOverride || {};
  const size = {};
  for (const k of ["l", "w", "h"]) {
    const v = isNum(ov[k]) ? ov[k] : imp[k];
    if (isNum(v)) size[k] = v;
  }
  return { ...size, ...stackK(project.pcbStack, size.h) };
}

// In plane: copper and dielectric in parallel. Through plane: layers in series, each copper layer a
// coverage weighted mix of copper and dielectric.
export function stackK(s, h) {
  const out = { copperTotalMm: s.copperLayers * s.copperThicknessMm };
  if (N(h) > 0) {
    out.copperFraction = (out.copperTotalMm * s.copperCoverage) / h;
    out.kxy = out.copperFraction * s.copperK + (1 - out.copperFraction) * s.dielectricK;
    out.kz = h / ((h - out.copperTotalMm) / s.dielectricK
      + out.copperTotalMm / (s.copperCoverage * s.copperK + (1 - s.copperCoverage) * s.dielectricK));
  }
  return out;
}

// Per body heat, density and check text; then the checks block.
export function geometryChecks(project, rows, board, reg) {
  const refs = new Set(project.parts.map(p => p.refDes));
  const bodies = rows.map(r => {
    const out = { key: r.key };
    const b = r.box;
    if (isNum(b.xMin) && isNum(b.l)) out.xCentre = b.xMin + b.l / 2;
    if (isNum(b.yMin) && isNum(b.w)) out.yCentre = b.yMin + b.w / 2;
    if (r.refDes) {
      out.pdW = N(reg.parts[r.refDes]?.pdW);
      const s = r.size;
      if (N(s.l) * N(s.w) > 0) out.densityWcm2 = out.pdW / ((s.l * s.w) / 100);
    }
    out.check = checkOf(r, rows, refs, board);
    return out;
  });
  const included = rows.filter(r => r.include && r.name);
  const placed = included.reduce((a, r) => a + (r.refDes ? N(reg.parts[r.refDes]?.pdW) : 0), 0);
  return {
    bodies,
    imported: rows.filter(r => !r.manual && r.importedName).length,
    included: included.length,
    registerPdW: reg.totalPdW,
    placedPdW: placed,
    notPlacedW: reg.totalPdW - placed,
    partsWithNoBody: project.parts.filter(p => !rows.some(r => r.include && r.refDes === p.refDes)).length,
    problems: bodies.filter(b => b.check && b.check !== "OK" && b.check !== "overhangs board edge").length,
  };
}

function checkOf(r, rows, refs, board) {
  if (!r.name) return "";
  const b = r.box;
  if (!fullBox(b)) return "missing geometry";
  if (r.refDes && !refs.has(r.refDes)) return "RefDes not in register";
  if (r.refDes && r.include && rows.filter(x => x.refDes === r.refDes && x.include).length > 1) return "RefDes used twice";
  if (r.refDes && !r.include) return "heat source not included";
  if (b.xMin < 0 || b.yMin < 0 || b.xMin + b.l > N(board.l) + 0.001 || b.yMin + b.w > N(board.w) + 0.001) return "overhangs board edge";
  return "OK";
}
