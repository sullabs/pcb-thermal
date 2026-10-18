// Spreader plates (7 Spreader, 7b Spreader bottom): plate face height, per chip gap, TIM and pedestal,
// and the clearance cuts for tall parts the plate does not cover. Heights are from the PCB bottom face (mm).
// Top plate: its face is the underside, above the parts. Bottom plate: its face is the top side, below the
// bottom parts (negative Z).

import { bodyOf } from "./geometry.js";
import { isNum, N } from "./util.js";

export function spreaderPlate(plate, side, rows) {
  const top = side === "top";
  const chips = plate.chips.map(c => {
    const r = { refDes: c.refDes, covered: c.covered === true };
    const body = bodyOf(rows, c.refDes);
    if (body) {
      const b = body.box;
      r.key = body.key;
      Object.assign(r, { xMin: b.xMin, yMin: b.yMin, l: b.l, w: b.w });
      // Chip face toward the plate: top of the body for the top plate, its bottom for the bottom plate.
      if (top ? isNum(b.zMin) && isNum(b.h) : isNum(b.zMin)) r.faceZ = top ? b.zMin + b.h : b.zMin;
    }
    if (r.covered && isNum(r.faceZ)) r.zNeed = top ? r.faceZ + N(c.timTMm) : r.faceZ - N(c.timTMm);
    else if (top) r.zNeed = 0;
    return { r, c, body };
  });

  let faceZMm;
  if (plate.heightMode === "Manual") faceZMm = N(plate.heightManualMm);
  else if (top) faceZMm = Math.max(0, ...chips.map(x => N(x.r.zNeed)));
  else { const z = chips.map(x => x.r.zNeed).filter(isNum); faceZMm = z.length ? Math.min(...z) : 0; }

  for (const { r, c, body } of chips) {
    if (r.covered && isNum(r.faceZ)) {
      r.gapMm = top ? faceZMm - r.faceZ : r.faceZ - faceZMm;
      r.timTMm = c.heightMode === "Thick TIM" ? r.gapMm : r.gapMm - N(c.timTMm) < plate.minPedestalMm ? r.gapMm : N(c.timTMm);
      r.pedestalMm = r.gapMm - r.timTMm;
      if (N(c.timK) > 0 && N(r.l) * N(r.w) > 0) r.timRCW = (r.timTMm * 1000) / (c.timK * r.l * r.w);
    }
    r.check = chipCheck(r, c, body, top);
  }

  const coveredKeys = new Set(chips.filter(x => x.r.covered && x.r.key).map(x => x.r.key));
  const cx = plate.clearanceXYMm, cz = plate.clearanceZMm;
  const cuts = [];
  for (const b of rows) {
    const z = b.box;
    if (!b.include || !isNum(z.zMin) || !isNum(z.h) || coveredKeys.has(b.key)) continue;
    const reaches = top ? z.zMin + z.h > faceZMm - cz : z.zMin < faceZMm + cz;
    if (!reaches) continue;
    cuts.push({ key: b.key, name: b.name, xMin: z.xMin - cx, yMin: z.yMin - cx, l: z.l + 2 * cx, w: z.w + 2 * cx,
      edgeZMm: top ? z.zMin + z.h : z.zMin });
  }
  return { faceZMm, covered: chips.filter(x => x.r.covered).length, chips: chips.map(x => x.r), cuts };
}

function chipCheck(r, c, body, top) {
  if (!body) return "RefDes not on CAD geometry";
  if (!body.include) return "body not included";
  if (!r.covered) return "not covered";
  if (top && body.side === "Bottom") return "chip is on the bottom side: cover it on 7b Spreader bottom";
  if (!top && body.side !== "Bottom") return "chip is on the top side: cover it on 7 Spreader";
  if (!isNum(r.gapMm)) return "missing geometry";
  if (r.gapMm < -0.0005) return top ? "chip is above the underside" : "chip is below the plate face";
  if (c.timMode === "Body" && r.timTMm < 0.01) return "TIM too thin for a body: use Contact R";
  if (c.timMode === "Contact R" && r.timTMm > 0.3) return "big gap for a contact resistance: use Body";
  return "OK";
}

// What Build Spreader needs for one plate (forSolidWorks.plates.top / bottom).
export function plateForSolidWorks(plate, calc) {
  const round9 = x => Math.round(x * 1e9) / 1e9;
  const chips = [];
  plate.chips.forEach((c, i) => {
    const r = calc.chips[i];
    if (!r.covered || !r.key) return;
    const o = { refDes: c.refDes, key: r.key, timTMm: round9(N(r.timTMm)), pedestalMm: Math.max(0, round9(N(r.pedestalMm))),
      timMode: c.timMode === "Contact R" ? "Contact R" : "Body" };
    if (N(c.timK) > 0) o.timK = c.timK;
    if (N(r.timRCW) > 0) o.timRCW = r.timRCW;
    chips.push(o);
  });
  return {
    faceZMm: calc.faceZMm, material: plate.material, k: plate.k, densityKgM3: plate.densityKgM3,
    thicknessMm: plate.thicknessMm, clearanceXYMm: plate.clearanceXYMm,
    chips, cuts: calc.cuts.filter(x => [x.xMin, x.yMin, x.l, x.w, x.edgeZMm].every(isNum) && x.l > 0 && x.w > 0),
  };
}

