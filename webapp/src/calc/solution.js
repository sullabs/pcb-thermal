// Thermal solution hand check (8 Thermal solution, 8b Thermal solution bottom): one common plate over the
// parts marked "under plate". Stack per part: Rjc, TIM1, local spreader (optional), TIM2, spreading in the
// plate; the plate loses heat to air.

import { leeSpreading } from "./lee.js";
import { bodyOf } from "./geometry.js";
import { isNum, N, sum } from "./util.js";

export function thermalSolution(project, sol, rows, reg) {
  const ambient = project.board.ambientC;
  const tightBelow = project.settings.flagLimits.critical.marginBelowC;
  const inp = sol.parts || {};
  const o = { ambientC: ambient };
  if (isNum(sol.plateLMm) && isNum(sol.plateWMm)) o.plateAreaMm2 = sol.plateLMm * sol.plateWMm;
  const R = sol.rPlateToAirOverrideCW;
  if (isNum(R)) { if (R > 0 && isNum(o.plateAreaMm2)) o.hEff = 1 / ((R * o.plateAreaMm2) / 1e6); }
  else if (isNum(sol.hAir) && isNum(sol.faces)) o.hEff = sol.hAir * sol.faces;
  if (isNum(R)) o.rPlateAirCW = R;
  else if (isNum(o.hEff) && isNum(o.plateAreaMm2) && o.hEff * o.plateAreaMm2 > 0) o.rPlateAirCW = 1e6 / (o.hEff * o.plateAreaMm2);

  const pd = ref => reg.parts[ref].pdW;
  const active = p => inp[p.refDes]?.underPlate === true && isNum(pd(p.refDes));
  o.partsUnder = project.parts.filter(p => inp[p.refDes]?.underPlate === true).length;
  o.pdPlateW = sum(project.parts.filter(active).map(p => pd(p.refDes)));
  if (isNum(o.rPlateAirCW)) o.plateC = ambient + o.pdPlateW * o.rPlateAirCW;

  const parts = {};
  for (const p of project.parts) {
    const x = inp[p.refDes] || {};
    const r = {};
    const body = bodyOf(rows, p.refDes);
    if (body && isNum(body.box.l) && isNum(body.box.w)) r.packageMm2 = body.box.l * body.box.w;
    if (active(p)) {
      const q = pd(p.refDes);
      r.footprintMm2 = isNum(x.localLMm) ? (isNum(x.localWMm) ? x.localLMm * x.localWMm : undefined) : r.packageMm2;
      if (isNum(x.plateShareOverrideMm2)) r.plateShareMm2 = x.plateShareOverrideMm2;
      else if (isNum(o.plateAreaMm2) && o.pdPlateW > 0) r.plateShareMm2 = (o.plateAreaMm2 * q) / o.pdPlateW;
      r.rjcCW = isNum(p.rjcCW) ? p.rjcCW : undefined;
      if (isNum(x.tim1TMm) && N(x.tim1K) > 0 && N(r.packageMm2) > 0) r.rTim1CW = (1000 * x.tim1TMm) / (x.tim1K * r.packageMm2);
      const local = [x.localLMm, x.localWMm, x.localTMm, x.localK, x.tim2TMm, x.tim2K];
      if (local.every(v => !isNum(v))) { r.rLocalCW = 0; r.rTim2CW = 0; }
      else if (local.every(v => N(v) > 0) && N(r.packageMm2) > 0 && N(r.footprintMm2) > 0) {
        r.local = leeSpreading({ sourceMm2: r.packageMm2, plateMm2: r.footprintMm2, tMm: x.localTMm, k: x.localK, h: (1000 * x.tim2K) / x.tim2TMm });
        r.rLocalCW = r.local.r;
        r.rTim2CW = (1000 * x.tim2TMm) / (x.tim2K * r.footprintMm2);
      }
      if (N(r.footprintMm2) > 0 && N(r.plateShareMm2) > 0 && isNum(sol.plateTMm) && N(sol.plateK) > 0 && N(o.hEff) > 0) {
        r.plate = leeSpreading({ sourceMm2: r.footprintMm2, plateMm2: r.plateShareMm2, tMm: sol.plateTMm, k: sol.plateK, h: o.hEff });
        r.rPlateSpreadCW = r.plate.r;
      }
      const stack = [r.rjcCW, r.rTim1CW, r.rLocalCW, r.rTim2CW, r.rPlateSpreadCW];
      if (stack.every(isNum)) {
        r.rStackCW = sum(stack);
        r.dtStackC = q * r.rStackCW;
        if (isNum(o.plateC)) r.tjC = o.plateC + r.dtStackC;
        if (isNum(r.tjC) && isNum(p.tjMaxC)) r.marginC = p.tjMaxC - r.tjC;
      }
    }
    r.status = inp[p.refDes]?.underPlate !== true ? "not under plate" : !isNum(pd(p.refDes)) ? "no Pd in register"
      : !isNum(r.tjC) ? "missing input" : !isNum(r.marginC) ? "no Tj max"
      : r.marginC < 0 ? "FAIL" : r.marginC < tightBelow ? "Tight" : "OK";
    parts[p.refDes] = r;
  }
  o.footprintsMm2 = sum(Object.values(parts).map(r => r.footprintMm2));
  if (isNum(o.plateAreaMm2) && o.footprintsMm2 > o.plateAreaMm2) o.warning = "footprints exceed plate";
  o.parts = parts;
  return o;
}
