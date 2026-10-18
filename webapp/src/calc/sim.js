// Per chip values for the simulation (9 Sim setup): the block model gets a k z that reproduces Rjc
// through its height: k z = H / (Rjc x A).

import { bodyOf } from "./geometry.js";
import { isNum, N } from "./util.js";

export function simChips(project, rows, reg) {
  const byRef = Object.fromEntries(project.parts.map(p => [p.refDes, p]));
  return (project.simSetup?.chips || []).map(c => {
    const p = byRef[c.refDes] || {};
    const r = { refDes: c.refDes, model: c.model };
    if (N(p.rjcCW) > 0) r.rjcCW = p.rjcCW;
    if (N(p.rjbCW) > 0) r.rjbCW = p.rjbCW;
    const body = bodyOf(rows, c.refDes);
    if (body) {
      r.hMm = body.box.h;
      if (isNum(body.box.l) && isNum(body.box.w)) r.areaMm2 = body.box.l * body.box.w;
    }
    if (N(r.rjcCW) > 0 && N(r.hMm) > 0 && N(r.areaMm2) > 0) r.kz = (r.hMm * 1000) / (r.rjcCW * r.areaMm2);
    r.kxy = isNum(c.kxy) ? c.kxy : r.kz;
    r.pdW = N(reg.parts[c.refDes]?.pdW);
    r.check = !body ? "RefDes not on CAD geometry" : !(r.pdW > 0) ? "no heat in the register"
      : c.model === "Two-resistor" ? "two-resistor model comes later (M8)"
      : !isNum(r.kz) ? "no Rjc: k of parts without heat is used" : "OK";
    return r;
  });
}
