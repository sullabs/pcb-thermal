// Calculation module: every number the page shows and the macros read, from project.json, cad.json and
// results.json. Pure functions, no file or page access. Formulas are the template's (checked against the
// RPi4 workbooks by tools/test_w2.mjs). Formula texts for the page come with the term catalogue (W4).
//
//   calcProject(project, cad, results) -> { register, triage, rails, pcb, geometry, bareBoard,
//                                           spreader: {top, bottom}, solution: {top, bottom}, simChips }
//   forSolidWorks(project, calc, now)  -> the forSolidWorks block of project.json, or undefined when the
//                                         PCB size or k is not known yet
//   calc.fsNow: forSolidWorks as Save would write it now; calc.runs: results per run (runs.js); calc.build: whether
//   the model in SolidWorks was built from the PCB and bodies of now (cad.json build.inputsHash)

import { register, triage, railBudget } from "./power.js";
import { bodyTable, bodyOf, pcb, geometryChecks, fullBox } from "./geometry.js";
import { bareBoard } from "./network.js";
import { spreaderPlate, plateForSolidWorks } from "./spreader.js";
import { thermalSolution } from "./solution.js";
import { simChips } from "./sim.js";
import { runResults, inputsHash } from "./runs.js";
import { flowCheck, flowForSw } from "./flow.js";
import { isNum, put } from "./util.js";

export const CALC_VERSION = "0.1.0";

export function calcProject(project, cad = {}, results = {}) {
  const rows = bodyTable(project, cad);
  const board = pcb(project, cad);
  const areaOf = ref => {
    const b = bodyOf(rows, ref)?.box;
    return b && isNum(b.l) && isNum(b.w) && b.l * b.w > 0 ? b.l * b.w : undefined;
  };
  const reg = register(project, areaOf);
  const out = {
    bodies: rows,
    register: reg,
    triage: triage(project, reg),
    rails: railBudget(project, reg),
    pcb: board,
    geometry: geometryChecks(project, rows, board, reg),
    bareBoard: bareBoard(project, results, rows, board, reg),
    spreader: {},
    solution: {},
    simChips: simChips(project, rows, reg),
  };
  for (const side of ["top", "bottom"]) {
    const pl = project.spreader?.[side];
    if (pl) out.spreader[side] = spreaderPlate(pl, side, rows);
    const s = project.handCheck?.[side === "top" ? "solutionTop" : "solutionBottom"];
    if (s) out.solution[side] = thermalSolution(project, s, rows, reg);
  }
  out.fsNow = forSolidWorks(project, out, "");
  out.flow = flowCheck(project, out);
  out.runs = runResults(project, results, out, out.fsNow);
  const built = cad?.build?.inputsHash;
  out.build = { hashNow: out.fsNow?.inputsHash, hashBuilt: built };
  if (cad?.build && built && out.fsNow) out.build.outOfDate = built !== out.fsNow.inputsHash;
  return out;
}

export function forSolidWorks(project, calc, now) {
  const p = calc.pcb;
  if (![p.l, p.w, p.h, p.kxy, p.kz].every(x => isNum(x) && x > 0)) return undefined;
  const bodies = calc.bodies.filter(b => b.include && fullBox(b.box)).map(b => {
    const o = put({ key: b.key, name: b.name }, "refDes", b.refDes);
    o.box = { xMin: b.box.xMin, yMin: b.box.yMin, zMin: b.box.zMin, l: b.box.l, w: b.box.w, h: b.box.h };
    if (b.rotDeg) { o.rotDeg = b.rotDeg; put(o, "ownL", b.ownL); put(o, "ownW", b.ownW); }
    return o;
  });
  const heat = {};
  for (const x of project.parts) {
    const pd = calc.register.parts[x.refDes].pdW;
    if (isNum(pd)) heat[x.refDes] = pd;
  }
  const chips = calc.simChips.map(c => {
    const o = { refDes: c.refDes, model: c.model };
    if (isNum(c.kz) && c.kz > 0) o.kz = c.kz;
    if (isNum(c.kxy) && c.kxy > 0) o.kxy = c.kxy;
    put(o, "rjcCW", c.rjcCW);
    put(o, "rjbCW", c.rjbCW);
    return o;
  });
  const fs = {
    computedAt: now, computedBy: "calc " + CALC_VERSION, ambientC: project.board.ambientC,
    pcb: { l: p.l, w: p.w, h: p.h, kxy: p.kxy, kz: p.kz }, bodies, heat, chips,
  };
  const plates = {};
  for (const side of ["top", "bottom"]) {
    const pl = project.spreader?.[side];
    if (pl?.enabled && calc.spreader[side]) plates[side] = plateForSolidWorks(pl, calc.spreader[side]);
  }
  if (Object.keys(plates).length) fs.plates = plates;
  const flow = flowForSw(project, calc, fs);
  if (Object.keys(flow.chips).length) fs.flow = flow;
  fs.inputsHash = inputsHash(fs);
  return fs;
}
