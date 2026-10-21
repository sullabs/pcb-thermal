// Compare and validate (Flow plan F7): per RefDes the hand check, the newest healthy Thermal run, the newest healthy
// Flow run and the newest measurement, with the gaps; and, when Flow and the measurement are far apart, what usually
// explains it, ranked from the inputs.
//
//   compareTable(project, results, calc, anyHealth) -> { thermal, flow, meas, refs, rows: { [ref]: {...} } }
//   explainGap(project, calc, run, ref, flowC, measC, method) -> [text], most likely first

import { isNum } from "./util.js";

// Flow and a measurement further apart than this (C) get a list of likely causes.
export const MEAS_GAP_LIMIT_C = 5;
export const MEAS_METHODS = ["thermocouple", "IR", "internal sensor"];

const newest = (list, f) => list.filter(f).reduce((a, r) => (!a || (r.date || "") >= (a.date || "") ? r : a), null);

function handTj(calc, ref) {
  const sol = ["top", "bottom"].map(s => calc.solution[s]?.parts?.[ref]).find(x => isNum(x?.tjC));
  if (sol) return { c: sol.tjC, from: "plate" };
  if (isNum(calc.bareBoard.parts[ref]?.tjC)) return { c: calc.bareBoard.parts[ref].tjC, from: "bare board" };
  return undefined;
}

export function compareTable(project, results, calc, anyHealth = false) {
  const runs = (results?.runs || []).filter(r => r.status === "read");
  const okRun = r => { const h = calc.runs?.[r.name]?.health?.level; return h === "ok" || (anyHealth && h === "check"); };
  const thermal = newest(runs, r => r.studyType === "Thermal" && okRun(r));
  const flow = newest(runs, r => r.studyType === "Flow" && okRun(r));
  const meas = newest(project.measurements || [], () => true);
  const refs = project.parts.filter(x => calc.register.parts[x.refDes]?.pdW > 0).map(x => x.refDes);
  const rows = {};
  for (const ref of refs) {
    const r = {};
    const h = handTj(calc, ref);
    if (h) { r.handC = h.c; r.handFrom = h.from; }
    if (isNum(thermal?.parts?.[ref]?.peakC)) r.thermalC = thermal.parts[ref].peakC;
    if (isNum(flow?.parts?.[ref]?.peakC)) r.flowC = flow.parts[ref].peakC;
    const m = meas?.parts?.[ref];
    if (isNum(m?.tC)) { r.measC = m.tC; r.method = m.method; }
    const base = isNum(r.measC) ? r.measC : undefined;
    if (isNum(base)) for (const [k, g] of [["handC", "gapHand"], ["thermalC", "gapThermal"], ["flowC", "gapFlow"]]) if (isNum(r[k])) r[g] = r[k] - base;
    if (isNum(r.flowC) && isNum(r.thermalC)) r.flowVsThermal = r.flowC - r.thermalC;
    if (isNum(r.gapFlow) && Math.abs(r.gapFlow) > MEAS_GAP_LIMIT_C) r.causes = explainGap(project, calc, flow, ref, r.flowC, r.measC, r.method);
    rows[ref] = r;
  }
  return { thermal: thermal?.name, flow: flow?.name, meas: meas?.name, measAmbientC: meas?.ambientC, refs, rows };
}

// Likely causes of a gap between Flow and a measurement, most likely first. Flow cooler than measured: heat missing,
// contacts set to 0, emissivity too high, no Rjc (package temperature). Flow warmer: emissivity too low, a hotter
// ambient than the test, the measurement on the case and not the junction.
export function explainGap(project, calc, run, ref, flowC, measC, method) {
  const out = [];
  const ss = run?.inputs?.simSetup || project.simSetup || {};
  const fsw = run?.inputs?.forSolidWorks || calc.fsNow || {};
  const ct = ss.contacts || {};
  const sf = ss.surfaces || {};
  const eps = ["pcbTop", "parts", "spreaderTop"].map(g => sf[g]?.emissivity).filter(isNum);
  const meas = newest(project.measurements || [], () => true);
  if (flowC < measC) {
    if (isNum(run?.heatRequestedW) && isNum(run?.heatAppliedW) && run.heatAppliedW < run.heatRequestedW - 0.01) out.push(`heat missing in the Flow run: ${(run.heatRequestedW - run.heatAppliedW).toFixed(2)} W of ${run.heatRequestedW.toFixed(2)} W was not applied`);
    if ((calc.heatNoBody || []).length) out.push(`parts with heat and no body: ${calc.heatNoBody.map(x => x.refDes).join(", ")}`);
    if (!(ct.partToPcb > 0) && !(ct.chipToTim > 0) && !(ct.timToSpreader > 0)) out.push("every contact resistance is 0 (bonded): real solder joints and TIM interfaces add resistance, so the real chip runs hotter (tab 8, Contacts)");
    if (run?.parts?.[ref]?.tjBasis === "package" || (fsw.flow?.chips?.[ref]?.basis === "package")) out.push(`${ref} has no Rjc: Flow gives its package temperature, which is cooler than Tj${method === "internal sensor" ? " (and an internal sensor reads Tj)" : ""}`);
    if (eps.some(e => e >= 0.8)) out.push(`emissivity ${Math.max(...eps)} on tab 5 or 8: bare metal (0.05 to 0.1) or a glossy surface radiates far less than that`);
    if (isNum(meas?.ambientC) && isNum(fsw.ambientC) && meas.ambientC > fsw.ambientC + 1) out.push(`the test ambient (${meas.ambientC} C) was warmer than the study ambient (${fsw.ambientC} C)`);
  } else {
    if (eps.some(e => e < 0.5)) out.push(`emissivity ${Math.min(...eps)} on tab 5 or 8 is low: a painted, anodised or solder mask surface is 0.8 to 0.95`);
    if (isNum(meas?.ambientC) && isNum(fsw.ambientC) && meas.ambientC < fsw.ambientC - 1) out.push(`the test ambient (${meas.ambientC} C) was cooler than the study ambient (${fsw.ambientC} C)`);
    if (method === "thermocouple" || method === "IR") out.push(`the ${method} reads the case surface, cooler than Tj; Flow's peak of a calibrated block is Tj`);
    if (method === "IR") out.push("an IR camera reads low on a glossy chip unless its emissivity is set or the chip carries a dab of matt paint or tape");
    out.push("the real board may lose more heat than the model: an enclosure opening, a cable, a metal case or a mounting post");
  }
  if (method === "thermocouple") out.push("a thermocouple bead and its wire can act as a small heat sink on a small chip");
  return out;
}
