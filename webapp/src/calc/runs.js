// Simulation runs read back (9 Results) and whether the model and the runs are still up to date (design 4.2):
// per run and RefDes the margin to Tj max, a status, the hand check Tj it is compared with and the gap; per run
// the hottest part, PASS or FAIL, and what changed in the inputs since the run.
//
//   runResults(project, results, calc, fsNow) -> { [run name]: { parts, worst, pass, fails, balanceOk, changes } }
//   inputsHash(fs)                             -> short hash of the PCB and bodies the macros build from
//   inputChanges(fsThen, fsNow, simThen, simNow) -> texts, one per input that differs ("Pd of U5 0.8 W, now 1.1 W")

import { isNum } from "./util.js";

// Energy balance above this (%) means the run is not converged or a heat load is missing.
export const BALANCE_LIMIT_PCT = 2;

export function runResults(project, results, calc, fsNow) {
  const out = {};
  const tight = project.settings.flagLimits.critical.marginBelowC;
  const byRef = Object.fromEntries(project.parts.map(p => [p.refDes, p]));
  for (const run of results?.runs || []) {
    const withPlate = !!run.spreader || (run.spreaderFiles || []).length > 0;
    const parts = {};
    for (const [ref, t] of Object.entries(run.parts || {})) {
      const p = byRef[ref];
      const r = {};
      if (p && isNum(p.tjMaxC) && isNum(t.peakC)) r.marginC = p.tjMaxC - t.peakC;
      r.status = !p ? "not on tab 1" : !isNum(p.tjMaxC) ? "no Tj max" : r.marginC < 0 ? "FAIL" : r.marginC < tight ? "Tight" : "OK";
      const sol = withPlate ? ["top", "bottom"].map(s => calc.solution[s]?.parts?.[ref]).find(x => isNum(x?.tjC)) : undefined;
      if (sol) { r.handC = sol.tjC; r.handFrom = "thermal solution"; }
      else if (isNum(calc.bareBoard.parts[ref]?.tjC)) { r.handC = calc.bareBoard.parts[ref].tjC; r.handFrom = "bare board"; }
      if (isNum(r.handC) && isNum(t.peakC)) r.gapC = r.handC - t.peakC;
      parts[ref] = r;
    }
    const peaks = Object.entries(run.parts || {}).filter(([, t]) => isNum(t.peakC));
    const worst = peaks.reduce((a, x) => (!a || x[1].peakC > a[1].peakC ? x : a), null);
    const fails = Object.values(parts).filter(r => r.status === "FAIL").length;
    const checked = Object.values(parts).filter(r => isNum(r.marginC)).length;
    const o = { parts, fails, checked, withPlate };
    if (worst) { o.worst = worst[0]; o.worstC = worst[1].peakC; }
    o.pass = !checked ? "not checked" : fails ? "FAIL" : "PASS";
    if (isNum(run.energyBalancePct)) o.balanceOk = Math.abs(run.energyBalancePct) <= BALANCE_LIMIT_PCT;
    o.changes = run.inputs?.forSolidWorks ? inputChanges(run.inputs.forSolidWorks, fsNow, run.inputs.simSetup, project.simSetup) : undefined;
    out[run.name] = o;
  }
  return out;
}

// FNV-1a over the JSON text of what Build Board builds from: the PCB and the bodies (sizes rounded to 1 um,
// so a value that only differs in the last digit does not count). Build Board copies it into cad.json build.
export function inputsHash(fs) {
  if (!fs) return undefined;
  const r = x => (isNum(x) ? Math.round(x * 1000) / 1000 : x);
  const text = JSON.stringify({
    pcb: [fs.pcb.l, fs.pcb.w, fs.pcb.h].map(r),
    bodies: fs.bodies.map(b => [b.key, ...["xMin", "yMin", "zMin", "l", "w", "h"].map(f => r(b.box[f])), r(b.rotDeg || 0)]),
  });
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) { h ^= text.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
  return h.toString(16).padStart(8, "0");
}

const near = (a, b, tol = 1e-6) => isNum(a) && isNum(b) ? Math.abs(a - b) <= tol * Math.max(1, Math.abs(a), Math.abs(b)) : a === b;
const num = x => (isNum(x) ? String(+x.toPrecision(4)) : "blank");

export function inputChanges(then = {}, now = {}, simThen, simNow) {
  const out = [];
  if (!now) return ["the inputs cannot be worked out now (PCB size or k unknown)"];
  if (!near(then.ambientC, now.ambientC)) out.push(`ambient ${num(then.ambientC)} C, now ${num(now.ambientC)} C`);
  for (const f of ["l", "w", "h", "kxy", "kz"]) if (!near(then.pcb?.[f], now.pcb?.[f])) out.push(`PCB ${f === "kxy" ? "k xy" : f === "kz" ? "k z" : f.toUpperCase()} ${num(then.pcb?.[f])}, now ${num(now.pcb?.[f])}`);
  const refs = new Set([...Object.keys(then.heat || {}), ...Object.keys(now.heat || {})]);
  for (const r of refs) if (!near(then.heat?.[r], now.heat?.[r])) out.push(`Pd of ${r} ${num(then.heat?.[r])} W, now ${num(now.heat?.[r])} W`);
  const tb = Object.fromEntries((then.bodies || []).map(b => [b.key, b])), nb = Object.fromEntries((now.bodies || []).map(b => [b.key, b]));
  let moved = 0, added = 0, gone = 0;
  for (const k of Object.keys(nb)) { if (!tb[k]) added++; else if (["xMin", "yMin", "zMin", "l", "w", "h"].some(f => !near(tb[k].box?.[f], nb[k].box?.[f], 1e-4))) moved++; }
  for (const k of Object.keys(tb)) if (!nb[k]) gone++;
  if (moved) out.push(`${moved} bod${moved === 1 ? "y" : "ies"} moved or resized`);
  if (added) out.push(`${added} bod${added === 1 ? "y" : "ies"} added to the model`);
  if (gone) out.push(`${gone} bod${gone === 1 ? "y" : "ies"} left out of the model`);
  const tc = Object.fromEntries((then.chips || []).map(c => [c.refDes, c])), nc = Object.fromEntries((now.chips || []).map(c => [c.refDes, c]));
  for (const r of new Set([...Object.keys(tc), ...Object.keys(nc)])) {
    if (tc[r]?.model !== nc[r]?.model) out.push(`chip model of ${r} ${tc[r]?.model || "none"}, now ${nc[r]?.model || "none"}`);
    else if (!near(tc[r]?.kz, nc[r]?.kz) || !near(tc[r]?.kxy, nc[r]?.kxy)) out.push(`k of chip ${r} changed`);
  }
  for (const s of ["top", "bottom"]) {
    const a = then.plates?.[s], b = now.plates?.[s];
    if (!a && !b) continue;
    if (!a || !b) { out.push(`${s} plate ${a ? "removed" : "added"}`); continue; }
    if (!near(a.faceZMm, b.faceZMm) || !near(a.thicknessMm, b.thicknessMm) || !near(a.k, b.k)) out.push(`${s} plate height, thickness or k changed`);
    if (JSON.stringify(a.chips) !== JSON.stringify(b.chips)) out.push(`${s} plate chips, TIMs or pedestals changed`);
  }
  if (simThen && simNow) {
    for (const [g, name] of [["spreaderTop", "spreader top"], ["spreaderOther", "spreader other faces"], ["pcbTop", "PCB top"], ["pcbBottom", "PCB bottom"], ["parts", "parts"]]) {
      const a = simThen.surfaces?.[g], b = simNow.surfaces?.[g];
      if (!near(a?.h, b?.h) || !near(a?.emissivity, b?.emissivity)) out.push(`h or emissivity of ${name} changed`);
    }
    for (const [f, name] of [["partToPcb", "part to PCB"], ["chipToTim", "chip to TIM"], ["timToSpreader", "TIM to spreader"]]) {
      if (!near(simThen.contacts?.[f], simNow.contacts?.[f])) out.push(`contact R ${name} changed`);
    }
    if (!near(simThen.meshGlobalMm, simNow.meshGlobalMm) || !near(simThen.meshFineMm, simNow.meshFineMm)) out.push("mesh size changed");
    if (simThen.includePartsWithoutHeat !== simNow.includePartsWithoutHeat || !near(simThen.kPartsWithoutHeat, simNow.kPartsWithoutHeat)) out.push("parts without heat changed");
  }
  return out;
}
