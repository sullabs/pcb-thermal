// Every value the page shows, by term id and key, and the formula of each calculated one, with its inputs.
// The page uses this for "Show formulas", the Explain panel and the trace from a value back to what was typed.
// tools/test_w4.mjs checks that each formula, worked out from its inputs, gives the value the calculation
// module gives, for every instance in both RPi4 examples. So the text the user reads is the sum that was done.
//
//   ctx = { project, cad, results, calc }       calc = calcProject(project, cad, results)
//   valueOf(ctx, term, key)  -> the value (number, text, true/false) or undefined
//   formulaOf(ctx, term, key) -> { text, expr, inputs: [[symbol, term, key]], fn, note, ref } or null
//   instances(ctx)          -> every [term, key] the project has, for the trace ("used by") and the tests
//
// Keys: part.* by RefDes, load.* by load id, body.* by body key, bbp.* (bare board per part) by RefDes,
// rail.* by rail name, plate.* by "top" / "bottom", pchip.* by "top/0" (side and chip index), sim.* by
// chip index, run.* by run name, runp.* by "run name|RefDes", change.* by change id. Others have key "".

import { CLASS_METHOD, STEFAN_BOLTZMANN } from "./constants.js";
import { bodyOf } from "./geometry.js";
import { isNum, N } from "./util.js";

const partOf = (c, ref) => c.project.parts.find(p => p.refDes === ref);
const loadOf = (c, id) => {
  for (const p of c.project.parts) for (const l of p.loads) if (l.id === id) return { part: p, load: l, r: c.calc.register.loads[id] };
  return null;
};
const bodyRow = (c, key) => c.calc.bodies.find(b => b.key === key);
const bodyOfRef = (c, ref) => bodyOf(c.calc.bodies, ref);
const bbp = (c, ref) => c.calc.bareBoard.parts[ref];
const surf = c => c.project.simSetup?.surfaces || {};
const rail = (c, name) => c.calc.rails.rails.find(r => r.rail === name);
const plateIn = (c, side) => c.project.spreader?.[side];
const pchip = (c, key) => { const [side, i] = key.split("/"); return { side, i: +i, c: plateIn(c, side)?.chips[+i], r: c.calc.spreader[side]?.chips[+i], pl: plateIn(c, side), calc: c.calc.spreader[side] }; };
const runOf = (c, name) => (c.results?.runs || []).find(r => r.name === name);
const runp = (c, key) => { const [run, ref] = key.split("|"); return runOf(c, run)?.parts?.[ref]; };
const change = (c, id) => (c.cad?.changes || []).find(x => x.id === id);
const lim = c => c.project.settings.flagLimits;
const geo = c => c.calc.geometry;
const geoBody = (c, key) => c.calc.geometry.bodies.find(b => b.key === key);
const choice = (c, key) => c.project.bodies[key] || {};
const bodySet = (c, key, f) => { const ch = choice(c, key); return ch.manual ? ch.manual.box?.[f] : ch.override?.[f]; };

// ---------- values ----------
export const VALUES = {
  "const.sigma": () => STEFAN_BOLTZMANN,
  "board.name": c => c.project.board.name,
  "board.stepFile": c => c.project.board.stepFile,
  "board.ambient": c => c.project.board.ambientC,
  "board.measuredInputPower": c => c.project.board.measuredInputPowerW,
  "board.notes": c => c.project.board.notes,
  "board.totalPd": c => c.calc.register.totalPdW,
  "limit.criticalMargin": c => lim(c).critical.marginBelowC,
  "limit.criticalDensity": c => lim(c).critical.densityAboveWcm2,
  "limit.criticalShare": c => lim(c).critical.shareAbove,
  "limit.watchMargin": c => lim(c).watch.marginBelowC,
  "limit.watchDensity": c => lim(c).watch.densityAboveWcm2,
  "limit.watchShare": c => lim(c).watch.shareAbove,

  "part.refDes": (c, k) => partOf(c, k)?.refDes,
  "part.partNumber": (c, k) => partOf(c, k)?.partNumber,
  "part.class": (c, k) => partOf(c, k)?.class,
  "part.package": (c, k) => partOf(c, k)?.package,
  "part.tjMax": (c, k) => partOf(c, k)?.tjMaxC,
  "part.rja": (c, k) => partOf(c, k)?.rjaCW,
  "part.rjc": (c, k) => partOf(c, k)?.rjcCW,
  "part.rjb": (c, k) => partOf(c, k)?.rjbCW,
  "part.taOverride": (c, k) => partOf(c, k)?.taOverrideC,
  "part.notes": (c, k) => partOf(c, k)?.notes,
  "part.pd": (c, k) => c.calc.register.parts[k]?.pdW,
  "part.ta": (c, k) => c.calc.register.parts[k]?.taC,
  "part.area": (c, k) => c.calc.register.parts[k]?.areaMm2,
  "part.rise": (c, k) => c.calc.register.parts[k]?.riseC,
  "part.tjEst": (c, k) => c.calc.register.parts[k]?.tjEstC,
  "part.margin": (c, k) => c.calc.register.parts[k]?.marginC,
  "part.density": (c, k) => c.calc.register.parts[k]?.densityWcm2,
  "part.share": (c, k) => c.calc.register.parts[k]?.share,
  "part.flag": (c, k) => c.calc.register.parts[k]?.flag,
  "part.why": (c, k) => c.calc.register.parts[k]?.why,
  "part.rankPd": (c, k) => c.calc.register.parts[k]?.rankPd,
  "part.rankDensity": (c, k) => c.calc.register.parts[k]?.rankDensity,

  "load.label": (c, k) => loadOf(c, k)?.load.label,
  "load.rail": (c, k) => loadOf(c, k)?.load.rail,
  "load.method": (c, k) => loadOf(c, k)?.load.method,
  "load.vin": (c, k) => loadOf(c, k)?.load.vinV,
  "load.iin": (c, k) => loadOf(c, k)?.load.iinA,
  "load.vout": (c, k) => loadOf(c, k)?.load.voutV,
  "load.iout": (c, k) => loadOf(c, k)?.load.ioutA,
  "load.eff": (c, k) => loadOf(c, k)?.load.efficiency,
  "load.r": (c, k) => loadOf(c, k)?.load.rOhm,
  "load.pdMeas": (c, k) => loadOf(c, k)?.load.pdMeasuredW,
  "load.methodUsed": (c, k) => loadOf(c, k)?.r.method,
  "load.pin": (c, k) => loadOf(c, k)?.r.pinW,
  "load.pout": (c, k) => loadOf(c, k)?.r.poutW,
  "load.pdCalc": (c, k) => loadOf(c, k)?.r.pdCalcW,
  "load.pd": (c, k) => loadOf(c, k)?.r.pdW,
  "load.problem": (c, k) => loadOf(c, k)?.r.problem,

  "rail.name": (c, k) => rail(c, k)?.rail,
  "rail.parts": (c, k) => rail(c, k)?.parts,
  "rail.pd": (c, k) => rail(c, k)?.pdW,
  "rail.share": (c, k) => rail(c, k)?.share,
  "rails.sum": c => c.calc.rails.sumRailsW,
  "rails.noRail": c => c.calc.rails.noRailW,
  "rails.measured": c => c.calc.rails.measuredW,
  "rails.unaccounted": c => c.calc.rails.unaccountedW,
  "rails.unaccountedFrac": c => c.calc.rails.unaccountedFrac,

  "triage.parts": c => c.calc.triage.parts,
  "triage.critical": c => c.calc.triage.critical,
  "triage.watch": c => c.calc.triage.watch,
  "triage.minor": c => c.calc.triage.minor,
  "triage.noPd": c => c.calc.triage.noPd,
  "triage.noRja": c => c.calc.triage.noRja,
  "triage.noBody": c => c.calc.triage.noBody,

  "import.model": c => c.cad?.import?.model,
  "import.date": c => c.cad?.import?.date,
  "import.bodies": c => c.cad?.import?.bodies?.length,
  "import.frame": c => c.cad?.import?.frame,
  "set.pcbNameContains": c => c.project.settings.import?.pcbNameContains,
  "set.minSize": c => c.project.settings.import?.minSizeMm,
  "set.topLevelOnly": c => c.project.settings.import?.topLevelOnly,
  "set.tolPos": c => c.project.settings.tolPosMm,
  "set.tolSize": c => c.project.settings.tolSizeMm,
  "pcbov.l": c => c.project.pcbOverride?.l,
  "pcbov.w": c => c.project.pcbOverride?.w,
  "pcbov.h": c => c.project.pcbOverride?.h,
  "pcb.l": c => c.calc.pcb.l,
  "pcb.w": c => c.calc.pcb.w,
  "pcb.h": c => c.calc.pcb.h,
  "stack.layers": c => c.project.pcbStack.copperLayers,
  "stack.cuT": c => c.project.pcbStack.copperThicknessMm,
  "stack.coverage": c => c.project.pcbStack.copperCoverage,
  "stack.cuK": c => c.project.pcbStack.copperK,
  "stack.dielK": c => c.project.pcbStack.dielectricK,
  "pcb.cuTotal": c => c.calc.pcb.copperTotalMm,
  "pcb.cuFraction": c => c.calc.pcb.copperFraction,
  "pcb.kxy": c => c.calc.pcb.kxy,
  "pcb.kz": c => c.calc.pcb.kz,

  "body.include": (c, k) => bodyRow(c, k)?.include,
  "body.name": (c, k) => bodyRow(c, k)?.name,
  "body.side": (c, k) => bodyRow(c, k)?.side,
  "body.refDes": (c, k) => bodyRow(c, k)?.refDes,
  "body.l": (c, k) => bodyRow(c, k)?.box.l,
  "body.w": (c, k) => bodyRow(c, k)?.box.w,
  "body.h": (c, k) => bodyRow(c, k)?.box.h,
  "body.xMin": (c, k) => bodyRow(c, k)?.box.xMin,
  "body.yMin": (c, k) => bodyRow(c, k)?.box.yMin,
  "body.zMin": (c, k) => bodyRow(c, k)?.box.zMin,
  "body.pd": (c, k) => geoBody(c, k)?.pdW,
  "body.density": (c, k) => geoBody(c, k)?.densityWcm2,
  "body.check": (c, k) => geoBody(c, k)?.check,
  "bodyset.name": (c, k) => choice(c, k).manual?.name ?? choice(c, k).nameOverride,
  "bodyset.side": (c, k) => choice(c, k).manual?.side,
  "bodyset.l": (c, k) => bodySet(c, k, "l"),
  "bodyset.w": (c, k) => bodySet(c, k, "w"),
  "bodyset.h": (c, k) => bodySet(c, k, "h"),
  "bodyset.xMin": (c, k) => bodySet(c, k, "xMin"),
  "bodyset.yMin": (c, k) => bodySet(c, k, "yMin"),
  "bodyset.zMin": (c, k) => bodySet(c, k, "zMin"),
  "geo.included": c => geo(c).included,
  "geo.placed": c => geo(c).placedPdW,
  "geo.notPlaced": c => geo(c).notPlacedW,
  "geo.noBody": c => geo(c).partsWithNoBody,
  "geo.problems": c => geo(c).problems,

  "surf.pcbTopH": c => surf(c).pcbTop?.h,
  "surf.pcbTopEps": c => surf(c).pcbTop?.emissivity,
  "surf.pcbBottomH": c => surf(c).pcbBottom?.h,
  "surf.pcbBottomEps": c => surf(c).pcbBottom?.emissivity,
  "surf.partsH": c => surf(c).parts?.h,
  "surf.partsEps": c => surf(c).parts?.emissivity,
  "hand.surfaceRise": c => c.project.handCheck?.surfaceRiseK,
  "hand.compareRun": c => c.project.handCheck?.compareRun,
  "simset.kPartsWithoutHeat": c => c.project.simSetup?.kPartsWithoutHeat,

  "bb.area": c => c.calc.bareBoard.board.areaMm2,
  "bb.hRadTop": c => c.calc.bareBoard.board.hRadTop,
  "bb.hRadBottom": c => c.calc.bareBoard.board.hRadBottom,
  "bb.hRadParts": c => c.calc.bareBoard.board.hRadParts,
  "bb.hPcb": c => c.calc.bareBoard.board.hPcb,
  "bb.hParts": c => c.calc.bareBoard.board.hParts,
  "bb.rBoardAir": c => c.calc.bareBoard.board.rBoardAirCW,
  "bb.totalPd": c => c.calc.bareBoard.board.totalPdW,
  "bb.rise": c => c.calc.bareBoard.board.riseC,
  "bb.boardC": c => c.calc.bareBoard.board.boardC,
  "bb.qBoard": c => c.calc.bareBoard.board.qBoardW,
  "bb.boardShare": c => (isNum(c.calc.bareBoard.board.qBoardW) && c.calc.bareBoard.board.totalPdW > 0 ? c.calc.bareBoard.board.qBoardW / c.calc.bareBoard.board.totalPdW : undefined),
  "bb.fail": c => c.calc.bareBoard.board.fail,
  "bb.tight": c => c.calc.bareBoard.board.tight,
  "bb.compareRun": c => c.calc.bareBoard.compare.run,
  "bb.compareCheck": c => c.calc.bareBoard.compare.check,
  "bb.meanGap": c => c.calc.bareBoard.compare.meanGapC,
  "bb.maxGap": c => c.calc.bareBoard.compare.maxGapC,
  "bbp.pd": (c, k) => bbp(c, k)?.pdW,
  "bbp.l": (c, k) => bbp(c, k)?.lMm,
  "bbp.w": (c, k) => bbp(c, k)?.wMm,
  "bbp.h": (c, k) => bbp(c, k)?.hMm,
  "bbp.rjcUsed": (c, k) => bbp(c, k)?.rjcUsedCW,
  "bbp.model": (c, k) => bbp(c, k)?.model,
  "bbp.rJTop": (c, k) => bbp(c, k)?.rJTopCW,
  "bbp.rJBoard": (c, k) => bbp(c, k)?.rJBoardCW,
  "bbp.exposed": (c, k) => bbp(c, k)?.exposedMm2,
  "bbp.rSurfaceAir": (c, k) => bbp(c, k)?.rSurfaceAirCW,
  "bbp.rTop": (c, k) => bbp(c, k)?.rTopCW,
  "bbp.footprint": (c, k) => bbp(c, k)?.footprintMm2,
  "bbp.share": (c, k) => bbp(c, k)?.shareMm2,
  "bbp.eps": (c, k) => bbp(c, k)?.spread?.eps,
  "bbp.tau": (c, k) => bbp(c, k)?.spread?.tau,
  "bbp.bi": (c, k) => bbp(c, k)?.spread?.bi,
  "bbp.lambda": (c, k) => bbp(c, k)?.spread?.lambda,
  "bbp.phi": (c, k) => bbp(c, k)?.spread?.phi,
  "bbp.rSpread": (c, k) => bbp(c, k)?.rSpreadCW,
  "bbp.rBoard": (c, k) => bbp(c, k)?.rBoardCW,
  "bbp.f": (c, k) => bbp(c, k)?.f,
  "bbp.c": (c, k) => bbp(c, k)?.cWK,
  "bbp.fPd": (c, k) => bbp(c, k)?.fPdW,
  "bbp.qBoard": (c, k) => bbp(c, k)?.qBoardW,
  "bbp.qTop": (c, k) => bbp(c, k)?.qTopW,
  "bbp.boardShare": (c, k) => (isNum(bbp(c, k)?.qBoardW) && bbp(c, k).pdW > 0 ? bbp(c, k).qBoardW / bbp(c, k).pdW : undefined),
  "bbp.topShare": (c, k) => (isNum(bbp(c, k)?.qTopW) && bbp(c, k).pdW > 0 ? bbp(c, k).qTopW / bbp(c, k).pdW : undefined),
  "bbp.tj": (c, k) => bbp(c, k)?.tjC,
  "bbp.margin": (c, k) => bbp(c, k)?.marginC,
  "bbp.status": (c, k) => bbp(c, k)?.status,
  "bbp.simC": (c, k) => c.calc.bareBoard.compare.parts?.[k]?.simC,
  "bbp.gap": (c, k) => c.calc.bareBoard.compare.parts?.[k]?.gapC,

  "build.assembly": c => c.cad?.build?.assembly,
  "build.date": c => c.cad?.build?.date,
  "build.mode": c => c.cad?.build?.mode,
  "build.wanted": c => c.cad?.build?.counts?.wanted,
  "build.added": c => c.cad?.build?.counts?.added,
  "build.moved": c => c.cad?.build?.counts?.moved,
  "build.resized": c => c.cad?.build?.counts?.resized,
  "build.suppressed": c => c.cad?.build?.counts?.suppressed,
  "build.broughtBack": c => c.cad?.build?.counts?.broughtBack,
  "build.failed": c => c.cad?.build?.counts?.failed,
  "build.overlaps": c => c.cad?.build?.overlaps?.length,
  "build.included": c => c.calc.bodies.filter(b => b.include).length,
  "buildbody.status": (c, k) => c.cad?.build?.bodies?.find(b => b.key === k)?.status,

  "plate.enabled": (c, k) => plateIn(c, k)?.enabled,
  "plate.material": (c, k) => plateIn(c, k)?.material,
  "plate.k": (c, k) => plateIn(c, k)?.k,
  "plate.thickness": (c, k) => plateIn(c, k)?.thicknessMm,
  "plate.minPedestal": (c, k) => plateIn(c, k)?.minPedestalMm,
  "plate.heightMode": (c, k) => plateIn(c, k)?.heightMode,
  "plate.heightManual": (c, k) => plateIn(c, k)?.heightManualMm,
  "plate.clearanceZ": (c, k) => plateIn(c, k)?.clearanceZMm,
  "plate.faceZ": (c, k) => c.calc.spreader[k]?.faceZMm,
  "plate.covered": (c, k) => c.calc.spreader[k]?.covered,
  "plate.cuts": (c, k) => c.calc.spreader[k]?.cuts.length,
  "pchip.refDes": (c, k) => pchip(c, k).c?.refDes,
  "pchip.covered": (c, k) => pchip(c, k).c?.covered === true,
  "pchip.heightMode": (c, k) => pchip(c, k).c?.heightMode,
  "pchip.timTSet": (c, k) => pchip(c, k).c?.timTMm,
  "pchip.chipFace": (c, k) => pchip(c, k).r?.faceZ,
  "pchip.zNeed": (c, k) => pchip(c, k).r?.zNeed,
  "pchip.gap": (c, k) => pchip(c, k).r?.gapMm,
  "pchip.timT": (c, k) => pchip(c, k).r?.timTMm,
  "pchip.pedestal": (c, k) => pchip(c, k).r?.pedestalMm,
  "pchip.check": (c, k) => pchip(c, k).r?.check,

  "sim.refDes": (c, k) => c.calc.simChips[+k]?.refDes,
  "sim.model": (c, k) => c.calc.simChips[+k]?.model,
  "sim.pd": (c, k) => c.calc.simChips[+k]?.pdW,
  "sim.rjc": (c, k) => c.calc.simChips[+k]?.rjcCW,
  "sim.h": (c, k) => c.calc.simChips[+k]?.hMm,
  "sim.area": (c, k) => c.calc.simChips[+k]?.areaMm2,
  "sim.kz": (c, k) => c.calc.simChips[+k]?.kz,
  "sim.kxySet": (c, k) => c.project.simSetup?.chips?.[+k]?.kxy,
  "sim.kxy": (c, k) => c.calc.simChips[+k]?.kxy,
  "sim.check": (c, k) => c.calc.simChips[+k]?.check,

  "run.name": (c, k) => runOf(c, k)?.name,
  "run.date": (c, k) => runOf(c, k)?.date,
  "run.studyType": (c, k) => runOf(c, k)?.studyType,
  "run.status": (c, k) => runOf(c, k)?.status,
  "run.energyBalance": (c, k) => runOf(c, k)?.energyBalancePct,
  "run.pcbMin": (c, k) => runOf(c, k)?.pcb?.minC,
  "run.pcbMax": (c, k) => runOf(c, k)?.pcb?.maxC,
  "run.spreaderMin": (c, k) => runOf(c, k)?.spreader?.minC,
  "run.spreaderMax": (c, k) => runOf(c, k)?.spreader?.maxC,
  "runp.peak": (c, k) => runp(c, k)?.peakC,
  "runp.avg": (c, k) => runp(c, k)?.avgC,

  "change.checked": (c, k) => !!c.project.checks?.[k],
  "change.date": (c, k) => change(c, k)?.date,
  "change.macro": (c, k) => change(c, k)?.macro,
  "change.body": (c, k) => change(c, k)?.name,
  "change.change": (c, k) => change(c, k)?.change,
  "change.field": (c, k) => change(c, k)?.field,
  "change.old": (c, k) => change(c, k)?.old,
  "change.new": (c, k) => change(c, k)?.new,
};

export function valueOf(ctx, term, key = "") {
  const f = VALUES[term];
  if (!f) return undefined;
  try { return f(ctx, key); } catch { return undefined; }
}

// ---------- formulas ----------
// text: the formula as the user reads it. expr: the right side with the input symbols, where the numbers go
// (default: the right side of text). fn: works the value out from the inputs, by symbol. A formula without fn
// is a rule in words (flags, counts, ranks).
const F = {};

F["board.totalPd"] = c => {
  const ins = [];
  for (const p of c.project.parts) for (const l of p.loads) if (isNum(c.calc.register.loads[l.id].pdW)) ins.push([`Pd(${l.label || p.refDes})`, "load.pd", l.id]);
  return { text: "Board Pd = sum of Pd of every load", expr: ins.map(x => x[0]).join(" + ") || "0", inputs: ins,
    fn: v => ins.reduce((a, x) => a + v[x[0]], 0) };
};

F["part.pd"] = (c, k) => {
  const p = partOf(c, k);
  if (!p) return null;
  const ins = p.loads.filter(l => isNum(c.calc.register.loads[l.id].pdW)).map((l, i) => [p.loads.length > 1 ? `Pd${i + 1}` : "Pd load", "load.pd", l.id]);
  if (!ins.length) return null;
  return { text: p.loads.length > 1 ? "Pd = sum of the Pd of its loads" : "Pd = Pd of its load", expr: ins.map(x => x[0]).join(" + "), inputs: ins,
    fn: v => ins.reduce((a, x) => a + v[x[0]], 0), note: p.loads.length > 1 ? "A part with several heat sources (a PMIC with four bucks) heats as one part." : undefined };
};
F["part.ta"] = (c, k) => (isNum(partOf(c, k)?.taOverrideC)
  ? { text: "Ta = Ta override", inputs: [["Ta override", "part.taOverride", k]], fn: v => v["Ta override"] }
  : { text: "Ta = ambient", inputs: [["ambient", "board.ambient", ""]], fn: v => v.ambient, note: "No Ta override for this part, so it sees the board ambient." });
F["part.area"] = (c, k) => { const b = bodyOfRef(c, k); return b && { text: "A = L × W", inputs: [["L", "body.l", b.key], ["W", "body.w", b.key]], fn: v => v.L * v.W, note: `Footprint of its body, ${b.name}.` }; };
F["body.pd"] = (c, k) => { const r = bodyRow(c, k)?.refDes; return r && { text: "Pd = Pd of its part", inputs: [["Pd", "part.pd", r]], fn: v => N(v.Pd) }; };
F["body.density"] = (c, k) => (isNum(geoBody(c, k)?.densityWcm2)
  ? { text: "Density = Pd / (L × W / 100)", inputs: [["Pd", "body.pd", k], ["L", "body.l", k], ["W", "body.w", k]], fn: v => v.Pd / (v.L * v.W / 100), note: "100 mm² = 1 cm²." } : null);
F["body.check"] = (c, k) => ({ text: "OK unless: missing geometry, RefDes not in register, RefDes used twice, heat source not included, or overhangs board edge (in this order)",
  inputs: [["L", "body.l", k], ["W", "body.w", k], ["H", "body.h", k], ["X min", "body.xMin", k], ["Y min", "body.yMin", k], ["Part", "body.refDes", k], ["Include", "body.include", k], ["PCB L", "pcb.l", ""], ["PCB W", "pcb.w", ""]] });
F["geo.included"] = () => ({ text: "Number of bodies with Include ticked", inputs: [] });
F["geo.placed"] = c => {
  const ins = c.calc.bodies.filter(b => b.include && b.name && b.refDes).map(b => [`Pd(${b.refDes})`, "body.pd", b.key]);
  return { text: "Pd placed = sum of Pd of every included body with a part", expr: ins.map(x => x[0]).join(" + ") || "0", inputs: ins, fn: v => ins.reduce((a, x) => a + N(v[x[0]]), 0) };
};
F["geo.notPlaced"] = () => ({ text: "Pd with no body = Board Pd - Pd placed", inputs: [["Board Pd", "board.totalPd", ""], ["Pd placed", "geo.placed", ""]], fn: v => v["Board Pd"] - v["Pd placed"] });
F["geo.noBody"] = () => ({ text: "Number of parts that no included body carries", inputs: [] });
F["geo.problems"] = () => ({ text: "Number of bodies whose check is not OK (overhangs only warn)", inputs: [] });
F["part.rise"] = (c, k) => ({ text: "ΔT = Pd × Rja", inputs: [["Pd", "part.pd", k], ["Rja", "part.rja", k]], fn: v => v.Pd * v.Rja });
F["part.tjEst"] = (c, k) => ({ text: "Tj est = Ta + Pd × Rja", inputs: [["Ta", "part.ta", k], ["Pd", "part.pd", k], ["Rja", "part.rja", k]], fn: v => v.Ta + v.Pd * v.Rja });
F["part.margin"] = (c, k) => ({ text: "Margin = Tj max - Tj est", inputs: [["Tj max", "part.tjMax", k], ["Tj est", "part.tjEst", k]], fn: v => v["Tj max"] - v["Tj est"] });
F["part.density"] = (c, k) => ({ text: "Density = Pd / (A / 100)", inputs: [["Pd", "part.pd", k], ["A", "part.area", k]], fn: v => v.Pd / (v.A / 100), note: "100 mm² = 1 cm²." });
F["part.share"] = (c, k) => ({ text: "Share = Pd / Board Pd", inputs: [["Pd", "part.pd", k], ["Board Pd", "board.totalPd", ""]], fn: v => v.Pd / v["Board Pd"] });
F["part.flag"] = (c, k) => ({
  text: "CRITICAL when margin < limit, or density > limit, or share > limit (CRITICAL limits); else Watch with the Watch limits; else Minor",
  expr: "margin vs critical margin, density vs critical density, share vs critical share; then the Watch limits",
  inputs: [["margin", "part.margin", k], ["density", "part.density", k], ["share", "part.share", k],
    ["critical margin", "limit.criticalMargin", ""], ["critical density", "limit.criticalDensity", ""], ["critical share", "limit.criticalShare", ""],
    ["Watch margin", "limit.watchMargin", ""], ["Watch density", "limit.watchDensity", ""], ["Watch share", "limit.watchShare", ""]],
  note: "A check that cannot run (no Rja, no body) does not flag. Why flagged says which limit was crossed.",
});
F["part.why"] = (c, k) => ({ text: "The limits this part crossed, with the numbers, and the checks that could not run", inputs: [["flag", "part.flag", k], ["margin", "part.margin", k], ["density", "part.density", k], ["share", "part.share", k]] });
F["part.rankPd"] = (c, k) => ({ text: "Rank = position by Pd, largest first (equal Pd: register order)", inputs: [["Pd", "part.pd", k]] });
F["part.rankDensity"] = (c, k) => ({ text: "Rank = position by density, largest first (equal density: register order)", inputs: [["density", "part.density", k]] });

F["load.methodUsed"] = (c, k) => {
  const x = loadOf(c, k);
  if (!x) return null;
  if (x.load.method) return { text: "Method = method set on the load", inputs: [["method", "load.method", k]] };
  return { text: "Method = the default of the part class", inputs: [["class", "part.class", x.part.refDes]], note: x.part.class ? `${x.part.class} uses ${CLASS_METHOD[x.part.class] || "no method (not in the class map)"}.` : "No class picked yet." };
};
F["load.pin"] = (c, k) => ({ text: "Pin = Vin × Iin", inputs: [["Vin", "load.vin", k], ["Iin", "load.iin", k]], fn: v => v.Vin * v.Iin });
F["load.pout"] = (c, k) => ({ text: "Pout = Vout × Iout", inputs: [["Vout", "load.vout", k], ["Iout", "load.iout", k]], fn: v => v.Vout * v.Iout });
F["load.pdCalc"] = (c, k) => {
  const x = loadOf(c, k);
  if (!x || !isNum(x.r.pdCalcW)) return null;
  const { method, pinW, poutW } = x.r, l = x.load;
  const both = isNum(pinW) && isNum(poutW);
  if ((method === "LDO / linear reg" || method === "Switching regulator") && both)
    return { text: "Pd = Pin - Pout", inputs: [["Pin", "load.pin", k], ["Pout", "load.pout", k]], fn: v => v.Pin - v.Pout, note: "Input and output power are both known, so the loss is their difference." };
  if (method === "LDO / linear reg")
    return { text: "Pd = (Vin - Vout) × Iout", inputs: [["Vin", "load.vin", k], ["Vout", "load.vout", k], ["Iout", "load.iout", k]], fn: v => (v.Vin - v.Vout) * v.Iout,
      note: "A linear regulator drops the voltage difference at the full load current." };
  if (method === "Switching regulator")
    return { text: "Pd = Pout × (1 / η - 1)", inputs: [["Pout", "load.pout", k], ["η", "load.eff", k]], fn: v => v.Pout * (1 / v["η"] - 1), note: "Pin = Pout / η, so the loss is Pin - Pout." };
  if (method === "Digital load (V x I)")
    return isNum(poutW) ? { text: "Pd = Pin - Pout", inputs: [["Pin", "load.pin", k], ["Pout", "load.pout", k]], fn: v => v.Pin - v.Pout }
      : { text: "Pd = Pin", inputs: [["Pin", "load.pin", k]], fn: v => v.Pin, note: "All the power a digital part draws ends up as heat in it." };
  if (method === "Resistive (I squared R)")
    return isNum(l.iinA) ? { text: "Pd = Iin² × R", inputs: [["Iin", "load.iin", k], ["R", "load.r", k]], fn: v => v.Iin ** 2 * v.R }
      : { text: "Pd = Iout² × R", inputs: [["Iout", "load.iout", k], ["R", "load.r", k]], fn: v => v.Iout ** 2 * v.R };
  return null;
};
F["load.pd"] = (c, k) => (isNum(loadOf(c, k)?.load.pdMeasuredW)
  ? { text: "Pd = Pd measured", inputs: [["Pd measured", "load.pdMeas", k]], fn: v => v["Pd measured"], note: "A measured or datasheet Pd replaces the calculated one." }
  : isNum(loadOf(c, k)?.r.pdCalcW) ? { text: "Pd = Pd calc", inputs: [["Pd calc", "load.pdCalc", k]], fn: v => v["Pd calc"] } : null);

F["rail.pd"] = (c, k) => {
  const ins = [];
  for (const p of c.project.parts) for (const l of p.loads) if (l.rail === k && isNum(c.calc.register.loads[l.id].pdW)) ins.push([`Pd(${l.label || p.refDes})`, "load.pd", l.id]);
  return { text: "Rail Pd = sum of the Pd of the loads on this rail", expr: ins.map(x => x[0]).join(" + ") || "0", inputs: ins, fn: v => ins.reduce((a, x) => a + v[x[0]], 0) };
};
F["rail.share"] = (c, k) => ({ text: "Share = Rail Pd / Board Pd", inputs: [["Rail Pd", "rail.pd", k], ["Board Pd", "board.totalPd", ""]], fn: v => v["Rail Pd"] / v["Board Pd"] });
F["rail.parts"] = (c, k) => ({ text: "Number of parts with a load on this rail", inputs: [] });
F["rails.sum"] = c => { const ins = c.calc.rails.rails.map(r => [`Pd(${r.rail})`, "rail.pd", r.rail]); return { text: "Sum of rails = sum of every rail Pd", expr: ins.map(x => x[0]).join(" + ") || "0", inputs: ins, fn: v => ins.reduce((a, x) => a + v[x[0]], 0) }; };
F["rails.noRail"] = () => ({ text: "No rail = Board Pd - Sum of rails", inputs: [["Board Pd", "board.totalPd", ""], ["Sum of rails", "rails.sum", ""]], fn: v => v["Board Pd"] - v["Sum of rails"] });
F["rails.measured"] = () => ({ text: "Measured input = measured board input power", inputs: [["measured", "board.measuredInputPower", ""]], fn: v => v.measured });
F["rails.unaccounted"] = () => ({ text: "Not accounted = Measured - Board Pd", inputs: [["Measured", "board.measuredInputPower", ""], ["Board Pd", "board.totalPd", ""]], fn: v => v.Measured - v["Board Pd"] });
F["rails.unaccountedFrac"] = () => ({ text: "Share not accounted = Not accounted / Measured", inputs: [["Not accounted", "rails.unaccounted", ""], ["Measured", "board.measuredInputPower", ""]], fn: v => v["Not accounted"] / v.Measured });

for (const [id, word] of [["critical", "CRITICAL"], ["watch", "Watch"], ["minor", "Minor"], ["noPd", "No Pd"]]) {
  F["triage." + id] = () => ({ text: `Number of parts flagged ${word}`, inputs: [] });
}
F["triage.parts"] = () => ({ text: "Number of parts on tab 1", inputs: [] });
F["triage.noRja"] = () => ({ text: "Number of parts without Rja (their margin is not checked)", inputs: [] });
F["triage.noBody"] = () => ({ text: "Number of parts without a body (their density is not checked)", inputs: [] });

F["pcb.cuTotal"] = () => ({ text: "Copper total = layers × t copper", inputs: [["layers", "stack.layers", ""], ["t copper", "stack.cuT", ""]], fn: v => v.layers * v["t copper"] });
F["pcb.cuFraction"] = () => ({ text: "f = Copper total × coverage / H", inputs: [["Copper total", "pcb.cuTotal", ""], ["coverage", "stack.coverage", ""], ["H", "pcb.h", ""]], fn: v => v["Copper total"] * v.coverage / v.H });
F["pcb.kxy"] = () => ({ text: "k xy = f × k Cu + (1 - f) × k diel", inputs: [["f", "pcb.cuFraction", ""], ["k Cu", "stack.cuK", ""], ["k diel", "stack.dielK", ""]],
  fn: v => v.f * v["k Cu"] + (1 - v.f) * v["k diel"], note: "In plane the copper and the dielectric carry heat side by side (in parallel), so copper dominates." });
F["pcb.kz"] = () => ({ text: "k z = H / ((H - Copper total) / k diel + Copper total / (coverage × k Cu + (1 - coverage) × k diel))",
  inputs: [["H", "pcb.h", ""], ["Copper total", "pcb.cuTotal", ""], ["k diel", "stack.dielK", ""], ["coverage", "stack.coverage", ""], ["k Cu", "stack.cuK", ""]],
  fn: v => v.H / ((v.H - v["Copper total"]) / v["k diel"] + v["Copper total"] / (v.coverage * v["k Cu"] + (1 - v.coverage) * v["k diel"])),
  note: "Through the board the layers are in series, so the dielectric dominates." });

// h of radiation, linearised: the same for the three surface groups.
const hRad = (epsTerm) => c => (isNum(valueOf(c, epsTerm))
  ? { text: "h rad = ε × σ × (T1² + T2²) × (T1 + T2), T1 = Ta + 273.15, T2 = Ta + ΔT + 273.15",
      expr: "ε × σ × ((Ta + 273.15)² + (Ta + ΔT + 273.15)²) × (2 × Ta + ΔT + 546.3)",
      inputs: [["ε", epsTerm, ""], ["σ", "const.sigma", ""], ["Ta", "board.ambient", ""], ["ΔT", "hand.surfaceRise", ""]],
      fn: v => { const t1 = v.Ta + 273.15, t2 = v.Ta + N(v["ΔT"]) + 273.15; return v["ε"] * v["σ"] * (t1 ** 2 + t2 ** 2) * (t1 + t2); },
      note: "Radiation written as an h, so it adds to convection. Temperatures in kelvin." }
  : { text: "h rad = 0 (no emissivity set)", inputs: [["ε", epsTerm, ""]], fn: () => 0 });
F["bb.hRadTop"] = hRad("surf.pcbTopEps");
F["bb.hRadBottom"] = hRad("surf.pcbBottomEps");
F["bb.hRadParts"] = hRad("surf.partsEps");
F["bb.area"] = () => ({ text: "A board = L × W", inputs: [["L", "pcb.l", ""], ["W", "pcb.w", ""]], fn: v => v.L * v.W });
F["bb.hPcb"] = () => ({ text: "h PCB = h top + h rad top + h bottom + h rad bottom",
  inputs: [["h top", "surf.pcbTopH", ""], ["h rad top", "bb.hRadTop", ""], ["h bottom", "surf.pcbBottomH", ""], ["h rad bottom", "bb.hRadBottom", ""]],
  fn: v => N(v["h top"]) + v["h rad top"] + N(v["h bottom"]) + v["h rad bottom"], note: "Both faces lose heat, so their h add up on the board area." });
F["bb.hParts"] = () => ({ text: "h parts = h + h rad", inputs: [["h", "surf.partsH", ""], ["h rad", "bb.hRadParts", ""]], fn: v => N(v.h) + v["h rad"] });
F["bb.rBoardAir"] = () => ({ text: "R board-air = 1 000 000 / (h PCB × A board)", inputs: [["h PCB", "bb.hPcb", ""], ["A board", "bb.area", ""]],
  fn: v => 1e6 / (v["h PCB"] * v["A board"]), note: "R = 1 / (h A). A is in mm², so 10⁶ turns it into m²." });
F["bb.totalPd"] = c => { const ins = c.project.parts.map(p => [`Pd(${p.refDes})`, "bbp.pd", p.refDes]); return { text: "Pd total = sum of every part's Pd", expr: ins.map(x => x[0]).join(" + "), inputs: ins, fn: v => ins.reduce((a, x) => a + v[x[0]], 0) }; };
F["bb.rise"] = c => {
  const ps = c.project.parts.filter(p => isNum(bbp(c, p.refDes)?.fPdW));
  const fins = ps.map(p => [`fPd(${p.refDes})`, "bbp.fPd", p.refDes]), cins = ps.map(p => [`C(${p.refDes})`, "bbp.c", p.refDes]);
  return { text: "ΔT board = sum of f × Pd / (1 / R board-air + sum of C)",
    expr: `(${fins.map(x => x[0]).join(" + ")}) / (1 / R board-air + ${cins.map(x => x[0]).join(" + ")})`,
    inputs: [...fins, ["R board-air", "bb.rBoardAir", ""], ...cins],
    fn: v => fins.reduce((a, x) => a + v[x[0]], 0) / (1 / v["R board-air"] + cins.reduce((a, x) => a + v[x[0]], 0)),
    note: "Heat balance of the shared board node: what the parts push into the board equals what the board loses to air, less what flows back up through the parts." };
};
F["bb.boardC"] = () => ({ text: "T board = Ta + ΔT board", inputs: [["Ta", "board.ambient", ""], ["ΔT board", "bb.rise", ""]], fn: v => v.Ta + v["ΔT board"] });
F["bb.qBoard"] = c => { const ins = c.project.parts.filter(p => isNum(bbp(c, p.refDes)?.qBoardW)).map(p => [`Q(${p.refDes})`, "bbp.qBoard", p.refDes]); return { text: "Q into board = sum of every part's heat into the board", expr: ins.map(x => x[0]).join(" + "), inputs: ins, fn: v => ins.reduce((a, x) => a + v[x[0]], 0) }; };
F["bb.fail"] = () => ({ text: "Number of parts with status FAIL (Tj above Tj max)", inputs: [] });
F["bb.tight"] = () => ({ text: "Number of parts with status Tight (margin under the CRITICAL margin limit)", inputs: [["limit", "limit.criticalMargin", ""]] });
F["bb.meanGap"] = c => { const ins = Object.keys(c.calc.bareBoard.compare.parts || {}).filter(r => isNum(c.calc.bareBoard.compare.parts[r].gapC)).map(r => [`gap(${r})`, "bbp.gap", r]); return { text: "Mean gap = average of (Tj hand - Tj sim)", expr: `(${ins.map(x => x[0]).join(" + ")}) / ${ins.length}`, inputs: ins, fn: v => ins.reduce((a, x) => a + v[x[0]], 0) / ins.length }; };
F["bb.maxGap"] = c => { const ins = Object.keys(c.calc.bareBoard.compare.parts || {}).filter(r => isNum(c.calc.bareBoard.compare.parts[r].gapC)).map(r => [`gap(${r})`, "bbp.gap", r]); return { text: "Largest gap = largest |Tj hand - Tj sim|", expr: `largest of |${ins.map(x => x[0]).join("|, |")}|`, inputs: ins, fn: v => Math.max(...ins.map(x => Math.abs(v[x[0]]))) }; };

F["bbp.pd"] = (c, k) => ({ text: "Pd = Pd of the part (tab 1)", inputs: [["Pd", "part.pd", k]], fn: v => N(v.Pd) });
F["bbp.rjcUsed"] = (c, k) => {
  const p = partOf(c, k), b = bodyOfRef(c, k);
  if (N(p?.rjcCW) > 0) return { text: "Rjc used = Rjc", inputs: [["Rjc", "part.rjc", k]], fn: v => v.Rjc };
  if (!b) return { text: "Rjc used = 0 (no Rjc and no body)", inputs: [], fn: () => 0 };
  return { text: "Rjc used = H × 1000 / (k parts × L × W)", inputs: [["H", "bbp.h", k], ["k parts", "simset.kPartsWithoutHeat", ""], ["L", "bbp.l", k], ["W", "bbp.w", k]],
    fn: v => (v.H * 1000) / (v["k parts"] * v.L * v.W), note: "No Rjc on the datasheet: the body is treated as a block of the k used for parts without heat." };
};
F["bbp.rJTop"] = (c, k) => (N(partOf(c, k)?.rjbCW) > 0
  ? { text: "R j-top = Rjc used", inputs: [["Rjc used", "bbp.rjcUsed", k]], fn: v => v["Rjc used"], note: "Two-resistor model: Rjc to the top, Rjb to the board." }
  : { text: "R j-top = Rjc used / 2", inputs: [["Rjc used", "bbp.rjcUsed", k]], fn: v => v["Rjc used"] / 2, note: "Block model: the junction sits in the middle of the block, so half of Rjc on each side." });
F["bbp.rJBoard"] = (c, k) => (N(partOf(c, k)?.rjbCW) > 0
  ? { text: "R j-board = Rjb", inputs: [["Rjb", "part.rjb", k]], fn: v => v.Rjb }
  : { text: "R j-board = Rjc used / 2", inputs: [["Rjc used", "bbp.rjcUsed", k]], fn: v => v["Rjc used"] / 2 });
F["bbp.exposed"] = (c, k) => ({ text: "A exposed = L × W + 2 × (L + W) × H", inputs: [["L", "bbp.l", k], ["W", "bbp.w", k], ["H", "bbp.h", k]], fn: v => v.L * v.W + 2 * (v.L + v.W) * v.H, note: "Top face plus the four sides of the body." });
F["bbp.rSurfaceAir"] = (c, k) => ({ text: "R surface-air = 1 000 000 / (h parts × A exposed)", inputs: [["h parts", "bb.hParts", ""], ["A exposed", "bbp.exposed", k]], fn: v => 1e6 / (v["h parts"] * v["A exposed"]) });
F["bbp.rTop"] = (c, k) => ({ text: "R top = R j-top + R surface-air", inputs: [["R j-top", "bbp.rJTop", k], ["R surface-air", "bbp.rSurfaceAir", k]], fn: v => v["R j-top"] + v["R surface-air"], note: "The top path: junction to case top, then case surface to air, in series." });
F["bbp.footprint"] = (c, k) => ({ text: "A footprint = L × W", inputs: [["L", "bbp.l", k], ["W", "bbp.w", k]], fn: v => v.L * v.W });
F["bbp.share"] = (c, k) => ({ text: "A share = A board × Pd / Pd total", inputs: [["A board", "bb.area", ""], ["Pd", "bbp.pd", k], ["Pd total", "bb.totalPd", ""]], fn: v => v["A board"] * v.Pd / v["Pd total"], note: "Each part gets the share of the board area that matches its share of the heat." });
F["bbp.eps"] = (c, k) => ({ text: "ε = min(1, √(A footprint / A share))", inputs: [["A footprint", "bbp.footprint", k], ["A share", "bbp.share", k]], fn: v => Math.min(1, Math.sqrt(v["A footprint"] / v["A share"])), ref: "Lee et al. 1995" });
F["bbp.tau"] = (c, k) => ({ text: "τ = t / √(A share / π)", inputs: [["t", "pcb.h", ""], ["A share", "bbp.share", k]], fn: v => v.t / Math.sqrt(v["A share"] / Math.PI), note: "Board thickness over the radius of its share.", ref: "Lee et al. 1995" });
F["bbp.bi"] = (c, k) => ({ text: "Bi = h PCB × √(A share / π) / 1000 / k xy", inputs: [["h PCB", "bb.hPcb", ""], ["A share", "bbp.share", k], ["k xy", "pcb.kxy", ""]], fn: v => v["h PCB"] * Math.sqrt(v["A share"] / Math.PI) / 1000 / v["k xy"], note: "Biot number: how easily heat leaves the board face compared with how easily it spreads in it.", ref: "Lee et al. 1995" });
F["bbp.lambda"] = (c, k) => ({ text: "λ = π + 1 / (ε × √π)", inputs: [["ε", "bbp.eps", k]], fn: v => Math.PI + 1 / (v["ε"] * Math.sqrt(Math.PI)), ref: "Lee et al. 1995" });
F["bbp.phi"] = (c, k) => ({ text: "Φ = (tanh(λ × τ) + λ / Bi) / (1 + λ / Bi × tanh(λ × τ))", inputs: [["λ", "bbp.lambda", k], ["τ", "bbp.tau", k], ["Bi", "bbp.bi", k]],
  fn: v => { const th = Math.tanh(v["λ"] * v["τ"]); return (th + v["λ"] / v.Bi) / (1 + v["λ"] / v.Bi * th); }, ref: "Lee et al. 1995" });
F["bbp.rSpread"] = (c, k) => ({ text: "R spread = (ε × τ / √π + (1 - ε) × Φ / √π) / (√π × k xy × √(A footprint / π) / 1000)",
  inputs: [["ε", "bbp.eps", k], ["τ", "bbp.tau", k], ["Φ", "bbp.phi", k], ["k xy", "pcb.kxy", ""], ["A footprint", "bbp.footprint", k]],
  fn: v => (v["ε"] * v["τ"] / Math.sqrt(Math.PI) + (1 - v["ε"]) * v["Φ"] / Math.sqrt(Math.PI)) / (Math.sqrt(Math.PI) * v["k xy"] * Math.sqrt(v["A footprint"] / Math.PI) / 1000),
  note: "Resistance of heat spreading out from the small footprint into the larger board share, and through the board.", ref: "Lee, Song, Au, Moran 1995, spreading resistance of a disc source on a plate" });
F["bbp.rBoard"] = (c, k) => ({ text: "R board = R j-board + R spread", inputs: [["R j-board", "bbp.rJBoard", k], ["R spread", "bbp.rSpread", k]], fn: v => v["R j-board"] + v["R spread"], note: "The board path: junction to board, then spreading into the board, in series." });
F["bbp.f"] = (c, k) => (isNum(bbp(c, k)?.rTopCW)
  ? { text: "f = R top / (R top + R board)", inputs: [["R top", "bbp.rTop", k], ["R board", "bbp.rBoard", k]], fn: v => v["R top"] / (v["R top"] + v["R board"]), note: "Share of the part's heat that goes into the board, if the board were at ambient." }
  : { text: "f = 1 (no top path)", inputs: [], fn: () => 1 });
F["bbp.c"] = (c, k) => (isNum(bbp(c, k)?.rTopCW)
  ? { text: "C = 1 / (R top + R board)", inputs: [["R top", "bbp.rTop", k], ["R board", "bbp.rBoard", k]], fn: v => 1 / (v["R top"] + v["R board"]), note: "Conductance from the board node back up through the part to air." }
  : { text: "C = 0 (no top path)", inputs: [], fn: () => 0 });
F["bbp.fPd"] = (c, k) => ({ text: "f × Pd", expr: "f × Pd", inputs: [["f", "bbp.f", k], ["Pd", "bbp.pd", k]], fn: v => v.f * v.Pd });
F["bbp.qBoard"] = (c, k) => ({ text: "Q board = f × Pd - C × ΔT board", inputs: [["f", "bbp.f", k], ["Pd", "bbp.pd", k], ["C", "bbp.c", k], ["ΔT board", "bb.rise", ""]], fn: v => v.f * v.Pd - v.C * v["ΔT board"] });
F["bbp.boardShare"] = (c, k) => ({ text: "Share = Q board / Pd", inputs: [["Q board", "bbp.qBoard", k], ["Pd", "bbp.pd", k]], fn: v => v["Q board"] / v.Pd });
F["bbp.topShare"] = (c, k) => ({ text: "Share = Q top / Pd", inputs: [["Q top", "bbp.qTop", k], ["Pd", "bbp.pd", k]], fn: v => v["Q top"] / v.Pd });
F["bb.boardShare"] = () => ({ text: "Share = Q board / Pd", inputs: [["Q board", "bb.qBoard", ""], ["Pd", "bb.totalPd", ""]], fn: v => v["Q board"] / v.Pd, note: "Q board: the heat all parts put into the board. Pd: the board's total Pd." });
F["bbp.qTop"] = (c, k) => ({ text: "Q top = Pd - Q board", inputs: [["Pd", "bbp.pd", k], ["Q board", "bbp.qBoard", k]], fn: v => v.Pd - v["Q board"] });
F["bbp.tj"] = (c, k) => (isNum(bbp(c, k)?.rTopCW)
  ? { text: "Tj = Ta + Q top × R top", inputs: [["Ta", "board.ambient", ""], ["Q top", "bbp.qTop", k], ["R top", "bbp.rTop", k]], fn: v => v.Ta + v["Q top"] * v["R top"] }
  : { text: "Tj = T board + Q board × R board", inputs: [["T board", "bb.boardC", ""], ["Q board", "bbp.qBoard", k], ["R board", "bbp.rBoard", k]], fn: v => v["T board"] + v["Q board"] * v["R board"] });
F["bbp.margin"] = (c, k) => ({ text: "Margin = Tj max - Tj", inputs: [["Tj max", "part.tjMax", k], ["Tj", "bbp.tj", k]], fn: v => v["Tj max"] - v.Tj });
F["bbp.status"] = (c, k) => ({ text: "FAIL when margin < 0; Tight when margin < the CRITICAL margin limit; else OK", inputs: [["margin", "bbp.margin", k], ["limit", "limit.criticalMargin", ""]] });
F["bbp.gap"] = (c, k) => ({ text: "Gap = Tj hand - Tj sim", inputs: [["Tj hand", "bbp.tj", k], ["Tj sim", "bbp.simC", k]], fn: v => v["Tj hand"] - v["Tj sim"] });
F["bbp.model"] = (c, k) => ({ text: "Two-resistor when Rjb is given; Block when only Rjc is given; else the part k is used", inputs: [["Rjc", "part.rjc", k], ["Rjb", "part.rjb", k]] });
F["load.problem"] = (c, k) => ({ text: "What the method of this load still needs", inputs: [["method", "load.methodUsed", k]] });
F["bb.compareCheck"] = () => ({ text: "OK to compare when the run is a bare board Thermal run; a run with a spreader, or a Flow run, is flagged", inputs: [["run", "bb.compareRun", ""]] });
F["build.included"] = () => ({ text: "Number of bodies with Include ticked on tab 3", inputs: [] });

F["plate.faceZ"] = (c, k) => {
  const pl = plateIn(c, k);
  if (!pl) return null;
  if (pl.heightMode === "Manual") return { text: "Face = manual height", inputs: [["manual height", "plate.heightManual", k]], fn: v => N(v["manual height"]) };
  const ch = c.calc.spreader[k].chips.map((r, i) => [r, i]).filter(([r]) => isNum(r.zNeed) && (k === "bottom" || r.covered));
  const ins = ch.map(([r, i]) => [`z(${r.refDes})`, "pchip.zNeed", `${k}/${i}`]);
  return k === "top"
    ? { text: "Face = highest (chip top + TIM) of the covered chips", expr: `max(0, ${ins.map(x => x[0]).join(", ")})`, inputs: ins, fn: v => Math.max(0, ...ins.map(x => v[x[0]])), note: "The plate underside sits on the tallest covered chip and its TIM." }
    : { text: "Face = lowest (chip bottom - TIM) of the covered chips", expr: ins.length ? `min(${ins.map(x => x[0]).join(", ")})` : "0", inputs: ins, fn: v => (ins.length ? Math.min(...ins.map(x => v[x[0]])) : 0) };
};
F["plate.covered"] = () => ({ text: "Number of chips with Covered ticked", inputs: [] });
F["plate.cuts"] = (c, k) => ({ text: "Number of included bodies, not covered, that reach the plate face (within the Z clearance): each gets a clearance cut", inputs: [["face", "plate.faceZ", k], ["Z clearance", "plate.clearanceZ", k]] });
F["pchip.zNeed"] = (c, k) => { const x = pchip(c, k); if (!x.r?.covered || !isNum(x.r.faceZ)) return null; return x.side === "top"
  ? { text: "z = chip top + TIM t", inputs: [["chip top", "pchip.chipFace", k], ["TIM t", "pchip.timTSet", k]], fn: v => v["chip top"] + N(v["TIM t"]) }
  : { text: "z = chip bottom - TIM t", inputs: [["chip bottom", "pchip.chipFace", k], ["TIM t", "pchip.timTSet", k]], fn: v => v["chip bottom"] - N(v["TIM t"]) }; };
F["pchip.chipFace"] = (c, k) => { const x = pchip(c, k); const b = x.r?.key; if (!b) return null; return x.side === "top"
  ? { text: "chip top = Z min + H", inputs: [["Z min", "body.zMin", b], ["H", "body.h", b]], fn: v => v["Z min"] + v.H }
  : { text: "chip bottom = Z min of the body", inputs: [["Z min", "body.zMin", b]], fn: v => v["Z min"] }; };
F["pchip.gap"] = (c, k) => { const x = pchip(c, k); if (!isNum(x.r?.gapMm)) return null; return x.side === "top"
  ? { text: "Gap = Face - chip top", inputs: [["Face", "plate.faceZ", "top"], ["chip top", "pchip.chipFace", k]], fn: v => v.Face - v["chip top"] }
  : { text: "Gap = chip bottom - Face", inputs: [["chip bottom", "pchip.chipFace", k], ["Face", "plate.faceZ", "bottom"]], fn: v => v["chip bottom"] - v.Face }; };
F["pchip.timT"] = (c, k) => { const x = pchip(c, k); if (!isNum(x.r?.timTMm)) return null;
  if (x.c.heightMode === "Thick TIM") return { text: "TIM = Gap (Thick TIM fills the whole gap)", inputs: [["Gap", "pchip.gap", k]], fn: v => v.Gap };
  if (x.r.gapMm - N(x.c.timTMm) < x.pl.minPedestalMm) return { text: "TIM = Gap (a pedestal would be thinner than the smallest pedestal)", inputs: [["Gap", "pchip.gap", k], ["TIM t", "pchip.timTSet", k], ["smallest pedestal", "plate.minPedestal", x.side]], fn: v => v.Gap };
  return { text: "TIM = TIM t", inputs: [["TIM t", "pchip.timTSet", k]], fn: v => N(v["TIM t"]) }; };
F["pchip.pedestal"] = (c, k) => (isNum(pchip(c, k).r?.pedestalMm) ? { text: "Pedestal = Gap - TIM", inputs: [["Gap", "pchip.gap", k], ["TIM", "pchip.timT", k]], fn: v => v.Gap - v.TIM, note: "A raised block on the plate that fills the rest of the gap." } : null);
F["pchip.check"] = (c, k) => ({ text: "OK unless: not covered, wrong side, chip above the plate face, TIM too thin for a body (< 0.01 mm), or a big gap for a contact resistance (> 0.3 mm)", inputs: [["Gap", "pchip.gap", k], ["TIM", "pchip.timT", k]] });

F["sim.pd"] = (c, k) => ({ text: "Pd = Pd of the part (tab 1)", inputs: [["Pd", "part.pd", c.calc.simChips[+k].refDes]], fn: v => N(v.Pd) });
F["sim.rjc"] = (c, k) => ({ text: "Rjc = Rjc of the part (tab 1)", inputs: [["Rjc", "part.rjc", c.calc.simChips[+k].refDes]], fn: v => v.Rjc });
F["sim.h"] = (c, k) => { const b = bodyOfRef(c, c.calc.simChips[+k].refDes); return b && { text: "H = height of its body", inputs: [["H", "body.h", b.key]], fn: v => v.H }; };
F["sim.area"] = (c, k) => { const b = bodyOfRef(c, c.calc.simChips[+k].refDes); return b && { text: "A = L × W of its body", inputs: [["L", "body.l", b.key], ["W", "body.w", b.key]], fn: v => v.L * v.W }; };
F["sim.kz"] = (c, k) => (isNum(c.calc.simChips[+k]?.kz)
  ? { text: "k z = H × 1000 / (Rjc × A)", inputs: [["H", "sim.h", k], ["Rjc", "sim.rjc", k], ["A", "sim.area", k]], fn: v => v.H * 1000 / (v.Rjc * v.A),
      note: "The chip is one block. This k makes the block's resistance top to bottom equal Rjc: R = H / (k A)." }
  : null);
F["sim.kxy"] = (c, k) => (isNum(c.project.simSetup?.chips?.[+k]?.kxy)
  ? { text: "k xy = k xy typed on tab 8", inputs: [["k xy", "sim.kxySet", k]], fn: v => v["k xy"] }
  : isNum(c.calc.simChips[+k]?.kz) ? { text: "k xy = k z (none typed)", inputs: [["k z", "sim.kz", k]], fn: v => v["k z"] } : null);
F["sim.check"] = (c, k) => ({ text: "OK unless: no body, no heat, two-resistor model (not built yet), or no Rjc", inputs: [["k z", "sim.kz", k], ["Pd", "sim.pd", k]] });

export const FORMULAS = F;
// Values worked out by the page (not typed, not from SolidWorks) without a formula entry would be a gap:
// tools/test_w4.mjs checks that every term of kind "calculated" in the catalogue has one.

export function formulaOf(ctx, term, key = "") {
  const f = F[term];
  if (!f) return null;
  let out;
  try { out = f(ctx, key); } catch { return null; }
  if (!out) return null;
  const expr = out.expr ?? out.text.split(" = ").slice(1).join(" = ").split(", ")[0];
  return { ...out, expr };
}

// The inputs of a formula with their values, and the result worked out from them.
export function evaluate(ctx, term, key = "") {
  const f = formulaOf(ctx, term, key);
  if (!f) return null;
  const vals = {};
  for (const [sym, t, k] of f.inputs) vals[sym] = valueOf(ctx, t, k);
  let result;
  if (f.fn) { try { result = f.fn(vals); } catch { result = undefined; } }
  return { ...f, vals, result, value: valueOf(ctx, term, key) };
}

// Every [term, key] in this project, for "used by" and the tests.
export function instances(ctx) {
  const out = [];
  const add = (prefix, keys) => { for (const t of Object.keys(VALUES)) if (t.startsWith(prefix + ".")) for (const k of keys) out.push([t, k]); };
  const { project: p, calc } = ctx;
  for (const t of Object.keys(VALUES)) if (/^(const|board|limit|rails|triage|import|set|pcb|pcbov|stack|geo|surf|hand|simset|bb|build)\./.test(t) && t !== "buildbody.status") out.push([t, ""]);
  add("part", p.parts.map(x => x.refDes));
  add("load", p.parts.flatMap(x => x.loads.map(l => l.id)));
  add("rail", calc.rails.rails.map(r => r.rail));
  add("body", calc.bodies.map(b => b.key));
  add("bodyset", calc.bodies.map(b => b.key));
  add("bbp", p.parts.map(x => x.refDes));
  add("buildbody", (ctx.cad?.build?.bodies || []).map(b => b.key));
  const sides = ["top", "bottom"].filter(s => p.spreader?.[s] && calc.spreader[s]);
  add("plate", sides);
  add("pchip", sides.flatMap(s => p.spreader[s].chips.map((_, i) => `${s}/${i}`)));
  add("sim", calc.simChips.map((_, i) => String(i)));
  const runs = ctx.results?.runs || [];
  add("run", runs.map(r => r.name));
  add("runp", runs.flatMap(r => Object.keys(r.parts || {}).map(x => `${r.name}|${x}`)));
  add("change", (ctx.cad?.changes || []).map(x => x.id));
  return out;
}

// Instances whose formula uses this one (the "used by" list of the Explain panel).
export function usedBy(ctx, term, key = "") {
  const out = [];
  for (const [t, k] of instances(ctx)) {
    const f = F[t] && formulaOf(ctx, t, k);
    if (f && f.inputs.some(([, it, ik]) => it === term && ik === key)) out.push([t, k]);
  }
  return out;
}
