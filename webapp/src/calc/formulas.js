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
import { leeSpreading } from "./lee.js";
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
// Thermal solution (tab 7 hand check with a common plate): sol.* by side, solp.* by "side|RefDes".
const solIn = (c, side) => c.project.handCheck?.[side === "bottom" ? "solutionBottom" : "solutionTop"];
const solOut = (c, side) => c.calc.solution?.[side];
const solp = (c, key) => { const [side, ref] = key.split("|"); return { side, ref, inp: solIn(c, side)?.parts?.[ref] || {}, r: solOut(c, side)?.parts?.[ref], s: solIn(c, side) }; };
const runCalc = (c, name) => c.calc.runs?.[name];
const cmpRow = (c, key) => { const [ref, mode] = key.split("|"); return c.calc.compare?.[mode === "any" ? "any" : "ok"]?.rows?.[ref]; };
const runpCalc = (c, key) => { const [run, ref] = key.split("|"); return c.calc.runs?.[run]?.parts?.[ref]; };
const sprb = (c, side) => c.cad?.spreaderBuild?.[side];
const asbp = (c, side) => c.cad?.asBuilt?.plates?.[side];

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

  // ---------- W7 ----------
  "plate.density": (c, k) => plateIn(c, k)?.densityKgM3,
  "plate.clearanceXY": (c, k) => plateIn(c, k)?.clearanceXYMm,
  "pchip.timMode": (c, k) => pchip(c, k).c?.timMode,
  "pchip.timK": (c, k) => pchip(c, k).c?.timK,
  "pchip.timR": (c, k) => pchip(c, k).r?.timRCW,
  "sol.plateL": (c, k) => solIn(c, k)?.plateLMm,
  "sol.plateW": (c, k) => solIn(c, k)?.plateWMm,
  "sol.plateT": (c, k) => solIn(c, k)?.plateTMm,
  "sol.plateK": (c, k) => solIn(c, k)?.plateK,
  "sol.hAir": (c, k) => solIn(c, k)?.hAir,
  "sol.faces": (c, k) => solIn(c, k)?.faces,
  "sol.rOverride": (c, k) => solIn(c, k)?.rPlateToAirOverrideCW,
  "sol.area": (c, k) => solOut(c, k)?.plateAreaMm2,
  "sol.hEff": (c, k) => solOut(c, k)?.hEff,
  "sol.rPlateAir": (c, k) => solOut(c, k)?.rPlateAirCW,
  "sol.partsUnder": (c, k) => solOut(c, k)?.partsUnder,
  "sol.pdPlate": (c, k) => solOut(c, k)?.pdPlateW,
  "sol.plateC": (c, k) => solOut(c, k)?.plateC,
  "sol.footprints": (c, k) => solOut(c, k)?.footprintsMm2,
  "solp.underPlate": (c, k) => solp(c, k).inp.underPlate === true,
  "solp.tim1T": (c, k) => solp(c, k).inp.tim1TMm,
  "solp.tim1K": (c, k) => solp(c, k).inp.tim1K,
  "solp.localL": (c, k) => solp(c, k).inp.localLMm,
  "solp.localW": (c, k) => solp(c, k).inp.localWMm,
  "solp.localT": (c, k) => solp(c, k).inp.localTMm,
  "solp.localK": (c, k) => solp(c, k).inp.localK,
  "solp.tim2T": (c, k) => solp(c, k).inp.tim2TMm,
  "solp.tim2K": (c, k) => solp(c, k).inp.tim2K,
  "solp.shareOverride": (c, k) => solp(c, k).inp.plateShareOverrideMm2,
  "solp.package": (c, k) => solp(c, k).r?.packageMm2,
  "solp.footprint": (c, k) => solp(c, k).r?.footprintMm2,
  "solp.plateShare": (c, k) => solp(c, k).r?.plateShareMm2,
  "solp.rjc": (c, k) => solp(c, k).r?.rjcCW,
  "solp.rTim1": (c, k) => solp(c, k).r?.rTim1CW,
  "solp.rLocal": (c, k) => solp(c, k).r?.rLocalCW,
  "solp.rTim2": (c, k) => solp(c, k).r?.rTim2CW,
  "solp.rPlateSpread": (c, k) => solp(c, k).r?.rPlateSpreadCW,
  "solp.rStack": (c, k) => solp(c, k).r?.rStackCW,
  "solp.dtStack": (c, k) => solp(c, k).r?.dtStackC,
  "solp.tj": (c, k) => solp(c, k).r?.tjC,
  "solp.margin": (c, k) => solp(c, k).r?.marginC,
  "solp.status": (c, k) => solp(c, k).r?.status,
  "sprb.file": (c, k) => sprb(c, k)?.file,
  "sprb.date": (c, k) => sprb(c, k)?.date,
  "sprb.pedestals": (c, k) => sprb(c, k)?.pedestals,
  "sprb.tims": (c, k) => sprb(c, k)?.tims?.length,
  "sprb.cuts": (c, k) => sprb(c, k)?.cuts,
  "sprb.failures": (c, k) => sprb(c, k)?.failures?.length,
  "asb.date": c => c.cad?.asBuilt?.date,
  "asb.bodies": c => c.cad?.asBuilt?.bodies?.length,
  "asb.failedChecks": c => c.cad?.asBuilt?.failedChecks?.length,
  "asbp.outlineArea": (c, k) => asbp(c, k)?.outlineAreaMm2,
  "asbp.topArea": (c, k) => asbp(c, k)?.topAreaMm2,
  "asbp.exposedArea": (c, k) => asbp(c, k)?.exposedAreaMm2,
  "asbp.volume": (c, k) => asbp(c, k)?.volumeMm3,
  "asbp.mass": (c, k) => asbp(c, k)?.massG,
  "asbp.thickness": (c, k) => asbp(c, k)?.thicknessMm,
  "asbp.holes": (c, k) => asbp(c, k)?.holes,
  "surf.spreaderTopH": c => surf(c).spreaderTop?.h,
  "surf.spreaderTopEps": c => surf(c).spreaderTop?.emissivity,
  "surf.spreaderOtherH": c => surf(c).spreaderOther?.h,
  "surf.spreaderOtherEps": c => surf(c).spreaderOther?.emissivity,
  "simset.studyType": c => c.project.simSetup?.studyType,
  "simset.meshGlobal": c => c.project.simSetup?.meshGlobalMm,
  "simset.meshFine": c => c.project.simSetup?.meshFineMm,
  "simset.includePartsWithoutHeat": c => c.project.simSetup?.includePartsWithoutHeat,
  "simset.contactPartPcb": c => c.project.simSetup?.contacts?.partToPcb,
  "simset.contactChipTim": c => c.project.simSetup?.contacts?.chipToTim,
  "simset.contactTimSpreader": c => c.project.simSetup?.contacts?.timToSpreader,
  "simset.layerFrac": c => c.project.simSetup?.layerFrac,
  "simset.kCore": c => c.project.simSetup?.kCore,
  "run.study": (c, k) => runOf(c, k)?.study,
  "run.meshGlobal": (c, k) => runOf(c, k)?.mesh?.globalMm,
  "run.meshFine": (c, k) => runOf(c, k)?.mesh?.fineMm,
  "run.nodes": (c, k) => runOf(c, k)?.mesh?.nodes,
  "run.elements": (c, k) => runOf(c, k)?.mesh?.elements,
  "run.plot": (c, k) => runOf(c, k)?.plotFile,
  "run.worst": (c, k) => runCalc(c, k)?.worst,
  "run.worstC": (c, k) => runCalc(c, k)?.worstC,
  "run.pass": (c, k) => runCalc(c, k)?.pass,
  "run.fails": (c, k) => runCalc(c, k)?.fails,
  "run.balanceOk": (c, k) => runCalc(c, k)?.balanceOk,
  "run.health": (c, k) => ({ ok: "OK", check: "Check", bad: "Bad" })[runCalc(c, k)?.health?.level],
  "run.healthReasons": (c, k) => { const h = runCalc(c, k)?.health; return h ? (h.reasons.length ? h.reasons.join("; ") : "none") : undefined; },
  "run.heatRequested": (c, k) => runOf(c, k)?.heatRequestedW,
  "run.heatApplied": (c, k) => runOf(c, k)?.heatAppliedW,
  "run.solverFinished": (c, k) => runOf(c, k)?.solver?.finished,
  "run.iterations": (c, k) => runOf(c, k)?.solver?.iterations,
  "run.travels": (c, k) => runOf(c, k)?.solver?.travels,
  "run.cellsTotal": (c, k) => runOf(c, k)?.solver?.cellsTotal,
  "run.cellsSolid": (c, k) => runOf(c, k)?.solver?.cellsSolid,
  "run.cellsFluid": (c, k) => runOf(c, k)?.solver?.cellsFluid,
  "run.meshLevel": (c, k) => runOf(c, k)?.mesh?.level,
  "run.localMesh": (c, k) => runOf(c, k)?.mesh?.localMesh,
  "run.meshCheckPair": (c, k) => runOf(c, k)?.meshCheck?.pairWith,
  "run.meshCheckOk": (c, k) => { const m = runCalc(c, k)?.meshCheck; return m?.ok === undefined ? undefined : m.ok ? "OK" : "Not OK"; },
  "run.meshCheckWorst": (c, k) => runCalc(c, k)?.meshCheck?.worstC,
  "runp.meshDelta": (c, k) => { const [run, ref] = k.split("|"); return runCalc(c, run)?.meshCheck?.parts?.[ref]; },
  "flow.meshLevel": c => c.calc.flow?.settings.meshLevel,
  "flow.localMesh": c => c.calc.flow?.settings.localMesh,
  "flow.solidRefine": c => c.calc.flow?.settings.solidRefine,
  "flow.meshCheck": c => c.calc.flow?.settings.meshCheck,
  "flow.meshCheckLimit": c => c.calc.flow?.settings.meshCheckLimitC,
  "runp.tjBasis": (c, k) => runpCalc(c, k)?.tjBasis,
  "fchip.basis": (c, k) => c.calc.fsNow?.flow?.chips?.[k]?.basis,
  "fchip.material": (c, k) => { const x = c.calc.fsNow?.flow?.chips?.[k]; return x?.material || (x?.layers ? `${x.layers.bottom} / ${x.layers.core} / ${x.layers.top}` : undefined); },
  "fchip.rjcStep": (c, k) => { const ch = c.calc.simChips.find(x => x.refDes === k), m = c.calc.fsNow?.flow?.materials.find(x => x.path === c.calc.fsNow?.flow?.chips?.[k]?.material); return ch && m && ch.hMm > 0 && ch.areaMm2 > 0 ? (ch.hMm * 1000) / (m.kz * ch.areaMm2) : undefined; },
  "fmat.name": (c, k) => c.calc.fsNow?.flow?.materials?.[+k]?.name,
  "fmat.kxy": (c, k) => c.calc.fsNow?.flow?.materials?.[+k]?.kxy,
  "fmat.kz": (c, k) => c.calc.fsNow?.flow?.materials?.[+k]?.kz,
  "fmat.usedBy": (c, k) => c.calc.fsNow?.flow?.materials?.[+k]?.usedBy?.join(", "),
  "fsw.pcbMat": c => c.calc.fsNow?.flow?.pcb,
  "fsw.timMat": (c, k) => c.calc.fsNow?.flow?.tims?.[k],
  "fsw.sprMat": (c, k) => c.calc.fsNow?.flow?.spreaders?.[k],
  "fsw.surface": (c, k) => c.calc.fsNow?.flow?.surfaces?.[k],
  "fsw.contactR": (c, k) => c.calc.fsNow?.flow?.contacts?.[k]?.rKW,
  "fsw.contactA": (c, k) => c.calc.fsNow?.flow?.contacts?.[k]?.areaMm2,
  "frad.name": (c, k) => c.calc.fsNow?.flow?.radiative?.[+k]?.name,
  "frad.eps": (c, k) => c.calc.fsNow?.flow?.radiative?.[+k]?.emissivity,
  "frad.usedBy": (c, k) => c.calc.fsNow?.flow?.radiative?.[+k]?.usedBy?.join(", "),
  "hnb.pd": (c, k) => c.calc.heatNoBody?.find(x => x.refDes === k)?.pdW,
  "hnb.accepted": (c, k) => c.calc.heatNoBody?.find(x => x.refDes === k)?.accepted,
  "env.orientation": c => c.calc.flow?.env.orientation,
  "env.pressure": c => c.project.simSetup?.flow?.environment?.pressureKPa,
  "env.altitude": c => c.project.simSetup?.flow?.environment?.altitudeM,
  "env.airSpeed": c => c.calc.flow?.env.airSpeedMs,
  "env.airDirection": c => c.project.simSetup?.flow?.environment?.airDirection,
  "env.pressureUsed": c => c.calc.flow?.env.pressureKPa,
  "env.up": c => c.calc.flow?.env.up,
  "cmp.includeCheck": () => undefined,
  "cmp.hand": (c, k) => cmpRow(c, k)?.handC,
  "cmp.thermal": (c, k) => cmpRow(c, k)?.thermalC,
  "cmp.flow": (c, k) => cmpRow(c, k)?.flowC,
  "cmp.meas": (c, k) => cmpRow(c, k)?.measC,
  "cmp.gapHand": (c, k) => cmpRow(c, k)?.gapHand,
  "cmp.gapThermal": (c, k) => cmpRow(c, k)?.gapThermal,
  "cmp.gapFlow": (c, k) => cmpRow(c, k)?.gapFlow,
  "cmp.flowVsThermal": (c, k) => cmpRow(c, k)?.flowVsThermal,
  "meas.name": (c, k) => c.project.measurements?.[+k]?.name,
  "meas.date": (c, k) => c.project.measurements?.[+k]?.date,
  "meas.ambient": (c, k) => c.project.measurements?.[+k]?.ambientC,
  "meas.note": (c, k) => c.project.measurements?.[+k]?.note,
  "meas.t": (c, k) => { const [i, r] = k.split("|"); return c.project.measurements?.[+i]?.parts?.[r]?.tC; },
  "meas.method": (c, k) => { const [i, r] = k.split("|"); return c.project.measurements?.[+i]?.parts?.[r]?.method; },
  "meas.where": (c, k) => { const [i, r] = k.split("|"); return c.project.measurements?.[+i]?.parts?.[r]?.where; },
  "flow.minTravels": c => c.calc.flow?.settings.minTravels,
  "flow.goalCriteria": c => c.calc.flow?.settings.goalCriteriaC,
  "run.goalsLoose": (c, k) => { const r = runOf(c, k), lim = c.calc.flow?.settings.goalCriteriaC; if (!r?.goalsHealth || !isNum(lim)) return undefined; const l = Object.entries(r.goalsHealth).filter(([, g]) => g.deltaC > lim).map(([n]) => n); return l.length ? l.join(", ") : "none"; },
  "run.worstDelta": (c, k) => { const g = Object.values(runOf(c, k)?.goalsHealth || {}).map(x => x.deltaC).filter(isNum); return g.length ? Math.max(...g) : undefined; },
  "flow.domainVol": c => { const d = c.calc.flow?.domain; return d ? (d.x[1] - d.x[0]) * (d.y[1] - d.y[0]) * (d.z[1] - d.z[0]) : undefined; },
  "flow.minCell": c => c.calc.flow?.minCellMm,
  "flow.thin": c => { const t = c.calc.flow?.thin; return t ? (t.length ? t.map(x => `${x.refDes} (${x.hMm} mm)`).join(", ") : "none") : undefined; },
  "run.unresolved": (c, k) => { const u = runOf(c, k)?.unresolved; return u?.length ? u.join(", ") : undefined; },
  "run.changes": (c, k) => { const ch = runCalc(c, k)?.changes; return ch === undefined ? undefined : ch.length ? ch.join("; ") : "none"; },
  "runp.margin": (c, k) => runpCalc(c, k)?.marginC,
  "runp.status": (c, k) => runpCalc(c, k)?.status,
  "runp.hand": (c, k) => runpCalc(c, k)?.handC,
  "runp.gap": (c, k) => runpCalc(c, k)?.gapC,
  "build.outOfDate": c => c.calc.build?.outOfDate,
  "help.search": () => undefined,
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

// ---------- W7: TIM resistance, thermal solution, runs ----------
F["pchip.timR"] = (c, k) => { const x = pchip(c, k); if (!isNum(x.r?.timRCW) || !x.r.key) return null;
  return { text: "R TIM = TIM × 1000 / (k TIM × L × W)", inputs: [["TIM", "pchip.timT", k], ["k TIM", "pchip.timK", k], ["L", "body.l", x.r.key], ["W", "body.w", x.r.key]],
    fn: v => v.TIM * 1000 / (v["k TIM"] * v.L * v.W), note: "A layer of thickness t and area A conducts with R = t / (k A); 1000 turns mm into m." }; };

F["sol.area"] = (c, k) => (isNum(solOut(c, k)?.plateAreaMm2) ? { text: "A plate = L plate × W plate", inputs: [["L plate", "sol.plateL", k], ["W plate", "sol.plateW", k]], fn: v => v["L plate"] * v["W plate"] } : null);
F["sol.hEff"] = (c, k) => {
  const s = solIn(c, k), o = solOut(c, k);
  if (!isNum(o?.hEff)) return null;
  return isNum(s.rPlateToAirOverrideCW)
    ? { text: "h eff = 1 / (R typed × A plate / 10⁶)", expr: "1 / (R typed × A plate / 1000000)", inputs: [["R typed", "sol.rOverride", k], ["A plate", "sol.area", k]], fn: v => 1 / ((v["R typed"] * v["A plate"]) / 1e6), note: "The h that gives the typed resistance over the plate area." }
    : { text: "h eff = h × faces", inputs: [["h", "sol.hAir", k], ["faces", "sol.faces", k]], fn: v => v.h * v.faces };
};
F["sol.rPlateAir"] = (c, k) => {
  const s = solIn(c, k), o = solOut(c, k);
  if (!isNum(o?.rPlateAirCW)) return null;
  return isNum(s.rPlateToAirOverrideCW)
    ? { text: "R plate-air = R typed", inputs: [["R typed", "sol.rOverride", k]], fn: v => v["R typed"] }
    : { text: "R plate-air = 1000000 / (h eff × A plate)", inputs: [["h eff", "sol.hEff", k], ["A plate", "sol.area", k]], fn: v => 1e6 / (v["h eff"] * v["A plate"]), note: "R = 1 / (h A); A in mm², so × 10⁶." };
};
F["sol.partsUnder"] = () => ({ text: "Number of parts with Under plate ticked", inputs: [] });
const solActive = (c, k) => c.project.parts.filter(p => solIn(c, k)?.parts?.[p.refDes]?.underPlate === true && isNum(c.calc.register.parts[p.refDes]?.pdW));
F["sol.pdPlate"] = (c, k) => {
  if (!solOut(c, k)) return null;
  const ins = solActive(c, k).map(p => [`Pd(${p.refDes})`, "part.pd", p.refDes]);
  return { text: "Pd plate = sum of Pd of the parts under the plate", expr: ins.map(x => x[0]).join(" + ") || "0", inputs: ins, fn: v => ins.reduce((a, x) => a + v[x[0]], 0) };
};
F["sol.plateC"] = (c, k) => (isNum(solOut(c, k)?.plateC) ? { text: "T plate = Ta + Pd plate × R plate-air", inputs: [["Ta", "board.ambient", ""], ["Pd plate", "sol.pdPlate", k], ["R plate-air", "sol.rPlateAir", k]], fn: v => v.Ta + v["Pd plate"] * v["R plate-air"] } : null);
F["sol.footprints"] = (c, k) => {
  if (!solOut(c, k)) return null;
  const ins = c.project.parts.filter(p => isNum(solOut(c, k).parts[p.refDes]?.footprintMm2)).map(p => [`A(${p.refDes})`, "solp.footprint", `${k}|${p.refDes}`]);
  return { text: "Footprints = sum of the footprints on the plate", expr: ins.map(x => x[0]).join(" + ") || "0", inputs: ins, fn: v => ins.reduce((a, x) => a + v[x[0]], 0) };
};
F["solp.package"] = (c, k) => { const x = solp(c, k); const b = bodyOfRef(c, x.ref); return b && isNum(x.r?.packageMm2) ? { text: "A package = L × W of its body", inputs: [["L", "body.l", b.key], ["W", "body.w", b.key]], fn: v => v.L * v.W } : null; };
F["solp.footprint"] = (c, k) => { const x = solp(c, k); if (!isNum(x.r?.footprintMm2)) return null;
  return isNum(x.inp.localLMm) ? { text: "A footprint = local L × local W", inputs: [["local L", "solp.localL", k], ["local W", "solp.localW", k]], fn: v => v["local L"] * v["local W"] }
    : { text: "A footprint = A package (no local spreader)", inputs: [["A package", "solp.package", k]], fn: v => v["A package"] }; };
F["solp.plateShare"] = (c, k) => { const x = solp(c, k); if (!isNum(x.r?.plateShareMm2)) return null;
  return isNum(x.inp.plateShareOverrideMm2) ? { text: "A share = plate share typed", inputs: [["typed", "solp.shareOverride", k]], fn: v => v.typed }
    : { text: "A share = A plate × Pd / Pd plate", inputs: [["A plate", "sol.area", x.side], ["Pd", "part.pd", x.ref], ["Pd plate", "sol.pdPlate", x.side]], fn: v => v["A plate"] * v.Pd / v["Pd plate"], note: "Each part heats a share of the plate in proportion to its Pd." }; };
F["solp.rjc"] = (c, k) => { const x = solp(c, k); return isNum(x.r?.rjcCW) ? { text: "Rjc = Rjc of the part (tab 1)", inputs: [["Rjc", "part.rjc", x.ref]], fn: v => v.Rjc } : null; };
F["solp.rTim1"] = (c, k) => (isNum(solp(c, k).r?.rTim1CW) ? { text: "R TIM1 = 1000 × t TIM1 / (k TIM1 × A package)", inputs: [["t TIM1", "solp.tim1T", k], ["k TIM1", "solp.tim1K", k], ["A package", "solp.package", k]], fn: v => 1000 * v["t TIM1"] / (v["k TIM1"] * v["A package"]) } : null);
const leeFx = (ins, note) => ({
  text: "R = Lee spreading (source A, plate A, t, k, h)", expr: `Lee(${ins.map(x => x[0]).join(", ")})`, inputs: ins,
  fn: v => leeSpreading({ sourceMm2: v[ins[0][0]], plateMm2: v[ins[1][0]], tMm: v[ins[2][0]], k: v[ins[3][0]], h: ins[4][3] ? ins[4][3](v) : v[ins[4][0]] }).r,
  note, ref: "Lee, Song, Au and Moran 1995",
});
F["solp.rLocal"] = (c, k) => { const x = solp(c, k); if (!isNum(x.r?.rLocalCW)) return null;
  if (!x.r.local) return { text: "R local = 0 (no local spreader)", inputs: [], fn: () => 0 };
  const f = leeFx([["A package", "solp.package", k], ["A footprint", "solp.footprint", k], ["t local", "solp.localT", k], ["k local", "solp.localK", k]],
    "The far face of the local spreader is cooled through TIM2, so its h is k TIM2 / t TIM2.");
  f.inputs.push(["k TIM2", "solp.tim2K", k], ["t TIM2", "solp.tim2T", k]);
  f.expr = "Lee(A package, A footprint, t local, k local, 1000 × k TIM2 / t TIM2)";
  f.fn = v => leeSpreading({ sourceMm2: v["A package"], plateMm2: v["A footprint"], tMm: v["t local"], k: v["k local"], h: 1000 * v["k TIM2"] / v["t TIM2"] }).r;
  return f; };
F["solp.rTim2"] = (c, k) => { const x = solp(c, k); if (!isNum(x.r?.rTim2CW)) return null;
  return x.r.local ? { text: "R TIM2 = 1000 × t TIM2 / (k TIM2 × A footprint)", inputs: [["t TIM2", "solp.tim2T", k], ["k TIM2", "solp.tim2K", k], ["A footprint", "solp.footprint", k]], fn: v => 1000 * v["t TIM2"] / (v["k TIM2"] * v["A footprint"]) }
    : { text: "R TIM2 = 0 (no local spreader)", inputs: [], fn: () => 0 }; };
F["solp.rPlateSpread"] = (c, k) => { const x = solp(c, k); if (!isNum(x.r?.rPlateSpreadCW)) return null;
  const f = leeFx([["A footprint", "solp.footprint", k], ["A share", "solp.plateShare", k], ["t plate", "sol.plateT", x.side], ["k plate", "sol.plateK", x.side], ["h eff", "sol.hEff", x.side]],
    "Heat enters the plate through the footprint and spreads over the part's share before it leaves to air.");
  f.expr = "Lee(A footprint, A share, t plate, k plate, h eff)";
  return f; };
F["solp.rStack"] = (c, k) => (isNum(solp(c, k).r?.rStackCW) ? { text: "R stack = Rjc + R TIM1 + R local + R TIM2 + R plate",
  inputs: [["Rjc", "solp.rjc", k], ["R TIM1", "solp.rTim1", k], ["R local", "solp.rLocal", k], ["R TIM2", "solp.rTim2", k], ["R plate", "solp.rPlateSpread", k]],
  fn: v => v.Rjc + v["R TIM1"] + v["R local"] + v["R TIM2"] + v["R plate"] } : null);
F["solp.dtStack"] = (c, k) => { const x = solp(c, k); return isNum(x.r?.dtStackC) ? { text: "ΔT stack = Pd × R stack", inputs: [["Pd", "part.pd", x.ref], ["R stack", "solp.rStack", k]], fn: v => v.Pd * v["R stack"] } : null; };
F["solp.tj"] = (c, k) => { const x = solp(c, k); return isNum(x.r?.tjC) ? { text: "Tj = T plate + ΔT stack", inputs: [["T plate", "sol.plateC", x.side], ["ΔT stack", "solp.dtStack", k]], fn: v => v["T plate"] + v["ΔT stack"] } : null; };
F["solp.margin"] = (c, k) => { const x = solp(c, k); return isNum(x.r?.marginC) ? { text: "Margin = Tj max - Tj", inputs: [["Tj max", "part.tjMax", x.ref], ["Tj", "solp.tj", k]], fn: v => v["Tj max"] - v.Tj } : null; };
F["solp.status"] = (c, k) => ({ text: "FAIL when the margin is below 0, Tight when it is below the CRITICAL margin limit, else OK (or what is missing)", inputs: [["Margin", "solp.margin", k], ["limit", "limit.criticalMargin", ""]] });

F["run.worst"] = (c, k) => ({ text: "The RefDes with the highest peak in this run", inputs: [["peak", "run.worstC", k]] });
F["run.worstC"] = (c, k) => {
  const r = runOf(c, k);
  const ins = Object.entries(r?.parts || {}).filter(([, t]) => isNum(t.peakC)).map(([ref]) => [`peak(${ref})`, "runp.peak", `${k}|${ref}`]);
  if (!ins.length) return null;
  return { text: "Hottest part peak = highest peak of all parts", expr: `max(${ins.map(x => x[0]).join(", ")})`, inputs: ins, fn: v => Math.max(...ins.map(x => v[x[0]])) };
};
F["run.pass"] = (c, k) => ({ text: "PASS when no part is above its Tj max (FAIL as soon as one is)", inputs: [["parts over", "run.fails", k]] });
F["run.fails"] = () => ({ text: "Number of parts with a peak above their Tj max", inputs: [] });
F["run.balanceOk"] = (c, k) => ({ text: "Yes when the energy balance error is 2 % or less", inputs: [["error", "run.energyBalance", k]] });
F["run.health"] = (c, k) => ({ text: "Bad when the solver did not finish, a part with heat was not meshed, or the energy balance error is above 2 %; Check when a goal is not converged, travels are below the minimum, Flow wrote other warnings or heat is left out; else OK",
  inputs: [["finished", "run.solverFinished", k], ["not meshed", "run.unresolved", k], ["error", "run.energyBalance", k], ["travels", "run.travels", k], ["heat asked", "run.heatRequested", k], ["heat applied", "run.heatApplied", k]] });
F["flow.domainVol"] = () => ({ text: "V = X × Y × Z of the domain RunFlow sets: the PCB and its bodies, plus half the largest board size at the sides and below and one and a half above", inputs: [] });
F["flow.minCell"] = c => (isNum(c.calc.flow?.minCellMm) ? {
  text: "Smallest cell = (V / (28000 × 7^(level - 3)))^(1/3), divided by 2^refinement with a local mesh",
  inputs: [["V", "flow.domainVol", ""], ["level", "flow.meshLevel", ""], ["local mesh", "flow.localMesh", ""], ["refinement", "flow.solidRefine", ""]],
  fn: v => Math.cbrt(v.V / (28000 * 7 ** (v.level - 3))) / (v["local mesh"] ? 2 ** v.refinement : 1),
  note: "An estimate calibrated on the RPi4 runs (level 3: 28 000 cells, level 5: 1.43 M). Flow also refines near small solids by itself, so a part a little thinner may still be meshed. The run health on tab 9 shows what happened." } : null);
F["flow.thin"] = () => ({ text: "Parts with heat whose body is thinner than the estimated smallest cell", inputs: [["smallest cell", "flow.minCell", ""]] });
F["run.meshCheckOk"] = (c, k) => ({ text: "OK when every chip moved less than the mesh check limit between the two levels", inputs: [["largest change", "run.meshCheckWorst", k], ["limit", "flow.meshCheckLimit", ""]] });
F["run.meshCheckWorst"] = () => ({ text: "The largest change of a chip peak between the two runs of the mesh check (finer minus coarser)", inputs: [] });
F["runp.meshDelta"] = () => ({ text: "Change = peak at the finer level - peak at the coarser level", inputs: [] });
F["run.goalsLoose"] = (c, k) => ({ text: "Goals whose last change (delta) is above the goal criterion of tab 8", inputs: [["criterion", "flow.goalCriteria", ""], ["largest delta", "run.worstDelta", k]] });
F["run.worstDelta"] = () => ({ text: "Largest delta of the temperature goals at the end of the solve (from the Flow info file)", inputs: [] });
F["runp.tjBasis"] = () => ({ text: "Calibrated block or two-resistor core: the peak is Tj. Package: no Rjc (or a generic material), the peak is the package temperature", inputs: [] });
F["fchip.basis"] = () => ({ text: "Two-resistor when the chip model is Two-resistor with Rjc and Rjb, calibrated block when the chip has a block k from Rjc, else package", inputs: [] });
F["fchip.material"] = () => ({ text: "User material named after the k of the chip in 10 % steps: k xy and k z rounded to 24 steps per decade", inputs: [] });
F["fchip.rjcStep"] = (c, k) => { const ch = c.calc.simChips.find(x => x.refDes === k); const i = c.calc.simChips.indexOf(ch);
  return ch ? { text: "Rjc given = H / (k z step × A)", inputs: [["H", "sim.h", String(i)], ["k z step", "fmat.kz", String(c.calc.fsNow?.flow?.materials.findIndex(x => x.path === c.calc.fsNow?.flow?.chips?.[k]?.material))], ["A", "sim.area", String(i)]],
    fn: v => (v.H * 1000) / (v["k z step"] * v.A), note: "The step is at most 5 % from the exact k, so the Rjc the block gives is within 5 % of the datasheet." } : null; };
F["fmat.name"] = () => ({ text: "k <k xy> <k z>, or k <k> when both are the same, after rounding to 10 % steps", inputs: [] });
F["fmat.kxy"] = () => ({ text: "k xy step = 10^(round(24 × log10(k xy)) / 24), 3 digits", inputs: [] });
F["fmat.kz"] = () => ({ text: "k z step = 10^(round(24 × log10(k z)) / 24), 3 digits", inputs: [] });
F["fmat.usedBy"] = () => ({ text: "The chips (and layers) that use this material", inputs: [] });
F["fsw.pcbMat"] = () => ({ text: "Material named after the PCB k xy and k z of the stack (tab 3), in 10 % steps", inputs: [["k xy", "pcb.kxy", ""], ["k z", "pcb.kz", ""]] });
F["fsw.timMat"] = () => ({ text: "Material named after k TIM eff = t / (t / k TIM + R chip to TIM), in 10 % steps", inputs: [["R chip to TIM", "simset.contactChipTim", ""]] });
F["fsw.sprMat"] = () => ({ text: "Material named after the plate k (tab 7), in 10 % steps", inputs: [] });
F["fsw.surface"] = () => ({ text: "Radiative surface named after the emissivity: PCB top (tab 5), parts (tab 5), spreader top (tab 8)", inputs: [] });
F["fsw.contactR"] = (c, k) => { const x = c.calc.fsNow?.flow?.contacts?.[k]; return x ? { text: "R integral = R contact / A", inputs: [["R contact", k === "partToPcb" ? "simset.contactPartPcb" : "simset.contactTimSpreader", ""], ["A", "fsw.contactA", k]], fn: v => v["R contact"] / (v.A * 1e-6) } : null; };
F["fsw.contactA"] = (c, k) => ({ text: k === "partToPcb" ? "A = sum of L × W of the bodies that sit on the PCB" : "A = sum of the TIM areas (1000 × t / (k × R TIM))", inputs: [] });
F["frad.name"] = () => ({ text: "e <emissivity, 2 decimals>", inputs: [] });
F["frad.eps"] = () => ({ text: "Emissivity rounded to 2 decimals", inputs: [] });
F["frad.usedBy"] = () => ({ text: "The component groups that get this radiative surface", inputs: [] });
F["hnb.pd"] = (c, k) => ({ text: "Pd of the part (tab 1)", inputs: [["Pd", "part.pd", k]], fn: v => v.Pd });
F["env.pressureUsed"] = c => { const e = c.project.simSetup?.flow?.environment || {};
  if (isNum(e.pressureKPa)) return { text: "p = the pressure typed", inputs: [["p", "env.pressure", ""]], fn: v => v.p };
  if (isNum(e.altitudeM)) return { text: "p = 101.325 × (1 - 2.25577e-5 × altitude)^5.25588 (standard atmosphere)", inputs: [["altitude", "env.altitude", ""]], fn: v => 101.325 * (1 - 2.25577e-5 * v.altitude) ** 5.25588 };
  return { text: "p = 101.325 kPa (sea level)", inputs: [], fn: () => 101.325 }; };
F["env.up"] = () => ({ text: "Against gravity: +Z flat parts up, -Z flat parts down, along the short board size for long side up, along the long size for short side up", inputs: [["orientation", "env.orientation", ""]] });
F["cmp.hand"] = (c, k) => ({ text: "Tj of the hand check: with the plate (tab 7) when the part is under it, else the bare board (tab 5)", inputs: [] });
F["cmp.thermal"] = (c, k) => ({ text: "Peak of the part in the newest Thermal run whose health is OK" + (k.endsWith("|any") ? " or Check" : ""), inputs: [] });
F["cmp.flow"] = (c, k) => ({ text: "Peak of the part in the newest Flow run whose health is OK" + (k.endsWith("|any") ? " or Check" : ""), inputs: [] });
F["cmp.meas"] = () => ({ text: "Temperature of the part in the newest measurement", inputs: [] });
F["cmp.gapHand"] = (c, k) => (isNum(cmpRow(c, k)?.gapHand) ? { text: "Gap = Tj hand - measured", inputs: [["Tj hand", "cmp.hand", k], ["measured", "cmp.meas", k]], fn: v => v["Tj hand"] - v.measured } : null);
F["cmp.gapThermal"] = (c, k) => (isNum(cmpRow(c, k)?.gapThermal) ? { text: "Gap = Thermal - measured", inputs: [["Thermal", "cmp.thermal", k], ["measured", "cmp.meas", k]], fn: v => v.Thermal - v.measured } : null);
F["cmp.gapFlow"] = (c, k) => (isNum(cmpRow(c, k)?.gapFlow) ? { text: "Gap = Flow - measured", inputs: [["Flow", "cmp.flow", k], ["measured", "cmp.meas", k]], fn: v => v.Flow - v.measured } : null);
F["cmp.flowVsThermal"] = (c, k) => (isNum(cmpRow(c, k)?.flowVsThermal) ? { text: "Flow - Thermal", inputs: [["Flow", "cmp.flow", k], ["Thermal", "cmp.thermal", k]], fn: v => v.Flow - v.Thermal } : null);
F["run.healthReasons"] = (c, k) => ({ text: "Each reason behind the health of the run, in plain words", inputs: [["health", "run.health", k]] });
F["run.changes"] = () => ({ text: "Inputs of the run (a copy kept in the results file) compared with the inputs now: ambient, PCB, Pd per part, bodies, chip models, plates, surfaces, contacts, mesh", inputs: [] });
F["runp.margin"] = (c, k) => { const ref = k.split("|")[1]; return isNum(runpCalc(c, k)?.marginC) ? { text: "Margin = Tj max - peak", inputs: [["Tj max", "part.tjMax", ref], ["peak", "runp.peak", k]], fn: v => v["Tj max"] - v.peak } : null; };
F["runp.status"] = (c, k) => ({ text: "FAIL when the peak is above Tj max, Tight when the margin is below the CRITICAL margin limit, else OK", inputs: [["Margin", "runp.margin", k], ["limit", "limit.criticalMargin", ""]] });
F["runp.hand"] = (c, k) => {
  const [run, ref] = k.split("|");
  const x = runpCalc(c, k);
  if (!isNum(x?.handC)) return null;
  if (x.handFrom === "thermal solution") {
    const side = ["top", "bottom"].find(s => isNum(c.calc.solution[s]?.parts?.[ref]?.tjC));
    return { text: "Tj hand = Tj with the plate (tab 7)", inputs: [["Tj plate", "solp.tj", `${side}|${ref}`]], fn: v => v["Tj plate"], note: `The run ${run} has a spreader and ${ref} is under the hand check plate.` };
  }
  return { text: "Tj hand = Tj of the bare board hand check (tab 5)", inputs: [["Tj bare", "bbp.tj", ref]], fn: v => v["Tj bare"] };
};
F["runp.gap"] = (c, k) => (isNum(runpCalc(c, k)?.gapC) ? { text: "Gap = Tj hand - peak", inputs: [["Tj hand", "runp.hand", k], ["peak", "runp.peak", k]], fn: v => v["Tj hand"] - v.peak } : null);
F["build.outOfDate"] = () => ({ text: "Yes when the hash of the PCB and bodies now differs from the one the last Build Board built from", inputs: [] });

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
  for (const t of Object.keys(VALUES)) if (/^(const|board|limit|rails|triage|import|set|pcb|pcbov|stack|geo|surf|hand|simset|bb|build|asb)\./.test(t) && t !== "buildbody.status") out.push([t, ""]);
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
  const solSides = ["top", "bottom"].filter(s => calc.solution?.[s]);
  add("sol", solSides);
  add("solp", solSides.flatMap(s => p.parts.map(x => `${s}|${x.refDes}`)));
  add("sprb", ["top", "bottom"].filter(s => ctx.cad?.spreaderBuild?.[s]));
  add("asbp", ["top", "bottom"].filter(s => ctx.cad?.asBuilt?.plates?.[s]));
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
