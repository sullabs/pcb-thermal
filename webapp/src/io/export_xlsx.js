// Workbook export (design section 9): a new .xlsx made by the page, values only, units in the headers and a note on
// every header from the term catalogue. The Method tab writes out every formula. Pure: takes the SheetJS module, so
// tools/test_w9.mjs runs it in Node and the page uses the copy built into it.
//
//   exportWorkbook(XLSX, ctx, runNames) -> SheetJS workbook (XLSX.write it to bytes)
//     ctx = { project, cad, results, calc }; runNames: the runs to put on the Results tab (all when empty)

import { valueOf, FORMULAS, formulaOf, instances } from "../calc/formulas.js";
import { TERMS } from "../terms/catalogue.js";
import { isNum } from "../calc/util.js";

export const SHEETS = ["Summary", "Parts", "Bodies", "Triage", "Rail budget", "Bare board network", "Spreader", "Sim setup", "Results", "Compare", "Changes", "Method"];

// One column: [term, key function (row) -> key, header text (default: label), value function (row) -> value (default: valueOf)]
function table(XLSX, ctx, cols, rows) {
  const head = cols.map(([t, , h]) => {
    const u = TERMS[t]?.unit;
    return (h || TERMS[t]?.label || t) + (u ? ` (${u})` : "");
  });
  const data = rows.map(r => cols.map(([t, kf, , vf]) => {
    const v = vf ? vf(r) : valueOf(ctx, t, kf ? kf(r) : "");
    return cell(v);
  }));
  const ws = XLSX.utils.aoa_to_sheet([head, ...data]);
  cols.forEach(([t], i) => note(XLSX, ws, i, t));
  ws["!cols"] = head.map((h, i) => ({ wch: Math.min(40, Math.max(8, h.length + 2, ...data.map(r => String(r[i] ?? "").length + 1))) }));
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  return ws;
}
const cell = v => (typeof v === "boolean" ? (v ? "Yes" : "No") : isNum(v) ? +v.toPrecision(10) : v ?? "");
function note(XLSX, ws, col, t) {
  const term = TERMS[t];
  const ref = XLSX.utils.encode_cell({ r: 0, c: col });
  if (!ws[ref] || !term) return;
  const kind = { input: "you type this", calculated: "calculated", cad: "from SolidWorks", run: "from a simulation run", constant: "constant" }[term.kind];
  ws[ref].c = [{ a: "PCB Thermal", t: `${term.label}${term.sym && term.sym !== term.label ? " (" + term.sym + ")" : ""}${term.unit ? ", " + term.unit : ""}. ${kind}.\n${term.short}` }];
  ws[ref].c.hidden = true;
}
// Two column list of [term, key] facts: label, value, unit.
function facts(XLSX, ctx, rows, title) {
  const aoa = [[title, "Value", "Unit"]];
  const terms = [];
  for (const [t, k = "", v] of rows) {
    const val = v !== undefined ? v : valueOf(ctx, t, k);
    if (val === undefined || val === "") continue;
    aoa.push([TERMS[t]?.label || t, cell(val), TERMS[t]?.unit || ""]);
    terms.push(t);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  terms.forEach((t, i) => {
    const ref = XLSX.utils.encode_cell({ r: i + 1, c: 0 });
    const term = TERMS[t];
    if (term) { ws[ref].c = [{ a: "PCB Thermal", t: term.short }]; ws[ref].c.hidden = true; }
  });
  ws["!cols"] = [{ wch: 36 }, { wch: 28 }, { wch: 10 }];
  return ws;
}

export function exportWorkbook(XLSX, ctx, runNames = []) {
  const { project: p, cad, results, calc } = ctx;
  const wb = XLSX.utils.book_new();
  const add = (name, ws) => XLSX.utils.book_append_sheet(wb, ws, name);
  const refs = p.parts.map(x => x.refDes);
  const runs = (results?.runs || []).filter(r => !runNames.length || runNames.includes(r.name));
  const last = runs[runs.length - 1];

  add("Summary", facts(XLSX, ctx, [["board.name"], ["board.stepFile"], ["board.ambient"], ["board.measuredInputPower"], ["board.totalPd"], ["triage.parts"], ["triage.critical"], ["triage.watch"],
    ["pcb.l"], ["pcb.w"], ["pcb.h"], ["pcb.kxy"], ["pcb.kz"], ["bb.boardC"], ["bb.fail"], ["bb.tight"], ["build.date"], ["build.outOfDate"],
    ...(last ? [["run.name", last.name], ["run.date", last.name], ["run.pass", last.name], ["run.worst", last.name], ["run.worstC", last.name], ["run.energyBalance", last.name]] : []),
    ["board.notes"]], "Board " + (p.board.name || "")));

  const loads = p.parts.flatMap(x => x.loads.map((l, j) => ({ x, l, j })));
  add("Parts", table(XLSX, ctx, [["part.refDes", r => r.x.refDes], ["part.partNumber", r => r.x.refDes], ["part.class", r => r.x.refDes], ["part.package", r => r.x.refDes],
    ["part.tjMax", r => r.x.refDes], ["part.rja", r => r.x.refDes], ["part.rjc", r => r.x.refDes], ["part.rjb", r => r.x.refDes],
    ["load.label", r => r.l.id], ["load.rail", r => r.l.id], ["load.methodUsed", r => r.l.id], ["load.vin", r => r.l.id], ["load.iin", r => r.l.id], ["load.vout", r => r.l.id],
    ["load.iout", r => r.l.id], ["load.eff", r => r.l.id], ["load.r", r => r.l.id], ["load.pdMeas", r => r.l.id], ["load.pd", r => r.l.id],
    ["part.pd", r => r.x.refDes, "Part Pd", r => (r.j === 0 ? valueOf(ctx, "part.pd", r.x.refDes) : "")], ["part.tjEst", r => r.x.refDes, , r => (r.j === 0 ? valueOf(ctx, "part.tjEst", r.x.refDes) : "")],
    ["part.margin", r => r.x.refDes, , r => (r.j === 0 ? valueOf(ctx, "part.margin", r.x.refDes) : "")], ["part.flag", r => r.x.refDes, , r => (r.j === 0 ? valueOf(ctx, "part.flag", r.x.refDes) : "")],
    ["part.notes", r => r.x.refDes, , r => (r.j === 0 ? r.x.notes : "")]], loads));

  add("Bodies", table(XLSX, ctx, [["body.name", b => b.key], ["body.include", b => b.key], ["body.side", b => b.key], ["body.refDes", b => b.key],
    ["body.xMin", b => b.key], ["body.yMin", b => b.key], ["body.zMin", b => b.key], ["body.l", b => b.key], ["body.w", b => b.key], ["body.h", b => b.key],
    ["body.pd", b => b.key], ["body.density", b => b.key], ["body.check", b => b.key]], calc.bodies));

  add("Triage", table(XLSX, ctx, [["part.refDes", r => r], ["part.rankPd", r => r], ["part.pd", r => r], ["part.rankDensity", r => r], ["part.density", r => r],
    ["part.share", r => r], ["part.tjEst", r => r], ["part.margin", r => r], ["part.flag", r => r], ["part.why", r => r]], calc.triage.byPd));

  add("Rail budget", table(XLSX, ctx, [["rail.name", r => r.rail], ["rail.parts", r => r.rail], ["rail.pd", r => r.rail], ["rail.share", r => r.rail]], calc.rails.rails));

  add("Bare board network", table(XLSX, ctx, ["part.refDes", "bbp.model", "bbp.pd", "bbp.rjcUsed", "bbp.rJTop", "bbp.rSurfaceAir", "bbp.rTop", "bbp.rJBoard", "bbp.footprint", "bbp.share",
    "bbp.rSpread", "bbp.rBoard", "bbp.qBoard", "bbp.qTop", "bbp.tj", "bbp.margin", "bbp.status"].map(t => [t, r => r]), refs.filter(r => calc.bareBoard.parts[r])));

  const plateRows = ["top", "bottom"].flatMap(s => (p.spreader?.[s]?.chips || []).map((_, i) => `${s}/${i}`));
  const solRows = ["top", "bottom"].filter(s => calc.solution[s]).flatMap(s => refs.filter(r => p.handCheck?.[s === "top" ? "solutionTop" : "solutionBottom"]?.parts?.[r]?.underPlate).map(r => `${s}|${r}`));
  const spr = table(XLSX, ctx, [["pchip.refDes", k => k, "RefDes"], [null, , "Plate", k => k.split("/")[0]], ["pchip.covered", k => k], ["pchip.heightMode", k => k], ["pchip.timMode", k => k],
    ["pchip.timTSet", k => k], ["pchip.timK", k => k], ["pchip.chipFace", k => k], ["pchip.gap", k => k], ["pchip.timT", k => k], ["pchip.pedestal", k => k], ["pchip.timR", k => k], ["pchip.check", k => k]], plateRows);
  XLSX.utils.sheet_add_aoa(spr, [[], ["Hand check with a plate"], ["RefDes and plate", ...["solp.rjc", "solp.rTim1", "solp.rLocal", "solp.rTim2", "solp.rPlateSpread", "solp.rStack", "solp.tj", "solp.margin", "solp.status"].map(t => `${TERMS[t].label}${TERMS[t].unit ? " (" + TERMS[t].unit + ")" : ""}`)],
    ...solRows.map(k => [k.replace("|", ", "), ...["solp.rjc", "solp.rTim1", "solp.rLocal", "solp.rTim2", "solp.rPlateSpread", "solp.rStack", "solp.tj", "solp.margin", "solp.status"].map(t => cell(valueOf(ctx, t, k)))])], { origin: -1 });
  add("Spreader", spr);

  const sim = table(XLSX, ctx, [["sim.refDes", r => r], ["sim.model", r => r], ["sim.pd", r => r], ["sim.rjc", r => r], ["sim.h", r => r], ["sim.area", r => r], ["sim.kz", r => r], ["sim.kxy", r => r], ["sim.check", r => r]],
    calc.simChips.map((_, i) => String(i)));
  const ss = p.simSetup || {};
  XLSX.utils.sheet_add_aoa(sim, [[], ["Study"], ["Study type", ss.studyType ?? ""], ["Global mesh size (mm)", cell(ss.meshGlobalMm)], ["Fine mesh size (mm)", cell(ss.meshFineMm)],
    ["Parts without heat in the study", cell(ss.includePartsWithoutHeat)], ["k of parts without heat (W/(m K))", cell(ss.kPartsWithoutHeat)], [],
    ["Surface", "h (W/(m² K))", "Emissivity"], ...Object.entries(ss.surfaces || {}).map(([g, x]) => [g, cell(x.h), cell(x.emissivity)]), [],
    ["Contact", "R (K m²/W)"], ...Object.entries(ss.contacts || {}).map(([g, x]) => [g, cell(x)])], { origin: -1 });
  add("Sim setup", sim);

  const resRows = runs.flatMap(r => Object.keys(r.parts || {}).map(ref => ({ run: r.name, ref })));
  const res = table(XLSX, ctx, [["run.name", r => r.run, "Run", r => r.run], ["part.refDes", r => r.ref, "RefDes", r => r.ref], ["part.tjMax", r => r.ref], ["runp.peak", r => `${r.run}|${r.ref}`], ["runp.avg", r => `${r.run}|${r.ref}`],
    ["runp.margin", r => `${r.run}|${r.ref}`], ["runp.status", r => `${r.run}|${r.ref}`], ["runp.hand", r => `${r.run}|${r.ref}`], ["runp.gap", r => `${r.run}|${r.ref}`]], resRows);
  XLSX.utils.sheet_add_aoa(res, [[], ["Runs", "Date", "Study", "Status", "Result", "Hottest part", "Peak (C)", "Energy balance error (%)", "PCB max (C)", "Spreader max (C)", "Inputs changed since"],
    ...runs.map(r => [r.name, r.date, r.studyType, r.status, ...["run.pass", "run.worst", "run.worstC", "run.energyBalance", "run.pcbMax", "run.spreaderMax", "run.changes"].map(t => cell(valueOf(ctx, t, r.name)))])], { origin: -1 });
  add("Results", res);

  const cmp = calc.compare?.ok;
  const cws = table(XLSX, ctx, [["part.refDes", r => r, "RefDes", r => r], ["cmp.hand", r => r], ["cmp.thermal", r => r], ["cmp.flow", r => r], ["cmp.meas", r => r], ["cmp.gapHand", r => r], ["cmp.gapThermal", r => r], ["cmp.gapFlow", r => r], ["cmp.flowVsThermal", r => r]], cmp?.refs || []);
  XLSX.utils.sheet_add_aoa(cws, [[], ["Thermal run", cmp?.thermal || "none"], ["Flow run", cmp?.flow || "none"], ["Measurement", cmp?.meas || "none"], ["Runs are the newest with health OK."]], { origin: -1 });
  add("Compare", cws);

  add("Changes", table(XLSX, ctx, [["change.date", x => x.id], ["change.macro", x => x.id], ["change.body", x => x.id], ["change.change", x => x.id], ["change.field", x => x.id],
    ["change.old", x => x.id], ["change.new", x => x.id], ["change.checked", x => x.id]], cad?.changes || []));

  // Method: every formula of the page, with the numbers of the first place it is used.
  const seen = new Set(), method = [["Value", "Unit", "Formula", "With the numbers of", "Note", "Reference"]];
  for (const [t, k] of instances(ctx)) {
    if (seen.has(t) || !FORMULAS[t]) continue;
    const f = formulaOf(ctx, t, k);
    if (!f) continue;
    seen.add(t);
    method.push([TERMS[t]?.label || t, TERMS[t]?.unit || "", f.text, k || "", f.note || "", f.ref || ""]);
  }
  for (const t of Object.keys(FORMULAS)) if (!seen.has(t) && TERMS[t]) method.push([TERMS[t].label, TERMS[t].unit || "", "(not used in this project)", "", "", ""]);
  const mws = XLSX.utils.aoa_to_sheet(method);
  mws["!cols"] = [{ wch: 30 }, { wch: 10 }, { wch: 90 }, { wch: 18 }, { wch: 60 }, { wch: 24 }];
  add("Method", mws);
  return wb;
}
