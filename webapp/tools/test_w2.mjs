// W2 tests: the calculation module gives the workbook's numbers.
//   node tools/test_w2.mjs              compare with tests/webapp/expected_*.json
//   node tools/test_w2.mjs --write      first refresh the expected files from the workbooks
// Numbers must agree to 6 significant digits, texts exactly. One deliberate change: the workbook worked per
// register row, the page per part (RefDes). For a part with several loads (RPi4 U3) the per row values are
// summed (Pd, share, density); its Tj est, margin and flag are the page's own and are only reported.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { convertFile } from "./convert_workbook.mjs";
import { expectedFromWorkbook } from "./expected_from_workbook.mjs";
import { makeValidator, validateFiles } from "./validate.mjs";
import { calcProject, forSolidWorks } from "../src/calc/index.js";
import { ranks } from "../src/calc/util.js";
import { loadPower } from "../src/calc/power.js";
import { spreaderPlate } from "../src/calc/spreader.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
const exDir = path.join(root, "example", "RaspberryPi4");
const expDir = path.join(root, "tests", "webapp");
const NOW = "2026-10-18T12:00:00";
const CASES = [["RaspberryPi4.xlsx", "expected_rpi4.json"], ["RPi4_fresh.xlsx", "expected_rpi4_fresh.json"]];

let fails = 0, compared = 0;
const failLines = [];
const check = (ok, what) => { console.log(`${ok ? "ok  " : "FAIL"} ${what}`); if (!ok) fails++; };
const same = (a, b) => {
  if (typeof b === "number") return typeof a === "number" && (Math.abs(a - b) <= 1e-6 * Math.max(Math.abs(a), Math.abs(b)) || Math.abs(a - b) < 1e-9);
  return a === b;
};
// Compare every key of `exp` with `act`; returns the list of differences.
function diff(label, act, exp, skip = []) {
  const out = [];
  for (const [k, e] of Object.entries(exp)) {
    if (skip.includes(k) || k === "row") continue;
    const a = act?.[k];
    if (e && typeof e === "object") { out.push(...diff(`${label}.${k}`, a, e, skip)); continue; }
    compared++;
    if (!same(a, e)) out.push(`${label}.${k}: page ${JSON.stringify(a)}, workbook ${JSON.stringify(e)}`);
  }
  return out;
}
const yn = b => (b ? "Yes" : "No");
function report(what, problems) {
  check(problems.length === 0, what + (problems.length ? `: ${problems.length} differ` : ""));
  for (const p of problems.slice(0, 12)) console.log("       " + p);
}

if (process.argv.includes("--write")) {
  for (const [xlsx, exp] of CASES) {
    if (!fs.existsSync(path.join(exDir, xlsx))) continue;
    fs.writeFileSync(path.join(expDir, exp), JSON.stringify(expectedFromWorkbook(path.join(exDir, xlsx)), null, 1) + "\n");
    console.log("     wrote " + exp);
  }
}

const ajv = makeValidator();
for (const [xlsx, expFile] of CASES) {
  const src = path.join(exDir, xlsx);
  if (!fs.existsSync(src)) { console.log(`skip ${xlsx} not here`); continue; }
  console.log(`\n== ${xlsx}`);
  const E = JSON.parse(fs.readFileSync(path.join(expDir, expFile), "utf8"));
  const conv = convertFile(src, { now: NOW, forSolidWorksFrom: "workbook" });
  const { project, cad, results } = conv;
  const C = calcProject(project, cad, results);

  // Register: per load, then per part.
  {
    const probs = [];
    const multi = new Set(project.parts.filter(p => p.loads.length > 1).map(p => p.refDes));
    for (const p of project.parts) {
      const rows = E.register.rows.filter(r => r.refDes === p.refDes);
      if (rows.length !== p.loads.length) { probs.push(`${p.refDes}: ${p.loads.length} loads, ${rows.length} rows`); continue; }
      p.loads.forEach((l, i) => {
        const L = C.register.loads[l.id], e = rows[i];
        const act = { method: L.method, pinW: L.pinW, poutW: L.poutW, pdCalc: L.pdCalcW ?? L.problem, pdW: L.pdW };
        probs.push(...diff(`${p.refDes} load ${i + 1}`, act, { method: e.method, pinW: e.pinW, poutW: e.poutW, pdCalc: e.pdCalc, pdW: e.pdW }));
      });
      const P = C.register.parts[p.refDes];
      if (!multi.has(p.refDes)) probs.push(...diff(p.refDes, P, rows[0], ["refDes", "method", "pinW", "poutW", "pdCalc", "rankPd", "rankDensity"]));
      else {
        const s = k => rows.reduce((a, r) => a + (r[k] ?? 0), 0);
        probs.push(...diff(p.refDes, P, { pdW: s("pdW"), areaMm2: rows[0].areaMm2, densityWcm2: s("densityWcm2"), share: s("share") }));
        console.log(`     ${p.refDes}: ${rows.length} rows became one part. Tj est per part ${P.tjEstC?.toFixed(2)} C ` +
          `(rows: ${rows.map(r => r.tjEstC?.toFixed(2)).join(", ")}), flag ${P.flag}`);
      }
    }
    // Ranks: the workbook's rule applied to its own per part sums.
    const per = k => project.parts.map(p => { const rs = E.register.rows.filter(r => r.refDes === p.refDes); return rs.some(r => typeof r[k] === "number") ? rs.reduce((a, r) => a + (r[k] ?? 0), 0) : undefined; });
    const rp = ranks(per("pdW")), rd = ranks(per("densityWcm2"));
    project.parts.forEach((p, i) => {
      probs.push(...diff(p.refDes, C.register.parts[p.refDes], { rankPd: rp[i], rankDensity: rd[i] }));
    });
    probs.push(...diff("register", C.register, { totalPdW: E.register.totalPdW }));
    report("register: Pd per load, Tj est, margin, density, share, flag, why, ranks", probs);
  }

  // Triage summary.
  {
    const t = C.triage;
    const probs = diff("triage", t, { totalPdW: E.triage.totalPdW, noRja: E.triage.noRja, noBody: E.triage.noBody });
    if (t.critical + t.watch + t.minor + t.noPd !== t.parts) probs.push("flag counts do not add up to the parts");
    report("triage: totals, parts with no Rja or body", probs);
  }

  // Rail budget (per load, as the rows were).
  {
    const probs = [];
    E.rails.rails.forEach((e, i) => {
      const a = C.rails.rails[i] || {};
      probs.push(...diff(`rail ${e.rail}`, { rail: a.rail, rows: a.loads, pdW: a.pdW, share: a.share, noPd: a.noPd }, e));
    });
    if (C.rails.rails.length !== E.rails.rails.length) probs.push(`${C.rails.rails.length} rails, workbook ${E.rails.rails.length}`);
    const { rails, rowsNoRail, ...sumE } = E.rails;
    probs.push(...diff("rails", { ...C.rails, rowsNoRail: C.rails.loadsNoRail }, { ...sumE, rowsNoRail }));
    report("rail budget", probs);
  }

  // PCB and body table.
  report("PCB size and stack k", diff("pcb", C.pcb, E.pcb));
  {
    const probs = [];
    const { bodies: eb, ...sumE } = E.geometry;
    probs.push(...diff("geometry", C.geometry, sumE));
    const cutIndex = side => { const m = {}; (C.spreader[side]?.cuts || []).forEach((c, i) => { m[c.key] = i + 1; }); return m; };
    const ct = cutIndex("top"), cb = cutIndex("bottom");
    for (const [key, e] of Object.entries(eb)) {
      const row = C.bodies.find(b => b.key === key), g = C.geometry.bodies.find(b => b.key === key);
      if (!row) { probs.push(`body ${key} missing`); continue; }
      const inc = project.bodies[key]?.include;
      probs.push(...diff(`body ${key.slice(0, 24)}`, { name: row.name, refDes: row.refDes, include: inc === undefined ? undefined : yn(inc),
        ...row.box, ...g, cutTop: ct[key], cutBottom: cb[key] }, e));
    }
    if (C.bodies.filter(b => b.name).length !== Object.keys(eb).length) probs.push(`${C.bodies.length} bodies, workbook ${Object.keys(eb).length}`);
    report("CAD geometry: used boxes, centres, Pd per body, checks, cut numbers", probs);
  }

  // Bare board network.
  {
    const B = C.bareBoard;
    const probs = diff("board", B.board, E.bareBoard.board);
    for (const [ref, e] of Object.entries(E.bareBoard.parts)) {
      const a = B.parts[ref];
      if (!a) { probs.push(`${ref} missing`); continue; }
      probs.push(...diff(ref, { ...a, ...a.spread, simC: B.compare.parts?.[ref]?.simC, gapC: B.compare.parts?.[ref]?.gapC }, e));
    }
    probs.push(...diff("compare", B.compare, E.bareBoard.compare));
    report("bare board network: resistances, heat split, Tj hand, status, sim compare", probs);
  }

  // Spreader plates.
  for (const side of ["top", "bottom"]) {
    const e = E.spreader[side];
    if (!e) continue;
    const a = C.spreader[side];
    if (!a) { check(false, `spreader ${side}: not computed`); continue; }
    const probs = diff(side, { ...a, cutCount: a.cuts.length }, { faceZMm: e.faceZMm, covered: e.covered, cutCount: e.cutCount });
    for (const [ref, ec] of Object.entries(e.chips)) {
      const ac = a.chips.find(c => c.refDes === ref);
      probs.push(...diff(`${side} ${ref}`, { ...ac, covered: yn(ac?.covered) }, ec));
    }
    e.cuts.forEach((ec, i) => probs.push(...diff(`${side} cut ${i + 1}`, a.cuts[i], ec)));
    report(`spreader ${side}: plate face, gaps, TIM, pedestals, checks, cuts`, probs);
  }

  // Thermal solution.
  for (const side of ["top", "bottom"]) {
    const e = E.solution[side];
    if (!e) continue;
    const a = C.solution[side];
    if (!a) { check(false, `thermal solution ${side}: not computed`); continue; }
    const { rows, ...sumE } = e;
    const probs = diff(side, a, sumE);
    for (const p of project.parts) {
      const rs = rows.filter(r => r.refDes === p.refDes);
      const ap = a.parts[p.refDes];
      if (rs.length === 1) probs.push(...diff(`${side} ${p.refDes}`, { ...ap, pdW: C.register.parts[p.refDes].pdW }, rs[0], ["refDes"]));
      else if (rs.every(r => r.status === "not under plate")) probs.push(...diff(`${side} ${p.refDes}`, ap, { status: "not under plate" }));
      else console.log(`     ${side} ${p.refDes}: several rows under the plate; not compared`);
    }
    report(`thermal solution ${side}: stack resistances, plate, Tj, status`, probs);
  }

  // Sim setup per chip.
  {
    const probs = [];
    for (const [ref, e] of Object.entries(E.simChips)) probs.push(...diff(ref, C.simChips.find(c => c.refDes === ref), e, e.kxy === undefined ? ["kxy"] : []));
    report("sim setup: k z per chip", probs);
  }

  // forSolidWorks from the module = the workbook's grey cells the macros read today (flow is page only: the workbook has no Flow block).
  {
    const { inputsHash: _h, flow: _f, ...fsCalc0 } = forSolidWorks(project, C, NOW) || {};
    const fsCalc = forSolidWorks(project, C, NOW) && fsCalc0;
    const fsBook = project.forSolidWorks;
    if (!fsBook) check(!fsCalc, "forSolidWorks: none, as in the workbook");
    else {
      const { computedAt, computedBy, inputsHash, ...book } = fsBook;
      report("forSolidWorks: same as the workbook (bodies, heat, chips, plates)", [...diff("fs", fsCalc, book), ...diff("fs(extra)", book, fsCalc ? (({ computedAt, computedBy, inputsHash, ...x }) => x)(fsCalc) : {})]);
      const files = { project: { ...project, forSolidWorks: fsCalc }, cad, results };
      const v = validateFiles(files, ajv);
      check(v.length === 0, "project.json with the computed forSolidWorks validates" + (v.length ? ": " + v.slice(0, 3).join("; ") : ""));
    }
  }
}

// Cases the RPi4 workbooks do not reach, worked by hand from the template formulas.
{
  console.log("\n== Hand worked cases");
  const lp = (part, load) => loadPower({ class: "Other", ...part }, load);
  const probs = [
    ...diff("LDO", lp({}, { method: "LDO / linear reg", vinV: 5, voutV: 3.3, ioutA: 1 }), { pdW: 1.7 }),
    ...diff("LDO Pin-Pout", lp({}, { method: "LDO / linear reg", vinV: 5, iinA: 1.1, voutV: 3.3, ioutA: 1 }), { pdW: 2.2 }),
    ...diff("LDO missing", lp({}, { method: "LDO / linear reg", vinV: 5 }), { problem: "missing input" }),
    ...diff("buck eta 0", lp({}, { method: "Switching regulator", voutV: 1, ioutA: 1, efficiency: 0 }), { problem: "missing input" }),
    ...diff("I2R from Iout", lp({ class: "Power inductor" }, { ioutA: 2, rOhm: 0.05 }), { method: "Resistive (I squared R)", pdW: 0.2 }),
    ...diff("manual", lp({}, {}), { method: "Manual entry", problem: "enter Pd meas" }),
    ...diff("measured wins", lp({ class: "Memory" }, { vinV: 1, iinA: 1, pdMeasuredW: 0.4 }), { pdCalcW: 1, pdW: 0.4 }),
    ...diff("no class", loadPower({}, { vinV: 1, iinA: 1 }), { problem: "pick class" }),
  ];
  const rows = [{ key: "a", refDes: "A", include: true, side: "Top", box: { xMin: 0, yMin: 0, zMin: 1.6, l: 10, w: 10, h: 2 } },
    { key: "b", refDes: "B", include: true, side: "Top", box: { xMin: 20, yMin: 0, zMin: 1.6, l: 5, w: 5, h: 1 } }];
  const plate = { minPedestalMm: 0.3, clearanceXYMm: 1, clearanceZMm: 1, heightMode: "Manual", heightManualMm: 4,
    chips: [{ refDes: "A", covered: true, heightMode: "Thick TIM", timMode: "Body", timTMm: 0.2, timK: 4 },
      { refDes: "B", covered: true, timMode: "Contact R", timTMm: 0.2, timK: 4 }] };
  const sp = spreaderPlate(plate, "top", rows);
  probs.push(...diff("thick TIM", sp.chips[0], { gapMm: 0.4, timTMm: 0.4, pedestalMm: 0, timRCW: 1, check: "OK" }));
  probs.push(...diff("contact R", sp.chips[1], { gapMm: 1.4, timTMm: 0.2, pedestalMm: 1.2, check: "OK" }));
  probs.push(...diff("manual face", sp, { faceZMm: 4 }));
  report("power methods, Thick TIM, Contact R, manual plate height", probs);
}

console.log(`\n${compared} values compared`);
console.log(fails ? `${fails} failed` : "All passed");
process.exit(fails ? 1 : 0);
