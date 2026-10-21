// W7 test: tabs 6 to 10, the SolidWorks cards, and the files the remaining macros write.
//   node tools/test_w7.mjs
// 1. The SolidWorks cards: the checks of each macro and the marker the page waits on.
// 2. Spreader, plate hand check, sim chips and sizes from SolidWorks: the editors' model functions.
// 3. Out of date: the inputs hash Build Board copies, and what changed since a run.
// 4. What the macros write in project folder mode (Build Spreader, Sync CAD, Run Thermal Study, Read Results,
//    Run Flow), made here the way the .bas files make it, is valid against the schemas and shows on tabs 6 to 10.
// 5. Tab states: build out of date, results PASS and FAIL, sim inputs changed.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { h } from "preact";
import { render } from "preact-render-to-string";
import { makeValidator } from "./validate.mjs";
import { calcProject, forSolidWorks } from "../src/calc/index.js";
import { valueOf } from "../src/calc/formulas.js";
import { inputsHash, inputChanges } from "../src/calc/runs.js";
import { addPlate, removePlate, addSolution, addSimChips, takeAsBuilt, projectForSave, setIn, heatedOnSide } from "../src/page/model.js";
import { macroChecks, marker, MACROS } from "../src/page/swcard.js";
import { tabStates, blockingProblems } from "../src/page/tabs.js";
import { Explain } from "../src/page/explain.js";
import * as Views from "../src/page/views.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const ex = path.join(here, "..", "..", "example", "RaspberryPi4");
const read = f => JSON.parse(fs.readFileSync(path.join(ex, f), "utf8"));
const ajv = makeValidator();
let fails = 0, n = 0;
const check = (ok, what, extra) => { n++; if (!ok) { fails++; console.log("FAIL " + what + (extra !== undefined ? ": " + JSON.stringify(extra).slice(0, 400) : "")); } };
const valid = (schema, data, what) => { const ok = ajv.validate(schema, data); check(ok, what + " is valid", ajv.errors?.slice(0, 3)); };
const NOW = "2026-10-19T12:00:00";
const P = read("project.json"), C = read("cad.json"), R = read("results.json");
const calcOf = (p, c = C, r = R) => calcProject(p, c, r);
const ctxOf = (p, c = C, r = R, extra = {}) => {
  const calc = calcOf(p, c, r);
  return { project: p, cad: c, results: r, calc, states: tabStates({ project: p, cad: c, results: r, calc, canSave: true }), folder: "RPi4", files: { project: "AB_project.json", cad: "AB_cad.json", results: "AB_results.json" },
    edit() {}, editFn() {}, go() {}, locked: false, learn: true, now: () => NOW, blocking: blockingProblems(p, calc), value: (t, k) => valueOf({ project: p, cad: c, results: r, calc }, t, k),
    dirty: false, mode: "folder", waits: {}, setWait() {}, icons: {}, ...extra };
};
const show = (tab, c) => render(h(Explain.Provider, { value: { ctx: { project: c.project, cad: c.cad, results: c.results, calc: c.calc }, open() {}, showFx: false } }, h(() => Views[tab](c))));

// ---------- 1. SolidWorks cards ----------
{
  const c = ctxOf(P);
  for (const id of Object.keys(MACROS)) {
    const ch = macroChecks(id, c);
    check(Array.isArray(ch) && ch.length > 0 && ch.every(x => typeof x.text === "string" && x.text && typeof x.ok === "boolean"), `${id}: checks listed`, ch);
  }
  check(macroChecks("BuildBoard", c).every(x => x.ok || x.soft), "RPi4 example: Build Board checks met", macroChecks("BuildBoard", c).filter(x => !x.ok));
  check(macroChecks("BuildBoard", { ...c, dirty: true }).some(x => !x.ok && /Save/.test(x.text)), "unsaved edits: Build Board says Save first");
  const cOpen = ctxOf({ ...P, checks: {} });
  check(macroChecks("BuildBoard", cOpen).some(x => !x.ok && /Tick the 15 changes/.test(x.text) && x.fix === 10), "unticked changes: Build Board names them and links tab 10");
  check(macroChecks("SetupSim", c).some(x => !x.ok && /Study type on this tab: Flow/.test(x.text)), "RPi4 is a Flow study: Run Thermal Study says the type is wrong");
  check(macroChecks("RunFlow", c).find(x => /Study type/.test(x.text))?.ok === true, "Run Flow Simulation accepts the Flow study type");
  check(marker("BuildBoard", C, R) === C.build.date && marker("BuildBoard", { ...C, build: { ...C.build, date: "2026-10-19T13:00:00" } }, R) !== C.build.date, "Build Board marker follows the build date");
  const r2 = { ...R, runs: [...R.runs, { name: "PT_x", date: NOW, studyType: "Thermal", status: "solved" }] };
  check(marker("SetupSim", C, R) !== marker("SetupSim", C, r2), "a new run changes the study marker");
  check(marker("ReadResults", C, r2) !== marker("ReadResults", C, { ...r2, runs: r2.runs.map(r => r.name === "PT_x" ? { ...r, status: "read" } : r) }), "a run read changes the Read Results marker");
  check(macroChecks("ReadResults", ctxOf(P, C, r2)).some(x => x.ok && /PT_x/.test(x.text)), "Read Results names the run waiting to be read");
  const out = show("BuildTab", c);
  check(out.includes("In SolidWorks: 3 Build Board") && out.includes("Before you press it") && out.includes("I pressed it"), "tab 6 has the Build Board card");
  const waiting = show("BuildTab", { ...c, waits: { BuildBoard: { since: NOW, marker: C.build.date } } });
  check(waiting.includes("Waiting for Build Board"), "the card waits after I pressed it");
  check(show("ImportTab", c).includes("In SolidWorks: 2 Import STEP"), "tab 2 has the Import STEP card");
}

// ---------- 2. editors ----------
{
  let calc = calcOf(P);
  const top = heatedOnSide(P, calc, "top");
  check(top.includes("U1") && !top.includes("J1") === !(calc.register.parts.J1?.pdW > 0), "heated parts on the top side", top);
  let p = removePlate(P, "top");
  check(!p.spreader, "removing the only plate removes the spreader section");
  p = addPlate(p, calcOf(p), "top");
  check(p.spreader.top.chips.length === top.length && p.spreader.top.chips.every(x => x.covered === false), "a new plate lists the heated parts, none covered");
  p = setIn(p, ["spreader", "top", "chips", 0, "covered"], true);
  let saved = projectForSave(p, C, R, NOW);
  valid("project.schema.json", saved, "project with a new plate");
  check(saved.forSolidWorks.plates?.top?.chips.length === 1, "one covered chip reaches forSolidWorks", saved.forSolidWorks.plates);
  p = addSolution(removePlate(P, "top"), calcOf(P), "bottom");
  calc = calcOf(p);
  check(!!calc.solution.bottom, "a bottom plate hand check is worked out");
  valid("project.schema.json", projectForSave(p, C, R, NOW), "project with a new plate hand check");
  const noChips = setIn(P, ["simSetup", "chips"], []);
  const withChips = addSimChips(noChips, calcOf(noChips));
  check(withChips.simSetup.chips.length >= 6 && withChips.simSetup.chips.every(x => x.model === "Block" || x.model === "Two-resistor"), "sim chips added for the heated parts", withChips.simSetup.chips);
  valid("project.schema.json", projectForSave(withChips, C, R, NOW), "project with added sim chips");
  const u1 = calc.bodies.find(b => b.refDes === "U1");
  const moved = takeAsBuilt(P, calcOf(P), u1.key, { ...u1.box, xMin: u1.box.xMin + 1, h: u1.box.h });
  check(moved.bodies[u1.key].override?.xMin === Math.round((u1.box.xMin + 1) * 1000) / 1000 && moved.bodies[u1.key].override?.h === undefined, "taking a SolidWorks size writes only the field that differs", moved.bodies[u1.key]);
  const ctx = ctxOf(P);
  const sp = show("SpreaderTab", ctx);
  check(sp.includes("Top plate") && sp.includes('data-term="pchip.timMode"') && sp.includes("Hand check with a top plate") && sp.includes('data-term="solp.tim1T"') && sp.includes("In SolidWorks: 4 Build Spreader") && sp.includes("In SolidWorks: 5 Sync CAD"), "tab 7 has the plate editor, the hand check with a plate and both cards");
  const sim = show("SimTab", ctx);
  check(sim.includes('data-term="simset.studyType"') && sim.includes('data-term="surf.spreaderTopH"') && sim.includes('data-term="simset.contactPartPcb"') && sim.includes("In SolidWorks: 7 Run Flow Simulation"), "tab 8 has the study, surfaces, contacts, chips and the Flow card");
}

// ---------- 3. out of date ----------
{
  const fs0 = forSolidWorks(P, calcOf(P), NOW);
  check(/^[0-9a-f]{8}$/.test(fs0.inputsHash) && fs0.inputsHash === inputsHash(fs0), "forSolidWorks carries the inputs hash");
  const cBuilt = { ...C, build: { ...C.build, inputsHash: fs0.inputsHash } };
  check(calcOf(P, cBuilt).build.outOfDate === false, "built from the inputs of now: up to date");
  const u1 = P.bodies[Object.keys(P.bodies).find(k => P.bodies[k].refDes === "U1")];
  const key = Object.keys(P.bodies).find(k => P.bodies[k].refDes === "U1");
  const p2 = setIn(P, ["bodies", key, "override", "h"], 3);
  check(calcOf(p2, cBuilt).build.outOfDate === true, "a body resized after the build: out of date");
  const p3 = setIn(P, ["board", "notes"], "changed notes");
  check(calcOf(p3, cBuilt).build.outOfDate === false, "a change that is not built (notes) keeps the model up to date");
  const st = Object.fromEntries(tabStates({ project: p2, cad: cBuilt, results: R, calc: calcOf(p2, cBuilt), canSave: true }).map(s => [s.id, s]));
  check(st.build.status === "attention" && /out of date/.test(st.build.note), "tab 6 says out of date", st.build);
  const fsNow = forSolidWorks(setIn(P, ["board", "ambientC"], 45), calcOf(setIn(P, ["board", "ambientC"], 45)), NOW);
  const ch = inputChanges(fs0, fsNow, P.simSetup, setIn(P, ["simSetup", "meshGlobalMm"], 2).simSetup);
  check(ch.some(x => /ambient 40 C, now 45 C/.test(x)) && ch.some(x => /mesh/.test(x)), "what changed since a run is named", ch);
  check(inputChanges(fs0, fs0, P.simSetup, P.simSetup).length === 0, "nothing changed: empty list");
  // A run's inputs come back from the macro with 15 digits: that is not a change.
  const chip = { refDes: "U1", key: "k", timTMm: 0.2, pedestalMm: 0, timMode: "Body", timK: 3 };
  const plate = chips => ({ faceZMm: 4.1, thicknessMm: 1.5, k: 167, chips });
  const withPlate = (fs, chips) => ({ ...fs, plates: { top: plate(chips) } });
  check(inputChanges(withPlate(fs0, [{ ...chip, timRCW: 0.296296296296296 }]), withPlate(fs0, [{ ...chip, timRCW: 0.2962962962962959 }])).length === 0,
    "plate chips written back with 15 digits: no change");
  check(inputChanges(withPlate(fs0, [{ ...chip, timRCW: 0.3 }]), withPlate(fs0, [{ ...chip, timRCW: 0.6, timTMm: 0.4 }])).some(x => /chips, TIMs/.test(x)),
    "a thicker TIM is a change");
}

// ---------- 4. files the macros write ----------
{
  const fs0 = forSolidWorks(P, calcOf(P), NOW);
  // Build Spreader: cad spreaderBuild, as BuildSpreader.bas SideRecord makes it.
  const cad = structuredClone(C);
  cad.build.inputsHash = fs0.inputsHash;
  cad.spreaderBuild = { top: { date: NOW, file: "C:\\b\\CAD\\AB_Spreader_r01.sldprt", pedestals: 6, tims: ["AB_TIM_U1.sldprt", "AB_TIM_U2.sldprt"], cuts: 5, failures: [], report: "C:\\b\\CAD\\AB_BuildSpreader_report.txt" } };
  // Sync CAD: asBuilt with a body that differs, the plate, failed checks; a change record.
  const u1 = calcOf(P).bodies.find(b => b.refDes === "U1");
  cad.asBuilt = { date: NOW, plates: { top: { outlineAreaMm2: 2798.2, topAreaMm2: 2803, exposedAreaMm2: 5989.7, volumeMm3: 4744.3, massG: 12.81, thicknessMm: 1.5, holes: 1, chips: [{ refDes: "U1", pedestalMm: 0, timTMm: 0.2 }] } },
    bodies: [{ key: u1.key, box: { ...u1.box, xMin: u1.box.xMin + 0.5 } }], failedChecks: ["TIM U2: The TIM part for U2 is suppressed."] };
  cad.changes = [{ id: "20261019T120000-1", date: NOW, macro: "SyncFromCAD", key: u1.key, name: u1.name, refDes: "U1", change: "differs in SolidWorks", field: "xMin", old: u1.box.xMin, new: u1.box.xMin + 0.5, delta: 0.5 }, ...cad.changes];
  valid("cad.schema.json", cad, "cad.json after Build Spreader and Sync CAD");
  // Run Thermal Study then Read Results: a run with the inputs copy; Run Flow: a Flow run.
  const res = structuredClone(R);
  const saved = projectForSave(setIn(P, ["simSetup", "studyType"], "Thermal"), cad, res, NOW);
  const peaks = { U1: 71.2, U2: 62.4, U3: 66.1, U4: 58.9, U5: 63.3, U6: 55.5, L1: 60.2, L2: 59.1 };
  res.runs.push({ name: "PT_20261019_1200", date: NOW, studyType: "Thermal", study: "PT_20261019_1200", status: "read", assembly: "C:\\b\\CAD\\AB_Thermal.sldasm",
    spreaderFiles: ["AB_Spreader_r01.sldprt"], mesh: { globalMm: 3, fineMm: 0.5, nodes: 120000, elements: 80000 },
    inputs: { forSolidWorks: saved.forSolidWorks, simSetup: saved.simSetup }, report: "C:\\b\\CAD\\AB_SetupSim_PT_20261019_1200.txt",
    spreader: { maxC: 58.2, minC: 52.1 }, pcb: { maxC: 61, minC: 47.3 }, energyBalancePct: 0.4,
    parts: Object.fromEntries(Object.entries(peaks).map(([k, v]) => [k, { peakC: v, avgC: v - 3 }])), plotFile: "C:\\b\\CAD\\results\\PT_20261019_1200.png", notes: ["Heat flow, whole model: in 6.9 W"] });
  res.runs.push({ name: "FL_20261019_1300", date: NOW, studyType: "Flow", study: "Flow PT_Flow", status: "read", assembly: "C:\\b\\CAD\\AB_Thermal.sldasm",
    inputs: { forSolidWorks: saved.forSolidWorks, simSetup: saved.simSetup }, pcb: { maxC: 70, minC: 50 }, parts: { U1: { peakC: 130, avgC: 120 } } });
  valid("results.schema.json", res, "results.json after Run Thermal Study, Read Results and Run Flow");
  const p = { ...saved, checks: { ...saved.checks, "20261019T120000-1": { at: NOW } } };
  const c = ctxOf(p, cad, res);
  const rc = c.calc.runs.PT_20261019_1200;
  check(rc.pass === "PASS" && rc.worst === "U1" && rc.balanceOk === true && rc.withPlate, "thermal run: PASS, hottest part, balance OK, with a plate", rc);
  check(rc.changes.length === 0, "the run made from the saved inputs is up to date", rc.changes);
  check(rc.parts.U1.handFrom === "thermal solution" && Number.isFinite(rc.parts.U1.gapC), "U1 is compared with the hand check with the plate", rc.parts.U1);
  check(c.calc.runs.FL_20261019_1300.pass === "FAIL", "Flow run with U1 above Tj max: FAIL");
  const res9 = show("ResultsTab", c);
  check(res9.includes("Run FL_20261019_1300") && res9.includes('data-term="runp.gap"') && res9.includes("All runs"), "tab 9 shows the newest run, hand vs simulation and all runs");
  const t10 = show("ChangesTab", c);
  check(t10.includes("Sizes changed in SolidWorks") && t10.includes("Take all 1"), "tab 10 offers the SolidWorks sizes");
  const t7 = show("SpreaderTab", c);
  check(t7.includes("Top plate as built") && t7.includes("Read back from the assembly") && t7.includes('data-term="asbp.mass"'), "tab 7 shows the plate as built and read back");
  const st = Object.fromEntries(c.states.map(s => [s.id, s]));
  check(st.results.status === "attention" && /over Tj max/.test(st.results.note), "tab 9 needs attention when the newest run FAILs", st.results);
  check(st.spreader.status === "done", "tab 7 done after Build Spreader", st.spreader);
  const pass = { ...res, runs: res.runs.slice(0, -1) };
  const st2 = Object.fromEntries(tabStates({ project: p, cad, results: pass, calc: calcOf(p, cad, pass), canSave: true }).map(s => [s.id, s]));
  check(st2.results.status === "done" && /PASS/.test(st2.results.note), "tab 9 done when the newest run PASSes", st2.results);
  const p4 = setIn(p, ["board", "ambientC"], 45);
  const st3 = Object.fromEntries(tabStates({ project: p4, cad, results: pass, calc: calcOf(p4, cad, pass), canSave: true }).map(s => [s.id, s]));
  check(st3.sim.status === "attention" && /inputs changed/.test(st3.sim.note), "tab 8 says the inputs changed after the run", st3.sim);
}

console.log(`W7: ${n - fails} of ${n} checks passed.`);
if (fails) process.exit(1);
