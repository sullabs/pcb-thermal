// Flow plan test (docs/FLOW_PLAN.md, phases F1 to F8): what RunFlow writes and what the page makes of it.
//   node tools/test_flow.mjs
// F1. Run health: the fields RunFlow writes are valid, runHealth gives ok / check / bad with plain reasons, tab 9
//     shows the badge and greys out PASS/FAIL of a bad run, the report and the tab state name it.
// F2. Mesh: simSetup.flow is valid, the smallest cell estimate matches RunFlow's rule and flags thin parts, tab 8
//     shows the Flow mesh card, a mesh check pair is compared per chip on tab 9.
// F3. Convergence: minTravels and goalCriteriaC are valid, default to 2 and 0.2, and drive the health reasons.
// F4. Tj basis: k steps and material names, forSolidWorks.flow per chip, two-resistor layers, tjBasis on tab 9.
// F5. PCB, TIM, spreader materials, radiative surfaces and contacts from the page, with the simplifications noted.
// F6. Environment: gravity from the orientation, pressure from altitude, forced air, domain margins, hand check note.
// F8. Heat with no body: listed on tab 6, accepted with a tick, open ones in forSolidWorks.openChecks (macros refuse).
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { h } from "preact";
import { render } from "preact-render-to-string";
import { makeValidator } from "./validate.mjs";
import { calcProject } from "../src/calc/index.js";
import { valueOf } from "../src/calc/formulas.js";
import { runHealth } from "../src/calc/runs.js";
import { flowSettings, kStep, flowMaterial, tjBasisOf, flowEmissivity, flowEnvironment, flowDomainMm } from "../src/calc/flow.js";
import { tabStates, blockingProblems } from "../src/page/tabs.js";
import { macroChecks } from "../src/page/swcard.js";
import { openIssues } from "../src/page/report.js";
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
const ctxOf = (p, c = C, r = R, extra = {}) => {
  const calc = calcProject(p, c, r);
  return { project: p, cad: c, results: r, calc, states: tabStates({ project: p, cad: c, results: r, calc, canSave: true }), folder: "RPi4", files: { project: "AB_project.json", cad: "AB_cad.json", results: "AB_results.json" },
    edit() {}, editFn() {}, go() {}, locked: false, learn: true, now: () => NOW, blocking: blockingProblems(p, calc), value: (t, k) => valueOf({ project: p, cad: c, results: r, calc }, t, k),
    dirty: false, mode: "folder", waits: {}, setWait() {}, icons: {}, ...extra };
};
const show = (tab, c) => render(h(Explain.Provider, { value: { ctx: { project: c.project, cad: c.cad, results: c.results, calc: c.calc }, open() {}, showFx: false } }, h(() => Views[tab](c))));

// A Flow run as RunFlow writes it since F1, from the numbers of the RPi4 run FL_20261001_0753: L1 and L2 unresolved,
// Serv Heatf 2.817 W + Radiation Energy Delta 3.604 W = 6.421 W out against 6.42 W applied of 6.67 W asked for.
const heat = { ...P.forSolidWorks.heat };
const flowRun = (over = {}) => ({
  name: "FL_20261019_2100", date: "2026-10-19T21:00:00", studyType: "Flow", study: "Flow RPi4_Flow", status: "read",
  inputs: { forSolidWorks: { ...P.forSolidWorks, heat }, simSetup: P.simSetup },
  heatRequestedW: 6.67, heatAppliedW: 6.42, energyBalancePct: Math.abs(6.42 - (2.817 + 3.604)) / 6.42 * 100,
  solver: { finished: true, iterations: 140, travels: 1.006, cellsTotal: 1429713, cellsSolid: 1093103, cellsFluid: 336610, projectDir: "C:\\RPi4\\CAD\\1" },
  goalsHealth: { "U1 max": { progressPct: 100, deltaC: 0.18, criteriaC: 2.15 } },
  warnings: ["Unresolved volume source: L1 heat", "Unresolved volume source: L2 heat"],
  unresolved: ["L1", "L2"],
  parts: { U1: { peakC: 91.6, avgC: 90.6 } },
  ...over,
});

// ---------- F1 ----------
{
  const run = flowRun();
  valid("results.schema.json", { ...R, runs: [...R.runs, run] }, "results.json with the F1 health fields");
  const hl = runHealth(run);
  check(hl.level === "bad", "the FL_20261001_0753 case is bad", hl);
  const lw = (heat.L1 || 0) + (heat.L2 || 0);
  check(/L1 and L2 were not meshed/.test(hl.reasons[0]) && hl.reasons[0].includes(`${lw.toFixed(2)} W left out`), "the reason names L1 and L2 and the heat left out", hl.reasons);
  check(!hl.reasons.some(r => /Flow warnings/.test(r)), "the unresolved warnings are not repeated as other warnings", hl.reasons);

  const ok = runHealth(flowRun({ unresolved: [], warnings: [], heatAppliedW: 6.67, energyBalancePct: 0.1, solver: { finished: true, travels: 2.1 } }));
  check(ok.level === "ok" && ok.reasons.length === 0, "finished, all meshed, balanced: ok", ok);
  const unfinished = runHealth(flowRun({ unresolved: [], warnings: [], solver: { finished: false, travels: 0.4 } }));
  check(unfinished.level === "bad" && unfinished.reasons.some(r => /did not finish/.test(r)) && unfinished.reasons.some(r => /fewer than 1/.test(r)), "not finished: bad, and too few travels named", unfinished);
  const offBalance = runHealth(flowRun({ unresolved: [], warnings: [], heatAppliedW: 6.67, energyBalancePct: 3.8 }));
  check(offBalance.level === "bad" && /energy balance error 3.80 %/.test(offBalance.reasons[0]), "balance above 2 %: bad", offBalance);
  const slow = runHealth(flowRun({ unresolved: [], warnings: [], heatAppliedW: 6.67, energyBalancePct: 0.1, goalsHealth: { "U3 max": { progressPct: 80, deltaC: 1, criteriaC: 0.5 } } }));
  check(slow.level === "check" && /goals not converged: U3 max/.test(slow.reasons[0]), "a goal below 100 %: check", slow);
  const leftOut = runHealth(flowRun({ unresolved: [], warnings: [], heatAppliedW: 6.0, energyBalancePct: 0.1 }));
  check(leftOut.level === "check" && /0.67 W of the 6.67 W/.test(leftOut.reasons[0]), "heat without a source: check", leftOut);
  const old = runHealth({ name: "FL_old", studyType: "Flow", status: "read" });
  check(old.level === "check" && /read it again/.test(old.reasons[0]), "a Flow run from before F1: check, read it again", old);
  check(runHealth({ studyType: "Thermal", energyBalancePct: 0 }).level === "ok", "Thermal run balanced: ok");
  check(runHealth({ studyType: "Thermal", energyBalancePct: 5 }).level === "bad", "Thermal run off balance: bad");

  const res = { ...R, runs: [...R.runs, run] };
  const c = ctxOf(P, C, res);
  check(c.value("run.health", run.name) === "Bad" && /L1 and L2/.test(c.value("run.healthReasons", run.name)), "health values on the page");
  check(c.value("run.cellsSolid", run.name) === 1093103 && c.value("run.unresolved", run.name) === "L1, L2", "solver values on the page");
  const t9 = show("ResultsTab", c);
  check(t9.includes('data-term="run.health"') && t9.includes("L1 and L2 were not meshed") && t9.includes("PASS or FAIL is greyed out"), "tab 9 shows the badge and the reasons");
  check(/class="cols dim"/.test(t9), "tab 9 greys out the verdict of a bad run");
  const st = Object.fromEntries(c.states.map(s => [s.id, s]));
  check(st.results.status === "attention" && /health Bad/.test(st.results.note), "tab 9 needs attention when the newest run is bad", st.results);
  const issues = openIssues(c);
  check(issues.some(x => /health Bad: L1 and L2/.test(x)), "the report lists the bad run as an open issue", issues);
}

// ---------- F2 ----------
{
  const withFlow = f => ({ ...P, simSetup: { ...P.simSetup, studyType: "Flow", flow: f } });
  valid("project.schema.json", withFlow({ meshLevel: 4, localMesh: true, solidRefine: 3, meshCheck: true, meshCheckLimitC: 1 }), "project.json with simSetup.flow");
  check(!ajv.validate("project.schema.json", withFlow({ meshLevel: 8 })), "mesh level 8 is refused by the schema");
  const c0 = ctxOf(withFlow({}));
  const s = c0.calc.flow.settings;
  check(s.meshLevel === 4 && s.localMesh === true && s.solidRefine === 3 && s.meshCheckLimitC === 1, "blank Flow settings take the defaults", s);
  const d = c0.calc.flow.domain, vol = (d.x[1] - d.x[0]) * (d.y[1] - d.y[0]) * (d.z[1] - d.z[0]);
  const expect = Math.cbrt(vol / (28000 * 7)) / 8;
  check(Math.abs(c0.calc.flow.minCellMm - expect) < 1e-9 && c0.calc.flow.thin.length === 0, "level 4 with local mesh: smallest cell = (V/(28000×7))^(1/3)/8, no thin part", [c0.calc.flow.minCellMm, c0.calc.flow.thin]);
  const c2 = ctxOf(withFlow({ meshLevel: 2, localMesh: false }));
  const thin = c2.calc.flow.thin.map(x => x.refDes);
  check(thin.includes("L1") && thin.includes("L2") && thin.includes("U3"), "level 2 without local mesh: L1, L2 and U3 may not be meshed", c2.calc.flow);
  const t8 = show("SimTab", c2);
  check(t8.includes("Flow mesh and convergence") && t8.includes('data-term="flow.meshLevel"') && t8.includes("May not be meshed: ") && t8.includes("L1 (1.2 mm)"), "tab 8 shows the Flow mesh card and the thin parts");
  check(!show("SimTab", ctxOf({ ...P, simSetup: { ...P.simSetup, studyType: "Thermal" } })).includes("Flow mesh"), "no Flow mesh card for a Thermal study");

  const a = flowRun({ name: "FL_20261019_2200", unresolved: [], warnings: [], heatAppliedW: 6.67, energyBalancePct: 0.1, mesh: { level: 4, localMesh: true }, meshCheck: { pairWith: "FL_20261019_2200_L5", level: 4 }, parts: { U1: { peakC: 91.6 }, U3: { peakC: 124.5 } } });
  const b = flowRun({ name: "FL_20261019_2200_L5", unresolved: [], warnings: [], heatAppliedW: 6.67, energyBalancePct: 0.1, mesh: { level: 5, localMesh: true }, meshCheck: { pairWith: "FL_20261019_2200", level: 5 }, parts: { U1: { peakC: 91.2 }, U3: { peakC: 126.0 } } });
  const res = { ...R, runs: [...R.runs, a, b] };
  valid("results.schema.json", res, "results.json with a mesh check pair");
  const c = ctxOf(withFlow({}), C, res);
  const m = c.calc.runs[b.name].meshCheck;
  check(m.coarse === a.name && m.fine === b.name && Math.abs(m.parts.U3 - 1.5) < 1e-9 && Math.abs(m.parts.U1 + 0.4) < 1e-9, "mesh check: finer minus coarser per chip", m);
  check(m.ok === false && m.over.join() === "U3" && m.worst === "U3", "U3 moved 1.5 C: not OK against the 1 C limit", m);
  check(c.calc.runs[a.name].meshCheck.fine === b.name, "both runs of the pair see the same comparison");
  const t9 = show("ResultsTab", c);
  check(t9.includes("Mesh check: level 4") && t9.includes('data-term="runp.meshDelta"') && t9.includes("Moved by the limit or more: U3"), "tab 9 shows the mesh check per chip");
  const okc = ctxOf(withFlow({ meshCheckLimitC: 2 }), C, res);
  check(okc.calc.runs[b.name].meshCheck.ok === true, "with a 2 C limit the same pair is OK");
}

// ---------- F3 ----------
{
  const withFlow = f => ({ ...P, simSetup: { ...P.simSetup, studyType: "Flow", flow: f } });
  valid("project.schema.json", withFlow({ minTravels: 2, goalCriteriaC: 0.2 }), "project.json with minTravels and goalCriteriaC");
  check(!ajv.validate("project.schema.json", withFlow({ minTravels: 0 })), "0 travels is refused by the schema");
  const s = flowSettings(withFlow({}));
  check(s.minTravels === 2 && s.goalCriteriaC === 0.2, "defaults: 2 travels, 0.2 C", s);
  const base = { unresolved: [], warnings: [], heatAppliedW: 6.67, energyBalancePct: 0.1 };
  const one = runHealth(flowRun({ ...base, solver: { finished: true, travels: 1.006 }, goalsHealth: { "U2 max": { progressPct: 100, deltaC: 0.36, criteriaC: 2.2 } } }), s);
  check(one.level === "check" && one.reasons.some(r => /only 1.01 travels, fewer than 2/.test(r)) && one.reasons.some(r => /moving more than 0.2 C: U2 max/.test(r)), "1 travel and a goal moving 0.36 C: check with both reasons", one);
  const two = runHealth(flowRun({ ...base, solver: { finished: true, travels: 2.04 }, goalsHealth: { "U2 max": { progressPct: 100, deltaC: 0.12, criteriaC: 2.2 } } }), s);
  check(two.level === "ok", "2 travels and every goal within 0.2 C: ok", two);
  const own = runHealth(flowRun({ ...base, solver: { finished: true, travels: 2.04 }, inputs: { ...flowRun().inputs, simSetup: { ...P.simSetup, flow: { minTravels: 3 } } } }), s);
  check(own.level === "check" && /fewer than 3/.test(own.reasons[0]), "a run is judged with the settings it was made with", own);
  const r = flowRun({ name: "FL_20261019_2300", ...base, solver: { finished: true, travels: 2.04 }, goalsHealth: { "U2 max": { progressPct: 100, deltaC: 0.36, criteriaC: 2.2 }, "U1 max": { progressPct: 100, deltaC: 0.1, criteriaC: 2 } } });
  const c = ctxOf(withFlow({}), C, { ...R, runs: [...R.runs, r] });
  check(c.value("run.goalsLoose", r.name) === "U2 max" && Math.abs(c.value("run.worstDelta", r.name) - 0.36) < 1e-9, "goals still moving and the largest delta on the page");
  const t8 = show("SimTab", c);
  check(t8.includes('data-term="flow.minTravels"') && t8.includes('data-term="flow.goalCriteria"'), "tab 8 has the convergence inputs");
}

// ---------- F4 ----------
{
  check(kStep(1) === 1 && kStep(10) === 10 && kStep(20.44) === 19.6, "k steps: 1, 10 exact, 20.44 to 19.6", [kStep(1), kStep(10), kStep(20.44)]);
  let worst = 0;
  for (let k = 0.2; k < 400; k *= 1.013) worst = Math.max(worst, Math.abs(kStep(k) / k - 1));
  check(worst < 0.05, `every step within 5 % of k (worst ${(worst * 100).toFixed(2)} %)`);
  const m = flowMaterial(20.76, 0.3287);
  check(m.name === "k 21.5 0.316" && m.path === "User Defined\\PCBThermal\\k 21.5 0.316" && flowMaterial(5).name === "k 5.11", "material names: k xy k z, or one k when the same", [m, flowMaterial(5)]);

  const c0 = ctxOf(P);
  const fl = c0.calc.fsNow.flow;
  valid("project.schema.json", { ...P, forSolidWorks: { ...c0.calc.fsNow, computedAt: NOW } }, "project.json with forSolidWorks.flow");
  check(fl.chips.U1.basis === "calibrated block" && fl.chips.U1.material.endsWith("k 19.6") && fl.chips.L1.basis === "package" && !fl.chips.J1 === !P.forSolidWorks.bodies.some(b => b.refDes === "J1"), "U1 calibrated block (k 19.6), L1 package", fl.chips);
  check(fl.materials.some(x => x.name === "k 1.33" && x.usedBy.join() === "U2,U4"), "U2 and U4 share the material k 1.33", fl.materials);
  const u1 = c0.calc.simChips.find(x => x.refDes === "U1");
  check(Math.abs(c0.value("fchip.rjcStep", "U1") - (u1.hMm * 1000) / (19.6 * u1.areaMm2)) < 1e-9 && Math.abs(c0.value("fchip.rjcStep", "U1") / 0.5 - 1) < 0.05, "Rjc given by the U1 step is within 5 % of 0.5 C/W", c0.value("fchip.rjcStep", "U1"));

  // two-resistor U1 (Rjb typed): three layer materials
  const P2 = structuredClone(P);
  P2.parts.find(x => x.refDes === "U1").rjbCW = 5;
  P2.simSetup.chips.find(x => x.refDes === "U1").model = "Two-resistor";
  const c2 = ctxOf(P2);
  const ly = c2.calc.fsNow.flow.chips.U1;
  const t = 0.1 * u1.hMm;
  check(ly.basis === "two-resistor" && ly.layers.core.endsWith("k 100") && ly.layers.top.endsWith(`k ${kStep((t * 1000) / (0.5 * u1.areaMm2))}`) && ly.layers.bottom.endsWith(`k ${kStep((t * 1000) / (5 * u1.areaMm2))}`), "two-resistor U1: bottom from Rjb, top from Rjc, core 100", ly);

  const fsT = { ...c0.calc.fsNow };
  const thermal = { name: "PT_X", studyType: "Thermal", status: "read", inputs: { forSolidWorks: fsT }, parts: { U1: { peakC: 80 }, L1: { peakC: 70 } } };
  check(tjBasisOf(thermal, "U1") === "calibrated block" && tjBasisOf(thermal, "L1") === "package", "Thermal run: basis from the run's chip models");
  check(tjBasisOf({ studyType: "Flow", parts: { U1: { peakC: 1 } } }, "U1") === "package", "Flow run from before F4: package");
  const fr = flowRun({ name: "FL_20261019_2350", unresolved: [], warnings: [], heatAppliedW: 6.67, energyBalancePct: 0.1, parts: { U1: { peakC: 91.6, tjBasis: "calibrated block" }, L1: { peakC: 99, tjBasis: "package" } } });
  const res = { ...R, runs: [...R.runs, fr] };
  valid("results.schema.json", res, "results.json with tjBasis per part");
  const c = ctxOf({ ...P, simSetup: { ...P.simSetup, studyType: "Flow" } }, C, res);
  check(c.calc.runs[fr.name].parts.U1.tjBasis === "calibrated block" && c.calc.runs[fr.name].parts.L1.tjBasis === "package", "Flow run: basis as Run Flow wrote it");
  const t9 = show("ResultsTab", c);
  check(t9.includes('data-term="runp.tjBasis"') && t9.includes("package, no "), "tab 9 shows the Tj basis and marks package parts");
  const t8 = show("SimTab", c);
  check(t8.includes("Flow materials") && t8.includes("Make these once in Flow") && t8.includes("k 19.6"), "tab 8 lists the Flow materials to make");
}

// ---------- F5 ----------
{
  const c0 = ctxOf(P);
  const fl = c0.calc.fsNow.flow;
  check(fl.pcb.endsWith(flowMaterial(P.forSolidWorks.pcb.kxy, P.forSolidWorks.pcb.kz).name) && fl.spreaders.top.endsWith("k 162"), "PCB from the stack k, spreader from the plate k (167 to 162)", [fl.pcb, fl.spreaders]);
  check(fl.tims.U1.endsWith("k 2.87") && !fl.notes, "TIM k 3 to step 2.87, no notes when contacts are 0 and emissivities agree", [fl.tims.U1, fl.notes]);
  check(fl.surfaces.pcb.endsWith("e 0.90") && fl.surfaces.spreader.endsWith("e 0.85") && fl.radiative.length === 2, "radiative surfaces from tab 5 and 8 emissivities", fl.surfaces);
  check(flowEmissivity(0.854).name === "e 0.85" && flowEmissivity(0.9).name === "e 0.90", "emissivity names to 2 decimals");

  const P2 = structuredClone(P);
  P2.simSetup.contacts = { partToPcb: 2e-5, chipToTim: 1e-5, timToSpreader: 4e-5 };
  P2.simSetup.surfaces.pcbBottom.emissivity = 0.8;
  const c2 = ctxOf(P2);
  const f2 = c2.calc.fsNow.flow;
  valid("project.schema.json", { ...P2, forSolidWorks: { ...c2.calc.fsNow, computedAt: NOW } }, "project.json with F5 contacts and surfaces");
  const tm = 0.2e-3, kEff = tm / (tm / 3 + 1e-5);
  check(f2.tims.U1.endsWith(`k ${kStep(kEff)}`), `chip to TIM folded into the TIM: k ${kEff.toFixed(3)} to step ${kStep(kEff)}`, f2.tims.U1);
  const onPcb = c2.calc.fsNow.bodies.filter(b => Math.abs(b.box.zMin - c2.calc.fsNow.pcb.h) < 0.05 || Math.abs(b.box.zMin + b.box.h) < 0.05).reduce((s, b) => s + b.box.l * b.box.w, 0);
  check(Math.abs(f2.contacts.partToPcb.rKW - 2e-5 / (onPcb * 1e-6)) < 1e-9 && f2.contacts.partToPcb.areaMm2 === onPcb, "part to PCB: integral R = R contact / sum of footprints", f2.contacts.partToPcb);
  check(f2.contacts.timToSpreader.rKW > 0 && Math.abs(c2.value("fsw.contactR", "timToSpreader") - f2.contacts.timToSpreader.rKW) < 1e-12, "TIM to spreader contact on the page");
  check(f2.notes.some(x => /folded into each TIM/.test(x)) && f2.notes.some(x => /PCB emissivity: top 0.9, bottom 0.8/.test(x)), "simplifications are noted", f2.notes);
  const t8 = show("SimTab", ctxOf({ ...P2, simSetup: { ...P2.simSetup, studyType: "Flow" } }));
  check(t8.includes('data-term="fsw.pcbMat"') && t8.includes('data-term="fsw.contactR"') && t8.includes("And these radiative surfaces") && t8.includes("is folded into each"), "tab 8 lists what goes to Flow, the radiative surfaces and the notes");
}

// ---------- F8 ----------
{
  const c0 = ctxOf(P);
  check(c0.calc.heatNoBody.length === 0 && !c0.calc.fsNow.openChecks, "RPi4 example: every part with heat has a body");
  const P2 = structuredClone(P);
  const key = Object.keys(P2.bodies).find(k => P2.bodies[k].refDes === "J1");
  P2.bodies[key].include = false;
  const c2 = ctxOf(P2);
  const j1 = c2.calc.heatNoBody.find(x => x.refDes === "J1");
  check(j1 && !j1.accepted && c2.calc.fsNow.openChecks?.[0]?.startsWith("J1 has 0.0547 W of heat but no body in the model."), "J1 left out: open check with its heat", c2.calc.fsNow.openChecks);
  valid("project.schema.json", { ...P2, forSolidWorks: { ...c2.calc.fsNow, computedAt: NOW } }, "project.json with openChecks");
  const bb = macroChecks("BuildBoard", c2);
  check(bb.some(x => !x.ok && /Heat with no body: J1/.test(x.text)), "Build Board card refuses while J1 is open", bb.map(x => x.text));
  const t6 = show("BuildTab", c2);
  check(t6.includes("Heat with no body") && t6.includes('data-term="hnb.accepted"') && t6.includes("Build Board and the studies refuse"), "tab 6 lists J1 with a tick box");
  const P3 = { ...P2, checks: { ...(P2.checks || {}), "heat-no-body/J1": { at: NOW } } };
  const c3 = ctxOf(P3);
  check(c3.calc.heatNoBody[0].accepted && !c3.calc.fsNow.openChecks, "ticked: no open check left");
  valid("project.schema.json", { ...P3, forSolidWorks: { ...c3.calc.fsNow, computedAt: NOW } }, "project.json with the accepted check");
  check(macroChecks("RunFlow", c3).some(x => x.ok && /on purpose: J1/.test(x.text)), "Run Flow card says J1 is left out on purpose");
}

// ---------- F6 ----------
{
  const withEnv = e => ({ ...P, simSetup: { ...P.simSetup, studyType: "Flow", flow: { environment: e } } });
  const fs = ctxOf(P).calc.fsNow;
  const flat = flowEnvironment(P, fs);
  check(flat.up === "+Z" && flat.gravity.join() === "0,0,-9.81" && flat.pressureKPa === 101.325 && flat.natural && flat.margins.zMax === 1.5 && flat.margins.xMin === 0.5, "default: flat, parts up, g -Z, sea level, margins 0.5 and 1.5 above", flat);
  check(JSON.stringify(fs.flow.environment.margins) === JSON.stringify(flat.margins) && fs.flow.environment.pressurePa === 101325, "forSolidWorks.flow.environment carries the same", fs.flow.environment);
  const old = flowDomainMm(fs), now = flowDomainMm(fs, flat);
  check(JSON.stringify(old) === JSON.stringify(now), "default margins give the domain RunFlow used before F6");
  const v = flowEnvironment(withEnv({ orientation: "vertical, long side up" }), fs);
  check(v.up === "+Y" && v.gravity.join() === "0,-9.81,0" && v.margins.yMax === 1.5 && v.margins.zMax === 0.5 && !v.natural, "RPi4 standing on its long edge: g -Y, domain grows along +Y", v);
  check(flowEnvironment(withEnv({ orientation: "vertical, short side up" }), fs).up === "+X" && flowEnvironment(withEnv({ orientation: "flat, parts down" }), fs).gravity[2] === 9.81, "short side up: +X; parts down: g +Z");
  const alt = flowEnvironment(withEnv({ altitudeM: 2000 }), fs);
  check(Math.abs(alt.pressureKPa - 79.5) < 0.1, `2000 m: ${alt.pressureKPa.toFixed(2)} kPa (standard atmosphere 79.5)`);
  check(flowEnvironment(withEnv({ altitudeM: 2000, pressureKPa: 90 }), fs).pressureKPa === 90, "a typed pressure wins over the altitude");
  const fan = flowEnvironment(withEnv({ airSpeedMs: 1.5, airDirection: "-X" }), fs);
  check(fan.velocity.join() === "-1.5,0,0" && fan.margins.xMin === 2 && fan.margins.zMax === 1.5, "forced air -X: velocity and 2 x downstream", fan);
  const Pv = withEnv({ orientation: "vertical, long side up", airSpeedMs: 1 });
  valid("project.schema.json", Pv, "project.json with simSetup.flow.environment");
  const cv = ctxOf(Pv);
  valid("project.schema.json", { ...Pv, forSolidWorks: { ...cv.calc.fsNow, computedAt: NOW } }, "forSolidWorks with the environment");
  check(Math.abs(cv.value("env.pressureUsed") - 101.325) < 1e-9 && cv.value("env.up") === "+Y", "environment values on the page");
  const t8 = show("SimTab", cv);
  check(t8.includes("Flow environment") && t8.includes('data-term="env.orientation"') && t8.includes("svg") && t8.includes("assumes a flat board in still air"), "tab 8 shows the environment, the drawing and the hand check note");
  check(show("BareTab", cv).includes("are not the same case"), "tab 5 notes that the hand check is a different case");
  check(!show("BareTab", ctxOf(withEnv({}))).includes("are not the same case"), "no note for a flat board in still air");
}

console.log(`Flow plan: ${n - fails} of ${n} checks passed.`);
if (fails) process.exit(1);
