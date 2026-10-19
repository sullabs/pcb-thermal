// Flow plan test (docs/FLOW_PLAN.md, phases F1 to F8): what RunFlow writes and what the page makes of it.
//   node tools/test_flow.mjs
// F1. Run health: the fields RunFlow writes are valid, runHealth gives ok / check / bad with plain reasons, tab 9
//     shows the badge and greys out PASS/FAIL of a bad run, the report and the tab state name it.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { h } from "preact";
import { render } from "preact-render-to-string";
import { makeValidator } from "./validate.mjs";
import { calcProject } from "../src/calc/index.js";
import { valueOf } from "../src/calc/formulas.js";
import { runHealth } from "../src/calc/runs.js";
import { tabStates, blockingProblems } from "../src/page/tabs.js";
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

console.log(`Flow plan: ${n - fails} of ${n} checks passed.`);
if (fails) process.exit(1);
