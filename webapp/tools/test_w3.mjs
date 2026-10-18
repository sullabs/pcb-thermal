// W3 test: page shell logic (unlock rules and reasons, edits, undo, save output) and the built page file.
//   node tools/test_w3.mjs
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { makeValidator, validateFiles } from "./validate.mjs";
import { build } from "./build.mjs";
import { calcProject } from "../src/calc/index.js";
import { tabStates, nextStep } from "../src/page/tabs.js";
import {
  newProject, setIn, projectForSave, parseNumber, historyInit, commit, undo, redo,
  addPart, duplicatePart, deletePart, renameRefDes,
} from "../src/page/model.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const ex = path.join(here, "..", "..", "example", "RaspberryPi4");
const read = f => JSON.parse(fs.readFileSync(path.join(ex, f), "utf8"));
const ajv = makeValidator();
let fails = 0, n = 0;
const check = (ok, what, extra) => { n++; if (!ok) { fails++; console.log("FAIL " + what + (extra !== undefined ? ": " + JSON.stringify(extra) : "")); } };
const states = (project, cad = {}, results = { runs: [] }, canSave = true) =>
  Object.fromEntries(tabStates({ project, cad, results, calc: project ? calcProject(project, cad, results) : null, canSave }).map(s => [s.id, s]));
const NOW = "2026-10-18T12:00:00";

// ---------- unlock rules ----------
{
  const s = states(null, {}, { runs: [] }, false);
  for (const id of ["parts", "import", "link", "triage", "bare", "build", "spreader", "sim", "results"]) {
    check(s[id].locked && s[id].reason === "Choose the project folder first, so your work can be saved." && s[id].fixTab === 0, `no project: ${id} locked with the folder reason`, s[id]);
  }
  for (const id of ["project", "changes", "help"]) check(!s[id].locked, `no project: ${id} open`);
}
const P = read("project.json"), C = read("cad.json"), R = read("results.json");
{
  const s = states(P, C, R);
  const locked = Object.values(s).filter(x => x.locked).map(x => x.id);
  check(locked.length === 0, "RPi4 example: every tab open", locked);
  check(s.bare.status === "attention" && /FAIL/.test(s.bare.note), "RPi4 example: bare board needs attention (parts FAIL)", s.bare);
  check(s.changes.status === "done" && !s.changes.badge, "RPi4 example: all 15 changes ticked");

  const notTicked = { ...P, checks: Object.fromEntries(Object.entries(P.checks).slice(2)) };
  const s2 = states(notTicked, C, R);
  check(s2.spreader.locked && s2.spreader.reason.startsWith("Tick the 2 changes on tab 10 first.") && s2.spreader.fixTab === 10, "2 unticked changes lock the spreader with the count", s2.spreader);
  check(s2.sim.locked && s2.changes.badge === 2, "and the simulation setup, and the Changes tab shows 2");

  const { build: _b, ...noBuild } = C;
  const s3 = states(P, noBuild, R);
  check(s3.spreader.reason === "Run Build Board first (tab 6)." && s3.spreader.fixTab === 6, "no build: spreader names Build Board", s3.spreader);
  check(states(P, C, { runs: [] }).results.reason === "Run the study first (tab 8). Results appear here when Read Results has finished.", "no runs: Results locked with its reason");

  const s4 = states({ ...P, bodies: {} }, {}, R);
  check(s4.link.locked && s4.link.reason.startsWith("Run Import STEP in SolidWorks first") && s4.link.fixTab === 2, "no cad.json and no typed in bodies: Link bodies locked", s4.link);
  check(s4.triage.locked && s4.triage.fixTab === 3, "no bodies: Triage locked (no part has a body)", s4.triage);
  check(s4.build.locked && s4.build.reason.startsWith("Include at least one body on tab 3"), "no bodies: Build model locked", s4.build);
  check(s4.bare.locked && /PCB size unknown/.test(s4.bare.reason), "no PCB: bare board names the problem", s4.bare);

  const bad = setIn(P, ["parts", 0, "loads", 0, "vinV"], undefined);
  const s5 = states(bad, C, R);
  check(s5.bare.locked && s5.bare.reason === "Fix the problems listed on tab 4 first: U1: missing input." && s5.bare.fixTab === 4, "U1 without Vin: bare board lists it", s5.bare.reason);
  check(s5.parts.status === "attention" && s5.triage.status === "attention", "and Parts and Triage need attention");
}
{
  // A new project, step by step.
  let p = newProject("Test board", NOW);
  check(ajv.validate("project.schema.json", p), "new project is valid", ajv.errors);
  let s = states(p);
  check(s.import.locked && s.import.reason === "Add at least one part with power first. The import links bodies to your parts by RefDes." && s.import.fixTab === 1, "new project: Import locked until a part has Pd", s.import);
  check(nextStep(Object.values(s)).id === "parts", "new project: next step is Parts and power");
  p = addPart(p);
  p = setIn(p, ["parts", 0, "class"], "Processor / SoC");
  check(states(p).import.locked, "part without Pd: still locked");
  p = setIn(p, ["parts", 0, "loads", 0, "vinV"], 1);
  p = setIn(p, ["parts", 0, "loads", 0, "iinA"], 2);
  s = states(p);
  check(!s.import.locked && s.link.locked, "part with Pd: Import open, Link waits for cad.json", s.link);
  const saved = projectForSave(p, {}, {}, NOW);
  check(ajv.validate("project.schema.json", saved), "new project with a part saves valid", ajv.errors);
  check(saved.forSolidWorks === undefined, "no PCB yet: no forSolidWorks block");
}

// ---------- edits ----------
{
  const o = { a: { b: { c: 1 } }, top: {} };
  check(JSON.stringify(setIn(o, ["a", "b", "c"], "")) === '{"top":{}}', "blank removes the key and empty parents, not top level keys", setIn(o, ["a", "b", "c"], ""));
  check(o.a.b.c === 1, "setIn does not change the original");
  const ov = setIn(setIn(P, ["bodies", "u1", "override", "l"], 12), ["bodies", "u1", "override", "l"], undefined);
  check(!("u1" in ov.bodies) && "bodies" in ov, "override removed again leaves no empty object");

  const ren = renameRefDes(P, "U1", "U100");
  check(ren.parts[0].refDes === "U100" && Object.values(ren.bodies).some(b => b.refDes === "U100") && !Object.values(ren.bodies).some(b => b.refDes === "U1"), "rename moves the body link");
  check(ren.spreader.top.chips.some(c => c.refDes === "U100"), "rename moves the spreader chip");
  const sv = projectForSave(ren, C, R, NOW);
  check(validateFiles({ project: sv, cad: C, results: R }, ajv).length === 0, "renamed project saves valid", validateFiles({ project: sv, cad: C, results: R }, ajv));
  check(sv.forSolidWorks.heat.U100 === P.forSolidWorks.heat.U1, "forSolidWorks heat follows the rename");

  const del = deletePart(P, "p1");
  const sd = projectForSave(del, C, R, NOW);
  check(validateFiles({ project: sd, cad: C, results: R }, ajv).length === 0 && !sd.forSolidWorks.heat.U1, "deleted U1: no links left, saves valid", validateFiles({ project: sd, cad: C, results: R }, ajv));
  const dup = duplicatePart(P, "p1");
  check(dup.parts[1].refDes === "U1_copy" && dup.parts[1].id === "p10" && dup.parts[1].loads[0].id === "p10l1", "duplicate gets a new id and RefDes", dup.parts[1]);
}

// ---------- undo ----------
{
  let h = historyInit({ v: 0 });
  h = commit(h, { v: 1 }, "type", "f", 1000);
  h = commit(h, { v: 12 }, "type", "f", 1500);   // same field, within 1.5 s: one step
  h = commit(h, { v: 3 }, "other", "g", 1600);
  check(h.past.length === 2, "typing in one field is one undo step", h.past.length);
  h = undo(h);
  check(h.present.v === 12 && h.did === "Undo: other", "undo goes back one step", h);
  h = undo(h);
  check(h.present.v === 0, "undo again: the start");
  h = redo(h);
  check(h.present.v === 12 && h.future.length === 1, "redo");
  h = commit(h, { v: 5 }, "new", null, 9000);
  check(h.future.length === 0, "a new edit clears redo");
}

// ---------- typed numbers ----------
{
  const eq = (t, u, v, rule) => { const r = parseNumber(t, u, rule); check(Math.abs(r.value - v) < 1e-12, `parse '${t}' ${u}`, r); };
  eq("1.5", "W", 1.5); eq("250 mW", "W", 0.25); eq("1 oz", "copperMm", 0.035); eq("35 um", "copperMm", 0.035); eq("10 mil", "mm", 0.254);
  eq("2,5", "CW", 2.5); eq("3 K/W", "CW", 3); eq("1e-3", "W", 0.001);
  check(parseNumber("abc", "W").error === "not a number", "text is refused");
  check(/not accepted/.test(parseNumber("3 oz", "W").error), "wrong unit is refused");
  check(parseNumber("-1", "W", { min: 0 }).error === "at least 0", "range is checked");
  check(parseNumber("", "C", { required: true }).error === "required", "required field");
  check(parseNumber("", "C").value === undefined, "blank = left out");
}

// ---------- save output equals the example ----------
{
  const out = projectForSave(P, C, R, NOW);
  const strip = x => { const { savedBy, savedAt, ...r } = x; const fs = r.forSolidWorks && { ...r.forSolidWorks }; if (fs) { delete fs.computedAt; r.forSolidWorks = fs; } return r; };
  check(JSON.stringify(strip(out)) === JSON.stringify(strip(P)), "saving the example unchanged gives the same project.json");
  check(out.savedBy === "page 0.1.0" && out.savedAt === NOW, "savedBy and savedAt");
  check(validateFiles({ project: out, cad: C, results: R }, ajv).length === 0, "saved example validates with its cad and results");
}

// ---------- the built page ----------
{
  const { page, size } = await build({ out: path.join(here, "..", "dist", "PCB_Thermal.html"), quiet: true });
  check(size < 4 * 1024 * 1024, "page under 4 MB", size);
  check(!/<script[^>]+src=/i.test(page) && !/<link[^>]+href=/i.test(page), "no external scripts or styles");
  check(!/https?:\/\/(?!www\.w3\.org|sheetjs|github|git\.io|bit\.ly|oss\.sheetjs)[^"' )]*\.(js|css)\b/.test(page), "no library loaded from the web");
  check((page.match(/<\/script>/g) || []).length === 2, "only the two script tags of the template close");
  check(page.includes('data-theme="dark"'), "dark theme is the default");
}

console.log(`W3: ${n - fails} of ${n} checks passed.`);
if (fails) process.exit(1);
