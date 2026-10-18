// W1 tests: every workbook converts and validates, the RPi4 example matches the committed files,
// and the validator rejects broken files.
//   node tools/test_w1.mjs            run the tests
//   node tools/test_w1.mjs --write    also rewrite the committed RPi4 example files
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { convertFile, writeProject } from "./convert_workbook.mjs";
import { makeValidator, validateFiles, readFolder } from "./validate.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
const exDir = path.join(root, "example", "RaspberryPi4");
const NOW = "2026-10-18T12:00:00";   // fixed, so the example files do not change on every run
const ajv = makeValidator();
let fails = 0;
const check = (ok, what) => { console.log(`${ok ? "ok  " : "FAIL"} ${what}`); if (!ok) fails++; };
const skip = what => console.log(`skip ${what}`);
// The CAD folder (build reports) and some test workbooks are not in git. Checks that need them skip.
const hasReport = f => fs.existsSync(path.join(exDir, "CAD", f));
const near = (a, b, rel = 1e-9) => Math.abs(a - b) <= rel * Math.max(1, Math.abs(b));

// 1. Every workbook we have converts and validates.
const books = [
  path.join(root, "workbook", "PCB_thermal_TEMPLATE.xlsx"),
  ...fs.readdirSync(exDir).filter(f => /\.xlsx$/i.test(f) && !f.startsWith("~$")).map(f => path.join(exDir, f)),
];
const outs = {};
for (const b of books) {
  const out = convertFile(b, { now: NOW });
  outs[path.basename(b)] = out;
  const problems = validateFiles(out, ajv);
  check(problems.length === 0, `${path.basename(b)} converts and validates${problems.length ? ": " + problems.slice(0, 5).join("; ") : ""}`);
}

// 2. RPi4 reference values.
const rp = outs["RaspberryPi4.xlsx"];
const p = rp.project, c = rp.cad, r = rp.results;
check(p.parts.length === 9, "RPi4: 9 parts (U3 rows merged)");
check(p.parts.find(x => x.refDes === "U3")?.loads.length === 4, "RPi4: U3 has 4 loads");
const totalPd = Object.values(p.forSolidWorks.heat).reduce((a, b) => a + b, 0);
check(near(totalPd, 6.728385147058824), `RPi4: total Pd ${totalPd.toFixed(6)} W = register D4`);
check(p.forSolidWorks.bodies.length === 20, "RPi4: 20 included bodies");
check(near(p.forSolidWorks.pcb.kxy, 20.75925) && near(p.forSolidWorks.pcb.kz, 0.328726731512456), "RPi4: PCB k in-plane and through-plane");
check(p.pcbStack.copperThicknessMm === 0.035, "RPi4: copper 35 um stored as 0.035 mm");
check(p.bodies.m_l1?.manual?.box.l === 2.5, "RPi4: hand-added body L1 kept with its box");
check(p.forSolidWorks.plates?.top?.cuts.length === 11 && p.forSolidWorks.plates.top.faceZMm === 4.1, "RPi4: top plate face 4.1 mm, 11 cuts");
check(c.import?.bodies.length === 18 && c.import.pcb.box.l === 85, "RPi4: 18 imported bodies, PCB 85 mm");
if (!hasReport("RaspberryPi4_BuildBoard_report.txt")) skip("RPi4: build report (CAD folder not here)");
else check(c.build?.bodies.length === 20 && c.build.bodies.filter(b => b.cutAtBoardMm).length === 11 && c.build.overlaps.length === 1, "RPi4: build report read (20 built, 11 cut, 1 overlap)");
check(c.changes.length === 15 && Object.keys(p.checks).length === 15, "RPi4: 15 changes, all checked");
check(r.runs.length === 4 && r.runs[0].parts.U1.peakC === 76, "RPi4: 4 runs, U1 peak 76 C in the first");

const fr = outs["RPi4_fresh.xlsx"];
if (!fr) skip("RPi4_fresh.xlsx not here");
else {
  check(fr.project.spreader?.bottom?.chips.some(x => x.refDes === "U7" && x.covered), "RPi4_fresh: bottom plate covers U7");
  if (hasReport("RPi4_fresh_BuildSpreader_report.txt")) check(fr.cad.spreaderBuild?.bottom?.tims.includes("U7"), "RPi4_fresh: BuildSpreader report read");
  else skip("RPi4_fresh: BuildSpreader report (CAD folder not here)");
}
const mt = outs["RPi4_manual_test.xlsx"];
if (mt && hasReport("RPi4_manual_test_BuildBoard_report.txt")) check(mt.cad.build?.bodies.filter(b => b.status === "failed").length === 3, "RPi4_manual_test: 3 failed bodies read from the report");
else skip("RPi4_manual_test: workbook or report not here");
check(!outs["PCB_thermal_TEMPLATE.xlsx"].project.forSolidWorks, "Template: no forSolidWorks before an import");

// 3. The committed example matches a fresh conversion.
if (process.argv.includes("--write")) { writeProject(rp, exDir); console.log("     wrote the example files"); }
const committed = readFolder(exDir);
for (const k of ["project", "cad", "results"]) {
  if (k === "cad" && !hasReport("RaspberryPi4_BuildBoard_report.txt")) { skip("example cad.json compare (needs the build report)"); continue; }
  check(JSON.stringify(committed[k]) === JSON.stringify(rp[k]), `example/RaspberryPi4/${k}.json is up to date (run with --write to refresh)`);
}

// 4. The validator rejects broken files.
const broken = [
  ["part without loads", f => { f.project.parts[0].loads = []; }],
  ["negative Rjc", f => { f.project.parts[0].rjcCW = -1; }],
  ["unknown class", f => { f.project.parts[0].class = "Toaster"; }],
  ["two parts with one RefDes", f => { f.project.parts[1].refDes = f.project.parts[0].refDes; }],
  ["body linked to a missing RefDes", f => { Object.values(f.project.bodies)[0].refDes = "U99"; }],
  ["check of a change that does not exist", f => { f.project.checks["20990101T000000-1"] = { at: NOW }; }],
  ["unknown field", f => { f.project.board.colour = "green"; }],
  ["null value", f => { f.project.board.notes = null; }],
  ["wrong schema name", f => { f.cad.schema = "pcb-thermal/project"; }],
  ["run without a name", f => { delete f.results.runs[0].name; }],
];
for (const [what, mutate] of broken) {
  const f = JSON.parse(JSON.stringify({ project: rp.project, cad: rp.cad, results: rp.results }));
  mutate(f);
  check(validateFiles(f, ajv).length > 0, `rejects: ${what}`);
}

console.log(fails ? `\n${fails} failed` : "\nAll passed");
process.exit(fails ? 1 : 0);
