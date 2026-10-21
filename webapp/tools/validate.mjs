// Check project.json, cad.json and results.json in a folder against the schemas, plus the links between them.
//   node tools/validate.mjs <project folder>
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import Ajv2020 from "ajv/dist/2020.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const schemaDir = path.join(here, "..", "schema");

export function makeValidator() {
  const ajv = new Ajv2020({ allErrors: true, strict: true, allowUnionTypes: true });
  for (const f of ["common", "project", "cad", "results"]) {
    ajv.addSchema(JSON.parse(fs.readFileSync(path.join(schemaDir, f + ".schema.json"), "utf8")), f + ".schema.json");
  }
  return ajv;
}

// Returns a list of problems (empty = valid).
export function validateFiles(files, ajv = makeValidator()) {
  const problems = [];
  for (const k of ["project", "cad", "results"]) {
    if (!files[k]) { problems.push(`${k}.json is missing`); continue; }
    const ok = ajv.validate(k + ".schema.json", files[k]);
    if (!ok) for (const e of ajv.errors) problems.push(`${k}.json ${e.instancePath || "/"}: ${e.message}${e.params?.additionalProperty ? " (" + e.params.additionalProperty + ")" : ""}`);
  }
  if (problems.length) return problems;

  // Links the schema cannot check.
  const { project: p, cad: c, results: r } = files;
  const refs = new Set(p.parts.map(x => x.refDes));
  const dup = p.parts.map(x => x.refDes).filter((x, i, a) => a.indexOf(x) !== i);
  if (dup.length) problems.push(`project.json: RefDes used by two parts: ${[...new Set(dup)].join(", ")}`);
  const ids = p.parts.flatMap(x => [x.id, ...x.loads.map(l => l.id)]);
  if (new Set(ids).size !== ids.length) problems.push("project.json: part or load id used twice");
  const importKeys = new Set((c.import?.bodies || []).map(b => b.key));
  for (const [key, b] of Object.entries(p.bodies)) {
    if (b.refDes && !refs.has(b.refDes)) problems.push(`project.json bodies.${key}: RefDes ${b.refDes} is not a part`);
    if (!b.manual && c.import && !importKeys.has(key)) problems.push(`project.json bodies.${key}: not in cad.json import and not a manual body`);
  }
  const fs = p.forSolidWorks;
  if (fs) {
    for (const b of fs.bodies) if (!p.bodies[b.key]?.include) problems.push(`forSolidWorks body ${b.key} is not included in bodies`);
    for (const k of Object.keys(fs.heat)) if (!refs.has(k)) problems.push(`forSolidWorks heat ${k} is not a part`);
  }
  const changeIds = new Set((c.changes || []).map(x => x.id));
  if (changeIds.size !== (c.changes || []).length) problems.push("cad.json: change id used twice");
  for (const id of Object.keys(p.checks || {})) if (!changeIds.has(id)) problems.push(`project.json checks: ${id} is not a change in cad.json`);
  const runNames = r.runs.map(x => x.name);
  if (new Set(runNames).size !== runNames.length) problems.push("results.json: run name used twice");
  return problems;
}

export function readFolder(dir) {
  const files = {};
  for (const k of ["project", "cad", "results"]) {
    const f = path.join(dir, k + ".json");
    if (fs.existsSync(f)) files[k] = JSON.parse(fs.readFileSync(f, "utf8"));
  }
  return files;
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, "$1"))) {
  const dir = process.argv[2];
  if (!dir) { console.error("Usage: node tools/validate.mjs <project folder>"); process.exit(2); }
  const problems = validateFiles(readFolder(dir));
  if (problems.length) { console.log(`${problems.length} problems:`); problems.forEach(x => console.log("  " + x)); process.exit(1); }
  console.log("Valid: project.json, cad.json, results.json");
}
