// Convert a filled PCB thermal workbook to the three project files.
//   node tools/convert_workbook.mjs <workbook.xlsx> <out folder> [--step <file>] [--name <board name>]
// Build reports are read from <workbook folder>/CAD/<prefix>BuildBoard_report.txt when they exist.
// Reads the workbook only; never writes it.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import * as XLSX from "xlsx";
import { workbookToProject } from "../src/io/workbook_import.js";

XLSX.set_fs(fs);

export function convertFile(xlsxPath, opts = {}) {
  const wb = XLSX.readFile(xlsxPath, { cellFormula: false, cellNF: false, cellDates: false });
  const dir = path.dirname(xlsxPath);
  const base = path.basename(xlsxPath, path.extname(xlsxPath));
  const prefix = opts.prefix || base + "_";
  const readIf = f => (fs.existsSync(f) ? fs.readFileSync(f, "utf8") : undefined);
  let stepFile = opts.stepFile;
  if (!stepFile) {
    const steps = fs.readdirSync(dir).filter(f => /\.(step|stp)$/i.test(f));
    if (steps.length === 1) stepFile = steps[0];
  }
  return workbookToProject(wb, XLSX, {
    fileName: path.basename(xlsxPath),
    boardName: opts.boardName,
    stepFile,
    now: opts.now,
    forSolidWorksFrom: opts.forSolidWorksFrom,
    savedBy: "convert_workbook 0.1.0",
    reports: {
      buildBoard: readIf(path.join(dir, "CAD", prefix + "BuildBoard_report.txt")),
      buildSpreader: readIf(path.join(dir, "CAD", prefix + "BuildSpreader_report.txt")),
    },
  });
}

export function writeProject(out, outDir) {
  fs.mkdirSync(outDir, { recursive: true });
  for (const k of ["project", "cad", "results"]) {
    fs.writeFileSync(path.join(outDir, k + ".json"), JSON.stringify(out[k], null, 2) + "\n", "utf8");
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, "$1"))) {
  const args = process.argv.slice(2);
  const flag = n => { const i = args.indexOf(n); return i >= 0 ? args.splice(i, 2)[1] : undefined; };
  const stepFile = flag("--step"), boardName = flag("--name");
  const [src, outDir] = args;
  if (!src || !outDir) { console.error("Usage: node tools/convert_workbook.mjs <workbook.xlsx> <out folder> [--step f] [--name n]"); process.exit(2); }
  const out = convertFile(src, { stepFile, boardName });
  writeProject(out, outDir);
  console.log(`Wrote project.json, cad.json, results.json to ${outDir}`);
  console.log(`${out.project.parts.length} parts, ${Object.keys(out.project.bodies).length} body choices, ${out.cad.import?.bodies.length ?? 0} imported bodies, ${out.results.runs.length} runs, ${out.cad.changes?.length ?? 0} changes.`);
  for (const n of out.notes) console.log("  note: " + n);
}
