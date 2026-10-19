// W9 test: exports.
//   node tools/test_w9.mjs [--write <folder>]   (--write keeps the RPi4 report, workbook and archive for a look)
// 1. Workbook: every tab of the design, units in the headers, a note on every header, values match the page,
//    the Method tab lists every formula. It opens again with SheetJS.
// 2. Archive: a valid zip (read back here), CRCs right, names with folders.
// 3. Report: every section, pass or fail per part, open issues.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { h } from "preact";
import { render } from "preact-render-to-string";
import * as XLSX from "xlsx";
import { calcProject } from "../src/calc/index.js";
import { valueOf, FORMULAS } from "../src/calc/formulas.js";
import { exportWorkbook, SHEETS } from "../src/io/export_xlsx.js";
import { makeZip, crc32 } from "../src/io/zip.js";
import { Report, openIssues } from "../src/page/report.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const ex = path.join(here, "..", "..", "example", "RaspberryPi4");
const read = f => JSON.parse(fs.readFileSync(path.join(ex, f), "utf8"));
let fails = 0, n = 0;
const check = (ok, what, extra) => { n++; if (!ok) { fails++; console.log("FAIL " + what + (extra !== undefined ? ": " + JSON.stringify(extra).slice(0, 300) : "")); } };
const P = read("project.json"), C = read("cad.json"), R = read("results.json");
const ctx = { project: P, cad: C, results: R, calc: calcProject(P, C, R) };
const outDir = process.argv[2] === "--write" ? process.argv[3] : null;

// ---------- 1. workbook ----------
{
  const wb = exportWorkbook(XLSX, ctx, R.runs.map(r => r.name));
  check(JSON.stringify(wb.SheetNames) === JSON.stringify(SHEETS), "workbook has the tabs of the design", wb.SheetNames);
  const bytes = XLSX.write(wb, { type: "buffer", bookType: "xlsx", compression: true });
  if (outDir) fs.writeFileSync(path.join(outDir, "RPi4_export.xlsx"), bytes);
  const back = XLSX.read(bytes, { type: "buffer" });
  check(back.SheetNames.length === SHEETS.length, "workbook opens again");
  const parts = XLSX.utils.sheet_to_json(back.Sheets.Parts, { header: 1 });
  check(parts[0].includes("Tj max (C)") && parts[0].includes("Part Pd (W)"), "units in the headers", parts[0]);
  const u1 = parts.find(r => r[0] === "U1");
  const col = parts[0].indexOf("Part Pd (W)");
  check(u1 && Math.abs(u1[col] - valueOf(ctx, "part.pd", "U1")) < 1e-9, "Part Pd of U1 equals the page", [u1?.[col], valueOf(ctx, "part.pd", "U1")]);
  const ws = wb.Sheets.Parts;
  const heads = XLSX.utils.decode_range(ws["!ref"]).e.c + 1;
  let noted = 0;
  for (let c = 0; c < heads; c++) if (ws[XLSX.utils.encode_cell({ r: 0, c })]?.c?.length) noted++;
  check(noted === heads, `a note on every Parts header (${noted} of ${heads})`);
  const back2 = back.Sheets.Parts[XLSX.utils.encode_cell({ r: 0, c: 0 })];
  check(back2?.c?.[0]?.t?.includes("Reference designator"), "header notes survive the file", back2?.c);
  const method = XLSX.utils.sheet_to_json(back.Sheets.Method, { header: 1 });
  check(method.length - 1 >= Object.keys(FORMULAS).length - 5 && method.some(r => /Tj est = Ta \+ Pd × Rja/.test(r[2])), `Method tab lists the formulas (${method.length - 1})`);
  const res = XLSX.utils.sheet_to_json(back.Sheets.Results, { header: 1 });
  check(res.some(r => r[0] === R.runs[0].name), "Results tab has the runs");
  const bb = XLSX.utils.sheet_to_json(back.Sheets["Bare board network"], { header: 1 });
  check(bb.length > 5 && bb[0].some(x => /Tj, hand check/.test(x)), "Bare board network tab");
  const spr = XLSX.utils.sheet_to_json(back.Sheets.Spreader, { header: 1 });
  check(spr.some(r => r[0] === "Hand check with a plate") && spr.some(r => r[0] === "top, U1"), "Spreader tab has the chips and the hand check with a plate");
  const last = exportWorkbook(XLSX, ctx, [R.runs[R.runs.length - 1].name]);
  const lr = XLSX.utils.sheet_to_json(last.Sheets.Results, { header: 1 });
  check(!lr.some(r => r[0] === R.runs[0].name), "one run picked: only that run on the Results tab");
}

// ---------- 2. archive ----------
{
  const files = [{ name: "AB_project.json", data: JSON.stringify(P) }, { name: "CAD/AB_BuildBoard_report.txt", data: "report ü" }, { name: "CAD/results/x.png", data: new Uint8Array([137, 80, 78, 71, 0, 1, 2]) }];
  const z = makeZip(files);
  if (outDir) fs.writeFileSync(path.join(outDir, "RPi4_archive.zip"), z);
  const dv = new DataView(z.buffer);
  const end = z.length - 22;
  check(dv.getUint32(end, true) === 0x06054b50 && dv.getUint16(end + 10, true) === 3, "end record with 3 files");
  let at = dv.getUint32(end + 16, true), ok = true;
  const enc = new TextEncoder();
  for (const f of files) {
    const nlen = dv.getUint16(at + 28, true), off = dv.getUint32(at + 42, true), size = dv.getUint32(at + 20, true), crc = dv.getUint32(at + 16, true);
    const name = new TextDecoder().decode(z.slice(at + 46, at + 46 + nlen));
    const data = z.slice(off + 30 + dv.getUint16(off + 26, true), off + 30 + dv.getUint16(off + 26, true) + size);
    const want = typeof f.data === "string" ? enc.encode(f.data) : f.data;
    if (name !== f.name || crc !== crc32(want) || crc32(data) !== crc || data.length !== want.length) ok = false;
    at += 46 + nlen;
  }
  check(ok, "every file reads back with its name and CRC");
  check(crc32(enc.encode("123456789")) === 0xcbf43926, "CRC-32 check value");
}

// ---------- 3. report ----------
{
  const runs = R.runs.map(r => r.name);
  const out = render(h(Report, { ctx, runs, images: {}, now: "2026-10-19T12:00:00" }));
  if (outDir) fs.writeFileSync(path.join(outDir, "RPi4_report.html"), `<!doctype html><meta charset="utf-8"><style>${fs.readFileSync(path.join(here, "..", "src", "page", "page.css"), "utf8")}</style>${out}`);
  for (const s of ["1 Assumptions", "2 Parts and power", "3 Hand check: bare board", "4 Hand check with a spreader plate", "5 Simulation", "6 Pass or fail per part", "7 Open issues"]) check(out.includes(s), `report section ${s}`);
  for (const r of runs) check(out.includes(`Run ${r}`), `report has run ${r}`);
  check(/FAIL: .*above Tj max/.test(out), "cover says FAIL with the parts above Tj max");
  const iss = openIssues(ctx);
  check(iss.some(x => /above Tj max/.test(x)), "open issues name the failing run", iss);
  const clean = openIssues({ ...ctx, project: { ...P, checks: {} } });
  check(clean.some(x => /15 change records not ticked/.test(x)), "open issues count the unticked changes");
}

console.log(`W9: ${n - fails} of ${n} checks passed.`);
if (fails) process.exit(1);
