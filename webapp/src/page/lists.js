// Filling a board fast: a parts list pasted from Excel or a text file, and body suggestions for tab 3.
// Pure functions, tested by tools/test_w5.mjs.
//
//   readPartsList(text)               -> { columns, rows: [{ line, part fields, load fields }], problems }
//   applyPartsList(project, list)     -> { project, added, updated }: one part per RefDes, one load per row
//   partsListText(project)            -> the parts as a tab separated table, with the header readPartsList reads
//   suggestLinks(project, calc)       -> [{ key, refDes, why }]: bodies whose name holds a part number or RefDes

import { parseNumber, nextPartId } from "./model.js";
import { CLASSES, METHODS } from "../calc/constants.js";
import { isNum } from "../calc/util.js";

// Column name (lower case, spaces and signs removed) -> [where, field, unit for parseNumber].
const COLS = [
  [["refdes", "ref", "designator", "reference"], "part", "refDes"],
  [["partnumber", "pn", "mpn", "part"], "part", "partNumber"],
  [["class", "type"], "part", "class"],
  [["package", "pkg", "case"], "part", "package"],
  [["tjmax", "tjmaxc", "tj"], "part", "tjMaxC", "C"],
  [["rja", "θja", "thetaja", "rthja", "rjacw"], "part", "rjaCW", "CW"],
  [["rjc", "θjc", "thetajc", "rthjc", "rjccw"], "part", "rjcCW", "CW"],
  [["rjb", "θjb", "thetajb", "rthjb", "rjbcw"], "part", "rjbCW", "CW"],
  [["taoverride", "localta", "ta"], "part", "taOverrideC", "C"],
  [["notes", "note", "comment"], "part", "notes"],
  [["label", "load", "loadlabel"], "load", "label"],
  [["rail", "supply", "net"], "load", "rail"],
  [["method", "powermethod"], "load", "method"],
  [["vin", "vinv"], "load", "vinV", "V"],
  [["iin", "iina"], "load", "iinA", "A"],
  [["vout", "voutv"], "load", "voutV", "V"],
  [["iout", "iouta"], "load", "ioutA", "A"],
  [["eff", "efficiency", "η"], "load", "efficiency", "frac"],
  [["r", "rohm", "dcr", "resistance"], "load", "rOhm", "ohm"],
  [["pd", "pdmeas", "pdmeasured", "pdw", "power"], "load", "pdMeasuredW", "W"],
];
const squash = s => String(s).toLowerCase().replace(/\(.*?\)/g, "").replace(/[\s_\-./°]/g, "");
const colOf = h => COLS.find(([names]) => names.includes(squash(h)));

export const LIST_HEADER = ["RefDes", "Part number", "Class", "Package", "Tj max", "Rja", "Rjc", "Rjb", "Ta override", "Label", "Rail", "Method", "Vin", "Iin", "Vout", "Iout", "Efficiency", "R", "Pd"];

const pick = (list, text) => {
  const t = squash(text);
  return list.find(x => squash(x) === t) || list.find(x => squash(x).startsWith(t) && t.length >= 3);
};

export function readPartsList(text) {
  const lines = String(text || "").split(/\r?\n/).filter(l => l.trim());
  const problems = [];
  if (!lines.length) return { columns: [], rows: [], problems: ["Nothing pasted."] };
  const sep = lines[0].includes("\t") ? "\t" : lines[0].includes(";") ? ";" : ",";
  const head = lines[0].split(sep).map(s => s.trim());
  const columns = head.map(colOf);
  const unknown = head.filter((h, i) => h && !columns[i]);
  if (unknown.length) problems.push(`Columns not used: ${unknown.join(", ")}.`);
  if (!columns.some(c => c?.[2] === "refDes")) return { columns, rows: [], problems: ["The first line must name the columns, and one must be RefDes. Example: " + LIST_HEADER.join(", ") + "."] };
  const rows = [];
  lines.slice(1).forEach((line, n) => {
    const cells = line.split(sep).map(s => s.trim());
    const row = { line: n + 2, part: {}, load: {} };
    columns.forEach((c, i) => {
      const cell = cells[i];
      if (!c || cell === undefined || cell === "") return;
      const [, where, field, unit] = c;
      let v = cell;
      if (field === "class") { v = pick(CLASSES, cell); if (!v) { problems.push(`Line ${row.line}: class '${cell}' is not one of the classes, left blank.`); return; } }
      else if (field === "method") { v = pick(METHODS, cell); if (!v) { problems.push(`Line ${row.line}: method '${cell}' is not known, left blank.`); return; } }
      else if (unit === "frac") {
        const pc = cell.match(/^([\d.,]+)\s*%$/);
        const p = parseNumber(pc ? pc[1] : cell, "", {});
        v = p.error ? undefined : pc ? p.value / 100 : p.value;
        if (!isNum(v) || v <= 0 || v > 1) { problems.push(`Line ${row.line}: efficiency '${cell}' must be 0 to 1 (or 0 to 100 %).`); return; }
      } else if (unit) {
        const p = parseNumber(cell, unit, {});
        if (p.error) { problems.push(`Line ${row.line}: ${head[i]} '${cell}': ${p.error}.`); return; }
        v = p.value;
      }
      row[where][field] = v;
    });
    if (!row.part.refDes) { problems.push(`Line ${row.line}: no RefDes, line skipped.`); return; }
    rows.push(row);
  });
  return { columns, rows, problems };
}

// One part per RefDes. The first row of a RefDes sets the datasheet values, each row with a load value adds a
// load. A RefDes already on the list gets the values given, and its loads are replaced when the list has some.
export function applyPartsList(project, list) {
  let p = { ...project, parts: project.parts.slice() };
  const added = [], updated = [];
  const byRef = new Map();
  for (const r of list.rows) {
    if (!byRef.has(r.part.refDes)) byRef.set(r.part.refDes, []);
    byRef.get(r.part.refDes).push(r);
  }
  for (const [ref, rows] of byRef) {
    const fields = Object.assign({}, ...rows.map(r => r.part));
    const loads = rows.map(r => r.load).filter(l => Object.keys(l).length);
    const at = p.parts.findIndex(x => x.refDes === ref);
    if (at >= 0) {
      const old = p.parts[at];
      const part = { ...old, ...fields };
      if (loads.length) part.loads = loads.map((l, j) => ({ id: `${old.id}l${j + 1}`, ...l }));
      p.parts[at] = part;
      updated.push(ref);
    } else {
      const id = nextPartId(p);
      const part = { id, ...fields, loads: [] };
      part.loads = (loads.length ? loads : [{}]).map((l, j) => ({ id: `${id}l${j + 1}`, ...l }));
      p.parts.push(part);
      added.push(ref);
    }
  }
  return { project: p, added, updated };
}

export function partsListText(project) {
  const F = ["refDes", "partNumber", "class", "package", "tjMaxC", "rjaCW", "rjcCW", "rjbCW", "taOverrideC"];
  const LF = ["label", "rail", "method", "vinV", "iinA", "voutV", "ioutA", "efficiency", "rOhm", "pdMeasuredW"];
  const cell = v => (v === undefined ? "" : String(v));
  const out = [LIST_HEADER.join("\t")];
  for (const x of project.parts) {
    const loads = x.loads.length ? x.loads : [{}];
    loads.forEach((l, j) => out.push([...F.map(f => (j === 0 || f === "refDes" ? cell(x[f]) : "")), ...LF.map(f => cell(l[f]))].join("\t")));
  }
  return out.join("\n") + "\n";
}

// ---------- body suggestions ----------
// A part number word (letters and digits, 4 or more characters, for example BCM2711B0 or MxL7704) or the RefDes
// as a word in the body name. Only bodies without a part, and parts without a body, get a suggestion.
const words = s => String(s || "").split(/[^A-Za-z0-9]+/).filter(Boolean);
const compact = s => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
export function suggestLinks(project, calc) {
  const linked = new Set(calc.bodies.filter(b => b.refDes).map(b => b.refDes));
  const free = calc.bodies.filter(b => !b.refDes);
  const out = [];
  const taken = new Set();
  for (const x of project.parts) {
    if (linked.has(x.refDes)) continue;
    const keys = words(x.partNumber).filter(w => w.length >= 4 && /\d/.test(w) && /[a-z]/i.test(w)).map(w => w.toLowerCase());
    const hits = free.filter(b => !taken.has(b.key)).map(b => {
      const name = b.name.split("/").pop();
      const ws = words(name);
      const pn = keys.filter(k => compact(name).includes(k));
      const rd = ws.includes(x.refDes);
      return { b, score: pn.length * 2 + (rd ? 3 : 0), why: rd ? `the body name has ${x.refDes}` : `the body name has ${pn.join(", ")}` };
    }).filter(h => h.score > 0).sort((a, z) => z.score - a.score);
    if (hits.length && (hits.length === 1 || hits[0].score > hits[1].score)) {
      taken.add(hits[0].b.key);
      out.push({ key: hits[0].b.key, refDes: x.refDes, why: hits[0].why });
    }
  }
  return out;
}
