// The parts register as a sheet: one row per load, the part's own cells on its first row (as in the workbook's
// Component register). Pure functions, tested by tools/test_w5.mjs.
//
//   SHEET_COLS                          -> the input columns in screen order (what a pasted block fills)
//   sheetRows(project)                  -> [{ i, j, part, load }] one per load, in screen order
//   insertPart(project, at)             -> { project, id }: a new part with one load at index `at`
//   movePart(project, from, to)         -> parts reordered; `to` is the index the part ends up at
//   addLoad(project, partId, at)        -> a new empty load at index `at` of that part
//   moveLoad(project, partId, from, to) -> that part's loads reordered
//   deleteLoad(project, partId, loadId) -> the load removed (a part keeps at least one)
//   pasteCells(project, start, col, text) -> { project, problems, added }: a block copied from Excel, written from
//                                            row `start` and column key `col` on, with new parts below the last row

import { parseNumber, nextPartId, nextLoadId, renameRefDes, setIn } from "./model.js";
import { CLASSES, METHODS } from "../calc/constants.js";

// key, where the value lives, field, term, kind, unit for parseNumber, rule, width.
const num = (key, where, field, term, unit, rule, width = "5em") => ({ key, where, field, term, kind: "num", unit, rule, width });
const txt = (key, where, field, term, width) => ({ key, where, field, term, kind: "text", width });
const sel = (key, where, field, term, options, width) => ({ key, where, field, term, kind: "select", options, width });
export const SHEET_COLS = [
  { ...txt("refDes", "part", "refDes", "part.refDes", "6.5em"), kind: "refdes" },
  txt("partNumber", "part", "partNumber", "part.partNumber", "11em"),
  sel("class", "part", "class", "part.class", CLASSES, "10em"),
  txt("package", "part", "package", "part.package", "7em"),
  num("tjMaxC", "part", "tjMaxC", "part.tjMax", "C", { min: -273.15 }),
  num("rjaCW", "part", "rjaCW", "part.rja", "CW", { minEx: 0 }),
  num("rjcCW", "part", "rjcCW", "part.rjc", "CW", { minEx: 0 }),
  num("rjbCW", "part", "rjbCW", "part.rjb", "CW", { minEx: 0 }),
  num("taOverrideC", "part", "taOverrideC", "part.taOverride", "C", { min: -273.15 }),
  txt("label", "load", "label", "load.label", "10em"),
  txt("rail", "load", "rail", "load.rail", "6em"),
  sel("method", "load", "method", "load.method", METHODS, "11em"),
  num("vinV", "load", "vinV", "load.vin", "V", {}),
  num("iinA", "load", "iinA", "load.iin", "A", { min: 0 }),
  num("voutV", "load", "voutV", "load.vout", "V", {}),
  num("ioutA", "load", "ioutA", "load.iout", "A", { min: 0 }),
  num("efficiency", "load", "efficiency", "load.eff", "", { minEx: 0, max: 1 }, "4em"),
  num("rOhm", "load", "rOhm", "load.r", "ohm", { min: 0 }),
  num("pdMeasuredW", "load", "pdMeasuredW", "load.pdMeas", "W", { min: 0 }),
  txt("notes", "part", "notes", "part.notes", "18em"),
];

// Short header texts where the term label is too long for a column.
const SHORT = { tjMaxC: "Tj max", rjaCW: "Rja", rjcCW: "Rjc", rjbCW: "Rjb", taOverrideC: "Ta override", efficiency: "Eff", pdMeasuredW: "Pd measured" };
for (const c of SHEET_COLS) if (SHORT[c.key]) c.short = SHORT[c.key];

export function sheetRows(project) {
  const out = [];
  project.parts.forEach((part, i) => part.loads.forEach((load, j) => out.push({ i, j, part, load })));
  return out;
}

const freeRef = (project, base = "NEW") => {
  const used = new Set(project.parts.map(x => x.refDes));
  let n = project.parts.length + 1;
  while (used.has(base + n)) n++;
  return base + n;
};

export function insertPart(project, at = project.parts.length, refDes) {
  const id = nextPartId(project);
  const part = { id, refDes: refDes || freeRef(project), loads: [{ id: id + "l1" }] };
  const k = Math.max(0, Math.min(at, project.parts.length));
  return { project: { ...project, parts: [...project.parts.slice(0, k), part, ...project.parts.slice(k)] }, id };
}

const move = (list, from, to) => {
  if (from === to || from < 0 || from >= list.length) return list;
  const out = list.slice();
  const [x] = out.splice(from, 1);
  out.splice(Math.max(0, Math.min(to, out.length)), 0, x);
  return out;
};
export const movePart = (project, from, to) => ({ ...project, parts: move(project.parts, from, to) });

const withPart = (project, partId, f) => ({ ...project, parts: project.parts.map(x => (x.id === partId ? f(x) : x)) });
export const addLoad = (project, partId, at) => withPart(project, partId, x => {
  const k = at === undefined ? x.loads.length : Math.max(0, Math.min(at, x.loads.length));
  return { ...x, loads: [...x.loads.slice(0, k), { id: nextLoadId(x) }, ...x.loads.slice(k)] };
});
export const moveLoad = (project, partId, from, to) => withPart(project, partId, x => ({ ...x, loads: move(x.loads, from, to) }));
export const deleteLoad = (project, partId, loadId) => withPart(project, partId, x => (x.loads.length > 1 ? { ...x, loads: x.loads.filter(l => l.id !== loadId) } : x));

// One cell of text into a column: { value } or { error }. Blank text clears the cell.
export function readCell(col, text) {
  const t = String(text ?? "").trim();
  if (col.kind === "num") return parseNumber(t, col.unit, col.rule);
  if (col.kind === "select") {
    if (!t) return { value: undefined };
    const hit = col.options.find(o => o.toLowerCase() === t.toLowerCase());
    return hit ? { value: hit } : { error: `'${t}' is not on the list` };
  }
  if (col.kind === "refdes" && !t) return { error: "a RefDes cannot be blank" };
  return { value: t || undefined };
}

// Excel copies a block as lines of tab separated cells. A single cell is left to the browser.
export const isBlock = text => /\t|\r?\n./.test(String(text).replace(/\r?\n$/, ""));

export function pasteCells(project, start, colKey, text) {
  const lines = String(text).replace(/\r/g, "").replace(/\n$/, "").split("\n").map(l => l.split("\t"));
  const c0 = SHEET_COLS.findIndex(c => c.key === colKey);
  const problems = [];
  let p = project, added = 0;
  if (c0 < 0) return { project, problems: ["Pick a cell to paste into first."], added };
  lines.forEach((cells, r) => {
    let rows = sheetRows(p);
    let row = rows[start + r];
    if (!row) {
      const refCell = c0 === 0 ? String(cells[0] ?? "").trim() : "";
      const free = refCell && !p.parts.some(x => x.refDes === refCell) ? refCell : undefined;
      p = insertPart(p, p.parts.length, free).project;
      added++;
      rows = sheetRows(p);
      row = rows[rows.length - 1];
    }
    cells.forEach((cell, k) => {
      const col = SHEET_COLS[c0 + k];
      if (!col) return;
      const v = readCell(col, cell);
      const where = `row ${start + r + 1}, ${col.key}`;
      if (v.error) { problems.push(`${where}: ${v.error}`); return; }
      const part = p.parts[row.i];
      if (col.kind === "refdes") {
        if (v.value === part.refDes) return;
        if (p.parts.some(x => x.refDes === v.value)) { problems.push(`${where}: ${v.value} is already on the list`); return; }
        p = renameRefDes(p, part.refDes, v.value);
      } else if (col.where === "part") p = setIn(p, ["parts", row.i, col.field], v.value);
      else p = setIn(p, ["parts", row.i, "loads", row.j, col.field], v.value);
    });
  });
  return { project: p, problems, added };
}
