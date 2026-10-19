// Page state helpers. Pure functions, no browser access, so tools/test_w3.mjs can run them in Node.
//
//   newProject(name, now, filePrefix)     -> a valid empty project file; emptyCad, emptyResults: the other two
//   setIn(obj, path, value)               -> copy of obj with path set; a blank value removes the key
//   projectForSave(project, cad, res, now) -> what Save writes: blanks removed, savedBy, savedAt, forSolidWorks
//   parseNumber(text, unit)               -> { value } or { error }, with the input helpers of section 13.4
//   history: historyInit, commit, undo, redo

import { calcProject, forSolidWorks } from "../calc/index.js";

export const PAGE_VERSION = "0.1.0";

export function isoLocal(d) {
  const p = v => String(v).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

// Template defaults (12 Settings, 3 CAD geometry PCB stack, register flag limits, 6 Bare board network air and
// surfaces, 11 Sim setup). Without the surfaces a new board's hand check would have no h.
export const SIM_DEFAULTS = {
  studyType: "Thermal", meshGlobalMm: 3, meshFineMm: 0.5, includePartsWithoutHeat: true, kPartsWithoutHeat: 0.5,
  surfaces: {
    spreaderTop: { h: 10, emissivity: 0.85 }, spreaderOther: { h: 5, emissivity: 0.85 },
    pcbTop: { h: 8, emissivity: 0.9 }, pcbBottom: { h: 5, emissivity: 0.9 }, parts: { h: 8, emissivity: 0.9 },
  },
  contacts: { partToPcb: 0, chipToTim: 0, timToSpreader: 0 },
  chips: [],
};
// A project made before the page (or by hand) may have no simSetup: give it the template values before an edit.
export const withSimSetup = p => (p.simSetup ? p : { ...p, simSetup: structuredClone(SIM_DEFAULTS) });

// ---------- W7: spreader plates, thermal solution, sim chips, sizes from SolidWorks ----------
// The parts with Pd whose included body is on this side of the board (bottom plate: bodies on the bottom side).
export function heatedOnSide(project, calc, side) {
  return project.parts.filter(x => calc.register.parts[x.refDes]?.pdW > 0).filter(x => {
    const b = calc.bodies.find(r => r.refDes === x.refDes && r.include);
    return b && (side === "bottom" ? b.side === "Bottom" : b.side !== "Bottom");
  }).map(x => x.refDes);
}
export const newPlateChip = refDes => ({ refDes, covered: false, heightMode: "Pedestal", timMode: "Body", timTMm: 0.2, timK: 3 });
// A plate with the template defaults (7 Spreader): aluminium 6061, 1.5 mm, every heated part on its side listed, none covered.
export function addPlate(project, calc, side) {
  const plate = { enabled: true, material: "Aluminium 6061", k: 167, densityKgM3: 2700, thicknessMm: 1.5, clearanceXYMm: 1, clearanceZMm: 1,
    minPedestalMm: 0.3, heightMode: "Auto", chips: heatedOnSide(project, calc, side).map(newPlateChip) };
  return { ...project, spreader: { ...(project.spreader || {}), [side]: plate } };
}
export function removePlate(project, side) {
  const { [side]: gone, ...rest } = project.spreader || {};
  const p = { ...project, spreader: rest };
  if (!Object.keys(rest).length) delete p.spreader;
  return p;
}
// Thermal solution hand check (8 Thermal solution): a plate the size of the PCB, 2 mm aluminium, h 10, one face.
export function addSolution(project, calc, side) {
  const key = side === "bottom" ? "solutionBottom" : "solutionTop";
  const sol = { plateTMm: 2, plateK: 167, hAir: 10, faces: 1, parts: {} };
  if (calc.pcb.l > 0) sol.plateLMm = calc.pcb.l;
  if (calc.pcb.w > 0) sol.plateWMm = calc.pcb.w;
  for (const r of heatedOnSide(project, calc, side)) sol.parts[r] = { underPlate: true, tim1TMm: 0.2, tim1K: 3 };
  return { ...project, handCheck: { ...(project.handCheck || {}), [key]: sol } };
}
export function removeSolution(project, side) {
  const key = side === "bottom" ? "solutionBottom" : "solutionTop";
  const { [key]: gone, ...rest } = project.handCheck || {};
  return { ...project, handCheck: rest };
}
// Sim chips: every part with Pd and a body gets a block model; chips already listed keep their settings.
export function addSimChips(project, calc) {
  const p = withSimSetup(project);
  const have = new Set((p.simSetup.chips || []).map(x => x.refDes));
  const add = project.parts.filter(x => calc.register.parts[x.refDes]?.pdW > 0 && calc.bodies.some(b => b.refDes === x.refDes && b.include) && !have.has(x.refDes))
    .map(x => ({ refDes: x.refDes, model: x.rjbCW > 0 && x.rjcCW > 0 ? "Two-resistor" : "Block" }));
  return { ...p, simSetup: { ...p.simSetup, chips: [...(p.simSetup.chips || []), ...add] } };
}
// Sync CAD found a body whose box in SolidWorks differs from the page: take the SolidWorks values as typed values
// (an override for an imported body, the box itself for a body typed in by hand). Only the fields that differ.
export function takeAsBuilt(project, calc, key, box) {
  const row = calc.bodies.find(b => b.key === key);
  if (!row) return project;
  let p = project;
  for (const f of ["xMin", "yMin", "zMin", "l", "w", "h"]) {
    if (!Number.isFinite(box[f]) || Math.abs(box[f] - (row.box[f] ?? NaN)) <= 1e-4) continue;
    p = setIn(p, row.manual ? ["bodies", key, "manual", "box", f] : ["bodies", key, "override", f], Math.round(box[f] * 1000) / 1000);
  }
  return p;
}

// A body typed in by hand (not in the STEP), key m_<n>. It starts as a 5 x 5 x 1 mm box on the board top.
export function addManualBody(project, pcbH) {
  let n = 1;
  while (project.bodies["m_" + n]) n++;
  const box = { xMin: 0, yMin: 0, zMin: Number.isFinite(pcbH) ? pcbH : 1.6, l: 5, w: 5, h: 1 };
  return { project: { ...project, bodies: { ...project.bodies, ["m_" + n]: { include: true, manual: { name: "New body " + n, box, side: "Top" } } } }, key: "m_" + n };
}

// The cad and results files of a new project: only their header, ready for the macros.
export const emptyCad = now => ({ schema: "pcb-thermal/cad", schemaVersion: 1, savedBy: "page " + PAGE_VERSION, savedAt: now });
export const emptyResults = now => ({ schema: "pcb-thermal/results", schemaVersion: 1, savedBy: "page " + PAGE_VERSION, savedAt: now, runs: [] });

// filePrefix (optional, with its "_") also names the CAD files the macros write.
export function newProject(name, now, filePrefix) {
  const p = {
    schema: "pcb-thermal/project", schemaVersion: 1, savedBy: "page " + PAGE_VERSION, savedAt: now,
    board: { name: name || "Board", ambientC: 40 },
    parts: [],
    bodies: {},
    pcbStack: { copperLayers: 4, copperThicknessMm: 0.035, copperCoverage: 0.6, copperK: 390, dielectricK: 0.3 },
    handCheck: { surfaceRiseK: 20 },
    simSetup: structuredClone(SIM_DEFAULTS),
    settings: {
      tolPosMm: 0.05, tolSizeMm: 0.05,
      import: { pcbNameContains: "PCB", minSizeMm: 0, topLevelOnly: true },
      flagLimits: {
        critical: { marginBelowC: 15, densityAboveWcm2: 2, shareAbove: 0.25 },
        watch: { marginBelowC: 30, densityAboveWcm2: 0.8, shareAbove: 0.1 },
      },
    },
  };
  if (filePrefix) p.settings.filePrefix = filePrefix;
  return p;
}

export const blank = v => v === undefined || v === null || v === "" || (typeof v === "number" && !Number.isFinite(v));

// Immutable set. path is an array of keys and array indexes. A blank value removes the key, and an object that
// becomes empty is removed too (an empty override is not allowed), except the sections project.json must have.
const KEEP = new Set(["board", "parts", "bodies", "pcbStack", "settings"]);
export function setIn(obj, path, value, depth = 0) {
  if (!path.length) return value;
  const [k, ...rest] = path;
  const src = obj ?? (typeof k === "number" ? [] : {});
  const copy = Array.isArray(src) ? src.slice() : { ...src };
  const next = rest.length ? setIn(src[k], rest, value, depth + 1) : value;
  const drop = blank(next) || (rest.length && (depth > 0 || !KEEP.has(k)) && isEmptyObject(next));
  if (drop && !Array.isArray(copy)) delete copy[k];
  else copy[k] = next;
  return copy;
}
export const getIn = (obj, path) => path.reduce((o, k) => (o == null ? undefined : o[k]), obj);
const isEmptyObject = x => x && typeof x === "object" && !Array.isArray(x) && Object.keys(x).length === 0;

// Remove blanks everywhere (the files never hold null or "").
export function clean(x) {
  if (Array.isArray(x)) return x.map(clean);
  if (x && typeof x === "object") {
    const o = {};
    for (const [k, v] of Object.entries(x)) if (!blank(v)) o[k] = clean(v);
    return o;
  }
  return x;
}

export function projectForSave(project, cad, results, now) {
  const p = clean(project);
  p.savedBy = "page " + PAGE_VERSION;
  p.savedAt = now;
  const fs = forSolidWorks(p, calcProject(p, cad || {}, results || {}), now);
  if (fs) p.forSolidWorks = fs; else delete p.forSolidWorks;
  return p;
}

// Ids that never change (schema: p<n>, p<n>l<m>).
export function nextPartId(project) {
  const n = Math.max(0, ...project.parts.map(x => Number(x.id.slice(1)) || 0));
  return "p" + (n + 1);
}
export function nextLoadId(part) {
  const n = Math.max(0, ...part.loads.map(l => Number(l.id.split("l")[1]) || 0));
  return part.id + "l" + (n + 1);
}

// A RefDes links a part to its body, spreader chip, thermal solution row and sim chip. Renaming or deleting a
// part keeps those links right. forSolidWorks is left alone: Save works it out again.
export function renameRefDes(project, from, to) {
  const walk = (x, top) => {
    if (Array.isArray(x)) return x.map(v => walk(v));
    if (!x || typeof x !== "object") return x;
    const o = {};
    for (const [k, v] of Object.entries(x)) {
      if (top && (k === "parts" || k === "forSolidWorks")) { o[k] = v; continue; }
      if (k === "refDes" && v === from) o[k] = to;
      else if (k === "parts" && v && typeof v === "object" && !Array.isArray(v)) {
        o[k] = Object.fromEntries(Object.entries(v).map(([r, w]) => [r === from ? to : r, walk(w)]));
      } else o[k] = walk(v);
    }
    return o;
  };
  const p = walk(project, true);
  p.parts = project.parts.map(x => (x.refDes === from ? { ...x, refDes: to } : x));
  return p;
}

export function deletePart(project, id) {
  const part = project.parts.find(x => x.id === id);
  if (!part) return project;
  const ref = part.refDes;
  const still = project.parts.some(x => x.id !== id && x.refDes === ref);
  let p = { ...project, parts: project.parts.filter(x => x.id !== id) };
  if (still) return p;
  p.bodies = Object.fromEntries(Object.entries(p.bodies).map(([k, b]) => {
    if (b.refDes !== ref) return [k, b];
    const { refDes, ...rest } = b;
    return [k, rest];
  }).filter(([, b]) => Object.keys(b).length));
  if (p.spreader) {
    p.spreader = { ...p.spreader };
    for (const s of ["top", "bottom"]) if (p.spreader[s]) p.spreader[s] = { ...p.spreader[s], chips: p.spreader[s].chips.filter(c => c.refDes !== ref) };
  }
  if (p.handCheck) {
    p.handCheck = { ...p.handCheck };
    for (const s of ["solutionTop", "solutionBottom"]) {
      const parts = p.handCheck[s]?.parts;
      if (parts && ref in parts) { const { [ref]: gone, ...rest } = parts; p.handCheck[s] = { ...p.handCheck[s], parts: rest }; }
    }
  }
  if (p.simSetup?.chips) p.simSetup = { ...p.simSetup, chips: p.simSetup.chips.filter(c => c.refDes !== ref) };
  return p;
}

export function duplicatePart(project, id) {
  const src = project.parts.find(x => x.id === id);
  const nid = nextPartId(project);
  const used = new Set(project.parts.map(x => x.refDes));
  let ref = src.refDes + "_copy", i = 2;
  while (used.has(ref)) ref = `${src.refDes}_copy${i++}`;
  const copy = { ...structuredClone(src), id: nid, refDes: ref };
  copy.loads = copy.loads.map((l, j) => ({ ...l, id: `${nid}l${j + 1}` }));
  const at = project.parts.indexOf(src) + 1;
  return { ...project, parts: [...project.parts.slice(0, at), copy, ...project.parts.slice(at)] };
}

export function addPart(project) {
  const id = nextPartId(project);
  const used = new Set(project.parts.map(x => x.refDes));
  let n = project.parts.length + 1;
  while (used.has("NEW" + n)) n++;
  return { ...project, parts: [...project.parts, { id, refDes: "NEW" + n, loads: [{ id: id + "l1" }] }] };
}

// ---------- numbers typed by the user ----------
// unit: the stored unit. Accepted extra units convert on entry (design 13.4).
const ALSO = {
  mm: { mil: 0.0254, um: 0.001, µm: 0.001, mm: 1 },
  copperMm: { oz: 0.035, um: 0.001, µm: 0.001, mil: 0.0254, mm: 1 },
  W: { mw: 0.001, w: 1 },
  CW: { "k/w": 1, "c/w": 1, "°c/w": 1 },
  C: { c: 1, "°c": 1 },
  dC: { k: 1, c: 1, "°c": 1 },
  V: { v: 1, mv: 0.001 },
  A: { a: 1, ma: 0.001 },
  ohm: { ohm: 1, "ω": 1, mohm: 0.001, "mω": 0.001 },
};

export function parseNumber(text, unit, rule = {}) {
  const t = String(text ?? "").trim().replace(",", ".");
  if (t === "") return rule.required ? { error: "required" } : { value: undefined };
  const m = t.match(/^([-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?)\s*(.*)$/i);
  if (!m) return { error: "not a number" };
  let v = Number(m[1]);
  const u = m[2].trim();
  if (u) {
    const table = ALSO[unit] || {};
    const key = Object.keys(table).find(k => k.toLowerCase() === u.toLowerCase());
    if (!key) return { error: `unit '${u}' not accepted here` + (Object.keys(table).length ? ` (use ${Object.keys(table).join(", ")})` : "") };
    v *= table[key];
  }
  if (rule.integer && !Number.isInteger(v)) return { error: "whole number only" };
  if (rule.min !== undefined && v < rule.min) return { error: `at least ${rule.min}` };
  if (rule.minEx !== undefined && v <= rule.minEx) return { error: `must be above ${rule.minEx}` };
  if (rule.max !== undefined && v > rule.max) return { error: `at most ${rule.max}` };
  return { value: v };
}

// ---------- undo and redo ----------
// Each commit is one undo step. Edits to the same field within 1.5 s merge into one step (typing a number).
export function historyInit(present) {
  return { past: [], present, future: [], lastKey: null, lastAt: 0 };
}
export function commit(h, next, label, key = null, t = Date.now()) {
  if (next === h.present) return h;
  if (key && key === h.lastKey && t - h.lastAt < 1500 && h.past.length) {
    return { ...h, present: next, future: [], lastAt: t };
  }
  const past = [...h.past, { state: h.present, label }].slice(-200);
  return { past, present: next, future: [], lastKey: key, lastAt: t };
}
export function undo(h) {
  if (!h.past.length) return h;
  const last = h.past[h.past.length - 1];
  return { past: h.past.slice(0, -1), present: last.state, future: [{ state: h.present, label: last.label }, ...h.future],
    lastKey: null, lastAt: 0, did: "Undo: " + last.label };
}
export function redo(h) {
  if (!h.future.length) return h;
  const [first, ...rest] = h.future;
  return { past: [...h.past, { state: h.present, label: first.label }], present: first.state, future: rest,
    lastKey: null, lastAt: 0, did: "Redo: " + first.label };
}
