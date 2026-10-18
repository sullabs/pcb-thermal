// Make project.json, cad.json and results.json from a filled PCB thermal workbook (template version 1).
// Pure module: no file access. Works in the page (SheetJS loaded) and in Node (tools/convert_workbook.mjs).
//
//   workbookToProject(wb, XLSX, options) -> { project, cad, results, notes }
//
// wb       SheetJS workbook, read with { cellFormula: false } (cached values are used, Excel computed them).
// XLSX     the SheetJS module (for its date code parser).
// options  { fileName, stepFile, reports: { buildBoard, buildSpreader }, now, savedBy }
//
// Reads tabs by the words after their number ("7 Spreader" -> "Spreader"), as the macros do.
// Units stay as the workbook has them (mm, C, W), except copper thickness: um in the workbook, mm in the file.

const CLASSES = ["Processor / SoC", "Memory", "LDO / linear reg", "Switching regulator", "PMIC", "Ethernet PHY",
  "USB / interface IC", "Power inductor", "Sense resistor", "MOSFET / driver", "Connector", "Other",
  "RF / wireless module"];
const METHODS = ["LDO / linear reg", "Switching regulator", "Digital load (V x I)", "Resistive (I squared R)", "Manual entry"];

export function workbookToProject(wb, XLSX, options = {}) {
  const notes = [];
  const now = options.now || isoLocal(new Date());
  const savedBy = options.savedBy || "workbook import";
  const R = reader(wb, XLSX);

  const reg = R.sheet("Component register", true);
  const geo = R.sheet("CAD geometry", true);
  const imp = R.sheet("STEP import", true);
  const impPrev = R.sheet("STEP import (previous)");
  const spr = R.sheet("Spreader");
  const sprB = R.sheet("Spreader bottom");
  const net = R.sheet("Bare board network");
  const sol = R.sheet("Thermal solution");
  const solB = R.sheet("Thermal solution bottom");
  const sim = R.sheet("Sim setup");
  const res = R.sheet("Sim results");
  const log = R.sheet("Change log");
  const set = R.sheet("Settings");
  const rail = R.sheet("Rail budget");

  // ---------- parts (register) ----------
  const regRows = R.nameRows("Reg_Table", 8, 53);
  const groups = new Map();
  for (let r = regRows[0]; r <= regRows[1]; r++) {
    const refDes = str(R.v(reg, "B" + r));
    if (!refDes) continue;
    if (!groups.has(refDes)) groups.set(refDes, []);
    groups.get(refDes).push(r);
  }
  const parts = [];
  let pn = 0;
  for (const [refDes, rows] of groups) {
    pn++;
    const first = rows[0];
    const part = { id: "p" + pn, refDes };
    const texts = rows.map(r => str(R.v(reg, "C" + r))).filter(Boolean);
    if (texts.length) part.partNumber = rows.length > 1 ? commonWords(texts) || texts[0] : texts[0];
    const cls = str(R.v(reg, "D" + first));
    if (cls) {
      if (CLASSES.includes(cls)) part.class = cls;
      else notes.push(`${refDes}: class '${cls}' is not in the class list; left blank.`);
    }
    put(part, "package", str(R.v(reg, "E" + first)));
    put(part, "tjMaxC", num(R.v(reg, "O" + first)));
    put(part, "rjaCW", pos(R.v(reg, "P" + first)));
    put(part, "rjcCW", pos(R.v(reg, "Q" + first)));
    put(part, "rjbCW", pos(R.v(reg, "R" + first)));
    put(part, "taOverrideC", num(R.v(reg, "S" + first)));
    // Part level values must agree across the rows of one RefDes.
    for (const [col, label] of [["D", "class"], ["E", "package"], ["O", "Tj max"], ["P", "Rja"], ["Q", "Rjc"], ["R", "Rjb"], ["S", "Ta override"]]) {
      const vals = new Set(rows.map(r => String(R.v(reg, col + r) ?? "")));
      if (vals.size > 1) notes.push(`${refDes}: rows ${rows.join(", ")} have different ${label}; the first row's value was kept.`);
    }
    part.loads = rows.map((r, i) => {
      const load = { id: part.id + "l" + (i + 1) };
      if (rows.length > 1) put(load, "label", stripPrefix(str(R.v(reg, "C" + r)), part.partNumber));
      put(load, "rail", str(R.v(reg, "F" + r)));
      const m = str(R.v(reg, "M" + r));
      if (m) {
        if (METHODS.includes(m)) load.method = m;
        else notes.push(`${refDes}: method override '${m}' is not a known method; left blank.`);
      }
      put(load, "vinV", num(R.v(reg, "G" + r)));
      put(load, "iinA", num(R.v(reg, "H" + r)));
      put(load, "voutV", num(R.v(reg, "I" + r)));
      put(load, "ioutA", num(R.v(reg, "J" + r)));
      put(load, "efficiency", num(R.v(reg, "K" + r)));
      put(load, "rOhm", num(R.v(reg, "L" + r)));
      put(load, "pdMeasuredW", num(R.v(reg, "N" + r)));
      return load;
    });
    if (rows.length > 1) notes.push(`${refDes}: ${rows.length} register rows became one part with ${rows.length} loads.`);
    parts.push(part);
  }

  // Text under the register table (assumptions) goes to the board notes.
  const boardNotes = [];
  for (let r = regRows[1] + 1; r <= regRows[1] + 40; r++) {
    const t = str(R.v(reg, "B" + r));
    if (t) boardNotes.push(t);
  }

  // ---------- board ----------
  const ambient = num(R.name("Reg_Ambient") ?? R.v(reg, "D5"));
  const board = { name: options.boardName || baseName(options.fileName) || str(R.v(imp, "D4")) || "Board", ambientC: ambient ?? 25 };
  if (ambient === undefined) notes.push("Ambient was blank; 25 C was used.");
  put(board, "stepFile", options.stepFile);
  if (rail) put(board, "measuredInputPowerW", num(R.v(rail, "E35")));
  if (boardNotes.length) board.notes = boardNotes.join("\n");

  // ---------- bodies (CAD geometry) ----------
  const geoRows = R.nameRows("Geo_Table", 26, 95);
  const bodies = {};
  const geoByName = new Map();      // name used -> key, for the build report
  const fsBodies = [];
  for (let r = geoRows[0]; r <= geoRows[1]; r++) {
    const nameUsed = str(R.v(geo, "E" + r));
    if (!nameUsed) continue;
    const key = str(R.v(geo, "AG" + r));
    if (!key) { notes.push(`CAD geometry row ${r}: no key; skipped.`); continue; }
    geoByName.set(nameUsed, key);
    const imported = str(R.v(geo, "C" + r));
    const choice = {};
    const inc = yesNo(R.v(geo, "G" + r));
    if (inc !== undefined) choice.include = inc;
    put(choice, "refDes", str(R.v(geo, "F" + r)));
    const ov = boxFrom(R, geo, r, "OPQRST");
    if (imported) {
      put(choice, "nameOverride", str(R.v(geo, "D" + r)));
      if (ov) choice.override = ov;
    } else {
      const full = boxFrom(R, geo, r, "OPQRST", true);
      if (!full) { notes.push(`Body ${key} was typed in by hand but its box is incomplete; skipped.`); continue; }
      choice.manual = { name: str(R.v(geo, "D" + r)) || nameUsed, box: full };
    }
    bodies[key] = choice;

    if (choice.include) {
      const used = boxFrom(R, geo, r, "UVWXYZ", true);
      if (used) {
        const b = { key, name: nameUsed, box: used };
        put(b, "refDes", choice.refDes);
        const rot = num(R.v(geo, "AI" + r));
        if (rot) {
          b.rotDeg = rot;
          put(b, "ownL", pos(R.v(geo, "AJ" + r)));
          put(b, "ownW", pos(R.v(geo, "AK" + r)));
        }
        fsBodies.push(b);
      }
    }
  }

  const pcbOverride = {};
  [["l", "D7"], ["w", "D8"], ["h", "D9"]].forEach(([k, a]) => put(pcbOverride, k, pos(R.v(geo, a))));

  const cuUm = num(R.v(geo, "D13"));
  const pcbStack = {
    copperLayers: int(R.v(geo, "D12")) ?? 4,
    copperThicknessMm: cuUm !== undefined ? round(cuUm / 1000, 9) : 0.035,
    copperCoverage: num(R.v(geo, "D14")) ?? 0.6,
    copperK: pos(R.v(geo, "D15")) ?? 390,
    dielectricK: pos(R.v(geo, "D16")) ?? 0.3,
  };

  // ---------- spreader plates ----------
  const spreader = {};
  const fsPlates = {};
  if (spr) { spreader.top = plateFrom(R, spr); fsPlates.top = plateForSw(R, spr, spreader.top); }
  if (sprB) { spreader.bottom = plateFrom(R, sprB); fsPlates.bottom = plateForSw(R, sprB, spreader.bottom); }
  for (const side of ["top", "bottom"]) if (fsPlates[side] === null) delete fsPlates[side];

  // ---------- hand check inputs ----------
  const handCheck = {};
  if (net) {
    put(handCheck, "surfaceRiseK", pos(R.v(net, "E17")));
    put(handCheck, "compareRun", str(R.v(net, "R7")));
  }
  if (sol) handCheck.solutionTop = solutionFrom(R, sol);
  if (solB) handCheck.solutionBottom = solutionFrom(R, solB);

  // ---------- sim setup ----------
  let simSetup;
  if (sim) {
    const surf = r => ({ h: num(R.v(sim, "C" + r)) ?? 0, emissivity: num(R.v(sim, "D" + r)) ?? 0.9 });
    const st = str(R.v(sim, "C5"));
    simSetup = {
      studyType: st === "Flow" ? "Flow" : "Thermal",
      meshGlobalMm: pos(R.v(sim, "C8")) ?? 3,
      meshFineMm: pos(R.v(sim, "C9")) ?? 0.5,
      includePartsWithoutHeat: yesNo(R.v(sim, "C10")) ?? true,
      kPartsWithoutHeat: pos(R.v(sim, "C11")) ?? 0.5,
      surfaces: { spreaderTop: surf(15), spreaderOther: surf(16), pcbTop: surf(17), pcbBottom: surf(18), parts: surf(19) },
      contacts: { partToPcb: num(R.v(sim, "C23")) ?? 0, chipToTim: num(R.v(sim, "C24")) ?? 0, timToSpreader: num(R.v(sim, "C25")) ?? 0 },
      chips: [],
    };
    for (let r = 29; r <= 88; r++) {
      const refDes = str(R.v(sim, "B" + r));
      if (!refDes) continue;
      const c = { refDes, model: str(R.v(sim, "C" + r)) === "Two-resistor" ? "Two-resistor" : "Block" };
      put(c, "kxy", pos(R.v(sim, "I" + r)));
      simSetup.chips.push(c);
    }
  }

  // ---------- settings ----------
  const settings = {
    tolPosMm: num(R.name("Set_TolPos") ?? (set && R.v(set, "C9"))) ?? 0.05,
    tolSizeMm: num(R.name("Set_TolSize") ?? (set && R.v(set, "C10"))) ?? 0.05,
    import: {
      pcbNameContains: str(R.v(imp, "M3")) || "PCB",
      minSizeMm: num(R.v(imp, "M4")) ?? 0,
      topLevelOnly: yesNo(R.v(imp, "M5")) ?? true,
    },
    flagLimits: {
      critical: { marginBelowC: num(R.v(reg, "AB4")) ?? 15, densityAboveWcm2: num(R.v(reg, "AC4")) ?? 2, shareAbove: num(R.v(reg, "AD4")) ?? 0.25 },
      watch: { marginBelowC: num(R.v(reg, "AB5")) ?? 30, densityAboveWcm2: num(R.v(reg, "AC5")) ?? 0.8, shareAbove: num(R.v(reg, "AD5")) ?? 0.1 },
    },
  };
  if (set) {
    put(settings, "partTemplate", str(R.v(set, "C5")));
    put(settings, "asmTemplate", str(R.v(set, "C6")));
    put(settings, "outputFolder", str(R.v(set, "C7")));
    put(settings, "filePrefix", str(R.v(set, "C8")));
    put(settings, "asmName", str(R.v(set, "C11")));
  }

  // ---------- forSolidWorks: copied from the workbook's cached values ----------
  const heat = {};
  for (let r = regRows[0]; r <= regRows[1]; r++) {
    const refDes = str(R.v(reg, "B" + r));
    const pd = num(R.v(reg, "Y" + r));
    if (refDes && pd !== undefined) heat[refDes] = (heat[refDes] || 0) + pd;
  }
  const fsChips = [];
  if (sim) {
    for (let r = 29; r <= 88; r++) {
      const refDes = str(R.v(sim, "B" + r));
      if (!refDes) continue;
      const c = { refDes, model: str(R.v(sim, "C" + r)) === "Two-resistor" ? "Two-resistor" : "Block" };
      put(c, "kz", pos(R.v(sim, "H" + r)));
      put(c, "kxy", pos(R.v(sim, "I" + r)) ?? c.kz);
      put(c, "rjcCW", pos(R.v(sim, "D" + r)));
      put(c, "rjbCW", pos(R.v(sim, "E" + r)));
      fsChips.push(c);
    }
  }
  const forSolidWorks = {
    computedAt: now,
    computedBy: "workbook",
    ambientC: board.ambientC,
    pcb: {
      l: pos(R.v(geo, "E7")), w: pos(R.v(geo, "E8")), h: pos(R.v(geo, "E9")),
      kxy: pos(R.name("Geo_kXY") ?? R.v(geo, "D19")), kz: pos(R.name("Geo_kZ") ?? R.v(geo, "D20")),
    },
    bodies: fsBodies,
    heat,
    chips: fsChips,
  };
  if (Object.keys(fsPlates).length) forSolidWorks.plates = fsPlates;

  // ---------- cad.json ----------
  const cad = { schema: "pcb-thermal/cad", schemaVersion: 1, savedBy, savedAt: now };
  const imported = importFrom(R, imp, notes);
  if (imported) cad.import = imported;
  const prev = impPrev && importFrom(R, impPrev, []);
  if (prev) cad.importPrevious = prev;

  const bb = options.reports?.buildBoard;
  if (bb) {
    const build = parseBuildReport(bb, geoByName, fsBodies, notes);
    if (build) cad.build = build;
  }
  const bs = options.reports?.buildSpreader;
  if (bs) {
    const sb = parseSpreaderReport(bs);
    if (sb) cad.spreaderBuild = sb;
  }
  const asBuilt = asBuiltFrom(R, spr, sprB);
  if (asBuilt) cad.asBuilt = asBuilt;

  const checks = {};
  if (log) {
    cad.changes = [];
    const seen = new Map();
    let empty = 0;
    for (let r = 8; r <= 5015 && empty < 20; r++) {
      const dRaw = R.v(log, "B" + r);
      const change = str(R.v(log, "G" + r));
      if (dRaw === undefined && !change) { empty++; continue; }
      empty = 0;
      const date = R.date(dRaw) || now;
      const stamp = date.replace(/[-:]/g, "").slice(0, 15).padEnd(15, "0");
      const n = (seen.get(stamp) || 0) + 1;
      seen.set(stamp, n);
      const c = { id: `${stamp}-${n}`, date, macro: str(R.v(log, "C" + r)) || "unknown", change: change || "change" };
      put(c, "key", str(R.v(log, "D" + r)));
      put(c, "name", str(R.v(log, "E" + r)));
      put(c, "refDes", str(R.v(log, "F" + r)));
      put(c, "field", str(R.v(log, "H" + r)));
      put(c, "old", scalar(R.v(log, "I" + r)));
      put(c, "new", scalar(R.v(log, "J" + r)));
      put(c, "delta", num(R.v(log, "K" + r)));
      cad.changes.push(c);
      if (str(R.v(log, "L" + r)) === "Yes") checks[c.id] = { at: now };
    }
  }

  // ---------- results.json ----------
  const results = { schema: "pcb-thermal/results", schemaVersion: 1, savedBy, savedAt: now, runs: [] };
  if (res) {
    const hdr = {};
    for (let c = 17; c <= 400; c++) {
      const h = str(R.v(res, colName(c) + "7"));
      if (!h) continue;
      const m = h.match(/^(.+?) (peak|avg) \(C\)$/);
      if (m) hdr[c] = m;
    }
    const byName = new Map();
    for (let r = 8; r <= 1000; r++) {
      const name = str(R.v(res, "B" + r));
      if (!name) { if (!str(R.v(res, "B" + (r + 1)))) break; continue; }
      const run = { name, date: R.date(R.v(res, "C" + r)) || now, studyType: name.startsWith("FL_") ? "Flow" : "Thermal", status: "read" };
      put(run, "study", str(R.v(res, "D" + r)));
      const sf = str(R.v(res, "E" + r));
      if (sf) run.spreaderFiles = sf.split(/\s*[;,]\s*/);
      const mesh = {};
      put(mesh, "globalMm", pos(R.v(res, "F" + r)));
      put(mesh, "fineMm", pos(R.v(res, "G" + r)));
      if (Object.keys(mesh).length) run.mesh = mesh;
      const summary = {};
      put(summary, "hSpreaderTop", num(R.v(res, "H" + r)));
      put(summary, "hPcbTop", num(R.v(res, "I" + r)));
      put(summary, "timModes", str(R.v(res, "J" + r)));
      if (Object.keys(summary).length) run.inputs = { summary };
      const sMax = num(R.v(res, "K" + r)), sMin = num(R.v(res, "L" + r));
      if (sMax !== undefined && sMin !== undefined) run.spreader = { maxC: sMax, minC: sMin };
      const pMax = num(R.v(res, "M" + r)), pMin = num(R.v(res, "N" + r));
      if (pMax !== undefined && pMin !== undefined) run.pcb = { maxC: pMax, minC: pMin };
      put(run, "energyBalancePct", num(R.v(res, "O" + r)));
      put(run, "plotFile", str(R.v(res, "P" + r)));
      const pts = {};
      for (const [c, m] of Object.entries(hdr)) {
        const t = num(R.v(res, colName(+c) + r));
        if (t === undefined) continue;
        pts[m[1]] = pts[m[1]] || {};
        pts[m[1]][m[2] === "peak" ? "peakC" : "avgC"] = t;
      }
      for (const [k, p] of Object.entries(pts)) if (p.peakC === undefined) delete pts[k];
      if (Object.keys(pts).length) run.parts = pts;
      if (byName.has(name)) {
        notes.push(`Sim results: run ${name} is listed twice; the later row was kept.`);
        results.runs[byName.get(name)] = run;
      } else {
        byName.set(name, results.runs.length);
        results.runs.push(run);
      }
    }
  }

  // ---------- project.json ----------
  const project = {
    schema: "pcb-thermal/project", schemaVersion: 1, savedBy, savedAt: now,
    importedFrom: { file: options.fileName || "workbook", date: now },
    board, parts, bodies,
  };
  if (Object.keys(pcbOverride).length) project.pcbOverride = pcbOverride;
  project.pcbStack = pcbStack;
  if (Object.keys(spreader).length) project.spreader = spreader;
  if (Object.keys(handCheck).length) project.handCheck = handCheck;
  if (simSetup) project.simSetup = simSetup;
  project.settings = settings;
  if (Object.values(forSolidWorks.pcb).every(x => x !== undefined)) project.forSolidWorks = forSolidWorks;
  else notes.push("PCB size or k is blank (no STEP import yet), so forSolidWorks was left out.");
  project.checks = checks;
  if (notes.length) project.importedFrom.notes = notes;

  return { project, cad, results, notes };
}

// ---------------------------------------------------------------- sections

function plateFrom(R, ws) {
  const p = {
    enabled: false,
    material: str(R.v(ws, "C5")) || "Aluminium 6061",
    k: pos(R.v(ws, "C6")) ?? 167,
    densityKgM3: pos(R.v(ws, "C7")) ?? 2700,
    thicknessMm: pos(R.v(ws, "C8")) ?? 1.5,
    clearanceXYMm: num(R.v(ws, "C9")) ?? 1,
    clearanceZMm: num(R.v(ws, "C10")) ?? 1,
    minPedestalMm: num(R.v(ws, "C11")) ?? 0.3,
    heightMode: str(R.v(ws, "C12")) === "Manual" ? "Manual" : "Auto",
    chips: [],
  };
  put(p, "heightManualMm", num(R.v(ws, "C13")));
  for (let r = 20; r <= 79; r++) {
    const refDes = str(R.v(ws, "B" + r));
    if (!refDes) continue;
    const c = { refDes, covered: yesNo(R.v(ws, "I" + r)) ?? false };
    const hm = str(R.v(ws, "J" + r));
    if (hm === "Pedestal" || hm === "Thick TIM") c.heightMode = hm;
    const tm = str(R.v(ws, "K" + r));
    if (tm === "Body" || tm === "Contact R") c.timMode = tm;
    put(c, "timTMm", pos(R.v(ws, "L" + r)));
    put(c, "timK", pos(R.v(ws, "M" + r)));
    if (c.covered) p.enabled = true;
    p.chips.push(c);
  }
  return p;
}

function plateForSw(R, ws, plate) {
  if (!plate.enabled) return null;
  const out = {
    faceZMm: num(R.v(ws, "C14")) ?? 0,
    material: plate.material, k: plate.k, densityKgM3: plate.densityKgM3,
    thicknessMm: plate.thicknessMm, clearanceXYMm: plate.clearanceXYMm,
    chips: [], cuts: [],
  };
  for (let r = 20; r <= 79; r++) {
    if (yesNo(R.v(ws, "I" + r)) !== true) continue;
    const c = {
      refDes: str(R.v(ws, "B" + r)), key: str(R.v(ws, "C" + r)),
      timTMm: num(R.v(ws, "O" + r)) ?? 0, pedestalMm: Math.max(0, round(num(R.v(ws, "P" + r)) ?? 0, 9)),
      timMode: str(R.v(ws, "K" + r)) === "Contact R" ? "Contact R" : "Body",
    };
    c.timTMm = round(c.timTMm, 9);
    put(c, "timK", pos(R.v(ws, "M" + r)));
    put(c, "timRCW", pos(R.v(ws, "Q" + r)));
    if (c.refDes && c.key) out.chips.push(c);
  }
  for (let r = 85; r <= 144; r++) {
    const key = str(R.v(ws, "C" + r));
    if (!key) continue;
    const c = { key, name: str(R.v(ws, "D" + r)) || key, xMin: num(R.v(ws, "E" + r)), yMin: num(R.v(ws, "F" + r)),
      l: pos(R.v(ws, "G" + r)), w: pos(R.v(ws, "H" + r)), edgeZMm: num(R.v(ws, "I" + r)) };
    if (Object.values(c).every(x => x !== undefined)) out.cuts.push(c);
  }
  return out;
}

function solutionFrom(R, ws) {
  const s = {};
  put(s, "plateLMm", pos(R.v(ws, "D6")));
  put(s, "plateWMm", pos(R.v(ws, "D7")));
  put(s, "plateTMm", pos(R.v(ws, "D8")));
  put(s, "plateK", pos(R.v(ws, "D9")));
  put(s, "hAir", pos(R.v(ws, "D10")));
  const f = int(R.v(ws, "D11"));
  if (f === 1 || f === 2) s.faces = f;
  put(s, "rPlateToAirOverrideCW", pos(R.v(ws, "D12")));
  const parts = {};
  const cols = [["tim1TMm", "J"], ["tim1K", "K"], ["localLMm", "L"], ["localWMm", "M"], ["localTMm", "N"],
    ["localK", "O"], ["tim2TMm", "P"], ["tim2K", "Q"], ["plateShareOverrideMm2", "R"]];
  for (let r = 16; r <= 68; r++) {
    const refDes = str(R.v(ws, "B" + r));
    if (!refDes) continue;
    const p = {};
    const u = yesNo(R.v(ws, "H" + r));
    if (u !== undefined) p.underPlate = u;
    for (const [k, c] of cols) put(p, k, pos(R.v(ws, c + r)));
    if (Object.keys(p).length) parts[refDes] = p;
  }
  if (Object.keys(parts).length) s.parts = parts;
  return s;
}

function importFrom(R, ws, notes) {
  const model = str(R.v(ws, "D4"));
  const pcbName = str(R.v(ws, "C8"));
  if (!model || !pcbName) return null;
  const pcbBox = boxFrom(R, ws, 8, "DEFGHI", true);
  if (!pcbBox) { notes.push("STEP import: PCB row has no size; import left out."); return null; }
  const out = {
    model, date: R.date(R.v(ws, "D5")) || isoLocal(new Date()), frame: str(R.v(ws, "D6")) || "",
    settingsUsed: {},
    pcb: { name: pcbName, box: pcbBox },
    bodies: [], messages: [],
  };
  put(out.settingsUsed, "pcbNameContains", str(R.v(ws, "M3")));
  put(out.settingsUsed, "minSizeMm", num(R.v(ws, "M4")));
  const tl = yesNo(R.v(ws, "M5"));
  if (tl !== undefined) out.settingsUsed.topLevelOnly = tl;
  for (let r = 9; r <= 400; r++) {
    const name = str(R.v(ws, "C" + r));
    if (!name) { if (!str(R.v(ws, "C" + (r + 1)))) break; continue; }
    const box = boxFrom(R, ws, r, "DEFGHI", true);
    const key = str(R.v(ws, "K" + r));
    if (!box || !key) { notes.push(`STEP import row ${r}: incomplete; skipped.`); continue; }
    const side = str(R.v(ws, "J" + r));
    const b = { key, name, box, side: ["Top", "Bottom", "Board", "Through"].includes(side) ? side : "Top" };
    const rot = num(R.v(ws, "N" + r));
    if (rot !== undefined) {
      b.rotDeg = rot;
      put(b, "ownL", pos(R.v(ws, "O" + r)));
      put(b, "ownW", pos(R.v(ws, "P" + r)));
    }
    out.bodies.push(b);
  }
  for (let r = 8; r <= 200; r++) {
    const t = str(R.v(ws, "L" + r));
    if (!t) continue;
    out.messages.push({ level: /could not|failed|error|no surface points|not found/i.test(t) ? "warning" : "info", text: t });
  }
  return out;
}

function asBuiltFrom(R, spr, sprB) {
  const one = ws => {
    if (!ws) return null;
    const p = {};
    put(p, "outlineAreaMm2", num(R.v(ws, "G5")));
    put(p, "topAreaMm2", num(R.v(ws, "G6")));
    put(p, "exposedAreaMm2", num(R.v(ws, "G7")));
    put(p, "volumeMm3", num(R.v(ws, "G8")));
    put(p, "massG", num(R.v(ws, "G9")));
    put(p, "thicknessMm", num(R.v(ws, "G10")));
    put(p, "holes", int(R.v(ws, "G11")));
    const chips = [];
    for (let r = 20; r <= 79; r++) {
      const refDes = str(R.v(ws, "B" + r));
      const ped = num(R.v(ws, "U" + r)), tim = num(R.v(ws, "V" + r));
      if (!refDes || (ped === undefined && tim === undefined)) continue;
      const c = { refDes };
      put(c, "pedestalMm", ped);
      put(c, "timTMm", tim);
      chips.push(c);
    }
    if (chips.length) p.chips = chips;
    return Object.keys(p).length ? { p, date: R.date(R.v(ws, "G12")) } : null;
  };
  const top = one(spr), bot = one(sprB);
  if (!top && !bot) return null;
  const out = { date: top?.date || bot?.date || isoLocal(new Date()), plates: {} };
  if (top) out.plates.top = top.p;
  if (bot) out.plates.bottom = bot.p;
  return out;
}

// ---------------------------------------------------------------- reports

export function parseBuildReport(text, nameToKey, fsBodies, notes = []) {
  const lines = text.split(/\r?\n/);
  const m1 = lines[0]?.match(/^BuildBoard report (\d{4}-\d{2}-\d{2}) (\d{2}:\d{2})/);
  if (!m1) { notes.push("BuildBoard report: first line not recognised; build left out."); return null; }
  const field = label => (lines.find(l => l.startsWith(label + ":")) || "").slice(label.length + 1).trim();
  const out = { assembly: field("Assembly"), date: `${m1[1]}T${m1[2]}:00`, mode: field("Mode") === "create" ? "create" : "update", counts: {}, bodies: [], overlaps: [] };
  const cm = text.match(/Parts wanted: (\d+)\s+added (\d+)\s+moved (\d+)\s+resized (\d+)\s+suppressed (\d+)\s+brought back (\d+)\s+failed (\d+)/);
  if (cm) ["wanted", "added", "moved", "resized", "suppressed", "broughtBack", "failed"].forEach((k, i) => out.counts[k] = +cm[i + 1]);
  const rm = text.match(/Rotated footprints: (\d+)/);
  if (rm) out.counts.rotated = +rm[1];

  const names = [...nameToKey.keys()].sort((a, b) => b.length - a.length);
  const keyOf = n => nameToKey.get(n);
  const section = title => {
    const i = lines.findIndex(l => l.startsWith(title));
    if (i < 0) return [];
    const out2 = [];
    for (let j = i + 1; j < lines.length && lines[j].trim() !== ""; j++) out2.push(lines[j]);
    return out2.filter(l => l.startsWith("  ") && !l.startsWith("   "));
  };
  const failed = new Map();
  for (const l of section("FAILED")) {
    const s = l.trim();
    const ca = s.match(/^Could not add (.+) \(([a-z0-9_]+)\): (.+)$/);
    const n = ca ? ca[1] : names.find(x => s.startsWith(x + ":"));
    if (ca) { failed.set(n, ca[3]); if (!nameToKey.has(n)) nameToKey.set(n, ca[2]); }
    else if (n) failed.set(n, s.slice(n.length + 1).trim());
    else notes.push(`BuildBoard report: failed body not matched: ${s}`);
  }
  const cut = new Map();
  for (const l of section("PARTS CUT AT THE BOARD SURFACE")) {
    const m = l.match(/^\s+(.+): box cut at the board surface \(([\d.]+) mm went into the PCB\)/);
    if (m) cut.set(m[1], +m[2]);
  }
  for (const b of fsBodies) {
    const e = { key: b.key, name: b.name, status: failed.has(b.name) ? "failed" : "built" };
    if (failed.has(b.name)) e.reason = failed.get(b.name);
    if (cut.has(b.name)) e.cutAtBoardMm = cut.get(b.name);
    out.bodies.push(e);
  }
  for (const n of failed.keys()) if (!fsBodies.some(b => b.name === n)) out.bodies.push({ key: keyOf(n), name: n, status: "failed", reason: failed.get(n) });
  for (const n of cut.keys()) if (!fsBodies.some(b => b.name === n)) notes.push(`BuildBoard report: cut body is no longer included: ${n}`);
  for (const l of section("OVERLAPPING BOXES")) {
    const s = l.trim();
    const mm = s.match(/ overlap by ([\d.]+) x ([\d.]+) x ([\d.]+) mm$/);
    if (!mm) continue;
    const pair = s.slice(0, mm.index);
    const a = names.find(x => pair.startsWith(x + " and "));
    if (!a) { notes.push(`BuildBoard report: overlap not matched: ${s}`); continue; }
    const b = pair.slice(a.length + 5);
    out.overlaps.push({ a: keyOf(a) || a, b: keyOf(b) || b, size: [+mm[1], +mm[2], +mm[3]] });
  }
  for (const e of out.bodies) if (!e.key) delete e.key;
  return out;
}

export function parseSpreaderReport(text) {
  const lines = text.split(/\r?\n/);
  const m1 = lines[0]?.match(/^BuildSpreader report (\d{4}-\d{2}-\d{2}) (\d{2}:\d{2})/);
  if (!m1) return null;
  const date = `${m1[1]}T${m1[2]}:00`;
  const out = {};
  for (const [title, side] of [["TOP SPREADER:", "top"], ["BOTTOM SPREADER:", "bottom"]]) {
    const i = lines.findIndex(l => l.startsWith(title));
    if (i < 0) continue;
    const p = { date, file: lines[i].slice(title.length).trim(), pedestals: 0, tims: [], cuts: 0, failures: [] };
    for (let j = i + 1; j < lines.length && lines[j].startsWith(" "); j++) {
      const c = lines[j].match(/^\s+(\S+): face Z [-\d.]+, TIM [\d.]+ \((\w+)[^)]*\), pedestal ([\d.]+) mm/);
      if (c) { if (+c[3] > 0) p.pedestals++; if (c[2] === "body") p.tims.push(c[1]); }
      const k = lines[j].match(/^\s+clearance cuts \((\d+)\)/);
      if (k) p.cuts = +k[1];
    }
    out[side] = p;
  }
  const pi = lines.findIndex(l => l.startsWith("PROBLEMS"));
  if (pi >= 0) {
    const probs = [];
    for (let j = pi + 1; j < lines.length && lines[j].startsWith(" "); j++) probs.push(lines[j].trim());
    for (const side of Object.keys(out)) out[side].failures = probs;
  }
  for (const side of Object.keys(out)) if (!out[side].failures.length) delete out[side].failures;
  return Object.keys(out).length ? out : null;
}

// ---------------------------------------------------------------- cell access

function reader(wb, XLSX) {
  const byWords = new Map();
  for (const n of wb.SheetNames) byWords.set(n.replace(/^\d+[a-z]?\s+/, ""), n);
  const names = new Map();
  for (const d of wb.Workbook?.Names || []) if (d.Sheet === undefined) names.set(d.Name, d.Ref);
  const R = {
    sheet(words, required) {
      const n = byWords.get(words);
      if (!n && required) throw new Error(`The workbook has no '${words}' tab. Is it a PCB thermal workbook?`);
      return n ? wb.Sheets[n] : null;
    },
    v(ws, addr) {
      if (!ws) return undefined;
      const c = ws[addr];
      if (!c || c.t === "z" || c.t === "e") return undefined;
      if (c.t === "s" && c.v.trim() === "") return undefined;
      return c.v;
    },
    // Value of a single-cell defined name.
    name(n) {
      const ref = names.get(n);
      if (!ref) return undefined;
      const m = ref.match(/^'?(.+?)'?!\$?([A-Z]+)\$?(\d+)$/);
      if (!m) return undefined;
      return R.v(wb.Sheets[m[1]], m[2] + m[3]);
    },
    // First and last row of a defined range, or the defaults.
    nameRows(n, a, b) {
      const ref = names.get(n);
      const m = ref && ref.match(/\$?[A-Z]+\$?(\d+):\$?[A-Z]+\$?(\d+)$/);
      return m ? [+m[1], +m[2]] : [a, b];
    },
    date(x) {
      if (typeof x === "number") {
        const d = XLSX.SSF.parse_date_code(x);
        const p = v => String(v).padStart(2, "0");
        return `${d.y}-${p(d.m)}-${p(d.d)}T${p(d.H)}:${p(d.M)}:${p(Math.floor(d.S))}`;
      }
      if (x instanceof Date) return isoLocal(x);
      return undefined;
    },
  };
  return R;
}

// ---------------------------------------------------------------- small helpers

function boxFrom(R, ws, r, cols, full) {
  const keys = ["xMin", "yMin", "zMin", "l", "w", "h"];
  const b = {};
  keys.forEach((k, i) => {
    const x = i < 3 ? num(R.v(ws, cols[i] + r)) : pos(R.v(ws, cols[i] + r));
    if (x !== undefined) b[k] = x;
  });
  const n = Object.keys(b).length;
  if (full) return n === 6 ? b : null;
  return n ? b : null;
}
function put(o, k, v) { if (v !== undefined && v !== "") o[k] = v; }
function str(x) { if (x === undefined || x === null) return ""; return String(x).trim(); }
function num(x) {
  if (typeof x === "number" && isFinite(x)) return x;
  if (typeof x === "string" && x.trim() !== "" && isFinite(+x)) return +x;
  return undefined;
}
function pos(x) { const n = num(x); return n !== undefined && n > 0 ? n : undefined; }
function int(x) { const n = num(x); return n !== undefined && Number.isInteger(n) ? n : undefined; }
function scalar(x) { return typeof x === "number" ? x : (x === undefined ? undefined : str(x)); }
function yesNo(x) { const s = str(x).toLowerCase(); return s === "yes" ? true : s === "no" ? false : undefined; }
function round(x, d) { const f = 10 ** d; return Math.round(x * f) / f; }
function baseName(f) { return f ? f.replace(/^.*[\\/]/, "").replace(/\.[^.]+$/, "") : ""; }
function colName(c) { let s = ""; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; }
// Leading whole words shared by all texts: "MxL7704 buck1 > A", "MxL7704 buck2 > B" -> "MxL7704".
function commonWords(texts) {
  const split = texts.map(t => t.split(/\s+/));
  const out = [];
  for (let i = 0; i < split[0].length; i++) {
    const w = split[0][i];
    if (split.every(s => s[i] === w)) out.push(w); else break;
  }
  return out.join(" ");
}
function stripPrefix(t, prefix) {
  if (!t) return "";
  if (prefix && t.startsWith(prefix + " ")) return t.slice(prefix.length + 1);
  return t;
}
export function isoLocal(d) {
  const p = v => String(v).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}
