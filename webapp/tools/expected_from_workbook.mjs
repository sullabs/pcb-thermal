// Read every calculated value of a filled workbook, as Excel last computed and saved it, into a JSON file
// the calculation module is tested against.
//   node tools/expected_from_workbook.mjs <workbook.xlsx> <out.json>
// The workbooks in example/ were last saved by Excel (they carry calcChain.xml), so the cached values are
// Excel's own results. No Excel is needed to read them. Reads only; never writes the workbook.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import * as XLSX from "xlsx";

XLSX.set_fs(fs);

export function expectedFromWorkbook(file) {
  const wb = XLSX.readFile(file, { cellFormula: false });
  const sheet = words => {
    const n = wb.SheetNames.find(s => s.replace(/^\d+[a-z]?\s+/, "").toLowerCase() === words.toLowerCase());
    return n ? wb.Sheets[n] : undefined;
  };
  const val = (ws, a) => {
    const c = ws?.[a];
    if (!c || c.t === "e" || c.v === "" || c.v === undefined) return undefined;
    return typeof c.v === "string" ? c.v.trim() || undefined : c.v;
  };
  const num = (ws, a) => { const v = val(ws, a); return typeof v === "number" ? v : undefined; };
  const obj = (ws, r, cols) => {
    const o = {};
    for (const [k, c] of Object.entries(cols)) { const v = val(ws, c + r); if (v !== undefined) o[k] = v; }
    return o;
  };
  const out = { source: path.basename(file),
    note: "Values Excel computed and saved in the workbook. Row tabs are keyed by RefDes (register and thermal solution: one entry per register row)." };

  // 1 Component register (one entry per row)
  const reg = sheet("Component register");
  out.register = { totalPdW: num(reg, "D4"), rows: [] };
  for (let r = 8; r <= 53; r++) {
    const refDes = val(reg, "B" + r);
    if (!refDes) continue;
    out.register.rows.push({ refDes, row: r, ...obj(reg, r, { method: "T", pinW: "U", poutW: "V", areaMm2: "W", pdCalc: "X",
      pdW: "Y", taC: "Z", riseC: "AA", tjEstC: "AB", marginC: "AC", densityWcm2: "AD", share: "AE", flag: "AF", why: "AG",
      rankPd: "AH", rankDensity: "AI" }) });
  }

  // 4 Triage summary
  const tri = sheet("Triage");
  out.triage = { critical: num(tri, "D5"), watch: num(tri, "D6"), minor: num(tri, "D7"), noPd: num(tri, "D8"),
    rows: num(tri, "H5"), totalPdW: num(tri, "H6"), noRja: num(tri, "H7"), noBody: num(tri, "H8") };

  // 5 Rail budget
  const rb = sheet("Rail budget");
  out.rails = { rails: [], totalPdW: num(rb, "E30"), sumRailsW: num(rb, "E31"), noRailW: num(rb, "E32"), rowsNoRail: num(rb, "E33"),
    measuredW: num(rb, "E35"), unaccountedW: num(rb, "E36"), unaccountedFrac: num(rb, "E37") };
  for (let r = 7; r <= 28; r++) {
    const rail = val(rb, "C" + r);
    if (rail) out.rails.rails.push({ rail, ...obj(rb, r, { rows: "D", pdW: "E", share: "F", noPd: "G" }) });
  }

  // 3 CAD geometry
  const geo = sheet("CAD geometry");
  out.pcb = { l: num(geo, "E7"), w: num(geo, "E8"), h: num(geo, "E9"), copperTotalMm: num(geo, "D17"),
    copperFraction: num(geo, "D18"), kxy: num(geo, "D19"), kz: num(geo, "D20") };
  out.geometry = { imported: num(geo, "M6"), included: num(geo, "M7"), registerPdW: num(geo, "M8"), placedPdW: num(geo, "M9"),
    notPlacedW: num(geo, "M10"), partsWithNoBody: num(geo, "M11"), problems: num(geo, "M12"), bodies: {} };
  for (let r = 26; r <= 185; r++) {
    const key = val(geo, "AG" + r);
    if (!key || !val(geo, "E" + r)) continue;
    out.geometry.bodies[key] = { row: r, ...obj(geo, r, { name: "E", refDes: "F", include: "G", xMin: "U", yMin: "V", zMin: "W",
      l: "X", w: "Y", h: "Z", xCentre: "AA", yCentre: "AB", pdW: "AC", densityWcm2: "AD", check: "AE", cutTop: "AH", cutBottom: "AL" }) };
  }

  // 6 Bare board network
  const net = sheet("Bare board network");
  out.bareBoard = {
    board: { ambientC: num(net, "L7"), areaMm2: num(net, "L8"), hRadTop: num(net, "L9"), hRadBottom: num(net, "L10"),
      hRadParts: num(net, "L11"), hPcb: num(net, "L12"), hParts: num(net, "L13"), rBoardAirCW: num(net, "L14"),
      totalPdW: num(net, "L15"), riseC: num(net, "L16"), boardC: num(net, "L17"), qBoardW: num(net, "L18"),
      qSurfacesW: num(net, "L19"), fail: num(net, "L20"), tight: num(net, "L21") },
    compare: { run: val(net, "R8"), check: val(net, "R12"), compared: num(net, "R13"), meanGapC: num(net, "R14"), maxGapC: num(net, "R15") },
    parts: {},
  };
  for (let r = 26; r <= 71; r++) {
    const refDes = val(net, "C" + r);
    if (!refDes) continue;
    out.bareBoard.parts[refDes] = obj(net, r, { pdW: "E", lMm: "F", wMm: "G", hMm: "H", rjcUsedCW: "M", model: "N", rJTopCW: "O",
      rJBoardCW: "P", exposedMm2: "Q", rSurfaceAirCW: "R", rTopCW: "S", footprintMm2: "T", shareMm2: "U", rSpreadCW: "V",
      rBoardCW: "W", f: "X", cWK: "Y", fPdW: "Z", qBoardW: "AA", qTopW: "AB", tjC: "AC", marginC: "AD", status: "AE",
      tjRjaC: "AF", simC: "AG", gapC: "AH", eps: "AJ", tau: "AK", bi: "AL", lambda: "AM", phi: "AN" });
  }

  // 7 Spreader, 7b Spreader bottom
  out.spreader = {};
  for (const [side, words] of [["top", "Spreader"], ["bottom", "Spreader bottom"]]) {
    const ws = sheet(words);
    if (!ws) continue;
    const s = { faceZMm: num(ws, "C14"), covered: num(ws, "C15"), cutCount: num(ws, "C16"), chips: {}, cuts: [] };
    for (let r = 20; r <= 79; r++) {
      const refDes = val(ws, "B" + r);
      if (refDes) s.chips[refDes] = obj(ws, r, { key: "C", xMin: "D", yMin: "E", l: "F", w: "G", faceZ: "H", covered: "I",
        gapMm: "N", timTMm: "O", pedestalMm: "P", timRCW: "Q", check: "R", zNeed: "S" });
    }
    for (let r = 85; r <= 144; r++) {
      const key = val(ws, "C" + r);
      if (key) s.cuts.push(obj(ws, r, { key: "C", name: "D", xMin: "E", yMin: "F", l: "G", w: "H", edgeZMm: "I" }));
    }
    out.spreader[side] = s;
  }

  // 8 Thermal solution, 8b Thermal solution bottom (one entry per register row)
  out.solution = {};
  for (const [side, words] of [["top", "Thermal solution"], ["bottom", "Thermal solution bottom"]]) {
    const ws = sheet(words);
    if (!ws) continue;
    const s = { ambientC: num(ws, "J5"), plateAreaMm2: num(ws, "J6"), partsUnder: num(ws, "J7"), footprintsMm2: num(ws, "J8"),
      hEff: num(ws, "J9"), rPlateAirCW: num(ws, "J10"), pdPlateW: num(ws, "J11"), plateC: num(ws, "J12"), warning: val(ws, "K8"), rows: [] };
    for (let r = 16; r <= 68; r++) {
      const refDes = val(ws, "B" + r);
      if (!refDes) continue;
      s.rows.push({ refDes, row: r, ...obj(ws, r, { pdW: "D", packageMm2: "S", footprintMm2: "T", plateShareMm2: "U", rjcCW: "V",
        rTim1CW: "W", rLocalCW: "X", rTim2CW: "Y", rPlateSpreadCW: "Z", rStackCW: "AA", dtStackC: "AB", tjC: "AC", marginC: "AD",
        status: "AE" }) });
    }
    out.solution[side] = s;
  }

  // 9 Sim setup
  const sim = sheet("Sim setup");
  out.simChips = {};
  for (let r = 29; r <= 88; r++) {
    const refDes = val(sim, "B" + r);
    if (refDes) out.simChips[refDes] = obj(sim, r, { rjcCW: "D", rjbCW: "E", hMm: "F", areaMm2: "G", kz: "H", kxy: "I", pdW: "J", check: "K" });
  }
  return out;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [src, dst] = process.argv.slice(2);
  if (!src || !dst) { console.error("Usage: node tools/expected_from_workbook.mjs <workbook.xlsx> <out.json>"); process.exit(2); }
  fs.mkdirSync(path.dirname(dst), { recursive: true });
  fs.writeFileSync(dst, JSON.stringify(expectedFromWorkbook(src), null, 1) + "\n", "utf8");
  console.log("Wrote " + dst);
}
