// Bare board network (6 Bare board network): each part has a top path (junction to case top to air) and a
// board path (junction to board, spreading into its share of the PCB). All parts heat one shared board node.

import { STEFAN_BOLTZMANN } from "./constants.js";
import { leeSpreading } from "./lee.js";
import { bodyOf } from "./geometry.js";
import { isNum, N, sum } from "./util.js";

// h of radiation, linearised around the ambient and ambient + rise (W/(m2 K)).
export function hRadiation(emissivity, ambientC, riseK) {
  if (!isNum(emissivity)) return 0;
  const t1 = ambientC + 273.15, t2 = ambientC + N(riseK) + 273.15;
  return emissivity * STEFAN_BOLTZMANN * (t1 ** 2 + t2 ** 2) * (t1 + t2);
}

export function bareBoard(project, results, rows, board, reg) {
  const s = project.simSetup?.surfaces || {};
  const rise = project.handCheck?.surfaceRiseK;
  const kParts = project.simSetup?.kPartsWithoutHeat;
  const ambient = project.board.ambientC;
  const tightBelow = project.settings.flagLimits.critical.marginBelowC;

  const b = { ambientC: ambient };
  if (isNum(board.l) && isNum(board.w)) b.areaMm2 = board.l * board.w;
  b.hRadTop = hRadiation(s.pcbTop?.emissivity, ambient, rise);
  b.hRadBottom = hRadiation(s.pcbBottom?.emissivity, ambient, rise);
  b.hRadParts = hRadiation(s.parts?.emissivity, ambient, rise);
  b.hPcb = N(s.pcbTop?.h) + b.hRadTop + N(s.pcbBottom?.h) + b.hRadBottom;
  b.hParts = N(s.parts?.h) + b.hRadParts;
  if (isNum(b.areaMm2) && b.hPcb > 0) b.rBoardAirCW = 1e6 / (b.hPcb * b.areaMm2);

  const parts = {};
  for (const p of project.parts) {
    const r = { pdW: N(reg.parts[p.refDes].pdW) };
    const body = bodyOf(rows, p.refDes);
    r.hasBody = !!body;
    if (body) { r.lMm = body.box.l; r.wMm = body.box.w; r.hMm = N(body.box.h); }
    const L = N(r.lMm), W = N(r.wMm), H = N(r.hMm);
    r.rjcUsedCW = N(p.rjcCW) > 0 ? p.rjcCW : H > 0 && L * W > 0 && N(kParts) > 0 ? (H * 1000) / (kParts * L * W) : 0;
    const twoR = N(p.rjbCW) > 0;
    r.model = twoR ? "Two-resistor" : N(p.rjcCW) > 0 ? "Block" : "No Rjc: part k";
    r.rJTopCW = twoR ? r.rjcUsedCW : r.rjcUsedCW / 2;
    r.rJBoardCW = twoR ? p.rjbCW : r.rjcUsedCW / 2;
    if (body) r.exposedMm2 = L * W + 2 * (L + W) * H;
    if (N(r.exposedMm2) > 0 && b.hParts > 0) r.rSurfaceAirCW = 1e6 / (b.hParts * r.exposedMm2);
    if (isNum(r.rSurfaceAirCW)) r.rTopCW = r.rJTopCW + r.rSurfaceAirCW;
    if (body) r.footprintMm2 = L * W;
    parts[p.refDes] = r;
  }
  b.totalPdW = sum(Object.values(parts).map(r => r.pdW));

  for (const p of project.parts) {
    const r = parts[p.refDes];
    if (r.pdW > 0 && b.totalPdW > 0 && isNum(b.areaMm2)) r.shareMm2 = (b.areaMm2 * r.pdW) / b.totalPdW;
    if (N(r.footprintMm2) > 0 && N(r.shareMm2) > 0 && N(board.h) > 0 && N(board.kxy) > 0 && b.hPcb > 0) {
      r.spread = leeSpreading({ sourceMm2: r.footprintMm2, plateMm2: r.shareMm2, tMm: board.h, k: board.kxy, h: b.hPcb });
      r.rSpreadCW = r.spread.r;
      r.rBoardCW = r.rJBoardCW + r.rSpreadCW;
      r.f = isNum(r.rTopCW) ? r.rTopCW / (r.rTopCW + r.rBoardCW) : 1;
      r.cWK = isNum(r.rTopCW) ? 1 / (r.rTopCW + r.rBoardCW) : 0;
      r.fPdW = r.f * r.pdW;
    }
  }
  const pl = Object.values(parts);
  if (isNum(b.rBoardAirCW)) {
    b.riseC = sum(pl.map(r => r.fPdW)) / (1 / b.rBoardAirCW + sum(pl.map(r => r.cWK)));
    b.boardC = ambient + b.riseC;
  }
  for (const p of project.parts) {
    const r = parts[p.refDes];
    if (isNum(r.f) && isNum(b.riseC)) {
      r.qBoardW = r.f * r.pdW - r.cWK * b.riseC;
      r.qTopW = r.pdW - r.qBoardW;
      r.tjC = isNum(r.rTopCW) ? ambient + r.qTopW * r.rTopCW : b.boardC + r.qBoardW * r.rBoardCW;
      if (isNum(p.tjMaxC)) r.marginC = p.tjMaxC - r.tjC;
    }
    r.status = statusOf(r, p, b, tightBelow);
    if (r.pdW > 0 && isNum(p.rjaCW)) r.tjRjaC = ambient + r.pdW * p.rjaCW;
  }
  if (pl.some(r => isNum(r.qBoardW))) {
    b.qBoardW = sum(pl.map(r => r.qBoardW));
    b.qSurfacesW = b.totalPdW - b.qBoardW;
  }
  b.fail = pl.filter(r => r.status === "FAIL").length;
  b.tight = pl.filter(r => r.status === "Tight").length;
  return { board: b, parts, compare: compareRun(project, results, parts) };
}

function statusOf(r, p, b, tightBelow) {
  if (!(r.pdW > 0)) return "no Pd: fill this part on the register";
  if (!r.hasBody) return "no body: set RefDes on CAD geometry col F";
  if (!(N(r.footprintMm2) > 0)) return "body has no L x W on CAD geometry";
  if (!isNum(b.areaMm2)) return "no PCB size on CAD geometry";
  if (!isNum(r.tjC)) return "missing input";
  if (!isNum(r.marginC)) return "no Tj max on the register";
  return r.marginC < 0 ? "FAIL" : r.marginC < tightBelow ? "Tight" : "OK";
}

// Hand vs simulation: the named run, or the newest one.
function compareRun(project, results, parts) {
  const runs = results?.runs || [];
  if (!runs.length) return { check: "no runs yet" };
  const want = project.handCheck?.compareRun;
  const run = want ? [...runs].reverse().find(r => r.name === want) : runs[runs.length - 1];
  if (!run) return { check: "run not found" };
  const out = { run: run.name, date: run.date };
  if (run.spreaderFiles?.length) out.spreader = run.spreaderFiles.join(", ");
  out.check = out.spreader ? "run had a spreader: pick a bare board run"
    : run.studyType === "Flow" ? "Flow run: check the Flow project has no spreader" : "bare board run, OK to compare";
  out.parts = {};
  const gaps = [];
  for (const [ref, r] of Object.entries(parts)) {
    const sim = run.parts?.[ref]?.peakC;
    if (!isNum(sim)) continue;
    const o = { simC: sim };
    if (isNum(r.tjC)) { o.gapC = r.tjC - sim; gaps.push(o.gapC); }
    out.parts[ref] = o;
  }
  out.compared = Object.keys(out.parts).length;
  if (gaps.length) {
    out.meanGapC = sum(gaps) / gaps.length;
    out.maxGapC = Math.max(Math.max(...gaps), -Math.min(...gaps));
  }
  return out;
}
