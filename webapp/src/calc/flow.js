// Flow Simulation settings and checks (docs/FLOW_PLAN.md). The page owns simSetup.flow; RunFlow reads it.
//
//   flowSettings(project)            -> simSetup.flow with the defaults filled in
//   flowDomainMm(fs, settings)       -> computational domain RunFlow sets: { x: [min, max], y, z } in mm
//   flowMinCellMm(settings, domain)  -> estimated smallest solid cell (mm); same rule as MinCellMm in RunFlow.bas
//   flowCheck(project, calc)         -> { settings, domain, minCellMm, thin: [{ refDes, hMm }] }
//   meshCheckResult(run, runs, limit) -> the change per chip between the two runs of a mesh check

import { isNum } from "./util.js";

export const FLOW_DEFAULTS = { meshLevel: 4, localMesh: true, solidRefine: 3, meshCheck: false, meshCheckLimitC: 1 };

export function flowSettings(project) {
  const f = project.simSetup?.flow || {};
  const out = { ...FLOW_DEFAULTS };
  for (const [k, v] of Object.entries(f)) if (v !== undefined) out[k] = v;
  return out;
}

// The box of the PCB and every included body, with RunFlow's margins: half the largest board size at the sides and
// below, one and a half above (warm air rises along +Z).
export function flowDomainMm(fs) {
  if (!fs) return undefined;
  const boxes = [{ xMin: 0, yMin: 0, zMin: 0, l: fs.pcb.l, w: fs.pcb.w, h: fs.pcb.h }, ...fs.bodies.map(b => b.box)];
  const lo = f => Math.min(...boxes.map(b => b[f + "Min"]));
  const hi = (f, s) => Math.max(...boxes.map(b => b[f + "Min"] + b[s]));
  const x = [lo("x"), hi("x", "l")], y = [lo("y"), hi("y", "w")], z = [lo("z"), hi("z", "h")];
  const m = Math.max(x[1] - x[0], y[1] - y[0]);
  return { x: [x[0] - 0.5 * m, x[1] + 0.5 * m], y: [y[0] - 0.5 * m, y[1] + 0.5 * m], z: [z[0] - 0.5 * m, z[1] + 1.5 * m] };
}

// Mean cell of the domain at the global level (calibrated on the RPi4 runs of 2026-10-19: level 3 gave 28 000
// cells, level 5 1.43 M, so about 7 times more cells per level), halved solidRefine times inside a local mesh.
// An estimate: Flow also refines near small solids by itself.
export function flowMinCellMm(s, d) {
  if (!d || !isNum(s.meshLevel)) return undefined;
  const vol = (d.x[1] - d.x[0]) * (d.y[1] - d.y[0]) * (d.z[1] - d.z[0]);
  const cell = Math.cbrt(vol / (28000 * 7 ** (s.meshLevel - 3)));
  return s.localMesh ? cell / 2 ** s.solidRefine : cell;
}

export function flowCheck(project, calc) {
  const settings = flowSettings(project);
  const fs = calc.fsNow;
  const domain = flowDomainMm(fs);
  const minCellMm = flowMinCellMm(settings, domain);
  const thin = [];
  if (fs && isNum(minCellMm)) {
    for (const b of fs.bodies) {
      if (!b.refDes || !(fs.heat[b.refDes] > 0) || !(b.box.h < minCellMm)) continue;
      if (!thin.some(t => t.refDes === b.refDes)) thin.push({ refDes: b.refDes, hMm: b.box.h });
    }
  }
  return { settings, domain, minCellMm, thin };
}

// Mesh check (F2): RunFlow solves at level N and N+1 and stores both runs with meshCheck.pairWith pointing at the
// other one. The change per chip is the finer run's peak minus the coarser run's peak.
export function meshCheckResult(run, runs, limitC) {
  if (!run?.meshCheck?.pairWith) return undefined;
  const other = runs.find(r => r.name === run.meshCheck.pairWith);
  if (!other) return { missing: run.meshCheck.pairWith };
  const lv = r => r.meshCheck?.level ?? r.mesh?.level;
  const [coarse, fine] = (lv(run) ?? 0) <= (lv(other) ?? 0) ? [run, other] : [other, run];
  const parts = {};
  for (const [ref, t] of Object.entries(fine.parts || {})) {
    const c = coarse.parts?.[ref];
    if (isNum(t.peakC) && isNum(c?.peakC)) parts[ref] = t.peakC - c.peakC;
  }
  const refs = Object.keys(parts);
  if (!refs.length) return { coarse: coarse.name, fine: fine.name, parts, ok: undefined };
  const worst = refs.reduce((a, r) => (Math.abs(parts[r]) > Math.abs(parts[a]) ? r : a), refs[0]);
  const over = refs.filter(r => Math.abs(parts[r]) >= limitC);
  return { coarse: coarse.name, fine: fine.name, coarseLevel: lv(coarse), fineLevel: lv(fine), parts, worst, worstC: parts[worst], over, ok: over.length === 0 };
}
