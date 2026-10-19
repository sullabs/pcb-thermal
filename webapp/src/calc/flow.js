// Flow Simulation settings and checks (docs/FLOW_PLAN.md). The page owns simSetup.flow; RunFlow reads it.
//
//   flowSettings(project)            -> simSetup.flow with the defaults filled in
//   flowDomainMm(fs, settings)       -> computational domain RunFlow sets: { x: [min, max], y, z } in mm
//   flowMinCellMm(settings, domain)  -> estimated smallest solid cell (mm); same rule as MinCellMm in RunFlow.bas
//   flowCheck(project, calc)         -> { settings, domain, minCellMm, thin: [{ refDes, hMm }] }
//   meshCheckResult(run, runs, limit) -> the change per chip between the two runs of a mesh check
//   kStep(k), flowMaterial(kxy, kz)  -> k in 10 % steps (24 per decade) and the user material named after it
//   flowForSw(project, calc, fs)     -> forSolidWorks.flow: per chip its Tj basis and material(s), and the materials list
//   flowEmissivity(e)               -> radiative surface named after the emissivity (2 decimals)
//   tjBasisOf(run, ref)              -> "calibrated block" | "two-resistor" | "package" for a run's part

import { isNum } from "./util.js";

export const FLOW_DEFAULTS = { meshLevel: 4, localMesh: true, solidRefine: 3, meshCheck: false, meshCheckLimitC: 1, minTravels: 2, goalCriteriaC: 0.2 };

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

// Flow cannot take a k from the API (F0, S4): only a material path from its Engineering Database. So every k is
// rounded to a step of 10 % and the material is named after it; the same step is reused by every project.
export const kStep = k => +(10 ** (Math.round(24 * Math.log10(k)) / 24)).toPrecision(3);
export function flowMaterial(kxy, kz = kxy) {
  const a = kStep(kxy), b = kStep(kz);
  const name = a === b ? `k ${a}` : `k ${a} ${b}`;
  return { name, path: `User Defined\\PCBThermal\\${name}`, kxy: a, kz: b };
}

export function flowEmissivity(e) {
  const v = Math.round(e * 100) / 100, name = `e ${v.toFixed(2)}`;
  return { name, path: `User Defined\\PCBThermal\\${name}`, emissivity: v };
}

// Chips with heat in the Flow project: a block with the Rjc k (calibrated block), the three bodies of a two-resistor
// chip (Run Thermal Study makes them; layer k = t / (theta × A), core kCore), or the generic package material.
export function flowForSw(project, calc, fs) {
  const mats = new Map(), chips = {};
  const use = (m, who) => { if (!mats.has(m.path)) mats.set(m.path, { ...m, usedBy: [] }); mats.get(m.path).usedBy.push(who); return m.path; };
  const ss = project.simSetup || {};
  const frac = isNum(ss.layerFrac) ? ss.layerFrac : 0.1, kCore = isNum(ss.kCore) ? ss.kCore : 100;
  for (const [ref, q] of Object.entries(fs.heat)) {
    if (!(q > 0) || !fs.bodies.some(b => b.refDes === ref)) continue;
    const c = calc.simChips.find(x => x.refDes === ref);
    if (c?.model === "Two-resistor" && c.rjcCW > 0 && c.rjbCW > 0 && c.hMm > 0 && c.areaMm2 > 0) {
      const t = frac * c.hMm, k = th => (t * 1000) / (th * c.areaMm2);
      chips[ref] = { basis: "two-resistor", layers: { bottom: use(flowMaterial(k(c.rjbCW)), `${ref} bottom`), core: use(flowMaterial(kCore), `${ref} core`), top: use(flowMaterial(k(c.rjcCW)), `${ref} top`) } };
    } else if (c && isNum(c.kz) && c.kz > 0) {
      chips[ref] = { basis: "calibrated block", material: use(flowMaterial(c.kxy > 0 ? c.kxy : c.kz, c.kz), ref) };
    } else chips[ref] = { basis: "package" };
  }
  // F5: PCB, TIMs, spreaders, radiative surfaces and contacts from the page, not Flow's library picks
  const out = { chips };
  const notes = [];
  out.pcb = use(flowMaterial(fs.pcb.kxy, fs.pcb.kz), "PCB");
  const ct = ss.contacts || {};
  const tims = {}, spreaders = {};
  let timArea = 0;
  for (const side of ["top", "bottom"]) {
    const pl = fs.plates?.[side];
    if (!pl) continue;
    if (pl.k > 0) spreaders[side] = use(flowMaterial(pl.k), `${side} spreader`);
    for (const ch of pl.chips) {
      if (ch.timMode !== "Body" || !(ch.timK > 0) || !(ch.timTMm > 0)) continue;
      const tm = ch.timTMm / 1000, rc = ct.chipToTim > 0 ? ct.chipToTim : 0;
      tims[ch.refDes] = use(flowMaterial(tm / (tm / ch.timK + rc)), `TIM ${ch.refDes}`);
      if (ch.timRCW > 0) timArea += (1000 * ch.timTMm) / (ch.timK * ch.timRCW);
    }
  }
  if (Object.keys(tims).length) out.tims = tims;
  if (Object.keys(spreaders).length) out.spreaders = spreaders;
  if (ct.chipToTim > 0 && Object.keys(tims).length) notes.push(`chip to TIM contact ${ct.chipToTim} K m²/W is folded into each TIM material: k = t / (t / k TIM + R contact)`);
  const sf = ss.surfaces || {}, rads = new Map();
  const rad = (e, who) => { const r = flowEmissivity(e); if (!rads.has(r.path)) rads.set(r.path, { ...r, usedBy: [] }); rads.get(r.path).usedBy.push(who); return r.path; };
  const surfaces = {};
  if (isNum(sf.pcbTop?.emissivity)) surfaces.pcb = rad(sf.pcbTop.emissivity, "PCB");
  if (isNum(sf.parts?.emissivity)) surfaces.parts = rad(sf.parts.emissivity, "parts");
  if (fs.plates && isNum(sf.spreaderTop?.emissivity)) surfaces.spreader = rad(sf.spreaderTop.emissivity, "spreader");
  if (Object.keys(surfaces).length) out.surfaces = surfaces;
  if (isNum(sf.pcbTop?.emissivity) && isNum(sf.pcbBottom?.emissivity) && sf.pcbTop.emissivity !== sf.pcbBottom.emissivity) notes.push(`PCB emissivity: top ${sf.pcbTop.emissivity}, bottom ${sf.pcbBottom.emissivity}; Flow gets the top value on the whole PCB`);
  if (fs.plates && isNum(sf.spreaderTop?.emissivity) && isNum(sf.spreaderOther?.emissivity) && sf.spreaderTop.emissivity !== sf.spreaderOther.emissivity) notes.push(`spreader emissivity: top ${sf.spreaderTop.emissivity}, other faces ${sf.spreaderOther.emissivity}; Flow gets the top value on the whole spreader`);
  const contacts = {};
  if (ct.partToPcb > 0) {
    const onPcb = fs.bodies.filter(b => Math.abs(b.box.zMin - fs.pcb.h) < 0.05 || Math.abs(b.box.zMin + b.box.h) < 0.05);
    const a = onPcb.reduce((s, b) => s + b.box.l * b.box.w, 0);
    if (a > 0) contacts.partToPcb = { rKW: ct.partToPcb / (a * 1e-6), areaMm2: a, rSpecific: ct.partToPcb };
  }
  if (ct.timToSpreader > 0 && timArea > 0) contacts.timToSpreader = { rKW: ct.timToSpreader / (timArea * 1e-6), areaMm2: timArea, rSpecific: ct.timToSpreader };
  if (Object.keys(contacts).length) out.contacts = contacts;
  if (notes.length) out.notes = notes;
  out.radiative = [...rads.values()];
  if (!out.radiative.length) delete out.radiative;
  out.materials = [...mats.values()];
  return out;
}

// What a run's peak means. Flow runs since F4 say it; older Flow runs used a generic plastic block (package).
// Thermal runs: from the chip models in the run's own copy of forSolidWorks.
export function tjBasisOf(run, ref) {
  const own = run.parts?.[ref]?.tjBasis;
  if (own) return own;
  if (run.studyType === "Flow") return "package";
  const c = (run.inputs?.forSolidWorks?.chips || []).find(x => x.refDes === ref);
  if (c?.model === "Two-resistor" && c.rjcCW > 0 && c.rjbCW > 0) return "two-resistor";
  if (c && c.kz > 0) return "calibrated block";
  return "package";
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
