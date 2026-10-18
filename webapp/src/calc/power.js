// Parts and power (1 Component register), triage (4 Triage) and rail budget (5 Rail budget).
// The workbook had one row per load; here a part holds its loads and Tj est uses the part's total Pd.

import { CLASS_METHOD } from "./constants.js";
import { isNum, N, fixed, percent, numText, ranks, sum } from "./util.js";

// One load: method, Pin, Pout, Pd calc and Pd used. `problem` says why Pd is missing.
export function loadPower(part, load) {
  const method = load.method || (part.class ? CLASS_METHOD[part.class] || "not in map" : "pick class");
  const { vinV: vin, iinA: iin, voutV: vout, ioutA: iout, efficiency: eff, rOhm: r, pdMeasuredW: meas } = load;
  const pinW = isNum(vin) && isNum(iin) ? vin * iin : undefined;
  const poutW = isNum(vout) && isNum(iout) ? vout * iout : undefined;
  const count = (...xs) => xs.filter(isNum).length;
  let pdCalcW, problem;
  switch (method) {
    case "LDO / linear reg":
      if (isNum(pinW) && isNum(poutW)) pdCalcW = pinW - poutW;
      else if (count(vin, vout, iout) < 3) problem = "missing input";
      else pdCalcW = (vin - vout) * iout;
      break;
    case "Switching regulator":
      if (isNum(pinW) && isNum(poutW)) pdCalcW = pinW - poutW;
      else if (count(vout, iout, eff) < 3 || N(eff) === 0) problem = "missing input";
      else pdCalcW = poutW * (1 / eff - 1);
      break;
    case "Digital load (V x I)":
      if (isNum(pinW)) pdCalcW = pinW - N(poutW);
      else problem = "missing input";
      break;
    case "Resistive (I squared R)":
      if ((isNum(iin) || isNum(iout)) && isNum(r)) pdCalcW = (isNum(iin) ? iin : iout) ** 2 * r;
      else problem = "missing input";
      break;
    case "Manual entry":
      if (!isNum(meas)) problem = "enter Pd meas";
      break;
    default:
      problem = method;   // "pick class" or "not in map"
  }
  const pdW = isNum(meas) ? meas : pdCalcW;
  if (isNum(pdW)) problem = undefined;
  return { method, pinW, poutW, pdCalcW, pdW, problem };
}

// Register per part. `areaOf(refDes)` gives the body footprint L x W (mm2) or undefined.
export function register(project, areaOf) {
  const lim = project.settings.flagLimits;
  const ambient = project.board.ambientC;
  const loads = {};
  for (const p of project.parts) for (const l of p.loads) loads[l.id] = loadPower(p, l);
  const totalPdW = sum(Object.values(loads).map(l => l.pdW));

  const parts = {};
  for (const p of project.parts) {
    const ls = p.loads.map(l => loads[l.id]);
    const withPd = ls.filter(l => isNum(l.pdW));
    const r = { pdW: withPd.length ? sum(withPd.map(l => l.pdW)) : undefined };
    r.taC = isNum(p.taOverrideC) ? p.taOverrideC : ambient;
    r.areaMm2 = areaOf(p.refDes);
    if (isNum(r.pdW) && isNum(p.rjaCW)) r.riseC = r.pdW * p.rjaCW;
    if (isNum(r.riseC)) r.tjEstC = r.taC + r.riseC;
    if (isNum(r.tjEstC) && isNum(p.tjMaxC)) r.marginC = p.tjMaxC - r.tjEstC;
    if (isNum(r.pdW) && isNum(r.areaMm2) && r.areaMm2 > 0) r.densityWcm2 = r.pdW / (r.areaMm2 / 100);
    if (isNum(r.pdW) && totalPdW > 0) r.share = r.pdW / totalPdW;
    Object.assign(r, flagOf(r, p, ls, lim));
    parts[p.refDes] = r;
  }
  const refs = project.parts.map(p => p.refDes);
  const rp = ranks(refs.map(x => parts[x].pdW));
  const rd = ranks(refs.map(x => parts[x].densityWcm2));
  refs.forEach((x, i) => { parts[x].rankPd = rp[i]; parts[x].rankDensity = rd[i]; });
  return { loads, parts, totalPdW };
}

function flagOf(r, part, ls, lim) {
  if (!isNum(r.pdW)) {
    const l = ls.find(x => x.problem) || ls[0];
    return { flag: "No Pd", why: "Pd not computed: " + (l?.problem || "missing input") };
  }
  const hit = L => ({
    margin: isNum(r.marginC) && r.marginC < L.marginBelowC,
    density: isNum(r.densityWcm2) && r.densityWcm2 > L.densityAboveWcm2,
    share: isNum(r.share) && r.share > L.shareAbove,
  });
  const c = hit(lim.critical), w = hit(lim.watch);
  const flag = c.margin || c.density || c.share ? "CRITICAL" : w.margin || w.density || w.share ? "Watch" : "Minor";
  let why = "";
  if (flag !== "Minor") {
    const L = flag === "CRITICAL" ? lim.critical : lim.watch, h = flag === "CRITICAL" ? c : w;
    if (h.margin) why += `; margin ${fixed(r.marginC, 1)} °C < ${numText(L.marginBelowC)}`;
    if (h.density) why += `; density ${fixed(r.densityWcm2, 2)} > ${numText(L.densityAboveWcm2)} W/cm²`;
    if (h.share) why += `; share ${percent(r.share)} > ${percent(L.shareAbove)}`;
  }
  if (!isNum(part.rjaCW)) why += "; no Rja, margin not checked";
  else if (!isNum(part.tjMaxC)) why += "; no Tj max, margin not checked";
  if (!isNum(r.areaMm2)) why += "; no body on CAD geometry, density not checked";
  const missing = ls.filter(l => !isNum(l.pdW)).length;
  if (missing) why += `; ${missing} of ${ls.length} loads have no Pd`;
  return { flag, why: why ? why.slice(2) : "below Watch limits" };
}

// Triage summary (4 Triage, top block). The ranked lists are the parts sorted by rankPd / rankDensity.
export function triage(project, reg) {
  const ps = project.parts;
  const count = f => ps.filter(p => reg.parts[p.refDes].flag === f).length;
  const by = k => ps.map(p => p.refDes).sort((a, b) => reg.parts[a][k] - reg.parts[b][k]);
  return {
    critical: count("CRITICAL"), watch: count("Watch"), minor: count("Minor"), noPd: count("No Pd"),
    parts: ps.length, totalPdW: reg.totalPdW,
    noRja: ps.filter(p => !isNum(p.rjaCW)).length,
    noBody: ps.filter(p => !isNum(reg.parts[p.refDes].areaMm2)).length,
    byPd: by("rankPd"), byDensity: by("rankDensity"),
  };
}

// Rail budget (5 Rail budget). Per load, as the register rows were.
export function railBudget(project, reg) {
  const map = new Map();
  let loadsNoRail = 0;
  for (const p of project.parts) for (const l of p.loads) {
    const pd = reg.loads[l.id].pdW;
    if (!l.rail) { loadsNoRail++; continue; }
    if (!map.has(l.rail)) map.set(l.rail, { rail: l.rail, loads: 0, parts: new Set(), pdW: 0, noPd: 0 });
    const r = map.get(l.rail);
    r.loads++; r.parts.add(p.refDes); r.pdW += N(pd); if (!isNum(pd)) r.noPd++;
  }
  const total = reg.totalPdW;
  const rails = [...map.values()].map(r => ({ ...r, parts: r.parts.size, share: total !== 0 ? r.pdW / total : undefined }));
  const sumW = sum(rails.map(r => r.pdW));
  const out = { rails, totalPdW: total, sumRailsW: sumW, noRailW: total - sumW, loadsNoRail };
  const meas = project.board.measuredInputPowerW;
  if (N(meas) !== 0) {
    out.measuredW = meas;
    out.unaccountedW = meas - total;
    out.unaccountedFrac = out.unaccountedW / meas;
  }
  return out;
}
