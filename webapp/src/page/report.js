// Report (design section 9): a print layout of the study, saved as PDF with the browser's Print > Save as PDF. Always
// light. Cover, assumptions, inputs, hand check, hand check with a plate, simulation runs with hand vs simulation,
// pass or fail per part, plot pictures, open issues.
//
//   <Report ctx=${{ project, cad, results, calc }} runs=${["PT_..."]} images=${{ run: url }} />
//   openIssues(ctx) -> texts, also used by tools/test_w9.mjs

import { html } from "./ui.js";
import { valueOf } from "../calc/formulas.js";
import { TERMS } from "../terms/catalogue.js";
import { unticked, blockingProblems } from "./tabs.js";
import { isNum } from "../calc/util.js";
import { shortName } from "./bits.js";
import { PAGE_VERSION } from "./model.js";

const fx = (v, d) => (typeof v === "boolean" ? (v ? "yes" : "no") : isNum(v) ? (d === undefined ? String(+v.toPrecision(4)) : v.toFixed(d)) : v ?? "");
const lbl = t => TERMS[t]?.label || t;
const unit = t => (TERMS[t]?.unit ? ` (${TERMS[t].unit})` : "");

function Table({ ctx, cols, rows }) {
  return html`<table class="rep"><thead><tr>${cols.map(([t, , d, h]) => html`<th class=${d !== undefined ? "num" : ""}>${h || lbl(t)}${unit(t)}</th>`)}</tr></thead>
    <tbody>${rows.map(r => html`<tr>${cols.map(([t, kf, d]) => { const v = valueOf(ctx, t, kf(r)); return html`<td class=${d !== undefined ? "num" : ""}>${fx(v, d)}</td>`; })}</tr>`)}</tbody></table>`;
}
function Facts({ ctx, rows }) {
  return html`<table class="rep facts"><tbody>${rows.map(([t, k = "", d]) => { const v = valueOf(ctx, t, k); return v === undefined || v === "" ? null : html`<tr><th>${lbl(t)}</th><td>${fx(v, d)}${TERMS[t]?.unit ? " " + TERMS[t].unit : ""}</td></tr>`; })}</tbody></table>`;
}

export function openIssues(ctx) {
  const { project: p, cad, results, calc } = ctx;
  const out = [];
  const open = unticked(p, cad);
  if (open) out.push(`${open} change record${open === 1 ? "" : "s"} not ticked (tab 10).`);
  for (const b of blockingProblems(p, calc)) out.push(`Hand check blocked: ${b}.`);
  const noBody = p.parts.filter(x => calc.register.parts[x.refDes]?.pdW > 0 && !calc.bodies.some(b => b.include && b.refDes === x.refDes)).map(x => x.refDes);
  if (noBody.length) out.push(`Parts with Pd but no body in the model: ${noBody.join(", ")}.`);
  const failed = (cad?.build?.bodies || []).filter(x => x.status === "failed");
  if (failed.length) out.push(`Build Board failed for ${failed.map(x => shortName(x.name)).join(", ")}.`);
  if ((cad?.build?.overlaps || []).length) out.push(`${cad.build.overlaps.length} overlapping pairs of boxes in the model.`);
  if (calc.build?.outOfDate) out.push("The model is out of date: the PCB or bodies changed after the last build.");
  if (!cad?.build) out.push("The thermal model has not been built yet.");
  const bbFail = p.parts.filter(x => calc.bareBoard.parts[x.refDes]?.status === "FAIL").map(x => x.refDes);
  if (bbFail.length) out.push(`Bare board hand check FAIL: ${bbFail.join(", ")}.`);
  const runs = results?.runs || [];
  if (!runs.length) out.push("No simulation run yet.");
  const last = runs[runs.length - 1];
  const rc = last && calc.runs[last.name];
  if (rc?.pass === "FAIL") out.push(`Run ${last.name}: ${rc.fails} part${rc.fails === 1 ? "" : "s"} above Tj max.`);
  if (rc?.balanceOk === false) out.push(`Run ${last.name}: energy balance error ${fx(last.energyBalancePct, 2)} %, above 2 %.`);
  if (rc?.changes?.length) out.push(`Run ${last.name} is older than the inputs: ${rc.changes.slice(0, 5).join("; ")}${rc.changes.length > 5 ? "; and more" : ""}.`);
  for (const r of runs.filter(x => x.studyType === "Thermal" && x.status !== "read")) out.push(`Run ${r.name} is not read yet (Read Results).`);
  const noTj = p.parts.filter(x => calc.register.parts[x.refDes]?.pdW > 0 && !isNum(x.tjMaxC)).map(x => x.refDes);
  if (noTj.length) out.push(`Parts with Pd but no Tj max (not checked): ${noTj.join(", ")}.`);
  return out;
}

// Pass or fail per part: the newest selected run where the part has a result, else the hand checks.
function verdicts(ctx, runNames) {
  const { project: p, calc, results } = ctx;
  const runs = (results?.runs || []).filter(r => runNames.includes(r.name)).reverse();
  return p.parts.filter(x => calc.register.parts[x.refDes]?.pdW > 0).map(x => {
    const ref = x.refDes;
    const run = runs.find(r => r.parts?.[ref]);
    if (run) { const r = calc.runs[run.name].parts[ref]; return { ref, from: `run ${run.name}`, t: run.parts[ref].peakC, margin: r.marginC, status: r.status }; }
    const sol = ["top", "bottom"].map(s => calc.solution[s]?.parts?.[ref]).find(r => isNum(r?.tjC));
    if (sol) return { ref, from: "hand check with plate", t: sol.tjC, margin: sol.marginC, status: sol.status };
    const bb = calc.bareBoard.parts[ref];
    return { ref, from: "bare board hand check", t: bb?.tjC, margin: bb?.marginC, status: bb?.status || "not worked out" };
  });
}

export function Report({ ctx, runs: runNames = [], images = {}, now }) {
  const { project: p, cad, results, calc } = ctx;
  const runs = (results?.runs || []).filter(r => runNames.includes(r.name));
  const refs = p.parts.map(x => x.refDes);
  const heated = refs.filter(r => calc.register.parts[r]?.pdW > 0);
  const issues = openIssues(ctx);
  const ver = verdicts(ctx, runNames);
  const fails = ver.filter(v => v.status === "FAIL");
  const sol = ["top", "bottom"].filter(s => calc.solution[s]);
  const s = p.simSetup?.surfaces || {};
  return html`<article class="report">
    <section class="cover">
      <h1>${p.board.name}</h1>
      <p class="sub">Thermal study report</p>
      <table class="rep facts"><tbody>
        <tr><th>Date</th><td>${(now || "").replace("T", " ").slice(0, 16)}</td></tr>
        <tr><th>Project file</th><td>saved ${(p.savedAt || "not saved").replace("T", " ")} by ${p.savedBy || ""}</td></tr>
        <tr><th>Runs in this report</th><td>${runs.length ? runs.map(r => r.name).join(", ") : "none"}</td></tr>
        <tr><th>Result</th><td><b>${fails.length ? `FAIL: ${fails.map(v => v.ref).join(", ")} above Tj max` : ver.some(v => v.status === "OK" || v.status === "Tight") ? "PASS: every checked part is below its Tj max" : "not checked"}</b></td></tr>
        <tr><th>Open issues</th><td>${issues.length || "none"}</td></tr>
      </tbody></table>
    </section>

    <h2>1 Assumptions</h2>
    <${Facts} ctx=${ctx} rows=${[["board.ambient", "", 1], ["board.measuredInputPower", "", 2], ["board.totalPd", "", 2], ["pcb.l", "", 1], ["pcb.w", "", 1], ["pcb.h", "", 2],
      ["stack.layers"], ["stack.cuT", "", 3], ["stack.coverage", "", 2], ["pcb.kxy", "", 1], ["pcb.kz", "", 3], ["hand.surfaceRise", "", 0]]} />
    <table class="rep"><thead><tr><th>Surface</th><th class="num">h (W/(m² K))</th><th class="num">Emissivity</th></tr></thead>
      <tbody>${[["pcbTop", "PCB top"], ["pcbBottom", "PCB bottom"], ["parts", "Parts"], ["spreaderTop", "Spreader top"], ["spreaderOther", "Spreader other faces"]].map(([g, n]) => s[g] && html`<tr><td>${n}</td><td class="num">${fx(s[g].h)}</td><td class="num">${fx(s[g].emissivity)}</td></tr>`)}</tbody></table>
    ${p.board.notes && html`<p class="notes">${p.board.notes}</p>`}

    <h2>2 Parts and power</h2>
    <${Table} ctx=${ctx} rows=${refs} cols=${[["part.refDes", r => r], ["part.partNumber", r => r], ["part.class", r => r], ["part.package", r => r], ["part.tjMax", r => r, 0], ["part.rja", r => r, 1], ["part.rjc", r => r, 2], ["part.rjb", r => r, 2],
      ["part.pd", r => r, 3], ["part.tjEst", r => r, 1], ["part.flag", r => r]]} />

    <h2>3 Hand check: bare board</h2>
    <${Facts} ctx=${ctx} rows=${[["bb.boardC", "", 1], ["bb.rise", "", 1], ["bb.boardShare", "", 2], ["bb.fail"], ["bb.tight"]]} />
    <${Table} ctx=${ctx} rows=${heated.filter(r => calc.bareBoard.parts[r])} cols=${[["part.refDes", r => r], ["bbp.pd", r => r, 3], ["bbp.rTop", r => r, 1], ["bbp.rSpread", r => r, 2], ["bbp.rBoard", r => r, 2], ["bbp.qBoard", r => r, 3], ["bbp.tj", r => r, 1], ["bbp.margin", r => r, 1], ["bbp.status", r => r]]} />

    ${sol.length > 0 && html`<h2>4 Hand check with a spreader plate</h2>${sol.map(side => html`<h3>${side === "top" ? "Top" : "Bottom"} plate</h3>
      <${Facts} ctx=${ctx} rows=${[["sol.area", side, 0], ["sol.hEff", side, 1], ["sol.rPlateAir", side, 3], ["sol.pdPlate", side, 3], ["sol.plateC", side, 1]]} />
      <${Table} ctx=${ctx} rows=${refs.filter(r => calc.solution[side].parts[r]?.status !== "not under plate").map(r => `${side}|${r}`)} cols=${[["part.refDes", k => k.split("|")[1]], ["solp.rStack", k => k, 3], ["solp.dtStack", k => k, 1], ["solp.tj", k => k, 1], ["solp.margin", k => k, 1], ["solp.status", k => k]]} />`)}`}

    <h2>5 Simulation</h2>
    ${!runs.length ? html`<p>No run in this report.</p>` : runs.map(r => html`<div class="run"><h3>Run ${r.name}</h3>
      <${Facts} ctx=${ctx} rows=${[["run.date", r.name], ["run.studyType", r.name], ["run.study", r.name], ["run.pass", r.name], ["run.worst", r.name], ["run.worstC", r.name, 1], ["run.energyBalance", r.name, 2], ["run.pcbMin", r.name, 1], ["run.pcbMax", r.name, 1], ["run.spreaderMin", r.name, 1], ["run.spreaderMax", r.name, 1], ["run.meshGlobal", r.name], ["run.meshFine", r.name], ["run.changes", r.name]]} />
      <${Table} ctx=${ctx} rows=${Object.keys(r.parts || {}).map(k => `${r.name}|${k}`)} cols=${[["part.refDes", k => k.split("|")[1], undefined, "RefDes"], ["part.tjMax", k => k.split("|")[1], 0], ["runp.peak", k => k, 1], ["runp.avg", k => k, 1], ["runp.margin", k => k, 1], ["runp.status", k => k], ["runp.hand", k => k, 1], ["runp.gap", k => k, 1]]} />
      ${images[r.name] && html`<figure><img src=${images[r.name]} alt=${"Temperature plot of " + r.name} /><figcaption>Temperature plot, run ${r.name}</figcaption></figure>`}
    </div>`)}

    <h2>6 Pass or fail per part</h2>
    <table class="rep"><thead><tr><th>RefDes</th><th class="num">Tj max (C)</th><th class="num">Temperature (C)</th><th class="num">Margin (C)</th><th>Status</th><th>From</th></tr></thead>
      <tbody>${ver.map(v => html`<tr class=${v.status === "FAIL" ? "fail" : ""}><td>${v.ref}</td><td class="num">${fx(p.parts.find(x => x.refDes === v.ref)?.tjMaxC, 0)}</td><td class="num">${fx(v.t, 1)}</td><td class="num">${fx(v.margin, 1)}</td><td>${v.status}</td><td>${v.from}</td></tr>`)}</tbody></table>

    <h2>7 Open issues</h2>
    ${issues.length ? html`<ul>${issues.map(x => html`<li>${x}</li>`)}</ul>` : html`<p>None.</p>`}
    <p class="foot">Made by PCB Thermal page ${PAGE_VERSION}. Values in mm, C, W, W/(m K) and C/W. Every formula is listed in the workbook export (Method tab) and on the page (Show formulas).</p>
  </article>`;
}
