// Tab contents. W3 gives each tab its working core: Project, Parts and power, Link bodies and Changes can be
// edited; the other tabs show what the calculation module and the macros' files hold. Hover texts from the term
// catalogue come with W4, the full tabs 1 to 5 with W5, the SolidWorks cards with W7.

import { html, NumField, TextField, Select, Check, Status, Card, KV, fmt, pct } from "./ui.js";
import { CLASSES, METHODS, CLASS_METHOD } from "../calc/constants.js";
import { addPart, duplicatePart, deletePart, renameRefDes, nextLoadId } from "./model.js";
import { nextStep, unticked } from "./tabs.js";
import { isNum } from "../calc/util.js";

// ---------- 0 Project ----------
export function ProjectTab(c) {
  const { project: p, cad, results, states, folder } = c;
  return html`
    <${Card} title="Project folder">${c.folderCard}<//>
    ${p && html`
      <${Card} title="Board">
        <div class="form">
          <label>Board name <${TextField} value=${p.board.name} label="Board name" required onCommit=${v => c.edit("Board name", ["board", "name"], v, "board.name")} /></label>
          <label>STEP file <${TextField} value=${p.board.stepFile} label="STEP file" width="22em" onCommit=${v => c.edit("STEP file", ["board", "stepFile"], v, "board.stepFile")} /></label>
          <label>Ambient <${NumField} value=${p.board.ambientC} unit="C" label="Ambient" rule=${{ required: true, min: -273.15 }} onCommit=${v => v !== undefined && c.edit("Ambient", ["board", "ambientC"], v, "board.ambientC")} /></label>
          <label>Measured board input power <${NumField} value=${p.board.measuredInputPowerW} unit="W" label="Measured board input power" rule=${{ min: 0 }} onCommit=${v => c.edit("Measured input power", ["board", "measuredInputPowerW"], v, "board.measured")} /> <span class="hint">optional: checks the sum of all part Pd</span></label>
          <label class="wide">Notes <${TextField} value=${p.board.notes} label="Notes" multiline width="100%" onCommit=${v => c.edit("Notes", ["board", "notes"], v, "board.notes")} /></label>
        </div>
      <//>
      <${Card} title="Where you are">
        ${(() => { const s = nextStep(states); return s ? html`<p class="next">Next step: <button class="link" onClick=${() => c.go(s.n)}>tab ${s.n}, ${s.title}</button>${s.locked ? html` <span class="hint">(locked: ${s.reason})</span>` : s.note ? html` <span class="hint">(${s.note})</span>` : ""}</p>` : html`<p>All steps done.</p>`; })()}
        <ul class="steps">${states.filter(s => s.n > 0 && s.n < 10).map(s => html`<li class=${s.status}><button class="link" onClick=${() => c.go(s.n)}>${s.n} ${s.title}</button> <span class="state">${stateWord(s)}</span>${s.note ? html` <span class="hint">${s.note}</span>` : ""}</li>`)}</ul>
      <//>
      <${Card} title="Files">
        <${KV} rows=${[
          ["Folder", folder || "(files opened by hand)"],
          ["project.json", p.savedAt ? `saved ${p.savedAt.replace("T", " ")} by ${p.savedBy}` : "not saved yet"],
          ["cad.json", cad?.import ? `import ${cad.import.date?.replace("T", " ") ?? ""}, ${cad.import.bodies?.length ?? 0} bodies${cad.build ? `; build ${cad.build.date?.replace("T", " ") ?? ""}` : ""}` : "no import yet"],
          ["results.json", results?.runs?.length ? `${results.runs.length} runs, last ${results.runs[results.runs.length - 1].name}` : "no runs yet"],
        ]} />
        ${c.backupCard}
      <//>`}`;
}
export const stateWord = s => ({ locked: "locked", ready: "to do", done: "done", attention: "needs attention" }[s.status]);

// ---------- 1 Parts and power ----------
export function PartsTab(c) {
  const { project: p, calc, locked } = c;
  const reg = calc.register;
  const count = {};
  for (const x of p.parts) count[x.refDes] = (count[x.refDes] || 0) + 1;
  const setPart = (i, key, label, v) => c.edit(label, ["parts", i, key], v, `parts.${p.parts[i].id}.${key}`);
  const setLoad = (i, j, key, label, v) => c.edit(label, ["parts", i, "loads", j, key], v, `loads.${p.parts[i].loads[j].id}.${key}`);
  return html`
    <div class="toolbar">
      <button disabled=${locked} onClick=${() => c.editFn("Add part", addPart)}>Add part</button>
      <span class="hint">${p.parts.length} parts, board Pd ${fmt(reg.totalPdW, 2)} W</span>
    </div>
    ${p.parts.map((x, i) => {
      const r = reg.parts[x.refDes] || {};
      const dup = count[x.refDes] > 1;
      const method = CLASS_METHOD[x.class];
      return html`<section class="card part" key=${x.id}>
        <div class="part-head">
          <span class="field ${dup ? "bad" : ""}">
            <input class="refdes" type="text" value=${x.refDes} aria-label="RefDes" disabled=${locked}
              title=${dup ? "RefDes used by two parts" : "Reference designator. Links the part to its body."}
              onChange=${e => { const v = e.currentTarget.value.trim(); if (v && v !== x.refDes) c.editFn(`RefDes ${x.refDes} to ${v}`, q => renameRefDes(q, x.refDes, v)); else e.currentTarget.value = x.refDes; }} />
            ${dup && html`<span class="err">used twice</span>`}
          </span>
          <${TextField} value=${x.partNumber} label="Part number" disabled=${locked} onCommit=${v => setPart(i, "partNumber", "Part number", v)} />
          <${Select} value=${x.class} options=${CLASSES} label="Class" blankText="(pick class)" disabled=${locked} onCommit=${v => setPart(i, "class", "Class", v)} />
          <span class="result">Pd <b>${fmt(r.pdW, 3)}</b> W · Tj est <b>${fmt(r.tjEstC, 1)}</b> C · margin <b>${fmt(r.marginC, 1)}</b> C <${Status} word=${r.flag} /></span>
          <span class="actions">
            <button disabled=${locked} onClick=${() => c.editFn(`Duplicate ${x.refDes}`, q => duplicatePart(q, x.id))}>Duplicate</button>
            <button disabled=${locked} class="danger" onClick=${() => c.editFn(`Delete ${x.refDes}`, q => deletePart(q, x.id))}>Delete</button>
          </span>
        </div>
        <div class="form row">
          <label>Package <${TextField} value=${x.package} label="Package" width="8em" disabled=${locked} onCommit=${v => setPart(i, "package", "Package", v)} /></label>
          <label>Tj max <${NumField} value=${x.tjMaxC} unit="C" label="Tj max" disabled=${locked} rule=${{ min: -273.15 }} onCommit=${v => setPart(i, "tjMaxC", "Tj max", v)} /></label>
          <label>Rja <${NumField} value=${x.rjaCW} unit="CW" label="Rja" disabled=${locked} rule=${{ minEx: 0 }} onCommit=${v => setPart(i, "rjaCW", "Rja", v)} /></label>
          <label>Rjc <${NumField} value=${x.rjcCW} unit="CW" label="Rjc" disabled=${locked} rule=${{ minEx: 0 }} onCommit=${v => setPart(i, "rjcCW", "Rjc", v)} /></label>
          <label>Rjb <${NumField} value=${x.rjbCW} unit="CW" label="Rjb" disabled=${locked} rule=${{ minEx: 0 }} onCommit=${v => setPart(i, "rjbCW", "Rjb", v)} /></label>
          <label>Ta override <${NumField} value=${x.taOverrideC} unit="C" label="Ta override" disabled=${locked} rule=${{ min: -273.15 }} onCommit=${v => setPart(i, "taOverrideC", "Ta override", v)} /></label>
        </div>
        ${r.flag && r.flag !== "OK" && html`<p class="why">${r.why}</p>`}
        <div class="scroll"><table class="grid loads">
          <thead><tr><th>Load</th><th>Rail</th><th>Method</th><th>Vin V</th><th>Iin A</th><th>Vout V</th><th>Iout A</th><th>Efficiency</th><th>R ohm</th><th>Pd measured W</th><th>Pd W</th><th></th></tr></thead>
          <tbody>${x.loads.map((l, j) => {
            const lr = reg.loads[l.id] || {};
            return html`<tr key=${l.id}>
              <td><${TextField} value=${l.label} label="Load label" width="10em" disabled=${locked} onCommit=${v => setLoad(i, j, "label", "Load label", v)} /></td>
              <td><${TextField} value=${l.rail} label="Rail" width="6em" disabled=${locked} onCommit=${v => setLoad(i, j, "rail", "Rail", v)} /></td>
              <td><${Select} value=${l.method} options=${METHODS} label="Method" blankText=${method ? `(class: ${method})` : "(class default)"} disabled=${locked} onCommit=${v => setLoad(i, j, "method", "Method", v)} /></td>
              <td><${NumField} value=${l.vinV} unit="V" label="Vin" width="4.5em" disabled=${locked} onCommit=${v => setLoad(i, j, "vinV", "Vin", v)} /></td>
              <td><${NumField} value=${l.iinA} unit="A" label="Iin" width="4.5em" disabled=${locked} rule=${{ min: 0 }} onCommit=${v => setLoad(i, j, "iinA", "Iin", v)} /></td>
              <td><${NumField} value=${l.voutV} unit="V" label="Vout" width="4.5em" disabled=${locked} onCommit=${v => setLoad(i, j, "voutV", "Vout", v)} /></td>
              <td><${NumField} value=${l.ioutA} unit="A" label="Iout" width="4.5em" disabled=${locked} rule=${{ min: 0 }} onCommit=${v => setLoad(i, j, "ioutA", "Iout", v)} /></td>
              <td><${NumField} value=${l.efficiency} unit="" label="Efficiency (0 to 1)" width="4em" disabled=${locked} rule=${{ minEx: 0, max: 1 }} onCommit=${v => setLoad(i, j, "efficiency", "Efficiency", v)} /></td>
              <td><${NumField} value=${l.rOhm} unit="ohm" label="R" width="4.5em" disabled=${locked} rule=${{ min: 0 }} onCommit=${v => setLoad(i, j, "rOhm", "R", v)} /></td>
              <td><${NumField} value=${l.pdMeasuredW} unit="W" label="Pd measured" width="4.5em" disabled=${locked} rule=${{ min: 0 }} onCommit=${v => setLoad(i, j, "pdMeasuredW", "Pd measured", v)} /></td>
              <td class="num">${isNum(lr.pdW) ? fmt(lr.pdW, 3) : html`<span class="warn">${lr.problem}</span>`}</td>
              <td>${x.loads.length > 1 && html`<button class="small" disabled=${locked} title="Delete this load" onClick=${() => c.editFn("Delete load", q => ({ ...q, parts: q.parts.map(y => (y.id === x.id ? { ...y, loads: y.loads.filter(z => z.id !== l.id) } : y)) }))}>x</button>`}</td>
            </tr>`;
          })}</tbody>
        </table></div>
        <button class="small" disabled=${locked} onClick=${() => c.editFn(`Add load to ${x.refDes}`, q => ({ ...q, parts: q.parts.map(y => (y.id === x.id ? { ...y, loads: [...y.loads, { id: nextLoadId(y) }] } : y)) }))}>Add load</button>
      </section>`;
    })}
    ${!p.parts.length && html`<p class="hint">No parts yet. Press Add part, then type the RefDes from the schematic.</p>`}`;
}

// ---------- 2 Import geometry ----------
export function ImportTab(c) {
  const { cad, calc } = c;
  const imp = cad?.import;
  return html`
    <${Card} title="In SolidWorks: 2 Import STEP">
      <p>Import STEP reads the STEP file named on tab 0, finds the PCB and one body per part, and writes them to <code>cad.json</code>.
      The page reloads that file on its own when the macro has finished.</p>
    <//>
    ${imp ? html`
      <${Card} title="What the import found">
        <${KV} rows=${[["Model", imp.model], ["Date", imp.date?.replace("T", " ")], ["PCB", imp.pcb ? `${fmt(calc.pcb.l, 1)} x ${fmt(calc.pcb.w, 1)} x ${fmt(calc.pcb.h, 2)} mm` : "not found"],
          ["Bodies", imp.bodies?.length], ["Frame", imp.frame]]} />
        <${BoardDrawing} calc=${calc} />
      <//>` : html`<p class="hint">No import yet.</p>`}`;
}

export function BoardDrawing({ calc }) {
  const p = calc.pcb;
  if (!(p.l > 0 && p.w > 0)) return null;
  const pad = 2;
  return html`<svg class="board" viewBox=${`${-pad} ${-pad} ${p.l + 2 * pad} ${p.w + 2 * pad}`} role="img" aria-label="Top view of the board and its bodies">
    <g transform=${`translate(0 ${p.w}) scale(1 -1)`}>
      <rect class="pcb" x="0" y="0" width=${p.l} height=${p.w} />
      ${calc.bodies.filter(b => isNum(b.box.xMin) && b.box.l > 0).map(b => html`<rect class=${`body ${b.include ? "" : "skip"} ${b.side === "Bottom" ? "bottom" : ""} ${b.refDes ? "linked" : ""}`}
        x=${b.box.xMin} y=${b.box.yMin} width=${b.box.l} height=${b.box.w}><title>${b.refDes ? b.refDes + ": " : ""}${b.name}</title></rect>`)}
    </g>
  </svg>
  <p class="hint legend"><span class="sw linked"></span> linked to a part <span class="sw"></span> not linked <span class="sw bottom"></span> bottom side <span class="sw skip"></span> not included</p>`;
}

// ---------- 3 Link bodies ----------
export function LinkTab(c) {
  const { project: p, calc, locked } = c;
  const refs = p.parts.map(x => x.refDes);
  const noBody = p.parts.filter(x => isNum(calc.register.parts[x.refDes]?.pdW) && !calc.bodies.some(b => b.refDes === x.refDes)).map(x => x.refDes);
  return html`
    ${noBody.length > 0 && html`<p class="warn">Parts with Pd and no body: ${noBody.join(", ")}. Link each to its body below.</p>`}
    <div class="scroll"><table class="grid">
      <thead><tr><th>Include</th><th>Body</th><th>Side</th><th>RefDes</th><th class="num">L mm</th><th class="num">W mm</th><th class="num">H mm</th></tr></thead>
      <tbody>${calc.bodies.map(b => html`<tr key=${b.key} class=${b.include ? "" : "dim"}>
        <td><${Check} value=${b.include} label=${"Include " + b.name} disabled=${locked} onCommit=${v => c.edit(`Include ${shortName(b.name)}`, ["bodies", b.key, "include"], v)} /></td>
        <td title=${b.name}>${shortName(b.name)}${b.manual ? html` <span class="hint">(typed in)</span>` : ""}</td>
        <td>${b.side}</td>
        <td><${Select} value=${b.refDes} options=${refs} label="RefDes" blankText="(none)" disabled=${locked} onCommit=${v => c.edit(`Link ${shortName(b.name)}`, ["bodies", b.key, "refDes"], v)} /></td>
        <td class="num">${fmt(b.box.l, 2)}</td><td class="num">${fmt(b.box.w, 2)}</td><td class="num">${fmt(b.box.h, 2)}</td>
      </tr>`)}</tbody>
    </table></div>`;
}
export const shortName = n => (n || "").split("/").pop().replace(/-\d+$/, "");

// ---------- 4 Triage ----------
export function TriageTab(c) {
  const { calc } = c;
  const reg = calc.register, t = calc.triage, rails = calc.rails;
  const problems = c.blocking;
  return html`
    ${problems.length ? html`<${Card} title="Fix first" cls="attention"><ul>${problems.map(x => html`<li>${x}</li>`)}</ul><//>` : ""}
    <${Card} title="Overview">
      <${KV} rows=${[["Parts", t.parts], ["Board Pd", fmt(t.totalPdW, 2) + " W"], ["CRITICAL", t.critical], ["Watch", t.watch], ["Minor", t.minor], ["No Pd", t.noPd], ["No Rja", t.noRja], ["No body", t.noBody]]} />
    <//>
    <div class="cols">
      <${Card} title="By Pd"><ol>${t.byPd.map(r => html`<li>${r} <span class="hint">${fmt(reg.parts[r].pdW, 3)} W</span> <${Status} word=${reg.parts[r].flag} /></li>`)}</ol><//>
      <${Card} title="By power density"><ol>${t.byDensity.map(r => html`<li>${r} <span class="hint">${fmt(reg.parts[r].densityWcm2, 2)} W/cm2</span></li>`)}</ol><//>
    </div>
    <${Card} title="Rail budget">
      <div class="scroll"><table class="grid"><thead><tr><th>Rail</th><th>Parts</th><th class="num">Pd W</th><th class="num">Share</th></tr></thead>
      <tbody>${rails.rails.map(r => html`<tr><td>${r.rail}</td><td>${r.parts.join(", ")}</td><td class="num">${fmt(r.pdW, 3)}</td><td class="num">${pct(r.share)}</td></tr>`)}</tbody></table></div>
      <${KV} rows=${[["Sum of rails", fmt(rails.sumRailsW, 3) + " W"], ["Loads with no rail", fmt(rails.noRailW, 3) + " W"],
        isNum(rails.measuredW) && ["Measured input", fmt(rails.measuredW, 2) + " W"], isNum(rails.unaccountedW) && ["Not accounted for", `${fmt(rails.unaccountedW, 2)} W (${pct(rails.unaccountedFrac)})`]]} />
    <//>`;
}

// ---------- 5 Hand check: bare board ----------
export function BareTab(c) {
  const { project: p, calc } = c;
  const bb = calc.bareBoard, b = bb.board;
  return html`
    <${Card} title="Board">
      <${KV} rows=${[["Board area", fmt(b.areaMm2, 0) + " mm2"], ["Board to air", fmt(b.rBoardAirCW, 2) + " C/W"], ["Heat into the board", fmt(b.qBoardW, 2) + " W"],
        ["Board temperature", fmt(b.boardC, 1) + " C"], ["FAIL", b.fail], ["Tight", b.tight]]} />
    <//>
    <div class="scroll"><table class="grid">
      <thead><tr><th>RefDes</th><th class="num">Pd W</th><th class="num">R top C/W</th><th class="num">R board C/W</th><th class="num">To board W</th><th class="num">Tj C</th><th class="num">Margin C</th><th>Status</th></tr></thead>
      <tbody>${p.parts.map(x => { const r = bb.parts[x.refDes]; return r && html`<tr><td>${x.refDes}</td><td class="num">${fmt(r.pdW, 3)}</td><td class="num">${fmt(r.rTopCW, 1)}</td>
        <td class="num">${fmt(r.rBoardCW, 2)}</td><td class="num">${fmt(r.qBoardW, 3)}</td><td class="num">${fmt(r.tjC, 1)}</td><td class="num">${fmt(r.marginC, 1)}</td><td><${Status} word=${r.status} /></td></tr>`; })}</tbody>
    </table></div>
    ${bb.compare?.compared ? html`<p class="hint">Compared with run ${bb.compare.run}: mean gap ${fmt(bb.compare.meanGapC, 1)} C, largest ${fmt(bb.compare.maxGapC, 1)} C. ${bb.compare.check}</p>` : ""}`;
}

// ---------- 6 Build model ----------
export function BuildTab(c) {
  const b = c.cad?.build;
  return html`
    <${Card} title="In SolidWorks: 3 Build Board"><p>Builds one box per included body and the PCB plate, in a new assembly. ${c.calc.bodies.filter(x => x.include).length} bodies are included now.</p><//>
    ${b ? html`<${Card} title="Last build">
      <${KV} rows=${[["Assembly", b.assembly], ["Date", b.date?.replace("T", " ")], ["Mode", b.mode], ...Object.entries(b.counts || {}).map(([k, v]) => [k, v])]} />
      ${(b.bodies || []).filter(x => x.status !== "built").map(x => html`<p class="warn">${shortName(x.name)}: ${x.status}${x.reason ? ", " + x.reason : ""}</p>`)}
      ${b.overlaps?.length ? html`<p class="hint">${b.overlaps.length} overlaps reported.</p>` : ""}
    <//>` : html`<p class="hint">No build yet.</p>`}`;
}

// ---------- 7 Spreader ----------
export function SpreaderTab(c) {
  const { project: p, calc } = c;
  const sides = ["top", "bottom"].filter(s => p.spreader?.[s]);
  if (!sides.length) return html`<p class="hint">No spreader plate in this project. Plates are optional; the editor comes with W7.</p>`;
  return sides.map(s => {
    const pl = p.spreader[s], r = calc.spreader[s];
    return html`<${Card} title=${`${s === "top" ? "Top" : "Bottom"} plate${pl.enabled ? "" : " (not enabled)"}`}>
      <${KV} rows=${[["Material", `${pl.material}, k ${pl.k} W/(m K)`], ["Thickness", pl.thicknessMm + " mm"], ["Face height", fmt(r.faceZMm, 3) + " mm"], ["Chips covered", r.covered], ["Clearance cuts", r.cuts.length]]} />
      <div class="scroll"><table class="grid"><thead><tr><th>RefDes</th><th>Covered</th><th class="num">Gap mm</th><th class="num">TIM mm</th><th class="num">Pedestal mm</th><th>Check</th></tr></thead>
      <tbody>${r.chips.map(x => html`<tr><td>${x.refDes}</td><td>${x.covered ? "yes" : "no"}</td><td class="num">${fmt(x.gapMm, 3)}</td><td class="num">${fmt(x.timTMm, 3)}</td><td class="num">${fmt(x.pedestalMm, 3)}</td><td class=${x.check === "OK" ? "" : "warn"}>${x.check}</td></tr>`)}</tbody></table></div>
    <//>`;
  });
}

// ---------- 8 Simulation setup ----------
export function SimTab(c) {
  const chips = c.calc.simChips;
  return html`
    <${Card} title="In SolidWorks: 5 Run Thermal Study"><p>Sets up and runs the steady state study from these values. The editor for h, contacts and mesh comes with W7.</p><//>
    <div class="scroll"><table class="grid"><thead><tr><th>RefDes</th><th>Model</th><th class="num">Pd W</th><th class="num">Rjc C/W</th><th class="num">H mm</th><th class="num">k z W/(m K)</th><th class="num">k xy W/(m K)</th><th>Check</th></tr></thead>
    <tbody>${chips.map(x => html`<tr><td>${x.refDes}</td><td>${x.model}</td><td class="num">${fmt(x.pdW, 3)}</td><td class="num">${fmt(x.rjcCW, 2)}</td><td class="num">${fmt(x.hMm, 2)}</td><td class="num">${fmt(x.kz, 2)}</td><td class="num">${fmt(x.kxy, 2)}</td><td class=${x.check === "OK" ? "" : "warn"}>${x.check}</td></tr>`)}</tbody></table></div>`;
}

// ---------- 9 Results ----------
export function ResultsTab(c) {
  const runs = c.results?.runs || [];
  return [...runs].reverse().map(r => html`<${Card} title=${r.name}>
    <${KV} rows=${[["Date", r.date?.replace("T", " ")], ["Study", r.studyType], ["Status", r.status], isNum(r.energyBalancePct) && ["Energy balance error", r.energyBalancePct + " %"],
      r.pcb && ["PCB", `${fmt(r.pcb.minC, 1)} to ${fmt(r.pcb.maxC, 1)} C`], r.spreader && ["Spreader", `${fmt(r.spreader.minC, 1)} to ${fmt(r.spreader.maxC, 1)} C`]]} />
    ${r.parts && html`<div class="scroll"><table class="grid"><thead><tr><th>RefDes</th><th class="num">Peak C</th><th class="num">Average C</th><th class="num">Tj max C</th></tr></thead>
      <tbody>${Object.entries(r.parts).map(([k, v]) => { const tj = c.project.parts.find(x => x.refDes === k)?.tjMaxC; return html`<tr><td>${k}</td><td class=${`num ${isNum(tj) && v.peakC > tj ? "warn" : ""}`}>${fmt(v.peakC, 1)}</td><td class="num">${fmt(v.avgC, 1)}</td><td class="num">${fmt(tj, 0)}</td></tr>`; })}</tbody></table></div>`}
  <//>`);
}

// ---------- 10 Changes ----------
export function ChangesTab(c) {
  const { project: p, cad, now } = c;
  const list = cad?.changes || [];
  if (!p) return html`<p class="hint">Open a project to see its change records.</p>`;
  if (!list.length) return html`<p class="hint">No change records. Import STEP, Build Board and Sync CAD write them when something moves or changes size.</p>`;
  const open = unticked(p, cad);
  const tickAll = () => c.editFn(`Tick ${open} changes`, q => { const checks = { ...(q.checks || {}) }; for (const x of list) if (!checks[x.id]) checks[x.id] = { at: now() }; return { ...q, checks }; });
  return html`
    <div class="toolbar"><span>${open ? `${open} of ${list.length} not ticked` : `All ${list.length} ticked`}</span>${open > 0 && html`<button onClick=${tickAll}>Tick all</button>`}</div>
    <div class="scroll"><table class="grid"><thead><tr><th>Checked</th><th>Date</th><th>Macro</th><th>Body</th><th>Change</th><th>Field</th><th class="num">Old</th><th class="num">New</th></tr></thead>
    <tbody>${list.map(x => html`<tr class=${p.checks?.[x.id] ? "dim" : ""}>
      <td><${Check} value=${!!p.checks?.[x.id]} label=${"Checked " + x.id} onCommit=${v => c.edit(v ? "Tick change" : "Untick change", ["checks", x.id], v ? { at: now() } : undefined)} /></td>
      <td>${x.date?.replace("T", " ")}</td><td>${x.macro}</td><td title=${x.name}>${shortName(x.name)}</td><td>${x.change}</td><td>${x.field ?? ""}</td><td class="num">${x.old ?? ""}</td><td class="num">${x.new ?? ""}</td></tr>`)}</tbody></table></div>`;
}

// ---------- 11 Help ----------
export function HelpTab() {
  return html`<${Card} title="Start here">
    <p>This page takes a PCB from STEP file to a checked thermal result. You do the engineering (power numbers, datasheet values, judging the results); the SolidWorks macros build the model and run the study.</p>
    <ol>
      <li>Tab 0: choose the board's project folder. One folder per board. The page saves <code>project.json</code> there.</li>
      <li>Tab 1: add each part that makes heat, with its datasheet values and its power.</li>
      <li>Tabs 2 to 9: follow the tabs in order. A grey tab says what to do first and links to it.</li>
    </ol>
    <h4>Saving</h4>
    <p>Save (Ctrl+S) writes <code>project.json</code>. While there are unsaved edits the page also saves on its own every 30 s. Before the first save of each day it copies the old file to <code>backup\\project_yyyymmdd.json</code>.</p>
    <h4>Undo</h4>
    <p>Ctrl+Z undoes the last edit, Ctrl+Y (or Ctrl+Shift+Z) redoes it. Undo works for the whole session, including deleted parts.</p>
    <h4>Files the macros write</h4>
    <p>The page never writes <code>cad.json</code> or <code>results.json</code>. It checks them every 2 s and reloads them when a macro has finished.</p>
    <p class="hint">The full help, with the heat basics and every term explained, comes with W8.</p>
  <//>`;
}
