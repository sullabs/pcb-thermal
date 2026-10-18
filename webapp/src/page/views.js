// Tab contents. Every value is a <V> and every label a <L> with a term id from the catalogue, so it has a hover
// text, opens the Explain panel, and shows its formula when Show formulas is on (tools/test_w4.mjs checks this).
// The full tabs 1 to 5 come with W5, the SolidWorks cards with W7.

import { html, NumField, TextField, Select, Check, Card } from "./ui.js";
import { V, L, marked } from "./explain.js";
import { CLASSES, METHODS, CLASS_METHOD } from "../calc/constants.js";
import { addPart, duplicatePart, deletePart, renameRefDes, nextLoadId } from "./model.js";
import { nextStep, unticked } from "./tabs.js";
import { isNum } from "../calc/util.js";

const Scroll = ({ children }) => html`<div class="scroll">${children}</div>`;
// A definition list of [label term, value term, key, decimals, value] rows; rows whose value is blank are left out.
function Facts({ ctx, rows }) {
  return html`<dl class="kv">${rows.filter(Boolean).filter(([, t, k, , v]) => { const x = v !== undefined ? v : ctx.value(t, k); return x !== undefined && x !== ""; })
    .map(([lt, t, k = "", d, v]) => html`<dt><${L} t=${lt || t} k=${k} /></dt><dd><${V} t=${t} k=${k} d=${d} v=${v} u /></dd>`)}</dl>`;
}
const date = s => (s ? s.replace("T", " ") : s);

// ---------- 0 Project ----------
export function ProjectTab(c) {
  const { project: p, cad, results, states, folder } = c;
  return html`
    <${Card} title="Project folder">${c.folderCard}<//>
    ${p && html`
      <${Card} title="Board">
        <div class="form">
          <label><${L} t="board.name" /> <${TextField} term="board.name" value=${p.board.name} label="Board name" required onCommit=${v => c.edit("Board name", ["board", "name"], v, "board.name")} /></label>
          <label><${L} t="board.stepFile" /> <${TextField} term="board.stepFile" value=${p.board.stepFile} label="STEP file" width="22em" onCommit=${v => c.edit("STEP file", ["board", "stepFile"], v, "board.stepFile")} /></label>
          <label><${L} t="board.ambient" /> <${NumField} term="board.ambient" value=${p.board.ambientC} unit="C" label="Ambient" rule=${{ required: true, min: -273.15 }} onCommit=${v => v !== undefined && c.edit("Ambient", ["board", "ambientC"], v, "board.ambientC")} /></label>
          <label><${L} t="board.measuredInputPower" /> <${NumField} term="board.measuredInputPower" value=${p.board.measuredInputPowerW} unit="W" label="Measured board input power" rule=${{ min: 0 }} onCommit=${v => c.edit("Measured input power", ["board", "measuredInputPowerW"], v, "board.measured")} /> <span class="hint">optional: checks the sum of all part ${marked("Pd")}</span></label>
          <label class="wide"><${L} t="board.notes" /> <${TextField} term="board.notes" value=${p.board.notes} label="Notes" multiline width="100%" onCommit=${v => c.edit("Notes", ["board", "notes"], v, "board.notes")} /></label>
        </div>
      <//>
      <${Card} title="Where you are">
        ${(() => { const s = nextStep(states); return s ? html`<p class="next">Next step: <button class="link" onClick=${() => c.go(s.n)}>tab ${s.n}, ${s.title}</button>${s.locked ? html` <span class="hint">(locked: ${marked(s.reason)})</span>` : s.note ? html` <span class="hint">(${marked(s.note)})</span>` : ""}</p>` : html`<p>All steps done.</p>`; })()}
        <ul class="steps">${states.filter(s => s.n > 0 && s.n < 10).map(s => html`<li class=${s.status}><button class="link" onClick=${() => c.go(s.n)}>${s.n} ${s.title}</button> <span class="state" data-status=${stateWord(s)}>${stateWord(s)}</span>${s.note ? html` <span class="hint">${marked(s.note)}</span>` : ""}</li>`)}</ul>
      <//>
      <${Card} title="Files">
        <dl class="kv">
          <dt><${L} t="file.folder" /></dt><dd><${V} t="file.folder" v=${folder || "(files opened by hand)"} /></dd>
          <dt><${L} t="file.project" /></dt><dd><${V} t="file.project" v=${p.savedAt ? `saved ${date(p.savedAt)} by ${p.savedBy}` : "not saved yet"} /></dd>
          <dt><${L} t="file.cad" /></dt><dd><${V} t="file.cad" v=${cad?.import ? `import ${date(cad.import.date) ?? ""}, ${cad.import.bodies?.length ?? 0} bodies${cad.build ? `; build ${date(cad.build.date) ?? ""}` : ""}` : "no import yet"} /></dd>
          <dt><${L} t="file.results" /></dt><dd><${V} t="file.results" v=${results?.runs?.length ? `${results.runs.length} runs, last ${results.runs[results.runs.length - 1].name}` : "no runs yet"} /></dd>
        </dl>
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
      <span class="hint"><${V} t="triage.parts" /> parts, <${L} t="board.totalPd" /> <${V} t="board.totalPd" d=${2} u /></span>
    </div>
    ${p.parts.map((x, i) => {
      const r = reg.parts[x.refDes] || {};
      const k = x.refDes;
      const dup = count[x.refDes] > 1;
      const method = CLASS_METHOD[x.class];
      return html`<section class="card part" key=${x.id}>
        <div class="part-head">
          <span class="field ${dup ? "bad" : ""}" data-term="part.refDes" data-key=${k}>
            <input class="refdes" type="text" value=${x.refDes} aria-label="RefDes" disabled=${locked}
              onChange=${e => { const v = e.currentTarget.value.trim(); if (v && v !== x.refDes) c.editFn(`RefDes ${x.refDes} to ${v}`, q => renameRefDes(q, x.refDes, v)); else e.currentTarget.value = x.refDes; }} />
            ${dup && html`<span class="err">used twice</span>`}
          </span>
          <${TextField} term="part.partNumber" tkey=${k} value=${x.partNumber} label="Part number" disabled=${locked} onCommit=${v => setPart(i, "partNumber", "Part number", v)} />
          <${Select} term="part.class" tkey=${k} value=${x.class} options=${CLASSES} label="Class" blankText="(pick class)" disabled=${locked} onCommit=${v => setPart(i, "class", "Class", v)} />
          <span class="result"><${L} t="part.pd" k=${k}>Pd</${L}> <b><${V} t="part.pd" k=${k} d=${3} u /></b> · <${L} t="part.tjEst" k=${k}>Tj est</${L}> <b><${V} t="part.tjEst" k=${k} d=${1} u /></b> · <${L} t="part.margin" k=${k}>margin</${L}> <b><${V} t="part.margin" k=${k} d=${1} u /></b> <${V} t="part.flag" k=${k} /></span>
          <span class="actions">
            <button disabled=${locked} onClick=${() => c.editFn(`Duplicate ${x.refDes}`, q => duplicatePart(q, x.id))}>Duplicate</button>
            <button disabled=${locked} class="danger" onClick=${() => c.editFn(`Delete ${x.refDes}`, q => deletePart(q, x.id))}>Delete</button>
          </span>
        </div>
        <div class="form row">
          <label><${L} t="part.package" k=${k} /> <${TextField} term="part.package" tkey=${k} value=${x.package} label="Package" width="8em" disabled=${locked} onCommit=${v => setPart(i, "package", "Package", v)} /></label>
          <label><${L} t="part.tjMax" k=${k}>Tj max</${L}> <${NumField} term="part.tjMax" tkey=${k} value=${x.tjMaxC} unit="C" label="Tj max" disabled=${locked} rule=${{ min: -273.15 }} onCommit=${v => setPart(i, "tjMaxC", "Tj max", v)} /></label>
          <label><${L} t="part.rja" k=${k}>Rja</${L}> <${NumField} term="part.rja" tkey=${k} value=${x.rjaCW} unit="CW" label="Rja" disabled=${locked} rule=${{ minEx: 0 }} onCommit=${v => setPart(i, "rjaCW", "Rja", v)} /></label>
          <label><${L} t="part.rjc" k=${k}>Rjc</${L}> <${NumField} term="part.rjc" tkey=${k} value=${x.rjcCW} unit="CW" label="Rjc" disabled=${locked} rule=${{ minEx: 0 }} onCommit=${v => setPart(i, "rjcCW", "Rjc", v)} /></label>
          <label><${L} t="part.rjb" k=${k}>Rjb</${L}> <${NumField} term="part.rjb" tkey=${k} value=${x.rjbCW} unit="CW" label="Rjb" disabled=${locked} rule=${{ minEx: 0 }} onCommit=${v => setPart(i, "rjbCW", "Rjb", v)} /></label>
          <label><${L} t="part.taOverride" k=${k} /> <${NumField} term="part.taOverride" tkey=${k} value=${x.taOverrideC} unit="C" label="Ta override" disabled=${locked} rule=${{ min: -273.15 }} onCommit=${v => setPart(i, "taOverrideC", "Ta override", v)} /></label>
          <span class="calcs"><${L} t="part.ta" k=${k}>Ta</${L}> <${V} t="part.ta" k=${k} d=${1} u /> · <${L} t="part.area" k=${k}>A</${L}> <${V} t="part.area" k=${k} d=${1} u /> · <${L} t="part.density" k=${k}>density</${L}> <${V} t="part.density" k=${k} d=${2} u /> · <${L} t="part.share" k=${k}>share</${L}> <${V} t="part.share" k=${k} d=${3} /></span>
        </div>
        ${r.flag && r.flag !== "Minor" && html`<p class="why"><${L} t="part.why" k=${k} />: <${V} t="part.why" k=${k} /></p>`}
        <${Scroll}><table class="grid loads">
          <thead><tr>${["load.label", "load.rail", "load.method", "load.vin", "load.iin", "load.vout", "load.iout", "load.eff", "load.r", "load.pdMeas", "load.pd"].map(t => html`<th><${L} t=${t} /></th>`)}<th></th></tr></thead>
          <tbody>${x.loads.map((l, j) => {
            const lr = reg.loads[l.id] || {};
            const lk = l.id;
            return html`<tr key=${l.id}>
              <td><${TextField} term="load.label" tkey=${lk} value=${l.label} label="Load label" width="10em" disabled=${locked} onCommit=${v => setLoad(i, j, "label", "Load label", v)} /></td>
              <td><${TextField} term="load.rail" tkey=${lk} value=${l.rail} label="Rail" width="6em" disabled=${locked} onCommit=${v => setLoad(i, j, "rail", "Rail", v)} /></td>
              <td><${Select} term="load.method" tkey=${lk} value=${l.method} options=${METHODS} label="Method" blankText=${method ? `(class: ${method})` : "(class default)"} disabled=${locked} onCommit=${v => setLoad(i, j, "method", "Method", v)} /></td>
              <td><${NumField} term="load.vin" tkey=${lk} value=${l.vinV} unit="V" label="Vin" width="4.5em" disabled=${locked} onCommit=${v => setLoad(i, j, "vinV", "Vin", v)} /></td>
              <td><${NumField} term="load.iin" tkey=${lk} value=${l.iinA} unit="A" label="Iin" width="4.5em" disabled=${locked} rule=${{ min: 0 }} onCommit=${v => setLoad(i, j, "iinA", "Iin", v)} /></td>
              <td><${NumField} term="load.vout" tkey=${lk} value=${l.voutV} unit="V" label="Vout" width="4.5em" disabled=${locked} onCommit=${v => setLoad(i, j, "voutV", "Vout", v)} /></td>
              <td><${NumField} term="load.iout" tkey=${lk} value=${l.ioutA} unit="A" label="Iout" width="4.5em" disabled=${locked} rule=${{ min: 0 }} onCommit=${v => setLoad(i, j, "ioutA", "Iout", v)} /></td>
              <td><${NumField} term="load.eff" tkey=${lk} value=${l.efficiency} unit="" label="Efficiency (0 to 1)" width="4em" disabled=${locked} rule=${{ minEx: 0, max: 1 }} onCommit=${v => setLoad(i, j, "efficiency", "Efficiency", v)} /></td>
              <td><${NumField} term="load.r" tkey=${lk} value=${l.rOhm} unit="ohm" label="R" width="4.5em" disabled=${locked} rule=${{ min: 0 }} onCommit=${v => setLoad(i, j, "rOhm", "R", v)} /></td>
              <td><${NumField} term="load.pdMeas" tkey=${lk} value=${l.pdMeasuredW} unit="W" label="Pd measured" width="4.5em" disabled=${locked} rule=${{ min: 0 }} onCommit=${v => setLoad(i, j, "pdMeasuredW", "Pd measured", v)} /></td>
              <td class="num">${isNum(lr.pdW) ? html`<${V} t="load.pd" k=${lk} d=${3} />` : html`<span class="warn"><${V} t="load.problem" k=${lk} /></span>`}</td>
              <td>${x.loads.length > 1 && html`<button class="small" disabled=${locked} title="Delete this load" onClick=${() => c.editFn("Delete load", q => ({ ...q, parts: q.parts.map(y => (y.id === x.id ? { ...y, loads: y.loads.filter(z => z.id !== l.id) } : y)) }))}>x</button>`}</td>
            </tr>`;
          })}</tbody>
        </table><//>
        <button class="small" disabled=${locked} onClick=${() => c.editFn(`Add load to ${x.refDes}`, q => ({ ...q, parts: q.parts.map(y => (y.id === x.id ? { ...y, loads: [...y.loads, { id: nextLoadId(y) }] } : y)) }))}>Add load</button>
      </section>`;
    })}
    ${!p.parts.length && html`<p class="hint">No parts yet. Press Add part, then type the ${marked("RefDes")} from the schematic.</p>`}`;
}

// ---------- 2 Import geometry ----------
export function ImportTab(c) {
  const { cad, calc } = c;
  const imp = cad?.import;
  return html`
    <${Card} title="In SolidWorks: 2 Import STEP">
      <p>${marked("Import STEP reads the STEP file named on tab 0, finds the PCB and one body per part, and writes them to cad.json. The page reloads that file on its own when the macro has finished.")}</p>
    <//>
    ${imp ? html`
      <${Card} title="What the import found">
        <${Facts} ctx=${c} rows=${[[, "import.model"], [, "import.date", "", undefined, date(imp.date)], [, "pcb.l", "", 1], [, "pcb.w", "", 1], [, "pcb.h", "", 2], [, "pcb.kxy", "", 1], [, "pcb.kz", "", 3], [, "import.bodies"], [, "import.frame"]]} />
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
    ${noBody.length > 0 && html`<p class="warn">${marked(`Parts with Pd and no body: ${noBody.join(", ")}. Link each to its body below.`)}</p>`}
    <${Scroll}><table class="grid">
      <thead><tr><th><${L} t="body.include" /></th><th><${L} t="body.name" /></th><th><${L} t="body.side" /></th><th><${L} t="body.refDes" /></th><th class="num"><${L} t="body.l" /></th><th class="num"><${L} t="body.w" /></th><th class="num"><${L} t="body.h" /></th></tr></thead>
      <tbody>${calc.bodies.map(b => html`<tr key=${b.key} class=${b.include ? "" : "dim"}>
        <td><${Check} term="body.include" tkey=${b.key} value=${b.include} label=${"Include " + b.name} disabled=${locked} onCommit=${v => c.edit(`Include ${shortName(b.name)}`, ["bodies", b.key, "include"], v)} /></td>
        <td><${V} t="body.name" k=${b.key} v=${shortName(b.name)} />${b.manual ? html` <span class="hint">(typed in)</span>` : ""}</td>
        <td><${V} t="body.side" k=${b.key} /></td>
        <td><${Select} term="body.refDes" tkey=${b.key} value=${b.refDes} options=${refs} label="RefDes" blankText="(none)" disabled=${locked} onCommit=${v => c.edit(`Link ${shortName(b.name)}`, ["bodies", b.key, "refDes"], v)} /></td>
        <td class="num"><${V} t="body.l" k=${b.key} d=${2} /></td><td class="num"><${V} t="body.w" k=${b.key} d=${2} /></td><td class="num"><${V} t="body.h" k=${b.key} d=${2} /></td>
      </tr>`)}</tbody>
    </table><//>`;
}
export const shortName = n => (n || "").split("/").pop().replace(/-\d+$/, "");

// ---------- 4 Triage ----------
export function TriageTab(c) {
  const { calc } = c;
  const t = calc.triage, rails = calc.rails;
  const problems = c.blocking;
  return html`
    ${problems.length ? html`<${Card} title="Fix first" cls="attention"><ul>${problems.map(x => html`<li>${marked(x)}</li>`)}</ul><//>` : ""}
    <${Card} title="Overview">
      <${Facts} ctx=${c} rows=${[[, "triage.parts"], [, "board.totalPd", "", 2], [, "triage.critical"], [, "triage.watch"], [, "triage.minor"], [, "triage.noPd"], [, "triage.noRja"], [, "triage.noBody"]]} />
    <//>
    <div class="cols">
      <${Card} title="By Pd"><ol>${t.byPd.map(r => html`<li><${V} t="part.refDes" k=${r} /> <${V} t="part.pd" k=${r} d=${3} u /> <${V} t="part.flag" k=${r} /></li>`)}</ol><//>
      <${Card} title="By power density"><ol>${t.byDensity.map(r => html`<li><${V} t="part.refDes" k=${r} /> <${V} t="part.density" k=${r} d=${2} u /></li>`)}</ol><//>
    </div>
    <${Card} title="Rail budget">
      <${Scroll}><table class="grid"><thead><tr><th><${L} t="rail.name" /></th><th class="num"><${L} t="rail.parts" /></th><th class="num"><${L} t="rail.pd" /> W</th><th class="num"><${L} t="rail.share" /></th></tr></thead>
      <tbody>${rails.rails.map(r => html`<tr><td><${V} t="rail.name" k=${r.rail} /></td><td class="num"><${V} t="rail.parts" k=${r.rail} /></td><td class="num"><${V} t="rail.pd" k=${r.rail} d=${3} /></td><td class="num"><${V} t="rail.share" k=${r.rail} d=${3} /></td></tr>`)}</tbody></table><//>
      <${Facts} ctx=${c} rows=${[[, "rails.sum", "", 3], [, "rails.noRail", "", 3], [, "rails.measured", "", 2], [, "rails.unaccounted", "", 2], [, "rails.unaccountedFrac", "", 3]]} />
    <//>`;
}

// ---------- 5 Hand check: bare board ----------
export function BareTab(c) {
  const { project: p, calc } = c;
  const bb = calc.bareBoard;
  const cols = [["bbp.pd", 3], ["bbp.rTop", 1], ["bbp.rSpread", 2], ["bbp.rBoard", 2], ["bbp.qBoard", 3], ["bbp.tj", 1], ["bbp.margin", 1]];
  return html`
    <${Card} title="Board and air">
      <${Facts} ctx=${c} rows=${[[, "pcb.kxy", "", 1], [, "bb.area", "", 0], [, "bb.hPcb", "", 2], [, "bb.hParts", "", 2], [, "bb.rBoardAir", "", 2],
        [, "bb.rise", "", 1], [, "bb.boardC", "", 1], [, "bb.qBoard", "", 2], [, "bb.fail"], [, "bb.tight"]]} />
    <//>
    <${Scroll}><table class="grid">
      <thead><tr><th>RefDes</th>${cols.map(([t]) => html`<th class="num"><${L} t=${t} /></th>`)}<th><${L} t="bbp.status" /></th></tr></thead>
      <tbody>${p.parts.map(x => bb.parts[x.refDes] && html`<tr><td><${V} t="part.refDes" k=${x.refDes} /></td>${cols.map(([t, d]) => html`<td class="num"><${V} t=${t} k=${x.refDes} d=${d} /></td>`)}<td><${V} t="bbp.status" k=${x.refDes} /></td></tr>`)}</tbody>
    </table><//>
    ${bb.compare?.compared ? html`<${Card} title="Hand check against the simulation">
      <${Facts} ctx=${c} rows=${[[, "bb.compareRun"], [, "bb.compareCheck"], [, "bb.meanGap", "", 1], [, "bb.maxGap", "", 1]]} />
      <${Scroll}><table class="grid"><thead><tr><th>RefDes</th><th class="num"><${L} t="bbp.tj" /></th><th class="num"><${L} t="bbp.simC" /></th><th class="num"><${L} t="bbp.gap" /></th></tr></thead>
      <tbody>${Object.keys(bb.compare.parts).map(r => html`<tr><td><${V} t="part.refDes" k=${r} /></td><td class="num"><${V} t="bbp.tj" k=${r} d=${1} /></td><td class="num"><${V} t="bbp.simC" k=${r} d=${1} /></td><td class="num"><${V} t="bbp.gap" k=${r} d=${1} /></td></tr>`)}</tbody></table><//>
    <//>` : ""}
    <p class="hint">Click any value to see its formula and walk back to its inputs.</p>`;
}

// ---------- 6 Build model ----------
export function BuildTab(c) {
  const b = c.cad?.build;
  return html`
    <${Card} title="In SolidWorks: 3 Build Board"><p>Builds one box per included body and the PCB plate, in a new assembly. <${L} t="build.included" />: <${V} t="build.included" />.</p><//>
    ${b ? html`<${Card} title="Last build">
      <${Facts} ctx=${c} rows=${[[, "build.assembly"], [, "build.date", "", undefined, date(b.date)], [, "build.mode"], [, "build.wanted"], [, "build.added"], [, "build.moved"], [, "build.resized"], [, "build.suppressed"], [, "build.broughtBack"], [, "build.failed"], [, "build.overlaps"]]} />
      ${(b.bodies || []).filter(x => x.status !== "built").map(x => html`<p class="warn">${shortName(x.name)}: <${V} t="buildbody.status" k=${x.key} />${x.reason ? ", " + x.reason : ""}</p>`)}
    <//>` : html`<p class="hint">No build yet.</p>`}`;
}

// ---------- 7 Spreader ----------
export function SpreaderTab(c) {
  const { project: p } = c;
  const sides = ["top", "bottom"].filter(s => p.spreader?.[s] && c.calc.spreader[s]);
  if (!sides.length) return html`<p class="hint">No spreader plate in this project. Plates are optional; the editor comes with W7.</p>`;
  return sides.map(s => {
    const pl = p.spreader[s];
    return html`<${Card} title=${`${s === "top" ? "Top" : "Bottom"} plate${pl.enabled ? "" : " (not enabled)"}`}>
      <${Facts} ctx=${c} rows=${[[, "plate.material", s], [, "plate.k", s], [, "plate.thickness", s], [, "plate.heightMode", s], [, "plate.faceZ", s, 3], [, "plate.covered", s], [, "plate.cuts", s]]} />
      <${Scroll}><table class="grid"><thead><tr><th><${L} t="pchip.refDes" /></th><th><${L} t="pchip.covered" /></th><th class="num"><${L} t="pchip.chipFace" /></th><th class="num"><${L} t="pchip.gap" /></th><th class="num"><${L} t="pchip.timT" /></th><th class="num"><${L} t="pchip.pedestal" /></th><th><${L} t="pchip.check" /></th></tr></thead>
      <tbody>${pl.chips.map((x, i) => { const k = `${s}/${i}`; const r = c.calc.spreader[s].chips[i]; return html`<tr><td><${V} t="pchip.refDes" k=${k} /></td><td><${V} t="pchip.covered" k=${k} /></td><td class="num"><${V} t="pchip.chipFace" k=${k} d=${3} /></td><td class="num"><${V} t="pchip.gap" k=${k} d=${3} /></td><td class="num"><${V} t="pchip.timT" k=${k} d=${3} /></td><td class="num"><${V} t="pchip.pedestal" k=${k} d=${3} /></td><td class=${r.check === "OK" ? "" : "warn"}><${V} t="pchip.check" k=${k} /></td></tr>`; })}</tbody></table><//>
    <//>`;
  });
}

// ---------- 8 Simulation setup ----------
export function SimTab(c) {
  const chips = c.calc.simChips;
  const cols = [["sim.pd", 3], ["sim.rjc", 2], ["sim.h", 2], ["sim.area", 1], ["sim.kz", 2], ["sim.kxy", 2]];
  return html`
    <${Card} title="In SolidWorks: 5 Run Thermal Study"><p>Sets up and runs the steady state study from these values. The editor for h, contacts and mesh comes with W7.</p>
      <${Facts} ctx=${c} rows=${[[, "surf.pcbTopH"], [, "surf.pcbTopEps"], [, "surf.pcbBottomH"], [, "surf.pcbBottomEps"], [, "surf.partsH"], [, "surf.partsEps"], [, "simset.kPartsWithoutHeat"]]} />
    <//>
    <${Scroll}><table class="grid"><thead><tr><th><${L} t="sim.refDes" /></th><th><${L} t="sim.model" /></th>${cols.map(([t]) => html`<th class="num"><${L} t=${t} /></th>`)}<th><${L} t="sim.check" /></th></tr></thead>
    <tbody>${chips.map((x, i) => html`<tr><td><${V} t="sim.refDes" k=${String(i)} /></td><td><${V} t="sim.model" k=${String(i)} /></td>${cols.map(([t, d]) => html`<td class="num"><${V} t=${t} k=${String(i)} d=${d} /></td>`)}<td class=${x.check === "OK" ? "" : "warn"}><${V} t="sim.check" k=${String(i)} /></td></tr>`)}</tbody></table><//>`;
}

// ---------- 9 Results ----------
export function ResultsTab(c) {
  const runs = c.results?.runs || [];
  return [...runs].reverse().map(r => html`<${Card} title=${r.name}>
    <${Facts} ctx=${c} rows=${[[, "run.date", r.name, undefined, date(r.date)], [, "run.studyType", r.name], [, "run.status", r.name], [, "run.energyBalance", r.name, 2],
      [, "run.pcbMin", r.name, 1], [, "run.pcbMax", r.name, 1], [, "run.spreaderMin", r.name, 1], [, "run.spreaderMax", r.name, 1]]} />
    ${r.parts && html`<${Scroll}><table class="grid"><thead><tr><th>RefDes</th><th class="num"><${L} t="runp.peak" /></th><th class="num"><${L} t="runp.avg" /></th><th class="num"><${L} t="part.tjMax">Tj max</${L}></th></tr></thead>
      <tbody>${Object.entries(r.parts).map(([k, v]) => { const tj = c.project.parts.find(x => x.refDes === k)?.tjMaxC; const key = `${r.name}|${k}`; return html`<tr><td><${V} t="part.refDes" k=${k} v=${k} /></td><td class=${`num ${isNum(tj) && v.peakC > tj ? "warn" : ""}`}><${V} t="runp.peak" k=${key} d=${1} /></td><td class="num"><${V} t="runp.avg" k=${key} d=${1} /></td><td class="num">${isNum(tj) ? html`<${V} t="part.tjMax" k=${k} d=${0} />` : ""}</td></tr>`; })}</tbody></table><//>`}
  <//>`);
}

// ---------- 10 Changes ----------
export function ChangesTab(c) {
  const { project: p, cad, now } = c;
  const list = cad?.changes || [];
  if (!p) return html`<p class="hint">Open a project to see its change records.</p>`;
  if (!list.length) return html`<p class="hint">${marked("No change records. Import STEP, Build Board and Sync CAD write them when something moves or changes size.")}</p>`;
  const open = unticked(p, cad);
  const tickAll = () => c.editFn(`Tick ${open} changes`, q => { const checks = { ...(q.checks || {}) }; for (const x of list) if (!checks[x.id]) checks[x.id] = { at: now() }; return { ...q, checks }; });
  const cols = ["change.date", "change.macro", "change.body", "change.change", "change.field"];
  return html`
    <div class="toolbar"><span>${open ? `${open} of ${list.length} not ticked` : `All ${list.length} ticked`}</span>${open > 0 && html`<button onClick=${tickAll}>Tick all</button>`}</div>
    <${Scroll}><table class="grid"><thead><tr><th><${L} t="change.checked" /></th>${cols.map(t => html`<th><${L} t=${t} /></th>`)}<th class="num"><${L} t="change.old" /></th><th class="num"><${L} t="change.new" /></th></tr></thead>
    <tbody>${list.map(x => html`<tr class=${p.checks?.[x.id] ? "dim" : ""}>
      <td><${Check} term="change.checked" tkey=${x.id} value=${!!p.checks?.[x.id]} label=${"Checked " + x.id} onCommit=${v => c.edit(v ? "Tick change" : "Untick change", ["checks", x.id], v ? { at: now() } : undefined)} /></td>
      <td><${V} t="change.date" k=${x.id} v=${date(x.date)} /></td><td><${V} t="change.macro" k=${x.id} /></td><td><${V} t="change.body" k=${x.id} v=${shortName(x.name)} /></td><td><${V} t="change.change" k=${x.id} /></td><td><${V} t="change.field" k=${x.id} /></td>
      <td class="num"><${V} t="change.old" k=${x.id} /></td><td class="num"><${V} t="change.new" k=${x.id} /></td></tr>`)}</tbody></table><//>`;
}

// ---------- 11 Help ----------
export function HelpTab() {
  return html`<${Card} title="Start here">
    <p>This page takes a ${marked("PCB")} from ${marked("STEP")} file to a checked thermal result. You do the engineering (power numbers, datasheet values, judging the results); the SolidWorks macros build the model and run the study.</p>
    <ol>
      <li>Tab 0: choose the board's project folder. One folder per board. The page saves <code>project.json</code> there.</li>
      <li>Tab 1: add each part that makes heat, with its datasheet values and its power.</li>
      <li>Tabs 2 to 9: follow the tabs in order. A grey tab says what to do first and links to it.</li>
    </ol>
    <h4>Every value explains itself</h4>
    <p>Hover over any value or label (or tab to it with the keyboard) to see what it is, its unit and whether you type it or the page works it out. Click a value, or the small i next to a label, to open the Explain panel: the longer text, the formula with your numbers, the inputs (click one to walk back to it) and what uses the value. Show formulas in the top bar writes the formula under every calculated value. Words with a dotted underline are abbreviations: hover for the meaning.</p>
    <h4>Saving</h4>
    <p>Save (Ctrl+S) writes <code>project.json</code>. While there are unsaved edits the page also saves on its own every 30 s. Before the first save of each day it copies the old file to <code>backup\\project_yyyymmdd.json</code>.</p>
    <h4>Undo</h4>
    <p>Ctrl+Z undoes the last edit, Ctrl+Y (or Ctrl+Shift+Z) redoes it. Undo works for the whole session, including deleted parts.</p>
    <h4>Files the macros write</h4>
    <p>The page never writes <code>cad.json</code> or <code>results.json</code>. It checks them every 2 s and reloads them when a macro has finished.</p>
    <p class="hint">The full help, with the heat basics and the glossary, comes with W8.</p>
  <//>`;
}
