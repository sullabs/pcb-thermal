// Tab contents. Every value is a <V> and every label a <L> with a term id from the catalogue, so it has a hover
// text, opens the Explain panel, and shows its formula when Show formulas is on (tools/test_w4.mjs checks this).
// Tabs that keep their own view state (open editors, the selected part) render a small component, because the
// page calls each tab as a plain function. Tabs 6 to 10 and their SolidWorks cards are in steps.js.

import { useState, useMemo, useContext, useEffect } from "preact/hooks";
import { html, NumField, TextField, Select, Check, Card, unitText } from "./ui.js";
import { V, L, marked, Explain, fmtValue } from "./explain.js";
import { CLASSES, METHODS, CLASS_METHOD } from "../calc/constants.js";
import { duplicatePart, deletePart, renameRefDes, nextPartId, setIn, withSimSetup, addManualBody } from "./model.js";
import { swFiles, withFound } from "./swfiles.js";
import { SHEET_COLS, sheetRows, insertPart, movePart, addLoad, moveLoad, deleteLoad, pasteCells, isBlock } from "./sheet.js";
import { readPartsList, applyPartsList, partsListText, suggestLinks, LIST_HEADER } from "./lists.js";
import { valueOf } from "../calc/formulas.js";
import { TERMS } from "../terms/catalogue.js";
import { LEARN } from "../terms/learn.js";
import { nextStep, unticked } from "./tabs.js";
import { isNum } from "../calc/util.js";
import { names } from "./files.js";
import { Scroll, Facts, date, shortName, pctV } from "./bits.js";
import { SwCard } from "./swcard.js";

export { shortName };

// ---------- 0 Project ----------
// Flag limits: CRITICAL and Watch, read by tab 4 (flags) and tab 5 (Tight). On tab 0 with the board settings.
function FlagLimits(c) {
  const { project: p, locked } = c;
  const lim = p.settings.flagLimits;
  const setLim = (lvl, f, label, v) => v !== undefined && c.edit(label, ["settings", "flagLimits", lvl, f], v, `limit.${lvl}.${f}`);
  return html`
    <${Card} title="Flag limits">
      <p class="hint">A part gets a flag when any one of its numbers crosses a limit. Tab 4 shows the flags; the CRITICAL margin also sets Tight on tab 5.</p>
      <table class="grid limits"><thead><tr><th></th><th><${L} t="part.margin">margin below</${L}></th><th><${L} t="part.density">density above</${L}></th><th><${L} t="part.share">share above</${L}></th></tr></thead>
      <tbody>${[["critical", "CRITICAL", "limit.critical"], ["watch", "Watch", "limit.watch"]].map(([lvl, word, t]) => html`<tr><td><${V} t="part.flag" v=${word} /></td>
        <td><${NumField} term=${t + "Margin"} value=${lim[lvl].marginBelowC} unit="dC" label=${word + " margin limit"} rule=${{ required: true }} disabled=${locked} onCommit=${v => setLim(lvl, "marginBelowC", word + " margin limit", v)} /></td>
        <td><${NumField} term=${t + "Density"} value=${lim[lvl].densityAboveWcm2} unit="W/cm²" accepts="none" label=${word + " density limit"} rule=${{ required: true, min: 0 }} disabled=${locked} onCommit=${v => setLim(lvl, "densityAboveWcm2", word + " density limit", v)} /></td>
        <td><${NumField} term=${t + "Share"} value=${lim[lvl].shareAbove} unit="" label=${word + " share limit"} width="4em" rule=${{ required: true, min: 0, max: 1 }} disabled=${locked} onCommit=${v => setLim(lvl, "shareAbove", word + " share limit", v)} /></td></tr>`)}</tbody></table>
    <//>`;
}

export function ProjectTab(c) {
  const { project: p, cad, results, states, folder } = c;
  const files = c.files || names("");
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
          <dt><${L} t="file.project" /></dt><dd><${V} t="file.project" v=${`${files.project}: ` + (p.savedAt ? `saved ${date(p.savedAt)} by ${p.savedBy}` : "not saved yet")} /></dd>
          <dt><${L} t="file.cad" /></dt><dd><${V} t="file.cad" v=${`${files.cad}: ` + (cad?.import ? `import ${date(cad.import.date) ?? ""}, ${cad.import.bodies?.length ?? 0} bodies${cad.build ? `; build ${date(cad.build.date) ?? ""}` : ""}` : "no import yet")} /></dd>
          <dt><${L} t="file.results" /></dt><dd><${V} t="file.results" v=${`${files.results}: ` + (results?.runs?.length ? `${results.runs.length} runs, last ${results.runs[results.runs.length - 1].name}` : "no runs yet")} /></dd>
        </dl>
        ${c.backupCard}
      <//>
      ${FlagLimits(c)}`}`;
}
export const stateWord = s => ({ locked: "locked", ready: "to do", done: "done", attention: "needs attention" }[s.status]);

// ---------- Learn mode ----------
// "What this step is for" and "What to look at" at the top of a tab (design 5.3).
export function LearnBox({ id }) {
  const l = LEARN[id];
  if (!l) return null;
  return html`<section class="learn" aria-label="Learn">
    <p><b>What this step is for.</b> ${marked(l.purpose)}</p>
    <p class="learn-h"><b>What to look at</b></p>
    <ul>${l.look.map(x => html`<li>${marked(x)}</li>`)}</ul>
  </section>`;
}

// ---------- 1 Parts and power ----------
// What each power method needs (calc/power.js). `used`: fields the method reads; `need`: the short list to fill.
const METHOD_NEEDS = {
  "LDO / linear reg": { used: ["vinV", "iinA", "voutV", "ioutA"], need: ["vinV", "voutV", "ioutA"], text: "Vin, Vout and Iout (or Vin, Iin, Vout and Iout)" },
  "Switching regulator": { used: ["vinV", "iinA", "voutV", "ioutA", "efficiency"], need: ["voutV", "ioutA", "efficiency"], text: "Vout, Iout and efficiency (or Vin, Iin, Vout and Iout)" },
  "Digital load (V x I)": { used: ["vinV", "iinA", "voutV", "ioutA"], need: ["vinV", "iinA"], text: "Vin and Iin (Vout and Iout only if it passes power on)" },
  "Resistive (I squared R)": { used: ["iinA", "ioutA", "rOhm"], need: ["iinA", "rOhm"], text: "a current (Iin or Iout) and R" },
  "Manual entry": { used: [], need: ["pdMeasuredW"], text: "Pd measured" },
};

// The register is a sheet like the workbook's: one row per load, the part's own cells span its load rows. Click a
// cell and type; Enter and the arrow keys move between cells; a block copied from Excel pastes from the cell it
// lands in, with new parts below the last row. Drag the grip of a part (or of a load) to move it; right-click a
// row, or click its number, for insert, duplicate, move and delete.
const MORE = [["part.ta", 1], ["part.area", 1], ["part.density", 2], ["part.share", 3]];

export function PartsTab(c) {
  return html`<${PartsSheet} c=${c} />`;
}

function PartsSheet({ c }) {
  const { project: p, calc, locked } = c;
  const reg = calc.register;
  const [more, setMore] = useState(false);
  const [menu, setMenu] = useState(null);
  const [drag, setDrag] = useState(null);
  const [over, setOver] = useState(null);
  const [note, setNote] = useState(null);
  const count = {};
  for (const x of p.parts) count[x.refDes] = (count[x.refDes] || 0) + 1;
  const rows = sheetRows(p);
  const firstRow = [];
  rows.forEach((r, n) => { if (r.j === 0) firstRow[r.i] = n; });
  const setPart = (i, key, label, v) => c.edit(label, ["parts", i, key], v, `parts.${p.parts[i].id}.${key}`);
  const setLoad = (i, j, key, label, v) => c.edit(label, ["parts", i, "loads", j, key], v, `loads.${p.parts[i].loads[j].id}.${key}`);

  useEffect(() => {
    if (!menu) return;
    const close = e => { if (e.type === "keydown" ? e.key === "Escape" : !e.target.closest?.(".sheet-menu")) setMenu(null); };
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", close);
    return () => { document.removeEventListener("mousedown", close); document.removeEventListener("keydown", close); };
  }, [menu]);

  // The new row exists only after the next render: try for up to half a second.
  const focusCell = (id, col = "refDes", tries = 20) => setTimeout(() => {
    const el = document.querySelector(`.sheet tbody[data-id="${id}"] td[data-col="${col}"] :is(input, select)`);
    if (el) { el.focus(); el.select?.(); el.scrollIntoView({ block: "nearest" }); } else if (tries > 0) focusCell(id, col, tries - 1);
  }, 25);
  const insertAt = at => {
    const id = nextPartId(p);
    c.editFn("Add part", q => insertPart(q, at).project);
    focusCell(id);
  };
  const act = (label, fn) => { setMenu(null); c.editFn(label, fn); };

  // Keyboard moves like Excel: Enter and the up/down arrows go down a column, left/right at the edge of the text
  // go along the row. A part cell spans its load rows, so a row without it falls back to the part's first row.
  const cellAt = (n, key) => {
    const r = rows[n];
    if (!r) return null;
    const tb = `.sheet tbody[data-id="${r.part.id}"]`;
    return document.querySelector(`${tb} tr[data-row="${n}"] td[data-col="${key}"] :is(input, select)`)
      || document.querySelector(`${tb} tr[data-row="${firstRow[r.i]}"] td[data-col="${key}"] :is(input, select)`);
  };
  const go = el => { if (el) { el.focus(); el.select?.(); return true; } return false; };
  const onKeyDown = e => {
    const td = e.target.closest?.("td[data-col]");
    const tr = td?.closest("tr[data-row]");
    if (!td || !tr || e.altKey || e.ctrlKey || e.metaKey) return;
    const n = Number(tr.dataset.row), key = td.dataset.col;
    const isSel = e.target.tagName === "SELECT";
    let dn = 0;
    if (e.key === "Enter") dn = e.shiftKey ? -1 : 1;
    else if (!isSel && e.key === "ArrowDown") dn = 1;
    else if (!isSel && e.key === "ArrowUp") dn = -1;
    if (dn) {
      const r = rows[n];
      const step = SHEET_COLS.find(x => x.key === key)?.where === "part" ? "part" : "row";
      for (let k = n + dn; k >= 0 && k < rows.length; k += dn) {
        if (step === "part" && rows[k].i === r.i) continue;
        if (step === "part" && rows[k].j !== 0) continue;
        if (go(cellAt(k, key))) { e.preventDefault(); return; }
      }
      return;
    }
    if (e.key === "ArrowLeft" || e.key === "ArrowRight") {
      const t = e.target, d = e.key === "ArrowLeft" ? -1 : 1;
      const whole = t.tagName === "INPUT" && t.selectionStart === 0 && t.selectionEnd === t.value.length;
      if (t.tagName === "INPUT" && !whole) {
        if (t.selectionStart !== t.selectionEnd) return;
        if (d < 0 ? t.selectionStart !== 0 : t.selectionEnd !== t.value.length) return;
      }
      for (let k = SHEET_COLS.findIndex(x => x.key === key) + d; k >= 0 && k < SHEET_COLS.length; k += d) {
        if (go(cellAt(n, SHEET_COLS[k].key))) { e.preventDefault(); return; }
      }
    }
  };
  const onPaste = e => {
    const td = e.target.closest?.("td[data-col]");
    const tr = td?.closest("tr[data-row]");
    const text = e.clipboardData?.getData("text/plain") ?? "";
    if (!td || !tr || locked || !isBlock(text)) return;
    e.preventDefault();
    const n = Number(tr.dataset.row), key = td.dataset.col;
    const res = pasteCells(p, n, key, text);
    const lines = text.replace(/\r/g, "").replace(/\n$/, "").split("\n").length;
    c.editFn(`Paste ${lines} rows`, q => pasteCells(q, n, key, text).project);
    setNote({ text: `Pasted ${lines} row${lines === 1 ? "" : "s"}${res.added ? `, ${res.added} new part${res.added === 1 ? "" : "s"}` : ""}.`, problems: res.problems });
  };

  // Drag a part by its row number grip, a load by its own grip. The drop goes above or below the row under the mouse.
  const startDrag = (e, d) => {
    if (locked) { e.preventDefault(); return; }
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", "");
    const g = e.currentTarget.closest(d.kind === "part" ? "tbody" : "tr");
    if (g && e.dataTransfer.setDragImage) e.dataTransfer.setDragImage(g, 12, 12);
    setDrag(d);
  };
  const half = e => { const b = e.currentTarget.getBoundingClientRect(); return e.clientY > b.top + b.height / 2; };
  const partOver = (e, i) => {
    if (drag?.kind !== "part") return;
    e.preventDefault();
    const after = half(e);
    if (over?.i !== i || over?.after !== after || over?.j !== undefined) setOver({ i, after });
  };
  const loadOver = (e, i, j) => {
    if (drag?.kind !== "load" || drag.i !== i) return;
    e.preventDefault();
    e.stopPropagation();
    const after = half(e);
    if (over?.i !== i || over?.j !== j || over?.after !== after) setOver({ i, j, after });
  };
  const drop = e => {
    e.preventDefault();
    if (drag && over) {
      if (drag.kind === "part") {
        let to = over.i + (over.after ? 1 : 0);
        if (drag.i < to) to--;
        if (to !== drag.i) c.editFn(`Move ${p.parts[drag.i].refDes}`, q => movePart(q, drag.i, to));
      } else if (over.j !== undefined) {
        let to = over.j + (over.after ? 1 : 0);
        if (drag.j < to) to--;
        if (to !== drag.j) c.editFn(`Move a load of ${p.parts[drag.i].refDes}`, q => moveLoad(q, p.parts[drag.i].id, drag.j, to));
      }
    }
    setDrag(null); setOver(null);
  };
  const endDrag = () => { setDrag(null); setOver(null); };

  const cell = (col, x, l, i, j) => {
    const k = x.refDes, lk = l.id;
    const label = TERMS[col.term]?.label || col.key;
    const onPart = col.where === "part";
    const tkey = onPart ? k : lk;
    const value = onPart ? x[col.field] : l[col.field];
    const commit = v => (onPart ? setPart(i, col.field, label, v) : setLoad(i, j, col.field, label, v));
    const span = onPart ? x.loads.length : 1;
    const tdc = `${col.kind === "num" ? "num" : ""} ${col.kind === "refdes" ? "refcol" : ""}`;
    if (col.kind === "refdes") {
      const dup = count[x.refDes] > 1;
      return html`<td rowspan=${span} data-col=${col.key} class=${tdc}><span class="field ${dup ? "bad" : ""}" data-term="part.refDes" data-key=${k}>
        <input class="refdes" type="text" value=${x.refDes} aria-label="RefDes" disabled=${locked} style=${{ width: col.width }}
          onChange=${e => { const v = e.currentTarget.value.trim(); if (v && v !== x.refDes) c.editFn(`RefDes ${x.refDes} to ${v}`, q => renameRefDes(q, x.refDes, v)); else e.currentTarget.value = x.refDes; }} />
        ${dup && html`<span class="err">used twice</span>`}</span></td>`;
    }
    if (col.kind === "select") {
      const method = CLASS_METHOD[x.class];
      const blankText = col.key === "class" ? "" : method ? `(class: ${method})` : "(class default)";
      return html`<td rowspan=${span} data-col=${col.key} class=${tdc} style=${{ minWidth: col.width }}><${Select} term=${col.term} tkey=${tkey} value=${value} options=${col.options} label=${label} blankText=${blankText} disabled=${locked} onCommit=${commit} /></td>`;
    }
    if (col.kind === "num") {
      let cls = "";
      if (!onPart) {
        const lr = reg.loads[l.id] || {};
        const needs = METHOD_NEEDS[lr.method];
        const missing = !isNum(lr.pdW);
        const f = col.field;
        cls = !needs || f === "pdMeasuredW" ? (missing && needs?.need.includes(f) ? "need" : "")
          : missing && needs.need.includes(f) && !isNum(l[f]) ? "need" : needs.used.includes(f) ? "" : "unused";
      }
      return html`<td rowspan=${span} data-col=${col.key} class=${tdc}><${NumField} term=${col.term} tkey=${tkey} value=${value} unit=${col.unit} label=${label} width=${col.width} cls=${cls}
        disabled=${locked} rule=${col.rule} placeholder=${col.key === "rjbCW" || col.key === "taOverrideC" ? "optional" : undefined} onCommit=${commit} /></td>`;
    }
    return html`<td rowspan=${span} data-col=${col.key} class=${tdc}><${TextField} term=${col.term} tkey=${tkey} value=${value} label=${label} width=${col.width} disabled=${locked}
      placeholder=${col.key === "notes" ? "datasheet page, load case" : undefined} onCommit=${commit} /></td>`;
  };

  const partCols = SHEET_COLS.filter(x => x.where === "part" && x.key !== "notes");
  const loadCols = SHEET_COLS.filter(x => x.where === "load");
  const notesCol = SHEET_COLS.find(x => x.key === "notes");
  // Section of a column, for its colour: part, datasheet, loads, results (calculated), notes.
  const DATASHEET = new Set(["tjMaxC", "rjaCW", "rjcCW", "rjbCW", "taOverrideC"]);
  const grp = col => "g-" + (col.key === "notes" ? "notes" : col.where === "load" ? "load" : DATASHEET.has(col.key) ? "ds" : "part");
  const head = col => html`<th class=${`${col.kind === "num" ? "num" : col.kind === "refdes" ? "refcol" : ""} ${grp(col)}`}><${L} t=${col.term}>${col.short || TERMS[col.term]?.label}</${L}>${col.unit ? html`<span class="unit">${unitText(col.unit)}</span>` : ""}</th>`;
  const m = menu && p.parts[menu.i];

  return html`
    <div class="toolbar">
      <button disabled=${locked} onClick=${() => insertAt(p.parts.length)}>Add part</button>
      <${PasteList} c=${c} />
      <button class=${`small ${more ? "on" : ""}`} aria-pressed=${more} onClick=${() => setMore(!more)}>More columns</button>
      <span class="hint"><${V} t="triage.parts" /> parts, <${L} t="board.totalPd" /> <${V} t="board.totalPd" d=${2} u /></span>
    </div>
    <p class="hint">Click a cell and type. Enter or the arrow keys move between cells. Drag the grip at the left of a row to move it. Right-click a row (or click its number) to insert, duplicate, move or delete. A block copied from Excel pastes from the cell you paste into; extra lines become new parts.</p>
    ${note && html`<div class="hint sheet-note">${note.text} ${note.problems.length ? html`<span class="warn">${note.problems.length} cells not taken:</span><ul class="warn">${note.problems.slice(0, 8).map(x => html`<li>${x}</li>`)}</ul>` : ""}
      <button class="small" onClick=${() => setNote(null)}>OK</button></div>`}
    <div class="sheet-wrap">
      <table class=${`sheet ${drag ? "dragging" : ""}`} onKeyDown=${onKeyDown} onPaste=${onPaste} onDragEnd=${endDrag}>
        <thead>
          <tr class="groups"><th rowspan="2" class="rownum-h"></th><th class="refcol g-part">Part</th><th colspan=${3} class="g-part g-cont"></th><th colspan=${5} class="g-ds">Datasheet</th><th class="grip-h g-load" rowspan="2"></th>
            <th colspan=${loadCols.length} class="g-load">Loads: one row per heat source</th><th colspan=${5 + (more ? MORE.length + 1 : 0)} class="calc g-res">Results</th><th class="g-notes">Notes</th></tr>
          <tr>${partCols.map(head)}${loadCols.map(head)}
            <th class="num calc g-res"><${L} t="load.pd" /><span class="unit">W</span></th>
            <th class="num calc g-res"><${L} t="part.pd">Part Pd</${L}><span class="unit">W</span></th>
            <th class="num calc g-res"><${L} t="part.tjEst">Tj est</${L}><span class="unit">C</span></th>
            <th class="num calc g-res"><${L} t="part.margin" /><span class="unit">C</span></th>
            <th class="calc g-res"><${L} t="part.flag" /></th>
            ${more && MORE.map(([t]) => html`<th class="num calc g-res"><${L} t=${t} /></th>`)}
            ${more && html`<th class="calc g-res"><${L} t="part.why" /></th>`}
            ${head(notesCol)}</tr>
        </thead>
        ${p.parts.map((x, i) => {
          const k = x.refDes;
          const n0 = firstRow[i];
          const span = x.loads.length;
          const r = reg.parts[k] || {};
          const dropCls = drag?.kind === "part" && over?.i === i && over.j === undefined ? (over.after ? "drop-after" : "drop-before") : "";
          return html`<tbody key=${x.id} data-id=${x.id} class=${`${dropCls} ${drag?.kind === "part" && drag.i === i ? "dragged" : ""} ${i % 2 ? "odd" : ""}`}
            onDragOver=${e => partOver(e, i)} onDrop=${drop}>
            ${x.loads.map((l, j) => {
              const lr = reg.loads[l.id] || {};
              const needs = METHOD_NEEDS[lr.method];
              const lk = l.id;
              const ldrop = drag?.kind === "load" && over?.i === i && over?.j === j ? (over.after ? "drop-after" : "drop-before") : "";
              return html`<tr key=${l.id} data-row=${n0 + j} class=${ldrop} onContextMenu=${e => { if (locked) return; e.preventDefault(); setMenu({ i, j, x: e.clientX, y: e.clientY }); }}
                onDragOver=${e => loadOver(e, i, j)}>
                ${j === 0 && html`<td rowspan=${span} class="rownum hint">
                  <span class="grip" draggable=${!locked} title="Drag to move this part" onDragStart=${e => startDrag(e, { kind: "part", i })}>⠿</span>
                  <button class="rowbtn" title="Row menu" disabled=${locked} onClick=${e => { const b = e.currentTarget.getBoundingClientRect(); setMenu({ i, j: 0, x: b.right, y: b.top }); }}>${i + 1}</button></td>`}
                ${j === 0 && partCols.map(col => cell(col, x, l, i, j))}
                <td class="grip-c hint">${span > 1 && html`<span class="grip" draggable=${!locked} title="Drag to move this load" onDragStart=${e => { e.stopPropagation(); startDrag(e, { kind: "load", i, j }); }}>⠿</span>`}</td>
                ${loadCols.map(col => cell(col, x, l, i, j))}
                <td class="num calc">${isNum(lr.pdW) ? html`<${V} t="load.pd" k=${lk} d=${3} />` : html`<span class="warn"><${V} t="load.problem" k=${lk} /></span>${needs ? html`<span class="needs">needs ${marked(needs.text)}</span>` : ""}`}</td>
                ${j === 0 && html`
                  <td rowspan=${span} class="num calc"><b><${V} t="part.pd" k=${k} d=${3} /></b></td>
                  <td rowspan=${span} class="num calc"><${V} t="part.tjEst" k=${k} d=${1} /></td>
                  <td rowspan=${span} class="num calc"><${V} t="part.margin" k=${k} d=${1} /></td>
                  <td rowspan=${span} class="calc"><${V} t="part.flag" k=${k} /></td>
                  ${more && MORE.map(([t, d]) => html`<td rowspan=${span} class="num calc"><${V} t=${t} k=${k} d=${d} /></td>`)}
                  ${more && html`<td rowspan=${span} class="calc why-c">${r.flag && r.flag !== "Minor" ? html`<${V} t="part.why" k=${k} />` : ""}</td>`}
                  ${cell(notesCol, x, l, i, j)}`}
              </tr>`;
            })}
          </tbody>`;
        })}
      </table>
      ${!locked && html`<button class="addrow" onClick=${() => insertAt(p.parts.length)}>+ Add part</button>`}
    </div>
    ${!p.parts.length && html`<p class="hint">No parts yet. Press Add part, then type the ${marked("RefDes")} from the schematic. Or paste the whole list from Excel.</p>`}
    ${m && html`<div class="sheet-menu" role="menu" style=${{ left: Math.min(menu.x, innerWidth - 230) + "px", top: Math.min(menu.y, innerHeight - 330) + "px" }}>
      <div class="menu-h">${m.refDes}${m.loads.length > 1 ? `, load ${menu.j + 1}` : ""}</div>
      <button role="menuitem" onClick=${() => { setMenu(null); insertAt(menu.i); }}>Insert part above</button>
      <button role="menuitem" onClick=${() => { setMenu(null); insertAt(menu.i + 1); }}>Insert part below</button>
      <button role="menuitem" onClick=${() => act(`Duplicate ${m.refDes}`, q => duplicatePart(q, m.id))}>Duplicate part</button>
      <button role="menuitem" disabled=${menu.i === 0} onClick=${() => act(`Move ${m.refDes} up`, q => movePart(q, menu.i, menu.i - 1))}>Move part up</button>
      <button role="menuitem" disabled=${menu.i === p.parts.length - 1} onClick=${() => act(`Move ${m.refDes} down`, q => movePart(q, menu.i, menu.i + 1))}>Move part down</button>
      <hr />
      <button role="menuitem" onClick=${() => act(`Add load to ${m.refDes}`, q => addLoad(q, m.id, menu.j + 1))}>Add load below</button>
      ${m.loads.length > 1 && html`
        <button role="menuitem" disabled=${menu.j === 0} onClick=${() => act(`Move a load of ${m.refDes}`, q => moveLoad(q, m.id, menu.j, menu.j - 1))}>Move load up</button>
        <button role="menuitem" disabled=${menu.j === m.loads.length - 1} onClick=${() => act(`Move a load of ${m.refDes}`, q => moveLoad(q, m.id, menu.j, menu.j + 1))}>Move load down</button>
        <button role="menuitem" class="danger" onClick=${() => act(`Delete a load of ${m.refDes}`, q => deleteLoad(q, m.id, m.loads[menu.j].id))}>Delete this load</button>`}
      <hr />
      <button role="menuitem" class="danger" onClick=${() => act(`Delete ${m.refDes}`, q => deletePart(q, m.id))}>Delete part</button>
    </div>`}`;
}

// Paste a parts list (from Excel, tab separated, or a CSV text). One line per load; lines with the same RefDes
// make one part.
function PasteList({ c }) {
  const [open, setOpen] = useState(false);
  const [text, setText] = useState("");
  const [copied, setCopied] = useState("");
  const list = useMemo(() => (text.trim() ? readPartsList(text) : null), [text]);
  const refs = list ? [...new Set(list.rows.map(r => r.part.refDes))] : [];
  const have = new Set(c.project.parts.map(x => x.refDes));
  const add = () => {
    let out;
    c.editFn(`Paste ${refs.length} parts`, q => { out = applyPartsList(q, list); return out.project; });
    setOpen(false); setText("");
  };
  const copy = async () => {
    try { await navigator.clipboard.writeText(partsListText(c.project)); setCopied("Copied. Paste it into Excel, fill it, then paste it back here."); }
    catch { setCopied("The browser did not allow copying."); }
  };
  if (!open) return html`<button disabled=${c.locked} onClick=${() => setOpen(true)}>Paste a parts list</button>`;
  return html`<div class="paste card" data-term="file.partsList">
    <h3>Paste a parts list</h3>
    <p class="hint">Copy the cells from Excel (first row: the column names) and paste them here. One row per load: rows with the same ${marked("RefDes")} make one part. Columns read: <code>${LIST_HEADER.join(", ")}</code>. Units may be typed in the cells (mA, mW, oz). A RefDes already on the list gets the new values.</p>
    <textarea rows="8" aria-label="Parts list" value=${text} onInput=${e => setText(e.currentTarget.value)} placeholder=${LIST_HEADER.join("\t")}></textarea>
    ${list && html`<p class="hint">${refs.length} parts (${refs.filter(r => !have.has(r)).length} new, ${refs.filter(r => have.has(r)).length} already on the list), ${list.rows.filter(r => Object.keys(r.load).length).length} loads.</p>
      ${list.problems.length > 0 && html`<ul class="warn">${list.problems.slice(0, 12).map(x => html`<li>${marked(x)}</li>`)}</ul>`}`}
    <p class="hint">
      <button class="primary" disabled=${!refs.length} onClick=${add}>Add to the list</button>
      <button onClick=${copy}>Copy the current list</button>
      <button onClick=${() => { setOpen(false); setText(""); }}>Cancel</button>
      ${copied && html` ${copied}`}
    </p>
  </div>`;
}

// ---------- 2 Import geometry ----------
export function ImportTab(c) {
  const { project: p, cad, calc, locked } = c;
  const imp = cad?.import;
  const s = p.settings;
  const set = (path, label, v) => c.edit(label, ["settings", ...path], v, "settings." + path.join("."));
  return html`
    <${SwCard} c=${c} id="ImportStep">
      <${Facts} ctx=${c} rows=${[[, "board.stepFile"]]} />
    <//>
    <${Card} title="Import settings">
      <div class="form">
        <label><${L} t="set.pcbNameContains" /> <${TextField} term="set.pcbNameContains" value=${s.import.pcbNameContains} label="PCB body name has" required width="8em" disabled=${locked} onCommit=${v => set(["import", "pcbNameContains"], "PCB body name", v)} /></label>
        <label><${L} t="set.minSize" /> <${NumField} term="set.minSize" value=${s.import.minSizeMm} unit="mm" label="Smallest body" rule=${{ required: true, min: 0 }} disabled=${locked} onCommit=${v => v !== undefined && set(["import", "minSizeMm"], "Smallest body", v)} /></label>
        <label><${L} t="set.topLevelOnly" /> <span class="field"><${Check} term="set.topLevelOnly" value=${s.import.topLevelOnly} label="One body per component" disabled=${locked} onCommit=${v => set(["import", "topLevelOnly"], "One body per component", v)} /></span></label>
        <label><${L} t="set.tolPos" /> <${NumField} term="set.tolPos" value=${s.tolPosMm} unit="mm" label="Move tolerance" rule=${{ required: true, min: 0 }} disabled=${locked} onCommit=${v => v !== undefined && set(["tolPosMm"], "Move tolerance", v)} /></label>
        <label><${L} t="set.tolSize" /> <${NumField} term="set.tolSize" value=${s.tolSizeMm} unit="mm" label="Size tolerance" rule=${{ required: true, min: 0 }} disabled=${locked} onCommit=${v => v !== undefined && set(["tolSizeMm"], "Size tolerance", v)} /></label>
      </div>
      <p class="hint">The macro reads these when it runs. Change them, save, then run Import STEP again.</p>
    <//>
    ${imp ? html`
      <${Card} title="What the import found">
        <${Facts} ctx=${c} rows=${[[, "import.model"], [, "import.date", "", undefined, date(imp.date)], [, "pcb.l", "", 1], [, "pcb.w", "", 1], [, "pcb.h", "", 2], [, "import.bodies"], [, "import.frame"]]} />
        <${BoardDrawing} calc=${calc} />
      <//>` : html`<p class="hint">No import yet.</p>`}`;
}

// The board seen from the top and, when there are bottom side bodies, from below (mirrored left to right, as
// the board looks when turned over). A turned body (rotDeg, ownL, ownW from the import) is drawn turned.
export function BoardDrawing({ calc }) {
  const p = calc.pcb;
  if (!(p.l > 0 && p.w > 0)) return null;
  const bodies = calc.bodies.filter(b => isNum(b.box.xMin) && isNum(b.box.yMin) && b.box.l > 0 && b.box.w > 0);
  const top = bodies.filter(b => b.side !== "Bottom");
  const bottom = bodies.filter(b => b.side === "Bottom" || b.side === "Through");
  const hasBottom = bodies.some(b => b.side === "Bottom");
  // One frame for both views, wide enough for bodies that hang over the board edge (connectors).
  const xs = [0, p.l], ys = [0, p.w];
  for (const b of bodies) for (const [x, y] of footprint(b).pts) { xs.push(x); ys.push(y); }
  const ext = { x0: Math.min(...xs), x1: Math.max(...xs), y0: Math.min(...ys), y1: Math.max(...ys) };
  return html`<div class="boards">
      <${BoardView} p=${p} bodies=${top} ext=${ext} title="Top view" note="seen from above" />
      ${hasBottom && html`<${BoardView} p=${p} bodies=${bottom} ext=${ext} mirror title="Bottom view" note="seen from below, so left and right are swapped" />`}
    </div>
    <div class="hint legend">
      <span><svg viewBox="0 0 14 10"><rect class="pcb" x=".5" y=".5" width="13" height="9" /></svg> board</span>
      <span><svg viewBox="0 0 14 10"><rect class="body linked" x=".5" y=".5" width="13" height="9" /></svg> body linked to a part (tab 3), label is its RefDes</span>
      <span><svg viewBox="0 0 14 10"><rect class="body" x=".5" y=".5" width="13" height="9" /></svg> body not linked to a part</span>
      <span><svg viewBox="0 0 14 10"><rect class="body skip" x=".5" y=".5" width="13" height="9" /></svg> body left out of the model (Include off)</span>
      <span><svg viewBox="0 0 14 10"><rect class="body linked turned" x=".5" y=".5" width="13" height="9" /></svg> turned body, drawn at its angle</span>
    </div>`;
}

// Same test as PT_ShapeIsRotated in the macros: draw the turned shape only when it still fits the box used.
function footprint(b) {
  const x = b.box, cx = x.xMin + x.l / 2, cy = x.yMin + x.w / 2;
  const a = (b.rotDeg || 0) * Math.PI / 180, c = Math.abs(Math.cos(a)), s = Math.abs(Math.sin(a));
  const turned = Math.abs(b.rotDeg || 0) >= 0.05 && b.ownL > 0 && b.ownW > 0
    && Math.abs(b.ownL * c + b.ownW * s - x.l) <= 0.02 && Math.abs(b.ownL * s + b.ownW * c - x.w) <= 0.02;
  const hx = (turned ? b.ownL : x.l) / 2, hy = (turned ? b.ownW : x.w) / 2, an = turned ? a : 0;
  const pts = [[-hx, -hy], [hx, -hy], [hx, hy], [-hx, hy]]
    .map(([dx, dy]) => [cx + dx * Math.cos(an) - dy * Math.sin(an), cy + dx * Math.sin(an) + dy * Math.cos(an)]);
  return { cx, cy, pts, turned, short: Math.min(hx, hy) * 2 };
}

function BoardView({ p, bodies, ext, mirror, title, note }) {
  const pad = 2;
  // Board X right and Y up; screen Y down. The bottom view also swaps X.
  const sx = x => (mirror ? p.l - x : x), sy = y => p.w - y;
  const fp = bodies.map(b => ({ b, f: footprint(b) }));
  return html`<figure class="board-view">
    <figcaption><b>${title}</b> <span class="hint">${note}</span></figcaption>
    <svg class="board" viewBox=${`${(mirror ? p.l - ext.x1 : ext.x0) - pad} ${p.w - ext.y1 - pad} ${ext.x1 - ext.x0 + 2 * pad} ${ext.y1 - ext.y0 + 2 * pad}`} role="img" aria-label=${`${title} of the board and its bodies`}>
      <rect class="pcb" x="0" y="0" width=${p.l} height=${p.w} />
      ${fp.map(({ b, f }) => html`<polygon class=${`body ${b.include ? "" : "skip"} ${b.refDes ? "linked" : ""} ${f.turned ? "turned" : ""}`}
        points=${f.pts.map(([x, y]) => `${sx(x).toFixed(2)},${sy(y).toFixed(2)}`).join(" ")}><title>${b.refDes ? b.refDes + ": " : ""}${b.name}${f.turned ? ` (turned ${Math.round(b.rotDeg * 10) / 10}°)` : ""}</title></polygon>`)}
      <g class="hint tags">
        ${fp.filter(({ b, f }) => b.refDes && f.short >= 2).map(({ b, f }) => html`<text x=${sx(f.cx).toFixed(2)} y=${sy(f.cy).toFixed(2)}
          font-size=${Math.min(3, Math.max(1.2, f.short * 0.45)).toFixed(2)}>${b.refDes}</text>`)}
      </g>
    </svg>
  </figure>`;
}

// ---------- 3 Link bodies ----------
const BOX_FIELDS = [["l", "L"], ["w", "W"], ["h", "H"], ["xMin", "X min"], ["yMin", "Y min"], ["zMin", "Z min"]];
const SIDES = ["Top", "Bottom", "Board", "Through"];

export function LinkTab(c) {
  return html`<${LinkView} c=${c} />`;
}
function LinkView({ c }) {
  const { project: p, calc, locked } = c;
  const [show, setShow] = useState("all");
  const [open, setOpen] = useState(null);
  useEffect(() => { const f = e => setOpen(e.detail); window.addEventListener("pt-open-body", f); return () => window.removeEventListener("pt-open-body", f); }, []);
  const refs = p.parts.map(x => x.refDes);
  const noBody = p.parts.filter(x => isNum(calc.register.parts[x.refDes]?.pdW) && !calc.bodies.some(b => b.refDes === x.refDes)).map(x => x.refDes);
  const sugg = suggestLinks(p, calc);
  const suggOf = Object.fromEntries(sugg.map(s => [s.key, s]));
  const geoOf = Object.fromEntries(calc.geometry.bodies.map(b => [b.key, b]));
  const shown = calc.bodies.filter(b => show === "all" || (show === "included" && b.include) || (show === "linked" && b.refDes)
    || (show === "problems" && geoOf[b.key]?.check && geoOf[b.key].check !== "OK"));
  const applyAll = () => c.editFn(`Link ${sugg.length} suggested bodies`, q => sugg.reduce((a, s) => setIn(setIn(a, ["bodies", s.key, "refDes"], s.refDes), ["bodies", s.key, "include"], true), q));
  const addBody = () => { let key; c.editFn("Add a body by hand", q => { const r = addManualBody(q, calc.pcb.h); key = r.key; return r.project; }); setTimeout(() => setOpen(key)); };
  return html`
    ${noBody.length > 0 && html`<p class="warn">${marked(`Parts with Pd and no body: ${noBody.join(", ")}. Link each to its body below, or add a body by hand.`)}</p>`}
    <div class="toolbar">
      ${sugg.length > 0 && html`<button class="primary" disabled=${locked} onClick=${applyAll} title=${sugg.map(s => `${s.refDes}: ${s.why}`).join("\n")}>Link the ${sugg.length} suggested bodies</button>`}
      <button disabled=${locked} onClick=${addBody}>Add a body by hand</button>
      <span class="hint">Show:</span>
      ${[["all", "all"], ["included", "included"], ["linked", "with a part"], ["problems", "with a problem"]].map(([v, t]) => html`<button class=${`small ${show === v ? "on" : ""}`} aria-pressed=${show === v} onClick=${() => setShow(v)}>${t}</button>`)}
    </div>
    <${Scroll}><table class="grid bodies">
      <thead><tr><th><${L} t="body.include" /></th><th><${L} t="body.name" /></th><th><${L} t="body.side" /></th><th><${L} t="body.refDes" /></th>
        ${["body.l", "body.w", "body.h", "body.pd", "body.density"].map(t => html`<th class="num"><${L} t=${t} /></th>`)}<th><${L} t="body.check" /></th><th></th></tr></thead>
      <tbody>${shown.map(b => {
        const s = suggOf[b.key];
        const chk = geoOf[b.key]?.check;
        return html`<tr key=${b.key} class=${b.include ? "" : "dim"}>
          <td><${Check} term="body.include" tkey=${b.key} value=${b.include} label=${"Include " + b.name} disabled=${locked} onCommit=${v => c.edit(`Include ${shortName(b.name)}`, ["bodies", b.key, "include"], v)} /></td>
          <td><${V} t="body.name" k=${b.key} v=${shortName(b.name)} />${b.manual ? html` <span class="hint">(typed in)</span>` : ""}</td>
          <td><${V} t="body.side" k=${b.key} /></td>
          <td><${Select} term="body.refDes" tkey=${b.key} value=${b.refDes} options=${refs} label="RefDes" blankText="(none)" disabled=${locked} onCommit=${v => c.editFn(`Link ${shortName(b.name)}`, q => { const r = setIn(q, ["bodies", b.key, "refDes"], v); return v ? setIn(r, ["bodies", b.key, "include"], true) : r; })} />
            ${s && html` <button class="small sugg" disabled=${locked} title=${"Suggested: " + s.why} onClick=${() => c.editFn(`Link ${shortName(b.name)} to ${s.refDes}`, q => setIn(setIn(q, ["bodies", b.key, "refDes"], s.refDes), ["bodies", b.key, "include"], true))}>${s.refDes}?</button>`}</td>
          <td class="num"><${V} t="body.l" k=${b.key} d=${2} /></td><td class="num"><${V} t="body.w" k=${b.key} d=${2} /></td><td class="num"><${V} t="body.h" k=${b.key} d=${2} /></td>
          <td class="num">${b.refDes ? html`<${V} t="body.pd" k=${b.key} d=${3} />` : ""}</td><td class="num">${b.refDes ? html`<${V} t="body.density" k=${b.key} d=${2} />` : ""}</td>
          <td class=${chk && chk !== "OK" ? "warn" : ""}><${V} t="body.check" k=${b.key} /></td>
          <td><button class="small" onClick=${() => setOpen(open === b.key ? null : b.key)} aria-expanded=${open === b.key}>${open === b.key ? "Close" : "Edit"}</button></td>
        </tr>${open === b.key && html`<tr class="edit"><td colspan="11"><${BodyEdit} c=${c} b=${b} close=${() => setOpen(null)} /></td></tr>`}`;
      })}</tbody>
    </table><//>
    <${Card} title="Bodies and heat">
      <${Facts} ctx=${c} rows=${[[, "geo.included"], [, "board.totalPd", "", 3], [, "geo.placed", "", 3], [, "geo.notPlaced", "", 3], [, "geo.noBody"], [, "geo.problems"]]} />
    <//>
    <${PcbCard} c=${c} />`;
}

function BodyEdit({ c, b, close }) {
  const { project: p, cad, locked } = c;
  const imported = cad?.import?.bodies?.find(x => x.key === b.key);
  const base = b.manual ? ["bodies", b.key, "manual"] : ["bodies", b.key];
  const boxPath = f => (b.manual ? [...base, "box", f] : [...base, "override", f]);
  const ch = p.bodies[b.key] || {};
  return html`<div class="form body-edit">
    <label><${L} t="bodyset.name" k=${b.key} /> <${TextField} term="bodyset.name" tkey=${b.key} value=${b.manual ? ch.manual.name : ch.nameOverride} label="Body name" width="18em" required=${b.manual}
      placeholder=${imported ? shortName(imported.name) : ""} disabled=${locked} onCommit=${v => c.edit("Body name", b.manual ? [...base, "name"] : [...base, "nameOverride"], v, `bodies.${b.key}.name`)} /></label>
    ${b.manual && html`<label><${L} t="bodyset.side" k=${b.key} /> <${Select} term="bodyset.side" tkey=${b.key} value=${ch.manual.side} options=${SIDES} label="Side" disabled=${locked} onCommit=${v => c.edit("Body side", [...base, "side"], v)} /></label>`}
    ${BOX_FIELDS.map(([f, lab]) => html`<label><${L} t=${"bodyset." + f} k=${b.key} /> <${NumField} term=${"bodyset." + f} tkey=${b.key} value=${b.manual ? ch.manual.box[f] : ch.override?.[f]} unit="mm" label=${lab}
      placeholder=${imported && isNum(imported.box?.[f]) ? String(+imported.box[f].toFixed(3)) : ""} rule=${{ required: !!b.manual, ...(["l", "w", "h"].includes(f) ? { minEx: 0 } : {}) }} disabled=${locked}
      onCommit=${v => (v !== undefined || !b.manual) && c.edit(`${lab} of ${shortName(b.name)}`, boxPath(f), v, `bodies.${b.key}.${f}`)} /></label>`)}
    <p class="hint wide">${b.manual ? "A body you added: it is built like an imported one." : "Blank fields keep the imported value (shown in grey). Typed values replace it in the model."}
      ${!b.manual && (ch.override || ch.nameOverride) && html` <button class="small" disabled=${locked} onClick=${() => c.editFn(`Clear typed values of ${shortName(b.name)}`, q => setIn(setIn(q, [...base, "override"], undefined), [...base, "nameOverride"], undefined))}>Clear typed values</button>`}
      ${b.manual && html` <button class="small danger" disabled=${locked} onClick=${() => { close(); c.editFn(`Delete ${shortName(b.name)}`, q => setIn(q, ["bodies", b.key], undefined)); }}>Delete this body</button>`}</p>
  </div>`;
}

function PcbCard({ c }) {
  const { project: p, cad, locked } = c;
  const imp = cad?.import?.pcb?.box || {};
  const st = p.pcbStack;
  const setStack = (f, label, v) => v !== undefined && c.edit(label, ["pcbStack", f], v, "pcbStack." + f);
  return html`<${Card} title="PCB">
    <div class="cols">
      <div>
        <h4>Size</h4>
        <div class="form">
          ${[["l", "pcbov.l", "pcb.l"], ["w", "pcbov.w", "pcb.w"], ["h", "pcbov.h", "pcb.h"]].map(([f, t, used]) => html`<label><${L} t=${t} /> <${NumField} term=${t} value=${p.pcbOverride?.[f]} unit="mm" label=${TERMS[t].label}
            placeholder=${isNum(imp[f]) ? String(+imp[f].toFixed(3)) : ""} rule=${{ minEx: 0 }} disabled=${locked} onCommit=${v => c.edit(TERMS[t].label, ["pcbOverride", f], v, "pcbOverride." + f)} />
            <span>used: <${V} t=${used} d=${2} u /></span></label>`)}
        </div>
        <p class="hint">Blank uses the size from Import STEP (shown in grey).</p>
      </div>
      <div>
        <h4>Stack</h4>
        <div class="form">
          <label><${L} t="stack.layers" /> <${NumField} term="stack.layers" value=${st.copperLayers} unit="" label="Copper layers" width="4em" rule=${{ required: true, integer: true, min: 1 }} disabled=${locked} onCommit=${v => setStack("copperLayers", "Copper layers", v)} /></label>
          <label><${L} t="stack.cuT" /> <${NumField} term="stack.cuT" value=${st.copperThicknessMm} unit="mm" accepts="copperMm" label="Copper thickness" rule=${{ required: true, minEx: 0 }} disabled=${locked} onCommit=${v => setStack("copperThicknessMm", "Copper thickness", v)} /></label>
          <label><${L} t="stack.coverage" /> <${NumField} term="stack.coverage" value=${st.copperCoverage} unit="" label="Copper coverage" width="4em" rule=${{ required: true, minEx: 0, max: 1 }} disabled=${locked} onCommit=${v => setStack("copperCoverage", "Copper coverage", v)} /></label>
          <label><${L} t="stack.cuK" /> <${NumField} term="stack.cuK" value=${st.copperK} unit="WmK" label="Copper k" rule=${{ required: true, minEx: 0 }} disabled=${locked} onCommit=${v => setStack("copperK", "Copper k", v)} /></label>
          <label><${L} t="stack.dielK" /> <${NumField} term="stack.dielK" value=${st.dielectricK} unit="WmK" label="Dielectric k" rule=${{ required: true, minEx: 0 }} disabled=${locked} onCommit=${v => setStack("dielectricK", "Dielectric k", v)} /></label>
        </div>
        <p class="hint">Copper thickness accepts oz (1 oz = 0.035 mm) and um.</p>
        <${Facts} ctx=${c} rows=${[[, "pcb.cuTotal", "", 3], [, "pcb.cuFraction", "", 4], [, "pcb.kxy", "", 2], [, "pcb.kz", "", 3]]} />
      </div>
    </div>
  <//>`;
}

// ---------- 4 Triage ----------
export function TriageTab(c) {
  const { project: p, calc, locked } = c;
  const t = calc.triage, rails = calc.rails;
  const problems = c.blocking;
  const top = t.byPd.find(r => isNum(calc.register.parts[r]?.pdW));
  const dense = t.byDensity.find(r => isNum(calc.register.parts[r]?.densityWcm2));
  return html`
    ${problems.length ? html`<${Card} title="Fix first" cls="attention"><ul>${problems.map(x => html`<li>${marked(x)}</li>`)}</ul><//>` : ""}
    ${c.learn && top && html`<${Card} title="What the numbers say" cls="reading">
      <p><${V} nofx t="part.refDes" k=${top} /> makes ${pctV("part.share", top, calc.register.parts[top].share)} of the board's heat (<${V} nofx t="part.pd" k=${top} d=${2} u />).
        ${dense && html` <${V} nofx t="part.refDes" k=${dense} /> has the highest power density, <${V} nofx t="part.density" k=${dense} d=${2} u />: the least area to lose its heat through.`}</p>
      <p>${t.critical > 0 ? html`<${V} nofx t="triage.critical" /> part${t.critical === 1 ? " is" : "s are"} CRITICAL: start with ${t.critical === 1 ? "it" : "them"} on <button class="link" onClick=${() => c.go(5)}>tab 5</button>.` : "No part is CRITICAL on this first estimate."}
        ${t.noRja > 0 ? html` <${V} nofx t="triage.noRja" /> part${t.noRja === 1 ? " has" : "s have"} no ${marked("Rja")}, so ${t.noRja === 1 ? "its" : "their"} margin is not checked here.` : ""}</p>
      ${isNum(rails.unaccountedFrac) && html`<p>The parts account for all but ${pctV("rails.unaccountedFrac", "", rails.unaccountedFrac)} of the measured input power.${Math.abs(rails.unaccountedFrac) > 0.2 ? " That is a lot: a heat source may be missing from tab 1, or one is counted twice." : ""}</p>`}
    <//>`}
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
      ${!isNum(rails.measuredW) && html`<p class="hint">Type the measured board input power on tab 0 to check that the parts account for all of it.</p>`}
    <//>
    <p class="hint">The flag limits are on <button class="link" onClick=${() => c.go(0)}>tab 0</button>.</p>`;
}

// ---------- 5 Hand check: bare board ----------
export function BareTab(c) {
  return html`<${BareView} c=${c} />`;
}
function BareView({ c }) {
  const { project: p, calc, results, locked } = c;
  const bb = calc.bareBoard;
  const withTj = p.parts.filter(x => isNum(bb.parts[x.refDes]?.tjC)).map(x => x.refDes);
  const worst = [...withTj].sort((a, z) => (isNum(bb.parts[a].marginC) ? bb.parts[a].marginC : 1e9) - (isNum(bb.parts[z].marginC) ? bb.parts[z].marginC : 1e9));
  const [sel, setSel] = useState(null);
  const k = sel && bb.parts[sel] ? sel : worst[0];
  const cols = [["bbp.pd", 3], ["bbp.rTop", 1], ["bbp.rSpread", 2], ["bbp.rBoard", 2], ["bbp.qBoard", 3], ["bbp.tj", 1], ["bbp.margin", 1]];
  const s = p.simSetup?.surfaces || {};
  const setSim = (path, label, v) => v !== undefined && c.editFn(label, q => setIn(withSimSetup(q), ["simSetup", ...path], v), "simSetup." + path.join("."));
  const runs = (results?.runs || []).map(r => r.name);
  const surf = [["pcbTop", "surf.pcbTopH", "surf.pcbTopEps", "PCB top"], ["pcbBottom", "surf.pcbBottomH", "surf.pcbBottomEps", "PCB bottom"], ["parts", "surf.partsH", "surf.partsEps", "parts"]];
  const env = calc.flow?.env;
  return html`
    ${c.learn && html`<${Reading} c=${c} parts=${worst} />`}
    ${p.simSetup?.studyType === "Flow" && env && !env.natural && html`<p class="warn">The h values here are for a flat board, parts up, in still air. The Flow study on tab 8 has the board ${env.orientation}${env.airSpeedMs > 0 ? ` with ${env.airSpeedMs} m/s of air` : ""}, so this hand check and the Flow run are not the same case.</p>`}
    <${Card} title="Air and surfaces">
      <table class="grid limits"><thead><tr><th></th><th><${L} t="surf.pcbTopH">h</${L}></th><th><${L} t="surf.pcbTopEps">emissivity</${L}></th></tr></thead>
      <tbody>${surf.map(([f, th, te, name]) => html`<tr><td>${name}</td>
        <td><${NumField} term=${th} value=${s[f]?.h} unit="W/(m² K)" accepts="none" label=${"h, " + name} rule=${{ required: true, min: 0 }} disabled=${locked} onCommit=${v => setSim(["surfaces", f, "h"], "h, " + name, v)} /></td>
        <td><${NumField} term=${te} value=${s[f]?.emissivity} unit="" label=${"Emissivity, " + name} width="4em" rule=${{ required: true, min: 0, max: 1 }} disabled=${locked} onCommit=${v => setSim(["surfaces", f, "emissivity"], "Emissivity, " + name, v)} /></td></tr>`)}</tbody></table>
      <div class="form row">
        <label><${L} t="board.ambient" /> <span><${V} t="board.ambient" d=${1} u /> <span class="hint">(tab 0)</span></span></label>
        <label><${L} t="hand.surfaceRise" /> <${NumField} term="hand.surfaceRise" value=${p.handCheck?.surfaceRiseK} unit="dC" label="Surface rise for radiation" rule=${{ min: 0 }} disabled=${locked} onCommit=${v => c.edit("Surface rise", ["handCheck", "surfaceRiseK"], v, "handCheck.surfaceRiseK")} /></label>
        <label><${L} t="simset.kPartsWithoutHeat" /> <${NumField} term="simset.kPartsWithoutHeat" value=${p.simSetup?.kPartsWithoutHeat} unit="WmK" label="k of parts without Rjc" rule=${{ required: true, minEx: 0 }} disabled=${locked} onCommit=${v => setSim(["kPartsWithoutHeat"], "k of parts without Rjc", v)} /></label>
      </div>
      <p class="hint">The simulation on tab 8 uses the same h and emissivity, so the hand check and the study see the same air.</p>
    <//>
    <${Card} title="Board and air">
      <${Facts} ctx=${c} rows=${[[, "pcb.kxy", "", 1], [, "bb.area", "", 0], [, "bb.hRadTop", "", 2], [, "bb.hRadBottom", "", 2], [, "bb.hPcb", "", 2], [, "bb.hRadParts", "", 2], [, "bb.hParts", "", 2], [, "bb.rBoardAir", "", 2],
        [, "bb.rise", "", 1], [, "bb.boardC", "", 1], [, "bb.qBoard", "", 2], [, "bb.boardShare", "", 2], [, "bb.fail"], [, "bb.tight"]]} />
    <//>
    <${Scroll}><table class="grid">
      <thead><tr><th><${L} t="part.refDes" /></th>${cols.map(([t]) => html`<th class="num"><${L} t=${t} /></th>`)}<th><${L} t="bbp.status" /></th><th></th></tr></thead>
      <tbody>${p.parts.map(x => bb.parts[x.refDes] && html`<tr class=${x.refDes === k ? "sel" : ""}><td><${V} t="part.refDes" k=${x.refDes} /></td>${cols.map(([t, d]) => html`<td class="num"><${V} t=${t} k=${x.refDes} d=${d} /></td>`)}<td><${V} t="bbp.status" k=${x.refDes} /></td>
        <td>${isNum(bb.parts[x.refDes].tjC) && html`<button class="small" aria-pressed=${x.refDes === k} onClick=${() => setSel(x.refDes)}>Network</button>`}</td></tr>`)}</tbody>
    </table><//>
    ${k && html`<${Card} title=${`Heat paths of ${k}`}>
      <div class="cols net">
        <${NetDrawing} k=${k} />
        <div>
          <${Facts} ctx=${c} rows=${[[, "bbp.model", k], [, "bbp.pd", k, 3], [, "bbp.rjcUsed", k, 3], [, "bbp.rJTop", k, 3], [, "bbp.exposed", k, 1], [, "bbp.rSurfaceAir", k, 2], [, "bbp.rTop", k, 2],
            [, "bbp.rJBoard", k, 3], [, "bbp.footprint", k, 1], [, "bbp.share", k, 0], [, "bbp.eps", k, 4], [, "bbp.tau", k, 4], [, "bbp.bi", k, 4], [, "bbp.lambda", k, 4], [, "bbp.phi", k, 4], [, "bbp.rSpread", k, 3], [, "bbp.rBoard", k, 3],
            [, "bbp.f", k, 3], [, "bbp.c", k, 4], [, "bbp.fPd", k, 3], [, "bbp.qBoard", k, 3], [, "bbp.qTop", k, 3], [, "bbp.boardShare", k, 2], [, "bbp.tj", k, 1], [, "bbp.margin", k, 1], [, "bbp.status", k]]} />
        </div>
      </div>
    <//>`}
    <${Card} title="Hand check against the simulation">
      <div class="form row"><label><${L} t="hand.compareRun" /> <${Select} term="hand.compareRun" value=${p.handCheck?.compareRun} options=${runs} label="Run to compare with" blankText="(newest run)" disabled=${locked}
        onCommit=${v => c.edit("Run to compare with", ["handCheck", "compareRun"], v)} /></label></div>
      ${bb.compare?.compared ? html`
        <${Facts} ctx=${c} rows=${[[, "bb.compareRun"], [, "bb.compareCheck"], [, "bb.meanGap", "", 1], [, "bb.maxGap", "", 1]]} />
        <${Scroll}><table class="grid"><thead><tr><th><${L} t="part.refDes" /></th><th class="num"><${L} t="bbp.tj" /></th><th class="num"><${L} t="bbp.simC" /></th><th class="num"><${L} t="bbp.gap" /></th></tr></thead>
        <tbody>${Object.keys(bb.compare.parts).map(r => html`<tr><td><${V} t="part.refDes" k=${r} /></td><td class="num"><${V} t="bbp.tj" k=${r} d=${1} /></td><td class="num"><${V} t="bbp.simC" k=${r} d=${1} /></td><td class="num"><${V} t="bbp.gap" k=${r} d=${1} /></td></tr>`)}</tbody></table><//>`
      : html`<p class="hint">No simulation run to compare with yet. The study comes on tab 8.</p>`}
    <//>
    <p class="hint">Click any value to see its formula and walk back to its inputs.</p>`;
}

// Plain reading of the hand check (Learn mode): the board, then each part from the lowest margin.
function Reading({ c, parts }) {
  const bb = c.calc.bareBoard;
  if (!isNum(bb.board.boardC)) return null;
  return html`<${Card} title="What the numbers say" cls="reading">
    <p>The board sits at <${V} nofx t="bb.boardC" d=${1} u />, <${V} nofx t="bb.rise" d=${1} u /> above the air. ${pctV("bb.boardShare", "", valueOfC(c, "bb.boardShare"))} of all the heat goes down into the board and leaves from its two faces; the rest leaves from the part tops.</p>
    <ul>${parts.slice(0, 8).map(k => {
      const r = bb.parts[k];
      const share = valueOfC(c, "bbp.boardShare", k);
      return html`<li><${V} nofx t="part.refDes" k=${k} /> ${isNum(r.marginC) ? (r.marginC >= 0 ? html`is <${V} nofx t="bbp.margin" k=${k} d=${1} u /> below its ${marked("Tj max")}.` : html`is over its ${marked("Tj max")}: margin <${V} nofx t="bbp.margin" k=${k} d=${1} u />.`) : html`has no ${marked("Tj max")} to check against.`}
        ${!isNum(share) ? "" : share < 0 ? " The board around it is warmer than it would make it alone, so it takes heat in from its neighbours: cooling the parts near it helps it most."
          : share >= 0.5 ? html` ${pctV("bbp.boardShare", k, share)} of its heat goes down into the board, so more copper under it, thermal vias or a bigger board would help more than a heatsink.`
          : html` ${pctV("bbp.topShare", k, valueOfC(c, "bbp.topShare", k))} of its heat leaves from its top, so a heatsink or a spreader on top (<button class="link" onClick=${() => c.go(7)}>tab 7</button>) would help more than extra copper.`}</li>`;
    })}</ul>
  <//>`;
}
const valueOfC = (c, t, k = "") => c.value(t, k);

// The two heat paths of one part, drawn as a resistor network with the numbers of this board.
function NetDrawing({ k }) {
  const X = useContext(Explain);
  if (!X.ctx) return null;
  const val = (t, kk) => valueOf(X.ctx, t, kk);
  const num = (x, y, t, kk = "", d = 2, anchor = "start") => html`<text x=${x} y=${y} text-anchor=${anchor} class="nv" data-term=${t} data-key=${kk} tabindex="0" role="button"
    onClick=${() => X.open(t, kk)} onKeyDown=${e => { if (e.key === "Enter") X.open(t, kk); }}>${isNum(val(t, kk)) ? fmtValue(val(t, kk), d) : "?"} ${TERMS[t].unit}</text>`;
  const res = (x, y, label, t, kk) => html`<rect class="res" x=${x - 8} y=${y - 18} width="16" height="36" rx="3" />
    <text class="rl" x=${x - 14} y=${y + 4} text-anchor="end">${label}</text>${num(x + 14, y + 4, t, kk, 2)}`;
  return html`<svg class="net" viewBox="0 0 560 330" role="img" aria-label=${"Heat paths of " + k}>
    <line class="node" x1="30" y1="24" x2="530" y2="24" /><text class="nl" x="30" y="16">Air</text>${num(62, 16, "board.ambient", "", 1)}
    <line class="wire" x1="150" y1="24" x2="150" y2="300" />
    <line class="wire" x1="430" y1="24" x2="430" y2="300" />
    ${res(150, 66, "R surface to air", "bbp.rSurfaceAir", k)}
    ${res(150, 120, "R junction to top", "bbp.rJTop", k)}
    <text class="flow" x="164" y="96" data-term="bbp.qTop" data-key=${k}>↑ Q top</text>${num(220, 96, "bbp.qTop", k, 3)}
    <circle class="tj" cx="150" cy="165" r="7" /><text class="nl" x="136" y="170" text-anchor="end">Junction</text>${num(164, 170, "bbp.tj", k, 1)}
    ${res(150, 210, "R junction to board", "bbp.rJBoard", k)}
    <text class="flow" x="164" y="238" data-term="bbp.qBoard" data-key=${k}>↓ Q board</text>${num(234, 238, "bbp.qBoard", k, 3)}
    ${res(150, 266, "R spreading", "bbp.rSpread", k)}
    ${res(430, 162, "R board to air", "bb.rBoardAir", "")}
    <line class="node" x1="30" y1="300" x2="530" y2="300" /><text class="nl" x="30" y="320">Board (shared by all parts)</text>${num(212, 320, "bb.boardC", "", 1)}
    <text class="hint-svg" x="290" y="292" text-anchor="middle">heat from the other parts joins here</text>
  </svg>`;
}

// ---------- 6 to 10: steps.js ----------
export { BuildTab, SpreaderTab, SimTab, ResultsTab, ChangesTab } from "./steps.js";

// ---------- 11 Help: help.js ----------
export { HelpTab } from "./help.js";
