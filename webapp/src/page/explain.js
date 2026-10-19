// Explaining every value (design section 5): hover text, the Explain panel with the formula in numbers and the
// trace back to the inputs, Show formulas, abbreviations and status words.
//
//   <V t="part.tjEst" k="U1" d=1 />   a value: hover text, click opens the Explain panel, formula when switched on
//   <L t="part.rja">Rja</L>          a label or column header: hover text and an info button
//   marked(text)                       text with abbreviations marked (dotted underline, meaning on hover)

import { createContext } from "preact";
import { useContext, useEffect, useState } from "preact/hooks";
import { html } from "./ui.js";
import { TERMS, GLOSSARY, STATUS } from "../terms/catalogue.js";
import { valueOf, evaluate, usedBy } from "../calc/formulas.js";
import { isNum } from "../calc/util.js";
import { editPlace } from "./where.js";

export const Explain = createContext({ ctx: null, open: () => {}, showFx: false });

export const KIND = { input: "you type this", calculated: "calculated", cad: "from SolidWorks", run: "from a simulation run", constant: "constant" };

export function sig(x, n = 4) {
  if (!isNum(x)) return "?";
  if (x === 0) return "0";
  const a = Math.abs(x);
  if (a >= 1e6 || a < 1e-3) return x.toExponential(n - 1).replace(/\.?0+e/, "e").replace("e", " × 10^");
  return String(+x.toPrecision(n));
}

export function fmtValue(v, d) {
  if (typeof v === "boolean") return v ? "yes" : "no";
  if (!isNum(v)) return v ?? "";
  if (d === undefined) return sig(v, 4);
  const a = Math.abs(v);
  if (a !== 0 && (a < 10 ** -d || a >= 1e7)) return v.toExponential(2);
  return v.toFixed(d);
}

const unitOf = id => TERMS[id]?.unit || "";
const withUnit = (v, id) => { const u = unitOf(id); return u && isNum(v) ? `${sig(v)} ${u}` : typeof v === "boolean" ? (v ? "yes" : "no") : isNum(v) ? sig(v) : v ?? "(blank)"; };

// The right side of the formula with the numbers put in place of the symbols.
export function withNumbers(expr, vals) {
  const syms = Object.keys(vals).sort((a, b) => b.length - a.length);
  if (!syms.length) return expr;
  const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(`(?<![\\w])(${syms.map(esc).join("|")})(?![\\w])`, "g");
  return expr.replace(re, s => { const v = vals[s]; return isNum(v) ? (v < 0 ? `(${sig(v)})` : sig(v)) : typeof v === "string" ? v : "?"; });
}

// "Tj est = Ta + Pd × Rja = 40 + 3.168 × 14 = 84.35 C"
export function formulaLine(e, term) {
  if (!e) return "";
  const lhs = e.text.split(" = ")[0];
  if (!e.fn) return e.text;
  const nums = withNumbers(e.expr, e.vals);
  const res = isNum(e.value) ? `${sig(e.value)}${unitOf(term) ? " " + unitOf(term) : ""}` : "";
  const mid = nums === e.expr || nums === sig(e.value) ? "" : ` = ${nums}`;
  return `${lhs} = ${e.expr}${mid}${res ? " = " + res : ""}`;
}

// Which part, load, body or run a key names, in words.
export function keyText(ctx, term, key) {
  if (!key) return "";
  const p = ctx?.project;
  if (term.startsWith("load.") && p) {
    for (const x of p.parts) { const j = x.loads.findIndex(l => l.id === key); if (j >= 0) return x.loads.length > 1 ? `${x.refDes}, ${x.loads[j].label || "load " + (j + 1)}` : x.refDes; }
  }
  if (term.startsWith("body.") || term.startsWith("buildbody.")) { const b = ctx?.calc.bodies.find(r => r.key === key); if (b) return b.name.split("/").pop(); }
  if (term.startsWith("sim.")) return ctx?.calc.simChips[+key]?.refDes || key;
  if (term.startsWith("pchip.")) { const [side, i] = key.split("/"); return `${p?.spreader?.[side]?.chips[+i]?.refDes || ""} (${side} plate)`; }
  if (term.startsWith("plate.")) return key + " plate";
  if (term.startsWith("change.")) return "change " + key;
  return key.replace("|", ", ");
}

// ---------- abbreviations ----------
const GKEYS = Object.keys(GLOSSARY).sort((a, b) => b.length - a.length);
const GRE = new RegExp(`(?<![\\w])(${GKEYS.map(s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")})(?![\\w])`, "g");
export function marked(text) {
  if (typeof text !== "string" || !text) return text;
  const out = [];
  let last = 0;
  for (const m of text.matchAll(GRE)) {
    if (m.index > last) out.push(text.slice(last, m.index));
    out.push(html`<abbr class="gl" data-gloss=${m[1]} tabindex="0">${m[1]}</abbr>`);
    last = m.index + m[1].length;
  }
  if (last < text.length) out.push(text.slice(last));
  return out;
}

// ---------- a value ----------
// nofx: no formula line even with Show formulas on (inside a sentence, where it would break the reading).
export function V({ t, k = "", v, d, u, cls = "", nofx }) {
  const X = useContext(Explain);
  const val = v !== undefined ? v : X.ctx ? valueOf(X.ctx, t, k) : undefined;
  const term = TERMS[t];
  const shown = fmtValue(val, d);
  const e = X.showFx && !nofx && term?.kind === "calculated" && X.ctx ? evaluate(X.ctx, t, k) : null;
  const isStatus = typeof val === "string" && STATUS[val];
  const why = t === "part.flag" && X.ctx ? valueOf(X.ctx, "part.why", k) : undefined;
  return html`<span class=${`v ${term?.kind || ""} ${cls}`} data-term=${t} data-key=${k} data-why=${why} data-nofx=${nofx ? "1" : undefined} tabindex="0" role="button"
      onClick=${ev => { ev.stopPropagation(); X.open(t, k); }} onKeyDown=${ev => { if (ev.key === "Enter") X.open(t, k); }}>${
      isStatus ? html`<span class=${`status ${statusCls(val)}`} data-status=${val}>${val}</span>` : typeof val === "string" ? marked(shown) : shown}${
      u && unitOf(t) && shown !== "" ? html`<span class="unit"> ${unitOf(t)}</span>` : ""}${
      e?.fn ? html`<span class="fx">${formulaLine(e, t)}</span>` : ""}</span>`;
}
export const statusCls = w => ({ CRITICAL: "crit", FAIL: "crit", Watch: "watch", Tight: "watch", Minor: "minor", OK: "ok", "No Pd": "watch" }[w] || "plain");

// ---------- a label ----------
export function L({ t, k = "", children }) {
  const X = useContext(Explain);
  return html`<span class="lbl" data-term=${t} data-key=${k}>${children ?? TERMS[t]?.label}<button class="info" type="button" tabindex="-1" aria-label=${"Explain " + (TERMS[t]?.label || t)}
    onClick=${ev => { ev.preventDefault(); ev.stopPropagation(); X.open(t, k); }}>i</button></span>`;
}

// ---------- hover text: one box for the whole page ----------
// on: false when the user switched hover texts off (Hover in the top bar); the Explain panel still opens on click.
export function Tip({ on = true }) {
  const [tip, setTip] = useState(null);
  useEffect(() => {
    if (!on) { setTip(null); return; }
    const show = e => {
      const el = e.target.closest?.("[data-term],[data-gloss],[data-status]");
      if (!el) return;
      const r = el.getBoundingClientRect();
      const at = { x: Math.min(r.left, innerWidth - 380), y: r.bottom + 6 < innerHeight - 140 ? r.bottom + 6 : Math.max(8, r.top - 150) };
      const st = e.target.closest("[data-status]"), gl = e.target.closest("[data-gloss]");
      if (gl) return setTip({ at, title: gl.dataset.gloss, body: GLOSSARY[gl.dataset.gloss] });
      if (st && STATUS[st.dataset.status]) {
        const owner = st.closest("[data-term]");
        const why = owner?.dataset.why || null;
        return setTip({ at, title: st.dataset.status, body: STATUS[st.dataset.status], extra: why });
      }
      const t = TERMS[el.dataset.term];
      if (!t) return;
      setTip({ at, title: t.label + (t.sym && t.sym !== t.label ? ` (${t.sym})` : ""), body: t.short, unit: t.unit, kind: KIND[t.kind],
        range: t.range ? `typical ${t.range[0]} to ${t.range[1]}${t.unit ? " " + t.unit : ""}` : null });
    };
    const hide = e => { if (!e.relatedTarget || !e.relatedTarget.closest?.("[data-term],[data-gloss],[data-status]")) setTip(null); };
    const esc = e => { if (e.key === "Escape") setTip(null); };
    document.addEventListener("mouseover", show);
    document.addEventListener("focusin", show);
    document.addEventListener("mouseout", hide);
    document.addEventListener("focusout", hide);
    document.addEventListener("keydown", esc);
    return () => { for (const [n, f] of [["mouseover", show], ["focusin", show], ["mouseout", hide], ["focusout", hide], ["keydown", esc]]) document.removeEventListener(n, f); };
  }, [on]);
  if (!tip) return null;
  return html`<div class="tip" role="tooltip" style=${{ left: tip.at.x + "px", top: tip.at.y + "px" }}>
    <b>${tip.title}</b>${tip.unit ? html` <span class="hint">${tip.unit}</span>` : ""}${tip.kind ? html` <span class="kind">${tip.kind}</span>` : ""}
    <div>${tip.body}</div>${tip.extra ? html`<div class="hint">${tip.extra}</div>` : ""}${tip.range ? html`<div class="hint">${tip.range}</div>` : ""}
  </div>`;
}

// ---------- the Explain panel ----------
export function ExplainPanel({ at, ctx, onOpen, onEdit, onBack, onClose, canBack, onHelp }) {
  if (!at) return null;
  const { t, k } = at;
  const term = TERMS[t] || { label: t, short: "No explanation yet." };
  const val = ctx ? valueOf(ctx, t, k) : undefined;
  const e = ctx && term.kind === "calculated" ? evaluate(ctx, t, k) : null;
  const users = ctx ? usedBy(ctx, t, k).slice(0, 30) : [];
  const where = k ? ` · ${keyText(ctx, t, k)}` : "";
  const place = ctx ? editPlace(ctx, t, k) : null;
  return html`<aside class="explain" aria-label="Explain">
    <div class="ex-head">
      ${canBack && html`<button class="small" onClick=${onBack} title="Back to the value you came from">Back</button>`}
      <span class="spacer"></span><button class="small" onClick=${onClose} aria-label="Close">Close</button>
    </div>
    <h3>${term.label}${term.sym && term.sym !== term.label ? html` <span class="sym">${term.sym}</span>` : ""}</h3>
    <p class="hint">${KIND[term.kind] || ""}${term.unit ? ` · ${term.unit}` : ""}${where}</p>
    <p class="ex-value">${val === undefined || val === "" ? html`<span class="hint">blank</span>` : withUnit(val, t)}</p>
    ${place && html`<p><button class="small primary" title=${place.why} onClick=${() => onEdit(place)}>Change it on tab ${place.tab}</button></p>`}
    <p>${marked(term.short)}</p>
    ${term.long && html`<p>${marked(term.long)}</p>`}
    ${term.range && html`<p class="hint">Typical: ${term.range[0]} to ${term.range[1]} ${term.unit}.${isNum(val) && (val < term.range[0] || val > term.range[1]) ? html` <span class="warn">This value is outside it: check it.</span>` : ""}</p>`}
    ${e && html`<div class="ex-fx">
      <h4>Formula</h4>
      <p class="mono">${e.text}</p>
      ${e.fn && html`<p class="mono">${formulaLine(e, t)}</p>`}
      ${e.note && html`<p class="hint">${marked(e.note)}</p>`}
      ${e.ref && html`<p class="hint">Reference: ${e.ref}</p>`}
      ${e.inputs.length > 0 && html`<h4>Inputs</h4><ul class="ex-in">${e.inputs.map(([sym, it, ik]) => html`<li>
        <button class="link" onClick=${() => onOpen(it, ik)}>${sym}</button> = ${withUnit(e.vals[sym], it)}
        <span class="hint"> · ${TERMS[it]?.label || it}, ${KIND[TERMS[it]?.kind] || ""}${ik && ik !== k ? `, ${keyText(ctx, it, ik)}` : ""}</span></li>`)}</ul>`}
    </div>`}
    ${term.kind === "calculated" && !e && html`<p class="hint">Not worked out yet: an input is missing.</p>`}
    ${users.length > 0 && html`<h4>Used by</h4><ul class="ex-in">${users.map(([ut, uk]) => html`<li><button class="link" onClick=${() => onOpen(ut, uk)}>${TERMS[ut]?.label || ut}</button>${uk && uk !== k ? html` <span class="hint">${keyText(ctx, ut, uk)}</span>` : ""}</li>`)}</ul>`}
    ${term.help && html`<p><button class="link" onClick=${() => onHelp(term.help)}>Help: more on this</button></p>`}
  </aside>`;
}
