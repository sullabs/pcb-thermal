// 11 Help (design section 7): the sections of terms/help.js, one section per tab (its Learn texts plus what you do,
// what you get and the usual mistakes), the SolidWorks steps, troubleshooting, and a searchable glossary made from the
// term catalogue, so a term's hover text and its glossary entry are the same text.

import { useState, useEffect, useMemo } from "preact/hooks";
import { html } from "./ui.js";
import { marked, KIND } from "./explain.js";
import { HELP, TAB_HELP, MACRO_HELP, TROUBLE } from "../terms/help.js";
import { LEARN } from "../terms/learn.js";
import { TERMS, GLOSSARY, STATUS } from "../terms/catalogue.js";
import { TABS } from "./tabs.js";
import { MACROS } from "./swcard.js";

// Section id for a term's help field ("spreading" opens The bare board network).
export const helpSection = id => HELP.find(s => s.id === id || (s.aliases || []).includes(id))?.id || "start";

// The glossary: every catalogue term, every abbreviation and every status word, sorted by name.
export function glossaryEntries() {
  const out = Object.values(TERMS).map(t => ({ key: "t:" + t.id, name: t.label, sym: t.sym && t.sym !== t.label ? t.sym : "", unit: t.unit, kind: KIND[t.kind], short: t.short, long: t.long, term: t.id }));
  for (const [k, v] of Object.entries(GLOSSARY)) out.push({ key: "g:" + k, name: k, short: v, kind: "abbreviation" });
  for (const [k, v] of Object.entries(STATUS)) out.push({ key: "s:" + k, name: k, short: v, kind: "status word" });
  return out.sort((a, b) => a.name.localeCompare(b.name, "en", { sensitivity: "base" }));
}

const para = (p, i) => (Array.isArray(p) ? html`<ol key=${i}>${p.map(x => html`<li>${marked(x)}</li>`)}</ol>` : html`<p key=${i}>${marked(p)}</p>`);

export function HelpTab(c = {}) {
  return html`<${HelpView} c=${c} />`;
}

function HelpView({ c }) {
  const [q, setQ] = useState("");
  const all = useMemo(glossaryEntries, []);
  const words = q.trim().toLowerCase().split(/\s+/).filter(Boolean);
  const shown = words.length ? all.filter(e => words.every(w => [e.name, e.sym, e.short, e.long, e.unit].join(" ").toLowerCase().includes(w))) : all;
  useEffect(() => {
    if (!c.helpAt) return;
    const el = document.getElementById("help-" + c.helpAt);
    if (el) el.scrollIntoView({ block: "start" });
  }, [c.helpAt]);
  const go = id => { const el = document.getElementById("help-" + id); if (el) el.scrollIntoView({ block: "start", behavior: "smooth" }); };
  return html`<div class="help">
    <nav class="help-toc" aria-label="Help sections">${HELP.map(s => html`<button class="link" onClick=${() => go(s.id)}>${s.title}</button>`)}</nav>
    ${HELP.map(s => html`<section class="card help-sec" id=${"help-" + s.id} key=${s.id}>
      <h3>${s.title}</h3>
      ${s.body.map(para)}
      ${s.id === "start" && c.openExample && html`<p><button class="small" onClick=${c.openExample}>Open the example study</button></p>`}
      ${s.id === "tabs" && TABS.filter(t => t.id !== "help").map(t => {
        const l = LEARN[t.id] || {}, h = TAB_HELP[t.id] || {};
        return html`<div class="help-tab" id=${"help-tab-" + t.id}><h4>Tab ${t.n}: ${t.title}${c.go ? html` <button class="link small" onClick=${() => c.go(t.n)}>open it</button>` : ""}</h4>
          ${l.purpose && html`<p><b>For:</b> ${marked(l.purpose)}</p>`}
          ${h.you && html`<p><b>You:</b> ${marked(h.you)}</p>`}
          ${h.out && html`<p><b>You get:</b> ${marked(h.out)}</p>`}
          ${l.look && html`<p><b>Look at:</b></p><ol>${l.look.map(x => html`<li>${marked(x)}</li>`)}</ol>`}
          ${h.mistakes && html`<p><b>Usual mistakes:</b></p><ol>${h.mistakes.map(x => html`<li>${marked(x)}</li>`)}</ol>`}
        </div>`;
      })}
      ${s.id === "solidworks" && Object.entries(MACROS).map(([id, m]) => html`<div class="help-tab" id=${"help-sw-" + id}>
        <h4>${m.n} ${m.name}</h4>
        <p>${marked(m.does)}</p>
        <p><b>What it changes:</b> ${marked(MACRO_HELP[id].changes)}</p>
        <p><b>What to check after:</b> ${marked(MACRO_HELP[id].check)}</p>
      </div>`)}
      ${s.id === "trouble" && html`<dl class="trouble">${TROUBLE.map(([a, b]) => html`<dt>${marked(a)}</dt><dd>${marked(b)}</dd>`)}</dl>`}
      ${s.id === "glossary" && html`
        <p class="help-search"><span class="field" data-term="help.search" data-key=""><input type="text" aria-label="Search the glossary" placeholder="Search: Rjc, spreading, TIM..." value=${q} onInput=${e => setQ(e.currentTarget.value)} style=${{ width: "22em" }} /></span>
          <span class="hint"> ${shown.length} of ${all.length} entries</span></p>
        <dl class="gloss">${shown.map(e => html`<div class="g-entry" key=${e.key} id=${e.term ? "gloss-" + e.term : undefined}>
          <dt><b>${e.name}</b>${e.sym ? html` <span class="sym">${e.sym}</span>` : ""}${e.unit ? html` <span class="hint">${e.unit}</span>` : ""}${e.kind ? html` <span class="kind">${e.kind}</span>` : ""}</dt>
          <dd><p>${marked(e.short)}</p>${e.long && html`<p class="hint">${marked(e.long)}</p>`}</dd></div>`)}</dl>`}
    </section>`)}
  </div>`;
}
