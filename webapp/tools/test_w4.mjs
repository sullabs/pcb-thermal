// W4 test: term catalogue, formulas and hover texts.
//   node tools/test_w4.mjs
// 1. Every formula, worked out from its inputs, equals the value of the calculation module, for every instance
//    in both RPi4 examples (RaspberryPi4 and RPi4_fresh, which has a bottom plate).
// 2. Every catalogue entry is complete; every calculated term has a formula; every term has a value getter.
// 3. Every tab, rendered for both examples with Show formulas and Learn mode on: every value, label and input carries a term id
//    from the catalogue, no number stands without one, and each calculated value with a number shows a formula.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { h } from "preact";
import { render } from "preact-render-to-string";
import { calcProject } from "../src/calc/index.js";
import { VALUES, FORMULAS, instances, evaluate, valueOf } from "../src/calc/formulas.js";
import { TERMS, GLOSSARY, STATUS } from "../src/terms/catalogue.js";
import { Explain, formulaLine } from "../src/page/explain.js";
import { tabStates, blockingProblems } from "../src/page/tabs.js";
import * as Views from "../src/page/views.js";
import { convertFile } from "./convert_workbook.mjs";

const here = path.dirname(fileURLToPath(import.meta.url));
const ex = path.join(here, "..", "..", "example", "RaspberryPi4");
const read = f => JSON.parse(fs.readFileSync(path.join(ex, f), "utf8"));
let fails = 0, n = 0;
const check = (ok, what, extra) => { n++; if (!ok) { fails++; console.log("FAIL " + what + (extra !== undefined ? ": " + JSON.stringify(extra).slice(0, 300) : "")); } };

const examples = {
  RaspberryPi4: { project: read("project.json"), cad: read("cad.json"), results: read("results.json") },
  RPi4_fresh: convertFile(path.join(ex, "RPi4_fresh.xlsx"), { now: "2026-10-18T12:00:00" }),
};

// ---------- 1. formulas equal the calculation ----------
for (const [name, f] of Object.entries(examples)) {
  const ctx = { ...f, calc: calcProject(f.project, f.cad, f.results) };
  let checked = 0, bad = 0;
  for (const [t, k] of instances(ctx)) {
    const e = FORMULAS[t] && evaluate(ctx, t, k);
    if (!e?.fn || !Number.isFinite(e.value)) continue;
    checked++;
    if (!(Math.abs(e.result - e.value) <= 1e-9 * Math.max(1, Math.abs(e.value)))) { bad++; check(false, `${name} ${t} ${k}: formula gives ${e.result}, calc ${e.value}`); }
  }
  check(checked > 300 && bad === 0, `${name}: ${checked} formulas give the calculated value`, { checked, bad });
  f.ctx = ctx;
}
{
  const ctx = examples.RaspberryPi4.ctx;
  const line = formulaLine(evaluate(ctx, "part.tjEst", "U1"), "part.tjEst");
  check(line === "Tj est = Ta + Pd × Rja = 40 + 3.168 × 14 = 84.35 C", "formula line with numbers", line);
  const sp = formulaLine(evaluate(ctx, "bbp.rSpread", "U1"), "bbp.rSpread");
  check(/^R spread = .* = \(0\.3168 × 0\.0599 \/ √π \+ \(1 - 0\.3168\) × 3\.413 \/ √π\) \/ \(√π × \d+(\.\d+)? × √\(225 \/ π\) \/ 1000\) = 4\.259 C\/W$/.test(sp), "spreading formula line", sp);
}

// ---------- 2. catalogue ----------
const KINDS = new Set(["input", "calculated", "cad", "run", "constant"]);
for (const [id, t] of Object.entries(TERMS)) {
  check(t.label && t.short && KINDS.has(t.kind) && typeof t.unit === "string", `term ${id} has label, short, kind and unit`, t);
  check(!/[—–]/.test(JSON.stringify(t)), `term ${id} has no long dashes`);
  if (t.kind === "calculated") check(!!FORMULAS[id], `calculated term ${id} has a formula`);
  if (!id.startsWith("file.")) check(!!VALUES[id], `term ${id} has a value getter`);
  if (t.range) check(t.range[0] < t.range[1], `term ${id} range is low to high`);
}
for (const id of Object.keys(VALUES)) check(!!TERMS[id], `value ${id} is in the catalogue`);
for (const id of Object.keys(FORMULAS)) check(!!TERMS[id], `formula ${id} is in the catalogue`);
for (const [id, f] of Object.entries(examples)) for (const [t, k] of instances(f.ctx)) {
  const e = FORMULAS[t] && evaluate(f.ctx, t, k);
  if (e) for (const [sym, it] of e.inputs) if (!TERMS[it]) check(false, `${t} input ${sym} (${it}) is in the catalogue`);
}
check(Object.keys(GLOSSARY).length >= 20 && Object.keys(STATUS).length >= 10, "glossary and status words");

// ---------- 3. every tab ----------
function parse(htmlText) {
  // Small HTML walker: yields text nodes with the stack of open elements.
  const out = [], stack = [];
  const re = /<(\/?)([a-zA-Z0-9]+)((?:\s+[^\s=>/]+(?:="[^"]*")?)*)\s*(\/?)>|([^<]+)/g;
  const VOID = new Set(["input", "br", "img", "meta", "hr", "rect", "path"]);
  let m;
  while ((m = re.exec(htmlText))) {
    if (m[5] !== undefined) { out.push({ text: m[5], stack: [...stack] }); continue; }
    const [, close, tag, attrs, self] = m;
    if (close) { const i = stack.map(s => s.tag).lastIndexOf(tag); if (i >= 0) stack.length = i; continue; }
    const a = {};
    for (const x of attrs.matchAll(/([^\s=]+)(?:="([^"]*)")?/g)) a[x[1]] = x[2] ?? "";
    const el = { tag, a };
    out.push({ el, stack: [...stack] });
    if (!self && !VOID.has(tag)) stack.push(el);
  }
  return out;
}
const unesc = s => s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'");

for (const [name, f] of Object.entries(examples)) {
  const { project, cad, results, ctx } = f;
  const states = tabStates({ project, cad, results, calc: ctx.calc, canSave: true });
  const c = { project, cad, results, calc: ctx.calc, states, folder: name, edit() {}, editFn() {}, go() {}, locked: false, learn: true, now: () => "2026-10-18T12:00:00",
    blocking: blockingProblems(project, ctx.calc), folderCard: null, backupCard: null, value: (t, k) => valueOf(ctx, t, k) };
  const tabs = { ProjectTab: 0, PartsTab: 1, ImportTab: 2, LinkTab: 3, TriageTab: 4, BareTab: 5, BuildTab: 6, SpreaderTab: 7, SimTab: 8, ResultsTab: 9, ChangesTab: 10, HelpTab: 11 };
  let values = 0, fx = 0;
  for (const tab of Object.keys(tabs)) {
    const out = render(h(Explain.Provider, { value: { ctx, open() {}, showFx: true } }, h(() => Views[tab](c))));
    const nodes = parse(out);
    const where = `${name} ${tab}`;
    for (const x of nodes) {
      if (x.el) {
        const t = x.el.a["data-term"];
        if (t !== undefined) { check(!!TERMS[t], `${where}: term ${t} is in the catalogue`); }
        if (["input", "select", "textarea"].includes(x.el.tag) && x.el.a.type !== "file") {
          check(t !== undefined || x.stack.some(s => s.a["data-term"]), `${where}: <${x.el.tag} ${x.el.a["aria-label"] || ""}> has a term`);
        }
        if (x.el.a.class?.split(" ").includes("v")) {
          values++;
          const term = TERMS[t];
          const val = valueOf(ctx, t, x.el.a["data-key"] || "");
          if (term?.kind === "calculated" && !x.el.a["data-nofx"] && Number.isFinite(val) && evaluate(ctx, t, x.el.a["data-key"] || "")?.fn) {
            const has = nodes.some(y => y.el?.a.class === "fx" && y.stack.includes(x.el));
            if (has) fx++; else check(false, `${where}: ${t} ${x.el.a["data-key"]} shows its formula`);
          }
        }
        continue;
      }
      const text = unesc(x.text).trim();
      if (!/\d/.test(text)) continue;
      const inside = tag => x.stack.some(s => s.tag === tag);
      const termed = x.stack.some(s => s.a["data-term"] !== undefined || s.a["data-status"] !== undefined || s.a["data-gloss"] !== undefined);
      // Numbers that are not values: tab numbers in buttons and headings, hints, code, notes, option texts, svg titles.
      const plain = ["button", "h2", "h3", "h4", "code", "option", "title", "textarea", "label", "p", "ol"].some(inside)
        && !inside("td") && !inside("dd");
      const hint = x.stack.some(s => /\b(hint|warn|err|toolbar|legend)\b/.test(s.a.class || ""));
      if (!termed && !plain && !hint) check(false, `${where}: number without a term: '${text.slice(0, 60)}'`, x.stack.map(s => s.tag + (s.a.class ? "." + s.a.class : "")).join(">"));
    }
  }
  console.log(`  ${name}: ${values} values with terms, ${fx} formulas shown`);
  check(values > 400 && fx > 150, `${name}: ${values} values with terms, ${fx} formulas shown`, { values, fx });
}

console.log(`W4: ${n - fails} of ${n} checks passed.`);
if (fails) process.exit(1);
