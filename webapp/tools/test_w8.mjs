// W8 test: the Help page and the glossary.
//   node tools/test_w8.mjs
// 1. Every section of the design (section 7) is there; every tab and every SolidWorks button has its help.
// 2. Every catalogue term, abbreviation and status word is in the glossary, with the same text as its hover.
// 3. Every help link of a term opens a section; the page renders with the search.
import { h } from "preact";
import { render } from "preact-render-to-string";
import { HELP, TAB_HELP, MACRO_HELP, TROUBLE } from "../src/terms/help.js";
import { LEARN } from "../src/terms/learn.js";
import { TERMS, GLOSSARY, STATUS } from "../src/terms/catalogue.js";
import { HelpTab, glossaryEntries, helpSection } from "../src/page/help.js";
import { TABS } from "../src/page/tabs.js";
import { MACROS } from "../src/page/swcard.js";
import { Explain } from "../src/page/explain.js";

let fails = 0, n = 0;
const check = (ok, what, extra) => { n++; if (!ok) { fails++; console.log("FAIL " + what + (extra !== undefined ? ": " + JSON.stringify(extra).slice(0, 300) : "")); } };

const want = ["start", "heat", "datasheet", "methods", "network", "spreaders", "simulation", "tabs", "solidworks", "trouble", "glossary"];
check(want.every(id => HELP.some(s => s.id === id)), "every help section of the design is there", want.filter(id => !HELP.some(s => s.id === id)));
check(HELP.filter(s => !["tabs", "trouble", "glossary"].includes(s.id)).every(s => s.body.length >= 1), "every text section has text");
for (const t of TABS.filter(t => t.id !== "help")) check(LEARN[t.id] && TAB_HELP[t.id]?.you && TAB_HELP[t.id]?.out && TAB_HELP[t.id]?.mistakes?.length, `tab ${t.n} ${t.title} has its help`);
for (const id of Object.keys(MACROS)) check(MACRO_HELP[id]?.changes && MACRO_HELP[id]?.check, `${id} has its help`);
check(TROUBLE.length >= 8, "troubleshooting list");
const text = JSON.stringify([HELP, TAB_HELP, MACRO_HELP, TROUBLE]);
check(!/[—–]/.test(text), "help has no long dashes");

const g = glossaryEntries();
const keys = new Set(g.map(e => e.key));
const missing = [...Object.keys(TERMS).map(k => "t:" + k), ...Object.keys(GLOSSARY).map(k => "g:" + k), ...Object.keys(STATUS).map(k => "s:" + k)].filter(k => !keys.has(k));
check(!missing.length, `every term, abbreviation and status word is in the glossary (${g.length})`, missing.slice(0, 10));
check(g.filter(e => e.term).every(e => e.short === TERMS[e.term].short), "glossary text is the hover text");

for (const [id, t] of Object.entries(TERMS)) if (t.help) check(helpSection(t.help) !== "start" || t.help === "start", `help link of ${id} (${t.help}) opens a section`);
check(helpSection("spreading") === "network" && helpSection("heat") === "heat", "help aliases");

const out = render(h(Explain.Provider, { value: { ctx: null, open() {}, showFx: false } }, h(() => HelpTab({ openExample() {}, go() {} }))));
for (const id of want) check(out.includes(`id="help-${id}"`), `help section ${id} renders`);
for (const t of TABS.filter(t => t.id !== "help")) check(out.includes(`id="help-tab-${t.id}"`), `tab ${t.n} section renders`);
check(out.includes('data-term="help.search"') && out.includes(`${g.length} of ${g.length} entries`), "glossary search renders with every entry");
check(out.includes('id="gloss-sim.kz"'), "glossary entries can be linked by term");

console.log(`W8: ${n - fails} of ${n} checks passed.`);
if (fails) process.exit(1);
