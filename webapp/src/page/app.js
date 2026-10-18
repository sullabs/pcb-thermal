// The page: top bar, step tabs, the open tab, messages. Holds the three files and the undo history.

import { render } from "preact";
import { useState, useEffect, useMemo, useRef, useCallback } from "preact/hooks";
import * as XLSX from "xlsx";
import { html, Card } from "./ui.js";
import { TABS, tabStates, blockingProblems } from "./tabs.js";
import { PAGE_VERSION, isoLocal, newProject, setIn, projectForSave, historyInit, commit, undo, redo } from "./model.js";
import * as F from "./files.js";
import * as V from "./views.js";
import { calcProject } from "../calc/index.js";
import { valueOf } from "../calc/formulas.js";
import { Explain, Tip, ExplainPanel } from "./explain.js";
import { workbookToProject } from "../io/workbook_import.js";
import { validateProject, validateCad, validateResults } from "virtual:validators";
import EXAMPLE from "virtual:example";

const store = {
  get: k => { try { return localStorage.getItem("pcbt." + k); } catch { return null; } },
  set: (k, v) => { try { localStorage.setItem("pcbt." + k, v); } catch { /* storage blocked: not remembered */ } },
};
const now = () => isoLocal(new Date());
const today = () => now().slice(0, 10).replace(/-/g, "");
const hhmm = t => (t ? t.slice(11, 16) : "");

function schemaProblems(validate, name, data) {
  if (validate(data)) return [];
  return (validate.errors || []).slice(0, 8).map(e => `${name} ${e.instancePath || "/"}: ${e.message}${e.params?.additionalProperty ? " (" + e.params.additionalProperty + ")" : ""}`);
}

function App() {
  const [hist, setHist] = useState(historyInit(null));
  const [cad, setCad] = useState({});
  const [results, setResults] = useState({ runs: [] });
  const [dir, setDir] = useState(null);           // folder handle
  const [mode, setMode] = useState("none");       // none, folder, manual (opened by hand), empty (folder without project.json)
  const [savedRef, setSavedRef] = useState(null); // the project object as last saved or loaded
  const [savedAt, setSavedAt] = useState(null);
  const [tab, setTab] = useState(Number(store.get("tab")) || 0);
  const [theme, setTheme] = useState(store.get("theme") === "light" ? "light" : "dark");
  const [msgs, setMsgs] = useState([]);
  const [problems, setProblems] = useState([]);
  const [last, setLast] = useState(null);         // remembered folder handle
  const [backups, setBackups] = useState([]);
  const [showFx, setShowFx] = useState(store.get("fx") === "1");
  const [learn, setLearn] = useState(store.get("learn") !== "0");   // on until the user turns it off
  const [explain, setExplain] = useState([]);     // stack of { t, k }: the last is shown, Back pops
  const stamps = useRef({});
  const project = hist.present;
  const dirty = !!project && project !== savedRef;

  const say = useCallback((text, kind = "info") => {
    const id = Math.random();
    setMsgs(m => [...m.slice(-3), { id, text, kind }]);
    setTimeout(() => setMsgs(m => m.filter(x => x.id !== id)), kind === "error" ? 12000 : 6000);
  }, []);

  useEffect(() => { document.documentElement.dataset.theme = theme; store.set("theme", theme); }, [theme]);
  useEffect(() => { store.set("tab", String(tab)); }, [tab]);
  useEffect(() => { store.set("fx", showFx ? "1" : "0"); }, [showFx]);
  useEffect(() => { store.set("learn", learn ? "1" : "0"); }, [learn]);
  useEffect(() => { if (F.folderSupported) F.lastFolder().then(setLast); }, []);

  const calc = useMemo(() => {
    if (!project) return null;
    try { return calcProject(project, cad, results); } catch (e) { console.error(e); return { error: e.message }; }
  }, [project, cad, results]);
  const calcOk = calc && !calc.error ? calc : null;
  const canSave = mode === "folder" || mode === "manual";
  const example = mode === "example";
  const fctx = useMemo(() => (calcOk ? { project, cad, results, calc: calcOk } : null), [project, cad, results, calcOk]);
  const openExplain = useCallback((t, k = "") => setExplain(s => [...s.slice(-30), { t, k }]), []);
  const explainCtx = useMemo(() => ({ ctx: fctx, open: openExplain, showFx }), [fctx, openExplain, showFx]);
  const states = useMemo(() => tabStates({ project, cad, results, calc: calcOk, canSave: canSave || example }), [project, cad, results, calcOk, canSave, example]);

  // ---------- edits ----------
  const edit = useCallback((label, path, value, key) => { if (live.current.mode !== "example") setHist(h => commit(h, setIn(h.present, path, value), label, key)); }, []);
  const editFn = useCallback((label, fn, key) => { if (live.current.mode !== "example") setHist(h => commit(h, fn(h.present), label, key)); }, []);
  const doUndo = () => setHist(h => { const n = undo(h); if (n.did) say(n.did); return n; });
  const doRedo = () => setHist(h => { const n = redo(h); if (n.did) say(n.did); return n; });

  // ---------- open ----------
  const loadFiles = (pj, cj, rj, origin) => {
    const probs = [
      ...(pj ? schemaProblems(validateProject, "project.json", pj) : []),
      ...(cj ? schemaProblems(validateCad, "cad.json", cj) : []),
      ...(rj ? schemaProblems(validateResults, "results.json", rj) : []),
    ];
    setProblems(probs);
    setHist(historyInit(pj));
    setSavedRef(pj);
    setSavedAt(pj?.savedAt || null);
    setCad(cj || {});
    setResults(rj || { runs: [] });
    if (pj) say(`Opened ${pj.board?.name || "project"} from ${origin}: ${pj.parts?.length ?? 0} parts${cj?.import ? `, ${cj.import.bodies?.length ?? 0} bodies` : ""}${rj?.runs?.length ? `, ${rj.runs.length} runs` : ""}.`);
    if (probs.length) say("Some files do not match the format. See tab 0.", "error");
  };

  const openFolder = async (handle, ask = true) => {
    if (dirty && !confirm("You have unsaved changes. Open another folder anyway?")) return;
    try {
      if (!(await F.allowed(handle, ask))) { say("The browser did not allow access to the folder.", "error"); return; }
      const [pj, cj, rj] = await Promise.all(["project.json", "cad.json", "results.json"].map(n => F.readJson(handle, n)));
      stamps.current = { project: pj?.stamp || 0, cad: cj?.stamp || 0, results: rj?.stamp || 0 };
      setDir(handle);
      setLast(handle);
      setBackups(await F.listBackups(handle));
      if (!pj) {
        setMode("empty");
        loadFiles(null, cj?.data, rj?.data, handle.name);
        say(`${handle.name} has no project.json yet. Start a new project or make one from an old workbook.`);
        return;
      }
      setMode("folder");
      loadFiles(pj.data, cj?.data, rj?.data, "folder " + handle.name);
    } catch (e) { say(e.message, "error"); }
  };
  const pick = async () => {
    try { await openFolder(await F.pickFolder()); } catch (e) { if (e.name !== "AbortError") say(e.message, "error"); }
  };

  const openByHand = async e => {
    const files = [...e.currentTarget.files];
    const got = {};
    try {
      for (const f of files) {
        const k = f.name.replace(/\.json$/i, "").replace(/ \(\d+\)$/, "");
        if (["project", "cad", "results"].includes(k)) got[k] = JSON.parse(await f.text());
      }
    } catch (err) { say("A file is not valid JSON: " + err.message, "error"); return; }
    if (!got.project) { say("Pick project.json (and cad.json and results.json if the folder has them).", "error"); return; }
    setDir(null);
    setMode("manual");
    loadFiles(got.project, got.cad, got.results, "files opened by hand");
  };

  // The Raspberry Pi 4 study, built into the page, read only (design 5.3).
  const openExample = () => {
    if (dirty && !confirm("You have unsaved changes. Open the example anyway?")) return;
    setDir(null);
    setMode("example");
    loadFiles(EXAMPLE.project, EXAMPLE.cad, EXAMPLE.results, "the built-in example");
    say("This is the example study, read only. Look around; to work on your own board, choose its folder on tab 0.");
  };

  const startNew = async () => {
    const p = newProject(dir.name, now());
    const steps = [];
    for await (const [name, h] of dir.entries()) if (h.kind === "file" && /\.(step|stp)$/i.test(name)) steps.push(name);
    if (steps.length === 1) p.board.stepFile = steps[0];
    setMode("folder");
    setHist(historyInit(p));
    setSavedRef(null);
    await save(false, p);
  };

  const fromWorkbook = async e => {
    const f = e.currentTarget.files[0];
    if (!f) return;
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellFormula: false, cellNF: false, cellDates: false });
      const base = f.name.replace(/\.xlsx$/i, "");
      const out = workbookToProject(wb, XLSX, {
        fileName: f.name, now: now(), savedBy: "page " + PAGE_VERSION,
        reports: { buildBoard: await F.readText(dir, `CAD/${base}_BuildBoard_report.txt`), buildSpreader: await F.readText(dir, `CAD/${base}_BuildSpreader_report.txt`) },
      });
      const wrote = [];
      if (await F.writeIfMissing(dir, "cad.json", out.cad)) wrote.push("cad.json");
      if (await F.writeIfMissing(dir, "results.json", out.results)) wrote.push("results.json");
      const [cj, rj] = await Promise.all(["cad.json", "results.json"].map(n => F.readJson(dir, n)));
      stamps.current = { ...stamps.current, cad: cj?.stamp || 0, results: rj?.stamp || 0 };
      setMode("folder");
      loadFiles(out.project, cj?.data, rj?.data, f.name);
      setSavedRef(null);
      await save(false, out.project, cj?.data, rj?.data);
      say(`Made the project from ${f.name}. Wrote ${["project.json", ...wrote].join(", ")}.${out.notes.length ? " Notes: " + out.notes.join(" ") : ""}`);
    } catch (err) { console.error(err); say("Could not read the workbook: " + err.message, "error"); }
  };

  // ---------- save ----------
  const live = useRef({});
  live.current = { project, cad, results, dir, mode, dirty };
  const save = async (auto, pIn, cIn, rIn) => {
    const L = live.current;
    const p = pIn || L.project, d = L.dir;
    if (!p) return;
    const t = now();
    const out = projectForSave(p, cIn || L.cad, rIn || L.results, t);
    const probs = schemaProblems(validateProject, "project.json", out);
    const refs = out.parts.map(x => x.refDes);
    const dup = [...new Set(refs.filter((x, i) => refs.indexOf(x) !== i))];
    if (dup.length) probs.push(`RefDes used by two parts: ${dup.join(", ")}. Give each part its own RefDes.`);
    if (probs.length) {
      setProblems(probs);
      if (!auto) say("Not saved: " + probs[0], "error");
      return;
    }
    setProblems([]);
    try {
      if ((pIn ? "folder" : L.mode) === "folder" && d) {
        const disk = await F.stamp(d, "project.json");
        if (disk && stamps.current.project && disk !== stamps.current.project) {
          if (auto) { say("Not saved on its own: project.json was changed outside this page. Press Save to replace it.", "error"); return; }
          if (!confirm("project.json was changed outside this page since it was opened. Replace it with this page's version?")) return;
        }
        const r = await F.saveProject(d, out, today());
        stamps.current.project = r.stamp;
        if (r.backup) { say(`Backup of the old file: ${r.backup}`); setBackups(await F.listBackups(d)); }
        if (!auto) say(`Saved project.json in ${d.name}.`);
      } else {
        F.download("project.json", JSON.stringify(out, null, 2) + "\n");
        say("project.json downloaded. Put it in the project folder, over the old one.");
      }
      setSavedRef(p);
      setSavedAt(t);
    } catch (e) { say("Not saved: " + e.message, "error"); }
  };

  // Save on its own every 30 s while there are unsaved edits (folder mode only).
  useEffect(() => {
    const id = setInterval(() => { const L = live.current; if (L.dirty && L.mode === "folder") save(true); }, 30000);
    return () => clearInterval(id);
  }, []);

  // Reload cad.json and results.json when a macro has written them.
  useEffect(() => {
    if (mode !== "folder" || !dir) return;
    const id = setInterval(async () => {
      for (const [name, key, set] of [["cad.json", "cad", setCad], ["results.json", "results", setResults]]) {
        try {
          const st = await F.stamp(dir, name);
          if (st === stamps.current[key]) continue;
          const r = await F.readJson(dir, name);
          stamps.current[key] = st;
          if (!r) continue;
          set(r.data);
          say(key === "cad" ? `cad.json reloaded (${r.data.savedBy || "macro"}): ${r.data.import?.bodies?.length ?? 0} bodies${r.data.build ? ", build " + (r.data.build.date || "") : ""}.`
            : `results.json reloaded: ${r.data.runs?.length ?? 0} runs.`);
        } catch { /* half written or locked: try again in 2 s */ }
      }
      try {
        const st = await F.stamp(dir, "project.json");
        if (st && stamps.current.project && st !== stamps.current.project && !stamps.current.warned) {
          stamps.current.warned = true;
          say("project.json was changed outside this page. Saving here will ask before replacing it.", "error");
        }
      } catch { /* ignore */ }
    }, 2000);
    return () => clearInterval(id);
  }, [mode, dir]);

  // Keys and leaving the page.
  useEffect(() => {
    const onKey = e => {
      if (!(e.ctrlKey || e.metaKey)) return;
      const k = e.key.toLowerCase();
      if (k === "s") { e.preventDefault(); save(false); }
      else if (k === "z" && !e.shiftKey) { e.preventDefault(); doUndo(); }
      else if (k === "y" || (k === "z" && e.shiftKey)) { e.preventDefault(); doRedo(); }
    };
    const onLeave = e => { if (live.current.dirty) { e.preventDefault(); e.returnValue = ""; } };
    addEventListener("keydown", onKey);
    addEventListener("beforeunload", onLeave);
    return () => { removeEventListener("keydown", onKey); removeEventListener("beforeunload", onLeave); };
  }, []);

  const openBackup = async name => {
    try {
      const r = await F.readBackup(dir, name);
      setHist(h => commit(h, r.data, "Open backup " + name));
      say(`Opened ${name}. Press Save to keep it as project.json (Ctrl+Z goes back).`);
    } catch (e) { say(e.message, "error"); }
  };

  // Development hook for tools/test_w3 checks in a browser (a folder handle from the private origin storage).
  window.PCBT_DEV = { openFolder: h => openFolder(h, false), state: () => ({ project, cad, results, mode, dirty, tab, states, past: hist.past.map(x => x.label), future: hist.future.map(x => x.label), savedIsPresent: savedRef === project }) };

  // ---------- view ----------
  const st = states[tab] || states[0];
  const folderCard = html`<div class="folder">
    ${F.folderSupported ? html`
      <button class="primary" onClick=${pick}>${dir ? "Choose another folder" : "Choose the project folder"}</button>
      ${last && (!dir || last.name !== dir.name) && html` <button onClick=${() => openFolder(last)}>Open ${last.name} again</button>`}
      ${dir && html`<p>Folder: <b>${dir.name}</b></p>`}
      ${mode === "empty" && html`<div class="empty-folder">
        <p>This folder has no <code>project.json</code> yet.</p>
        <button class="primary" onClick=${startNew}>Start a new project here</button>
        <label class="button">Make it from an old workbook (.xlsx)<input type="file" accept=".xlsx" hidden onChange=${fromWorkbook} /></label>
      </div>`}`
    : html`<p class="warn">This browser cannot save into a folder. Use Chrome or Edge for that. Here you can open the files by hand, and Save downloads <code>project.json</code>.</p>`}
    <p class="hint">${F.folderSupported ? "Or open the files by hand:" : ""} <label class="button small">Open project.json, cad.json, results.json<input type="file" accept=".json" multiple hidden onChange=${openByHand} /></label></p>
    <p class="hint">New to this? <button class="small" onClick=${openExample}>Open the example study (Raspberry Pi 4, read only)</button></p>
    ${problems.length > 0 && html`<div class="problems"><p class="warn">Format problems (Save is refused until they are fixed):</p><ul>${problems.map(x => html`<li>${x}</li>`)}</ul></div>`}
    ${calc?.error && html`<p class="warn">The calculation stopped: ${calc.error}</p>`}
  </div>`;
  const backupCard = mode === "folder" && backups.length > 0 && html`<p>Open a backup: ${backups.slice(0, 7).map(b => html` <button class="small" onClick=${() => openBackup(b)}>${b.slice(8, 16)}</button>`)}</p>`;

  const ctx = { project, cad, results, calc: calcOk, states, value: (t, k) => (fctx ? valueOf(fctx, t, k) : undefined), folder: example ? "built-in example (read only)" : dir?.name, edit, editFn, go: setTab, locked: st.locked || example, now, learn, openExample,
    blocking: calcOk ? blockingProblems(project, calcOk) : [], folderCard, backupCard };
  const View = { project: V.ProjectTab, parts: V.PartsTab, import: V.ImportTab, link: V.LinkTab, triage: V.TriageTab, bare: V.BareTab,
    build: V.BuildTab, spreader: V.SpreaderTab, sim: V.SimTab, results: V.ResultsTab, changes: V.ChangesTab, help: V.HelpTab }[st.id];
  const needsProject = !["project", "help"].includes(st.id);

  return html`
    <header class="top">
      <span class="brand">PCB Thermal</span>
      <span class="name">${project?.board?.name || "No project open"}</span>
      <span class=${`save ${dirty ? "dirty" : ""}`}>${!project ? "" : dirty ? "Unsaved changes" : savedAt ? "Saved " + hhmm(savedAt) : ""}</span>
      <span class="spacer"></span>
      <button onClick=${() => save(false)} disabled=${!project || !canSave} title="Save project.json (Ctrl+S)">Save</button>
      <button onClick=${doUndo} disabled=${!hist.past.length} title=${hist.past.length ? "Undo: " + hist.past[hist.past.length - 1].label + " (Ctrl+Z)" : "Nothing to undo"}>Undo</button>
      <button onClick=${doRedo} disabled=${!hist.future.length} title=${hist.future.length ? "Redo: " + hist.future[0].label + " (Ctrl+Y)" : "Nothing to redo"}>Redo</button>
      <button aria-pressed=${learn} class=${learn ? "on" : ""} onClick=${() => setLearn(!learn)} title="Show what each step is for, what to look at, and a plain reading of the results">Learn</button>
      <button aria-pressed=${showFx} class=${showFx ? "on" : ""} onClick=${() => setShowFx(!showFx)} title="Write the formula, with your numbers, under every calculated value">Show formulas</button>
      <button onClick=${() => setTheme(theme === "dark" ? "light" : "dark")} title="Switch between dark and light">${theme === "dark" ? "Light" : "Dark"}</button>
      <button onClick=${() => setTab(11)}>Help</button>
    </header>
    <nav class="tabs" role="tablist">
      ${states.map(s => html`<button role="tab" aria-selected=${s.n === tab} class=${`tab ${s.status} ${s.n === tab ? "on" : ""}`}
        title=${s.locked ? "Locked: " + s.reason : s.note || ""} onClick=${() => setTab(s.n)}>
        <span class="n">${s.n}</span> ${s.title}${s.locked ? html` <span class="lock" aria-label="locked">🔒</span>` : ""}${s.badge ? html` <span class="badge">${s.badge}</span>` : ""}
      </button>`)}
    </nav>
    <${Explain.Provider} value=${explainCtx}>
    <div class=${`layout ${explain.length ? "with-panel" : ""}`}>
    <main>
      <h2>${st.n} ${st.title}</h2>
      ${st.locked && html`<div class="locked-banner" role="note"><b>This step is locked.</b> ${st.reason}
        ${st.fixTab !== undefined && html` <button onClick=${() => setTab(st.fixTab)}>Go to tab ${st.fixTab}</button>`}</div>`}
      ${example && html`<div class="example-banner" role="note"><b>Example study, read only.</b> Edits are off and nothing is saved. <button class="small" onClick=${() => setTab(0)}>Choose your own folder</button></div>`}
      ${learn && html`<${V.LearnBox} id=${st.id} />`}
      ${needsProject && !project ? html`<p class="hint">Open a project on tab 0 first.</p>`
        : needsProject && !calcOk ? html`<${Card} title="The calculation stopped"><p class="warn">${calc?.error}</p><//>`
        : html`<fieldset class="content" disabled=${st.locked || example}>${View(ctx)}</fieldset>`}
    </main>
    <${ExplainPanel} at=${explain[explain.length - 1]} ctx=${fctx} onOpen=${openExplain} canBack=${explain.length > 1}
      onBack=${() => setExplain(s => s.slice(0, -1))} onClose=${() => setExplain([])} onHelp=${() => setTab(11)} />
    </div>
    <${Tip} />
    <//>
    <div class="msgs" aria-live="polite">${msgs.map(m => html`<div class=${`msg ${m.kind}`} key=${m.id}>${m.text}</div>`)}</div>
    <footer>page ${PAGE_VERSION} · ${mode === "folder" ? "saves into the folder" : mode === "manual" ? "files opened by hand: Save downloads project.json" : example ? "example study, read only" : "no folder"}</footer>`;
}

export function start(el) {
  el.textContent = "";
  render(html`<${App} />`, el);
}
