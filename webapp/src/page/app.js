// The page: top bar, step tabs, the open tab, messages. Holds the three files and the undo history.

import { render } from "preact";
import { useState, useEffect, useMemo, useRef, useCallback } from "preact/hooks";
import * as XLSX from "xlsx";
import { html, Card } from "./ui.js";
import { TABS, tabStates, blockingProblems } from "./tabs.js";
import { PAGE_VERSION, isoLocal, newProject, emptyCad, emptyResults, setIn, projectForSave, historyInit, commit, undo, redo } from "./model.js";
import * as F from "./files.js";
import * as V from "./views.js";
import { calcProject } from "../calc/index.js";
import { valueOf } from "../calc/formulas.js";
import { Explain, Tip, ExplainPanel } from "./explain.js";
import { findField } from "./where.js";
import { workbookToProject } from "../io/workbook_import.js";
import { validateProject, validateCad, validateResults } from "virtual:validators";
import EXAMPLE from "virtual:example";
import ICONS from "virtual:icons";
import { MACROS, marker } from "./swcard.js";
import { swFolder } from "./swfiles.js";
import { helpSection } from "./help.js";
import { Report } from "./report.js";
import { exportWorkbook } from "../io/export_xlsx.js";
import { makeZip } from "../io/zip.js";
import { safeFile } from "./swfiles.js";

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
  const [mode, setMode] = useState("none");       // none, folder, manual (opened by hand), empty (folder without a project file), choose (several)
  const [pre, setPre] = useState("");             // file prefix of the open project with its "_" ("AB_"), "" for project.json
  const [found, setFound] = useState([]);         // prefixes of the project files in the folder (choose mode)
  const [newPre, setNewPre] = useState(store.get("initials") || "");   // new project form
  const [newName, setNewName] = useState("");
  const [savedRef, setSavedRef] = useState(null); // the project object as last saved or loaded
  const [savedAt, setSavedAt] = useState(null);
  const [tab, setTab] = useState(Number(store.get("tab")) || 0);
  const [theme, setTheme] = useState(store.get("theme") === "light" ? "light" : "dark");
  const [msgs, setMsgs] = useState([]);
  const [problems, setProblems] = useState([]);
  const [last, setLast] = useState(null);         // remembered folder handle
  const [backups, setBackups] = useState([]);
  const [showFx, setShowFx] = useState(store.get("fx") === "1");
  const [hover, setHover] = useState(store.get("hover") !== "0");   // hover texts, on until the user turns them off
  const [learn, setLearn] = useState(store.get("learn") !== "0");   // on until the user turns it off
  const [explain, setExplain] = useState([]);     // stack of { t, k }: the last is shown, Back pops
  const [helpAt, setHelpAt] = useState(null);    // Help section to scroll to
  const [waits, setWaits] = useState({});
  const [exp, setExp] = useState(false);          // Export menu open
  const [expRuns, setExpRuns] = useState("last"); // runs for the report and workbook: "last" or "all"
  const [printing, setPrinting] = useState(null); // report preview: { runs, images }         // SolidWorks cards: macro id -> { since, marker } after "I pressed it"
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
  useEffect(() => { store.set("hover", hover ? "1" : "0"); }, [hover]);
  useEffect(() => { store.set("learn", learn ? "1" : "0"); }, [learn]);
  useEffect(() => { if (F.folderSupported) F.lastFolder().then(setLast); }, []);

  const calc = useMemo(() => {
    if (!project) return null;
    try { return calcProject(project, cad, results); } catch (e) { console.error(e); return { error: e.message }; }
  }, [project, cad, results]);
  const calcOk = calc && !calc.error ? calc : null;
  const canSave = mode === "folder" || mode === "manual";
  const files = F.names(pre);
  const newPrefixOk = F.PREFIX_RE.test(newPre.trim());
  const example = mode === "example";
  const fctx = useMemo(() => (calcOk ? { project, cad, results, calc: calcOk } : null), [project, cad, results, calcOk]);
  const openExplain = useCallback((t, k = "") => setExplain(s => [...s.slice(-30), { t, k }]), []);
  // "Change it" in the Explain panel: open the tab, open the body's Edit row if needed, then focus the field.
  const goEdit = useCallback(pl => {
    setTab(pl.tab);
    let n = 0;
    const step = () => {
      if (pl.body) window.dispatchEvent(new CustomEvent("pt-open-body", { detail: pl.body }));
      const main = document.querySelector("main");
      const f = pl.t && main && findField(main, pl.t, pl.k || "");
      if (!f) { if (pl.t && ++n < 12) setTimeout(step, 50); return; }
      f.scrollIntoView({ block: "center" });
      f.focus({ preventScroll: true });
      const box = f.closest("td, label") || f;
      box.classList.add("flash");
      setTimeout(() => box.classList.remove("flash"), 1800);
    };
    setTimeout(step, 30);
  }, []);
  const explainCtx = useMemo(() => ({ ctx: fctx, open: openExplain, showFx }), [fctx, openExplain, showFx]);
  const states = useMemo(() => tabStates({ project, cad, results, calc: calcOk, canSave: canSave || example }), [project, cad, results, calcOk, canSave, example]);

  // ---------- edits ----------
  const edit = useCallback((label, path, value, key) => { if (live.current.mode !== "example") setHist(h => commit(h, setIn(h.present, path, value), label, key)); }, []);
  const editFn = useCallback((label, fn, key) => { if (live.current.mode !== "example") setHist(h => commit(h, fn(h.present), label, key)); }, []);
  const doUndo = () => setHist(h => { const n = undo(h); if (n.did) say(n.did); return n; });
  const doRedo = () => setHist(h => { const n = redo(h); if (n.did) say(n.did); return n; });

  // ---------- open ----------
  const loadFiles = (pj, cj, rj, origin, n = F.names("")) => {
    const probs = [
      ...(pj ? schemaProblems(validateProject, n.project, pj) : []),
      ...(cj ? schemaProblems(validateCad, n.cad, cj) : []),
      ...(rj ? schemaProblems(validateResults, n.results, rj) : []),
    ];
    setProblems(probs);
    setWaits({});
    setHist(historyInit(pj));
    setSavedRef(pj);
    setSavedAt(pj?.savedAt || null);
    setCad(cj || {});
    setResults(rj || { runs: [] });
    if (pj && origin) say(`Opened ${pj.board?.name || "project"} from ${origin}: ${pj.parts?.length ?? 0} parts${cj?.import ? `, ${cj.import.bodies?.length ?? 0} bodies` : ""}${rj?.runs?.length ? `, ${rj.runs.length} runs` : ""}.`);
    if (probs.length) say("Some files do not match the format. See tab 0.", "error");
  };

  // which: the prefix to open when the folder has several projects (else the one there, or the last one used).
  const openFolder = async (handle, ask = true, which) => {
    if (live.current.dirty && !confirm("You have unsaved changes. Open another folder anyway?")) return;
    try {
      if (!(await F.allowed(handle, ask))) { say("The browser did not allow access to the folder.", "error"); return; }
      const list = await F.listProjects(handle);
      const remembered = store.get("pre." + handle.name);
      const p0 = which ?? (list.length === 1 ? list[0] : list.includes(remembered) ? remembered : null);
      setDir(handle);
      setLast(handle);
      setFound(list);
      if (!list.length || p0 === null) {
        stamps.current = {};
        setPre("");
        setBackups([]);
        setMode(list.length ? "choose" : "empty");
        loadFiles(null, null, null, handle.name);
        if (!list.length) setNewName(handle.name);
        say(list.length ? `${handle.name} has ${list.length} projects. Pick one on tab 0.` : `${handle.name} has no project yet. Type your prefix on tab 0 to start one.`);
        return;
      }
      const n = F.names(p0);
      const [pj, cj, rj] = await Promise.all([n.project, n.cad, n.results].map(x => F.readJson(handle, x)));
      stamps.current = { project: pj?.stamp || 0, cad: cj?.stamp || 0, results: rj?.stamp || 0 };
      store.set("pre." + handle.name, p0);
      setPre(p0);
      setBackups(await F.listBackups(handle, p0));
      setMode("folder");
      loadFiles(pj.data, cj?.data, rj?.data, `${handle.name}\\${n.project}`, n);
    } catch (e) { say(e.message, "error"); }
  };
  const pick = async () => {
    try { await openFolder(await F.pickFolder()); } catch (e) { if (e.name !== "AbortError") say(e.message, "error"); }
  };

  const openByHand = async e => {
    const files = [...e.currentTarget.files];
    const got = {};
    let p0 = "";
    try {
      for (const f of files) {
        const n = F.nameParts(f.name.replace(/ \(\d+\)(\.json)$/i, "$1"));   // "AB_project (1).json" from a download
        if (!n) continue;
        got[n.kind] = JSON.parse(await f.text());
        if (n.kind === "project") p0 = n.pre;
      }
    } catch (err) { say("A file is not valid JSON: " + err.message, "error"); return; }
    if (!got.project) { say("Pick the project file (AB_project.json), with its cad and results files if the folder has them.", "error"); return; }
    setDir(null);
    setPre(p0);
    setMode("manual");
    loadFiles(got.project, got.cad, got.results, "files opened by hand", F.names(p0));
  };

  // The Raspberry Pi 4 study, built into the page, read only (design 5.3).
  const openExample = () => {
    if (dirty && !confirm("You have unsaved changes. Open the example anyway?")) return;
    setDir(null);
    setPre("");
    setMode("example");
    loadFiles(EXAMPLE.project, EXAMPLE.cad, EXAMPLE.results, "the built-in example");
    say("This is the example study, read only. Look around; to work on your own board, choose its folder on tab 0.");
  };

  // The prefix typed in the new project form, with its "_", or null (the user has been told why).
  const newPrefix = async () => {
    const typed = newPre.trim();
    if (!F.PREFIX_RE.test(typed)) { say("Type a prefix first: your initials, letters, digits and - only (for example AB).", "error"); return null; }
    const p0 = typed + "_";
    if ((await F.listProjects(dir)).some(x => x.toLowerCase() === p0.toLowerCase())) { say(`${F.names(p0).project} is already in ${dir.name}. Use another prefix.`, "error"); return null; }
    store.set("initials", typed);
    return p0;
  };

  // A new project: the project file, empty cad and results files and the CAD folder, all with the prefix.
  const startNew = async () => {
    const p0 = await newPrefix();
    if (p0 === null) return;
    const n = F.names(p0), t = now();
    try {
      const p = newProject(newName.trim() || dir.name, t, p0);
      const steps = [];
      for await (const [name, h] of dir.entries()) if (h.kind === "file" && /\.(step|stp)$/i.test(name)) steps.push(name);
      if (steps.length === 1) p.board.stepFile = steps[0];
      const wrote = [];
      if (await F.writeIfMissing(dir, n.cad, emptyCad(t))) wrote.push(n.cad);
      if (await F.writeIfMissing(dir, n.results, emptyResults(t))) wrote.push(n.results);
      await F.makeFolder(dir, "CAD");
      const [cj, rj] = await Promise.all([n.cad, n.results].map(x => F.readJson(dir, x)));
      stamps.current = { project: 0, cad: cj?.stamp || 0, results: rj?.stamp || 0 };
      store.set("pre." + dir.name, p0);
      setPre(p0);
      setFound(f => [...f, p0].sort());
      setBackups([]);
      setMode("folder");
      loadFiles(p, cj?.data, rj?.data, null, n);
      setSavedRef(null);
      if (!(await save(false, p, cj?.data, rj?.data, p0))) return;
      say(`New project ready in ${dir.name}: ${[n.project, ...wrote].join(", ")} and the CAD folder.${steps.length === 1 ? " STEP file: " + steps[0] + "." : ""}`);
    } catch (err) { console.error(err); say("Could not make the project: " + err.message, "error"); }
  };

  const fromWorkbook = async e => {
    const f = e.currentTarget.files[0];
    e.currentTarget.value = "";
    if (!f) return;
    const p0 = await newPrefix();
    if (p0 === null) return;
    const n = F.names(p0);
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellFormula: false, cellNF: false, cellDates: false });
      const base = f.name.replace(/\.xlsx$/i, "");
      const out = workbookToProject(wb, XLSX, {
        fileName: f.name, now: now(), savedBy: "page " + PAGE_VERSION,
        reports: { buildBoard: await F.readText(dir, `CAD/${base}_BuildBoard_report.txt`), buildSpreader: await F.readText(dir, `CAD/${base}_BuildSpreader_report.txt`) },
      });
      const wrote = [];
      if (await F.writeIfMissing(dir, n.cad, out.cad)) wrote.push(n.cad);
      if (await F.writeIfMissing(dir, n.results, out.results)) wrote.push(n.results);
      await F.makeFolder(dir, "CAD");
      const [cj, rj] = await Promise.all([n.cad, n.results].map(x => F.readJson(dir, x)));
      stamps.current = { project: 0, cad: cj?.stamp || 0, results: rj?.stamp || 0 };
      store.set("pre." + dir.name, p0);
      setPre(p0);
      setFound(x => [...x, p0].sort());
      setBackups([]);
      setMode("folder");
      loadFiles(out.project, cj?.data, rj?.data, f.name, n);
      setSavedRef(null);
      if (!(await save(false, out.project, cj?.data, rj?.data, p0))) return;
      say(`Made the project from ${f.name}. Wrote ${[n.project, ...wrote].join(", ")}.${out.notes.length ? " Notes: " + out.notes.join(" ") : ""}`);
    } catch (err) { console.error(err); say("Could not read the workbook: " + err.message, "error"); }
  };

  // ---------- save ----------
  const live = useRef({});
  live.current = { project, cad, results, dir, mode, dirty, pre };
  // True when saved. pIn (with cIn, rIn, preIn) saves a project just made, before the page state has caught up.
  const save = async (auto, pIn, cIn, rIn, preIn) => {
    const L = live.current;
    const p = pIn || L.project, d = L.dir;
    if (!p) return false;
    const pr = preIn ?? L.pre, fname = F.names(pr).project;
    const t = now();
    const out = projectForSave(p, cIn || L.cad, rIn || L.results, t);
    const probs = schemaProblems(validateProject, fname, out);
    const refs = out.parts.map(x => x.refDes);
    const dup = [...new Set(refs.filter((x, i) => refs.indexOf(x) !== i))];
    if (dup.length) probs.push(`RefDes used by two parts: ${dup.join(", ")}. Give each part its own RefDes.`);
    if (probs.length) {
      setProblems(probs);
      if (!auto) say("Not saved: " + probs[0], "error");
      return false;
    }
    setProblems([]);
    try {
      if ((pIn ? "folder" : L.mode) === "folder" && d) {
        const disk = await F.stamp(d, fname);
        if (disk && stamps.current.project && disk !== stamps.current.project) {
          if (auto) { say(`Not saved on its own: ${fname} was changed outside this page. Press Save to replace it.`, "error"); return false; }
          if (!confirm(`${fname} was changed outside this page since it was opened. Replace it with this page's version?`)) return false;
        }
        const r = await F.saveProject(d, out, today(), pr);
        stamps.current.project = r.stamp;
        if (r.backup) { say(`Backup of the old file: ${r.backup}`); setBackups(await F.listBackups(d, pr)); }
        if (!auto) say(`Saved ${fname} in ${d.name}.`);
      } else {
        F.download(fname, JSON.stringify(out, null, 2) + "\n");
        say(`${fname} downloaded. Put it in the project folder, over the old one.`);
      }
      setSavedRef(p);
      setSavedAt(t);
      return true;
    } catch (e) { say("Not saved: " + e.message, "error"); return false; }
  };

  // Save on its own every 30 s while there are unsaved edits (folder mode only).
  useEffect(() => {
    const id = setInterval(() => { const L = live.current; if (L.dirty && L.mode === "folder") save(true); }, 30000);
    return () => clearInterval(id);
  }, []);

  // Reload the cad and results files when a macro has written them.
  useEffect(() => {
    if (mode !== "folder" || !dir) return;
    const n = F.names(pre);
    const id = setInterval(async () => {
      for (const [name, key, set] of [[n.cad, "cad", setCad], [n.results, "results", setResults]]) {
        try {
          const st = await F.stamp(dir, name);
          if (st === stamps.current[key]) continue;
          const r = await F.readJson(dir, name);
          stamps.current[key] = st;
          if (!r) continue;
          set(r.data);
          say(key === "cad" ? `${name} reloaded (${r.data.savedBy || "macro"}): ${r.data.import?.bodies?.length ?? 0} bodies${r.data.build ? ", build " + (r.data.build.date || "") : ""}.`
            : `${name} reloaded: ${r.data.runs?.length ?? 0} runs.`);
        } catch { /* half written or locked: try again in 2 s */ }
      }
      try {
        const st = await F.stamp(dir, n.project);
        if (st && stamps.current.project && st !== stamps.current.project && !stamps.current.warned) {
          stamps.current.warned = true;
          say(`${n.project} was changed outside this page. Saving here will ask before replacing it.`, "error");
        }
      } catch { /* ignore */ }
    }, 2000);
    return () => clearInterval(id);
  }, [mode, dir, pre]);

  // SolidWorks cards: stop waiting when the macro's file shows it has run, and say what it did.
  const setWait = useCallback((id, mark) => setWaits(w => { const n = { ...w }; if (mark === null) delete n[id]; else n[id] = { since: now(), marker: mark }; return n; }), []);
  useEffect(() => {
    for (const [id, w] of Object.entries(waits)) {
      if (marker(id, cad, results) === w.marker) continue;
      setWait(id, null);
      const runs = results?.runs || [], last = runs[runs.length - 1];
      const what = id === "ImportStep" ? `${cad.import?.bodies?.length ?? 0} bodies, PCB ${cad.import?.pcb?.box ? [cad.import.pcb.box.l, cad.import.pcb.box.w, cad.import.pcb.box.h].map(x => +x.toFixed(2)).join(" × ") + " mm" : "not found"}`
        : id === "BuildBoard" ? `${cad.build?.counts?.failed ? cad.build.counts.failed + " failed, " : ""}${cad.build?.counts?.wanted ?? 0} bodies wanted`
        : id === "SyncFromCAD" ? `${cad.asBuilt?.bodies?.length ?? 0} bodies differ, ${cad.asBuilt?.failedChecks?.length ?? 0} checks failed`
        : id === "BuildSpreader" ? ["top", "bottom"].filter(s => cad.spreaderBuild?.[s]).map(s => `${s} plate ${cad.spreaderBuild[s].pedestals ?? 0} pedestals, ${(cad.spreaderBuild[s].failures || []).length} problems`).join("; ")
        : last ? `run ${last.name}, ${last.status}` : "";
      say(`${MACROS[id].name} finished: ${what}.`);
    }
  }, [cad, results]);
  const imageUrl = useCallback(async rel => {
    if (!dir || mode !== "folder") return null;
    try { const f = await F.readFile(dir, rel); return f ? URL.createObjectURL(f) : null; } catch { return null; }
  }, [dir, mode]);

  // ---------- export (design 9) ----------
  const runChoice = () => { const runs = live.current.results?.runs || []; return expRuns === "all" ? runs.map(r => r.name) : runs.length ? [runs[runs.length - 1].name] : []; };
  const baseName = () => (pre || "") + safeFile(project?.board?.name || "board") + "_" + today();
  const doReport = async () => {
    setExp(false);
    const runs = runChoice(), images = {};
    for (const r of (results?.runs || []).filter(x => runs.includes(x.name) && x.plotFile)) {
      const u = await imageUrl(`${swFolder(project) || "CAD"}/results/${r.plotFile.split(/[\\/]/).pop()}`);
      if (u) images[r.name] = u;
    }
    setPrinting({ runs, images });
  };
  const doXlsx = () => {
    setExp(false);
    try {
      const wb = exportWorkbook(XLSX, fctx, runChoice());
      const bytes = XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true });
      F.download(baseName() + ".xlsx", new Uint8Array(bytes), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      say(`Workbook ${baseName()}.xlsx downloaded: ${wb.SheetNames.length} tabs, values only, a note on every header.`);
    } catch (e) { console.error(e); say("The workbook could not be made: " + e.message, "error"); }
  };
  const doZip = async () => {
    setExp(false);
    try {
      const n = F.names(pre), t = now(), files = [];
      files.push({ name: n.project, data: JSON.stringify(projectForSave(project, cad, results, t), null, 2) + "\n" });
      if (cad && Object.keys(cad).length) files.push({ name: n.cad, data: JSON.stringify(cad, null, 2) + "\n" });
      if (results) files.push({ name: n.results, data: JSON.stringify(results, null, 2) + "\n" });
      let extra = 0;
      if (dir && mode === "folder") {
        const out = swFolder(project);
        for (const [path, re] of out ? [[out, /\.txt$/i], [out + "/results", /\.(png|jpe?g|bmp)$/i]] : []) {
          for (const name of await F.listNames(dir, path)) {
            if (!re.test(name)) continue;
            const f = await F.readFile(dir, path + "/" + name);
            if (f) { files.push({ name: path + "/" + name, data: new Uint8Array(await f.arrayBuffer()), date: new Date(f.lastModified) }); extra++; }
          }
        }
      }
      F.download(baseName() + "_archive.zip", makeZip(files), "application/zip");
      say(`Archive ${baseName()}_archive.zip downloaded: the three project files${extra ? ` and ${extra} reports and pictures` : ""}.`);
    } catch (e) { console.error(e); say("The archive could not be made: " + e.message, "error"); }
  };

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
      say(`Opened ${name}. Press Save to keep it as ${files.project} (Ctrl+Z goes back).`);
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
      ${mode === "choose" && html`<div class="empty-folder">
        <p>This folder has ${found.length} projects. Open one:</p>
        ${found.map(x => html`<button onClick=${() => openFolder(dir, false, x)}>${F.names(x).project}</button>`)}
      </div>`}
      ${mode === "empty" && html`<div class="new-project">
        <p>This folder has no project yet. Start one here. The page writes all its files, empty and ready to use.</p>
        <div class="form">
          <label>Prefix (your initials) <input type="text" value=${newPre} maxlength="16" placeholder="AB" style="width:7em" aria-invalid=${!!newPre.trim() && !newPrefixOk}
            onInput=${e => setNewPre(e.currentTarget.value)} onKeyDown=${e => e.key === "Enter" && startNew()} /></label>
          <label>Board name <input type="text" value=${newName} style="width:18em" onInput=${e => setNewName(e.currentTarget.value)} onKeyDown=${e => e.key === "Enter" && startNew()} /></label>
        </div>
        <p class="hint">${newPrefixOk
          ? html`Files: ${["project", "cad", "results"].map(k => html`<code>${F.names(newPre.trim() + "_")[k]}</code>, `)}and the <code>CAD</code> folder. The prefix keeps the names unique in your file system.`
          : "Letters, digits and - only. The prefix goes in front of every file name, so names stay unique in your file system."}</p>
        <div class="empty-folder">
          <button class="primary" onClick=${startNew} disabled=${!newPrefixOk}>Start a new project here</button>
          <label class=${`button ${newPrefixOk ? "" : "disabled"}`}>Make it from an old workbook (.xlsx)<input type="file" accept=".xlsx" hidden disabled=${!newPrefixOk} onChange=${fromWorkbook} /></label>
        </div>
      </div>`}`
    : html`<p class="warn">This browser cannot save into a folder. Use Chrome or Edge for that. Here you can open the files by hand, and Save downloads <code>${files.project}</code>.</p>`}
    <p class="hint">${F.folderSupported ? "Or open the files by hand:" : ""} <label class="button small">Open the project, cad and results files (.json)<input type="file" accept=".json" multiple hidden onChange=${openByHand} /></label></p>
    <p class="hint">New to this? <button class="small" onClick=${openExample}>Open the example study (Raspberry Pi 4, read only)</button></p>
    ${problems.length > 0 && html`<div class="problems"><p class="warn">Format problems (Save is refused until they are fixed):</p><ul>${problems.map(x => html`<li>${x}</li>`)}</ul></div>`}
    ${calc?.error && html`<p class="warn">The calculation stopped: ${calc.error}</p>`}
  </div>`;
  const backupCard = mode === "folder" && backups.length > 0 && html`<p>Open a backup: ${backups.slice(0, 7).map(b => html` <button class="small" onClick=${() => openBackup(b)}>${b.slice(-13, -5)}</button>`)}</p>`;

  const listFolder = useMemo(() => (dir && mode === "folder" ? path => F.listNames(dir, path) : null), [dir, mode]);
  const ctx = { project, cad, results, calc: calcOk, states, listFolder, value: (t, k) => (fctx ? valueOf(fctx, t, k) : undefined), folder: example ? "built-in example (read only)" : dir?.name, edit, editFn, go: setTab, locked: st.locked || example, now, learn, openExample,
    blocking: calcOk ? blockingProblems(project, calcOk) : [], folderCard, backupCard, files,
    dirty, mode, tab, helpAt, waits, setWait, icons: ICONS, imageUrl: mode === "folder" ? imageUrl : null, outFolder: project ? swFolder(project) : "CAD" };
  const View = { project: V.ProjectTab, parts: V.PartsTab, import: V.ImportTab, link: V.LinkTab, triage: V.TriageTab, bare: V.BareTab,
    build: V.BuildTab, spreader: V.SpreaderTab, sim: V.SimTab, results: V.ResultsTab, changes: V.ChangesTab, help: V.HelpTab }[st.id];
  const needsProject = !["project", "help"].includes(st.id);

  if (printing && fctx) return html`<div class="print-root">
      <div class="print-bar">
        <button class="primary" onClick=${() => window.print()}>Print or save as PDF</button>
        <button onClick=${() => { for (const u of Object.values(printing.images)) URL.revokeObjectURL(u); setPrinting(null); }}>Back to the page</button>
        <span class="hint">In the print window choose Save as PDF as the printer. The report is always light.</span>
      </div>
      <${Report} ctx=${fctx} runs=${printing.runs} images=${printing.images} now=${now()} />
    </div>`;

  return html`
    <header class="top">
      <span class="brand">PCB Thermal</span>
      <span class="name">${project?.board?.name || "No project open"}</span>
      <span class=${`save ${dirty ? "dirty" : ""}`}>${!project ? "" : dirty ? "Unsaved changes" : savedAt ? "Saved " + hhmm(savedAt) : ""}</span>
      <span class="spacer"></span>
      <button onClick=${() => save(false)} disabled=${!project || !canSave} title=${`Save ${files.project} (Ctrl+S)`}>Save</button>
      <button onClick=${doUndo} disabled=${!hist.past.length} title=${hist.past.length ? "Undo: " + hist.past[hist.past.length - 1].label + " (Ctrl+Z)" : "Nothing to undo"}>Undo</button>
      <button onClick=${doRedo} disabled=${!hist.future.length} title=${hist.future.length ? "Redo: " + hist.future[0].label + " (Ctrl+Y)" : "Nothing to redo"}>Redo</button>
      <button aria-pressed=${learn} class=${learn ? "on" : ""} onClick=${() => setLearn(!learn)} title="Show what each step is for, what to look at, and a plain reading of the results">Learn</button>
      <button aria-pressed=${showFx} class=${showFx ? "on" : ""} onClick=${() => setShowFx(!showFx)} title="Write the formula, with your numbers, under every calculated value">Show formulas</button>
      <button aria-pressed=${hover} class=${hover ? "on" : ""} onClick=${() => setHover(!hover)} title=${hover ? "Hover texts are on: pointing at a value, label or abbreviation explains it. Click to turn them off (clicking a value still opens the Explain panel)." : "Hover texts are off. Click to turn them on again."}>Hover help</button>
      <button onClick=${() => setTheme(theme === "dark" ? "light" : "dark")} title="Switch between dark and light">${theme === "dark" ? "Light" : "Dark"}</button>
      <span class="menu-wrap">
        <button aria-expanded=${exp} disabled=${!fctx} onClick=${() => setExp(!exp)} title="Report (PDF), workbook (.xlsx) or project archive (.zip)">Export</button>
        ${exp && html`<div class="export-menu" role="menu">
          <label class="hint">Runs: <select value=${expRuns} onChange=${e => setExpRuns(e.currentTarget.value)} aria-label="Runs to export"><option value="last">the last run</option><option value="all">all runs</option></select></label>
          <button role="menuitem" onClick=${doReport}>Report (PDF)</button>
          <button role="menuitem" onClick=${doXlsx}>Workbook (.xlsx)</button>
          <button role="menuitem" onClick=${doZip}>Project archive (.zip)</button>
        </div>`}
      </span>
      <button onClick=${() => { setHelpAt(null); setTab(11); }}>Help</button>
    </header>
    <nav class="tabs" role="tablist">
      ${states.map(s => html`<button role="tab" aria-selected=${s.n === tab} class=${`tab ${s.status} ${s.n === tab ? "on" : ""}`}
        title=${s.locked ? "Locked: " + s.reason : s.note || ""} onClick=${() => setTab(s.n)}>
        <span class="n">${s.n}</span> <span class="t">${s.title}</span>${s.locked ? html` <span class="lock" aria-label="locked">🔒</span>` : ""}${s.badge ? html` <span class="badge">${s.badge}</span>` : ""}
      </button>`)}
    </nav>
    <${Explain.Provider} value=${explainCtx}>
    <div class=${`layout ${explain.length ? "with-panel" : ""}`}>
    <main>
      <h2><span class="step-n">${st.n}</span> ${st.title}</h2>
      ${st.locked && html`<div class="locked-banner" role="note"><b>This step is locked.</b> ${st.reason}
        ${st.fixTab !== undefined && html` <button onClick=${() => setTab(st.fixTab)}>Go to tab ${st.fixTab}</button>`}</div>`}
      ${example && html`<div class="example-banner" role="note"><b>Example study, read only.</b> Edits are off and nothing is saved. <button class="small" onClick=${() => setTab(0)}>Choose your own folder</button></div>`}
      ${learn && html`<${V.LearnBox} id=${st.id} />`}
      ${needsProject && !project ? html`<p class="hint">Open a project on tab 0 first.</p>`
        : needsProject && !calcOk ? html`<${Card} title="The calculation stopped"><p class="warn">${calc?.error}</p><//>`
        : html`<fieldset class="content" disabled=${(st.locked || example) && st.id !== "help"}>${View(ctx)}</fieldset>`}
    </main>
    <${ExplainPanel} at=${explain[explain.length - 1]} ctx=${fctx} onOpen=${openExplain} onEdit=${goEdit} canBack=${explain.length > 1}
      onBack=${() => setExplain(s => s.slice(0, -1))} onClose=${() => setExplain([])} onHelp=${id => { setHelpAt(helpSection(id)); setTab(11); }} />
    </div>
    <${Tip} on=${hover} />
    <//>
    <div class="msgs" aria-live="polite">${msgs.map(m => html`<div class=${`msg ${m.kind}`} key=${m.id}>${m.text}</div>`)}</div>
    <footer>page ${PAGE_VERSION} · ${mode === "folder" ? "saves into the folder" : mode === "manual" ? "files opened by hand: Save downloads " + files.project : example ? "example study, read only" : "no folder"}</footer>`;
}

export function start(el) {
  el.textContent = "";
  render(html`<${App} />`, el);
}
