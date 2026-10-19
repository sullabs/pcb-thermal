// Step tabs and when they unlock (design 4.2). Pure, tested by tools/test_w3.mjs.
//
//   tabStates({ project, cad, results, calc, canSave }) -> one entry per tab:
//     { id, n, title, locked, reason, fixTab, status, note, badge }
//   status: "locked", "ready", "done" or "attention". A locked tab can still be opened to read.

import { isNum } from "../calc/util.js";

export const TABS = [
  { id: "project", n: 0, title: "Project" },
  { id: "parts", n: 1, title: "Parts and power" },
  { id: "import", n: 2, title: "Import geometry" },
  { id: "link", n: 3, title: "Link bodies" },
  { id: "triage", n: 4, title: "Triage" },
  { id: "bare", n: 5, title: "Hand check: bare board" },
  { id: "build", n: 6, title: "Build model" },
  { id: "spreader", n: 7, title: "Spreader" },
  { id: "sim", n: 8, title: "Simulation setup" },
  { id: "results", n: 9, title: "Results" },
  { id: "changes", n: 10, title: "Changes" },
  { id: "help", n: 11, title: "Help" },
];

export const unticked = (project, cad) => (cad?.changes || []).filter(c => !project?.checks?.[c.id]).length;

// Problems that stop the hand check: parts whose Pd cannot be worked out, and an unknown PCB.
export function blockingProblems(project, calc) {
  const out = [];
  for (const x of project.parts) for (const l of x.loads) {
    const pr = calc.register.loads[l.id]?.problem;
    if (pr) out.push(`${x.refDes}: ${pr}`);
  }
  const p = calc.pcb;
  if (![p.l, p.w, p.h].every(x => isNum(x) && x > 0)) out.push("PCB size unknown");
  if (![p.kxy, p.kz].every(x => isNum(x) && x > 0)) out.push("PCB k unknown (check the PCB stack)");
  return [...new Set(out)];
}

export function tabStates({ project, cad = {}, results = {}, calc, canSave }) {
  const loaded = !!project && !!calc;
  const withPd = loaded ? project.parts.filter(x => isNum(calc.register.parts[x.refDes]?.pdW)) : [];
  const linked = loaded ? project.parts.filter(x => calc.bodies.some(b => b.refDes === x.refDes)) : [];
  const included = loaded ? calc.bodies.filter(b => b.include).length : 0;
  const blocking = loaded ? blockingProblems(project, calc) : [];
  const open = loaded ? unticked(project, cad) : 0;
  const runs = results?.runs || [];

  // First failing rule gives the reason. Every tab past the project needs a project that can be saved.
  const need = (ok, reason, fixTab) => (ok ? null : { reason, fixTab });
  const base = need(loaded && canSave, "Choose the project folder first, so your work can be saved.", 0);
  const built = need(!!cad.build, "Run Build Board first (tab 6).", 6);
  const ticked = need(open === 0, `Tick the ${open} change${open === 1 ? "" : "s"} on tab 10 first. Building on top of unchecked changes can hide a mistake.`, 10);
  const rules = {
    project: [],
    parts: [base],
    import: [base, need(withPd.length > 0, "Add at least one part with power first. The import links bodies to your parts by RefDes.", 1)],
    link: [base, need(!!cad.import, "Run Import STEP in SolidWorks first (button 2). This tab lists the bodies it finds.", 2)],
    triage: [base, need(linked.length > 0, "Link at least one part to a body on tab 3. Power density needs the body footprint.", 3)],
    bare: [base, need(blocking.length === 0, `Fix the problems listed on tab 4 first: ${blocking.join("; ")}.`, 4)],
    build: [base, need(included > 0, "Include at least one body on tab 3. Build Board builds a box for each included body.", 3)],
    spreader: [base, built, ticked],
    sim: [base, built, ticked],
    results: [base, need(runs.length > 0, "Run the study first (tab 8). Results appear here when Read Results has finished.", 8)],
    changes: [],
    help: [],
  };

  const problems = loaded ? Object.values(calc.register.loads).filter(l => l.problem) : [];
  const status = {
    project: () => (!loaded ? ["ready", "Choose a project folder"] : project.board?.name && isNum(project.board?.ambientC) ? ["done"] : ["attention", "Board name and ambient are needed"]),
    parts: () => (!project.parts.length ? ["ready", "No parts yet"] : problems.length ? ["attention", `${problems.length} load${problems.length === 1 ? "" : "s"} without Pd`] : ["done"]),
    import: () => (cad.import ? ["done", `${cad.import.bodies?.length ?? 0} bodies`] : ["ready", "Waiting for Import STEP"]),
    link: () => { const n = withPd.filter(x => !linked.includes(x)).length; return n ? ["attention", `${n} part${n === 1 ? "" : "s"} with Pd have no body`] : ["done"]; },
    triage: () => (blocking.length ? ["attention", `${blocking.length} problem${blocking.length === 1 ? "" : "s"}`] : ["done"]),
    bare: () => { const f = calc.bareBoard.board.fail || 0; return f ? ["attention", `${f} part${f === 1 ? "" : "s"} FAIL`] : ["done"]; },
    build: () => (!cad.build ? ["ready", "Waiting for Build Board"] : cad.build.counts?.failed ? ["attention", `${cad.build.counts.failed} failed`]
      : calc.build?.outOfDate ? ["attention", "out of date: the PCB or bodies changed after the last build"] : ["done"]),
    spreader: () => (!project.spreader ? ["ready", "Optional"] : cad.spreaderBuild ? ["done"] : ["ready", "Waiting for Build Spreader"]),
    sim: () => {
      if (!runs.length) return ["ready", "Waiting for a run"];
      const type = project.simSetup?.studyType || "Thermal";
      const last = [...runs].reverse().find(r => r.studyType === type);
      const ch = last && calc.runs?.[last.name]?.changes;
      return ch?.length ? ["attention", `inputs changed after run ${last.name}`] : ["done", `${runs.length} run${runs.length === 1 ? "" : "s"}`];
    },
    results: () => {
      const last = runs[runs.length - 1], r = last && calc.runs?.[last.name];
      if (!r) return ["ready"];
      if (r.pass === "FAIL") return ["attention", `${r.fails} part${r.fails === 1 ? "" : "s"} over Tj max in ${last.name}`];
      if (r.balanceOk === false) return ["attention", `energy balance of ${last.name} above 2 %`];
      return r.pass === "PASS" ? ["done", `${last.name} PASS`] : ["ready", `${runs.length} run${runs.length === 1 ? "" : "s"}`];
    },
    changes: () => (open ? ["attention", `${open} to tick`] : ["done"]),
    help: () => ["ready"],
  };

  return TABS.map(t => {
    const fail = rules[t.id].find(Boolean);
    const out = { ...t, locked: !!fail, reason: fail?.reason, fixTab: fail?.fixTab };
    if (fail) out.status = "locked";
    else if (loaded || t.id === "project") [out.status, out.note] = status[t.id]();
    else out.status = "ready";
    if (t.id === "changes" && open) out.badge = open;
    return out;
  });
}

// The next step to show on the Project tab: the first open tab that is not done.
export function nextStep(states) {
  return states.find(s => s.n > 0 && s.n < 10 && !s.locked && s.status !== "done" && s.id !== "spreader" && s.id !== "results")
    || states.find(s => s.n > 0 && s.n < 10 && s.locked);
}
