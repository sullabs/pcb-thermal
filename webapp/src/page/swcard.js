// SolidWorks cards (design 4.3). The page cannot start SolidWorks, so each SolidWorks step is a card: what the macro
// does, the checks before pressing it (each ticked or crossed, with a link to the fix), which button to press, a
// waiting state after the user says they pressed it, and the outcome once the macro has written its file.
//
//   MACROS[id]                     the seven macros: button number, name, icon, what it does, what the page shows after
//   macroChecks(id, c)              [{ ok, text, fix, soft }] for the card; pure, tested by tools/test_w7.mjs
//   marker(id, cad, results)        a text that changes when the macro has written its file (the app compares it)
//   <SwCard c=${c} id="BuildBoard"> the card; children are the outcome

import { html, Card } from "./ui.js";
import { marked } from "./explain.js";
import { unticked } from "./tabs.js";
import { isNum } from "../calc/util.js";
import { plural } from "./bits.js";

export const MACROS = {
  ImportStep: { n: 2, name: "Import STEP", icon: "2_ImportStep", file: "cad",
    does: "Reads the STEP file named on tab 0 (or the model open in SolidWorks), finds the PCB and one body per component, and writes their sizes and positions to the cad file.",
    after: "Tab 2 shows the board drawing and tab 3 lists the bodies." },
  BuildBoard: { n: 3, name: "Build Board", icon: "3_BuildBoard", file: "cad",
    does: "Builds or updates the thermal assembly in the CAD folder: one box for the PCB and one per included body. A second run only changes what moved, resized or was switched off.",
    after: "This tab shows what was built, what failed and why, and the overlaps." },
  BuildSpreader: { n: 4, name: "Build Spreader", icon: "4_BuildSpreader", file: "cad",
    does: "Completes the spreader plate in the assembly: plate thickness, a pedestal per covered chip, clearance cuts for tall parts, and a TIM part per chip in TIM mode Body. Draw the plate outline first, or let it make a rectangle the size of the PCB.",
    after: "This tab shows the plate file, pedestals, TIM parts and problems." },
  SyncFromCAD: { n: 5, name: "Sync CAD", icon: "5_SyncCADtoExcel", file: "cad",
    does: "Reads the thermal assembly back: real sizes of the PCB and parts, the plate as built (area, mass, holes), pedestals and TIMs as found, and checks for missing faces and interference. It changes no SolidWorks file.",
    after: "Tab 10 lists what differs from the page, with a button to take the SolidWorks sizes." },
  SetupSim: { n: 6, name: "Run Thermal Study", icon: "6_RunThermalStudy", file: "results",
    does: "Makes a new steady state thermal study (PT_date_time) on the thermal assembly from this tab: materials, heat per part, contacts, convection and radiation per surface group, mesh. Then meshes and solves it. Older studies are never changed.",
    after: "The results file gets the new run. Then press Read Results." },
  RunFlow: { n: 7, name: "Run Flow Simulation", icon: "7_RunFlowSim", file: "results",
    does: "Sets the heat of every Flow volume source from tab 1 (the source name holds the RefDes) and the ambient, solves the Flow project you made once by hand, and reads its goals.",
    after: "Tab 9 shows the new run (FL_date_time)." },
  ReadResults: { n: 8, name: "Read Results", icon: "8_ReadResults", file: "results",
    does: "Reads the solved thermal study: peak and average temperature per RefDes, PCB and spreader minimum and maximum, the energy balance, and a picture of the plot.",
    after: "Tab 9 shows the run with every part against its Tj max and the hand check." },
};

// What changes when the macro has finished: the page stops waiting when this is different from when the user pressed.
export function marker(id, cad = {}, results = {}) {
  const runs = results?.runs || [];
  switch (id) {
    case "ImportStep": return cad.import?.date || "";
    case "BuildBoard": return cad.build?.date || "";
    case "BuildSpreader": return JSON.stringify([cad.spreaderBuild?.top?.date, cad.spreaderBuild?.bottom?.date]);
    case "SyncFromCAD": return cad.asBuilt?.date || "";
    default: return JSON.stringify(runs.map(r => [r.name, r.status]));
  }
}

const ok = (cond, text, fix, soft) => ({ ok: !!cond, text, fix, soft });

export function macroChecks(id, c) {
  const { project: p, cad = {}, results = {}, calc } = c;
  const out = [];
  const open = unticked(p, cad);
  const savedOnce = !!p.savedAt;
  out.push(ok(!c.dirty && savedOnce, c.dirty ? "Save the page first (Ctrl+S): the macro reads the project file, not this page." : savedOnce ? "Project saved" : "Save the project once first (Ctrl+S).", undefined));
  const ticked = ok(open === 0, open ? `Tick the ${plural(open, "change")} on tab 10 first. The macro refuses while changes are unticked.` : "No unticked changes", 10);
  const built = ok(!!cad.build, cad.build ? "Thermal assembly built" : "Run Build Board first (tab 6).", 6);
  if (id === "ImportStep") {
    out.push(ok(!!p.board.stepFile, p.board.stepFile ? `STEP file: ${p.board.stepFile}` : "Type the STEP file name on tab 0 (or open the model in SolidWorks first).", 0, true));
    out.push(ok(Object.values(calc.register.parts).some(x => isNum(x.pdW)), "At least one part with Pd on tab 1", 1, true));
  } else if (id === "BuildBoard") {
    out.push(ok(!!cad.import, cad.import ? "STEP imported" : "Run Import STEP first (tab 2).", 2));
    out.push(ticked);
    const inc = calc.bodies.filter(b => b.include).length;
    out.push(ok(inc > 0, inc ? `${plural(inc, "included body", "included bodies")} to build` : "Include at least one body on tab 3.", 3));
    out.push(ok(!!calc.fsNow, calc.fsNow ? "PCB size and k known" : "The PCB size or k is not known yet: check tab 3.", 3));
    const savedAfter = !cad.import?.date || (p.savedAt && p.savedAt >= cad.import.date);
    out.push(ok(savedAfter, savedAfter ? "Saved after the last import" : "Save after the last Import STEP: Build Board builds from what the page worked out on Save.", undefined));
    const probs = calc.geometry.bodies.filter(b => b.include && b.check && b.check !== "OK" && !/overhang/i.test(b.check)).length;
    out.push(ok(!probs, probs ? `${plural(probs, "included body", "included bodies")} with a problem on tab 3` : "No included body with a problem", 3, true));
  } else if (id === "BuildSpreader") {
    out.push(built, ticked);
    const sides = ["top", "bottom"].filter(s => p.spreader?.[s]?.enabled);
    out.push(ok(sides.length, sides.length ? `Plate${sides.length > 1 ? "s" : ""}: ${sides.join(" and ")}` : "Add and enable a plate on this tab first.", 7));
    for (const s of sides) {
      const ch = calc.spreader[s]?.chips || [];
      const cov = ch.filter(x => x.covered);
      const bad = cov.filter(x => x.check !== "OK");
      out.push(ok(cov.length && !bad.length, !cov.length ? `${s} plate: no chip is covered` : bad.length ? `${s} plate: ${bad.map(x => x.refDes).join(", ")} not OK (see the Check column)` : `${s} plate: ${plural(cov.length, "covered chip")}, all OK`, 7, !cov.length));
    }
    out.push(ok(true, "Draw the plate outline in the assembly first if you want a shape other than the PCB rectangle (a part with Spreader in its name, sketch named Outline).", undefined, true));
  } else if (id === "SyncFromCAD") {
    out.push(built);
  } else if (id === "SetupSim" || id === "RunFlow") {
    out.push(built, ticked);
    const type = p.simSetup?.studyType || "Thermal";
    out.push(ok(type === (id === "SetupSim" ? "Thermal" : "Flow"), `Study type on this tab: ${type}${type === (id === "SetupSim" ? "Thermal" : "Flow") ? "" : `. This button needs ${id === "SetupSim" ? "Thermal" : "Flow"}.`}`, 8));
    if (calc.build?.outOfDate) out.push(ok(false, "The bodies or the PCB changed after the last build: run Build Board again first.", 6));
    const heatNoBody = p.parts.filter(x => calc.register.parts[x.refDes]?.pdW > 0 && !calc.bodies.some(b => b.include && b.refDes === x.refDes)).length;
    out.push(ok(!heatNoBody, heatNoBody ? `${plural(heatNoBody, "part")} with Pd but no included body: their heat is not in the model` : "Every part with Pd has a body", 3, true));
    const plates = ["top", "bottom"].filter(s => p.spreader?.[s]?.enabled);
    if (plates.length) out.push(ok(!!cad.spreaderBuild, cad.spreaderBuild ? "Spreader built" : "A plate is enabled but Build Spreader has not run: the study will be a bare board study.", 7, true));
    if (id === "SetupSim") {
      const bad = calc.simChips.filter(x => x.check !== "OK");
      out.push(ok(!bad.length, bad.length ? `Chips not OK: ${bad.map(x => x.refDes).join(", ")} (they get the k of parts without heat)` : "Every chip model OK", 8, true));
    } else {
      out.push(ok(true, "The Flow project was made once by hand on the thermal assembly: one volume source per chip, with the RefDes in its name, and temperature goals named like 'U1 max'.", undefined, true));
    }
  } else if (id === "ReadResults") {
    const waiting = (results.runs || []).filter(r => r.studyType === "Thermal" && (r.status === "solved" || r.status === "set up"));
    out.push(ok(waiting.length, waiting.length ? `Run to read: ${waiting[waiting.length - 1].name}` : "No solved thermal run waiting to be read. Run the study first.", 8));
  }
  return out;
}

const hhmm = t => (t ? t.slice(11, 16) : "");

export function SwCard({ c, id, children, title }) {
  const m = MACROS[id];
  const checks = macroChecks(id, c);
  const wait = c.waits?.[id];
  const hard = checks.filter(x => !x.ok && !x.soft);
  const proj = c.files?.project || "project.json";
  return html`<${Card} title=${title || `In SolidWorks: ${m.n} ${m.name}`} cls="sw">
    <div class="sw-head">
      ${c.icons?.[m.icon] && html`<img class="sw-icon" src=${c.icons[m.icon]} alt="" width="32" height="32" />`}
      <p>${marked(m.does)}</p>
    </div>
    <h4>Before you press it</h4>
    <div class="sw-checks">${checks.map(x => html`<p class=${`chk ${x.ok ? "ok" : x.soft ? "soft" : "bad"}`}>
      <span class="mark" aria-label=${x.ok ? "done" : "not yet"}>${x.ok ? "✓" : x.soft ? "!" : "✗"}</span> ${marked(x.text)}${!x.ok && x.fix !== undefined && x.fix !== c.tab ? html` <button class="link" onClick=${() => c.go(x.fix)}>Go to tab ${x.fix}</button>` : ""}</p>`)}</div>
    <p class="sw-press"><b>Press:</b> in SolidWorks, <b>PCB Thermal</b> tab, button <b>${m.n} ${m.name}</b>. When it asks which project, pick <code>${proj}</code> in ${c.folder || "the project folder"} (next time Yes is enough).</p>
    <p class="hint">${marked(m.after)}</p>
    ${wait ? html`<p class="sw-wait" role="status"><span class="spin" aria-hidden="true"></span> Waiting for ${m.name}... (pressed ${hhmm(wait.since)}). The page reads the ${m.file} file on its own when the macro has finished.
        <button class="small" onClick=${() => c.setWait(id, null)}>Stop waiting</button></p>`
      : html`<p><button class=${hard.length ? "" : "primary"} disabled=${c.mode !== "folder"} title=${c.mode !== "folder" ? "Open the project from its folder so the page can see when the macro has finished" : ""}
          onClick=${() => c.setWait(id, marker(id, c.cad, c.results))}>I pressed it</button>
        ${hard.length > 0 && html` <span class="hint">${plural(hard.length, "check")} above not met: the macro will refuse or build the wrong thing.</span>`}</p>`}
    ${children}
  <//>`;
}
