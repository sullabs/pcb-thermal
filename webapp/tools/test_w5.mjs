// W5 test: a new user fills a board without the workbook (tabs 1 to 5), and the Learn mode texts.
//   node tools/test_w5.mjs
// 1. Start a new project, paste the RPi4 parts list, link bodies from the suggestions, add the two inductors by
//    hand, and the hand check gives the same numbers as the RPi4 example made from the workbook.
// 2. The parts list reader: units, percent, class names, separators, problems.
// 3. Body suggestions, bodies typed in by hand, the PCB stack and the air defaults of a new project.
// 4. Learn mode texts for every step tab, and the readings and the network drawing on tabs 4 and 5.
// 5. The parts sheet of tab 1: rows, insert, move, loads, and a block of cells pasted from Excel.
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { h } from "preact";
import { render } from "preact-render-to-string";
import { makeValidator } from "./validate.mjs";
import { build } from "./build.mjs";
import { calcProject } from "../src/calc/index.js";
import { valueOf } from "../src/calc/formulas.js";
import { newProject, setIn, projectForSave, addManualBody, withSimSetup } from "../src/page/model.js";
import { readPartsList, applyPartsList, partsListText, suggestLinks } from "../src/page/lists.js";
import { CLASSES } from "../src/calc/constants.js";
import { SHEET_COLS, sheetRows, insertPart, movePart, addLoad, moveLoad, deleteLoad, pasteCells, isBlock } from "../src/page/sheet.js";
import { tabStates, blockingProblems, TABS } from "../src/page/tabs.js";
import { Explain } from "../src/page/explain.js";
import * as Views from "../src/page/views.js";
import { LEARN } from "../src/terms/learn.js";
import { editPlace } from "../src/page/where.js";
import { swFiles, withFound, swPrefix, swFolder, safeFile } from "../src/page/swfiles.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const ex = path.join(here, "..", "..", "example", "RaspberryPi4");
const read = f => JSON.parse(fs.readFileSync(path.join(ex, f), "utf8"));
const ajv = makeValidator();
let fails = 0, n = 0;
const check = (ok, what, extra) => { n++; if (!ok) { fails++; console.log("FAIL " + what + (extra !== undefined ? ": " + JSON.stringify(extra).slice(0, 400) : "")); } };
const NOW = "2026-10-18T12:00:00";
const P = read("project.json"), C = read("cad.json"), R = read("results.json");
const close = (a, b) => (a === undefined && b === undefined) || (typeof a === "number" && Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)));

// ---------- 1. a new board from a pasted list ----------
{
  let p = newProject("RPi4 again", NOW);
  check(ajv.validate("project.schema.json", p), "new project is valid", ajv.errors);
  check(p.simSetup?.surfaces?.pcbTop?.h > 0 && p.handCheck?.surfaceRiseK === 20, "new project has the template air and surface values");

  const list = readPartsList(partsListText(P));
  check(list.problems.length === 0, "the RPi4 list reads without problems", list.problems);
  const r = applyPartsList(p, list);
  p = r.project;
  check(r.added.length === 9 && r.updated.length === 0, "9 parts added", r);
  check(JSON.stringify(p.parts) === JSON.stringify(P.parts), "pasted parts equal the example's, loads and ids too", p.parts.find((x, i) => JSON.stringify(x) !== JSON.stringify(P.parts[i])));

  // Tab 2 is open once a part has Pd; tab 3 once cad.json has the import.
  let calc = calcProject(p, C, {});
  let st = Object.fromEntries(tabStates({ project: p, cad: C, results: {}, calc, canSave: true }).map(s => [s.id, s]));
  check(!st.import.locked && !st.link.locked, "import and link tabs open", [st.import.reason, st.link.reason]);
  check(st.triage.locked && st.triage.fixTab === 3, "triage waits for a link", st.triage.reason);

  const sugg = suggestLinks(p, calc);
  const want = Object.fromEntries(Object.entries(P.bodies).filter(([k, b]) => b.refDes && !b.manual).map(([k, b]) => [k, b.refDes]));
  const got = Object.fromEntries(sugg.map(s => [s.key, s.refDes]));
  for (const [k, ref] of Object.entries(got)) check(want[k] === ref, `suggestion ${ref} is the right body`, k);
  check(["U1", "U2", "U3", "U4", "U5", "U6"].every(x => Object.values(got).includes(x)), "suggestions find U1 to U6 by part number", got);
  check(!Object.values(got).includes("J1"), "no guess for J1 (its part number has no digits)", got);
  p = sugg.reduce((a, s) => setIn(setIn(a, ["bodies", s.key, "refDes"], s.refDes), ["bodies", s.key, "include"], true), p);

  // The rest by hand, as the user would on tab 3: J1, the includes, the SD card position, two inductors.
  p = setIn(p, ["bodies", "femaleusbtypecconnectorrpi4modelb", "refDes"], "J1");
  for (const [k, b] of Object.entries(P.bodies)) if (!b.manual) p = setIn(p, ["bodies", k, "include"], b.include);
  p = setIn(p, ["bodies", "sdcardslotrpi4modelb", "override", "xMin"], 3.5);
  for (const k of ["m_l1", "m_l2"]) {
    const m = addManualBody(p, calcProject(p, C, {}).pcb.h);
    check(m.key.startsWith("m_"), "a body typed in by hand gets an m_ key", m.key);
    p = m.project;
    const src = P.bodies[k];
    for (const f of ["xMin", "yMin", "zMin", "l", "w", "h"]) p = setIn(p, ["bodies", m.key, "manual", "box", f], src.manual.box[f]);
    p = setIn(p, ["bodies", m.key, "manual", "name"], src.manual.name);
    p = setIn(p, ["bodies", m.key, "refDes"], src.refDes);
  }
  calc = calcProject(p, C, {});
  const ref = calcProject(P, C, {});
  check(blockingProblems(p, calc).length === 0, "no blocking problem: the hand check runs", blockingProblems(p, calc));
  st = Object.fromEntries(tabStates({ project: p, cad: C, results: {}, calc, canSave: true }).map(s => [s.id, s]));
  check(!st.triage.locked && !st.bare.locked && !st.build.locked, "triage, hand check and build open");
  for (const x of P.parts) {
    const a = calc.register.parts[x.refDes], b = ref.register.parts[x.refDes];
    check(close(a.pdW, b.pdW) && close(a.tjEstC, b.tjEstC) && a.flag === b.flag, `${x.refDes}: Pd, Tj est and flag as in the example`, [a.pdW, b.pdW, a.flag, b.flag]);
    const u = calc.bareBoard.parts[x.refDes], v = ref.bareBoard.parts[x.refDes];
    check(close(u.tjC, v.tjC) && u.status === v.status, `${x.refDes}: hand check Tj as in the example`, [u.tjC, v.tjC, u.status, v.status]);
  }
  check(close(calc.bareBoard.board.boardC, ref.bareBoard.board.boardC), "board temperature as in the example", [calc.bareBoard.board.boardC, ref.bareBoard.board.boardC]);
  const saved = projectForSave(p, C, {}, NOW);
  check(ajv.validate("project.schema.json", saved), "the new board saves valid", ajv.errors);
  check(saved.forSolidWorks?.bodies.length === ref.bodies.filter(b => b.include).length, "and has a forSolidWorks block for the build", saved.forSolidWorks?.bodies.length);
}

// ---------- 2. the parts list reader ----------
{
  const t = "RefDes\tPart number\tClass\tTj max\tRja\tRail\tVin\tIin\tEff\tVout\tIout\tPd\tColour\n"
    + "U9\tTPS62130\tswitch\t125\t35 K/W\t12V\t12\t\t85 %\t3.3\t2\t\tred\n"
    + "U9\t\t\t\t\t12V\t\t\t0.9\t1.8\t500 mA\t\t\n"
    + "R1\tsense\tSense\t155\t\t\t\t2 A\t\t\t\t\t\n"
    + "X1\tmystery\tunknown class\t\t\t\t\t\t\t\t\t250 mW\t\n"
    + "\tno ref\t\t\t\t\t\t\t\t\t\t\t\n";
  const l = readPartsList(t);
  const rows = l.rows;
  check(rows.length === 4, "4 rows with a RefDes", rows.length);
  check(rows[0].part.class === "Switching regulator" && rows[0].part.rjaCW === 35 && rows[0].load.efficiency === 0.85, "class by its start, K/W read as C/W, 85 % as 0.85", rows[0]);
  check(rows[1].load.ioutA === 0.5 && rows[2].load.iinA === 2 && rows[3].load.pdMeasuredW === 0.25, "mA and mW convert", [rows[1].load, rows[2].load, rows[3].load]);
  check(rows[2].part.class === "Sense resistor", "class name in any case", rows[2].part.class);
  check(l.problems.some(x => /Colour/.test(x)) && l.problems.some(x => /unknown class/.test(x)) && l.problems.some(x => /no RefDes/.test(x)), "unknown column, unknown class and a missing RefDes are reported", l.problems);
  const base = newProject("t", NOW);
  const a = applyPartsList(base, l).project;
  check(a.parts.length === 3 && a.parts[0].loads.length === 2 && a.parts[0].loads[1].id === "p1l2", "rows with one RefDes make one part with two loads", a.parts[0]);
  const calc = calcProject(a, {}, {});
  check(Math.abs(calc.register.parts.U9.pdW - (3.3 * 2 * (1 / 0.85 - 1) + 1.8 * 0.5 * (1 / 0.9 - 1))) < 1e-12, "the pasted buck works out its Pd", calc.register.parts.U9.pdW);
  check(calc.register.loads[a.parts[1].loads[0].id].problem === "missing input", "R1 without R: missing input", calc.register.loads[a.parts[1].loads[0].id]);
  check(ajv.validate("project.schema.json", projectForSave(a, {}, {}, NOW)), "pasted project saves valid", ajv.errors);
  const upd = applyPartsList(a, readPartsList("RefDes;Rja;R\nR1;;0.01\n"));
  check(upd.updated[0] === "R1" && upd.project.parts[1].loads[0].rOhm === 0.01 && upd.project.parts[1].tjMaxC === 155, "a semicolon list updates a part already on the list", upd.project.parts[1]);
  check(readPartsList("Part number,Rja\nX,3\n").rows.length === 0 && readPartsList("Part number,Rja\nX,3\n").problems[0].includes("RefDes"), "a list without RefDes says so");
  check(readPartsList("RefDes,Rja\nU1,abc\n").problems[0].includes("not a number"), "a wrong number is reported", readPartsList("RefDes,Rja\nU1,abc\n").problems);
}

// ---------- 3. suggestions, bodies by hand, stack, defaults ----------
{
  const calc = calcProject(P, C, R);
  check(suggestLinks(P, calc).length === 0, "the finished example has nothing left to suggest", suggestLinks(P, calc));
  const twin = { ...P, parts: [...P.parts, { id: "p99", refDes: "U99", partNumber: "Female Micro HDMI 2", loads: [{ id: "p99l1" }] }], bodies: {} };
  check(!suggestLinks(twin, calcProject(twin, C, R)).some(s => s.refDes === "U99"), "two bodies that match equally: no guess");
  const m = addManualBody(P, 1.6);
  check(m.key === "m_1" && m.project.bodies.m_1.manual.box.zMin === 1.6, "new body by hand sits on the board top", m.project.bodies.m_1);
  check(ajv.validate("project.schema.json", projectForSave(m.project, C, R, NOW)), "and saves valid", ajv.errors);
  const del = setIn(m.project, ["bodies", "m_1"], undefined);
  check(!del.bodies.m_1 && Object.keys(del.bodies).length === Object.keys(P.bodies).length, "deleting it removes only it");
  const { simSetup, ...noSim } = P;
  const ws = setIn(withSimSetup(noSim), ["simSetup", "surfaces", "pcbTop", "h"], 12);
  check(ajv.validate("project.schema.json", projectForSave(ws, C, R, NOW)) && ws.simSetup.surfaces.pcbBottom.h === 5, "editing h on a project without simSetup fills the template values first", ajv.errors);
  const two = setIn(P, ["pcbStack", "copperLayers"], 2);
  check(calcProject(two, C, R).pcb.kxy < calc.pcb.kxy, "fewer copper layers: less spreading in plane");
}

// ---------- 4. Learn mode, readings, network drawing ----------
{
  const steps = TABS.filter(t => t.id !== "help");
  for (const t of steps) {
    const l = LEARN[t.id];
    check(l && l.purpose && l.look.length >= 1, `Learn text for tab ${t.n}`);
    check(l && !/[—–]/.test(JSON.stringify(l)), `Learn text for tab ${t.n} has no long dashes`);
    if (l) {
      const out = render(h(Views.LearnBox, { id: t.id }));
      check(out.includes("What this step is for") && out.includes("What to look at"), `Learn box renders for tab ${t.n}`);
    }
  }
  const ctx = { project: P, cad: C, results: R, calc: calcProject(P, C, R) };
  const c = { ...ctx, states: tabStates({ ...ctx, canSave: true }), folder: "x", edit() {}, editFn() {}, go() {}, locked: false, now: () => NOW, learn: true,
    blocking: blockingProblems(P, ctx.calc), value: (t, k) => valueOf(ctx, t, k) };
  const show = (tab, learn) => render(h(Explain.Provider, { value: { ctx, open() {}, showFx: false } }, h(() => Views[tab]({ ...c, learn }))));
  const bare = show("BareTab", true);
  check(bare.includes("What the numbers say") && /The board sits at/.test(bare), "hand check reading in Learn mode");
  check(/goes down into the board|leaves from its top|takes heat in/.test(bare), "each part gets a plain reading of where its heat goes");
  check(!show("BareTab", false).includes("What the numbers say"), "no reading with Learn mode off");
  const worst = Object.entries(ctx.calc.bareBoard.parts).filter(([, r]) => Number.isFinite(r.marginC)).sort((a, b) => a[1].marginC - b[1].marginC)[0][0];
  check(bare.includes(`Heat paths of ${worst}`) && bare.includes(`data-term="bbp.rSpread" data-key="${worst}"`) && bare.includes('<svg class="net"'), `network drawing opens on the part with the lowest margin (${worst})`);
  const tri = show("TriageTab", true);
  check(/makes .* of the board&#39;s heat|makes .* of the board's heat/.test(tri) && /highest power density/.test(tri), "triage reading names the biggest and the densest part");
  const parts = show("PartsTab", true);
  check(parts.includes("Paste a parts list") && parts.includes('data-term="part.notes"'), "parts tab has the paste button and part notes");
  const link = show("LinkTab", true);
  check(link.includes("Add a body by hand") && link.includes('data-term="stack.cuT"') && link.includes('data-term="pcbov.h"') && link.includes('data-term="body.check"'), "link tab has bodies by hand, PCB size and stack, and the body check");
  const imp = show("ImportTab", true);
  check(imp.includes('data-term="set.pcbNameContains"') && imp.includes('data-term="set.minSize"'), "import tab has the import settings");
  // Change it in the Explain panel: every field on every tab is found on that tab, with its own term and key.
  {
    const TAB_FN = { ProjectTab: 0, PartsTab: 1, ImportTab: 2, LinkTab: 3, TriageTab: 4, BareTab: 5, BuildTab: 6, SpreaderTab: 7, SimTab: 8, ResultsTab: 9, ChangesTab: 10 };
    const wrong = []; let fields = 0;
    for (const [fn, n] of Object.entries(TAB_FN)) {
      const out = show(fn, false);
      const re = /data-term="([^"]+)" data-key="([^"]*)"[^>]*>(?:(?!<\/?span)[\s\S]){0,400}?<(?:input|select|textarea)\b|<(?:input|select|textarea)\b[^>]*data-term="([^"]+)" data-key="([^"]*)"/g;
      for (const m of out.matchAll(re)) {
        const t = m[1] || m[3], k = m[2] ?? m[4] ?? "";
        if (t.startsWith("file.")) continue;
        fields++;
        const pl = editPlace(ctx, t, k);
        if (!pl || pl.tab !== n || pl.t !== t || (pl.k || "") !== k) wrong.push(`${t} ${k} on tab ${n}: ${JSON.stringify(pl)}`);
      }
    }
    check(fields > 100 && !wrong.length, `Change it finds all ${fields} fields on their tab`, wrong.slice(0, 5));
    const ov = editPlace(ctx, "pcb.l", ""), bb = editPlace(ctx, "bbp.l", "U1");
    check(ov?.tab === 3 && ov.t === "pcbov.l" && bb?.tab === 3 && bb.t === "bodyset.l" && ctx.calc.bodies.find(b => b.key === bb.k)?.refDes === "U1", "PCB L and a part L lead to their overwrite fields on tab 3");
    check(editPlace(ctx, "bb.area", "") === null, "a calculated value has no Change it button");
  }
  // SolidWorks files on tab 6: names as Build Board writes them, found or missing from a folder listing.
  {
    const sw = swFiles(ctx.project, ctx.cad, ctx.calc);
    const pre = swPrefix(ctx.project);
    const inc = ctx.calc.bodies.filter(b => b.include);
    check(sw.folder === "CAD" && sw.rows[0].kind === "assembly" && sw.rows[0].file.toLowerCase().endsWith(".sldasm") && sw.rows[1].file === pre + "PCB.sldprt", "tab 6 files: assembly and PCB in the CAD folder");
    check(inc.every(b => sw.rows.some(r => r.key === b.key && r.file === pre + "P_" + safeFile(b.key) + ".sldprt")), "tab 6 files: one part file per included body, named as Build Board names it");
    check(safeFile("a b/c.d-e_f") === "a_b_c_d-e_f" && swPrefix({ board: { name: "My board" }, settings: {} }) === "My_board_" && swPrefix({ board: { name: "x" }, settings: { filePrefix: "SA_" } }) === "SA_", "file prefix follows PJ_Prefix and PT_SafeFile");
    check(swFolder({ settings: { outputFolder: "out\\cad\\" } }) === "out/cad" && swFolder({ settings: { outputFolder: "D:\\x" } }) === null, "output folder: relative kept, absolute not checked");
    const part = sw.rows.find(r => r.kind === "part");
    const got = withFound(sw.rows, [sw.rows[0].file, sw.rows[1].file.toUpperCase(), pre + "P_old_body.sldprt", "other.txt"], pre);
    check(got[0].found === true && got[1].found === true && got.find(r => r.key === part.key).found === false && got.some(r => r.kind === "extra" && r.name === "old_body") && !got.some(r => r.file === "other.txt"),
      "found, missing (case ignored as on Windows) and part files left from an earlier build");
  }
  const proj = show("ProjectTab", true);
  check(proj.includes('data-term="limit.criticalMargin"') && !tri.includes('data-term="limit.criticalMargin"') && bare.includes('data-term="surf.pcbTopH"') && bare.includes('data-term="hand.compareRun"'), "flag limits on tab 0 (not tab 4), air and surfaces and the compare run on tab 5");
  // A load without its inputs says what its method needs.
  const noVin = setIn(P, ["parts", 0, "loads", 0, "vinV"], undefined);
  const nctx = { project: noVin, cad: C, results: R, calc: calcProject(noVin, C, R) };
  const np = render(h(Explain.Provider, { value: { ctx: nctx, open() {}, showFx: false } }, h(() => Views.PartsTab({ ...c, ...nctx, value: (t, k) => valueOf(nctx, t, k) }))));
  check(np.includes("needs Vin and Iin") && /class="field\s+need"/.test(np), "a load without Vin says what it needs and marks the field");
}

// ---------- 5. the parts sheet ----------
{
  const rows = sheetRows(P);
  check(rows.length === P.parts.reduce((a, x) => a + x.loads.length, 0) && rows[0].j === 0, "one sheet row per load");
  const ins = insertPart(P, 1);
  check(ins.project.parts[1].id === ins.id && ins.project.parts.length === P.parts.length + 1 && !P.parts.some(x => x.refDes === ins.project.parts[1].refDes), "insert a part between rows with a free RefDes");
  const mv = movePart(P, 0, 2);
  check(mv.parts[2].id === P.parts[0].id && mv.parts[0].id === P.parts[1].id && mv.parts.length === P.parts.length, "move a part down two rows");
  check(movePart(mv, 2, 0).parts.map(x => x.id).join() === P.parts.map(x => x.id).join(), "and back");
  const multi = P.parts.find(x => x.loads.length > 1);
  const ml = moveLoad(P, multi.id, 0, multi.loads.length - 1).parts.find(x => x.id === multi.id);
  check(ml.loads[multi.loads.length - 1].id === multi.loads[0].id, "move a load to the end of its part");
  const al = addLoad(P, multi.id, 1).parts.find(x => x.id === multi.id);
  check(al.loads.length === multi.loads.length + 1 && !multi.loads.some(l => l.id === al.loads[1].id), "add a load in the middle with a new id");
  const one = P.parts.find(x => x.loads.length === 1);
  check(deleteLoad(P, one.id, one.loads[0].id).parts.find(x => x.id === one.id).loads.length === 1, "the last load of a part is kept");
  check(isBlock("1\t2") && isBlock("a\nb") && !isBlock("1.5") && !isBlock("1.5\r\n"), "a block is told from a single cell");
  // Paste RefDes, part number, class into the last row, with one line more than the sheet has.
  const last = sheetRows(P).length;
  const pc = pasteCells(P, last - 1, "refDes", "Q7\tBSS138\t" + CLASSES[0].toLowerCase() + "\nQ8\tBSS84\tnot a class\n");
  const q7 = pc.project.parts.find(x => x.refDes === "Q7"), q8 = pc.project.parts.find(x => x.refDes === "Q8");
  check(q7 && q7.partNumber === "BSS138" && q7.class === CLASSES[0] && q8 && q8.partNumber === "BSS84" && pc.added === 1, "pasted block renames the last part and adds one", pc);
  check(pc.problems.length === 1 && /not on the list/.test(pc.problems[0]), "a wrong class is reported, not written", pc.problems);
  const lc = SHEET_COLS.findIndex(x => x.key === "vinV");
  const pv = pasteCells(P, 0, "vinV", "5 V\t300 mA\nabc");
  const r0 = sheetRows(pv.project)[0];
  check(r0.load.vinV === 5 && r0.load.iinA === 0.3 && pv.problems.length === 1 && lc > 0, "numbers pasted with units, a bad number reported", pv.problems);
  check(ajv.validate("project.schema.json", projectForSave(pc.project, C, R, NOW)), "the project after a paste saves valid", ajv.errors);
}

// ---------- the example inside the page ----------
{
  const out = path.join(os.tmpdir(), "pcbt_w5_build.html");
  const { page } = await build({ out, quiet: true });
  check(page.includes("BCM2711B0 SoC") && page.includes("Open the example study"), "the built page carries the example study");
  fs.rmSync(out, { force: true });
}

console.log(`W5: ${n - fails} of ${n} checks passed.`);
if (fails) process.exit(1);
