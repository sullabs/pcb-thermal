// Tabs 6 to 10: build the model, spreader, simulation setup, results and changes, each with its SolidWorks cards.
// Same rules as views.js: every value is a <V>, every label a <L> with a catalogue term, every input carries its term.

import { useState, useMemo, useEffect } from "preact/hooks";
import { html, NumField, TextField, Select, Check, Card } from "./ui.js";
import { V, L, marked } from "./explain.js";
import { Scroll, Facts, date, shortName, pctV, plural } from "./bits.js";
import { SwCard } from "./swcard.js";
import { setIn, withSimSetup, addPlate, removePlate, newPlateChip, addSolution, removeSolution, addSimChips, takeAsBuilt, heatedOnSide } from "./model.js";
import { swFiles, withFound } from "./swfiles.js";
import { unticked } from "./tabs.js";
import { BALANCE_LIMIT_PCT } from "../calc/runs.js";
import { TERMS } from "../terms/catalogue.js";
import { isNum } from "../calc/util.js";

const SIDE = { top: "Top", bottom: "Bottom" };

// ---------- 6 Build model ----------
export function BuildTab(c) {
  return html`<${BuildView} c=${c} />`;
}
function BuildView({ c }) {
  const b = c.cad?.build;
  const sw = useMemo(() => swFiles(c.project, c.cad, c.calc), [c.project, c.cad, c.calc]);
  const [names, setNames] = useState(null);
  useEffect(() => {
    let live = true;
    setNames(null);
    if (c.listFolder && sw.folder) c.listFolder(sw.folder).then(n => live && setNames(n || []), () => live && setNames([]));
    return () => { live = false; };
  }, [c.listFolder, sw.folder, c.cad]);
  const rows = names ? withFound(sw.rows, names, sw.prefix) : sw.rows;
  const added = b?.counts?.added, wanted = b?.counts?.wanted;
  const failed = (b?.bodies || []).filter(x => x.status === "failed");
  const skipped = (b?.bodies || []).filter(x => x.status === "skipped");
  const found = r => r.found === true ? html`<span class="status ok" data-status="found">found</span>`
    : r.found === false ? html`<span class="status watch" data-status="missing">missing</span>` : html`<span class="hint">not checked</span>`;
  const why = !c.listFolder ? "Files are checked when the project is opened from its folder (not by hand, not the example)."
    : !sw.folder ? "The output folder is outside the project folder, so the page cannot look in it." : "";
  return html`
    <${SwCard} c=${c} id="BuildBoard">
      <p><${L} t="build.included" />: <${V} t="build.included" /></p>
      ${c.calc.build?.outOfDate && html`<p class="warn"><b>Out of date:</b> the PCB or the bodies changed on the page after the last build. Save, then run Build Board again.</p>`}
    <//>
    ${b ? html`<${Card} title="Last build" cls=${failed.length || c.calc.build?.outOfDate ? "attention" : ""}>
      ${failed.length > 0 && html`<div class="problems"><p class="warn"><b>Failed first:</b></p>${failed.map(x => html`<p class="warn">${shortName(x.name)}: ${x.reason || "no reason given"}</p>`)}</div>`}
      <${Facts} ctx=${c} rows=${[[, "build.assembly"], [, "build.date", "", undefined, date(b.date)], [, "build.mode"], [, "build.wanted"], [, "build.added"], [, "build.moved"], [, "build.resized"], [, "build.suppressed"], [, "build.broughtBack"], [, "build.failed"], [, "build.overlaps"], [, "build.outOfDate"]]} />
      ${isNum(added) && isNum(wanted) && added > 0 && html`<p class="hint">Added counts the PCB part too, so a first build adds one more than Wanted (the assembly file is not counted).</p>`}
      ${skipped.length > 0 && html`<p class="hint">Skipped: ${skipped.map(x => `${shortName(x.name)} (${x.reason || "no reason"})`).join("; ")}.</p>`}
      ${(b.overlaps || []).length > 0 && html`<h4>Overlapping boxes</h4><p class="hint">An overlap counts the material twice. Fix the boxes on tab 3 (Edit row) and build again.</p>
        <${Scroll}><table class="grid"><thead><tr><th><${L} t="body.name">Body A</${L}></th><th><${L} t="body.name">Body B</${L}></th><th class="num"><${L} t="build.overlaps">X, Y, Z overlap</${L}></th></tr></thead>
        <tbody>${b.overlaps.map(o => html`<tr><td>${bodyLabel(c, o.a)}</td><td>${bodyLabel(c, o.b)}</td><td class="num"><span class="hint">${o.size.map(x => x.toFixed(2)).join(" × ")} mm</span></td></tr>`)}</tbody></table><//>`}
    <//>` : html`<p class="hint">No build yet.</p>`}
    <${Card} title="SolidWorks files of this project">
      <p class="hint">The assembly, the PCB and one part per included body, as Build Board names them in ${sw.folder ? html`the <code>${sw.folder}</code> folder` : "the output folder"}. ${why}</p>
      <${Scroll}><table class="grid swfiles">
        <thead><tr><th><${L} t="body.name" /></th><th><${L} t="body.refDes" /></th><th><${L} t="body.include" /></th><th><${L} t="buildbody.status" /></th><th><${L} t="file.sw" /></th><th><${L} t="file.swFound" /></th></tr></thead>
        <tbody>${rows.map(r => html`<tr class=${r.kind === "part" && !r.included ? "dim" : ""}>
          <td>${r.kind === "part" ? html`<${V} t="body.name" k=${r.key} v=${shortName(r.name)} />` : html`<b>${r.kind === "extra" ? r.file : r.name}</b>`}</td>
          <td>${r.refDes ? html`<${V} t="body.refDes" k=${r.key} />` : ""}</td>
          <td>${r.kind === "part" ? html`<${V} t="body.include" k=${r.key} v=${r.included} />` : r.kind === "extra" ? html`<span class="hint">no: left out after an earlier build</span>` : ""}</td>
          <td>${r.status ? html`<${V} t="buildbody.status" k=${r.key} v=${r.status} />${r.reason ? html` <span class="hint">${r.reason}</span>` : ""}` : ""}</td>
          <td><${V} t="file.sw" k=${r.key} v=${r.file} /></td>
          <td>${found(r)}</td>
        </tr>`)}</tbody>
      </table><//>
    <//>
    <p class="hint">Changed a box by hand in SolidWorks? Run Sync CAD (<button class="link" onClick=${() => c.go(7)}>tab 7</button>): it reads the assembly back and tab 10 offers to take the new sizes.</p>`;
}
const bodyLabel = (c, key) => { const b = c.calc.bodies.find(x => x.key === key); return b ? html`<${V} t="body.name" k=${key} v=${shortName(b.name)} />` : shortName(key); };

// ---------- 7 Spreader ----------
export function SpreaderTab(c) {
  return html`<${SpreaderView} c=${c} />`;
}
function SpreaderView({ c }) {
  const { project: p, calc, locked } = c;
  const sides = ["top", "bottom"];
  return html`
    ${!p.spreader && html`<p class="hint">${marked("No spreader plate yet. A plate is optional: add one when a part is FAIL or Tight on tab 5 and most of its heat leaves from its top. Try the hand check below first: it shows what a plate would do before anything is built.")}</p>`}
    ${sides.map(s => p.spreader?.[s] ? html`<${PlateCard} c=${c} side=${s} />` : "")}
    <div class="toolbar">${sides.filter(s => !p.spreader?.[s]).map(s => html`<button disabled=${locked} onClick=${() => c.editFn(`Add ${s} plate`, q => addPlate(q, calc, s))}>Add a ${s} plate</button>`)}
      ${sides.filter(s => !p.handCheck?.[s === "top" ? "solutionTop" : "solutionBottom"]).map(s => html`<button disabled=${locked} onClick=${() => c.editFn(`Add ${s} plate hand check`, q => addSolution(q, calc, s))}>Add a ${s} plate hand check</button>`)}</div>
    ${sides.map(s => p.handCheck?.[s === "top" ? "solutionTop" : "solutionBottom"] && c.calc.solution[s] ? html`<${SolutionCard} c=${c} side=${s} />` : "")}
    ${p.spreader && html`<${SwCard} c=${c} id="BuildSpreader">${sides.filter(s => c.cad?.spreaderBuild?.[s]).map(s => html`<${PlateBuilt} c=${c} side=${s} />`)}<//>`}
    ${c.cad?.build && html`<${SwCard} c=${c} id="SyncFromCAD">${c.cad?.asBuilt && html`<${AsBuilt} c=${c} />`}<//>`}`;
}

const HEIGHT_MODES = ["Auto", "Manual"], CHIP_HEIGHT = ["Pedestal", "Thick TIM"], TIM_MODES = ["Body", "Contact R"];

function PlateCard({ c, side }) {
  const { project: p, calc, locked } = c;
  const pl = p.spreader[side];
  const r = calc.spreader[side];
  const base = ["spreader", side];
  const set = (f, label, v, need = true) => (v !== undefined || !need) && c.edit(`${SIDE[side]} plate ${label}`, [...base, f], v, `spreader.${side}.${f}`);
  const setChip = (i, f, label, v) => c.edit(`${pl.chips[i].refDes} ${label}`, [...base, "chips", i, f], v, `spreader.${side}.${i}.${f}`);
  const listed = new Set(pl.chips.map(x => x.refDes));
  const more = p.parts.map(x => x.refDes).filter(x => !listed.has(x));
  const heated = heatedOnSide(p, calc, side).filter(x => !listed.has(x));
  const [pick, setPick] = useState("");
  const f = (t, field, unit, rule, label, extra = {}) => html`<label><${L} t=${t} k=${side} /> <${NumField} term=${t} tkey=${side} value=${pl[field]} unit=${unit} label=${label} rule=${rule} disabled=${locked} onCommit=${v => set(field, label, v, rule.required)} ...${extra} /></label>`;
  return html`<${Card} title=${`${SIDE[side]} plate${pl.enabled ? "" : " (not enabled)"}`} cls="plate">
    <div class="form">
      <label><${L} t="plate.enabled" k=${side} /> <span class="field"><${Check} term="plate.enabled" tkey=${side} value=${pl.enabled} label="Plate enabled" disabled=${locked} onCommit=${v => set("enabled", "enabled", v)} /></span></label>
      <label><${L} t="plate.material" k=${side} /> <${TextField} term="plate.material" tkey=${side} value=${pl.material} label="Plate material" required width="11em" disabled=${locked} onCommit=${v => set("material", "material", v)} /></label>
      ${f("plate.k", "k", "WmK", { required: true, minEx: 0 }, "Plate k")}
      ${f("plate.density", "densityKgM3", "kg/m³", { required: true, minEx: 0 }, "Plate density", { accepts: "none" })}
      ${f("plate.thickness", "thicknessMm", "mm", { required: true, minEx: 0 }, "Plate thickness")}
      ${f("plate.clearanceXY", "clearanceXYMm", "mm", { required: true, min: 0 }, "XY clearance")}
      ${f("plate.clearanceZ", "clearanceZMm", "mm", { required: true, min: 0 }, "Z clearance")}
      ${f("plate.minPedestal", "minPedestalMm", "mm", { required: true, min: 0 }, "Smallest pedestal")}
      <label><${L} t="plate.heightMode" k=${side} /> <${Select} term="plate.heightMode" tkey=${side} value=${pl.heightMode} options=${HEIGHT_MODES} label="Height mode" disabled=${locked} onCommit=${v => set("heightMode", "height mode", v || "Auto")} /></label>
      ${pl.heightMode === "Manual" && f("plate.heightManual", "heightManualMm", "mm", { required: true }, "Manual height")}
    </div>
    <${Facts} ctx=${c} rows=${[[, "plate.faceZ", side, 3], [, "plate.covered", side], [, "plate.cuts", side]]} />
    <${Scroll}><table class="grid chips">
      <thead><tr><th><${L} t="pchip.refDes" /></th><th><${L} t="pchip.covered" /></th><th><${L} t="pchip.heightMode" /></th><th><${L} t="pchip.timMode" /></th>
        <th class="num"><${L} t="pchip.timTSet">TIM t</${L}></th><th class="num"><${L} t="pchip.timK" /></th>
        <th class="num"><${L} t="pchip.chipFace" /></th><th class="num"><${L} t="pchip.gap" /></th><th class="num"><${L} t="pchip.timT" /></th><th class="num"><${L} t="pchip.pedestal" /></th><th class="num"><${L} t="pchip.timR" /></th><th><${L} t="pchip.check" /></th><th></th></tr></thead>
      <tbody>${pl.chips.map((x, i) => { const k = `${side}/${i}`; const rr = r?.chips[i] || {}; return html`<tr key=${x.refDes + i} class=${x.covered ? "" : "dim"}>
        <td><${V} t="pchip.refDes" k=${k} /></td>
        <td><${Check} term="pchip.covered" tkey=${k} value=${x.covered} label=${"Covered " + x.refDes} disabled=${locked} onCommit=${v => setChip(i, "covered", "covered", v)} /></td>
        <td><${Select} term="pchip.heightMode" tkey=${k} value=${x.heightMode} options=${CHIP_HEIGHT} label="Height mode" blankText="(Pedestal)" disabled=${locked} onCommit=${v => setChip(i, "heightMode", "height mode", v)} /></td>
        <td><${Select} term="pchip.timMode" tkey=${k} value=${x.timMode} options=${TIM_MODES} label="TIM mode" blankText="(Body)" disabled=${locked} onCommit=${v => setChip(i, "timMode", "TIM mode", v)} /></td>
        <td class="num"><${NumField} term="pchip.timTSet" tkey=${k} value=${x.timTMm} unit="mm" label="TIM thickness" width="4.5em" rule=${{ minEx: 0 }} disabled=${locked} onCommit=${v => setChip(i, "timTMm", "TIM thickness", v)} /></td>
        <td class="num"><${NumField} term="pchip.timK" tkey=${k} value=${x.timK} unit="" label="TIM k" width="4em" rule=${{ minEx: 0 }} disabled=${locked} onCommit=${v => setChip(i, "timK", "TIM k", v)} /></td>
        <td class="num"><${V} t="pchip.chipFace" k=${k} d=${3} /></td><td class="num"><${V} t="pchip.gap" k=${k} d=${3} /></td><td class="num"><${V} t="pchip.timT" k=${k} d=${3} /></td>
        <td class="num"><${V} t="pchip.pedestal" k=${k} d=${3} /></td><td class="num"><${V} t="pchip.timR" k=${k} d=${2} /></td>
        <td class=${rr.check === "OK" || rr.check === "not covered" ? "" : "warn"}><${V} t="pchip.check" k=${k} /></td>
        <td><button class="small danger" disabled=${locked} title="Take this chip off the list" onClick=${() => c.editFn(`Remove ${x.refDes} from the ${side} plate`, q => setIn(q, [...base, "chips"], q.spreader[side].chips.filter((_, j) => j !== i)))}>Remove</button></td>
      </tr>`; })}</tbody>
    </table><//>
    <div class="toolbar">
      ${heated.length > 0 && html`<button class="small" disabled=${locked} onClick=${() => c.editFn(`List ${heated.length} heated parts on the ${side} plate`, q => setIn(q, [...base, "chips"], [...q.spreader[side].chips, ...heated.map(newPlateChip)]))}>List the ${plural(heated.length, "heated part")} on this side</button>`}
      <span class="field" data-term="pchip.refDes" data-key=""><select aria-label="Chip to add" value=${pick} disabled=${locked} onChange=${e => setPick(e.currentTarget.value)}><option value="">Add a part...</option>${more.map(x => html`<option value=${x}>${x}</option>`)}</select></span>
      <button class="small" disabled=${locked || !pick} onClick=${() => { c.editFn(`Add ${pick} to the ${side} plate`, q => setIn(q, [...base, "chips"], [...q.spreader[side].chips, newPlateChip(pick)])); setPick(""); }}>Add</button>
      <span class="spacer"></span>
      <button class="small danger" disabled=${locked} onClick=${() => c.editFn(`Remove the ${side} plate`, q => removePlate(q, side))}>Remove this plate</button>
    </div>
    <p class="hint">Tick Covered for each chip the plate touches. The face height (Auto) follows the tallest covered chip plus its TIM; the others get a pedestal. Parts it does not cover but that reach the plate get a clearance cut.</p>
  <//>`;
}

function PlateBuilt({ c, side }) {
  const sb = c.cad.spreaderBuild[side];
  return html`<div class="built"><h4>${SIDE[side]} plate as built</h4>
    ${(sb.failures || []).map(x => html`<p class="warn">${x}</p>`)}
    <${Facts} ctx=${c} rows=${[[, "sprb.file", side], [, "sprb.date", side, undefined, date(sb.date)], [, "sprb.pedestals", side], [, "sprb.tims", side], [, "sprb.cuts", side], [, "sprb.failures", side]]} />
  </div>`;
}

function AsBuilt({ c }) {
  const ab = c.cad.asBuilt;
  return html`<div class="built"><h4>Read back from the assembly</h4>
    <${Facts} ctx=${c} rows=${[[, "asb.date", "", undefined, date(ab.date)], [, "asb.bodies"], [, "asb.failedChecks"]]} />
    ${(ab.failedChecks || []).map(x => html`<p class="warn">${x}</p>`)}
    ${["top", "bottom"].filter(s => ab.plates?.[s]).map(s => html`<h4>${SIDE[s]} plate</h4><${Facts} ctx=${c} rows=${[[, "asbp.outlineArea", s, 0], [, "asbp.exposedArea", s, 0], [, "asbp.volume", s, 0], [, "asbp.mass", s, 1], [, "asbp.thickness", s, 3], [, "asbp.holes", s]]} />`)}
    ${(ab.bodies || []).length > 0 && html`<p><button class="link" onClick=${() => c.go(10)}>Tab 10</button> lists the bodies that differ, with a button to take the SolidWorks sizes.</p>`}
  </div>`;
}

const SOL_STACK = [["solp.package", 1], ["solp.footprint", 1], ["solp.plateShare", 0], ["solp.rjc", 3], ["solp.rTim1", 3], ["solp.rLocal", 3], ["solp.rTim2", 3], ["solp.rPlateSpread", 3], ["solp.rStack", 3], ["solp.tj", 1], ["solp.margin", 1]];
const SOL_IN = [["solp.tim1T", "tim1TMm", "mm"], ["solp.tim1K", "tim1K", ""], ["solp.localL", "localLMm", "mm"], ["solp.localW", "localWMm", "mm"], ["solp.localT", "localTMm", "mm"], ["solp.localK", "localK", ""],
  ["solp.tim2T", "tim2TMm", "mm"], ["solp.tim2K", "tim2K", ""], ["solp.shareOverride", "plateShareOverrideMm2", ""]];

function SolutionCard({ c, side }) {
  const { project: p, calc, locked } = c;
  const key = side === "bottom" ? "solutionBottom" : "solutionTop";
  const s = p.handCheck[key];
  const o = calc.solution[side];
  const base = ["handCheck", key];
  const set = (f, label, v) => c.edit(`Hand check ${label}`, [...base, f], v, `${key}.${f}`);
  const setPart = (ref, f, label, v) => c.edit(`${ref} ${label}`, [...base, "parts", ref, f], v, `${key}.${ref}.${f}`);
  const [detail, setDetail] = useState(null);
  const refs = p.parts.map(x => x.refDes);
  const under = refs.filter(r => s.parts?.[r]?.underPlate);
  const worst = under.filter(r => isNum(o.parts[r]?.marginC)).sort((a, z) => o.parts[a].marginC - o.parts[z].marginC);
  const num = (t, f, unit, label, rule = { minEx: 0 }) => html`<label><${L} t=${t} k=${side} /> <${NumField} term=${t} tkey=${side} value=${s[f]} unit=${unit} label=${label} rule=${rule} disabled=${locked} onCommit=${v => set(f, label, v)} /></label>`;
  return html`<${Card} title=${`Hand check with a ${side} plate`} cls="solution">
    <p class="hint">${marked("One common plate over the parts you tick. Each part's heat goes up through its Rjc, the TIM, an optional local spreader and the plate, then from the plate to air. Nothing is built: this shows whether a plate is worth it, and how big.")}</p>
    ${c.learn && worst.length > 0 && html`<div class="reading"><p><${V} nofx t="part.refDes" k=${worst[0]} /> is the closest to its limit with the plate: Tj <${V} nofx t="solp.tj" k=${`${side}|${worst[0]}`} d=${1} u />, margin <${V} nofx t="solp.margin" k=${`${side}|${worst[0]}`} d=${1} u />.
      ${isNum(o.plateC) && html` The plate itself sits at <${V} nofx t="sol.plateC" k=${side} d=${1} u />: ${o.plateC - p.board.ambientC > 0.6 * Math.max(1, (o.parts[worst[0]]?.tjC ?? 0) - p.board.ambientC) ? "most of the rise is the plate to air, so a bigger plate, fins or more air help most." : "most of the rise is in the stack, so a thinner TIM, a better TIM or a local copper spreader help most."}`}</p></div>`}
    <div class="form">
      ${num("sol.plateL", "plateLMm", "mm", "Plate length")}${num("sol.plateW", "plateWMm", "mm", "Plate width")}${num("sol.plateT", "plateTMm", "mm", "Plate thickness")}
      ${num("sol.plateK", "plateK", "WmK", "Plate k")}${num("sol.hAir", "hAir", "W/(m² K)", "h plate to air")}
      <label><${L} t="sol.faces" k=${side} /> <${Select} term="sol.faces" tkey=${side} value=${s.faces === undefined ? "" : String(s.faces)} options=${["1", "2"]} label="Faces to air" disabled=${locked} onCommit=${v => set("faces", "faces to air", v ? Number(v) : undefined)} /></label>
      ${num("sol.rOverride", "rPlateToAirOverrideCW", "CW", "R plate to air typed")}
    </div>
    ${o.warning && html`<p class="warn">The footprints add up to more than the plate area: the plate is too small for these parts.</p>`}
    <${Facts} ctx=${c} rows=${[[, "sol.area", side, 0], [, "sol.hEff", side, 2], [, "sol.rPlateAir", side, 3], [, "sol.partsUnder", side], [, "sol.pdPlate", side, 3], [, "sol.plateC", side, 1], [, "sol.footprints", side, 0]]} />
    <h4>Parts</h4>
    <${Scroll}><table class="grid sol">
      <thead><tr><th><${L} t="part.refDes" /></th><th><${L} t="solp.underPlate" /></th>${SOL_IN.map(([t]) => html`<th class="num"><${L} t=${t} /></th>`)}
        <th class="num"><${L} t="solp.tj" /></th><th class="num"><${L} t="solp.margin" /></th><th><${L} t="solp.status" /></th><th></th></tr></thead>
      <tbody>${refs.map(ref => { const k = `${side}|${ref}`; const x = s.parts?.[ref] || {}; const r = o.parts[ref] || {}; return html`<tr key=${ref} class=${x.underPlate ? "" : "dim"}>
        <td><${V} t="part.refDes" k=${ref} /></td>
        <td><${Check} term="solp.underPlate" tkey=${k} value=${x.underPlate} label=${"Under plate " + ref} disabled=${locked} onCommit=${v => setPart(ref, "underPlate", "under plate", v || undefined)} /></td>
        ${SOL_IN.map(([t, f, unit]) => html`<td class="num"><${NumField} term=${t} tkey=${k} value=${x[f]} unit=${unit} label=${TERMS[t].label} width="4.2em" rule=${{ minEx: 0 }} disabled=${locked || !x.underPlate} onCommit=${v => setPart(ref, f, TERMS[t].label, v)} /></td>`)}
        <td class="num"><${V} t="solp.tj" k=${k} d=${1} /></td><td class="num"><${V} t="solp.margin" k=${k} d=${1} /></td><td class=${r.status === "OK" || r.status === "not under plate" ? "" : "warn"}><${V} t="solp.status" k=${k} /></td>
        <td>${isNum(r.tjC) && html`<button class="small" aria-pressed=${detail === ref} onClick=${() => setDetail(detail === ref ? null : ref)}>Stack</button>`}</td>
      </tr>`; })}</tbody>
    </table><//>
    ${detail && o.parts[detail] && html`<div class="stack"><h4>Stack of ${detail}</h4><${Facts} ctx=${c} rows=${SOL_STACK.map(([t, d]) => [, t, `${side}|${detail}`, d]).concat([[, "solp.dtStack", `${side}|${detail}`, 2]])} /></div>`}
    <p class="hint">Blank local spreader fields mean no local spreader. <button class="small danger" disabled=${locked} onClick=${() => c.editFn(`Remove the ${side} plate hand check`, q => removeSolution(q, side))}>Remove this hand check</button></p>
  <//>`;
}

// ---------- 8 Simulation setup ----------
export function SimTab(c) {
  return html`<${SimView} c=${c} />`;
}
const SURF = [["spreaderTop", "surf.spreaderTopH", "surf.spreaderTopEps", "spreader top"], ["spreaderOther", "surf.spreaderOtherH", "surf.spreaderOtherEps", "other spreader faces"],
  ["pcbTop", "surf.pcbTopH", "surf.pcbTopEps", "PCB top"], ["pcbBottom", "surf.pcbBottomH", "surf.pcbBottomEps", "PCB bottom and sides"], ["parts", "surf.partsH", "surf.partsEps", "parts"]];
const CONTACTS = [["partToPcb", "simset.contactPartPcb", "part to PCB"], ["chipToTim", "simset.contactChipTim", "chip to TIM"], ["timToSpreader", "simset.contactTimSpreader", "TIM to spreader"]];

function SimView({ c }) {
  const { project: p, calc, locked } = c;
  const ss = p.simSetup || {};
  const setSim = (path, label, v, need = true) => (v !== undefined || !need) && c.editFn(label, q => setIn(withSimSetup(q), ["simSetup", ...path], v), "simSetup." + path.join("."));
  const chips = calc.simChips;
  const cols = [["sim.pd", 3], ["sim.rjc", 2], ["sim.h", 2], ["sim.area", 1], ["sim.kz", 2], ["sim.kxy", 2]];
  const type = ss.studyType || "Thermal";
  const two = (ss.chips || []).some(x => x.model === "Two-resistor");
  const listed = new Set((ss.chips || []).map(x => x.refDes));
  const missing = p.parts.filter(x => calc.register.parts[x.refDes]?.pdW > 0 && calc.bodies.some(b => b.refDes === x.refDes && b.include) && !listed.has(x.refDes)).length;
  const last = [...(c.results?.runs || [])].reverse().find(r => r.studyType === type);
  const lastCalc = last && calc.runs[last.name];
  return html`
    ${lastCalc?.changes?.length > 0 && html`<p class="warn"><b>Out of date:</b> ${marked(`since run ${last.name}: ${lastCalc.changes.slice(0, 6).join("; ")}${lastCalc.changes.length > 6 ? "; and more" : ""}. Run the study again to see the effect.`)}</p>`}
    <${SwCard} c=${c} id=${type === "Flow" ? "RunFlow" : "SetupSim"} />
    ${type === "Thermal" && (c.results?.runs || []).some(r => r.studyType === "Thermal" && r.status !== "read") && html`<${SwCard} c=${c} id="ReadResults" />`}
    <${Card} title="Study">
      <div class="form">
        <label><${L} t="simset.studyType" /> <${Select} term="simset.studyType" value=${ss.studyType} options=${["Thermal", "Flow"]} label="Study type" disabled=${locked} onCommit=${v => setSim(["studyType"], "Study type", v || "Thermal")} /></label>
        <label><${L} t="simset.meshGlobal" /> <${NumField} term="simset.meshGlobal" value=${ss.meshGlobalMm} unit="mm" label="Global mesh size" rule=${{ required: true, minEx: 0 }} disabled=${locked} onCommit=${v => setSim(["meshGlobalMm"], "Global mesh size", v)} /></label>
        <label><${L} t="simset.meshFine" /> <${NumField} term="simset.meshFine" value=${ss.meshFineMm} unit="mm" label="Fine mesh size" rule=${{ required: true, minEx: 0 }} disabled=${locked} onCommit=${v => setSim(["meshFineMm"], "Fine mesh size", v)} /></label>
        <label><${L} t="simset.includePartsWithoutHeat" /> <span class="field"><${Check} term="simset.includePartsWithoutHeat" value=${ss.includePartsWithoutHeat} label="Parts without heat in the study" disabled=${locked} onCommit=${v => setSim(["includePartsWithoutHeat"], "Parts without heat", v)} /></span></label>
        <label><${L} t="simset.kPartsWithoutHeat" /> <span><${V} t="simset.kPartsWithoutHeat" u /> <button class="link" type="button" onClick=${() => c.go(5)}>tab 5</button></span></label>
        ${two && html`<label><${L} t="simset.layerFrac" /> <${NumField} term="simset.layerFrac" value=${ss.layerFrac} unit="" label="Layer thickness" width="4em" placeholder="0.1" rule=${{ minEx: 0, max: 0.44 }} disabled=${locked} onCommit=${v => setSim(["layerFrac"], "Layer thickness", v, false)} /></label>
          <label><${L} t="simset.kCore" /> <${NumField} term="simset.kCore" value=${ss.kCore} unit="WmK" label="Core k" placeholder="100" rule=${{ minEx: 0 }} disabled=${locked} onCommit=${v => setSim(["kCore"], "Core k", v, false)} /></label>`}
      </div>
      <p class="hint">${type === "Flow" ? "Flow: the macro builds the Flow project itself (air, gravity, radiation, materials, a heat source and goals per chip) and sets heat and ambient. Surfaces and contacts below are for Thermal studies only." : "Thermal: SolidWorks Simulation, steady state, h and radiation on the outer faces."}</p>
    <//>
    <${Card} title="Air and surfaces">
      <table class="grid limits"><thead><tr><th></th><th><${L} t="surf.pcbTopH">h</${L}></th><th><${L} t="surf.pcbTopEps">emissivity</${L}></th></tr></thead>
      <tbody>${SURF.map(([g, th, te, name]) => g.startsWith("spreader") ? html`<tr><td>${name}</td>
        <td><${NumField} term=${th} value=${ss.surfaces?.[g]?.h} unit="W/(m² K)" accepts="none" label=${"h, " + name} rule=${{ required: true, min: 0 }} disabled=${locked} onCommit=${v => setSim(["surfaces", g, "h"], "h, " + name, v)} /></td>
        <td><${NumField} term=${te} value=${ss.surfaces?.[g]?.emissivity} unit="" label=${"Emissivity, " + name} width="4em" rule=${{ required: true, min: 0, max: 1 }} disabled=${locked} onCommit=${v => setSim(["surfaces", g, "emissivity"], "Emissivity, " + name, v)} /></td></tr>`
        : html`<tr><td>${name}</td><td><${V} t=${th} u /></td><td><${V} t=${te} /></td></tr>`)}</tbody></table>
      <p class="hint">PCB and parts are typed on <button class="link" type="button" onClick=${() => c.go(5)}>tab 5</button>, so the hand check and the study see the same air. Ambient: <${V} t="board.ambient" d=${1} u /> (tab 0).</p>
    <//>
    <${Card} title="Contacts">
      <div class="form">${CONTACTS.map(([f, t, name]) => html`<label><${L} t=${t} /> <${NumField} term=${t} value=${ss.contacts?.[f]} unit="K m²/W" accepts="none" label=${"Contact R " + name} rule=${{ required: true, min: 0 }} disabled=${locked} onCommit=${v => setSim(["contacts", f], "Contact R " + name, v)} /></label>`)}</div>
      <p class="hint">0 means bonded. A chip in TIM mode Contact R (tab 7) gets its TIM resistance as a contact on its own.</p>
    <//>
    <${Card} title="Chip models">
      <div class="toolbar">${missing > 0 && html`<button class="small" disabled=${locked} onClick=${() => c.editFn(`Add ${missing} chips`, q => addSimChips(q, calc))}>Add the ${plural(missing, "part")} with heat that are not listed</button>`}
        <span class="hint">Chips not listed get the k of parts without heat.</span></div>
      <${Scroll}><table class="grid"><thead><tr><th><${L} t="sim.refDes" /></th><th><${L} t="sim.model" /></th>${cols.map(([t]) => html`<th class="num"><${L} t=${t} /></th>`)}<th class="num"><${L} t="sim.kxySet" /></th><th><${L} t="sim.check" /></th><th></th></tr></thead>
      <tbody>${chips.map((x, i) => html`<tr key=${x.refDes + i}><td><${V} t="sim.refDes" k=${String(i)} /></td>
        <td><${Select} term="sim.model" tkey=${String(i)} value=${ss.chips[i].model} options=${["Block", "Two-resistor"]} label="Chip model" disabled=${locked} onCommit=${v => setSim(["chips", i, "model"], `Chip model of ${x.refDes}`, v || "Block")} /></td>
        ${cols.map(([t, d]) => html`<td class="num"><${V} t=${t} k=${String(i)} d=${d} /></td>`)}
        <td class="num"><${NumField} term="sim.kxySet" tkey=${String(i)} value=${ss.chips[i].kxy} unit="" label="k xy typed" width="4.5em" placeholder="= k z" rule=${{ minEx: 0 }} disabled=${locked} onCommit=${v => setSim(["chips", i, "kxy"], `k xy of ${x.refDes}`, v, false)} /></td>
        <td class=${x.check === "OK" ? "" : "warn"}><${V} t="sim.check" k=${String(i)} />${x.model === "Two-resistor" && isNum(x.rjcCW) && isNum(x.rjbCW) ? html` <span class="hint">(Run Thermal Study builds it in three layers)</span>` : ""}</td>
        <td><button class="small danger" disabled=${locked} onClick=${() => c.editFn(`Remove chip ${x.refDes}`, q => setIn(withSimSetup(q), ["simSetup", "chips"], q.simSetup.chips.filter((_, j) => j !== i)))}>Remove</button></td></tr>`)}</tbody></table><//>
    <//>`;
}

// ---------- 9 Results ----------
export function ResultsTab(c) {
  return html`<${ResultsView} c=${c} />`;
}
function ResultsView({ c }) {
  const runs = c.results?.runs || [];
  const [sel, setSel] = useState(null);
  const run = runs.find(r => r.name === sel) || runs[runs.length - 1];
  if (!run) return html`<p class="hint">No runs yet.</p>`;
  const rc = c.calc.runs[run.name] || {};
  const waiting = runs.some(r => r.studyType === "Thermal" && r.status !== "read");
  const refs = Object.keys(run.parts || {});
  const parts = new Map(c.project.parts.map(x => [x.refDes, x]));
  return html`
    ${waiting && html`<${SwCard} c=${c} id="ReadResults" />`}
    <${Card} title=${`Run ${run.name}`} cls=${rc.pass === "FAIL" || rc.balanceOk === false || rc.health?.level !== "ok" ? "attention" : ""}>
      <${Health} c=${c} run=${run} rc=${rc} />
      ${rc.changes?.length > 0 && html`<p class="warn"><b>Older than the inputs:</b> ${marked(rc.changes.slice(0, 8).join("; "))}${rc.changes.length > 8 ? "; and more" : ""}.</p>`}
      ${rc.balanceOk === false && run.studyType !== "Flow" && html`<p class="warn"><b>Energy balance error above ${BALANCE_LIMIT_PCT} %:</b> the run is not converged, or a heat load is missing. Do not trust its temperatures yet: check the SetupSim report, refine the mesh and run again.</p>`}
      <div class=${`cols ${rc.health?.level === "bad" ? "dim" : ""}`}>
        <div><${Facts} ctx=${c} rows=${[[, "run.pass", run.name], [, "run.worst", run.name], [, "run.worstC", run.name, 1], [, "run.fails", run.name], [, "run.date", run.name, undefined, date(run.date)], [, "run.studyType", run.name], [, "run.study", run.name], [, "run.status", run.name]]} /></div>
        <div><${Facts} ctx=${c} rows=${[[, "run.energyBalance", run.name, 2], [, "run.balanceOk", run.name], [, "run.pcbMin", run.name, 1], [, "run.pcbMax", run.name, 1], [, "run.spreaderMin", run.name, 1], [, "run.spreaderMax", run.name, 1], [, "run.meshGlobal", run.name], [, "run.meshFine", run.name], [, "run.nodes", run.name], [, "run.elements", run.name], [, "run.heatRequested", run.name, 2], [, "run.heatApplied", run.name, 2], [, "run.solverFinished", run.name], [, "run.travels", run.name, 2], [, "run.iterations", run.name], [, "run.cellsTotal", run.name], [, "run.cellsSolid", run.name], [, "run.cellsFluid", run.name]]} /></div>
      </div>
      ${c.learn && html`<${RunReading} c=${c} run=${run} rc=${rc} />`}
      ${refs.length > 0 && html`<${Scroll}><table class="grid"><thead><tr><th><${L} t="part.refDes" /></th><th class="num"><${L} t="part.tjMax" /></th><th class="num"><${L} t="runp.peak" /></th><th class="num"><${L} t="runp.avg" /></th>
        <th class="num"><${L} t="runp.margin" /></th><th><${L} t="runp.status" /></th><th class="num"><${L} t="runp.hand" /></th><th class="num"><${L} t="runp.gap" /></th></tr></thead>
        <tbody>${refs.map(k => { const key = `${run.name}|${k}`; const r = rc.parts?.[k] || {}; return html`<tr>
          <td><${V} t="part.refDes" k=${k} v=${k} /></td><td class="num">${parts.get(k) && isNum(parts.get(k).tjMaxC) ? html`<${V} t="part.tjMax" k=${k} d=${0} />` : ""}</td>
          <td class=${`num ${r.status === "FAIL" ? "warn" : ""}`}><${V} t="runp.peak" k=${key} d=${1} /></td><td class="num"><${V} t="runp.avg" k=${key} d=${1} /></td>
          <td class="num"><${V} t="runp.margin" k=${key} d=${1} /></td><td><${V} t="runp.status" k=${key} /></td>
          <td class="num"><${V} t="runp.hand" k=${key} d=${1} />${r.handFrom === "thermal solution" ? html` <span class="hint">plate</span>` : ""}</td><td class="num"><${V} t="runp.gap" k=${key} d=${1} /></td></tr>`; })}</tbody></table><//>`}
      <${Plot} c=${c} run=${run} />
    <//>
    <${Card} title="All runs">
      <${Scroll}><table class="grid runs"><thead><tr><th><${L} t="run.name" /></th><th><${L} t="run.date" /></th><th><${L} t="run.studyType" /></th><th><${L} t="run.status" /></th><th><${L} t="run.worst" /></th><th class="num"><${L} t="run.worstC" /></th><th class="num"><${L} t="run.energyBalance" /></th><th><${L} t="run.health" /></th><th><${L} t="run.pass" /></th><th></th></tr></thead>
      <tbody>${[...runs].reverse().map(r => html`<tr class=${r.name === run.name ? "sel" : ""}><td><${V} t="run.name" k=${r.name} /></td><td><${V} t="run.date" k=${r.name} v=${date(r.date)} /></td><td><${V} t="run.studyType" k=${r.name} /></td><td><${V} t="run.status" k=${r.name} /></td>
        <td><${V} t="run.worst" k=${r.name} /></td><td class="num"><${V} t="run.worstC" k=${r.name} d=${1} /></td><td class="num"><${V} t="run.energyBalance" k=${r.name} d=${2} /></td><td><span class=${`status ${HEALTH_CLS[c.calc.runs[r.name]?.health?.level] || ""}`}><${V} t="run.health" k=${r.name} /></span></td><td class=${c.calc.runs[r.name]?.health?.level === "bad" ? "dim" : ""}><${V} t="run.pass" k=${r.name} /></td>
        <td><button class="small" aria-pressed=${r.name === run.name} onClick=${() => setSel(r.name)}>Show</button></td></tr>`)}</tbody></table><//>
      <p class="hint">Each run keeps a copy of the inputs it was made with, so an old run always shows its own case. The hand check on tab 5 compares with the run set there.</p>
    <//>`;
}

const HEALTH_CLS = { ok: "ok", check: "watch", bad: "crit" };
// Health badge of a run (Flow plan F1) and the reasons in plain words. Read it before the temperatures.
function Health({ c, run, rc }) {
  const h = rc.health;
  if (!h) return null;
  return html`<div class="health"><p><span class=${`status ${HEALTH_CLS[h.level]}`}><${V} t="run.health" k=${run.name} /></span>
    ${h.level === "ok" ? " Finished, every part meshed, energy balance within limits: read the temperatures." : h.level === "bad" ? " Do not trust the temperatures of this run yet. PASS or FAIL is greyed out." : " Look at these before you trust the temperatures."}</p>
    ${h.reasons.length > 0 && html`<ul class="reasons">${h.reasons.map(x => html`<li>${marked(x.charAt(0).toUpperCase() + x.slice(1) + ".")}</li>`)}</ul>`}</div>`;
}

// Plain reading of a run (Learn mode): the hottest part, how it compares with the hand check, and what to trust.
function RunReading({ c, run, rc }) {
  const w = rc.worst;
  if (!w) return null;
  const key = `${run.name}|${w}`;
  const r = rc.parts[w] || {};
  const gaps = Object.entries(rc.parts || {}).filter(([, x]) => isNum(x.gapC));
  const big = gaps.filter(([, x]) => Math.abs(x.gapC) > 5).map(([k]) => k);
  return html`<div class="reading"><p><${V} nofx t="part.refDes" k=${w} v=${w} /> is the hottest part at <${V} nofx t="runp.peak" k=${key} d=${1} u />${isNum(r.marginC) ? html`, ${r.marginC >= 0 ? html`<${V} nofx t="runp.margin" k=${key} d=${1} u /> below its ${marked("Tj max")}` : html`over its ${marked("Tj max")} by <${V} nofx t="runp.margin" k=${key} v=${-r.marginC} d=${1} u />`}` : ""}.
    ${rc.pass === "PASS" ? " Every part with a Tj max is below it in this run." : rc.pass === "FAIL" ? ` ${plural(rc.fails, "part")} over the limit: fix the design (tab 7 spreader, copper, air) or check the inputs, then run again.` : ""}</p>
    ${gaps.length > 0 && html`<p>${big.length ? marked(`The hand check and the simulation differ by more than 5 C for ${big.join(", ")}. Look for an input that differs between them: h, emissivity, the PCB k, a body left out, or a chip block k.`) : "The hand check and the simulation agree within 5 C for every part: you can trust both."}</p>`}
  </div>`;
}

function Plot({ c, run }) {
  const [url, setUrl] = useState(null);
  const rel = run.plotFile ? `${c.outFolder || "CAD"}/results/${run.plotFile.split(/[\\/]/).pop()}` : null;
  useEffect(() => {
    let live = true, made = null;
    setUrl(null);
    if (rel && c.imageUrl) c.imageUrl(rel).then(u => { made = u; if (live) setUrl(u); });
    return () => { live = false; if (made) URL.revokeObjectURL(made); };
  }, [rel, c.imageUrl]);
  if (!run.plotFile) return null;
  return html`<figure class="plot">${url ? html`<img src=${url} alt=${"Temperature plot of " + run.name} />` : html`<p class="hint">Plot picture: <${V} t="run.plot" k=${run.name} />${c.imageUrl ? " (not found in the project folder)" : " (open the project from its folder to see it here)"}</p>`}</figure>`;
}

// ---------- 10 Changes ----------
export function ChangesTab(c) {
  const { project: p, cad, now, calc, locked } = c;
  const list = cad?.changes || [];
  if (!p) return html`<p class="hint">Open a project to see its change records.</p>`;
  const ab = cad?.asBuilt;
  const diffs = (ab?.bodies || []).filter(x => calc.bodies.some(b => b.key === x.key));
  const asBuilt = diffs.length > 0 && html`<${Card} title="Sizes changed in SolidWorks" cls="attention">
    <p>${marked(`Sync CAD (${date(ab.date)}) found ${plural(diffs.length, "body", "bodies")} whose box in the assembly differs from the page. Take the SolidWorks sizes to make the page match, or run Build Board to put the page's sizes back.`)}</p>
    <${Scroll}><table class="grid"><thead><tr><th><${L} t="body.name" /></th>${["body.xMin", "body.yMin", "body.zMin", "body.l", "body.w", "body.h"].map(t => html`<th class="num"><${L} t=${t} /></th>`)}<th></th></tr></thead>
    <tbody>${diffs.map(x => { const b = calc.bodies.find(r => r.key === x.key); return html`<tr><td><${V} t="body.name" k=${x.key} v=${shortName(b.name)} /></td>
      ${["xMin", "yMin", "zMin", "l", "w", "h"].map(f => html`<td class="num">${Math.abs(x.box[f] - b.box[f]) > 1e-4 ? html`<span class="hint">${(+b.box[f]).toFixed(2)} to </span><b><span class="hint">${x.box[f].toFixed(2)}</span></b>` : html`<span class="hint">same</span>`}</td>`)}
      <td><button class="small" disabled=${locked} onClick=${() => c.editFn(`Take the SolidWorks size of ${shortName(b.name)}`, q => takeAsBuilt(q, calc, x.key, x.box))}>Take these</button></td></tr>`; })}</tbody></table><//>
    <p><button class="primary" disabled=${locked} onClick=${() => c.editFn(`Take ${diffs.length} SolidWorks sizes`, q => diffs.reduce((a, x) => takeAsBuilt(a, calc, x.key, x.box), q))}>Take all ${diffs.length}</button> <span class="hint">They become typed values on tab 3 (Edit row), so a new import does not undo them. Save after.</span></p>
  <//>`;
  if (!list.length) return html`${asBuilt}<p class="hint">${marked("No change records. Import STEP, Build Board, Build Spreader and Sync CAD write them when something moves, changes size or a check fails.")}</p>`;
  const open = unticked(p, cad);
  const tickAll = () => c.editFn(`Tick ${open} changes`, q => { const checks = { ...(q.checks || {}) }; for (const x of list) if (!checks[x.id]) checks[x.id] = { at: now() }; return { ...q, checks }; });
  const cols = ["change.date", "change.macro", "change.body", "change.change", "change.field"];
  return html`
    ${asBuilt}
    <div class="toolbar"><span>${open ? `${open} of ${list.length} not ticked` : `All ${list.length} ticked`}</span>${open > 0 && html`<button onClick=${tickAll}>Tick all</button>`}
      ${open > 0 && html`<span class="hint">Tick a change once you know why it happened. Unticked changes lock tabs 7 and 8.</span>`}</div>
    <${Scroll}><table class="grid"><thead><tr><th><${L} t="change.checked" /></th>${cols.map(t => html`<th><${L} t=${t} /></th>`)}<th class="num"><${L} t="change.old" /></th><th class="num"><${L} t="change.new" /></th></tr></thead>
    <tbody>${list.map(x => html`<tr class=${p.checks?.[x.id] ? "dim" : ""}>
      <td><${Check} term="change.checked" tkey=${x.id} value=${!!p.checks?.[x.id]} label=${"Checked " + x.id} onCommit=${v => c.edit(v ? "Tick change" : "Untick change", ["checks", x.id], v ? { at: now() } : undefined)} /></td>
      <td><${V} t="change.date" k=${x.id} v=${date(x.date)} /></td><td><${V} t="change.macro" k=${x.id} /></td><td><${V} t="change.body" k=${x.id} v=${shortName(x.name)} /></td><td><${V} t="change.change" k=${x.id} /></td><td><${V} t="change.field" k=${x.id} /></td>
      <td class="num"><${V} t="change.old" k=${x.id} /></td><td class="num"><${V} t="change.new" k=${x.id} /></td></tr>`)}</tbody></table><//>`;
}
