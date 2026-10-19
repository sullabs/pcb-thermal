// Where a value is typed, for the "Change it" button of the Explain panel. Pure, tested by tools/test_w5.mjs.
//
//   editPlace(ctx, t, k) -> { tab, t, k, body, why } or null
//     tab: the step tab with the field; t, k: the field's term and key (data-term, data-key on the page);
//     body: a body key whose Edit row on tab 3 must be open first; why: the button's hover text.
//   findField(root, t, k) -> the input, select or textarea of that field, or null

import { TERMS } from "../terms/catalogue.js";

// Input terms by prefix: the tab where they are typed. The term is the field's own term.
const INPUT_TAB = [
  ["board.", 0], ["limit.", 0],
  ["part.", 1], ["load.", 1],
  ["set.", 2],
  ["pcbov.", 3], ["stack.", 3], ["body.include", 3], ["body.refDes", 3], ["bodyset.", 3],
  ["surf.spreader", 8], ["simset.kPartsWithoutHeat", 5], ["simset.", 8],
  ["surf.", 5], ["hand.", 5],
  ["plate.", 7], ["pchip.", 7], ["sol.", 7], ["solp.", 7],
  ["sim.refDes", 8], ["sim.model", 8], ["sim.kxySet", 8],
  ["change.checked", 10],
];

export function editPlace(ctx, t, k = "") {
  // Values from SolidWorks that the page lets you overwrite: the PCB size and each body's name and box.
  const pcb = /^pcb\.(l|w|h)$/.exec(t);
  if (pcb) return { tab: 3, t: "pcbov." + pcb[1], k: "", why: "Tab 3, PCB size: type a value to overwrite the imported one" };
  const body = /^body\.(name|l|w|h|xMin|yMin|zMin)$/.exec(t);
  if (body && k) return { tab: 3, t: "bodyset." + body[1], k, body: k, why: "Tab 3, Edit row of this body: type a value to overwrite the imported one" };
  const part = /^bbp\.(l|w|h)$/.exec(t);
  if (part) {
    const b = ctx?.calc?.bodies?.find(x => x.refDes === k && x.include) || ctx?.calc?.bodies?.find(x => x.refDes === k);
    return b ? { tab: 3, t: "bodyset." + part[1], k: b.key, body: b.key, why: `Tab 3, Edit row of the body linked to ${k}` }
      : { tab: 3, why: `Tab 3: link a body to ${k}` };
  }
  if (t === "rail.name") return { tab: 1, t: "load.rail", why: "Tab 1, Rail column" };
  if (TERMS[t]?.kind !== "input") return null;
  const hit = INPUT_TAB.find(([p]) => t.startsWith(p));
  if (!hit) return null;
  return { tab: hit[1], t, k, body: t.startsWith("bodyset.") && k ? k : undefined, why: `Tab ${hit[1]}: where this value is typed` };
}

const FIELD = "input:not([type=hidden]), select, textarea";
export function findField(root, t, k = "") {
  const esc = s => String(s).replace(/["\\]/g, "\\$&");
  const pick = sel => {
    for (const el of root.querySelectorAll(sel)) {
      const f = el.matches(FIELD) ? el : el.querySelector(FIELD);
      if (f) return f;
    }
    return null;
  };
  return pick(`[data-term="${esc(t)}"][data-key="${esc(k)}"]`) || (k ? null : pick(`[data-term="${esc(t)}"]`));
}
