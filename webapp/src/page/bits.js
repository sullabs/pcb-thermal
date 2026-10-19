// Small view pieces shared by the tab views (views.js for tabs 0 to 5, steps.js for 6 to 10, report.js).

import { html } from "./ui.js";
import { V, L } from "./explain.js";
import { isNum } from "../calc/util.js";

export const Scroll = ({ children }) => html`<div class="scroll">${children}</div>`;

// A definition list of [label term, value term, key, decimals, value] rows; rows whose value is blank are left out.
export function Facts({ ctx, rows }) {
  return html`<dl class="kv">${rows.filter(Boolean).filter(([, t, k, , v]) => { const x = v !== undefined ? v : ctx.value(t, k); return x !== undefined && x !== ""; })
    .map(([lt, t, k = "", d, v]) => html`<dt><${L} t=${lt || t} k=${k} /></dt><dd><${V} t=${t} k=${k} d=${d} v=${v} u /></dd>`)}</dl>`;
}

export const date = s => (s ? s.replace("T", " ") : s);
export const shortName = n => (n || "").split("/").pop().replace(/-\d+$/, "");
export const pctV = (t, k, x) => html`<${V} nofx t=${t} k=${k} v=${isNum(x) ? `${Math.round(x * 100)} %` : undefined} />`;
export const plural = (n, one, many = one + "s") => `${n} ${n === 1 ? one : many}`;
