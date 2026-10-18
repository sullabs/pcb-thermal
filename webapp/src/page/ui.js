// Small input and display widgets. Inputs keep what the user types locally and store only a valid value, so a
// half typed number never reaches the file. A wrong entry is marked at once with the reason.

import { h } from "preact";
import { useState, useEffect } from "preact/hooks";
import htm from "htm";
import { parseNumber } from "./model.js";
import { isNum } from "../calc/util.js";

export const html = htm.bind(h);

export function fmt(x, d = 2) {
  if (!isNum(x)) return "";
  const a = Math.abs(x);
  if (a !== 0 && (a < 10 ** -d || a >= 1e7)) return x.toExponential(2);
  return x.toFixed(d);
}
export const pct = x => (isNum(x) ? (x * 100).toFixed(1) + " %" : "");

// Number with a fixed unit. onCommit(value, key) gets undefined for a blank entry.
export function NumField({ value, unit, label, onCommit, rule = {}, disabled, title, accepts, width = "6em", id }) {
  const shown = v => (isNum(v) ? String(+v.toPrecision(12)) : "");
  const [text, setText] = useState(shown(value));
  const [error, setError] = useState(null);
  useEffect(() => {
    const p = parseNumber(text, accepts || unit, rule);
    if (p.error || p.value !== value) { setText(shown(value)); setError(null); }
  }, [value]);
  const onInput = e => {
    const t = e.currentTarget.value;
    setText(t);
    const p = parseNumber(t, accepts || unit, rule);
    setError(p.error || null);
    if (!p.error && p.value !== value) onCommit(p.value);
  };
  return html`<span class="field ${error ? "bad" : ""}">
    <input id=${id} type="text" inputmode="decimal" value=${text} onInput=${onInput} disabled=${disabled}
      aria-label=${label} aria-invalid=${!!error} title=${error ? label + ": " + error : title || label} style=${{ width }} />
    ${unit && html`<span class="unit">${unitText(unit)}</span>`}
    ${error && html`<span class="err" role="alert">${error}</span>`}
  </span>`;
}
const UNIT = { copperMm: "mm", dC: "C", CW: "C/W", W: "W", C: "C", mm: "mm", V: "V", A: "A", ohm: "ohm", WmK: "W/(m K)", frac: "" };
const unitText = u => UNIT[u] ?? u;

export function TextField({ value, label, onCommit, disabled, required, width = "12em", multiline, title }) {
  const [text, setText] = useState(value ?? "");
  useEffect(() => { if ((value ?? "") !== text.trim()) setText(value ?? ""); }, [value]);
  const bad = required && !text.trim();
  const onInput = e => {
    const t = e.currentTarget.value;
    setText(t);
    if (!(required && !t.trim())) onCommit(t.trim() || undefined);
  };
  const props = { value: text, onInput, disabled, "aria-label": label, title: bad ? label + ": required" : title || label, style: { width } };
  return html`<span class="field ${bad ? "bad" : ""}">
    ${multiline ? html`<textarea rows="5" ...${props}></textarea>` : html`<input type="text" ...${props} />`}
    ${bad && html`<span class="err" role="alert">required</span>`}
  </span>`;
}

export function Select({ value, options, label, onCommit, disabled, blankText = "", title }) {
  return html`<select aria-label=${label} title=${title || label} disabled=${disabled} value=${value ?? ""}
    onChange=${e => onCommit(e.currentTarget.value || undefined)}>
    <option value="">${blankText}</option>
    ${options.map(o => html`<option value=${o} selected=${o === value}>${o}</option>`)}
  </select>`;
}

export function Check({ value, label, onCommit, disabled, title }) {
  return html`<input type="checkbox" aria-label=${label} title=${title || label} checked=${!!value} disabled=${disabled}
    onChange=${e => onCommit(e.currentTarget.checked)} />`;
}

export function Status({ word }) {
  if (!word) return null;
  const cls = { CRITICAL: "crit", FAIL: "crit", Watch: "watch", Tight: "watch", Minor: "minor", OK: "ok", PASS: "ok" }[word] || "plain";
  return html`<span class="status ${cls}">${word}</span>`;
}

export function Card({ title, children, cls = "" }) {
  return html`<section class="card ${cls}">${title && html`<h3>${title}</h3>`}${children}</section>`;
}

export function KV({ rows }) {
  return html`<dl class="kv">${rows.filter(r => r && r[1] !== undefined && r[1] !== "").map(([k, v]) => html`<dt>${k}</dt><dd>${v}</dd>`)}</dl>`;
}
