// Checks the .bas files in src/ for the VBA rules of CLAUDE.md that have cost a failed run before:
//   node tools/lint_vba.mjs [file.bas ...]
// - ASCII only; first line Attribute VB_Name = "<file name>"
// - no reserved word or built-in function name as a variable, constant or parameter
// - every parameter says ByVal or ByRef (no implicit ByRef)
// - no name declared twice in one procedure (Dim, Const, parameters)
// - no procedure declared Public in two modules (Update Macros puts all modules in one project; implicit ones stay module-local)
// - every procedure called with a name that looks like one of ours (PT_, PJ_ or defined in the module) exists
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const srcDir = path.join(here, "..", "..", "src");
const files = process.argv.slice(2).length ? process.argv.slice(2) : fs.readdirSync(srcDir).filter(f => f.endsWith(".bas")).map(f => path.join(srcDir, f));
const RESERVED = new Set(["imp", "eqv", "mid", "len", "name", "date", "error", "left", "right", "val", "str", "type", "get", "set", "open", "close",
  "line", "input", "print", "write", "string", "time", "now", "format", "dir", "kill", "end", "exit", "next", "loop", "step", "to", "is", "like", "mod",
  "and", "or", "not", "xor", "let", "call", "case", "select", "with", "new", "me", "nothing", "empty", "null", "true", "false", "sub", "function",
  "property", "dim", "redim", "const", "public", "private", "static", "global", "option", "on", "goto", "gosub", "return", "resume", "stop", "each",
  "in", "for", "while", "wend", "do", "until", "if", "then", "else", "elseif", "byval", "byref", "optional", "paramarray", "as", "integer", "long",
  "double", "single", "boolean", "variant", "object", "currency", "byte", "array", "erase", "lbound", "ubound", "abs", "int", "fix", "sgn", "sqr",
  "chr", "asc", "trim", "ltrim", "rtrim", "ucase", "lcase", "space", "instr", "replace", "split", "join", "round", "cdbl", "clng", "cstr", "cint",
  "isnumeric", "isempty", "isnull", "isobject", "isarray", "environ", "shell", "seek", "lock", "unlock", "width", "binary", "random", "append",
  "output", "access", "shared", "spc", "tab", "rnd", "randomize", "timer", "day", "month", "year", "hour", "minute", "second", "weekday", "beep",
  "doevents", "debug", "err", "attribute", "enum", "event", "friend", "implements", "raiseevent", "typeof", "addressof", "declare", "lib", "alias",
  "decimal", "date", "vbcrlf"]);

let problems = 0;
const say = (f, ln, msg) => { problems++; console.log(`${path.basename(f)}:${ln}: ${msg}`); };
const publicNames = new Map();
const allProcs = new Set();
const moduleNames = new Set();
const parsed = [];

for (const f of files) {
  const raw = fs.readFileSync(f);
  const text = raw.toString("latin1");
  for (let i = 0; i < raw.length; i++) if (raw[i] > 127) { say(f, text.slice(0, i).split("\n").length, "non-ASCII character"); break; }
  const lines = text.replace(/\r\n/g, "\n").split("\n");
  const mod = path.basename(f, ".bas");
  if (lines[0] !== `Attribute VB_Name = "${mod}"`) say(f, 1, `first line must be Attribute VB_Name = "${mod}"`);
  // join continuation lines, keep the first line number
  const logical = [];
  for (let i = 0; i < lines.length; i++) {
    let s = lines[i], n = i + 1;
    while (/\s_\s*$/.test(stripComment(s)) && i + 1 < lines.length) { s = stripComment(s).replace(/\s_\s*$/, " ") + lines[++i]; }
    logical.push([n, stripComment(s)]);
  }
  const procs = [];
  let cur = null;
  for (const [n, s] of logical) {
    const m = s.match(/^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*\((.*)\)/i);
    if (m) {
      cur = { name: m[4], scope: (m[1] || "implicit").trim().toLowerCase(), start: n, names: new Map(), body: [] };
      procs.push(cur);
      allProcs.add(m[4].toLowerCase());
      if (cur.scope === "public") {
        const key = m[4].toLowerCase();
        if (publicNames.has(key) && publicNames.get(key) !== mod) say(f, n, `public ${m[4]} is also defined in ${publicNames.get(key)}`);
        publicNames.set(key, mod);
      }
      for (const p of splitTop(m[5])) {
        const pm = p.trim().match(/^(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?(\w+)(\(\))?/i);
        if (!pm || !pm[4]) continue;
        if (!pm[2] && !pm[3]) say(f, n, `parameter ${pm[4]} of ${m[4]} has no ByVal or ByRef`);
        declare(f, n, cur, pm[4]);
      }
      continue;
    }
    if (/^\s*End\s+(Sub|Function|Property)\b/i.test(s)) { cur = null; continue; }
    const dm = s.match(/^\s*(?:(?:Private|Public|Global)\s+)?(Dim|Static|Const|Private|Public|Global)\s+(?!Sub\b|Function\b|Property\b|Enum\b|Type\b|Declare\b)(.*)$/i);
    if (dm) {
      for (const part of splitTop(dm[2])) {
        const vm = part.trim().match(/^(?:WithEvents\s+)?(\w+)/i);
        if (!vm) continue;
        if (cur) declare(f, n, cur, vm[1]);
        else { moduleNames.add(vm[1].toLowerCase()); if (RESERVED.has(vm[1].toLowerCase())) say(f, n, `module variable ${vm[1]} is a reserved word or built-in name`); }
      }
      // more statements on the same line after ":" are rare; ignore
    }
    if (cur) cur.body.push([n, s]);
  }
  parsed.push({ f, mod, procs });
}

// calls to our own helpers that do not exist anywhere
for (const { f, procs } of parsed) {
  for (const p of procs) for (const [n, s] of p.body) {
    const noStr = s.replace(/"[^"]*"/g, '""');
    for (const m of noStr.matchAll(/\b(P[TJ]_\w+)\b/g)) {
      const k = m[1].toLowerCase();
      if (!allProcs.has(k) && !isVar(k, procs, p)) say(f, n, `${m[1]} is not defined in any module`);
    }
  }
}

function isVar(k, procs, p) {
  return p.names.has(k) || moduleNames.has(k);
}
function declare(f, n, proc, nm) {
  const k = nm.toLowerCase();
  if (RESERVED.has(k)) say(f, n, `${nm} in ${proc.name} is a reserved word or built-in name`);
  if (proc.names.has(k)) say(f, n, `${nm} declared twice in ${proc.name} (first on line ${proc.names.get(k)})`);
  else proc.names.set(k, n);
}
function stripComment(s) {
  let inStr = false;
  for (let i = 0; i < s.length; i++) {
    if (s[i] === '"') inStr = !inStr;
    else if (s[i] === "'" && !inStr) return s.slice(0, i);
  }
  return s;
}
function splitTop(s) {
  const out = []; let depth = 0, inStr = false, curS = "";
  for (const ch of s) {
    if (ch === '"') inStr = !inStr;
    if (!inStr && ch === "(") depth++;
    if (!inStr && ch === ")") depth--;
    if (!inStr && depth === 0 && ch === ",") { out.push(curS); curS = ""; } else curS += ch;
  }
  if (curS.trim()) out.push(curS);
  return out;
}

console.log(problems ? `VBA lint: ${problems} problems.` : `VBA lint: ${files.length} files clean.`);
if (problems) process.exit(1);
