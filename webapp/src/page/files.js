// Project folder access in Chrome and Edge (File System Access), with open and download as the fallback for
// other browsers. Each project's files carry the user's prefix: AB_project.json, AB_cad.json, AB_results.json
// (pre = "AB_"; pre = "" for the old names). The page writes only the project file (and its daily backup), and the
// empty cad and results files of a new project. After that the macros own those two.

export const folderSupported = typeof window !== "undefined" && "showDirectoryPicker" in window;

// ---------- remember the folder (IndexedDB; a folder handle cannot go in localStorage) ----------
function idb() {
  return new Promise((ok, fail) => {
    const r = indexedDB.open("pcb-thermal", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("kv");
    r.onsuccess = () => ok(r.result);
    r.onerror = () => fail(r.error);
  });
}
async function kv(mode, fn) {
  const db = await idb();
  return new Promise((ok, fail) => {
    const tx = db.transaction("kv", mode);
    const req = fn(tx.objectStore("kv"));
    tx.oncomplete = () => ok(req?.result);
    tx.onerror = () => fail(tx.error);
  });
}
export async function rememberFolder(handle) {
  try { await kv("readwrite", s => s.put(handle, "folder")); } catch { /* private window: not remembered */ }
}
export async function lastFolder() {
  try { return (await kv("readonly", s => s.get("folder"))) || null; } catch { return null; }
}

export async function pickFolder() {
  const handle = await window.showDirectoryPicker({ id: "pcb-thermal", mode: "readwrite" });
  await rememberFolder(handle);
  return handle;
}

// Needs a click when the browser asks again (each new browser session).
export async function allowed(handle, ask) {
  const opts = { mode: "readwrite" };
  if ((await handle.queryPermission(opts)) === "granted") return true;
  return ask ? (await handle.requestPermission(opts)) === "granted" : false;
}

// ---------- names ----------
// pre is "" or the prefix with its "_". The prefix the user types: letters, digits and "-", up to 16.
export const PREFIX_RE = /^[A-Za-z0-9][A-Za-z0-9-]{0,15}$/;
export const names = pre => ({ project: pre + "project.json", cad: pre + "cad.json", results: pre + "results.json" });
// "AB_project.json" -> { pre: "AB_", kind: "project" }; null for other names.
export function nameParts(fileName) {
  const m = /^(?:(.+)_)?(project|cad|results)\.json$/i.exec(fileName);
  return m ? { pre: m[1] ? m[1] + "_" : "", kind: m[2].toLowerCase() } : null;
}
// The prefixes of the project files in the folder, sorted ("" for project.json).
export async function listProjects(dir) {
  const out = [];
  for await (const [name, h] of dir.entries()) {
    const n = h.kind === "file" && nameParts(name);
    if (n && n.kind === "project") out.push(n.pre);
  }
  return out.sort();
}

// ---------- files ----------
async function fileIn(dir, name) {
  try { return await (await dir.getFileHandle(name)).getFile(); } catch (e) {
    if (e.name === "NotFoundError" || e.name === "TypeMismatchError") return null;
    throw e;
  }
}

export async function stamp(dir, name) {
  const f = await fileIn(dir, name);
  return f ? f.lastModified : 0;
}

// { data, stamp } or null when the file is not there. A file that does not parse throws, with the file name.
export async function readJson(dir, name) {
  const f = await fileIn(dir, name);
  if (!f) return null;
  const text = await f.text();
  try { return { data: JSON.parse(text), stamp: f.lastModified }; } catch (e) {
    throw new Error(`${name} is not valid JSON (${e.message}).`);
  }
}

export async function readText(dir, path) {
  let d = dir;
  const parts = path.split("/");
  try {
    for (const p of parts.slice(0, -1)) d = await d.getDirectoryHandle(p);
  } catch { return undefined; }
  const f = await fileIn(d, parts[parts.length - 1]);
  return f ? f.text() : undefined;
}

// Names of the files in a folder below the project folder ("CAD"); [] when the folder is not there.
export async function listNames(dir, path) {
  let d = dir;
  try { for (const p of path.split("/").filter(Boolean)) d = await d.getDirectoryHandle(p); } catch { return []; }
  const out = [];
  for await (const [name, h] of d.entries()) if (h.kind === "file") out.push(name);
  return out;
}

async function write(dir, name, text) {
  const h = await dir.getFileHandle(name, { create: true });
  const w = await h.createWritable(); // writes a temporary file, swapped in on close: never half written
  await w.write(text);
  await w.close();
  return (await h.getFile()).lastModified;
}

// Before the first save of the day: copy the project file on disk to backup\<pre>project_yyyymmdd.json.
async function backupOnce(dir, day, pre) {
  const cur = await fileIn(dir, names(pre).project);
  if (!cur) return null;
  const bdir = await dir.getDirectoryHandle("backup", { create: true });
  const name = `${pre}project_${day}.json`;
  if (await fileIn(bdir, name)) return null;
  await write(bdir, name, await cur.text());
  return "backup\\" + name;
}

export async function saveProject(dir, project, day, pre) {
  const backup = await backupOnce(dir, day, pre);
  const st = await write(dir, names(pre).project,JSON.stringify(project, null, 2) + "\n");
  return { stamp: st, backup };
}

// Only for a new project: the cad and results files are written once, never over a file a macro wrote.
export async function writeIfMissing(dir, name, data) {
  if (await fileIn(dir, name)) return false;
  await write(dir, name, JSON.stringify(data, null, 2) + "\n");
  return true;
}

// The CAD folder of a new project (the macros' default output folder).
export async function makeFolder(dir, name) {
  await dir.getDirectoryHandle(name, { create: true });
}

export async function listBackups(dir, pre) {
  let bdir;
  try { bdir = await dir.getDirectoryHandle("backup"); } catch { return []; }
  const out = [];
  for await (const [name, h] of bdir.entries()) if (h.kind === "file" && name.startsWith(pre) && /^project_\d{8}\.json$/.test(name.slice(pre.length))) out.push(name);
  return out.sort().reverse();
}
export async function readBackup(dir, name) {
  return readJson(await dir.getDirectoryHandle("backup"), name);
}

// ---------- fallback: open by hand, save as a download ----------
export function download(name, text, type = "application/json") {
  const url = URL.createObjectURL(new Blob([text], { type }));
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
