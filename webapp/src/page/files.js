// Project folder access in Chrome and Edge (File System Access), with open and download as the fallback for
// other browsers. The page writes only project.json (and its daily backup). cad.json and results.json are read only.

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

async function write(dir, name, text) {
  const h = await dir.getFileHandle(name, { create: true });
  const w = await h.createWritable(); // writes a temporary file, swapped in on close: never half written
  await w.write(text);
  await w.close();
  return (await h.getFile()).lastModified;
}

// Before the first save of the day: copy the project.json on disk to backup\project_yyyymmdd.json.
async function backupOnce(dir, day) {
  const cur = await fileIn(dir, "project.json");
  if (!cur) return null;
  const bdir = await dir.getDirectoryHandle("backup", { create: true });
  const name = `project_${day}.json`;
  if (await fileIn(bdir, name)) return null;
  await write(bdir, name, await cur.text());
  return "backup\\" + name;
}

export async function saveProject(dir, project, day) {
  const backup = await backupOnce(dir, day);
  const st = await write(dir, "project.json", JSON.stringify(project, null, 2) + "\n");
  return { stamp: st, backup };
}

// Only for a new project made from an old workbook: cad.json and results.json are written once, never over a
// file a macro wrote.
export async function writeIfMissing(dir, name, data) {
  if (await fileIn(dir, name)) return false;
  await write(dir, name, JSON.stringify(data, null, 2) + "\n");
  return true;
}

export async function listBackups(dir) {
  let bdir;
  try { bdir = await dir.getDirectoryHandle("backup"); } catch { return []; }
  const out = [];
  for await (const [name, h] of bdir.entries()) if (h.kind === "file" && /^project_\d{8}\.json$/.test(name)) out.push(name);
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
