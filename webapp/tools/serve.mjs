// Development only: serve the repository on http://localhost:8173 so the built page and the example files can be
// opened in a browser for checks. The user opens dist/PCB_Thermal.html straight from disk instead.
//   node tools/serve.mjs
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
const types = { ".html": "text/html; charset=utf-8", ".json": "application/json", ".js": "text/javascript", ".css": "text/css" };
const port = Number(process.env.PORT) || 8173;

http.createServer((req, res) => {
  const p = path.resolve(root, "." + decodeURIComponent(new URL(req.url, "http://x").pathname));
  if (!p.startsWith(root) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); res.end("not found"); return; }
  res.writeHead(200, { "content-type": types[path.extname(p)] || "application/octet-stream", "cache-control": "no-store" });
  fs.createReadStream(p).pipe(res);
}).listen(port, "127.0.0.1", () => console.log(`http://localhost:${port}/webapp/dist/PCB_Thermal.html`));
