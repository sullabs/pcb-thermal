// Build the page: one HTML file with every library inside it, dist/PCB_Thermal.html.
//   node tools/build.mjs
// The schema validators are compiled here (Ajv standalone), so the page checks project.json before it saves
// without compiling code at run time.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";
import Ajv2020 from "ajv/dist/2020.js";
import standalone from "ajv/dist/standalone/index.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(here, "..");
const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));

export async function build({ out = path.join(root, "dist", "PCB_Thermal.html"), quiet = false } = {}) {
  const ajv = new Ajv2020({ allErrors: true, strict: true, allowUnionTypes: true, code: { source: true, esm: true } });
  for (const f of ["common", "project", "cad", "results"]) {
    ajv.addSchema(JSON.parse(fs.readFileSync(path.join(root, "schema", f + ".schema.json"), "utf8")), f + ".schema.json");
  }
  const validators = standalone(ajv, { validateProject: "project.schema.json", validateCad: "cad.schema.json", validateResults: "results.schema.json" });

  // The example study (Help > example, read only). Board data stays in example/, out of src/.
  const ex = path.join(root, "..", "example", "RaspberryPi4");
  const example = JSON.stringify(Object.fromEntries(["project", "cad", "results"].map(n => [n, JSON.parse(fs.readFileSync(path.join(ex, n + ".json"), "utf8"))])));
  // The SolidWorks button icons for the SolidWorks cards, from the macros' icon folder (src/icons/*_32.bmp).
  const iconDir = path.join(root, "..", "src", "icons");
  const icons = JSON.stringify(Object.fromEntries(fs.readdirSync(iconDir).filter(f => /_32.bmp$/i.test(f))
    .map(f => [f.replace(/_32.bmp$/i, ""), "data:image/bmp;base64," + fs.readFileSync(path.join(iconDir, f)).toString("base64")])));
  const virtual = {
    name: "virtual",
    setup(b) {
      b.onResolve({ filter: /^virtual:(validators|example|icons)$/ }, a => ({ path: a.path.slice(8), namespace: "virtual" }));
      b.onLoad({ filter: /.*/, namespace: "virtual" }, a => ({ contents: a.path === "example" ? example : a.path === "icons" ? icons : validators, loader: a.path === "validators" ? "js" : "json", resolveDir: root }));
    },
  };
  const res = await esbuild.build({
    entryPoints: [path.join(root, "src", "page", "main.js")],
    bundle: true, format: "iife", minify: true, write: false, target: ["chrome110", "edge110", "firefox115"],
    legalComments: "inline", charset: "utf8", plugins: [virtual], logLevel: "warning",
  });
  const js = res.outputFiles[0].text.replace(/<\/script/gi, "<\\/script").replace(/<!--/g, "<\\!--");
  const css = fs.readFileSync(path.join(root, "src", "page", "page.css"), "utf8");
  const built = new Date().toISOString().slice(0, 10);
  const page = fs.readFileSync(path.join(root, "src", "page", "index.html"), "utf8")
    .replace("__VERSION__", pkg.version).replace("__BUILT__", built)
    .replace("__STYLE__", () => css).replace("__SCRIPT__", () => js);
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, page, "utf8");
  if (!quiet) console.log(`Wrote ${path.relative(process.cwd(), out)} (${(page.length / 1024).toFixed(0)} KB)`);
  return { out, size: page.length, page };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  build().catch(e => { console.error(e.message); process.exit(1); });
}
