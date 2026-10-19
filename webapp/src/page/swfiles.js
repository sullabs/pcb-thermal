// The SolidWorks files Build Board makes for this project, and whether they are in the folder. Pure, tested by
// tools/test_w5.mjs. The names follow the macros (PTJson.bas PJ_OutputFolder, PJ_Prefix, PJ_AsmPath; BuildBoard.bas):
//   <output folder>\<prefix>Thermal.sldasm, <prefix>PCB.sldprt, <prefix>P_<body key>.sldprt
//
//   swFiles(project, cad, calc) -> { folder, rows }
//     folder: the output folder relative to the project folder ("CAD"), or null when it is outside it
//     rows: [{ kind: "assembly" | "pcb" | "part", key, name, refDes, included, status, reason, file }]
//   withFound(rows, names, prefix) -> rows with found (true or false) from a folder listing, plus a row for each
//     part file of this prefix that no row names (kind "extra": built before, then left out)

export const safeFile = s => String(s ?? "").replace(/[^A-Za-z0-9_-]/g, "_").slice(0, 60);

export function swPrefix(project) {
  return safeFile(project?.settings?.filePrefix || safeFile(project?.board?.name || "Board") + "_");
}

export function swFolder(project) {
  const s = (project?.settings?.outputFolder || "").trim();
  if (!s) return "CAD";
  if (s.includes(":") || s.startsWith("\\\\")) return null;
  return s.replace(/\\/g, "/").replace(/\/+$/, "");
}

const base = p => String(p || "").split(/[\\/]/).pop();

export function swFiles(project, cad, calc) {
  const pre = swPrefix(project), b = cad?.build;
  const asmSet = (project?.settings?.asmName || "").trim();
  let asm = b?.assembly ? base(b.assembly) : asmSet ? base(asmSet) : pre + "Thermal";
  if (!/\.sldasm$/i.test(asm)) asm += ".sldasm";
  const rows = [
    { kind: "assembly", key: "", name: "Thermal assembly", status: b ? "built" : undefined, file: asm, elsewhere: !!asmSet && /[\\:]/.test(asmSet) },
    { kind: "pcb", key: "pcb", name: "PCB", status: b ? "built" : undefined, file: pre + "PCB.sldprt" },
  ];
  const rec = Object.fromEntries((b?.bodies || []).filter(x => x.key).map(x => [x.key, x]));
  const bodies = calc?.bodies || [];
  const keys = [...bodies.filter(x => x.include || rec[x.key]).map(x => x.key), ...Object.keys(rec).filter(k => !bodies.some(x => x.key === k))];
  for (const key of keys) {
    const body = bodies.find(x => x.key === key), r = rec[key];
    rows.push({ kind: "part", key, name: body?.name || r?.name || key, refDes: body?.refDes, included: !!body?.include,
      status: r?.status || (body?.include ? "not built yet" : undefined), reason: r?.reason,
      file: r?.file ? base(r.file) : pre + "P_" + safeFile(key) + ".sldprt" });
  }
  return { folder: swFolder(project), prefix: pre, rows };
}

export function withFound(rows, names, prefix) {
  const have = new Set(names.map(n => n.toLowerCase()));
  const out = rows.map(r => ({ ...r, found: r.elsewhere ? undefined : have.has(r.file.toLowerCase()) }));
  const named = new Set(out.map(r => r.file.toLowerCase()));
  const partRe = new RegExp(`^${prefix.replace(/[-]/g, "\\-")}P_.+\\.sldprt$`, "i");
  for (const n of names) if (partRe.test(n) && !named.has(n.toLowerCase())) out.push({ kind: "extra", key: "", name: n.slice(prefix.length + 2, -7), file: n, found: true });
  return out;
}
