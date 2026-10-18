# PCB thermal page

The page that replaces the workbook as the user interface. Design and milestones: `docs/WEBAPP_DESIGN.md`.

## Folders

| Folder | What |
|---|---|
| `schema/` | JSON Schemas of the three project files: `project.json` (page writes), `cad.json` and `results.json` (macros write). `common.schema.json` holds shared parts. |
| `src/calc/` | The calculation module: register, triage, rail budget, PCB stack, body checks, bare board network, spreader plates, thermal solution, sim k z, and `forSolidWorks`. Pure functions on the three files; the page and the converter both use it. |
| `src/io/workbook_import.js` | Makes the three files from a filled workbook. Used by the converter below and by the page ("Make it from an old workbook"). |
| `src/page/` | The page: `app.js` (files, save, undo, tabs), `tabs.js` (unlock rules and reasons), `model.js` (edits, undo history, save output, typed units), `views.js` (tab contents), `files.js` (folder access), `ui.js` (inputs), `page.css`, `index.html`. |
| `tools/` | Development tools (Node). Not shipped. |
| `dist/` | The built page, `PCB_Thermal.html`. The only file the user gets. Rebuilt by `npm run build` (also by `npm test`). |

## Tools (development PC only)

```
cd webapp
npm install
npm test                                                     # W1, W2 and W3 tests (W3 also rebuilds the page)
npm run build                                                # dist/PCB_Thermal.html
npm run serve                                                # http://localhost:8173/webapp/dist/PCB_Thermal.html
npm run expected                                             # refresh tests/webapp/expected_*.json from the workbooks, then test
node tools/convert_workbook.mjs <workbook.xlsx> <out folder>  # workbook to JSON
node tools/validate.mjs <project folder>                     # check the three files
node tools/test_w1.mjs --write                               # refresh example/RaspberryPi4/*.json
```

The converter reads the workbook's saved values (what Excel last calculated). It never writes the workbook.
Build reports are read from the workbook's `CAD` folder when they are there. `forSolidWorks` is worked out by `src/calc/`
(the W2 test checks it equals the workbook's grey cells).

W2 test: `tests/webapp/expected_rpi4.json` and `expected_rpi4_fresh.json` hold every calculated value of `RaspberryPi4.xlsx`
and `RPi4_fresh.xlsx` as Excel saved them (the files were last saved by Excel, so the cached values are Excel's own results;
no Excel is needed to read them). `src/calc/` must match them to 6 significant digits. The register and thermal solution are
compared per RefDes: for a part with several loads (U3) the per row Pd, share and density are summed, and its Tj est is the
page's own (per part, by design).

## The page (W3)

Open `dist/PCB_Thermal.html` in Chrome or Edge and choose the board's project folder. The page:

- reads `project.json`, `cad.json` and `results.json`; an empty folder offers "Start a new project here" or
  "Make it from an old workbook (.xlsx)" (writes `cad.json` and `results.json` only when they are not there yet);
- saves `project.json` on Save or Ctrl+S, and on its own every 30 s while there are unsaved edits. Before the first save of
  the day it copies the old file to `backupproject_yyyymmdd.json`. It refuses to save a file that does not match the schema
  (validators compiled into the page by `tools/build.mjs`) or has a RefDes twice;
- checks `cad.json` and `results.json` every 2 s and reloads them when a macro has written them;
- keeps undo and redo for the session (Ctrl+Z, Ctrl+Y). Typing in one field within 1.5 s is one undo step.

Other browsers: open the three files by hand; Save downloads `project.json`.

`window.PCBT_DEV` is a development hook (open a folder handle from the browser's private storage, read the state) used for
browser checks; it changes nothing on its own.
