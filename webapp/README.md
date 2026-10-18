# PCB thermal page

The page that replaces the workbook as the user interface. Design and milestones: `docs/WEBAPP_DESIGN.md`.

## Folders

| Folder | What |
|---|---|
| `schema/` | JSON Schemas of the three project files: `project.json` (page writes), `cad.json` and `results.json` (macros write). `common.schema.json` holds shared parts. |
| `src/calc/` | The calculation module: register, triage, rail budget, PCB stack, body checks, bare board network, spreader plates, thermal solution, sim k z, and `forSolidWorks`. Pure functions on the three files; the page and the converter both use it. |
| `src/io/workbook_import.js` | Makes the three files from a filled workbook. Used by the converter below and, later, by the page. |
| `tools/` | Development tools (Node). Not shipped. |
| `dist/` | The built page, `PCB_Thermal.html` (from W3). The only file the user gets. |

## Tools (development PC only)

```
cd webapp
npm install
npm test                                                     # W1 and W2 tests
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
