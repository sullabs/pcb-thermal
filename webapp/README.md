# PCB thermal page

The page that replaces the workbook as the user interface. Design and milestones: `docs/WEBAPP_DESIGN.md`.

## Folders

| Folder | What |
|---|---|
| `schema/` | JSON Schemas of the three project files: `project.json` (page writes), `cad.json` and `results.json` (macros write). `common.schema.json` holds shared parts. |
| `src/io/workbook_import.js` | Makes the three files from a filled workbook. Used by the converter below and, later, by the page. |
| `tools/` | Development tools (Node). Not shipped. |
| `dist/` | The built page, `PCB_Thermal.html` (from W3). The only file the user gets. |

## Tools (development PC only)

```
cd webapp
npm install
npm test                                                     # W1 tests
node tools/convert_workbook.mjs <workbook.xlsx> <out folder>  # workbook to JSON
node tools/validate.mjs <project folder>                     # check the three files
node tools/test_w1.mjs --write                               # refresh example/RaspberryPi4/*.json
```

The converter reads the workbook's saved values (what Excel last calculated). It never writes the workbook.
Build reports are read from the workbook's `CAD` folder when they are there.
