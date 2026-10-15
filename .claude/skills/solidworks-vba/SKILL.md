---
name: solidworks-vba
description: Write or review SolidWorks 2023 VBA macros, Simulation and Flow Simulation API calls for this PCB thermal project. Use before adding or changing any SolidWorks API call, and when a macro fails with a type mismatch, empty sketch, zero point, or similar.
---

# SolidWorks VBA for this project

Never guess an API call. Look it up on this PC first, then write the code. Read `CLAUDE.md` for the project rules
(ByVal, ASCII, CRLF, no reserved words, mm in Excel and metres in SolidWorks).

## 1. Look up a call (offline, SolidWorks need not be running)

Interop DLLs: `E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\`. Use the helper next to this file:

    powershell -NoProfile -File .claude/skills/solidworks-vba/api_lookup.ps1 -Type IFace2 -Member Tess
    powershell -NoProfile -File .claude/skills/solidworks-vba/api_lookup.ps1 -Type IBody2
    powershell -NoProfile -File .claude/skills/solidworks-vba/api_lookup.ps1 -Enum "swEndConditions_e"
    powershell -NoProfile -File .claude/skills/solidworks-vba/api_lookup.ps1 -Dll cosworks -Type ICWStudy -Member Mesh

It gives exact method names, parameter order and types, return types, and enum values. It does NOT give remarks
or the meaning of returned arrays. For those use the help files (section 2).

Online docs do not work: `help.solidworks.com` is script rendered and returns 403 to scripts. The old
`FeatureExtrusion2 fails with more than 12 parameters` advice (from the third party solidworks-api-skill) does not apply:
`PT_Boss` passes all 23 arguments from VBA and it is the real signature.

## 2. Remarks, parameters, array layouts (full reference, offline)

    powershell -NoProfile -File .claude/skills/solidworks-vba/api_doc.ps1 -Type IFace2 -Member GetTessTriangles
    powershell -NoProfile -File .claude/skills/solidworks-vba/api_doc.ps1 -Type ITessellation

Prints the help page text (syntax, parameter meaning, return value, remarks). First run unpacks the 2023 API reference
(14000 pages) to `%LOCALAPPDATA%\swapi_help` with 7-Zip. Use `-Ns cosworks` for Simulation (check the file prefix there).
Other help files: `E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\API_GB.chm` (getting started and example index; the example code
itself is online only, see 2c) and `E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\cworksapi.chm` (Simulation); Flow Simulation:
`C:\Program Files\SOLIDWORKS Corp\SolidWorks Flow Simulation\api SDK 2023\english\cfwapi.chm` (7-Zip extracts chm files).
Flow call chain and ids: memory file `flow-simulation-api`. When you learn something from a page, add it to section 5.

## 2b. Reference files (grep them, do not read whole)

Generated from the local help by `build_reference.py` (rerun it if the help changes; needs the unpacked help from `api_doc.ps1`).
- `reference/core.md`   signatures, parameters, remarks for the core calls the macros use (about 4700 lines).
- `reference/simulation.md`   every member of the thermal-relevant Simulation interfaces (about 6600 lines). Entry point:
  `ICosmosWorks` -> `ICWModelDoc` -> `ICWStudyManager` -> `ICWStudy` -> `ICWLoadsAndRestraintsManager` (AddHeatPower,
  AddConvection, AddRadiation), `ICWMesh`, `ICWResults`. Use `grep -n "### ICWStudy\." reference/simulation.md`.
- `reference/simulation_examples.md`   VB6 example code from cworksapi.chm.
- `reference/core_examples.md`   VBA examples fetched from the online help (tessellation bounding box, component transforms,
  bodies in components).

## 2c. Tool: official VBA examples on demand

When a call is new or a signature alone is not enough, fetch the official example before writing code:

    python .claude/skills/solidworks-vba/api_example.py IFace2.GetTessTriangles
    python .claude/skills/solidworks-vba/api_example.py IFeatureManager.FeatureExtrusion3 --list
    python .claude/skills/solidworks-vba/api_example.py IBody2.GetTessellation --title "Non-manifold"
    python .claude/skills/solidworks-vba/api_example.py ICWLoadsAndRestraintsManager.AddHeatPower --ns cosworks

It opens the member's online help page in headless Chromium (the site blocks plain scripts), follows the "(VBA)" example links
and prints the code. Results are cached in `example_cache/`, so repeat lookups work offline. Use the newest member name:
obsolete ones have no examples (`FeatureExtrusion2` has none, `FeatureExtrusion3` has). Needs `pip install playwright` and any
Chromium (Playwright's own under `%LOCALAPPDATA%\ms-playwright`, or Chrome). If it prints "did not load fully", run it again.
Do not crawl the whole site; the method reference is already local.

## 3. Rules that cost failed runs here

- Units: API metres, workbook mm. Convert at the boundary only.
- Typed locals for math objects: `SldWorks.MathUtility`, `MathPoint`, `MathTransform`. Late bound
  `MathUtility.CreatePoint(array)` returns a zero point (fixed in `PT_ToSketch`).
- Do not look up planes, sketches or features by display name (language dependent). Use `PT_RefPlane`, feature type
  names (`ProfileFeature`) and the `PT_` feature names we set.
- A sketch with no entities is deleted by `InsertSketch True`. If the rectangle call fails, no `ProfileFeature` exists
  and `PT_LastFeatureOfType` returns Nothing ("the sketch was not made"). The message now carries the size.
- Component transform: `Component2.Transform2` may be a part-to-assembly or its inverse, and `GetBox` is axis aligned in
  assembly space. ImportStep self-checks three modes against `GetBox` (`PickMode`). Do the same for any new code.
- Tessellation (confirmed from the help and the online example "Get Part Bounding Box"): call `GetTessTriangles(True)`
  (True = no unit conversion, raw metres; the SolidWorks example does this). It returns a flat array of 9 numbers per
  triangle (x,y,z of three points, in METRES). These are display triangles, fine for a bounding rectangle, not for
  machining. `IBody2.GetTessellation(FaceList)` returns an `ITessellation` (empty FaceList = whole body): set
  `NeedFaceFacetMap/NeedVertexParams/NeedVertexNormal = False`, call `Tessellate` (Boolean), then `GetVertexCount` and
  `GetVertexPoint(i)` (array of 3 doubles). `IComponent2.GetTessTriangles` needs a lightweight component, so do not use it.
  Whether body points are in part space or assembly space is not documented: the self-check in `PickMode` decides.
- Arrays from the API come back as `Variant`; test with `IsEmpty` and `IsArray` before `UBound`.
- Booleans from Flow Simulation come back as raw 1: test with `CLng(x) <> 0`.
- Passing objects: `ByVal x As Object`. An element of a Variant array cannot go to a `ByRef ... As Object` parameter.
- Save and close part documents you open from a macro, or the assembly rebuild uses stale copies.
- Simulation: steady state thermal is `SolutionType 1`. Create objects late bound. Temperatures to the solver in kelvin,
  conductivity W/m K. The API has no exclude-from-analysis call: use configuration `PT_HeatOnly`.
- `SelectByID2` with a named entity is fragile. Prefer `Feature.Select2` and `Face2.Select4`.

## 4. How to test

- Compile gate: user runs Debug > Compile VBAProject. Re-read the file for CLAUDE.md rules first.
- COM from the shell only works when Claude Code runs in the user's desktop session (SessionId 1). In session 0 Excel and
  SolidWorks give `E_ACCESSDENIED`. Check with `(Get-Process -Id $PID).SessionId` before planning a COM test.
- Harness: `tests\swp\PTTest.swp` holds PTCommon plus every macro module. See memory file `pc-environment`.
- Test board: `Raspberry Pi 4 Model B.STEP`, then `docs/TEST_PLAN.md`.

## 5. Verified facts (add to this list)

- `FeatureExtrusion2` has 23 parameters (checked 2026-10-14 in the 2023 interop DLL), `FeatureCut4` has 27.
- `IFace2.GetTessTriangles(Boolean NoConversion) -> Object`: 9 values per triangle, metres (help, 2026-10-14).
- `IBody2.GetTessellation(Object FaceList) -> Object` (ITessellation); `Tessellate() -> Boolean`; `GetVertexPoint(Int) -> Object` (3 doubles).
- `IBody2.GetBodyBox` is approximate and may change after a rebuild: do not use it for comparison.
- `IComponent2.GetBodies3(BodyType, ByRef BodiesInfo)` supports lightweight components; order of bodies may vary.
- `IComponent2.Transform2` is relative to the root component of the active assembly.
- `IModelDocExtension.SelectByID2(Name, Type, X, Y, Z, Append, Mark, Callout, SelectOption)`.
- Simulation enums checked in the interop DLL (2026-10-14): swsAnalysisStudyTypeThermal = 3, swsThermalSolutionTypeSteadyState = 1,
  swsContactSetTypeThermalResistance = 0 / Bonded = 1, swsResistanceTypeTotal = 0 / Distributed = 1, swsRadiationTypeSurfaceToAmbient = 0,
  swsMeshQualityHigh = 1, swsTemperatureUnitCelsius = 2, swsSelectionFaceEdgeVertexPoint = 0. SetupSim and ReadResults use these.
- `ICWResults.GetHeatPowerOrEnergy2` returns { sel in, sel out, sel net, model in, model out, model net }, net = in - out. In steady
  state the model in and out are equal, so the energy balance error is |model in - model out| / larger of the two.
- Name clash: VBA is not case sensitive. A parameter `roles` and a local `roleS` is a compile error (duplicate declaration).
  `.claude/skills/solidworks-vba/lint_vba.py` finds these (reserved names, duplicate Dim, missing ByVal, non-ASCII, CRLF).
