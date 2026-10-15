# Flow Simulation API notes (SW 2023), result of the M7 spike

Source: `cfwapi.chm` from the Flow API SDK (installed to `C:\Program Files\SolidWorks Corp\SolidWorks Flow Simulation\api SDK 2023\`).
Checked offline against the help and the registered type libraries. Not yet run on a live project.

## What the API can do
| Need | API |
|---|---|
| Heat sources | project feature `nikVolumeSource` (4), parameters `nikVolumeSourceType` (43, 1 = heat rate) and `nikHeatGenerationRate` (21, W) |
| Materials | `nikSolidMaterial` (8), string parameter `nikPathToSolidMaterial` (32), references by component or body |
| Ambient | `GlobalInitialConditions` parameter `nikTemperature` (10, kelvin) |
| Boundary conditions | `nikBoundaryCondition` (2), type parameter 41, values by parameter |
| Goals | read only through results (`GetGoalsCalculationResults2`); creating goals is possible as features 12 and 13 but RunFlow does not do it |
| Run | `FDAProject.Solve2(buildMesh, recalc, closeMonitor, runBRP)` |
| Read | `FDAProject.LoadResults(True)`, then goal name and `GetLastCalculatedValue` (SI) |
| Mesh | `GlobalMeshSettings`, local mesh feature 14 |

## Call chain
`BaseApiObject` -> `LoadProductAPI2("Flow Simulation","2023")` -> `Attach2RunningObject2(pid)` -> `ActiveDocument`
-> `GetActiveProject` -> `EnumFeatures` / `CreateTemporaryFeature` -> `GetInterface("IParametrizedFeature")`
-> `EnumParameters` -> `SetValue` / `SetLongValue` / `SetStringValue` (SI) -> `Solve2` -> `LoadResults`.

## Connecting
- Direct class: `NIKCommonApi2.BaseApiObject` (in-process, `system32\nikcommonapi2.dll`, 64-bit).
- Out-of-process server for 32-bit hosts: class `{478B60C3-62E1-4643-A9E0-8A448E7C3E68}`
  (`NIKCommonApiSrvLib`, `system32\nikcommonapisrv.exe`), method `LoadNIKCommonApiLibrary`. No ProgID, so `GetObject("new:{...}")`.
- The product must be running with the Flow add-in on; the API attaches to it by process ID.

## Decision for M7
Route from the plan holds: the user builds one Flow project by hand, `RunFlow` updates heat by RefDes, sets the initial
temperature, solves and reads goals into `Sim results`. Creating features from scratch is possible but not needed yet.

## Open questions for the first live run
1. Does `Solve2` block until the solve ends? Assumed yes (the SDK sample solves then reads a goal).
2. Does the macro host reach the API directly or only through the server class? Both are tried.
3. Are `Feature.Type` values returned as numbers? Assumed numeric text or number.
4. Do source parameters change live or only after `Rebuild`? The macro calls `Rebuild` on each source.
