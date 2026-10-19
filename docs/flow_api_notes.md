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

## F0 spike for the Flow plan (2026-10-19)

Checked offline in `cfwapi.chm` (2023) and against the files of the solved RPi4 run
(`CAD\1\1.info.json`, `calculation_warnings.txt`). Nothing here is run live yet; the first live RunFlow run confirms it.
Ids are the numbers of `nikFeaturesTypes_e` (feature) and `nikParametersTypes_e` (parameter).

| # | Question | Answer | Fallback chosen |
|---|---|---|---|
| S1 | Project folder from `FDAProject` | `fda.ProjectFiles.ProjectDirectory` (full path of `CAD\<n>\`) | newest `<n>\<n>.info.json` under the assembly folder whose `project_name` is ours |
| S2 | Global mesh level, gap; local mesh | `fda.GlobalMeshSettings` (a ParametrizedFeature): `nikInitialMeshLevel` 7, `nikGapSize` 8 (m). Local mesh = feature 14: `nikFluidCellsRefinementLevel` 67, `nikSolidCellsRefinementLevel` 68, `nikFluidSolidBoundaryCellsRefinementLevel` 69, `nikAdjustChannelRefinement` 82, references with `ITopologyBasedFeature.AddComponent` | a parameter the project does not offer is noted in the report |
| S3 | Finish conditions and goal criteria | `fda.GetCalculationControlOptions`: `FinishConditionsRule` (0 one, 1 all satisfied), `UseGoalsConvergence`, `UseMaximumTravels` + `UseManualMaximumTravels` + `MaximumTravels`, `UseMaximumIterations` + `MaximumIterations`. **No API for the criterion of a goal**, and no "minimum travels" | min travels = rule "all satisfied" with goals convergence on and manual max travels = minTravels (the solve stops only when both hold). Goal criteria not set; F1 flags a goal whose delta is above `goalCriteriaC` |
| S4 | User material from the API | **Not possible.** A solid material feature (8) only takes a database path, `nikPathToSolidMaterial` 32 | the page names one material per value (`User Defined\PCBThermal\<name>`) and lists them with k xy, k z, density and cp; the user adds the missing ones once in Flow's Engineering Database (docs/HOWTO_FLOW_MATERIALS.md). The macro tries the user path, then the old library pick, and says which it used |
| S5 | Radiative surface | feature 9, database path `nikPathToToRadiationMedia` 45. No emissivity value parameter | same as S4: `User Defined\PCBThermal\e 0.90` style names listed by the page; when missing, the template's surface stays and the report says so |
| S6 | Contact resistance | feature 15, `nikThermalContactResistanceType` 89 (1 = integral, K/W), `nikIntegralThermalResistance` 88, `nikApplyThermalContactResistanceToSolidSolidBoundaryOnly` 90, references by component | none needed (value is set directly) |
| S7 | Gravity, pressure, air speed | `fda.GravitationSettings` with `nikGravityVectorX/Y/ZValue` 83 to 85 (m/s2, `fda.EnableGravitation = True`); `fda.GlobalInitialConditions`: `nikEnvironmentPressure` 9 (Pa), `nikVelocityAlongX/Y/Zdirection` 11 to 13 (m/s; ambient = initial conditions in an external analysis) | none needed |

Other facts found on the way:
- `<n>.info.json` lists only the project goals (unit `temperature` for solid temperature goals), with `progress` (%),
  `criteria` and `delta` (K), plus `finished`, `mesh.cells_total/cells_fluid/cells_solid` and `telemetry.iteration/travel`.
  The service goals (`Serv Heatf`, `Radiation Energy Delta`) are **not** in it; RunFlow reads them from
  `GetGoalsCalculationResults2` (SI: W). RPi4 run FL_20261019_2003: 3.103 W + 3.765 W = 6.868 W against 6.859 W applied (0.1 %).
- `calculation_warnings.txt` lists the warnings after the line `Warnings on last iteration:` (empty when there are none).
- `FDAResultFileInfo.GetMeshCellsCount` gives the total cells too, but info.json has the split, so it is not used.
