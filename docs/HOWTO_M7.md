# M7 RunFlow (Flow Simulation)

- `RunFlow.swp` = `PTCommon.bas` + `RunFlow.bas`
- Needs: SOLIDWORKS Flow Simulation turned on (Tools > Add-Ins) and the Flow API SDK installed
  (`flow simulation api sdk x64.exe` in the `api` folder of the Flow Simulation install).
- Set `Sim setup` study type to `Flow`.

## The Flow project is built by the macro
RunFlow makes the project `<prefix>Flow` (for example `CL_Flow`) from `PCBThermal_Flow.fwp` the first time. The template
sits next to the macro (`macros\`) or in `C:\ProgramData\COSMOS Applications\Flow Simulation 2023	emplate\`. It holds
external air, conduction in solids, gravity -Z, radiation with default walls `Shellac, black matte` (e about 0.9) and
a "PCB material" whose axis is Z. On every run the macro points these features (fixed names) at the current components:
- Materials: `PCB material` (Flow's `PCB 4-layers`, `8-layers` or `12-layers` from the copper layer count), `Parts material`
  (typical plastic package), `Spreader material` (aluminium, copper when the spreader k is 300 or more), `TIM material` (GR25A).
- One volume source `<RefDes> heat` per part with heat. A source you made yourself whose name holds the RefDes is kept instead.
- Goals `<RefDes> max`, `<RefDes> avg`, `PCB max/min`, `Spreader max/min` (volume goals, solid temperature).
- The computational domain: the assembly box plus half its longest side around it and 1.5 times above it.
A feature whose component is gone is deleted. Anything else you add by hand (a fan, a local mesh, a lid) stays.

To change what every new project starts with: open a project made from the template, change it (General Settings,
mesh), then Flow Simulation > Project > Create Template, name `PCBThermal_Flow`, and copy the `.fwp` from the
ProgramData template folder to `macros\`.

## Run
1. Open the thermal assembly so it is the active window.
2. Run `RunFlow`, pick the workbook.
3. It sets each source to the register heat, sets the initial temperature to the ambient, then asks:
   Yes = solve now and read the goals, No = read the goals of the last solved run, Cancel = stop after the update.
4. Result: `Sim results` row (run name `FL_<date>_<time>`), `Thermal solution` column AF, report
   `CAD\<prefix>RunFlow_<run>.txt` with every source and goal it handled.

## Check
- Total heat set = register total (the message says both). A RefDes with heat and no source is listed.
- AF fills for the chips that have a max goal. AG = hand minus Flow.

## Known limits
- RunFlow creates and deletes only its own named features. Materials come from the Flow library, so the page's chip,
  TIM and PCB k values are not used: Flow and Thermal chip temperatures differ where those matter.
- The anisotropy axis and the default wall radiative surface have no API parameter: they come from the template.
- Boundary condition temperatures are not changed, only the initial temperature.
- The API units are SI: goals are kelvin, the macro writes degrees C.
- The first run on a PC may fail at connecting to the API. The macro then says which step failed. Report the message.
  The help (cfwapi.chm) says a 32-bit VBA host needs a reference to `NIKCommonApiSrvLib`; RunFlow already tries the
  out-of-process server class for that case.
- Tested 2026-10-01 on the RPi4 thermal assembly (test project: external air, conduction, 8 sources, 27 goals). Connects, sets
  sources and ambient, solves (Yes) and reads last run (No), writes Sim results and column AF. See `docs/flow_api_notes.md`.
- The Flow API returns Boolean results as a raw 1, so the macro tests them with `IsTrue()` (non-zero), never `Not value`.
- A goal that comes back as 0 K (body not meshed or not solid) is skipped and listed in the notes, not written to Excel.
- Flow Simulation must be loaded in the running SolidWorks (Tools > Add-Ins) or the API cannot attach.
- RunFlow sets three temperatures: fluid, initial solid and radiation environment (parameter ids 10, 26, 30).
- Realistic test run 2026-10-01 (external air 40 C, gravity, radiation on, PCB FR4, chips silicon, rest aluminium, mesh level 2,
  6.67 W): chips 97 to 99 C, PCB min 47 C. Tiny parts (L1, L2 inductors) were not meshed at level 2 and came back as 0 K, so
  they were skipped with a note. Use a higher mesh level or a local mesh on them. No radiative surfaces were defined.
- Macro-built run 2026-10-19 (RPi4, 40 C, 6.86 W, radiation on, spreader): chips 89 to 124 C, PCB 72 to 123 C, spreader
  86 to 88 C. The Thermal study gave U1 90, U4 88, U2 82, U6 82, U5 113 and U3 229 C (its own chip k). Without radiation
  the same Flow project ran 40 C hotter.
