# M7 RunFlow (Flow Simulation)

- `RunFlow.swp` = `PTCommon.bas` + `RunFlow.bas`
- Needs: SOLIDWORKS Flow Simulation turned on (Tools > Add-Ins) and the Flow API SDK installed
  (`flow simulation api sdk x64.exe` in the `api` folder of the Flow Simulation install).
- Set `Sim setup` study type to `Flow`.

## Make the Flow project once, by hand
On the thermal assembly (the one BuildBoard made), with the normal configuration active:
1. Flow Simulation > Wizard. Internal or External as you need, conduction in solids on, gravity and radiation as you need.
2. One **volume heat source per chip with heat**. Name each one with the RefDes in it: `U1`, `Q U1`, `U1 heat`.
   Pick the chip body as its reference. Any heat value: RunFlow overwrites it. Keep it a constant, not a formula.
3. Solid materials, boundary conditions, mesh: as you need. RunFlow does not touch them.
4. Goals (volume goals, temperature of solid). Name each one with the RefDes plus `max` or `avg`:
   `U1 max`, `U1 avg`. No word means max. For the plate and board use `Spreader max`, `Spreader min`, `PCB max`, `PCB min`.
5. Save the assembly.

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
- RunFlow does not create or delete features, materials or boundary conditions.
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
