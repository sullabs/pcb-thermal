# M6 SetupSim and ReadResults

- `SetupSim.swp` = `PTCommon.bas` + `SetupSim.bas`
- `ReadResults.swp` = `PTCommon.bas` + `ReadResults.bas`
Both use SOLIDWORKS Simulation through late binding: no extra reference to tick. The add-in must be on
(Tools > Add-Ins > SOLIDWORKS Simulation); SetupSim tries to load it if it is off.

## SetupSim
Reads Sim setup, Spreader, CAD geometry and the Component register, then in the thermal assembly:
1. New steady-state thermal study `PT_yyyymmdd_hhmm` (never overwrites an older one). The solver is left on Automatic (forcing FFEPlus crashed SolidWorks 2023 SP5).
2. Materials: PCB orthotropic (k in-plane, k through-plane, Z normal to the PCB Front plane);
   chips with a calibrated k z on Sim setup as orthotropic blocks (k xy blank = k z);
   parts without heat: one generic k; spreader: Spr_k; TIMs: TIM k of their chip.
3. Heat power per RefDes = sum of Q used over its register rows, on the bodies with that RefDes.
4. Global contact bonded. Thermal resistance sets when the Sim setup value is above 0
   (part to PCB, chip to TIM, TIM to spreader). Chips in TIM mode Contact R: total resistance TIM R.
5. Convection (h, ambient) and radiation (emissivity, view factor 1) per surface group.
   Faces left out: part bottoms on the PCB, chip tops under the spreader, TIM tops and bottoms,
   pedestal contact faces.
6. High quality mesh, global size, fine size on chips with heat and TIMs. Run.
Anything missing (RefDes with heat but no body, a named face, a refused load): the study is made but not
meshed or run, and the message lists what to fix. Report: `CAD\<prefix>SetupSim_<run>.txt`.

"Include parts without heat" = No: the study is made in configuration `PT_HeatOnly` where parts without heat
are suppressed. BuildBoard, BuildSpreader and SyncFromCAD switch back to the normal configuration themselves.

## ReadResults
Reads the study named in the run name cell of Sim setup:
- Sim results: one new row per run (inputs snapshot, spreader and PCB min/max, energy balance, plot file),
  peak and average per RefDes (columns added for new RefDes).
- Thermal solution column AF = peak per RefDes, so AG = hand minus sim.
- Picture of the first result plot: `CAD\results\<run>.png`.

## Known limits
- Convection on PCB top and bottom covers the whole face, also where parts sit (small error).
- Contact R mode leaves a real gap between chip and spreader. If SolidWorks refuses the resistance over a gap,
  use TIM mode Body for that chip.
- Energy balance: read from Simulation's own heat totals. Check on the first run that the heat in the message equals the register total (see "Needs a test run" in WORKFLOW_BARE_BOARD.md).

## While SetupSim runs
The SolidWorks status bar shows a progress bar with the step (assembly, materials, heat, meshing, solving).
Meshing and solving take several minutes on the Raspberry Pi 4; SolidWorks is busy and does not react. Do not close it.
