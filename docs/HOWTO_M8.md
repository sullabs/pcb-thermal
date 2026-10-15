# M8 two-resistor chip model

Part of `SetupSim`. No new macro. Files: `PTCommon.bas`, `SetupSim.bas`, `ReadResults.bas` (changed), `SetupWorkbook.bas` (changed).

## What it does
For a chip with Chip model = Two-resistor on Sim setup, SetupSim rebuilds its part file as three stacked bodies:
- bottom layer (feature `PT_Box`): thermal resistance theta_jb to the board
- junction core (`PT_LayerCore`): carries all the chip heat, k = `Sim_kCore` (default 100 W/m K)
- top layer (`PT_LayerTop`): thermal resistance theta_jc to the top face

Layer k = layer thickness / (theta x footprint area). Layer thickness = `Sim_LayerFrac` x chip height (default 0.1, each layer).
Same outer size as before, so BuildBoard, BuildSpreader and SyncFromCAD see no change. Tag `Layers_um` in the part says what is built.
Switching the chip back to Block makes SetupSim rebuild it as one box. The core temperature is the chip Tj (ReadResults reports the core).

## To use it
1. Update the macro project: re-import `PTCommon.bas`, `SetupSim.bas`, `ReadResults.bas`. Compile.
2. Existing workbook: the template needs `Sim_LayerFrac` and `Sim_kCore` on Sim setup (C12, C13). The script `tools/Update-TemplateColors.ps1` adds them (see HOWTO_COLOURS.md). By hand: label in B12, C12 = 0.1 named `Sim_LayerFrac`, label in B13, C13 = 100 named `Sim_kCore`.
   (new copies get them from SetupWorkbook). If the names are missing, SetupSim uses the defaults, so this is optional.
   Also update the Check formula if you want it to say "OK (two-resistor)" (SetupWorkbook writes it).
3. Sim setup: Chip model = Two-resistor, type theta_jb in column E (theta_jc comes from the register).
4. Run SetupSim, then ReadResults.

## What you should see
- SetupSim message and report: notes "bottom layer ... k ..." and "top layer ... k ..." per chip. Heat applied = register total.
- Simulation tree: three solid bodies for that chip, heat power on the middle one.
- Chip with no theta_jb or no heat: note, built as a block.
- Sanity check: junction to top face temperature difference = Q x theta_jc when the top face is the only path (roughly).

## Not tested
The macros were written but not run on SolidWorks: computer-use access was declined in this session, so the layer build
(`PT_BuildLayeredBox`, start-offset boss with Merge off) and the body-name matching need a first run. If a layer fails,
SetupSim says so in the report and falls back to the block model. Send the compile error or the report notes.
Thin layers with a fine mesh of 0.5 mm may need a smaller `Sim_MeshFine`.
