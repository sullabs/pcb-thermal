# Workflow: bare board simulation, then the spreader

Goal: build the board, simulate it with no spreader, compare with the hand calculation, and only then move to the spreader.
Nothing in this list needs the spreader. SetupSim makes a "bare board study" on its own when the assembly has no
unsuppressed spreader.

**Build the spreader only after the bare board simulation has been compared with the hand calculation (stages C to G).**

Macros are in the toolbar order of `HOWTO_TOOLBAR.md`. Each step says what to check before moving on.

## Stage A. Hand calculation first (Excel only)

1. `1 Component register`: fill every chip with heat. Check Pd (D4), ambient (D5), theta_jc, Rjb.
2. `6 Bare board network`: the hand Tj per RefDes with no heatsink.
   Top path goes through the case, board path through PCB spreading.
   Surface h and emissivity are read from `9 Sim setup`, so set those first.
3. Write down the hand Tj for each chip. This is the number the simulation must match.

## Stage B. Import and build the board

4. Run **ImportStep**. It fills `CAD geometry`: one row per body, envelope plus rotation.
   New rows have Include = No.
5. In `CAD geometry` set **Include = Yes** for:
   - the PCB,
   - every part that has heat in the register (a heat part left at No gets no heat source),
   - any part you want in the model for shape or shadowing.
   ImportStep sets Include = Yes by itself for a body whose name key equals a register RefDes (first link only; a row you
   set to No later stays No). The PCB has no RefDes, so switch it on yourself. Leave the rest at No. Link each heat body to its RefDes (column RefDes) if ImportStep did not do it.
6. Tick the rows in `Change log`. SetupSim refuses to run while any row is unchecked.
7. Run **BuildBoard**. Read `CAD\<prefix>BuildBoard_report.txt`:
   - failed parts: fix the row and rerun,
   - overlaps: fix before simulating (override on CAD geometry or Include = No on one body),
   - parts cut at the board surface: check they are intended.
8. Look at the model in SolidWorks. Compare with the STEP: positions, rotations, heights.

## Stage C. Check the model before simulating

9. Run **SyncFromCAD**. It reads back what you changed in SolidWorks and fills the as-built numbers.
   Tick the new Change log rows.
10. Confirm in the tree: one PCB body, a body for each heat RefDes, no spreader and no TIM in the assembly.
    If a spreader exists from an earlier run, suppress it with its TIMs, or the study will not be a bare board study.

## Stage D. Sim setup

11. `9 Sim setup`:
    - Sim_Type = Thermal
    - Sim_Ambient = register D5
    - PCB k in-plane and through-plane (from `19 Materials` or the board stack-up)
    - h and emissivity for PCB top, PCB bottom and Parts. Same values the hand network uses.
    - Interface resistances: start with 0 so the first comparison is clean.
    - Chip model = Block. Check the calibrated k z for each chip.
    - Mesh: global and fine (chips).
    - Sim_IncludeNonHeat: Yes for the first run (all built parts have a generic k). No gives the heat-only configuration.
12. Do not change the hand-network inputs between runs. One variable at a time.

## Stage E. Run

13. Run **SetupSim** (Run Thermal Study). Read the message and `CAD\<prefix>SetupSim_<run>.txt`:
    - "Bare board study" must be in the report,
    - heat applied must equal the register total,
    - nothing listed as missing. If something is, the study is made but not run: fix and rerun.
14. If the study did not solve, open it in SolidWorks Simulation and run it there.
15. Run **ReadResults**. It writes `10 Sim results` (new row), `Thermal solution` column AF
    (peak per RefDes) and the plot `CAD\results\<run>.png`.

## Stage F. Compare

16. `6 Bare board network` reads the Sim Tj by run name (blank = newest). Check hand Tj against sim Tj per chip.
    Column AG on `Thermal solution` is hand minus sim.
17. Sanity checks, in this order:
    - energy balance in `10 Sim results` within 2 %,
    - total heat in = register total,
    - board temperature field looks like the heat sources and the edges,
    - chip rise above its board spot is about Q x theta_jc when the top is the main path.
18. Reading the difference (hand minus sim):
    - Sim hotter on every chip: the model has less surface area or the hand h is too high. Check which faces got convection.
    - One chip off: wrong body size, wrong RefDes link, wrong heat, or the mesh is too coarse there.
    - Spread across chips with the same sign: PCB k or copper spreading in the hand network.
    - Large gap from rotation or merged parts: compare the box with the STEP again.
19. Change one thing, rerun SetupSim and ReadResults, compare again. Each run adds a row. Older studies stay.
    Stop when hand and sim agree within the tolerance you set (suggest 10 % on the rise above ambient, or 5 C).

## Stage G. Accept and move on

20. Record in `10 Sim results` which run is the accepted bare board baseline (note the run name).
21. Optional second pass: Chip model = Two-resistor with theta_jb (M8), then compare again.
22. Only then: `7 Spreader` tab (Covered column, TIM settings, Check column), **BuildSpreader**, SyncFromCAD,
    SetupSim again. The bare board run is the reference for how much the spreader helps.

## Risks to keep in mind

- Convection on PCB top and bottom covers the whole face, also under parts (small error).
- A heat part left at Include = No gets no heat. The total applied heat check catches it.
- Include on CAD geometry applies to parts. The PCB size comes from the Used column (Length X, Width Y, Thickness Z) and has no Include.
- The M8 two-resistor chip model has not been run yet.
