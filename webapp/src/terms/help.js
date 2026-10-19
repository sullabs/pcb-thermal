// Help page texts (design section 7). One entry per section; tab sections add to the Learn texts (learn.js), and the
// glossary is built from the term catalogue, so a term's hover text and its glossary entry never disagree.
// Paragraphs are plain strings (abbreviations are marked by the page); lists are arrays. Write Pd, Rja, TIM as they are.
//
//   HELP: [{ id, title, body: [paragraph or [list items]], aliases }]   aliases: other ids that open this section
//   TAB_HELP[tab id]: { you, out, mistakes }                         the extra text of each tab's section
//   MACRO_HELP[macro id]: { changes, check }                          the SolidWorks steps section

export const HELP = [
  { id: "start", title: "Start here", body: [
    "This page takes a PCB from its STEP file to a checked thermal result. You do the engineering: the power of each part, the datasheet values, and judging the results. The SolidWorks macros do the manual work: they build the model and set up and run the study.",
    "The steps in one line: list the parts and their power (tab 1), bring in the 3D shape (tab 2, Import STEP), link each part to its body (tab 3), see which parts matter (tab 4), check the board by hand (tab 5), build the model (tab 6, Build Board), add a spreader if needed (tab 7), run the study (tab 8), and read the results against the limits and the hand check (tab 9).",
    "Who writes what: the page writes only the project file. The macros write the cad and results files. The page reads those every 2 s and shows what a macro did as soon as it has finished. Neither ever overwrites the other's file.",
    ["Tab 0: choose the board's folder. For a new project type your prefix (your initials) and the board name: the page writes AB_project.json, AB_cad.json, AB_results.json and the CAD folder there.",
      "Follow the tabs in order. A grey tab says what it waits for and links to the step that unlocks it.",
      "Learn in the top bar explains each step; Show formulas writes every formula with your numbers; Hover help turns the texts that appear when you point at a value on or off.",
      "The example study (Raspberry Pi 4) shows every tab filled in. It opens read only, so nothing can break."],
  ] },
  { id: "heat", title: "Heat basics", body: [
    "Heat leaves a part in three ways. Conduction: through solid material, from hot to cold, like current through a resistor. Convection: from a surface into the air moving past it. Radiation: from a surface to everything around it, even in still air.",
    "Thermal resistance is Ohm's law for heat. A temperature difference drives heat through a resistance: ΔT = Q × R. Q is the heat in W (like current), ΔT the rise in C (like voltage), R in C/W. Resistances in series add; parallel paths share the heat.",
    "Conduction through a layer: R = t / (k × A), with t the thickness, k the conductivity and A the area. A thin, wide layer of a good conductor is a small resistance; a thick TIM with a low k is a large one.",
    "Convection from a surface: R = 1 / (h × A). Still air gives h of 5 to 10 W/(m² K), a fan 25 to 100. Radiation adds roughly 5 to 6 W/(m² K) for a dark surface at moderate temperatures; the page works it out as an extra h from the emissivity and the surface rise.",
    "The points that matter on a board: the junction (the silicon, Tj), the case top of the package, the board under the part, and the air (Ta). Every limit in a datasheet is on Tj, so every calculation here ends at Tj.",
  ] },
  { id: "datasheet", title: "Reading a datasheet", body: [
    "Find the thermal table, usually near the absolute maximum ratings or in a section called Thermal information or Package thermal characteristics.",
    ["Tj max: under absolute maximum ratings or recommended operating conditions. Use the operating limit, not the storage temperature.",
      "Rja (θJA): junction to ambient on the JEDEC test board, a 76 × 114 mm board in still air. It ranks parts; it does not describe your board.",
      "Rjc (θJC top): junction to the top of the case. This is the one to use under a heatsink or a spreader.",
      "Rjb (θJB): junction to the board under the part. Some datasheets call the exposed pad path θJC(bottom): type that one as Rjb.",
      "ΨJT and ΨJB (psi) are not resistances: they are for estimating Tj from a measured case or board temperature. Do not type them as Rjc or Rjb."],
    "Traps: Rja of an exposed pad package assumes the pad is soldered to a big copper area with vias. Rjc of a BGA with a lid is to the lid, not the die. Values for a 2-layer and a 4-layer JEDEC board differ a lot: take the one closest to your board.",
    "For parts without a junction (inductors, connectors, resistors) use the hottest allowed temperature as Tj max, and leave Rjc and Rja blank.",
  ] },
  { id: "methods", title: "Power methods", body: [
    "Pd is the electrical power turned into heat in the part. It is not the power the part delivers. Each load of a part has a method that says how its Pd is worked out; the class of the part picks a default method.",
    ["LDO / linear regulator: Pd = (Vin - Vout) × Iout, plus Vin × the quiescent current if you know Iin. All the dropped voltage becomes heat.",
      "Switching regulator: Pd = Pout × (1 / η - 1), with η the efficiency at this load from the datasheet curve. A 3 W buck at 85 % loses about 0.53 W.",
      "Digital load (V × I): Pd = Vin × Iin. A processor, memory or interface IC turns almost all its supply power into heat. If it passes power on to another part (a load switch), subtract Vout × Iout.",
      "Resistive (I squared R): Pd = I² × R. Inductors (DCR), sense resistors, MOSFETs (RDS(on)), connectors (contact resistance per pin).",
      "Manual entry: type the Pd measured or given by the datasheet. It replaces any calculated value."],
    "Count each watt once. A regulator's loss belongs to the regulator; the part it feeds has its own Pd from its own current. Compare the sum of all Pd with the measured board input power on tab 0: the difference is heat you have missed or counted twice.",
    "Use the load case you study: full load, worst supply voltage, hottest ambient. A thermal result is only true for the case it was worked out for.",
  ] },
  { id: "network", title: "The bare board network", aliases: ["spreading"], body: [
    "Tab 5 works out Tj of every part on the bare board, with no heatsink. Each part has two heat paths, drawn on tab 5 as a resistor network.",
    ["Top path: junction to the case top (R j-top, from Rjc), then from the case surfaces to air (1 / (h × A exposed)). For a small part in still air this is hundreds of C/W.",
      "Board path: junction to the board (R j-board, from Rjb or half of Rjc), then spreading into the board (R spread), into one board node that all parts share. The board then loses all its heat from its two faces to air."],
    "Spreading resistance: heat enters the board through a small footprint and must spread sideways before the whole share of the board can lose it. The page uses the model of Lee, Song, Au and Moran (1995): squares are turned into circles of the same area, and the result depends on the footprint over the share (ε), the board thickness over the share radius (τ) and how easily the faces lose heat compared with how easily the board spreads it (Bi). More copper (a higher in-plane k) lowers it most.",
    "The shared board node: every part heats the same board, so the board rises above the air by all the heat together times R board to air. A part next to a hot neighbour can even take heat in from the board (a negative share on tab 5).",
    "What the network leaves out: the real board is hotter under the parts than at its edges, and parts heat each other locally. The simulation shows that. A good hand check lands within a few degrees of the simulation for the parts that matter; a big gap means an input differs between the two.",
  ] },
  { id: "spreaders", title: "Spreaders and TIMs", body: [
    "A spreader is a metal plate over (or under) the board, touching the hot chips through a thermal interface material (TIM). It helps when most of a part's heat leaves from its top (tab 5 says so), or when the board alone cannot get rid of the total heat.",
    "Each covered chip needs the gap between its top and the plate filled. The plate sits on the tallest covered chip plus its TIM; lower chips get a pedestal, a raised block on the plate, so their TIM stays thin. A pedestal thinner than the smallest pedestal setting is not worth making: the TIM fills the whole gap instead.",
    "The TIM resistance is R = t / (k × A). A 0.2 mm pad with k = 3 on a 10 × 10 mm chip is 0.67 C/W; a 1 mm pad would be 3.3 C/W. Thin and well squeezed beats a high datasheet k.",
    "TIM mode Body builds the TIM as a solid part in the model; Contact R puts its resistance as a contact between chip and plate, with no part. Use Body above about 0.3 mm.",
    "The hand check with a plate (tab 7) sends all heat of the parts under the plate through the stack Rjc, TIM1, optional local spreader, TIM2 and spreading in the plate, then from the plate to air. If most of the rise is the plate to air, a bigger plate or more air helps; if most is in the stack, a thinner or better TIM helps.",
    "Tall parts the plate does not cover get a clearance cut (a pocket) so the plate clears them by the Z clearance; the cut is the XY clearance bigger than the part on each side.",
  ] },
  { id: "simulation", title: "Simulation", body: [
    "Run Thermal Study makes a SolidWorks Simulation steady state thermal study of the thermal assembly: one box per part on the PCB plate, plus the spreader and TIMs. It solves conduction inside the solids; the air is not modelled, only an h and an emissivity on the outer faces, the same as the hand check uses.",
    ["Materials: the PCB is orthotropic, with its in-plane and through-plane k from the copper stack (tab 3). Each chip on tab 8 is a block whose k z makes its resistance top to bottom equal Rjc: k z = H / (Rjc × A). Two-resistor chips are three layers (Rjb at the bottom, a hot core, Rjc at the top). Parts without heat get one generic k.",
      "Heat: each part's Pd (tab 1) as a heat power on its body.",
      "Contacts: bonded unless a contact resistance is typed on tab 8; chips in TIM mode Contact R get their TIM resistance.",
      "Mesh: the global size, finer on chips with heat and on TIMs. A chip should be at least three elements thick."],
    "Energy balance: the heat leaving the model must equal the heat put in. Read Results works out the error; above 2 % the run is not converged or a heat load is missing, and its temperatures should not be trusted.",
    "Why hand and simulation differ: the simulation sees the real spreading in the board and the hot spots under parts; the hand check uses one board temperature. Small parts near big ones run hotter in the simulation. Differences above about 5 C usually come from an input that differs: h or emissivity, the PCB k, a body left out of the model, or a chip block k.",
    "Flow Simulation (Run Flow Simulation, button 7) models the air too. The macro makes the Flow project from the PCBThermal_Flow template the first time (external air, gravity, radiation) and on every run sets its materials, a volume heat source and max/avg goals per chip, PCB and spreader goals and the domain, then the heat and ambient from the page. It solves and reads the goals. Since the Flow plan (F4, F5) Flow gets the same k values as the Thermal study: chips with an Rjc as calibrated blocks, the PCB k of the stack, the TIM and plate k, the emissivities and the contact resistances, as user materials and radiative surfaces named after their values. Make the ones tab 8 lists once in Flow's Engineering Database. Tab 8 also sets the mesh (level, local mesh, mesh check) and the minimum travels; tab 9 shows the run health before the temperatures.",
  ] },
  { id: "tabs", title: "Each tab", body: [] },
  { id: "solidworks", title: "SolidWorks steps", body: [
    "Every macro is a button on the PCB Thermal tab in SolidWorks (made once by SetupToolbar.cmd). Each asks which project the first time: pick <prefix>_project.json in the board folder. Next time Yes is enough. On the page, each step has a card that lists the checks, which button to press and what to look at after; press I pressed it and the card waits until the macro has written its file.",
  ] },
  { id: "trouble", title: "Troubleshooting", body: [] },
  { id: "glossary", title: "Glossary", body: [] },
];

export const TAB_HELP = {
  project: { you: "Choose the folder, name the board, set the ambient and the flag limits.", out: "The next step to do and the state of every step.",
    mistakes: ["Using room temperature as ambient. Use the air inside the enclosure at the worst case.", "Two boards in one folder. One folder per board keeps each study self-contained."] },
  parts: { you: "One row per heat source: RefDes, class, datasheet values, and the values its power method needs.", out: "Pd per load and per part, the Rja estimate of Tj, the margin and a flag.",
    mistakes: ["Typing the delivered power of a regulator instead of its loss.", "Counting the same watts twice (a regulator's loss and the full input power of its load).", "Rja of a part on a thin, crowded board: the estimate is optimistic."] },
  import: { you: "Name the STEP file, set the import settings, run Import STEP in SolidWorks.", out: "The PCB size and one body per component, on a drawing of the board.",
    mistakes: ["A STEP with the parts merged into one body: export it with one body per component.", "Small passives in the import: raise Smallest body."] },
  link: { you: "Link each part with Pd to its body; include the bodies to build; correct sizes; set the PCB stack.", out: "Pd and power density per body, the PCB k in plane and through.",
    mistakes: ["Leaving a part with Pd without a body: its heat is not in the model.", "Including hundreds of tiny bodies: the model gets slow for no gain."] },
  triage: { you: "Read only.", out: "The parts ranked by Pd and by power density, the rail budget and the problems to fix first.",
    mistakes: ["Skipping the rail budget check against the measured input power."] },
  bare: { you: "Set h and emissivity of the PCB and parts, the surface rise for radiation.", out: "The board temperature, Tj of every part with its two heat paths, FAIL or Tight, and the comparison with a run.",
    mistakes: ["An h far from the case you study (a fan on a still-air board).", "Ignoring a FAIL because the Rja estimate on tab 1 looked fine."] },
  build: { you: "Run Build Board in SolidWorks.", out: "What was built, failed or skipped, overlapping boxes, and the SolidWorks files of the project.",
    mistakes: ["Building before saving the page: the macro builds from the last save.", "Leaving overlaps: the material is counted twice."] },
  spreader: { you: "Add a plate, cover the chips, set TIMs; try the hand check with a plate; run Build Spreader and Sync CAD.", out: "Plate face height, pedestals, TIM resistance per chip, clearance cuts; Tj per part with the plate.",
    mistakes: ["A thick TIM to fill a big gap: use a pedestal.", "Covering a chip on the other side of the board than the plate."] },
  sim: { you: "Set the study type, mesh, surfaces of the spreader, contacts and the chip models; run the study.", out: "k z and k xy of every chip block, and the SolidWorks card for the study.",
    mistakes: ["Different h on tab 5 and tab 8: the hand check and the study cannot agree.", "A fine mesh larger than the TIM thickness."] },
  results: { you: "Read only; pick a run.", out: "Peak and average per part against Tj max and the hand check, the energy balance, the plot, every run.",
    mistakes: ["Trusting a run with an energy balance error above 2 %.", "Reading an old run after the inputs changed: the card says what changed."] },
  changes: { you: "Tick every change once you know why it happened; take sizes changed in SolidWorks.", out: "Every move, resize and failed check the macros found.",
    mistakes: ["Ticking all without looking: a wrong import then goes into the model unseen."] },
};

export const MACRO_HELP = {
  ImportStep: { changes: "Writes the import to the cad file: the PCB and every body with its box. The previous import is kept, and a change record is written for every body that moved, resized, appeared or went.", check: "The PCB size, the number of bodies, and the board drawing on tab 2." },
  BuildBoard: { changes: "Makes or updates the thermal assembly and its part files in the CAD folder. Writes the build to the cad file, with failures, overlaps and change records.", check: "Failures first, then overlaps, then that every heated part has a built body (tab 6)." },
  BuildSpreader: { changes: "Completes the spreader part (plate, pedestals, clearance cuts, named faces) and builds the TIM parts. Writes the plate build to the cad file.", check: "Problems on tab 7, then that each covered chip got its pedestal and TIM." },
  SyncFromCAD: { changes: "Reads the assembly only. Writes what it found (boxes that differ, the plate as built, failed checks) to the cad file.", check: "Tab 10: take the SolidWorks sizes or build again; tick the failed checks once fixed." },
  SetupSim: { changes: "Adds a new study to the assembly and a new run, with a copy of the inputs, to the results file. Never changes an older study.", check: "The message: solved, or the problems that stopped it. Then press Read Results." },
  RunFlow: { changes: "Makes or updates the Flow project <prefix>Flow (materials, heat sources, goals, domain, initial temperature), solves it, and adds a run to the results file.", check: "Every chip with heat has a source (the message lists the ones without), the goals were read, and the run health on tab 9 is OK (no part left unmeshed, energy balance within 2 %)." },
  ReadResults: { changes: "Puts the temperatures, energy balance and the plot picture into the run in the results file.", check: "Tab 9: the energy balance first, then each part against Tj max and the hand check." },
};

export const TROUBLE = [
  ["\"Work on this project?\" names the wrong board", "Answer No and pick the right <prefix>_project.json. The macros remember the last project, not the open model."],
  ["Build Board refuses: changes not ticked", "Open tab 10, look at each change, tick it, save, run again. The macro will not build on top of unchecked changes."],
  ["Build Board refuses: the page has not saved since the last import", "Press Save on the page. The macros read what the page worked out on its last save."],
  ["The sketch was not made (Build Board)", "A body is too small or flat for SolidWorks to sketch. Check its size on tab 3; raise Smallest body, or leave it out."],
  ["The page does not see the macro finish", "The project must be opened from its folder (not by hand). Check that the macro wrote the file named in its message; the page reads it every 2 s."],
  ["The page says a file was changed outside the page", "Someone (or a backup tool) wrote the project file. Save asks before replacing it; Open a backup on tab 0 goes back to an earlier day."],
  ["Wrong workbook or old workbook", "Old workbooks still work with the macros: pick the .xlsx when a macro asks. To move to the page, use Make it from an old workbook on tab 0."],
  ["Trend Micro or another virus scanner blocks the macro", "Ask IT to allow PCBThermal.swp in its folder. The macros only read and write files in the project folder and %APPDATA%\\PCBThermal."],
  ["Missing material or k in the study (errors 16, 19, 20)", "A k is zero or blank: check the PCB stack (tab 3), k of parts without heat (tab 5) and the plate k (tab 7). Save and run the study again."],
  ["Energy balance error above 2 %", "Refine the mesh on tab 8, check that every heated part has a body in the model, and run again."],
  ["Run Flow Simulation finds no sources or goals", "Name each volume source and goal with the RefDes in it (for example 'U1' and 'U1 max'). See the Simulation section."],
];
