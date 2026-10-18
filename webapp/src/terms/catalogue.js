// The term catalogue: what every value on the page is, for a graduate engineer. One entry per term id (the
// same ids as src/calc/formulas.js). Hover shows label, unit, kind and `short`; the Explain panel adds `long`,
// the formula with the user's numbers, where the inputs come from and what uses the value.
//
// kind: input (you type it), calculated (the page works it out), cad (Import STEP or Build Board wrote it),
//       run (a simulation run wrote it), constant.
// range: typical values [low, high]; the page warns, it does not refuse, outside it.
// Started from the header notes of the workbook template (docs/webapp/workbook_terms.json), rewritten.

const T = {};
const def = (id, label, sym, unit, kind, short, long, extra = {}) => { T[id] = { id, label, sym, unit, kind, short, long, ...extra }; };

// ---------- constants ----------
def("const.sigma", "Stefan-Boltzmann constant", "σ", "W/(m² K⁴)", "constant", "Physical constant of thermal radiation, 5.67 × 10⁻⁸.",
  "A black surface at absolute temperature T radiates σ T⁴ watts per square metre. Real surfaces radiate ε times that.");

// ---------- board ----------
def("board.name", "Board name", "", "", "input", "Name of the board. Used in reports and as the start of the SolidWorks file names.");
def("board.stepFile", "STEP file", "", "", "input", "The board's STEP file, in the project folder. Import STEP reads it.",
  "Export the board from the ECAD tool with its parts as separate bodies (one per component). The name is relative to the project folder.");
def("board.ambient", "Ambient", "Ta", "C", "input", "Air temperature around the board inside its enclosure, at the worst case.",
  "Not room temperature: the air inside a closed box is warmer than the room. Use the hottest operating case (for example 40 C inside a case in a 25 C room). Every hand calculation starts from it, and the simulations use it as the ambient.",
  { range: [-40, 85], help: "heat" });
def("board.measuredInputPower", "Measured board input power", "P in", "W", "input", "Power measured at the board's supply input. Optional: checks the sum of all part Pd.",
  "Measure supply voltage × current at the input connector in the load case you study. If the sum of part Pd is far below it, a heat source is missing from tab 1 (or a regulator's loss is counted twice).",
  { range: [0, 500] });
def("board.notes", "Notes", "", "", "input", "Assumptions of this study: load case, enclosure, airflow, where the numbers came from.",
  "Write them down so the next engineer can trust or redo the study. For example: full CPU load, closed plastic case, still air, Pd of U3 from the datasheet efficiency curve at 2 A.");
def("board.totalPd", "Board Pd", "Σ Pd", "W", "calculated", "Sum of the power dissipated by every load of every part.",
  "All the heat the board must get rid of. Compare it with the measured input power on tab 4.");

// ---------- flag limits ----------
def("limit.criticalMargin", "CRITICAL margin limit", "", "C", "input", "A part is CRITICAL when its margin is below this.", undefined, { range: [0, 50] });
def("limit.criticalDensity", "CRITICAL density limit", "", "W/cm²", "input", "A part is CRITICAL when its power density is above this.", undefined, { range: [0, 20] });
def("limit.criticalShare", "CRITICAL share limit", "", "", "input", "A part is CRITICAL when its share of board Pd is above this (0 to 1).", undefined, { range: [0, 1] });
def("limit.watchMargin", "Watch margin limit", "", "C", "input", "A part is Watch when its margin is below this (and it is not CRITICAL).", undefined, { range: [0, 80] });
def("limit.watchDensity", "Watch density limit", "", "W/cm²", "input", "A part is Watch when its power density is above this.", undefined, { range: [0, 20] });
def("limit.watchShare", "Watch share limit", "", "", "input", "A part is Watch when its share of board Pd is above this (0 to 1).", undefined, { range: [0, 1] });

// ---------- parts ----------
def("part.refDes", "RefDes", "", "", "input", "Reference designator from the schematic or BOM (U1, L3, R12). Links the part to its body.",
  "Type it exactly as in the STEP file names, so Import STEP can suggest the body. Each part has its own RefDes; renaming it keeps its body, spreader and simulation links.");
def("part.partNumber", "Part number", "", "", "input", "Manufacturer part number, so the datasheet can be found again.");
def("part.class", "Class", "", "", "input", "What kind of part it is. The class picks the default way its power loss is worked out.",
  "Processor, memory and interface ICs: digital load (V × I). LDO: linear regulator. Buck, boost, PMIC: switching regulator. Inductors, sense resistors, MOSFETs, connectors: I²R. Other: type the Pd yourself.");
def("part.package", "Package", "", "", "input", "Package name from the datasheet (QFN-16, FCBGA). For reference only: the size used comes from the body.");
def("part.tjMax", "Tj max", "Tj max", "C", "input", "Highest junction temperature the datasheet allows.",
  "Found under Absolute maximum ratings or Recommended operating conditions. Use the operating limit, not the storage limit. For parts without a junction (inductors, connectors) use the hot spot or case limit. Processors often throttle below Tj max: use the throttle point if that matters.",
  { range: [60, 175] });
def("part.rja", "Junction to ambient resistance", "Rja", "C/W", "input", "How many degrees the junction rises above the air for every watt, on the JEDEC test board.",
  "Also written θJA. Measured on a standard test board in still air, not on your board: use it only to rank parts and for the first Tj estimate. A real board with more copper usually does better, a crowded enclosure worse.",
  { range: [1, 300] });
def("part.rjc", "Junction to case resistance", "Rjc", "C/W", "input", "How many degrees the junction is above its case top for every watt it makes.",
  "Also written θJC(top). The datasheet thermal table lists it. Some datasheets give θJC(bottom) for parts with an exposed pad: that one is the path to the board, so type it as Rjb. Rja is not Rjc.",
  { range: [0.05, 50] });
def("part.rjb", "Junction to board resistance", "Rjb", "C/W", "input", "How many degrees the junction is above the board under it for every watt. Optional.",
  "Also written θJB. When given, the part uses the two-resistor model: Rjc to its top, Rjb to the board. Without it, the part is a block with half of Rjc each way.",
  { range: [0.1, 100] });
def("part.taOverride", "Ta override", "", "C", "input", "Local air temperature for this part, if it sits in a hotter spot than the board ambient. Optional.", undefined, { range: [-40, 125] });
def("part.pd", "Pd", "Pd", "W", "calculated", "Power the part turns into heat: the sum of its loads.",
  "Pd is the electrical power lost inside the part. For a regulator it is the input power minus what it delivers; for a digital chip it is all the power it draws.");
def("part.ta", "Ta", "Ta", "C", "calculated", "Air temperature the part sees: its Ta override, else the board ambient.");
def("part.area", "Footprint area", "A", "mm²", "calculated", "L × W of the part's body, from tab 3.");
def("part.rise", "Rise", "ΔT", "C", "calculated", "Junction rise above the air on the JEDEC board: Pd × Rja.");
def("part.tjEst", "Tj estimate", "Tj est", "C", "calculated", "First estimate of the junction temperature, from Rja.",
  "Ta + Pd × Rja. Good for ranking which parts to worry about, not for sign-off: Rja is for the JEDEC test board, not yours. The hand check on tab 5 and the simulation give better numbers.");
def("part.margin", "Margin", "", "C", "calculated", "Tj max minus Tj estimate. Negative means over the limit on this estimate.");
def("part.density", "Power density", "", "W/cm²", "calculated", "Pd per cm² of footprint. High density means the heat has little area to leave through.",
  "Above about 1 W/cm² a part on a plain board usually needs copper area, vias or a heatsink. Small parts with modest Pd (inductors) can have high density.");
def("part.share", "Share of board Pd", "", "", "calculated", "This part's Pd as a share of the whole board's Pd.");
def("part.flag", "Flag", "", "", "calculated", "CRITICAL, Watch or Minor, from the margin, density and share limits. No Pd: Pd not worked out yet.");
def("part.why", "Why flagged", "", "", "calculated", "Which limits the part crossed, with the numbers, and which checks could not run.");
def("part.rankPd", "Rank by Pd", "", "", "calculated", "Position by Pd, largest first.");
def("part.rankDensity", "Rank by density", "", "", "calculated", "Position by power density, largest first.");

// ---------- loads ----------
def("load.label", "Load", "", "", "input", "What this heat source is, e.g. 'buck1 > VDD_CORE'. Needed when a part has several loads.");
def("load.rail", "Rail", "", "", "input", "Supply rail the load draws from, typed the same way every time. Tab 4 groups power by rail.");
def("load.method", "Method", "", "", "input", "How Pd is worked out for this load. Blank uses the default of the part class.",
  "LDO: (Vin - Vout) × Iout. Switching regulator: Pout × (1/η - 1), or Pin - Pout. Digital load: Vin × Iin. Resistive: I² × R. Manual entry: type Pd measured.");
def("load.vin", "Vin", "Vin", "V", "input", "Input voltage: the rail voltage for a load, the input voltage for a regulator.", undefined, { range: [0, 60] });
def("load.iin", "Iin", "Iin", "A", "input", "Input current: what the load draws from the rail. Resistive parts: the current through them.", undefined, { range: [0, 50] });
def("load.vout", "Vout", "Vout", "V", "input", "Output voltage. Regulators only.", undefined, { range: [0, 60] });
def("load.iout", "Iout", "Iout", "A", "input", "Output (load) current. Regulators only; resistive parts can use it instead of Iin.", undefined, { range: [0, 50] });
def("load.eff", "Efficiency", "η", "", "input", "Efficiency of a switching regulator at this load, 0 to 1 (0.88 = 88 %).",
  "Read it from the datasheet efficiency curve at your output current and input voltage, not the headline peak. Not needed when Vin and Iin are typed (then Pd = Pin - Pout).",
  { range: [0.5, 0.99] });
def("load.r", "R", "R", "ohm", "input", "Resistance for I²R: inductor DCR, MOSFET Rds(on) when hot, resistor value.",
  "Use the value at operating temperature: copper DCR rises about 0.4 % per C, Rds(on) often 1.5× from 25 C to 100 C.", { range: [0, 10] });
def("load.pdMeas", "Pd measured", "", "W", "input", "A measured or datasheet Pd. When typed, it replaces the calculated Pd.", undefined, { range: [0, 100] });
def("load.methodUsed", "Method used", "", "", "calculated", "The method set on the load, else the default of the part class.");
def("load.pin", "Pin", "Pin", "W", "calculated", "Input power: Vin × Iin.");
def("load.pout", "Pout", "Pout", "W", "calculated", "Output power: Vout × Iout.");
def("load.pdCalc", "Pd calc", "", "W", "calculated", "Pd worked out by the method from the electrical values.");
def("load.pd", "Pd of the load", "Pd", "W", "calculated", "Pd measured if typed, else Pd calc.");
def("load.problem", "Why no Pd", "", "", "calculated", "What is missing before Pd can be worked out.");

// ---------- rails ----------
def("rail.name", "Rail", "", "", "input", "Supply rail name, as typed on the loads.");
def("rail.parts", "Parts on the rail", "", "", "calculated", "Number of parts with a load on this rail.");
def("rail.pd", "Rail Pd", "", "W", "calculated", "Heat made by all loads on this rail.");
def("rail.share", "Share of board Pd", "", "", "calculated", "This rail's Pd as a share of the board Pd.");
def("rails.sum", "Sum of rails", "", "W", "calculated", "Pd of every load that has a rail.");
def("rails.noRail", "Loads with no rail", "", "W", "calculated", "Pd of loads with no rail typed.");
def("rails.measured", "Measured input", "", "W", "calculated", "The measured board input power from tab 0.");
def("rails.unaccounted", "Not accounted for", "", "W", "calculated", "Measured input minus board Pd. Large means a heat source is missing, or the measurement was at another load.",
  "Some input power leaves the board as useful output (USB power to a device, LEDs, a fan), so a gap is not always a missing part.");
def("rails.unaccountedFrac", "Share not accounted for", "", "", "calculated", "Not accounted for, as a share of the measured input.");

// ---------- triage ----------
def("triage.parts", "Parts", "", "", "calculated", "Number of parts on tab 1.");
def("triage.critical", "CRITICAL", "", "", "calculated", "Parts flagged CRITICAL: look at these first.");
def("triage.watch", "Watch", "", "", "calculated", "Parts flagged Watch.");
def("triage.minor", "Minor", "", "", "calculated", "Parts below every Watch limit.");
def("triage.noPd", "No Pd", "", "", "calculated", "Parts whose Pd cannot be worked out yet.");
def("triage.noRja", "No Rja", "", "", "calculated", "Parts without Rja: their margin is not checked.");
def("triage.noBody", "No body", "", "", "calculated", "Parts without a body: their density is not checked.");

// ---------- import and PCB ----------
def("import.model", "Model", "", "", "cad", "Name of the STEP model Import STEP read.");
def("import.date", "Import date", "", "", "cad", "When Import STEP ran.");
def("import.bodies", "Bodies", "", "", "cad", "Number of bodies Import STEP found (one per part in the STEP).");
def("import.frame", "Board frame", "", "", "cad", "How the board's axes sit in the STEP model. X along the board length, Y along its width, Z up from the PCB bottom face.");
def("pcb.l", "PCB length", "L", "mm", "cad", "Board size along X, from Import STEP (or the override on tab 3).");
def("pcb.w", "PCB width", "W", "mm", "cad", "Board size along Y, from Import STEP (or the override on tab 3).");
def("pcb.h", "PCB thickness", "H", "mm", "cad", "Board thickness, from Import STEP (or the override on tab 3).", undefined, { range: [0.4, 6] });
def("stack.layers", "Copper layers", "", "", "input", "Number of copper layers in the PCB.", undefined, { range: [1, 32] });
def("stack.cuT", "Copper thickness", "t Cu", "mm", "input", "Thickness of one copper layer. 1 oz = 0.035 mm.", undefined, { range: [0.009, 0.21] });
def("stack.coverage", "Copper coverage", "", "", "input", "Average share of each layer that is copper (0 to 1). Planes about 0.9, signal layers about 0.2.", undefined, { range: [0.05, 1] });
def("stack.cuK", "Copper k", "k Cu", "W/(m K)", "input", "Thermal conductivity of copper, about 390.", undefined, { range: [300, 400] });
def("stack.dielK", "Dielectric k", "k diel", "W/(m K)", "input", "Thermal conductivity of the board laminate. FR-4 about 0.3.", undefined, { range: [0.1, 3] });
def("pcb.cuTotal", "Copper total", "", "mm", "calculated", "All copper layers stacked: layers × copper thickness.");
def("pcb.cuFraction", "Copper fraction", "f", "", "calculated", "Share of the board thickness that is copper, after coverage.");
def("pcb.kxy", "PCB k in plane", "k xy", "W/(m K)", "calculated", "How well the board spreads heat sideways. The thin copper layers carry most of it.",
  "A typical 4 layer board is 20 to 50 W/(m K) in plane but only about 0.4 through its thickness: heat spreads far more easily than it crosses the board.");
def("pcb.kz", "PCB k through", "k z", "W/(m K)", "calculated", "How well heat crosses the board thickness. The laminate limits it.");

// ---------- bodies ----------
def("body.include", "Include", "", "", "input", "Build this body in the thermal model. Skip parts that make no heat and do not matter for air or contact (labels, tiny passives).");
def("body.name", "Body", "", "", "cad", "Name of the body in the STEP file.");
def("body.side", "Side", "", "", "cad", "Which face of the PCB the body sits on: Top, Bottom, Board, or Through (passes through it).");
def("body.refDes", "Part", "", "", "input", "The part this body belongs to. Its heat goes into this body.");
def("body.l", "L", "L", "mm", "cad", "Body size along X (from Import STEP, or the override).");
def("body.w", "W", "W", "mm", "cad", "Body size along Y.");
def("body.h", "H", "H", "mm", "cad", "Body height along Z.");
def("body.xMin", "X min", "", "mm", "cad", "Lowest X of the body.");
def("body.yMin", "Y min", "", "mm", "cad", "Lowest Y of the body.");
def("body.zMin", "Z min", "", "mm", "cad", "Lowest Z of the body, from the PCB bottom face. A top part starts at the board thickness.");

// ---------- air and surfaces ----------
def("surf.pcbTopH", "h, PCB top", "h", "W/(m² K)", "input", "Convection coefficient of the PCB top face.",
  "Still air about 5 to 10, a gentle fan 20 to 50. One value for the whole face; the simulation uses the same.", { range: [1, 200] });
def("surf.pcbTopEps", "Emissivity, PCB top", "ε", "", "input", "How well the PCB top radiates heat, 0 to 1. Solder mask about 0.9.", undefined, { range: [0.02, 1] });
def("surf.pcbBottomH", "h, PCB bottom", "h", "W/(m² K)", "input", "Convection coefficient of the PCB bottom face. Usually lower than the top in still air.", undefined, { range: [1, 200] });
def("surf.pcbBottomEps", "Emissivity, PCB bottom", "ε", "", "input", "How well the PCB bottom radiates heat, 0 to 1.", undefined, { range: [0.02, 1] });
def("surf.partsH", "h, parts", "h", "W/(m² K)", "input", "Convection coefficient of the part surfaces.", undefined, { range: [1, 200] });
def("surf.partsEps", "Emissivity, parts", "ε", "", "input", "How well the part surfaces radiate, 0 to 1. Black mould compound about 0.9, bare metal 0.1.", undefined, { range: [0.02, 1] });
def("hand.surfaceRise", "Surface rise for radiation", "ΔT", "K", "input", "Guess of how much warmer the surfaces are than the air. Radiation h depends on it a little.", undefined, { range: [0, 100] });
def("simset.kPartsWithoutHeat", "k of parts without Rjc", "k parts", "W/(m K)", "input", "Conductivity given to bodies that have no Rjc. Also used as the block k when a part has no Rjc.", undefined, { range: [0.2, 400] });

// ---------- bare board ----------
def("bb.area", "Board area", "A board", "mm²", "calculated", "L × W of the PCB. Each part gets a share of it.");
def("bb.hRadTop", "h rad, PCB top", "h rad", "W/(m² K)", "calculated", "Radiation from the PCB top written as an h, so it adds to convection.");
def("bb.hRadBottom", "h rad, PCB bottom", "h rad", "W/(m² K)", "calculated", "Radiation from the PCB bottom as an h.");
def("bb.hRadParts", "h rad, parts", "h rad", "W/(m² K)", "calculated", "Radiation from the part surfaces as an h.");
def("bb.hPcb", "h of the PCB", "h PCB", "W/(m² K)", "calculated", "Convection and radiation of both board faces together, on the board area.");
def("bb.hParts", "h of the parts", "h parts", "W/(m² K)", "calculated", "Convection plus radiation of a part's exposed surface.");
def("bb.rBoardAir", "Board to air", "R board-air", "C/W", "calculated", "Resistance from the board to air, both faces: 1 / (h × A).");
def("bb.totalPd", "Pd total", "Pd total", "W", "calculated", "Pd of all parts together.");
def("bb.rise", "Board rise", "ΔT board", "C", "calculated", "How much the board sits above the air, all parts heating it at once.");
def("bb.boardC", "Board temperature", "T board", "C", "calculated", "Ambient plus the board rise. One number for the whole board: the hand check treats it as one node.",
  "The real board is hotter under the parts and cooler at the edges. The simulation shows that; the hand check gives the average level.");
def("bb.qBoard", "Heat into the board", "Q board", "W", "calculated", "Of all the Pd, how much goes down into the board (the rest leaves through the part tops).");
def("bb.fail", "FAIL", "", "", "calculated", "Parts whose hand check Tj is above Tj max.");
def("bb.tight", "Tight", "", "", "calculated", "Parts with a positive margin below the CRITICAL margin limit.");
def("bb.compareRun", "Run compared", "", "", "run", "The simulation run the hand check is compared with: the one set on tab 5, else the newest.");
def("bb.compareCheck", "Comparison check", "", "", "calculated", "Whether that run is a fair match for a bare board hand check.");
def("bb.meanGap", "Mean gap", "", "C", "calculated", "Average of hand Tj minus simulated Tj. Positive: the hand check is warmer (on the safe side).");
def("bb.maxGap", "Largest gap", "", "C", "calculated", "Largest difference, either way, between hand and simulated Tj.");
def("bbp.pd", "Pd", "Pd", "W", "calculated", "Pd of the part, from tab 1.");
def("bbp.l", "L", "L", "mm", "cad", "Body length of the part.");
def("bbp.w", "W", "W", "mm", "cad", "Body width of the part.");
def("bbp.h", "H", "H", "mm", "cad", "Body height of the part.");
def("bbp.rjcUsed", "Rjc used", "Rjc used", "C/W", "calculated", "Rjc from the datasheet, or a block estimate when there is none.");
def("bbp.model", "Chip model", "", "", "calculated", "Two-resistor (Rjc and Rjb), Block (Rjc only), or the part k when there is no Rjc.");
def("bbp.rJTop", "Junction to top", "R j-top", "C/W", "calculated", "Resistance from the junction to the part's top surface.");
def("bbp.rJBoard", "Junction to board", "R j-board", "C/W", "calculated", "Resistance from the junction down to the board.");
def("bbp.exposed", "Exposed area", "A exposed", "mm²", "calculated", "Top and four sides of the body: the area that loses heat to air.");
def("bbp.rSurfaceAir", "Surface to air", "R surface-air", "C/W", "calculated", "Resistance from the part surface to air: 1 / (h × A exposed).");
def("bbp.rTop", "Top path", "R top", "C/W", "calculated", "Junction to air through the part top: R j-top + R surface-air.",
  "Usually large (hundreds of C/W) for a small part in still air: most heat of a bare-board part goes into the board, not out of its top.");
def("bbp.footprint", "Footprint", "A footprint", "mm²", "calculated", "L × W of the body: where its heat enters the board.");
def("bbp.share", "Board share", "A share", "mm²", "calculated", "The part of the board area this part heats, in proportion to its Pd.");
def("bbp.eps", "Area ratio", "ε", "", "calculated", "√(footprint / share): 1 means the part covers its whole share.", undefined, { help: "spreading" });
def("bbp.tau", "Relative thickness", "τ", "", "calculated", "Board thickness over the radius of the part's share.");
def("bbp.bi", "Biot number", "Bi", "", "calculated", "How easily heat leaves the board face compared with how easily it spreads in the board.");
def("bbp.lambda", "λ", "λ", "", "calculated", "A term of the spreading model, from ε.");
def("bbp.phi", "Φ", "Φ", "", "calculated", "A term of the spreading model, from λ, τ and Bi.");
def("bbp.rSpread", "Spreading resistance", "R spread", "C/W", "calculated", "Resistance of heat spreading out from the footprint into the board.",
  "Heat enters the board through a small footprint and must spread sideways before the whole share can lose it. The narrower the footprint and the thinner the copper, the larger this is. Model: Lee, Song, Au and Moran (1995), squares turned into circles of the same area.", { help: "spreading" });
def("bbp.rBoard", "Board path", "R board", "C/W", "calculated", "Junction to the board node: R j-board + R spread.");
def("bbp.f", "Board fraction", "f", "", "calculated", "Share of the part's heat that would go into the board if the board were at ambient.");
def("bbp.c", "Back conductance", "C", "W/K", "calculated", "How much heat flows back from the board up through the part to air, per degree of board rise.");
def("bbp.fPd", "f × Pd", "", "W", "calculated", "Heat the part pushes into the board.");
def("bbp.qBoard", "Into the board", "Q board", "W", "calculated", "Heat this part puts into the board, after the shared board has warmed.");
def("bbp.qTop", "Out of the top", "Q top", "W", "calculated", "Heat that leaves through the part top.");
def("bbp.tj", "Tj, hand check", "Tj", "C", "calculated", "Junction temperature from the bare board network.",
  "Better than the Rja estimate because it uses your board size, copper and air. Still a model: compare it with the simulation.");
def("bbp.margin", "Margin", "", "C", "calculated", "Tj max minus the hand check Tj.");
def("bbp.status", "Status", "", "", "calculated", "FAIL, Tight or OK from the margin, or what is missing.");
def("bbp.simC", "Tj, simulation", "Tj sim", "C", "run", "Peak temperature of the part in the run compared.");
def("bbp.gap", "Gap", "", "C", "calculated", "Hand Tj minus simulated Tj.");

// ---------- build ----------
def("build.assembly", "Assembly", "", "", "cad", "The SolidWorks assembly Build Board made.");
def("build.date", "Build date", "", "", "cad", "When Build Board ran.");
def("build.mode", "Build mode", "", "", "cad", "New: built from nothing. Update: only what changed since the last build.");
def("build.wanted", "Wanted", "", "", "cad", "Bodies the build was asked for (included bodies plus the PCB).");
def("build.added", "Added", "", "", "cad", "Bodies built new.");
def("build.moved", "Moved", "", "", "cad", "Bodies moved to a new position.");
def("build.resized", "Resized", "", "", "cad", "Bodies built again with a new size.");
def("build.suppressed", "Suppressed", "", "", "cad", "Bodies taken out of the model because they are no longer included.");
def("build.broughtBack", "Brought back", "", "", "cad", "Suppressed bodies included again.");
def("build.failed", "Failed", "", "", "cad", "Bodies that could not be built. Each has a reason below.");
def("build.overlaps", "Overlaps", "", "", "cad", "Pairs of bodies that overlap. An overlap counts the material twice; fix the boxes on tab 3.");
def("build.included", "Included bodies", "", "", "calculated", "Bodies with Include ticked on tab 3.");
def("buildbody.status", "Build status", "", "", "cad", "What Build Board did with this body.");

// ---------- spreader ----------
def("plate.enabled", "Plate enabled", "", "", "input", "Build this plate and use it in the study.");
def("plate.material", "Plate material", "", "", "input", "Material of the spreader plate.");
def("plate.k", "Plate k", "k", "W/(m K)", "input", "Thermal conductivity of the plate. Aluminium 6061 about 167, copper about 390.", undefined, { range: [10, 400] });
def("plate.thickness", "Plate thickness", "t", "mm", "input", "Thickness of the plate.", undefined, { range: [0.3, 10] });
def("plate.minPedestal", "Smallest pedestal", "", "mm", "input", "Thinner pedestals are not made: the TIM fills the gap instead.", undefined, { range: [0, 3] });
def("plate.heightMode", "Height mode", "", "", "input", "Auto: the plate sits on the tallest covered chip. Manual: the height typed in.");
def("plate.heightManual", "Manual height", "", "mm", "input", "Height of the plate face, from the PCB bottom face, when the height mode is Manual.");
def("plate.clearanceZ", "Z clearance", "", "mm", "input", "Air gap kept between the plate and parts it does not cover.");
def("plate.faceZ", "Plate face height", "Face", "mm", "calculated", "Height of the plate face that meets the chips, from the PCB bottom face.");
def("plate.covered", "Chips covered", "", "", "calculated", "Chips the plate touches through a TIM.");
def("plate.cuts", "Clearance cuts", "", "", "calculated", "Pockets cut into the plate so it clears tall parts it does not cover.");
def("pchip.refDes", "RefDes", "", "", "input", "The chip under the plate.");
def("pchip.covered", "Covered", "", "", "input", "The plate touches this chip through a TIM.");
def("pchip.heightMode", "Height mode", "", "", "input", "Pedestal: a raised block on the plate plus a thin TIM. Thick TIM: the TIM fills the whole gap.");
def("pchip.timTSet", "TIM thickness typed", "TIM t", "mm", "input", "TIM thickness you want between chip and pedestal.", undefined, { range: [0.02, 3] });
def("pchip.chipFace", "Chip face", "", "mm", "calculated", "Height of the chip face toward the plate (top of a top chip, bottom of a bottom chip).");
def("pchip.zNeed", "Height needed", "z", "mm", "calculated", "Where the plate face would have to be for this chip: chip face plus its TIM.");
def("pchip.gap", "Gap", "", "mm", "calculated", "Space between the chip face and the plate face.");
def("pchip.timT", "TIM", "", "mm", "calculated", "TIM thickness built: the typed one, or the whole gap.");
def("pchip.pedestal", "Pedestal", "", "mm", "calculated", "Height of the block on the plate that fills the rest of the gap.");
def("pchip.check", "Check", "", "", "calculated", "OK, or what is wrong with this chip under the plate.");

// ---------- simulation setup ----------
def("sim.refDes", "RefDes", "", "", "input", "Chip that gets its own model in the simulation.");
def("sim.model", "Chip model", "", "", "input", "Block: one block whose k reproduces Rjc. Two-resistor: Rjc and Rjb (later).");
def("sim.pd", "Pd", "Pd", "W", "calculated", "Heat applied to the chip body in the study.");
def("sim.rjc", "Rjc", "Rjc", "C/W", "calculated", "Rjc of the part, from tab 1.");
def("sim.h", "Body height", "H", "mm", "cad", "Height of the chip body.");
def("sim.area", "Body area", "A", "mm²", "cad", "L × W of the chip body.");
def("sim.kz", "Chip k through", "k z", "W/(m K)", "calculated", "Conductivity given to the chip block, top to bottom, so that it has the datasheet Rjc.",
  "A real chip is die, mould compound and leadframe. The simulation uses one block; k z = H / (Rjc × A) gives the block the same resistance from bottom to top as the datasheet Rjc.");
def("sim.kxySet", "k xy typed", "", "W/(m K)", "input", "Optional in-plane conductivity of the chip block.");
def("sim.kxy", "Chip k in plane", "k xy", "W/(m K)", "calculated", "The k xy typed on tab 8, else the same as k z.");
def("sim.check", "Check", "", "", "calculated", "OK, or why the chip model cannot be set up.");

// ---------- results ----------
def("run.name", "Run", "", "", "run", "Name of the simulation run.");
def("run.date", "Run date", "", "", "run", "When the run was made.");
def("run.studyType", "Study", "", "", "run", "Thermal (SolidWorks Simulation, conduction with h on the surfaces) or Flow (Flow Simulation, air modelled).");
def("run.status", "Run status", "", "", "run", "Where the run is: set up, solved, read.");
def("run.energyBalance", "Energy balance error", "", "%", "run", "Heat in minus heat out, as a share. Above 1 % the mesh or the set up is suspect.");
def("run.pcbMin", "PCB coolest", "", "C", "run", "Lowest temperature on the PCB.");
def("run.pcbMax", "PCB hottest", "", "C", "run", "Highest temperature on the PCB.");
def("run.spreaderMin", "Spreader coolest", "", "C", "run", "Lowest temperature on the spreader plate.");
def("run.spreaderMax", "Spreader hottest", "", "C", "run", "Highest temperature on the spreader plate.");
def("runp.peak", "Peak", "", "C", "run", "Highest temperature in the part body. For a block chip model this is close to Tj.");
def("runp.avg", "Average", "", "C", "run", "Average temperature of the part body.");

// ---------- changes ----------
def("change.checked", "Checked", "", "", "input", "Tick when you have looked at this change and agree with it. Building on unticked changes is refused.");
def("change.date", "Date", "", "", "cad", "When the change was found.");
def("change.macro", "Macro", "", "", "cad", "The macro that found it.");
def("change.body", "Body", "", "", "cad", "The body that changed.");
def("change.change", "Change", "", "", "cad", "What happened: added, moved, resized, gone, override from CAD.");
def("change.field", "Field", "", "", "cad", "The value that changed.");
def("change.old", "Old", "", "", "cad", "Value before.");
def("change.new", "New", "", "", "cad", "Value after.");

// ---------- files ----------
def("file.folder", "Folder", "", "", "input", "The project folder. One folder per board; the page saves project.json there.");
def("file.project", "project.json", "", "", "input", "Everything you decide. Only this page writes it.");
def("file.cad", "cad.json", "", "", "cad", "What SolidWorks knows: the import, the build, the changes. Only the macros write it.");
def("file.results", "results.json", "", "", "run", "The simulation runs. Only the macros write it.");

export const TERMS = T;

// ---------- glossary: abbreviations and symbols in any text ----------
export const GLOSSARY = {
  "Pd": "Power dissipation: electrical power turned into heat in a part (W).",
  "Tj max": "Highest junction temperature the datasheet allows (C).",
  "Tj est": "Junction temperature estimate from Rja (C).",
  "Tj": "Junction temperature: the temperature of the silicon inside the part (C).",
  "Ta": "Ambient: the air temperature around the part (C).",
  "Rja": "Junction to ambient thermal resistance, on the JEDEC test board (C/W). Also θJA.",
  "Rjc": "Junction to case (top) thermal resistance (C/W). Also θJC.",
  "Rjb": "Junction to board thermal resistance (C/W). Also θJB.",
  "TIM": "Thermal interface material: the pad, paste or gel between a chip and a heatsink or plate.",
  "RefDes": "Reference designator: the part's name on the schematic and silkscreen (U1, L3).",
  "PCB": "Printed circuit board.",
  "PMIC": "Power management IC: several regulators in one package.",
  "LDO": "Low dropout linear regulator. Its loss is (Vin - Vout) × Iout.",
  "STEP": "A 3D CAD exchange file format (.step, .stp).",
  "DCR": "DC resistance of an inductor winding.",
  "JEDEC": "The standards body whose test board datasheet Rja values are measured on.",
  "W/cm²": "Watts per square centimetre: power density. 100 mm² = 1 cm².",
  "Bi": "Biot number: surface heat loss compared with conduction inside the plate.",
  "k xy": "Thermal conductivity in plane (W/(m K)).",
  "k z": "Thermal conductivity through the thickness (W/(m K)).",
  "η": "Efficiency of a regulator: Pout / Pin.",
  "ε": "Emissivity (0 to 1), or in the spreading model the area ratio.",
  "τ": "Relative thickness in the spreading model.",
  "λ": "A term of the spreading model.",
  "Φ": "A term of the spreading model.",
  "σ": "Stefan-Boltzmann constant, 5.67 × 10⁻⁸ W/(m² K⁴).",
};

// ---------- status words ----------
export const STATUS = {
  "CRITICAL": "A flag limit for CRITICAL is crossed: low margin, high power density or a big share of the board's heat. Look at these first.",
  "Watch": "A Watch limit is crossed, no CRITICAL one. Check these after the CRITICAL parts.",
  "Minor": "Below every Watch limit on this first estimate.",
  "No Pd": "Pd cannot be worked out yet: a value the method needs is missing.",
  "FAIL": "The hand check Tj is above Tj max.",
  "Tight": "Tj is below Tj max, by less than the CRITICAL margin limit.",
  "OK": "Inside the limits.",
  "locked": "This step cannot run yet. The tab says what to do first.",
  "to do": "Open and not done yet.",
  "done": "Nothing left to do here for now.",
  "needs attention": "Something on this tab needs fixing or a look.",
};
