// Learn mode texts (design 5.3): for each tab, what the step is for and what to look at. Shown at the top of the
// tab while Learn mode is on. Abbreviations are marked by the page, so write Pd, Rja, TIM as they are.

export const LEARN = {
  project: {
    purpose: "Set up the study: one folder per board, the board name, and the air temperature the board works in.",
    look: [
      "Ambient is the air inside the enclosure at the worst case, not the room. Every temperature on the later tabs starts from it, so a wrong ambient moves every result by the same amount.",
      "Write the load case and the enclosure in Notes. A thermal result is only true for the case it was worked out for.",
      "Where you are lists the steps and the next one to do. A grey step says what it waits for.",
    ],
  },
  parts: {
    purpose: "List every part that makes heat, with its datasheet values and its power loss Pd. This is the input everything else depends on.",
    look: [
      "Start with the parts that draw the most power: processor, memory, regulators, PHYs, power inductors. Small logic and passives can wait.",
      "Pd is the power turned into heat, not the power the part delivers. A buck that delivers 3 W at 85 % efficiency loses about 0.53 W.",
      "A regulator's loss belongs to the regulator. Its load (the processor) has its own Pd, worked out from its own supply current. Do not count the same watts twice.",
      "Tj est uses Rja from the JEDEC test board. It ranks the parts; it is not the final answer. The flag says which parts to look at first.",
      "Have the list in Excel already? Paste a parts list does it in one go.",
    ],
  },
  import: {
    purpose: "Bring the board's 3D shape in from the STEP file, so each part has a real footprint and height.",
    look: [
      "Check the PCB size and thickness against the drawing. A wrong board size changes every board temperature.",
      "Check the number of bodies. Far too many means tiny passives came in too: raise Smallest body. Too few: the export merged parts.",
      "The top view should look like the board. Bodies on the bottom side are drawn dashed.",
    ],
  },
  link: {
    purpose: "Tell the page which body belongs to which part, so each part's heat goes into the right place in the model.",
    look: [
      "Every part with Pd needs exactly one body. Suggest links matches body names with part numbers and RefDes; check each one.",
      "Include the bodies with heat and the big ones that block air or touch a heatsink. Skip labels, tiny passives and parts far from the heat.",
      "If a body is a placeholder or has the wrong size (a connector with its leads, a chip shown as a flat outline), type its real size.",
      "A part that is not in the STEP can be added by hand with its size and position.",
      "Set the PCB stack here: the copper in the board decides how far heat spreads sideways.",
    ],
  },
  triage: {
    purpose: "Rank the parts to see where to spend your time, and check that the power adds up.",
    look: [
      "By Pd shows which parts make the most heat. By power density shows which parts have the least area to lose it through. A small part with modest Pd can be the real problem.",
      "The rail budget sums the Pd fed from each rail. If you measured the board input power, the unaccounted share tells you whether a heat source is missing.",
      "Fix first lists what stops the hand check on tab 5.",
    ],
  },
  bare: {
    purpose: "Work out every part's junction temperature by hand on the bare board, before any simulation. Here you learn where the heat goes.",
    look: [
      "Each part has two paths: up through its case top into the air, and down into the board, which spreads it and loses it from both faces.",
      "All parts heat one shared board, so a hot neighbour raises everyone. The board temperature is the first number to check.",
      "Look at the share of heat that goes into the board. Above about half, more copper or a bigger board helps more than a heatsink on top.",
      "A FAIL or Tight part needs a change (copper, airflow, a spreader on tab 7) or a better number (a measured Pd, a real Rjc).",
      "The simulation on tab 8 should land within a few degrees of this. A big gap means an input is wrong in one of the two.",
    ],
  },
  build: {
    purpose: "Build the thermal model in SolidWorks: one simple box per included body, on a PCB plate.",
    look: ["Read the failures first. Each one names the body and the reason.", "Overlapping bodies share volume in the model: fix the sizes on tab 3."],
  },
  spreader: {
    purpose: "Optional: a metal plate over (or under) the board that takes heat from the hot chips through a TIM.",
    look: ["The TIM must be thin and fill the gap: a thick TIM can make things worse than no plate.", "Check every covered chip has an OK check."],
  },
  sim: {
    purpose: "Set up the SolidWorks thermal study: air and surfaces, contacts, mesh and how each chip is modelled.",
    look: ["Each chip is a block whose k makes its resistance top to bottom equal Rjc. Check the k z column for silly values.", "Use the same h and emissivity as the hand check, or the two cannot agree."],
  },
  results: {
    purpose: "Read the simulation results: temperature per part, against Tj max and against the hand check.",
    look: ["Check the energy balance first: above a few percent the result is not converged.", "A part above Tj max is marked."],
  },
  changes: {
    purpose: "Check every move and size change the macros found, so a wrong import or build is caught before you trust a result.",
    look: ["Tick a change once you know why it happened. Unticked changes lock the build and study steps."],
  },
};
