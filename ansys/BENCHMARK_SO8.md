# Benchmark: SO-8 on five test boards (measured, 1999)

A tested reference to check our setup against: first in Icepak, later in SolidWorks Flow Simulation. One SO-8
thermal test package on five different single-component boards, still air, measured junction temperature.
Same package and power on every board, so the five results check how well we model the board.

**Source:** J. Lohan, P. Tiilikka, P. Rodgers, C.-M. Fager, J. Rantala, *Effect of PCB Thermal Conductivity on the
Operating Temperature of an SO-8 Package in a Natural Convection Environment: Experimental Measurement versus
Numerical Prediction*, THERMINIC 1999, Rome, 4-6 Oct 1999. Nokia Research Center and University of Limerick.
Full text (7 pages): https://forums.futura-sciences.com/attachments/electronique/173759d1329913598-radiateurs-publication_17.pdf

Every value below says where it comes from:
- **paper**: printed in the text or a table (exact).
- **figure**: printed as a dimension on a drawing in the paper (exact, but our reading of what it dimensions is ours).
- **graph**: read off a plot by eye (about ±2 °C or ±3 °C/W).
- **ours**: not in the paper; a standard or an assumption we make. Check these first if results disagree.

---

## 1. What is measured (the truth)

| Quantity | How | Accuracy | Source |
|---|---|---|---|
| Junction temperature Tj | Diode on an Infineon G423 thermal test die (0.63 x 0.63 mm), calibrated to JESD51-1 | ±1 °C | paper |
| Repeatability | Two samples per board, average reported | pairs within 1.8 °C, typically 1 °C | paper |
| Board and package surface temperature | IR camera (Agema 550, 3.6-5 µm), all surfaces painted, emissivity 0.92 | ±0.5 °C, spot about 0.45 mm, averages over about 1.5 mm | paper |
| Package junction-to-case, cold plate (SEMI G30-88) | Cold plate at 20 °C, paste k = 3.0 W/mK | ±2.3 °C/W | paper |

The paint did not change temperatures (within 1 °C). Use emissivity 0.92 on all surfaces in the model.

### 1.1 Main result: Tj at 0.5 W (paper, Table 5)

| Board | Tj measured (°C) | θja = (Tj - Ta)/P, graph (°C/W) |
|---|---|---|
| FR4 "1" (JEDEC 1S) | **124.3** | about 210 |
| FR4 "2" (JEDEC 2S2P) | **84.9** | about 130 |
| FR4 "3" (4S4P) | **83.3** | about 122 |
| IMS Cu | **71.5** | about 101 |
| IMS Al | **70.6** | about 99 |

### 1.2 Other power levels (graph, Figure 6, about ±2 °C)

| Board | 0.2 W | 0.3 W | 0.5 W | 0.75 W |
|---|---|---|---|---|
| FR4 "1" | 66 | 87 | 128 | |
| FR4 "2" | 50 | 63 | 89 | |
| FR4 "3" | 49 | 61 | 86 | |
| IMS Al | | 54 | 75 | 100 |
| IMS Cu | | 55 | 75.5 | 102 |

The 0.5 W points in Figure 6 sit 3 to 4 °C above Table 5 (for example FR4 "1" 128 vs 124.3). The paper does not
say why. **Use Table 5 as the truth at 0.5 W**; use Figure 6 only for the trend with power.
θja does not depend on power (Figure 7), so the model should give a straight line Tj = Ta + θja x P.

### 1.3 Board surface profile at 0.5 W (graph, Figure 9)

IR line along Z through the package centre, board edge at 0, package at about 36-40 mm.
FR4 "1": peak about 102 °C on the package, about 50 °C at 10 mm from the package, about 23 °C at both board edges.
The other boards peak at about 55-67 °C and fall to about 25 °C within 15 mm. Figure 10 (FR4 "1") and Figure 11
(FR4 "2", X and Z lines) give the same profiles with the paper's own CFD on top. Compare shapes only; the peak on
the package reads low because the camera averages over 1.5 mm.

### 1.4 Package alone on a cold plate (paper)

A check of the package model before any board. Cold plate 20 °C, three powers giving Tj - Tplate = 25, 50, 75 °C.

| Orientation | θjc measured (°C/W) |
|---|---|
| Dead-bug (package top on the plate) | 82.3 ± 2.3 |
| Live-bug (leads and base on the plate) | 51.4 |

### 1.5 Ambient temperature (not printed)

The paper gives absolute Tj but not the ambient. From Tj (Table 5) minus θja x 0.5 W (Figure 7) every board gives
**Ta = 19 to 22 °C** (FR4 "1": 124.3 - 105 = 19; FR4 "2": 84.9 - 65 = 20; FR4 "3": 83.3 - 61 = 22; IMS: about 21).
**Use 21 °C** and judge results as temperature rise Tj - Ta as well as absolute Tj (graph-derived, ours).

---

## 2. Test environment

| Item | Value | Source |
|---|---|---|
| Environment | JEDEC natural convection still air (JESD51-2) | paper |
| Enclosure | 1 cubic foot (305 mm cube) closed box, board centred | ours (JESD51-2) |
| Board orientation | Horizontal, component on top, gravity normal to the board | paper (Figure 4c: gravity along Y, board in XZ) |
| Flow | Laminar | paper |
| Paper's own CFD domain | Quarter model 70 x 80 x 50 mm (X, Y up, Z), two symmetry planes, free-air (opening) boundaries elsewhere | paper, Table 3 |

Run both: the JEDEC box (closest to the test) and the paper's open domain (closest to their CFD). The difference is
part of what we learn.

---

## 3. Test boards

All boards: **76.2 x 114 mm**, JEDEC single-component layout, package near the centre of the 76.2 x 76 mm test area,
edge connector at one short end (paper, Figure 3). FR4 boards 1.6 mm thick (paper, Table 1).

### 3.1 Surface trace layer (JESD51-3 / 51-7 layout)

Copper traces **70 µm thick, 250 µm wide**, one per lead, running from the pads to the edge connector (paper).
Exact trace paths and lengths are in JESD51-3 Figure 2 and not dimensioned in the paper (ours: draw them to the
standard, or use the paper's compact model below).

On which side: FR4 "1", IMS Al and IMS Cu on the component side only; FR4 "2" and FR4 "3" on both sides (paper).

**Paper's compact trace model** (Strategy 3, Figure 4c): the surface trace layer as three flat blocks, 70 µm thick,
in steps away from the package, with isotropic k = **96.3, 77.0, 55.0 W/mK** (blocks 1, 2, 3, block 1 nearest the
leads). Block extents are not dimensioned (figure only).

### 3.2 Stack-ups (paper, Table 1; spacings from the drawings)

| Board | Build | Layers, top to bottom |
|---|---|---|
| **FR4 "1"** | JESD51-3 low-K, 1S | 70 µm Cu traces / 1.6 mm FR4 |
| **FR4 "2"** | JESD51-5/51-7 high-K, 2S2P | 70 µm Cu traces (20 %) / 0.36 FR4 / 35 µm Cu plane (100 %) / 0.71 FR4 / 35 µm Cu plane (100 %) / 0.36 FR4 / 70 µm Cu traces (20 %). Sum 1.64 mm |
| **FR4 "3"** | Application style, 4S4P | 70 µm Cu traces (20 %) on both surfaces, 4 internal 35 µm Cu planes (100 %), 2 internal 35 µm Cu signal layers (20 %, the 74 x 74 mm hatched area of Figure 3b). Layer spacing "about 0.20 mm". Layer order: figure only, not readable with certainty |
| **IMS Al** | 1S1P insulated metal substrate | 70 µm Cu signal / 150 µm dielectric / 70 µm Cu plane / aluminium baseplate. Drawing gives 1.6 and 2.0 mm; read as 1.6 mm baseplate in a 2.0 mm board, which leaves about 0.1 mm unexplained |
| **IMS Cu** | Same build as IMS Al, copper baseplate | as IMS Al |

Internal planes cover the hatched area of Figure 3b (the 76 x 76 mm test area), 100 % copper there (paper).
Order of plane layer and dielectric in the IMS boards is drawn as signal / plane / dielectric / baseplate in the
labels, which is unusual; we read it as signal / dielectric / plane / baseplate (ours, check).

### 3.3 Effective board conductivities the paper used (W/mK)

| Board | In-plane | Through-plane | Source |
|---|---|---|---|
| FR4 "1" | 0.81 | 0.29 | paper (accepted values) |
| FR4 "2" | 24.36 / 17.7 / **18.6** | 0.303 / 0.32 / **0.35** | paper, Table 4 (series-parallel / Azar-Graebner / **Punch-Davies**) |
| FR4 "3" | 41.16 / 34.6 / **36.5** | 0.318 / 0.34 / **0.38** | paper, Table 4 |
| IMS Al | 164 | | paper (parallel resistance) |
| IMS Cu | 328 | | paper |

Bold is the Punch-Davies value the paper plots against (Figure 8). Isotropic alternatives in Table 4: arithmetic
mean 12.33 and 20.74, geometric mean 2.72 and 3.62 (FR4 "2" and "3").

---

## 4. Package: SO-8 (Infineon P-DSO-8-1)

Dimensions in mm from Figure 2 (figure). Axes as the paper: X along the package length (lead rows parallel to X),
Z across, Y up. Values marked "half" are from the centre line.

| Part | Size | Position |
|---|---|---|
| Body (encapsulant) | 5.0 (X, 2.5 half) x 4.0 (Z) x 1.45 (Y) | sits on the board, standoff not given (ours: 0.1 mm, JEDEC MS-012) |
| Paddle (copper) | 4.3 (X, 2.15 half) x 2.23 (Z) x 0.2 thick | bottom of paddle 0.3 above body bottom |
| Die attach | 0.63 x 0.63, 0.01 thick (vendor); measured 3-5 µm | on the paddle |
| Die (silicon) | 0.63 x 0.63 x 0.4 | on the die attach, centred |
| Inner leadframe | two strips, 0.685 wide (Z) x 0.2 thick, running along X | one at each Z side, 0.2 inside the body edge, bottom 0.55 above body bottom; gap to the paddle 0.5 (Z) |
| Leads (outside the body) | not dimensioned | ours: 8 gull-wing leads, pitch 1.27, width 0.41, span 6.0 over the leads (JEDEC MS-012) |

Material conductivity (paper, Table 2, vendor values):

| Element | k (W/mK) |
|---|---|
| Die (silicon) | 117.5 - 0.42 (T - 100), T in °C |
| Die attach | 1.25 |
| Copper paddle | 260 |
| Encapsulant | 0.62 |
| Leadframe | 260 |
| Inner leadframe block (volume averaged) | 208 |
| Outer leadframe block, package leads (volume averaged) | 96.0 |

Die attach thickness is the most sensitive input (paper): run **5 µm and 10 µm** and report both, as the paper does.

Heat path at 0.5 W from the paper's detailed model (Table 6, % of power): leads 72-80, package base 14-17,
sides 3.5-8, package top 2.3-5.4. Our model should show most of the heat leaving through the leads.

---

## 5. What the paper's own CFD got (Flotherm 2.1, 0.5 W)

Error = predicted - measured Tj in °C, die attach 5 µm (10 µm in brackets). Paper, Table 5.

| Board | Strategy 1: all layers in detail | Strategy 2: one anisotropic block | Strategy 3: anisotropic block + trace blocks |
|---|---|---|---|
| FR4 "1" | -3.1 (+0.6) | | |
| FR4 "2" | +6.8 (+11.6) | +28.7 (+32.6) | +3.2 (+7.0) |
| FR4 "3" | +3.7 (+7.6) | +24.9 (+28.7) | +1.5 (+5.4) |
| IMS Cu | +5.0 (+8.9) | +2.2 (+6.2) | |
| IMS Al | +6.5 (+10.4) | +3.9 (+7.8) | |

Cold plate model: dead-bug 82.0 °C/W at 10 µm (measured 82.3); live-bug +5.6 °C/W over measured at 10 µm; at 5 µm
both underpredict, by 7.3 and 3.4 °C/W.

---

## 6. How we run it

Order: the package first, then one board, then all five.

| Step | Model | Pass when |
|---|---|---|
| B0 | Package alone on the cold plate, dead-bug and live-bug | θjc within ±10 % of 82.3 and 51.4 °C/W |
| B1 | FR4 "1", 0.5 W, Ta 21 °C, JEDEC box | Tj within ±5 °C of 124.3; surface profile shape matches Figure 9 |
| B2 | FR4 "2", FR4 "3", IMS Al, IMS Cu, 0.5 W | each Tj within ±5 °C of Table 5; board order kept (FR4 "1" > "2" > "3" > IMS Cu > IMS Al) |
| B3 | All five at 0.2, 0.3, 0.75 W | Tj - Ta linear in power; within ±5 °C of Figure 6 |
| B4 | Board as detailed layers vs anisotropic block vs block + traces (paper's three strategies) | same ranking of errors as section 5 |
| B5 | Same B0-B2 in SolidWorks Flow Simulation | same pass limits; Icepak and Flow within 3 °C of each other |

±5 °C is about what the paper's own best models achieved (1.5 to 6.8 °C), plus the ±1 °C diode accuracy and the
ambient uncertainty.

Every run records: board, power, die attach thickness, board model (strategy 1/2/3), domain (JEDEC box or open),
mesh count, Tj, θja, error vs measured.

### Not covered by our Altium route

The package is built by hand (blocks per section 4), not from an EDB. The boards can be built either as Icepak PCB
objects with the stack-up in section 3.2, or as blocks with the effective k in section 3.3. Our two-resistor
junction model (Rjc, Rjb) is not usable here because the paper gives θjc for the cold plate only, not a JEDEC
two-resistor pair; it could be fitted from B0 as an extra study.

---

## 7. Open points

- Ambient temperature is derived (section 1.5), not stated.
- Table 5 vs Figure 6 differ by 3-4 °C at 0.5 W; reason not given.
- Surface trace paths: draw to JESD51-3 or use the compact trace blocks; the block extents are not dimensioned.
- FR4 "3" layer order and IMS total thickness are only readable approximately from the drawings.
- Lead geometry and standoff are from the JEDEC outline, not the paper.
- The paper's companion work (Rodgers et al., SO16 / TSOP48 / PQFP208 on single- and multi-component boards, still
  air and 1-2.25 m/s) would be the next benchmark; full text not found free yet.
- Dead-bug θjc comes out about 20 °C/W above the paper's own model with the paper's values (section 8).

---

## 8. Results

### B0: package on the cold plate (2026-10-28, Icepak 2026 R1)

Script: `benchmark/so8_coldplate.py`, results in `benchmark/runs/b0_*.json`. Conduction only (flow off, the paper has
insulation all round), plate = the region's bottom face and the paste's bottom face at 20 °C, paste 25 µm (ours),
standoff 0.1 mm (ours), silicon k 130 (Table 2 value near 70 °C), outer leads as one gull-wing block per row with
k 96. θjc from the die maximum; the die mean is about 1 °C/W lower. Dead-bug at 0.6 W, live-bug at 1.0 W (about 50 °C
rise). Mesh: a package mesh region at level 3, 4 and 5 moves θjc by at most 1.3 °C/W, so level 4 is used.

| Case (mesh level 4) | θjc ours (°C/W) | Measured | Error | Paper's model |
|---|---|---|---|---|
| Live-bug, attach 10 µm | 57.3 | 51.4 | +11.5 % | 57.0 |
| Live-bug, attach 5 µm | 48.9 | 51.4 | **-4.8 %, pass** | 48.0 |
| Live-bug, air in the standoff gap, 10 / 5 µm | 97.1 / 89.0 | 51.4 | +89 / +73 % | |
| Dead-bug, attach 10 µm | 103.5 | 82.3 | +25.8 % | 82.0 |
| Dead-bug, attach 5 µm | 95.6 | 82.3 | **+16.2 %, fail** | 75.0 |

- **Live-bug agrees with the paper's own model within 1 °C/W** and passes at the measured attach thickness (3-5 µm).
- The body base must sit in the paste: with air in the 0.1 mm standoff gap θjc nearly doubles. So "base on the
  plate" in the paper means paste under the whole body.
- **Dead-bug is about 20 °C/W above both the measurement and the paper's model.** Here all the heat crosses the mould
  into the plate; a hand estimate with k 0.62 through the 0.95 mm above the paddle also gives about 100 °C/W. The
  paper's model must have had an easier path than its printed dimensions and k give. Not found which.

Mould conductivity check (the paper says θjc is very sensitive to it), attach 10 µm:

| Mould k (W/mK) | Dead-bug θjc | Error | Live-bug θjc | Error |
|---|---|---|---|---|
| 0.62 (paper) | 103.5 | +25.8 % | 57.3 | +11.5 % |
| 0.8 | 84.9 | +3.1 % | 50.2 | -2.4 % |
| 0.9 | 77.9 | -5.4 % | | |

With k 0.8 both orientations pass at the vendor's 10 µm attach. That is a fit, not a paper value, so it is noted here
and not used: B1 onward run with the paper's 0.62 and both attach thicknesses, and k 0.8 as a sensitivity case.

**B0 verdict:** partial pass. Live-bug passes, dead-bug fails with the paper's own values. Live-bug is the
orientation that matters for B1-B2 (heat leaves mainly through the leads into the board, as on the test boards).
