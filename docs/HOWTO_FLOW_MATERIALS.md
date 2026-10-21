# Flow materials and radiative surfaces (Flow plan F4, F5)

Run Flow Simulation gives every chip, the PCB, the TIMs and the spreader the k values of the page, and every surface
group the emissivity of the page. The Flow API cannot make a material or a radiative surface; it can only pick one
from the Engineering Database by name. So the page names them after their value and you make each one once.
After that every project that needs the same value finds it.

## What to make

Tab 8 (study type Flow), card **Flow materials**, lists them:

- **Materials** (Solids): name `k <k xy> <k z>`, or `k <k>` when both are the same. k is rounded to 10 % steps
  (24 per decade), so a chip block's Rjc stays within 5 % of the datasheet.
- **Radiative surfaces**: name `e <emissivity>` with 2 decimals.

## How (once per name)

1. In SolidWorks: Tools > Flow Simulation > Tools > Engineering Database.
2. Materials > Solids > User Defined. Make a folder `PCBThermal` the first time (right click > New Folder).
3. In it, New Item for each material on tab 8:
   - Name: exactly as listed, for example `k 19.6` or `k 21.5 0.316`.
   - Density 2000 kg/m3, specific heat 1000 J/(kg K) for every one (a steady state result does not depend on them).
   - Conductivity type: Isotropic when there is one k; Axisymmetrical/Biaxial (axial, axis Z) with
     k axial = k z and k radial = k xy when there are two.
   - Save.
4. Radiative Surfaces > User Defined > folder `PCBThermal` > New Item for each `e ...` name: type Real surface,
   emissivity as listed. Save.

## What Run Flow does with them

- Path used: `User Defined\PCBThermal\<name>`. A missing one is noted in the report and on tab 9 notes; the macro
  then uses Flow's generic pick (Typical Plastic Package, PCB 4/8/12 layers, GR25A, aluminium or copper) and a chip
  that falls back is marked `package` on tab 9.
- An anisotropic material made new by the macro needs its axis set to Z once in SolidWorks (the report says so).
- Contacts are set directly (integral value in K/W over the contact area); they need nothing in the database.

## Simplifications in Flow (tab 8 lists the ones that apply)

- One emissivity per component group: PCB (top value), parts, spreader (top value).
- Chip to TIM contact is folded into the TIM material: k = t / (t / k TIM + R contact).
- Part to PCB contact sits on the PCB component, TIM to spreader contact on the spreader component.
