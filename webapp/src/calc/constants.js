// Fixed lists of the template (tab 17 Q methods, tab 12 Settings). Not board data.

export const METHODS = ["LDO / linear reg", "Switching regulator", "Digital load (V x I)", "Resistive (I squared R)", "Manual entry"];

// Class -> default power method (17 Q methods, class to method map).
export const CLASS_METHOD = {
  "Processor / SoC": "Digital load (V x I)",
  "Memory": "Digital load (V x I)",
  "LDO / linear reg": "LDO / linear reg",
  "Switching regulator": "Switching regulator",
  "PMIC": "Switching regulator",
  "Ethernet PHY": "Digital load (V x I)",
  "USB / interface IC": "Digital load (V x I)",
  "Power inductor": "Resistive (I squared R)",
  "Sense resistor": "Resistive (I squared R)",
  "MOSFET / driver": "Resistive (I squared R)",
  "Connector": "Resistive (I squared R)",
  "Other": "Manual entry",
  "RF / wireless module": "Digital load (V x I)",
};
export const CLASSES = Object.keys(CLASS_METHOD);

export const STEFAN_BOLTZMANN = 0.0000000567;   // W/(m2 K4), as the template uses it
