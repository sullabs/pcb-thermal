Set sw = GetObject(, "SldWorks.Application")
Set md = sw.OpenDoc("C:\Users\Sultan\Desktop\thermal_similation\example\RaspberryPi4\CAD\RaspberryPi4_Thermal.sldasm", CLng(2))
WScript.Echo TypeName(md)
