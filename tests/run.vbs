Set sw = GetObject(, "SldWorks.Application")
ok = sw.RunMacro("C:\Users\Sultan\Desktop\thermal_similation\tests\swp\PTTest.swp", WScript.Arguments(0), "main")
WScript.Echo WScript.Arguments(0) & " ok=" & ok
