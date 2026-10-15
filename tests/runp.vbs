Set sw = GetObject(, "SldWorks.Application")
ok = sw.RunMacro("C:\Users\Sultan\Desktop\thermal_similation\tests\swp\PTTest.swp", WScript.Arguments(0), WScript.Arguments(1))
WScript.Echo "ok=" & ok
