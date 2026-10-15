' Patches the open test workbook with the M8 inputs (same as SetupWorkbook now does for new copies).
Set xl = GetObject(, "Excel.Application")
Set wb = Nothing
For Each w In xl.Workbooks
  If InStr(w.Name, "_test.xlsx") > 0 Then Set wb = w
Next
If wb Is Nothing Then WScript.Echo "test workbook not open": WScript.Quit 1
Set ws = wb.Worksheets("Sim setup")
th = ChrW(952)
ws.Range("B12").Value = "Two-resistor: outer layer thickness"
ws.Range("C12").Value = 0.1
ws.Range("D12").Value = "fraction of the chip height, for each of the top and bottom layers"
ws.Range("B13").Value = "Two-resistor: k of the junction core"
ws.Range("C13").Value = 100
ws.Range("D13").Value = "W/m K, high so the core is one temperature"
ws.Range("B11").Copy: ws.Range("B12:B13").PasteSpecial -4122
ws.Range("C11").Copy: ws.Range("C12:C13").PasteSpecial -4122
ws.Range("D11").Copy: ws.Range("D12:D13").PasteSpecial -4122
xl.CutCopyMode = False
wb.Names.Add "Sim_LayerFrac", "='Sim setup'!$C$12"
wb.Names.Add "Sim_kCore", "='Sim setup'!$C$13"
f = "=IF($B29="""","""",IF($L29="""",""RefDes not on CAD geometry"",IF(N($J29)<=0,""no heat in the register"",IF($C29=""Two-resistor"",IF(AND(N($D29)>0,N($E29)>0),""OK (two-resistor)"",""needs " & th & "jc and " & th & "jb: block model is used""),IF($H29="""",""no " & th & "jc: k of parts without heat is used"",""OK"")))))"
ws.Range("K29:K88").Formula = f
If WScript.Arguments.Count > 0 Then
  ' test chip: U1 (row 30) as two-resistor
  ws.Range("C30").Value = "Two-resistor"
  ws.Range("E30").Value = WScript.Arguments(0)
End If
xl.Calculate
WScript.Echo "B30=" & ws.Range("B30").Value & " C30=" & ws.Range("C30").Value & " D30=" & ws.Range("D30").Value & " E30=" & ws.Range("E30").Value & " F30=" & ws.Range("F30").Value & " G30=" & ws.Range("G30").Value & " J30=" & ws.Range("J30").Value & " K30=" & ws.Range("K30").Value
WScript.Echo "names: " & wb.Names("Sim_LayerFrac").RefersTo & " " & wb.Names("Sim_kCore").RefersTo
