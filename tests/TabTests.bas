Attribute VB_Name = "TabTests"
Option Explicit
' Test: PT_TabBase and PT_Sheet against many tab-name styles. Writes tests\tabtests_out.txt

Dim outTxt As String
Dim nFail As Long

Sub Chk(ByVal label As String, ByVal got As String, ByVal want As String)
    If got = want Then
        outTxt = outTxt & "ok   " & label & vbCrLf
    Else
        outTxt = outTxt & "FAIL " & label & ": got [" & got & "] want [" & want & "]" & vbCrLf
        nFail = nFail + 1
    End If
End Sub

Sub main()
    Dim xl As Object, wb As Object, names As Variant, i As Long, ws As Object, f As Long, r As Variant
    outTxt = "": nFail = 0
    ' PT_TabBase
    Chk "base 7 Spreader", PT_TabBase("7 Spreader"), "Spreader"
    Chk "base 07_Spreader", PT_TabBase("07_Spreader"), "Spreader"
    Chk "base 7. Spreader", PT_TabBase("7. Spreader"), "Spreader"
    Chk "base 7-Spreader", PT_TabBase("7-Spreader"), "Spreader"
    Chk "base 7)Spreader", PT_TabBase("7)Spreader"), "Spreader"
    Chk "base nbsp", PT_TabBase("7" & Chr(160) & "Spreader"), "Spreader"
    Chk "base double space", PT_TabBase("7  Spreader  "), "Spreader"
    Chk "base 14 Read me", PT_TabBase("14 Read me"), "Read me"
    Chk "base no number", PT_TabBase("Spreader"), "Spreader"
    Chk "base 7b Spreader bottom", PT_TabBase("7b Spreader bottom"), "Spreader bottom"
    Chk "base 8b Thermal solution bottom", PT_TabBase("8b Thermal solution bottom"), "Thermal solution bottom"
    Chk "base 1 Component register", PT_TabBase("1 Component register"), "Component register"
    Chk "base Bare board", PT_TabBase("6 Bare board network"), "Bare board network"
    Chk "base starts with letter word", PT_TabBase("7 b"), "b"
    ' PT_Sheet on a workbook with the real tab style
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False: xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    names = Array("1 Component register", "7 Spreader", "7b Spreader bottom", "8 Thermal solution", "8b Thermal solution bottom", "9 Sim setup", "10 Sim results", "CAD geometry")
    For i = 0 To UBound(names)
        If i = 0 Then
            wb.Worksheets(1).Name = names(i)
        Else
            Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
            ws.Name = names(i)
        End If
    Next
    Chk "sheet Spreader", NameOf(PT_Sheet(wb, "Spreader")), "7 Spreader"
    Chk "sheet Spreader bottom", NameOf(PT_Sheet(wb, "Spreader bottom")), "7b Spreader bottom"
    Chk "sheet Thermal solution", NameOf(PT_Sheet(wb, "Thermal solution")), "8 Thermal solution"
    Chk "sheet Thermal solution bottom", NameOf(PT_Sheet(wb, "Thermal solution bottom")), "8b Thermal solution bottom"
    Chk "sheet Sim setup", NameOf(PT_Sheet(wb, "Sim setup")), "9 Sim setup"
    Chk "sheet Sim results", NameOf(PT_Sheet(wb, "Sim results")), "10 Sim results"
    Chk "sheet CAD geometry (no number)", NameOf(PT_Sheet(wb, "CAD geometry")), "CAD geometry"
    Chk "sheet exact full name", NameOf(PT_Sheet(wb, "7 Spreader")), "7 Spreader"
    Chk "sheet case insensitive", NameOf(PT_Sheet(wb, "sim SETUP")), "9 Sim setup"
    Chk "sheet missing", NameOf(PT_Sheet(wb, "Nothing here")), "(none)"
    Chk "sheet Component register", NameOf(PT_Sheet(wb, "Component register")), "1 Component register"
    wb.Close False
    xl.Quit
    outTxt = outTxt & vbCrLf & "failures=" & nFail & vbCrLf
    f = FreeFile
    Open "C:\Users\Sultan\Desktop\thermal_similation\tests\tabtests_out.txt" For Output As #f
    Print #f, outTxt
    Close #f
End Sub

Function NameOf(ByVal ws As Object) As String
    If ws Is Nothing Then NameOf = "(none)" Else NameOf = ws.Name
End Function
