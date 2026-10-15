Attribute VB_Name = "SpreaderTests"
Option Explicit
' Test helpers that stand in for the user's hand work on the spreader part (run them from the same macro project as BuildSpreader).
' PlaceIdentity : put the first spreader component (no "bottom" in its name) at the assembly origin and fix it.
' PlaceZ        : put it 1.6 mm above the origin (as if the part had been made on the PCB top face).
' AddHole       : cut a 4 mm hole into the plate at board X 70, Y 10 (inside the plate), named UserHole.

Const FOLDER As String = "C:\Users\Sultan\Desktop\thermal_similation\example\RaspberryPi4\CAD\"
Const ASM_FILE As String = "RPi4_fresh_Thermal.sldasm"

Function FindComp(ByVal asm As Object, ByVal wantBottom As Boolean) As Object
    Dim comps As Variant, i As Long, nm As String
    comps = asm.GetComponents(True)
    For i = 0 To UBound(comps)
        nm = LCase(comps(i).Name2)
        If InStr(nm, "spreader") > 0 And (InStr(nm, "bottom") > 0) = wantBottom Then
            If comps(i).GetSuppression2 <> 0 Then Set FindComp = comps(i): Exit Function
        End If
    Next
End Function

Sub Place(ByVal zmm As Double, ByVal wantBottom As Boolean)
    Dim swApp As Object, asm As Object, c As Object, mu As SldWorks.MathUtility, xf As SldWorks.MathTransform
    Dim d(15) As Double, v As Variant, errs As Long
    Set swApp = Application.SldWorks
    Set asm = swApp.GetOpenDocumentByName(FOLDER & ASM_FILE)
    swApp.ActivateDoc3 asm.GetTitle, False, 0, errs
    Set c = FindComp(asm, wantBottom)
    If c Is Nothing Then MsgBox "no spreader component": Exit Sub
    d(0) = 1: d(4) = 1: d(8) = 1: d(11) = zmm / 1000#: d(12) = 1
    v = d
    Set mu = swApp.GetMathUtility
    Set xf = mu.CreateTransform(v)
    asm.ClearSelection2 True
    If c.IsFixed Then
        c.Select4 False, Nothing, False
        asm.UnfixComponent
        asm.ClearSelection2 True
    End If
    c.Transform2 = xf
    asm.EditRebuild3
    c.Select4 False, Nothing, False
    asm.FixComponent
    asm.ClearSelection2 True
    MsgBox "Spreader component placed at Z " & zmm & " mm: " & c.Name2
End Sub

Sub PlaceIdentity()
    Place 0, False
End Sub

Sub PlaceZ()
    Place 1.6, False
End Sub

Sub AddHole()
    Dim swApp As Object, asm As Object, c As Object, md As Object, pl As Object, f As Object, sk As Object, feat As Object
    Dim xf As Object, a As Variant, d As Variant, ox As Double, oy As Double, oz As Double, errs As Long
    Set swApp = Application.SldWorks
    Set asm = swApp.GetOpenDocumentByName(FOLDER & ASM_FILE)
    Set c = FindComp(asm, False)
    If c Is Nothing Then MsgBox "no spreader component": Exit Sub
    d = c.Transform2.ArrayData
    ox = d(9) * 1000#: oy = d(10) * 1000#: oz = d(11) * 1000#
    Set md = c.GetModelDoc2
    swApp.ActivateDoc3 md.GetTitle, False, 0, errs
    Set pl = PT_FindFeature(md, "PT_Under")
    md.ClearSelection2 True
    pl.Select2 False, 0
    md.SketchManager.InsertSketch True
    Set sk = md.SketchManager.ActiveSketch
    Set xf = sk.ModelToSketchTransform
    a = PT_ToSketch(xf, 70 - ox, 10 - oy, 0)
    md.SketchManager.CreateCircleByRadius a(0), a(1), 0, 0.002
    md.SketchManager.InsertSketch True
    md.ClearSelection2 True
    Set f = PT_LastFeatureOfType(md, "ProfileFeature")
    f.Name = "UserHole_S"
    Set feat = PT_CutThrough(md, f)
    If feat Is Nothing Then MsgBox "hole cut failed": Exit Sub
    feat.Name = "UserHole"
    md.Save3 1, errs, errs
    swApp.ActivateDoc3 asm.GetTitle, False, 0, errs
    MsgBox "Hole UserHole cut in " & md.GetTitle
End Sub

' CleanSpreaders: delete every spreader and TIM component from the assembly, save it, then delete their files.
Sub CleanSpreaders()
    Dim swApp As Object, asm As Object, comps As Variant, i As Long, nm As String, n As Long, f As String, errs As Long, warns As Long
    Set swApp = Application.SldWorks
    Set asm = swApp.GetOpenDocumentByName(FOLDER & ASM_FILE)
    swApp.ActivateDoc3 asm.GetTitle, False, 0, errs
    comps = asm.GetComponents(True)
    asm.ClearSelection2 True
    For i = 0 To UBound(comps)
        nm = LCase(comps(i).Name2)
        If InStr(nm, "spreader") > 0 Or InStr(nm, "_tim_") > 0 Then
            If comps(i).IsFixed Then
                comps(i).Select4 False, Nothing, False
                asm.UnfixComponent
                asm.ClearSelection2 True
            End If
            comps(i).Select4 True, Nothing, False
            n = n + 1
        End If
    Next
    If n > 0 Then asm.EditDelete
    asm.ClearSelection2 True
    asm.Save3 1, errs, warns
    Dim dm As Object, ttls As String, parts As Variant, ttl As String
    Set dm = swApp.GetFirstDocument
    Do While Not dm Is Nothing
        ttl = dm.GetTitle
        If InStr(LCase(ttl), "spreader") > 0 Or InStr(LCase(ttl), "_tim_") > 0 Then ttls = ttls & ttl & "|"
        Set dm = dm.GetNext
    Loop
    parts = Split(ttls, "|")
    For i = 0 To UBound(parts)
        If parts(i) <> "" Then swApp.CloseDoc parts(i)
    Next
    On Error Resume Next
    f = Dir(FOLDER & "RPi4_fresh_Spreader*.sldprt")
    Do While f <> ""
        Kill FOLDER & f
        f = Dir
    Loop
    f = Dir(FOLDER & "RPi4_fresh_TIM_*.sldprt")
    Do While f <> ""
        Kill FOLDER & f
        f = Dir
    Loop
    MsgBox n & " components removed from the assembly, spreader and TIM files deleted."
End Sub

Sub PlaceBottomIdentity()
    Place 0, True
End Sub

' CleanBottom: delete the bottom spreader stand-in and its stand-in TIMs (names with spreader_bottom or _tim_tst) from the assembly.
Sub CleanBottom()
    Dim swApp As Object, asm As Object, comps As Variant, i As Long, nm As String, n As Long, errs As Long, warns As Long
    Set swApp = Application.SldWorks
    Set asm = swApp.GetOpenDocumentByName(FOLDER & ASM_FILE)
    swApp.ActivateDoc3 asm.GetTitle, False, 0, errs
    comps = asm.GetComponents(True)
    asm.ClearSelection2 True
    For i = 0 To UBound(comps)
        nm = LCase(comps(i).Name2)
        If InStr(nm, "spreader_bottom") > 0 Or InStr(nm, "_tim_tst") > 0 Then
            If comps(i).IsFixed Then
                comps(i).Select4 False, Nothing, False
                asm.UnfixComponent
                asm.ClearSelection2 True
            End If
            comps(i).Select4 True, Nothing, False
            n = n + 1
        End If
    Next
    If n > 0 Then asm.EditDelete
    asm.ClearSelection2 True
    asm.Save3 1, errs, warns
    MsgBox n & " bottom stand-in components removed."
End Sub
