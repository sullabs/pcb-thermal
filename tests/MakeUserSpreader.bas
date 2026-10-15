Attribute VB_Name = "MakeUserSpreader"
Option Explicit
' Test helper: stands in for the user. Makes a part with ONE sketch (a chamfered outline) on its Front plane,
' saves it as <folder>\RPi4_fresh_Spreader.sldprt and inserts it, loose, into the thermal assembly.
' ZOFF = how far above the assembly origin the part is placed (mm); 0 = Front plane, 1.6 = on the PCB top face.

Const FOLDER As String = "C:\Users\Sultan\Desktop\thermal_similation\example\RaspberryPi4\CAD\"
Const ASM_FILE As String = "RPi4_fresh_Thermal.sldasm"
Const ZOFF As Double = 0

Sub main()
    Build False
End Sub

Sub mainBottom()
    Build True
End Sub

Sub Build(ByVal bottom As Boolean)
    Dim swApp As Object, md As Object, tmpl As String, pl As Object, f As Object, last As Object
    Dim pts As Variant, i As Long, n As Long, errs As Long, warns As Long, ok As Boolean
    Dim asm As Object, c As Object, p As String, ttl As String
    Set swApp = Application.SldWorks
    tmpl = swApp.GetUserPreferenceStringValue(8)
    Set md = swApp.NewDocument(tmpl, 0, 0, 0)
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then Set pl = f: Exit Do
        Set f = f.GetNextFeature
    Loop
    md.ClearSelection2 True
    pl.Select2 False, 0
    md.SketchManager.InsertSketch True
    ' outline in board mm: the upper right corner is cut off
    If bottom Then pts = Array(32, 2, 62, 2, 62, 26, 56, 30, 32, 30) Else pts = Array(2, 2, 83, 2, 83, 40, 70, 54, 2, 54)
    n = (UBound(pts) + 1) / 2
    For i = 0 To n - 1
        md.SketchManager.CreateLine pts(2 * i) / 1000#, pts(2 * i + 1) / 1000#, 0, _
                                    pts(2 * ((i + 1) Mod n)) / 1000#, pts(2 * ((i + 1) Mod n) + 1) / 1000#, 0
    Next
    md.SketchManager.InsertSketch True
    md.ClearSelection2 True
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "ProfileFeature" Then Set last = f
        Set f = f.GetNextFeature
    Loop
    If last Is Nothing Then MsgBox "no sketch made": Exit Sub
    last.Name = "Outline"
    p = FOLDER & "RPi4_fresh_Spreader" & IIf(bottom, "_Bottom2", "") & ".sldprt"
    ok = md.Extension.SaveAs(p, 0, 1, Nothing, errs, warns)
    ttl = md.GetTitle
    Set asm = swApp.GetOpenDocumentByName(FOLDER & ASM_FILE)
    swApp.ActivateDoc3 asm.GetTitle, False, 0, errs
    Set c = asm.AddComponent5(p, 0, "", False, "", 0, 0, ZOFF / 1000#)
    swApp.CloseDoc ttl
    swApp.ActivateDoc3 asm.GetTitle, False, 0, errs
    MsgBox "Test spreader part saved=" & ok & ", inserted=" & (Not c Is Nothing) & ", sketch Outline with " & n & " lines."
End Sub
