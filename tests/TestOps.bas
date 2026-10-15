Attribute VB_Name = "TestOps"
Option Explicit

Private Sub Lg(ByVal t As String)
    Dim fn As Integer
    fn = FreeFile
    Open "C:\Users\Sultan\AppData\Local\Temp\testops.txt" For Append As #fn
    Print #fn, Format(Now, "hh:nn:ss") & " " & t
    Close #fn
End Sub

Private Function ShiftComp(ByVal asm As SldWorks.AssemblyDoc, ByVal c As SldWorks.Component2, ByVal dxM As Double) As Boolean
    Dim xf As SldWorks.MathTransform, nx As SldWorks.MathTransform, mu As SldWorks.MathUtility
    Dim a As Variant, d(15) As Double, k As Long
    Set xf = c.Transform2
    a = xf.ArrayData
    For k = 0 To 15: d(k) = a(k): Next
    d(9) = d(9) + dxM
    Set mu = Application.SldWorks.GetMathUtility
    Set nx = mu.CreateTransform(d)
    If c.IsFixed Then
        asm.ClearSelection2 True
        c.Select4 False, Nothing, False
        asm.UnfixComponent
        asm.ClearSelection2 True
    End If
    Lg "before x=" & d(9) - dxM & " target x=" & d(9)
    c.Transform2 = nx
    asm.EditRebuild3
    Dim b2 As Variant
    b2 = c.Transform2.ArrayData
    Lg "readback x=" & b2(9) & " y=" & b2(10) & " z=" & b2(11)
    Dim bx As Variant
    bx = c.GetBox(False, False)
    Lg "box xmin=" & bx(0) & " xmax=" & bx(3)
    ShiftComp = True
End Function

' open the STEP assembly (or activate it) so ImportStep can use it
Sub StepOpen()
    Dim sw As SldWorks.SldWorks, md As SldWorks.ModelDoc2, e As Long, w As Long
    Set sw = Application.SldWorks
    Set md = sw.OpenDoc6("C:\Users\Sultan\Desktop\thermal_similation\example\RaspberryPi4\Raspberry Pi 4 Model B.STEP", 2, 1, "", e, w)
    Lg "StepOpen nothing=" & (md Is Nothing) & " err=" & e
End Sub

' move the component whose name contains the text (set in the env file) by dx mm; text and dx are read from testops_arg.txt
Sub StepMove()
    Dim sw As SldWorks.SldWorks, asm As SldWorks.AssemblyDoc, comps As Variant, i As Long, c As SldWorks.Component2
    Dim fn As Integer, key As String, dx As Double, s As String, n As Long
    fn = FreeFile
    Open "C:\Users\Sultan\AppData\Local\Temp\testops_arg.txt" For Input As #fn
    Line Input #fn, key
    Line Input #fn, s
    Close #fn
    dx = CDbl(s)
    Set sw = Application.SldWorks
    Set asm = sw.ActiveDoc
    Lg "StepMove active=" & asm.GetTitle & " key=" & key & " dx=" & dx
    comps = asm.GetComponents(False)
    For i = 0 To UBound(comps)
        Set c = comps(i)
        If InStr(1, c.Name2, key, vbTextCompare) > 0 Then
            n = n + 1
            Lg "match " & c.Name2 & " parent-level fixed=" & c.IsFixed
            On Error Resume Next
            ShiftComp asm, c, dx / 1000#
            Lg "shift err=" & Err.Number & " " & Err.Description
            On Error GoTo 0
            Exit For
        End If
    Next
    Lg "StepMove done matches=" & n
End Sub

' move a component of the thermal assembly (key text in testops_arg.txt), unfix, shift, fix again
Sub AsmMove()
    Dim sw As SldWorks.SldWorks, asm As SldWorks.AssemblyDoc, comps As Variant, i As Long, c As SldWorks.Component2
    Dim fn As Integer, key As String, dx As Double, s As String, n As Long
    fn = FreeFile
    Open "C:\Users\Sultan\AppData\Local\Temp\testops_arg.txt" For Input As #fn
    Line Input #fn, key
    Line Input #fn, s
    Close #fn
    dx = CDbl(s)
    Set sw = Application.SldWorks
    Set asm = sw.ActiveDoc
    comps = asm.GetComponents(True)
    For i = 0 To UBound(comps)
        Set c = comps(i)
        If InStr(1, c.Name2, key, vbTextCompare) > 0 Then
            n = n + 1
            ShiftComp asm, c, dx / 1000#
            c.Select4 False, Nothing, False
            asm.FixComponent
            asm.ClearSelection2 True
            Lg "AsmMove moved " & c.Name2 & " by " & dx & " mm"
            Exit For
        End If
    Next
    Lg "AsmMove done matches=" & n
End Sub

' circular through hole in the spreader part
Sub AddHole()
    Dim sw As SldWorks.SldWorks, asm As SldWorks.AssemblyDoc, comps As Variant, i As Long, c As SldWorks.Component2
    Dim md As SldWorks.ModelDoc2, f As SldWorks.Feature, pl As SldWorks.Feature, e As Long, feat As Object
    Set sw = Application.SldWorks
    Set asm = sw.ActiveDoc
    comps = asm.GetComponents(True)
    For i = 0 To UBound(comps)
        Set c = comps(i)
        If InStr(1, c.Name2, "Spreader", vbTextCompare) > 0 Then Exit For
    Next
    Set md = c.GetModelDoc2
    Lg "AddHole part=" & md.GetTitle
    sw.ActivateDoc3 md.GetTitle, False, 0, e
    Set md = sw.ActiveDoc
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then Set pl = f: Exit Do
        Set f = f.GetNextFeature
    Loop
    md.ClearSelection2 True
    pl.Select2 False, 0
    md.SketchManager.InsertSketch True
    md.SketchManager.CreateCircleByRadius 0.030, 0.044, 0, 0.002
    md.SketchManager.InsertSketch True
    md.ClearSelection2 True
    Set f = md.FirstFeature
    Dim sk As SldWorks.Feature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "ProfileFeature" Then Set sk = f
        Set f = f.GetNextFeature
    Loop
    sk.Select2 False, 0
    Set feat = md.FeatureManager.FeatureCut4(False, False, False, 1, 1, 0.01, 0.01, False, False, False, False, _
                                             0, 0, False, False, False, False, False, True, True, False, False, False, _
                                             0, 0, False, False)
    Lg "AddHole cut nothing=" & (feat Is Nothing)
    md.Save3 1, 0, 0
    sw.ActivateDoc3 asm.GetTitle, False, 0, e
End Sub

Sub ProbeBox()
    Dim sw As SldWorks.SldWorks, asm As SldWorks.AssemblyDoc, comps As Variant, ch As Variant, i As Long, c As SldWorks.Component2
    Dim v As Variant, top As Variant
    Set sw = Application.SldWorks
    Set asm = sw.ActiveDoc
    Lg "ProbeBox doc=" & asm.GetTitle
    asm.ResolveAllLightWeightComponents False
    comps = asm.GetComponents(True)
    Lg "top comps=" & (UBound(comps) + 1)
    Do While UBound(comps) = 0
        ch = comps(0).GetChildren
        If IsEmpty(ch) Then Exit Do
        If UBound(ch) < 0 Then Exit Do
        comps = ch
    Loop
    Lg "after stepping in: n=" & (UBound(comps) + 1)
    For i = 0 To UBound(comps)
        Set c = comps(i)
        If InStr(1, c.Name2, "BCM2711B0", vbTextCompare) > 0 Then
            v = c.GetBox(False, False)
            Lg "via children: " & c.Name2 & " xmin=" & v(0) & " xmax=" & v(3)
        End If
    Next
    comps = asm.GetComponents(False)
    For i = 0 To UBound(comps)
        Set c = comps(i)
        If InStr(1, c.Name2, "PCB", vbTextCompare) > 0 Or InStr(1, c.Name2, "Model B.STEP-1/Female Micro HDMI Connector, RPi4ModelB.STEP-1", vbTextCompare) > 0 Or InStr(1, c.Name2, "BCM2711B0", vbTextCompare) > 0 Then
            v = c.GetBox(False, False)
            Lg "via all: " & c.Name2 & " xmin=" & v(0) & " xmax=" & v(3)
        End If
    Next
End Sub

Private Sub LogBoxes(ByVal asm As SldWorks.AssemblyDoc, ByVal tag As String)
    Dim comps As Variant, i As Long, c As SldWorks.Component2, v As Variant
    comps = asm.GetComponents(False)
    For i = 0 To UBound(comps)
        Set c = comps(i)
        If InStr(1, c.Name2, "PCB", vbTextCompare) > 0 Or InStr(1, c.Name2, "BCM2711B0", vbTextCompare) > 0 Then
            v = c.GetBox(False, False)
            Lg tag & " " & c.Name2 & " xmin=" & v(0) & " xmax=" & v(3)
        End If
    Next
End Sub

Sub StepMoveSub()
    Dim sw As SldWorks.SldWorks, asm As SldWorks.AssemblyDoc, sub1 As SldWorks.AssemblyDoc, comps As Variant, cs As Variant
    Dim i As Long, topC As SldWorks.Component2, c As SldWorks.Component2
    Dim fn As Integer, key As String, dx As Double, s As String
    fn = FreeFile
    Open "C:\Users\Sultan\AppData\Local\Temp\testops_arg.txt" For Input As #fn
    Line Input #fn, key
    Line Input #fn, s
    Close #fn
    dx = CDbl(s)
    Set sw = Application.SldWorks
    Set asm = sw.ActiveDoc
    LogBoxes asm, "before"
    comps = asm.GetComponents(True)
    Set topC = comps(0)
    Set sub1 = topC.GetModelDoc2
    Lg "sub doc type=" & sub1.GetType & " title=" & sub1.GetTitle
    cs = sub1.GetComponents(True)
    For i = 0 To UBound(cs)
        Set c = cs(i)
        If InStr(1, c.Name2, key, vbTextCompare) > 0 Then
            Lg "sub match " & c.Name2 & " fixed=" & c.IsFixed
            ShiftComp sub1, c, dx / 1000#
            Exit For
        End If
    Next
    asm.EditRebuild3
    LogBoxes asm, "after"
End Sub
