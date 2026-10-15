Attribute VB_Name = "ColorProbe"
Option Explicit

Private Function Fmt(ByVal v As Variant) As String
    Dim i As Long, s As String
    If IsEmpty(v) Then Fmt = "(empty)": Exit Function
    If Not IsArray(v) Then Fmt = "(not array)": Exit Function
    For i = LBound(v) To UBound(v): s = s & Format(CDbl(v(i)), "0.00") & " ": Next
    Fmt = s
End Function

Sub main()
    Dim swApp As Object, md As Object, root As Object, vc As Variant, i As Long, c As Object
    Dim out As String, mv As Variant, rd As Object, bodies As Variant, b As Object, f As Variant, fc As Object
    Dim vv(8) As Double, cfgNames As Variant
    Set swApp = Application.SldWorks
    Set md = swApp.ActiveDoc
    Set root = md.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)
    vc = root.GetChildren
    For i = 0 To UBound(vc)
        Set c = vc(i)
        mv = c.GetMaterialPropertyValues2(1, Empty)
        out = out & c.Name2 & " comp(this cfg): " & Fmt(mv)
        mv = c.GetMaterialPropertyValues2(2, Empty)
        out = out & " | all cfg: " & Fmt(mv)
        Set rd = c.GetModelDoc2
        If Not rd Is Nothing Then
            out = out & " | part mat: " & Fmt(rd.MaterialPropertyValues)
            bodies = rd.GetBodies2(0, True)
            If Not IsEmpty(bodies) Then
                Set b = bodies(0)
                out = out & " | body: " & Fmt(b.MaterialPropertyValues2)
                f = b.GetFaces
                Set fc = f(0)
                out = out & " | face0: " & Fmt(fc.MaterialPropertyValues)
            End If
        End If
        out = out & vbCrLf
    Next
    Dim fn As Integer: fn = FreeFile
    Open "C:\Users\Sultan\Desktop\thermal_similation\tests\colorprobe_out.txt" For Output As #fn
    Print #fn, out
    Close #fn
End Sub
