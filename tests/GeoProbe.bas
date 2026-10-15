Attribute VB_Name = "GeoProbe"
Option Explicit
' Writes the bounding box (assembly frame, mm) of every top-level component to tests\geoprobe_out.txt.
Sub main()
    Dim swApp As Object, asm As Object, comps As Variant, i As Long, c As Object, bb As Variant, s As String, errs As Long
    Set swApp = Application.SldWorks
    Set asm = swApp.GetOpenDocumentByName("C:\Users\Sultan\Desktop\thermal_similation\example\RaspberryPi4\CAD\RPi4_fresh_Thermal.sldasm")
    swApp.ActivateDoc3 asm.GetTitle, False, 0, errs
    comps = asm.GetComponents(True)
    For i = 0 To UBound(comps)
        Set c = comps(i)
        bb = c.GetBox(False, False)
        s = s & c.Name2 & ": x " & Format(bb(0) * 1000, "0.000") & " to " & Format(bb(3) * 1000, "0.000") & _
            " | y " & Format(bb(1) * 1000, "0.000") & " to " & Format(bb(4) * 1000, "0.000") & _
            " | z " & Format(bb(2) * 1000, "0.000") & " to " & Format(bb(5) * 1000, "0.000") & vbCrLf
    Next
    Dim fn As Integer: fn = FreeFile
    Open "C:\Users\Sultan\Desktop\thermal_similation\tests\geoprobe_out.txt" For Output As #fn
    Print #fn, s
    Close #fn
End Sub
