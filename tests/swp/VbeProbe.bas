Attribute VB_Name = "VbeProbe"
' Test only: can a SolidWorks macro reach its own VBA project (to import .bas files)?
' Import into any new macro, run VbeProbe.main, send the text of the message box. Do not save the macro.
Option Explicit

Sub main()
    Dim txt As String, appO As Object, vbeO As Object, proj As Object, comp As Object
    Dim fso As Object, ts As Object, tmpPath As String

    On Error Resume Next
    Set appO = Application
    Set vbeO = appO.VBE
    If Err.Number <> 0 Or vbeO Is Nothing Then
        txt = "1 VBE: FAILED " & Err.Number & " " & Err.Description
        Err.Clear
        MsgBox txt, vbExclamation, "VbeProbe"
        Exit Sub
    End If
    txt = "1 VBE: OK, version " & vbeO.Version
    txt = txt & vbCrLf & "2 Projects open: " & vbeO.VBProjects.Count
    If Err.Number <> 0 Then txt = txt & " FAILED " & Err.Number & " " & Err.Description: Err.Clear

    Set proj = vbeO.ActiveVBProject
    txt = txt & vbCrLf & "3 Active project: " & proj.Name & ", file " & proj.FileName
    If Err.Number <> 0 Then txt = txt & " FAILED " & Err.Number & " " & Err.Description: Err.Clear
    For Each comp In proj.VBComponents
        txt = txt & vbCrLf & "   module " & comp.Name
    Next comp
    If Err.Number <> 0 Then txt = txt & vbCrLf & "   list FAILED " & Err.Number & " " & Err.Description: Err.Clear

    Set fso = CreateObject("Scripting.FileSystemObject")
    tmpPath = Environ$("TEMP") & "\ProbeImported.bas"
    Set ts = fso.CreateTextFile(tmpPath, True, False)
    ts.Write "Attribute VB_Name = ""ProbeImported""" & vbCrLf & "Sub hello()" & vbCrLf & "End Sub" & vbCrLf
    ts.Close
    Set comp = proj.VBComponents.Import(tmpPath)
    If Err.Number <> 0 Then
        txt = txt & vbCrLf & "4 Import: FAILED " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        txt = txt & vbCrLf & "4 Import: OK, module " & comp.Name
        proj.VBComponents.Remove comp
        If Err.Number <> 0 Then
            txt = txt & vbCrLf & "5 Remove: FAILED " & Err.Number & " " & Err.Description
            Err.Clear
        Else
            txt = txt & vbCrLf & "5 Remove: OK"
        End If
    End If
    fso.DeleteFile tmpPath
    Err.Clear

    MsgBox txt, vbInformation, "VbeProbe"
End Sub
