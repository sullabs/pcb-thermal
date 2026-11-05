Attribute VB_Name = "BuildMacros"
Option Explicit
' =====================================================================================
' BuildMacros  -  UpdateMacros with no dialogs, for BuildMacros.cmd (run from a shell)
'
' main imports every .bas file of the folder of PCBThermal.swp into it (UpdateMacros too,
' but not this module) and asks the VBA editor to compile and save. The editor does both
' only after the macro has ended, so BuildMacros.vbs waits for the new .swp and then runs
' Check, which tells whether the project compiled. Both write %TEMP%\PCBThermal_build.log.
' UpdateMacros updates this module, this module updates UpdateMacros.
' A compile error shows the VBA editor's own message: click OK, Check then names the line.
' =====================================================================================

Private Const OWN_NAME As String = "BuildMacros"
Private Const VBEXT_STD_MODULE As Long = 1
Private Const ID_SAVE As Long = 3
Private Const ID_COMPILE As Long = 578
Private lastErr As String

Sub main()
    Dim vbeO As Object, proj As Object, fso As Object, fil As Object
    Dim srcDir As String, modName As String, logS As String
    Dim nDone As Long, nFail As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set vbeO = GetVbe()
    If vbeO Is Nothing Then
        WriteLog fso, "FAIL no VBA editor object (Application.VBE)"
        Exit Sub
    End If
    Set proj = OwnProject(vbeO)
    If proj Is Nothing Then
        WriteLog fso, "FAIL no open project holds " & OWN_NAME
        Exit Sub
    End If
    srcDir = fso.GetParentFolderName(ProjectFile(proj))
    logS = "project " & ProjectFile(proj) & vbCrLf

    For Each fil In fso.GetFolder(srcDir).Files
        If LCase$(fso.GetExtensionName(fil.Path)) = "bas" Then
            modName = ModuleNameOf(fso, fil.Path)
            If Len(modName) = 0 Then
                logS = logS & "FAIL " & fil.Name & ": no Attribute VB_Name line" & vbCrLf
                nFail = nFail + 1
            ElseIf StrComp(modName, OWN_NAME, vbTextCompare) <> 0 Then
                If ReplaceModule(proj, modName, fil.Path) Then
                    nDone = nDone + 1
                Else
                    logS = logS & "FAIL " & fil.Name & ": " & lastErr & vbCrLf
                    nFail = nFail + 1
                End If
            End If
        End If
    Next fil
    logS = logS & "imported " & nDone & " modules" & vbCrLf
    If nFail > 0 Then logS = logS & "FAIL import" & vbCrLf
    RunCommand vbeO, proj, ID_COMPILE
    RunCommand vbeO, proj, ID_SAVE
    WriteLog fso, logS
End Sub

' After main, once the editor has compiled and saved: the Compile command is greyed out when the project is
' compiled, so a command that is still enabled means a compile error, and the editor has its line selected.
Sub Check()
    Dim vbeO As Object, proj As Object, fso As Object, ctl As Object, cp As Object
    Dim sl As Long, sc As Long, el As Long, ec As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set vbeO = GetVbe()
    If vbeO Is Nothing Then Exit Sub
    Set proj = OwnProject(vbeO)
    If proj Is Nothing Then Exit Sub
    On Error Resume Next
    Set vbeO.ActiveVBProject = proj
    Set ctl = vbeO.CommandBars.FindControl(, ID_COMPILE)
    If ctl Is Nothing Then
        WriteLog fso, "FAIL no Compile command in the VBA editor"
        Exit Sub
    End If
    If Not ctl.Enabled Then
        WriteLog fso, "compiled" & vbCrLf & "OK"
        Exit Sub
    End If
    Set cp = vbeO.ActiveCodePane
    If cp Is Nothing Then
        WriteLog fso, "FAIL compile error (no code pane to show where)"
    Else
        cp.GetSelection sl, sc, el, ec
        WriteLog fso, "FAIL compile error in " & cp.CodeModule.Parent.Name & " line " & sl & ": " & Trim$(cp.CodeModule.Lines(sl, 1))
    End If
    Err.Clear
End Sub

Private Sub WriteLog(ByVal fso As Object, ByVal textS As String)
    Dim ts As Object
    On Error Resume Next
    Set ts = fso.OpenTextFile(Environ("TEMP") & "\PCBThermal_build.log", 8, True)      ' append
    ts.Write Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & textS & vbCrLf
    ts.Close
    Err.Clear
End Sub

Private Function GetVbe() As Object
    Dim appO As Object
    On Error Resume Next
    Set appO = Application
    Set GetVbe = appO.VBE
    Err.Clear
End Function

Private Function OwnProject(ByVal vbeO As Object) As Object
    Dim proj As Object
    For Each proj In vbeO.VBProjects
        If Not FindComponent(proj, OWN_NAME) Is Nothing Then
            Set OwnProject = proj
            Exit Function
        End If
    Next proj
End Function

Private Function FindComponent(ByVal proj As Object, ByVal modName As String) As Object
    On Error Resume Next
    Set FindComponent = proj.VBComponents(modName)
    Err.Clear
End Function

Private Function ProjectFile(ByVal proj As Object) As String
    On Error Resume Next
    ProjectFile = proj.FileName
    Err.Clear
End Function

Private Function ModuleNameOf(ByVal fso As Object, ByVal pathS As String) As String
    Dim ts As Object, lineS As String, p As Long
    On Error Resume Next
    Set ts = fso.OpenTextFile(pathS, 1)
    If ts Is Nothing Then Exit Function
    lineS = ts.ReadLine
    ts.Close
    Err.Clear
    If InStr(1, lineS, "Attribute VB_Name", vbTextCompare) <> 1 Then Exit Function
    p = InStr(lineS, """")
    If p = 0 Then Exit Function
    ModuleNameOf = Mid$(lineS, p + 1, InStrRev(lineS, """") - p - 1)
End Function

' Same as UpdateMacros.ReplaceModule: rename first, VBA removes the old module only when the macro ends.
Private Function ReplaceModule(ByVal proj As Object, ByVal modName As String, ByVal pathS As String) As Boolean
    Dim comp As Object, newComp As Object
    lastErr = ""
    On Error GoTo failed
    Set comp = FindComponent(proj, modName)
    If Not comp Is Nothing Then
        If comp.Type <> VBEXT_STD_MODULE Then
            lastErr = modName & " exists but is not a standard module"
            Exit Function
        End If
        comp.Name = modName & "_old"
        proj.VBComponents.Remove comp
    End If
    Set newComp = proj.VBComponents.Import(pathS)
    If StrComp(newComp.Name, modName, vbTextCompare) <> 0 Then newComp.Name = modName
    ReplaceModule = True
    Exit Function
failed:
    lastErr = Err.Description
End Function

Private Sub RunCommand(ByVal vbeO As Object, ByVal proj As Object, ByVal idN As Long)
    Dim ctl As Object
    On Error Resume Next
    vbeO.MainWindow.Visible = True        ' the editor's commands do nothing until its window was shown once
    Set vbeO.ActiveVBProject = proj
    Set ctl = vbeO.CommandBars.FindControl(, idN)
    If Not ctl Is Nothing Then
        If ctl.Enabled Then ctl.Execute
    End If
    Err.Clear
End Sub
