Attribute VB_Name = "UpdateMacros"
Option Explicit
' =====================================================================================
' UpdateMacros  -  brings the macro project it sits in up to date from the .bas files
'
' Lives in PCBThermal.swp next to every other module. It imports every .bas file
' from the folder of that .swp (removing the old copy of each module first), compiles
' and saves the project. Its own module is never replaced: after a change to
' UpdateMacros.bas, remove it and import it again by hand.
'
' Needs nothing else from the project, so it also works on a new, empty .swp:
'   Tools > Macro > New, save as PCBThermal.swp, import UpdateMacros.bas, run main.
' =====================================================================================

Private Const OWN_NAME As String = "UpdateMacros"
Private Const VBEXT_STD_MODULE As Long = 1
Private Const ID_SAVE As Long = 3
Private Const ID_COMPILE As Long = 578
Private lastErr As String

Sub main()
    Dim vbeO As Object, proj As Object, fso As Object, fil As Object
    Dim srcDir As String, modName As String, doneS As String, failS As String
    Dim nDone As Long, nFail As Long, msgS As String

    Set vbeO = GetVbe()
    If vbeO Is Nothing Then
        MsgBox "This macro cannot reach its own VBA project, so nothing was changed." & vbCrLf & _
               "Import the .bas files by hand (File > Import File).", vbExclamation, "Update macros"
        Exit Sub
    End If
    Set proj = OwnProject(vbeO)
    If proj Is Nothing Then
        MsgBox "Could not find the macro project that holds " & OWN_NAME & ". Nothing was changed.", vbExclamation, "Update macros"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    srcDir = ProjectFolder(fso, proj)
    If Not fso.FileExists(srcDir & "\PTCommon.bas") Then srcDir = PickFolder()
    If Len(srcDir) = 0 Then Exit Sub
    If Not fso.FileExists(srcDir & "\PTCommon.bas") Then
        MsgBox "No PTCommon.bas in " & srcDir & ". Nothing was changed." & vbCrLf & _
               "Pick the folder with the .bas files.", vbExclamation, "Update macros"
        Exit Sub
    End If

    For Each fil In fso.GetFolder(srcDir).Files
        If LCase$(fso.GetExtensionName(fil.Path)) = "bas" Then
            modName = ModuleNameOf(fso, fil.Path)
            If Len(modName) = 0 Then
                failS = failS & vbCrLf & "  " & fil.Name & ": no Attribute VB_Name line"
                nFail = nFail + 1
            ElseIf StrComp(modName, OWN_NAME, vbTextCompare) <> 0 Then
                If ReplaceModule(proj, modName, fil.Path) Then
                    doneS = doneS & " " & modName
                    nDone = nDone + 1
                Else
                    failS = failS & vbCrLf & "  " & fil.Name & ": " & lastErr
                    nFail = nFail + 1
                End If
            End If
        End If
    Next fil

    ' The editor runs these two only after this macro has ended (seen 2026-11-05), so their result is not known here
    RunCommand vbeO, proj, ID_COMPILE
    RunCommand vbeO, proj, ID_SAVE

    msgS = "Updated " & nDone & " modules in " & ProjectFile(proj) & " from " & srcDir & ":" & vbCrLf & " " & Trim$(doneS)
    If nFail > 0 Then msgS = msgS & vbCrLf & vbCrLf & nFail & " failed:" & failS
    msgS = msgS & vbCrLf & vbCrLf & "The VBA editor compiles and saves the project when you click OK. A compile error shows " & _
           "there with its line; fix the .bas file and run Update Macros again."
    MsgBox msgS, IIf(nFail > 0, vbExclamation, vbInformation), "Update macros"
End Sub

' The VBA editor object, or Nothing when the host does not give it out.
Private Function GetVbe() As Object
    Dim appO As Object
    On Error Resume Next
    Set appO = Application
    Set GetVbe = appO.VBE
    Err.Clear
End Function

' The open project that contains this module.
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

Private Function ProjectFolder(ByVal fso As Object, ByVal proj As Object) As String
    Dim fileS As String
    fileS = ProjectFile(proj)
    If Len(fileS) > 0 Then ProjectFolder = fso.GetParentFolderName(fileS)
End Function

' Name from the first line (Attribute VB_Name = "X"), or "" when the file has none.
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

' Removes the old module (renamed first: VBA removes it only when the macro ends, and the
' import would otherwise get the name PTCommon1) and imports the file. False with lastErr set on failure.
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

' Runs a VBA editor menu command on the project by its control id (3 Save, 578 Compile). Does nothing when missing.
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

Private Function PickFolder() As String
    Dim sh As Object, fld As Object
    On Error Resume Next
    Set sh = CreateObject("Shell.Application")
    Set fld = sh.BrowseForFolder(0, "Pick the folder with the .bas files", 0)
    If Not fld Is Nothing Then PickFolder = fld.Self.Path
    Err.Clear
End Function
