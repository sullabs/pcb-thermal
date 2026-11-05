Attribute VB_Name = "PTJson"
Option Explicit
' =====================================================================================
' PTJson  -  project folder and JSON files for the PCB thermal macros (SolidWorks 2023 VBA)
'
' The page (webapp/dist/PCB_Thermal.html) writes <prefix>_project.json in the board's project folder.
' The macros read it and write <prefix>_cad.json (and later <prefix>_results.json) next to it, with the same
' prefix (the user's initials, so file names stay unique). No prefix (project.json) works too. Schemas: webapp/schema/.
' Units in the files: mm, C, W, W/(m K), C/W.
'
' Import this module next to PTCommon in each macro that can work on a project folder:
'   ImportStep.swp = PTCommon.bas + PTJson.bas + ImportStep.bas
'   BuildBoard.swp = PTCommon.bas + PTJson.bas + BuildBoard.bas
'
' JSON values in VBA: object = Scripting.Dictionary (keys keep their order), array = Collection,
' text = String, number = Double, true/false = Boolean. Blank values are left out of the files, never null.
' Pure VBA, no references. Files are read as UTF-8 and written as ASCII (other characters as \u escapes).
' =====================================================================================

Public PJ_Dir As String           ' project folder of this run, no "\" at the end ("" in workbook mode)
Public PJ_Project As Object       ' project.json as read at the start of the run
Public PJ_WbPath As String        ' the workbook picked instead (workbook mode)
Public PJ_Pre As String           ' file prefix of this project, e.g. "AB_" for AB_project.json ("" for project.json)

Private pjS As String, pjP As Long, pjN As Long, pjDec As String      ' parser state
Private wBuf() As String, wN As Long                                  ' writer buffer

' ---------------------------------------------------------------- which project
' Asks which project this run works on, with the last one as the default:
'   <prefix>_project.json in a board folder (made by the page) -> "json", PJ_Dir, PJ_Pre and PJ_Project are set
'   an old workbook (.xlsx, .xlsm)                    -> "xlsx", PJ_WbPath is set
' Returns "" when the user cancels or the file cannot be used (the user has been told why).
Public Function PJ_PickSource(ByVal macroName As String) As String
    Dim lastF As String, f As String, ans As VbMsgBoxResult, opt As Long, cfg As String, disp As String
    Dim fso As Object, ext As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    PJ_Dir = "": PJ_WbPath = "": PJ_Pre = ""
    Set PJ_Project = Nothing
    lastF = GetSetting(PT_APP, "Common", "Source", "")
    If lastF = "" Then lastF = GetSetting(PT_APP, "Common", "Workbook", "")
    If lastF <> "" Then
        If Not fso.FileExists(lastF) Then lastF = ""
    End If
    If lastF <> "" Then
        ans = PT_Msg("Work on this project?" & vbCrLf & lastF & vbCrLf & vbCrLf & _
                     "Yes = use it." & vbCrLf & "No = pick another one (<prefix>_project.json in the board folder, or an old workbook).", _
                     vbYesNoCancel + vbQuestion, macroName)
        If ans = vbCancel Then Exit Function
        If ans = vbYes Then f = lastF
    End If
    If f = "" Then
        f = swApp.GetOpenFileName("Pick <prefix>_project.json in the board folder (or an old workbook)", lastF, _
                                  "Project (*project.json)|*project.json|Old workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*||", _
                                  opt, cfg, disp)
        If f = "" Then Exit Function
    End If
    ext = LCase$(fso.GetExtensionName(f))
    If ext = "xlsx" Or ext = "xlsm" Then
        PJ_WbPath = f
        SaveSetting PT_APP, "Common", "Source", f
        PJ_PickSource = "xlsx"
        Exit Function
    End If
    If PJ_PreOf(fso.GetFileName(f)) = "?" Then
        PT_Msg macroName & " did not run. Pick the file named <prefix>_project.json in the board folder (the page saves it there), " & _
               "or an old workbook." & vbCrLf & vbCrLf & "You picked: " & f, vbExclamation, macroName
        Exit Function
    End If
    If Not PJ_LoadProject(f, macroName) Then Exit Function
    SaveSetting PT_APP, "Common", "Source", f
    PJ_PickSource = "json"
End Function

' Prefix of a project file name: "AB_project.json" -> "AB_", "project.json" -> "", anything else -> "?".
Public Function PJ_PreOf(ByVal fileName As String) As String
    Dim n As Long
    PJ_PreOf = "?"
    n = Len(fileName)
    If n < 12 Then Exit Function
    If LCase$(Right$(fileName, 12)) <> "project.json" Then Exit Function
    If n = 12 Then
        PJ_PreOf = ""
        Exit Function
    End If
    If Mid$(fileName, n - 12, 1) <> "_" Then Exit Function
    PJ_PreOf = Left$(fileName, n - 12)
End Function

' A file of this project: PJ_File("cad") -> "AB_cad.json" (name only, in PJ_Dir).
Public Function PJ_File(ByVal kindS As String) As String
    PJ_File = PJ_Pre & kindS & ".json"
End Function

' Reads the project file f (<folder>\<prefix>_project.json) into PJ_Project and sets PJ_Dir and PJ_Pre.
' False if it cannot be read (the user has been told).
Public Function PJ_LoadProject(ByVal f As String, ByVal macroName As String) As Boolean
    Dim d As Object, fso As Object, folderS As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderS = fso.GetParentFolderName(f)
    On Error GoTo bad
    Set d = PJ_Parse(PJ_ReadFile(f))
    On Error GoTo 0
    If PJ_Str(d, "schema", "") <> "pcb-thermal/project" Then
        PT_Msg macroName & " did not run. This is not a project file of the page:" & vbCrLf & f, vbExclamation, macroName
        Exit Function
    End If
    PJ_Dir = folderS
    PJ_Pre = PJ_PreOf(fso.GetFileName(f))
    If PJ_Pre = "?" Then PJ_Pre = ""
    Set PJ_Project = d
    PT_Trace macroName & ": project folder " & folderS & " (board " & PJ_Str(d, "board/name", "?") & _
             ", " & fso.GetFileName(f) & " saved " & PJ_Str(d, "savedAt", "?") & ")"
    PJ_LoadProject = True
    Exit Function
bad:
    PT_Msg macroName & " did not run. The project file could not be read:" & vbCrLf & f & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, macroName
End Function

' ---------------------------------------------------------------- Ansys page (ansys\page.html) board folder
' Asks for study.json in a board folder of the Ansys page, the last one as the default. Sets PJ_Dir, PJ_Project
' (the study) and PJ_Pre = "". False when the user cancels or the file cannot be used (the user has been told why).
Public Function PJ_PickStudy(ByVal macroName As String) As Boolean
    Dim lastF As String, f As String, ans As VbMsgBoxResult, opt As Long, cfg As String, disp As String
    Dim fso As Object, d As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    PJ_Dir = "": PJ_WbPath = "": PJ_Pre = ""
    Set PJ_Project = Nothing
    lastF = GetSetting(PT_APP, "Common", "Study", "")
    If lastF <> "" Then
        If Not fso.FileExists(lastF) Then lastF = ""
    End If
    If lastF <> "" Then
        ans = PT_Msg("Work on this board folder?" & vbCrLf & fso.GetParentFolderName(lastF) & vbCrLf & vbCrLf & _
                     "Yes = use it." & vbCrLf & "No = pick study.json in another board folder.", vbYesNoCancel + vbQuestion, macroName)
        If ans = vbCancel Then Exit Function
        If ans = vbYes Then f = lastF
    End If
    If f = "" Then
        f = swApp.GetOpenFileName("Pick study.json in the board folder (the Ansys page saves it there)", lastF, _
                                  "Board study (study.json)|study.json|All files (*.*)|*.*||", opt, cfg, disp)
        If f = "" Then Exit Function
    End If
    On Error GoTo bad
    Set d = PJ_Parse(PJ_ReadFile(f))
    On Error GoTo 0
    If PJ_Str(d, "schema", "") <> "pcb-thermal-ansys/study" Then
        PT_Msg macroName & " did not run. This is not the study.json of the Ansys page:" & vbCrLf & f, vbExclamation, macroName
        Exit Function
    End If
    PJ_Dir = fso.GetParentFolderName(f)
    Set PJ_Project = d
    SaveSetting PT_APP, "Common", "Study", f
    PT_Trace macroName & ": board folder " & PJ_Dir
    PJ_PickStudy = True
    Exit Function
bad:
    PT_Msg macroName & " did not run. study.json could not be read:" & vbCrLf & f & vbCrLf & vbCrLf & Err.Description, vbExclamation, macroName
End Function

' The board name in board.json next to study.json ("Board" when it cannot be read).
Public Function PJ_BoardName() As String
    On Error Resume Next
    PJ_BoardName = PJ_Str(PJ_Parse(PJ_ReadFile(PJ_Dir & "\board.json")), "name", "Board")
    If PJ_BoardName = "" Then PJ_BoardName = "Board"
End Function

' <folder>\solidworks\<board>_solution.sldasm: the assembly Build Solution makes and Send Solution reads.
Public Function PJ_SolutionAsm() As String
    PJ_SolutionAsm = PJ_Dir & "\solidworks\" & PT_SafeFile(PJ_BoardName()) & "_solution.sldasm"
End Function

' ---------------------------------------------------------------- the macros' own files
' The project's cad or results file (fileName from PJ_File). A file that is not there yet gives a new one with only its header.
' Nothing when the file is there but cannot be read (reason in PT_LastErr): never overwrite it then.
Public Function PJ_ReadOwn(ByVal fileName As String, ByVal schemaS As String) As Object
    Dim f As String, d As Object
    PT_LastErr = ""
    f = PJ_Dir & "\" & fileName
    If Dir(f) = "" Then
        Set d = PJ_Dict()
        PJ_Put d, "schema", schemaS
        PJ_Put d, "schemaVersion", 1
        PJ_Put d, "savedBy", ""
        PJ_Put d, "savedAt", PJ_Now()
        Set PJ_ReadOwn = d
        Exit Function
    End If
    On Error GoTo bad
    Set d = PJ_Parse(PJ_ReadFile(f))
    On Error GoTo 0
    If PJ_Str(d, "schema", "") <> schemaS Then
        PT_LastErr = "its schema is '" & PJ_Str(d, "schema", "") & "', not " & schemaS
        Exit Function
    End If
    Set PJ_ReadOwn = d
    Exit Function
bad:
    PT_LastErr = Err.Description
End Function

' Sets savedBy and savedAt, makes the daily backup of the old file and writes the new one. True when written.
Public Function PJ_SaveOwn(ByVal d As Object, ByVal fileName As String, ByVal macroName As String) As Boolean
    PJ_Put d, "schemaVersion", 1
    PJ_Put d, "savedBy", macroName & " (VBA macro)"
    PJ_Put d, "savedAt", PJ_Now()
    PJ_DailyBackup fileName
    PJ_SaveOwn = PJ_WriteFile(PJ_Dir & "\" & fileName, PJ_Text(d))
    If PJ_SaveOwn Then PT_Trace macroName & ": wrote " & PJ_Dir & "\" & fileName
End Function

' One copy per day in backup\<name>_yyyymmdd.json, made before the first write of the day. A failed copy is ignored.
Public Sub PJ_DailyBackup(ByVal fileName As String)
    Dim fso As Object, src As String, dst As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    src = PJ_Dir & "\" & fileName
    If Not fso.FileExists(src) Then Exit Sub
    PT_MakeDirs PJ_Dir & "\backup"
    dst = PJ_Dir & "\backup\" & fso.GetBaseName(fileName) & "_" & PJ_Day(Now) & ".json"
    If Not fso.FileExists(dst) Then fso.CopyFile src, dst
End Sub

' ---------------------------------------------------------------- settings from project.json
' settings.outputFolder (blank = CAD folder in the project folder; a relative path starts at the project folder).
' Created if missing, ends with "\".
Public Function PJ_OutputFolder() As String
    Dim s As String
    s = PJ_Str(PJ_Project, "settings/outputFolder", "")
    If s = "" Then
        s = PJ_Dir & "\CAD"
    ElseIf InStr(s, ":") = 0 And Left$(s, 2) <> "\\" Then
        s = PJ_Dir & "\" & s
    End If
    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    PT_MakeDirs s
    PJ_OutputFolder = s & "\"
End Function

' settings.filePrefix, blank = board name + "_".
Public Function PJ_Prefix() As String
    PJ_Prefix = PJ_Str(PJ_Project, "settings/filePrefix", "")
    If PJ_Prefix = "" Then PJ_Prefix = PT_SafeFile(PJ_Str(PJ_Project, "board/name", "Board")) & "_"
End Function

' settings.asmName, blank = <prefix>Thermal.sldasm in the output folder.
Public Function PJ_AsmPath() As String
    Dim s As String
    s = PJ_Str(PJ_Project, "settings/asmName", "")
    If s = "" Then s = PT_SafeFile(PJ_Prefix()) & "Thermal.sldasm"
    If InStr(s, "\") = 0 Then s = PJ_OutputFolder() & s
    If LCase$(Right$(s, 7)) <> ".sldasm" Then s = s & ".sldasm"
    PJ_AsmPath = s
End Function

' Template from settings (partTemplate, asmTemplate), else the SolidWorks default; "" if the file is not there.
Public Function PJ_Template(ByVal settingS As String, ByVal swPref As Long) As String
    Dim s As String
    s = PJ_Str(PJ_Project, "settings/" & settingS, "")
    If s = "" Then s = swApp.GetUserPreferenceStringValue(swPref)
    If s <> "" Then
        If Dir(s) = "" Then s = ""
    End If
    PJ_Template = s
End Function

' ---------------------------------------------------------------- change records (cad.json changes[])
' Moves the rows collected by PT_LogAdd into d("changes"), newest first. Each gets an id yyyymmddThhmmss-n,
' under which the page keeps the user's tick (project.json checks). Returns the number added.
Public Function PJ_FlushChanges(ByVal d As Object) As Long
    Dim oldC As Object, newC As Collection, seenIds As Object, it As Variant, i As Long, k As Long, rw As Variant
    Dim rec As Object, baseS As String, idS As String, cnt As Long
    cnt = PT_LogCount()
    If cnt = 0 Then Exit Function
    Set seenIds = CreateObject("Scripting.Dictionary")
    Set oldC = PJ_Obj(d, "changes")
    If Not oldC Is Nothing Then
        For Each it In oldC
            If IsObject(it) Then
                idS = PJ_Str(it, "id", "")
                If idS <> "" Then seenIds(idS) = 1
            End If
        Next
    End If
    Set newC = New Collection
    For i = 1 To cnt
        rw = PT_LogItem(i)
        baseS = PJ_Day(rw(0)) & "T" & PJ_Two(Hour(rw(0))) & PJ_Two(Minute(rw(0))) & PJ_Two(Second(rw(0)))
        k = 1
        Do While seenIds.Exists(baseS & "-" & k): k = k + 1: Loop
        idS = baseS & "-" & k
        seenIds(idS) = 1
        Set rec = PJ_Dict()
        PJ_Put rec, "id", idS
        PJ_Put rec, "date", PJ_Stamp(rw(0))
        PJ_Put rec, "macro", PT_Str(rw(1))
        PJ_Put rec, "key", PT_Str(rw(2))
        PJ_Put rec, "name", PT_Str(rw(3))
        PJ_Put rec, "refDes", PT_Str(rw(4))
        PJ_Put rec, "change", PT_Str(rw(5))
        PJ_Put rec, "field", PT_Str(rw(6))
        PJ_Put rec, "old", ChangeValue(rw(7))
        PJ_Put rec, "new", ChangeValue(rw(8))
        If PT_IsNum(rw(9)) Then PJ_Put rec, "delta", CDbl(rw(9))
        newC.Add rec
    Next
    If Not oldC Is Nothing Then
        For Each it In oldC
            newC.Add it
        Next
    End If
    PJ_Put d, "changes", newC
    PT_LogReset
    PJ_FlushChanges = cnt
End Function

Private Function ChangeValue(ByVal v As Variant) As Variant
    If PT_IsNum(v) Then
        ChangeValue = CDbl(v)
    ElseIf PT_Str(v) <> "" Then
        ChangeValue = PT_Str(v)
    End If
End Function

' Build macros call this first. True = refuse: cad.json has change records the user has not ticked on the page.
' The user has been told which ones.
Public Function PJ_RefuseIfUnchecked(ByVal cad As Object, ByVal macroName As String) As Boolean
    Dim ch As Object, it As Variant, cnt As Long, s As String, idS As String, oc As Object
    ' page checks not yet accepted (forSolidWorks.openChecks, for example heat with no body)
    Set oc = PJ_Obj(PJ_Project, "forSolidWorks/openChecks")
    If Not oc Is Nothing Then
        For Each it In oc
            cnt = cnt + 1
            If cnt <= 8 Then s = s & "  " & CStr(it) & vbCrLf
        Next
        If cnt > 0 Then
            PT_Msg macroName & " did not run. Nothing was changed." & vbCrLf & vbCrLf & s & vbCrLf & _
                   "Fix it on the page (tab 6, Heat with no body), save, and run again.", vbExclamation, macroName
            PJ_RefuseIfUnchecked = True
            Exit Function
        End If
    End If
    Set ch = PJ_Obj(cad, "changes")
    If ch Is Nothing Then Exit Function
    For Each it In ch
        If IsObject(it) Then
            idS = PJ_Str(it, "id", "")
            If idS <> "" And Not PJ_Has(PJ_Project, "checks/" & idS) Then
                cnt = cnt + 1
                If cnt <= 8 Then s = s & "  " & Replace(PJ_Str(it, "date", ""), "T", " ") & "  " & PJ_Str(it, "change", "") & " " & _
                                     PJ_Str(it, "field", "") & "  " & PJ_Str(it, "name", PJ_Str(it, "key", "")) & vbCrLf
            End If
        End If
    Next
    If cnt = 0 Then Exit Function
    If cnt > 8 Then s = s & "  ..." & vbCrLf
    PT_Msg macroName & " did not run. Nothing was changed." & vbCrLf & vbCrLf & _
           cnt & " change records in cad.json are not ticked:" & vbCrLf & s & vbCrLf & _
           "Open the page, look at each one on tab 10 (Changes), tick it, save, and run again.", vbExclamation, macroName
    PJ_RefuseIfUnchecked = True
End Function

' ---------------------------------------------------------------- reading values
' Value at a path of names, e.g. "settings/import/minSizeMm". v is Empty when a name is missing.
Public Sub PJ_GetV(ByVal root As Object, ByVal pathS As String, ByRef v As Variant)
    Dim segs As Variant, k As Long, cur As Object, seg As String
    v = Empty
    If root Is Nothing Then Exit Sub
    Set cur = root
    segs = Split(pathS, "/")
    For k = 0 To UBound(segs)
        seg = segs(k)
        If TypeName(cur) <> "Dictionary" Then Exit Sub
        If Not cur.Exists(seg) Then Exit Sub
        If k = UBound(segs) Then
            If IsObject(cur(seg)) Then Set v = cur(seg) Else v = cur(seg)
            Exit Sub
        End If
        If Not IsObject(cur(seg)) Then Exit Sub
        Set cur = cur(seg)
    Next
End Sub

Public Function PJ_Has(ByVal root As Object, ByVal pathS As String) As Boolean
    Dim v As Variant
    PJ_GetV root, pathS, v
    If IsObject(v) Then PJ_Has = Not v Is Nothing Else PJ_Has = Not (IsEmpty(v) Or IsNull(v))
End Function

Public Function PJ_Num(ByVal root As Object, ByVal pathS As String, ByVal dflt As Double) As Double
    Dim v As Variant
    PJ_Num = dflt
    PJ_GetV root, pathS, v
    If IsObject(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbLong, vbInteger, vbSingle, vbCurrency, vbDecimal, vbByte: PJ_Num = CDbl(v)
    End Select
End Function

' True when the path holds a number (not text, not missing).
Public Function PJ_IsNum(ByVal root As Object, ByVal pathS As String) As Boolean
    Dim v As Variant
    PJ_GetV root, pathS, v
    If IsObject(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbLong, vbInteger, vbSingle, vbCurrency, vbDecimal, vbByte: PJ_IsNum = True
    End Select
End Function

Public Function PJ_Str(ByVal root As Object, ByVal pathS As String, ByVal dflt As String) As String
    Dim v As Variant
    PJ_Str = dflt
    PJ_GetV root, pathS, v
    If IsObject(v) Then Exit Function
    If VarType(v) = vbString Then
        If v <> "" Then PJ_Str = v
    End If
End Function

Public Function PJ_Bool(ByVal root As Object, ByVal pathS As String, ByVal dflt As Boolean) As Boolean
    Dim v As Variant
    PJ_Bool = dflt
    PJ_GetV root, pathS, v
    If IsObject(v) Then Exit Function
    If VarType(v) = vbBoolean Then PJ_Bool = v
End Function

' Object or array at the path, or Nothing.
Public Function PJ_Obj(ByVal root As Object, ByVal pathS As String) As Object
    Dim v As Variant
    PJ_GetV root, pathS, v
    If IsObject(v) Then Set PJ_Obj = v
End Function

' ---------------------------------------------------------------- making values
Public Function PJ_Dict() As Object
    Set PJ_Dict = CreateObject("Scripting.Dictionary")
End Function

' Sets d(k) = v in place. A blank value (Empty, Null, "", Nothing) removes the name instead: blanks are left out.
Public Sub PJ_Put(ByVal d As Object, ByVal k As String, ByVal v As Variant)
    Dim blank As Boolean
    If IsObject(v) Then
        blank = (v Is Nothing)
    ElseIf IsEmpty(v) Or IsNull(v) Then
        blank = True
    ElseIf VarType(v) = vbString Then
        blank = (v = "")
    End If
    If blank Then
        If d.Exists(k) Then d.Remove k
    ElseIf IsObject(v) Then
        If d.Exists(k) Then Set d(k) = v Else d.Add k, v
    Else
        If d.Exists(k) Then d(k) = v Else d.Add k, v
    End If
End Sub

' Box in board coordinates (mm): xMin, yMin, zMin, l, w, h. Sizes stay above 0 (the schema needs that).
Public Function PJ_Box(ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double, _
                       ByVal lx As Double, ByVal wy As Double, ByVal hz As Double) As Object
    Dim d As Object
    Set d = PJ_Dict()
    PJ_Put d, "xMin", Round(x0, 4)
    PJ_Put d, "yMin", Round(y0, 4)
    PJ_Put d, "zMin", Round(z0, 4)
    PJ_Put d, "l", PJ_Pos(lx)
    PJ_Put d, "w", PJ_Pos(wy)
    PJ_Put d, "h", PJ_Pos(hz)
    Set PJ_Box = d
End Function

Public Function PJ_Pos(ByVal mm As Double) As Double
    PJ_Pos = Round(mm, 4)
    If PJ_Pos < 0.001 Then PJ_Pos = 0.001
End Function

' {level, text} for the messages lists. level: info, warning or error.
Public Function PJ_Message(ByVal lvl As String, ByVal textS As String) As Object
    Dim d As Object
    Set d = PJ_Dict()
    PJ_Put d, "level", lvl
    PJ_Put d, "text", textS
    Set PJ_Message = d
End Function

' ---------------------------------------------------------------- dates (local time, no zone)
Public Function PJ_Now() As String
    PJ_Now = PJ_Stamp(Now)
End Function

' 2026-10-18T11:20:00. Built from numbers, so the Windows date format does not matter.
Public Function PJ_Stamp(ByVal d As Date) As String
    PJ_Stamp = CStr(Year(d)) & "-" & PJ_Two(Month(d)) & "-" & PJ_Two(Day(d)) & "T" & _
               PJ_Two(Hour(d)) & ":" & PJ_Two(Minute(d)) & ":" & PJ_Two(Second(d))
End Function

Public Function PJ_Day(ByVal d As Date) As String
    PJ_Day = CStr(Year(d)) & PJ_Two(Month(d)) & PJ_Two(Day(d))
End Function

Public Function PJ_Two(ByVal num As Long) As String
    PJ_Two = Right$("0" & CStr(num), 2)
End Function

' ---------------------------------------------------------------- files
' Whole file as text, read as UTF-8 (the page writes UTF-8). Raises an error when the file cannot be read.
Public Function PJ_ReadFile(ByVal fpath As String) As String
    Dim st As Object, fnum As Integer, raw As String
    On Error GoTo noAdo
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile fpath
    PJ_ReadFile = st.ReadText(-1)
    st.Close
    Exit Function
noAdo:
    ' no ADODB: read the bytes as they are (only texts outside ASCII come out wrong)
    fnum = FreeFile
    Open fpath For Binary Access Read As #fnum
    raw = Space$(LOF(fnum))
    Get #fnum, , raw
    Close #fnum
    PJ_ReadFile = raw
End Function

' Writes textS (ASCII) to a temporary file next to fpath, then puts it in place of fpath, so the page never
' reads half a file. True when written; reason in PT_LastErr otherwise.
Public Function PJ_WriteFile(ByVal fpath As String, ByVal textS As String) As Boolean
    Dim fso As Object, tmpP As String, ts As Object, tryN As Long, t0 As Single
    PT_LastErr = ""
    On Error GoTo failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmpP = fpath & ".tmp"
    If fso.FileExists(tmpP) Then fso.DeleteFile tmpP, True
    Set ts = fso.CreateTextFile(tmpP, True, False)
    ts.Write textS
    ts.Close
    On Error Resume Next
    For tryN = 1 To 5
        Err.Clear
        If fso.FileExists(fpath) Then fso.DeleteFile fpath, True
        If Err.Number = 0 Then fso.MoveFile tmpP, fpath
        If Err.Number = 0 Then PJ_WriteFile = True: Exit Function
        PT_LastErr = Err.Description
        t0 = Timer
        Do While Timer - t0 < 0.5 And Timer >= t0       ' the page or a virus scanner may hold the file for a moment
            DoEvents
        Loop
    Next
    PT_LastErr = PT_LastErr & " (the new text is in " & tmpP & ")"
    Exit Function
failed:
    PT_LastErr = Err.Description
End Function

' ---------------------------------------------------------------- JSON text -> values
' Object or array from JSON text. Raises an error naming the character position when the text is not JSON.
Public Function PJ_Parse(ByVal textS As String) As Object
    Dim v As Variant
    pjS = textS: pjP = 1: pjN = Len(textS)
    pjDec = Mid$(CStr(1.5), 2, 1)                 ' decimal separator of this PC, for CDbl
    If pjN > 0 Then
        If Mid$(pjS, 1, 1) = ChrW(&HFEFF) Then pjP = 2
    End If
    PValue v
    PSkip
    If pjP <= pjN Then PFail "more text after the end"
    If Not IsObject(v) Then PFail "the file is not a JSON object"
    Set PJ_Parse = v
    pjS = ""
End Function

Private Sub PFail(ByVal why As String)
    Err.Raise vbObjectError + 600, "PTJson", "JSON: " & why & " (character " & pjP & ")"
End Sub

Private Sub PSkip()
    Dim ch As String
    Do While pjP <= pjN
        ch = Mid$(pjS, pjP, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        pjP = pjP + 1
    Loop
End Sub

Private Sub PValue(ByRef v As Variant)
    Dim ch As String
    PSkip
    If pjP > pjN Then PFail "the text ends too early"
    ch = Mid$(pjS, pjP, 1)
    Select Case ch
        Case "{": Set v = PObject()
        Case "[": Set v = PArray()
        Case """": v = PString()
        Case "t": PWord "true": v = True
        Case "f": PWord "false": v = False
        Case "n": PWord "null": v = Null
        Case Else: v = PNumber()
    End Select
End Sub

Private Sub PWord(ByVal w As String)
    If Mid$(pjS, pjP, Len(w)) <> w Then PFail "unknown word"
    pjP = pjP + Len(w)
End Sub

Private Function PObject() As Object
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    pjP = pjP + 1
    PSkip
    If Mid$(pjS, pjP, 1) = "}" Then
        pjP = pjP + 1
        Set PObject = d
        Exit Function
    End If
    Do
        PSkip
        If Mid$(pjS, pjP, 1) <> """" Then PFail "a name in quotes was expected"
        k = PString()
        PSkip
        If Mid$(pjS, pjP, 1) <> ":" Then PFail "':' was expected"
        pjP = pjP + 1
        v = Empty
        PValue v
        If d.Exists(k) Then d.Remove k
        d.Add k, v
        PSkip
        Select Case Mid$(pjS, pjP, 1)
            Case ",": pjP = pjP + 1
            Case "}": pjP = pjP + 1: Exit Do
            Case Else: PFail "',' or '}' was expected"
        End Select
    Loop
    Set PObject = d
End Function

Private Function PArray() As Collection
    Dim c As Collection, v As Variant
    Set c = New Collection
    pjP = pjP + 1
    PSkip
    If Mid$(pjS, pjP, 1) = "]" Then
        pjP = pjP + 1
        Set PArray = c
        Exit Function
    End If
    Do
        v = Empty
        PValue v
        c.Add v
        PSkip
        Select Case Mid$(pjS, pjP, 1)
            Case ",": pjP = pjP + 1
            Case "]": pjP = pjP + 1: Exit Do
            Case Else: PFail "',' or ']' was expected"
        End Select
    Loop
    Set PArray = c
End Function

Private Function PString() As String
    Dim outS As String, q As Long, bs As Long, e As String
    pjP = pjP + 1
    Do
        q = InStr(pjP, pjS, """")
        If q = 0 Then PFail "a text has no closing quote"
        bs = InStr(pjP, pjS, "\")
        If bs = 0 Or bs > q Then
            outS = outS & Mid$(pjS, pjP, q - pjP)
            pjP = q + 1
            Exit Do
        End If
        outS = outS & Mid$(pjS, pjP, bs - pjP)
        e = Mid$(pjS, bs + 1, 1)
        pjP = bs + 2
        Select Case e
            Case """", "\", "/": outS = outS & e
            Case "b": outS = outS & Chr$(8)
            Case "f": outS = outS & Chr$(12)
            Case "n": outS = outS & vbLf
            Case "r": outS = outS & vbCr
            Case "t": outS = outS & vbTab
            Case "u": outS = outS & ChrW(HexNum(Mid$(pjS, pjP, 4))): pjP = pjP + 4
            Case Else: PFail "unknown escape \" & e
        End Select
    Loop
    PString = outS
End Function

Private Function HexNum(ByVal h As String) As Long
    Dim k As Long, dg As Long, ch As String
    If Len(h) <> 4 Then PFail "bad \u escape"
    For k = 1 To 4
        ch = UCase$(Mid$(h, k, 1))
        dg = InStr("0123456789ABCDEF", ch) - 1
        If dg < 0 Then PFail "bad \u escape"
        HexNum = HexNum * 16 + dg
    Next
End Function

Private Function PNumber() As Double
    Dim q As Long, ch As String, tok As String
    q = pjP
    Do While pjP <= pjN
        ch = Mid$(pjS, pjP, 1)
        If InStr("+-0123456789.eE", ch) = 0 Then Exit Do
        pjP = pjP + 1
    Loop
    If pjP = q Then PFail "unexpected character '" & Mid$(pjS, q, 1) & "'"
    tok = Mid$(pjS, q, pjP - q)
    If pjDec <> "." Then tok = Replace(tok, ".", pjDec)
    If Not IsNumeric(tok) Then PFail "bad number " & Mid$(pjS, q, pjP - q)
    PNumber = CDbl(tok)
End Function

' ---------------------------------------------------------------- values -> JSON text
' Two spaces per level. Small objects of plain values (a box) and short lists of numbers stay on one line.
' Blank values (Empty, Null, "") inside an object are left out.
Public Function PJ_Text(ByVal v As Variant) As String
    wN = 0
    ReDim wBuf(0 To 1023)
    WValue v, ""
    WAdd vbLf
    ReDim Preserve wBuf(0 To wN - 1)
    PJ_Text = Join(wBuf, "")
    Erase wBuf
End Function

Private Sub WAdd(ByVal s As String)
    If wN > UBound(wBuf) Then ReDim Preserve wBuf(0 To 2 * (UBound(wBuf) + 1) - 1)
    wBuf(wN) = s
    wN = wN + 1
End Sub

Private Sub WValue(ByVal v As Variant, ByVal ind As String)
    Dim k As Variant, it As Variant, first As Boolean, flat As Boolean
    If Not IsObject(v) Then
        WAdd ScalarText(v)
        Exit Sub
    End If
    If v Is Nothing Then
        WAdd "{}"
        Exit Sub
    End If
    first = True
    If TypeName(v) = "Dictionary" Then
        flat = IsFlat(v, True)
        WAdd "{"
        For Each k In v.Keys
            If Not IsBlank(v(k)) Then
                If Not first Then WAdd ","
                If flat Then
                    If Not first Then WAdd " "
                Else
                    WAdd vbLf & ind & "  "
                End If
                WAdd QuoteText(CStr(k)) & ": "
                WValue v(k), ind & "  "
                first = False
            End If
        Next
        If Not first And Not flat Then WAdd vbLf & ind
        WAdd "}"
    Else
        flat = IsFlat(v, False)
        WAdd "["
        For Each it In v
            If Not first Then WAdd ","
            If flat Then
                If Not first Then WAdd " "
            Else
                WAdd vbLf & ind & "  "
            End If
            WValue it, ind & "  "
            first = False
        Next
        If Not first And Not flat Then WAdd vbLf & ind
        WAdd "]"
    End If
End Sub

' One line for: an object of at most 6 numbers, true/false or short texts; a list of at most 16 numbers.
Private Function IsFlat(ByVal v As Variant, ByVal isDict As Boolean) As Boolean
    Dim it As Variant, cnt As Long
    If isDict Then
        For Each it In v.Items
            If IsObject(it) Then Exit Function
            If VarType(it) = vbString Then
                If Len(it) > 30 Then Exit Function
            End If
            cnt = cnt + 1
        Next
        IsFlat = (cnt <= 6)
    Else
        For Each it In v
            If IsObject(it) Then Exit Function
            If VarType(it) = vbString Then Exit Function
            cnt = cnt + 1
        Next
        IsFlat = (cnt <= 16)
    End If
End Function

Private Function IsBlank(ByVal v As Variant) As Boolean
    If IsObject(v) Then
        IsBlank = (v Is Nothing)
    ElseIf IsEmpty(v) Or IsNull(v) Then
        IsBlank = True
    ElseIf VarType(v) = vbString Then
        IsBlank = (v = "")
    End If
End Function

Private Function ScalarText(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbBoolean: If v Then ScalarText = "true" Else ScalarText = "false"
        Case vbString: ScalarText = QuoteText(CStr(v))
        Case vbDate: ScalarText = QuoteText(PJ_Stamp(v))
        Case vbDouble, vbLong, vbInteger, vbSingle, vbCurrency, vbDecimal, vbByte: ScalarText = PJ_NumText(CDbl(v))
        Case Else: ScalarText = "null"
    End Select
End Function

' Number as JSON text, with a point as decimal separator whatever the Windows settings ("0.5", not ".5" or "0,5").
Public Function PJ_NumText(ByVal d As Double) As String
    Dim s As String
    s = Trim$(Str$(d))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    PJ_NumText = s
End Function

' "text" with JSON escapes. Everything outside printable ASCII becomes \uXXXX, so the file is pure ASCII.
Private Function QuoteText(ByVal s As String) As String
    Dim i As Long, c As Long, q As Long, outS As String, esc As String
    q = 1
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c < 32 Or c > 126 Or c = 34 Or c = 92 Then
            Select Case c
                Case 34: esc = "\"""
                Case 92: esc = "\\"
                Case 10: esc = "\n"
                Case 13: esc = "\r"
                Case 9: esc = "\t"
                Case Else: esc = "\u" & Right$("000" & LCase$(Hex$(c)), 4)
            End Select
            outS = outS & Mid$(s, q, i - q) & esc
            q = i + 1
        End If
    Next
    QuoteText = """" & outS & Mid$(s, q) & """"
End Function
