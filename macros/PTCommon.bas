Attribute VB_Name = "PTCommon"
Option Explicit
' =====================================================================================
' PTCommon  -  shared code for the PCB thermal macros (SolidWorks 2023 VBA)
'
' Import this module into EVERY macro project, next to that macro's own module:
'   ImportStep.swp    = PTCommon.bas + ImportStep.bas
'   SetupWorkbook.swp = PTCommon.bas + SetupWorkbook.bas
'   (and so on for the later macros)
'
' Nothing here is specific to one board. All cell positions come from workbook names
' (defined in PCB_thermal_TEMPLATE.xlsx, or added to an older workbook by SetupWorkbook).
' Tabs are found by base name with or without the template's step number (PT_Sheet).
' =====================================================================================

Public Const PT_APP As String = "PCBThermal"
Public Const PT_LOG_SHEET As String = "Change log"
Public Const PT_LOG_COLS As Long = 11     ' Date Macro Key Name RefDes Change Field Old New Delta Checked

Public swApp As Object
Private logRows As Collection
Private reInst As Object, reExt As Object, reJunk As Object, reTail As Object, reDup As Object      ' regex objects made once (NormKey, PT_Key)
Public PT_LastErr As String
Public PT_ShapeRot As Double      ' rotated box part: degrees, counter-clockwise seen from +Z, about the footprint centre
Public PT_ShapeL As Double        ' rotated box part: rectangle size along its own X (mm)
Public PT_ShapeW As Double        ' rotated box part: rectangle size along its own Y (mm)
Private Const PT_PI As Double = 3.14159265358979
Public Const PT_swSolidBody As Long = 0
Public Const PT_swSuppressed As Long = 0
Public Const PT_swFullyResolved As Long = 2
Public Const PT_swDeleteAbsorbed As Long = 2
Public Const PT_swSaveSilent As Long = 1
Public Const PT_swSaveReferenced As Long = 4
Public Const PT_swOpenSilent As Long = 1
Public Const PT_swDocPART As Long = 1
Public Const PT_swDocASSEMBLY As Long = 2

' ---------------------------------------------------------------- Excel and workbook
Public Function PT_GetExcel() As Object
    On Error Resume Next
    Set PT_GetExcel = GetObject(, "Excel.Application")
    On Error GoTo 0
    If PT_GetExcel Is Nothing Then Set PT_GetExcel = CreateObject("Excel.Application")
    PT_GetExcel.Visible = True
End Function

' Asks which workbook to use (starts at the last one used), opens it if needed and
' checks the tabs in needSheets ("Tab one|Tab two"). Returns Nothing if the user cancels
' or a check fails (the user has been told why).
Public Function PT_PickWorkbook(ByVal xl As Object, ByVal macroName As String, ByVal needSheets As String) As Object
    Dim f As String, opt As Long, cfg As String, disp As String, lastF As String
    Dim wbk As Object, s As Variant, missing As String
    lastF = GetSetting(PT_APP, "Common", "Workbook", "")
    If lastF = "" Then lastF = GetSetting(PT_APP, "ImportStep", "Workbook", "")
    ' the workbook used last time is still open in Excel: use it without asking (close it to pick another one)
    If lastF <> "" Then
        For Each wbk In xl.Workbooks
            If StrComp(wbk.FullName, lastF, vbTextCompare) = 0 Then Set PT_PickWorkbook = wbk: Exit For
        Next
    End If
    If PT_PickWorkbook Is Nothing Then
        f = swApp.GetOpenFileName("Pick the project workbook (a copy of PCB_thermal_TEMPLATE)", lastF, _
                                  "Excel workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm||", opt, cfg, disp)
        If f = "" Then Exit Function
        For Each wbk In xl.Workbooks
            If StrComp(wbk.FullName, f, vbTextCompare) = 0 Then Set PT_PickWorkbook = wbk: Exit For
        Next
        If PT_PickWorkbook Is Nothing Then Set PT_PickWorkbook = xl.Workbooks.Open(f)
    End If
    f = PT_PickWorkbook.FullName
    If PT_PickWorkbook.ReadOnly Then
        MsgBox "That workbook is open read-only (maybe open on another PC or in another Excel window)." & vbCrLf & _
               "Close it there and run again.", vbExclamation, macroName
        Set PT_PickWorkbook = Nothing: Exit Function
    End If
    PT_DailyBackup PT_PickWorkbook
    For Each s In Split(needSheets, "|")
        If Not PT_HasSheet(PT_PickWorkbook, CStr(s)) Then missing = missing & vbCrLf & "  " & s
    Next
    If missing <> "" Then
        MsgBox "That workbook is missing these tabs:" & missing & vbCrLf & vbCrLf & _
               "Pick a copy of PCB_thermal_TEMPLATE.xlsx.", vbExclamation, macroName
        Set PT_PickWorkbook = Nothing: Exit Function
    End If
    SaveSetting PT_APP, "Common", "Workbook", f
End Function

' One copy of the workbook per day in the folder "backup" next to it (<name>_yyyymmdd.xlsx), made before the first macro of the
' day writes to it. Never stops a macro: a failed copy is ignored.
Public Sub PT_DailyBackup(ByVal wb As Object)
    Dim dirS As String, nm As String, dotP As Long, target As String
    On Error Resume Next
    If wb.Path = "" Then Exit Sub
    dirS = wb.Path & "\backup"
    PT_MakeDirs dirS
    nm = wb.Name
    dotP = InStrRev(nm, ".")
    If dotP > 1 Then target = dirS & "\" & Left(nm, dotP - 1) & "_" & Format(Now, "yyyymmdd") & Mid(nm, dotP) Else target = dirS & "\" & nm & "_" & Format(Now, "yyyymmdd")
    If Dir(target) = "" Then wb.SaveCopyAs target
End Sub

' Keys ("key|text") of the Change log rows whose Change column is changeText and whose Checked cell is not Yes.
' SyncFromCAD uses it so a check that failed before and is still unchecked is not logged a second time.
Public Function PT_OpenLogKeys(ByVal wb As Object, ByVal changeText As String) As Object
    Dim d As Object, hdr As Object, ws As Object, lastR As Long, r As Long, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    Set PT_OpenLogKeys = d
    Set hdr = PT_Rng(wb, "Log_Header")
    If hdr Is Nothing Then Exit Function
    Set ws = hdr.Worksheet
    lastR = ws.Cells(ws.Rows.Count, hdr.Column + 5).End(-4162).Row
    If lastR <= hdr.Row Then Exit Function
    v = ws.Range(ws.Cells(hdr.Row + 1, hdr.Column), ws.Cells(lastR, hdr.Column + PT_LOG_COLS - 1)).Value
    For r = 1 To UBound(v, 1)
        If PT_Str(v(r, 6)) = changeText And UCase(PT_Str(v(r, 11))) <> "YES" Then
            If Not d.Exists(PT_Str(v(r, 3)) & "|" & PT_Str(v(r, 7))) Then d.Add PT_Str(v(r, 3)) & "|" & PT_Str(v(r, 7)), 1
        End If
    Next
End Function

' True when every workbook name in needNames ("A|B|C") exists. Otherwise tells the user.
Public Function PT_NeedNames(ByVal wb As Object, ByVal needNames As String, ByVal macroName As String) As Boolean
    Dim s As Variant, missing As String
    For Each s In Split(needNames, "|")
        If PT_Rng(wb, CStr(s)) Is Nothing Then missing = missing & " " & s
    Next
    If missing <> "" Then
        MsgBox macroName & " did not run. Nothing was changed." & vbCrLf & vbCrLf & _
               "The workbook has no cells named:" & missing & vbCrLf & vbCrLf & _
               "Use a copy of PCB_thermal_TEMPLATE.xlsx. For an older workbook, run SetupWorkbook (Prepare Workbook) once, " & _
               "then run " & macroName & " again.", vbExclamation, macroName
        Exit Function
    End If
    PT_NeedNames = True
End Function

Public Function PT_HasSheet(ByVal wb As Object, ByVal sheetName As String) As Boolean
    PT_HasSheet = Not PT_Sheet(wb, sheetName) Is Nothing
End Function

' Tab by its base name. The template puts a step number in front ("7 Spreader"), older copies
' do not ("Spreader"). Both are found. Returns Nothing if there is no such tab.
Public Function PT_Sheet(ByVal wb As Object, ByVal baseName As String) As Object
    Dim ws As Object, nm As String, p As Long
    On Error Resume Next
    Set PT_Sheet = wb.Worksheets(baseName)
    On Error GoTo 0
    If Not PT_Sheet Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        nm = ws.Name
        p = InStr(nm, " ")
        If p > 1 And Left(nm, 1) Like "[0-9]" Then
            If StrComp(Mid(nm, p + 1), baseName, vbTextCompare) = 0 Then Set PT_Sheet = ws: Exit Function
        End If
    Next
End Function

' Range behind a workbook name, or Nothing.
Public Function PT_Rng(ByVal wb As Object, ByVal nameS As String) As Object
    On Error Resume Next
    Set PT_Rng = wb.Names(nameS).RefersToRange
End Function

Public Function PT_NameNum(ByVal wb As Object, ByVal nameS As String, ByVal dflt As Double) As Double
    Dim r As Object
    Set r = PT_Rng(wb, nameS)
    If r Is Nothing Then PT_NameNum = dflt Else PT_NameNum = PT_Num(r.Cells(1, 1).Value, dflt)
End Function

Public Function PT_NameStr(ByVal wb As Object, ByVal nameS As String, ByVal dflt As String) As String
    Dim r As Object
    Set r = PT_Rng(wb, nameS)
    If Not r Is Nothing Then PT_NameStr = PT_Str(r.Cells(1, 1).Value)
    If PT_NameStr = "" Then PT_NameStr = dflt
End Function

'  Fills the chip rows of the Spreader tab (Spr_Chips) and the Sim setup tab (Sim_Chips) from the
' RefDes column of CAD geometry, one row per included RefDes. Only a table that is still empty is filled,
' so typed rows are never overwritten. Returns the number of rows written.
Public Function PT_FillChipLists(ByVal wb As Object) As Long
    Dim geo As Variant, rSpr As Object, rSim As Object, seen As Object
    Dim r As Long, n As Long, refD As String, sideS As String, zmin As Double, hgt As Double, tPcb As Double
    Dim onTop As Boolean, doSpr As Boolean, doSim As Boolean
    Set rSpr = PT_Rng(wb, "Spr_Chips")
    Set rSim = PT_Rng(wb, "Sim_Chips")
    If PT_Rng(wb, "Geo_Table") Is Nothing Then Exit Function
    If Not rSpr Is Nothing Then doSpr = (wb.Application.WorksheetFunction.CountA(rSpr.Columns(1)) = 0)
    If Not rSim Is Nothing Then doSim = (wb.Application.WorksheetFunction.CountA(rSim.Columns(1)) = 0)
    If Not (doSpr Or doSim) Then Exit Function
    geo = PT_Rng(wb, "Geo_Table").Value          ' column 1 = B; 5 RefDes, 6 Include, 13 side, 22 Z min, 25 H
    tPcb = PT_Num(PT_Rng(wb, "Geo_PCB_Used").Cells(3, 1).Value, 0)
    Set seen = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(geo, 1)
        refD = PT_Str(geo(r, 5))
        If refD <> "" And UCase(PT_Str(geo(r, 6))) = "YES" And Not seen.Exists(UCase(refD)) Then
            seen.Add UCase(refD), 1
            If doSpr Then
                If n >= rSpr.Rows.Count Then Exit For
            End If
            If doSim Then
                If n >= rSim.Rows.Count Then Exit For
            End If
            sideS = PT_Str(geo(r, 13))
            zmin = PT_Num(geo(r, 22), 0)
            hgt = PT_Num(geo(r, 25), 99)
            If sideS = "" Then onTop = (zmin >= tPcb - 0.01) Else onTop = (UCase(sideS) = "TOP")
            If doSpr Then
                rSpr.Cells(n + 1, 1).Value = refD
                ' low parts on the top side are covered by default; tall ones (connectors, big inductors) are not
                rSpr.Cells(n + 1, 8).Value = IIf(onTop And hgt <= 3.5, "Yes", "No")
                rSpr.Cells(n + 1, 9).Value = "Pedestal"
                rSpr.Cells(n + 1, 10).Value = "Body"
                rSpr.Cells(n + 1, 11).Value = 0.2
                rSpr.Cells(n + 1, 12).Value = 3
            End If
            If doSim Then
                rSim.Cells(n + 1, 1).Value = refD
                rSim.Cells(n + 1, 2).Value = "Block"
            End If
            n = n + 1
        End If
    Next
    PT_FillChipLists = n
End Function

' ---------------------------------------------------------------- values
Public Function PT_IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    PT_IsNum = IsNumeric(v)
End Function

Public Function PT_Num(ByVal v As Variant, ByVal dflt As Double) As Double
    PT_Num = dflt
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim(v) = "" Then Exit Function
    End If
    If IsNumeric(v) Then PT_Num = CDbl(v)
End Function

Public Function PT_Str(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then Exit Function
    PT_Str = Trim(CStr(v))
End Function

' ---------------------------------------------------------------- keys
' Matching key of a model name: last path segment, lower case, no instance suffix
' (-1, <1>, #2), no file extension (.STEP), letters and digits only.
Public Function NormKey(ByVal s As String) As String
    If reInst Is Nothing Then
        Set reInst = CreateObject("VBScript.RegExp"): reInst.Global = True: reInst.Pattern = "(-\d+|<\d+>|\s#\d+)$"
        Set reDup = CreateObject("VBScript.RegExp"): reDup.Global = True: reDup.Pattern = "_\d+(?=\.(step|stp)(?![a-z0-9]))"
        Set reExt = CreateObject("VBScript.RegExp"): reExt.Global = True: reExt.Pattern = "\.(step|stp|sldprt|sldasm|x_t|igs|iges)$"
        Set reJunk = CreateObject("VBScript.RegExp"): reJunk.Global = True: reJunk.Pattern = "[^a-z0-9]"
    End If
    s = Mid(s, InStrRev(s, "/") + 1)
    s = reDup.Replace(LCase(Trim(s)), "")         ' SolidWorks adds _1 to a STEP name that is already open in the session
    s = reInst.Replace(s, "")
    s = reExt.Replace(s, "")
    NormKey = reJunk.Replace(s, "")
End Function

' Identity key of a body: NormKey plus "_<n>" when the instance number is above 1,
' so two copies of the same part (for example two identical connectors) get different keys.
Public Function PT_Key(ByVal s As String) As String
    Dim seg As String, inst As Long, mt As Object, k As Long
    seg = Trim(Mid(s, InStrRev(s, "/") + 1))
    If reTail Is Nothing Then
        Set reTail = CreateObject("VBScript.RegExp"): reTail.Pattern = "(?:-(\d+)|<(\d+)>|\s#(\d+))$"
    End If
    If reTail.Test(seg) Then
        Set mt = reTail.Execute(seg)
        For k = 0 To 2
            If mt(0).SubMatches(k) <> "" Then inst = CLng(mt(0).SubMatches(k))
        Next
    End If
    PT_Key = NormKey(s)
    If PT_Key = "" Then PT_Key = "body"
    If inst > 1 Then PT_Key = PT_Key & "_" & inst
End Function

' Returns k, or k plus "_d2", "_d3" ... if k is already in the dictionary. Adds the result.
Public Function PT_UniqueKey(ByVal used As Object, ByVal k As String) As String
    Dim i As Long, s As String
    s = k
    If used.Exists(s) Then
        i = 2
        Do While used.Exists(k & "_d" & i): i = i + 1: Loop
        s = k & "_d" & i
    End If
    used.Add s, 1
    PT_UniqueKey = s
End Function

' ---------------------------------------------------------------- change log
Public Sub PT_LogReset()
    Set logRows = New Collection
End Sub

Public Sub PT_LogAdd(ByVal macroName As String, ByVal key As String, ByVal nameS As String, ByVal refDes As String, _
                     ByVal change As String, ByVal field As String, ByVal oldV As Variant, ByVal newV As Variant)
    Dim dl As Variant
    If logRows Is Nothing Then Set logRows = New Collection
    dl = ""
    If PT_IsNum(oldV) And PT_IsNum(newV) Then dl = Round(CDbl(newV) - CDbl(oldV), 4)
    logRows.Add Array(Now, macroName, key, nameS, refDes, change, field, oldV, newV, dl, "")
End Sub

Public Function PT_LogCount() As Long
    If Not logRows Is Nothing Then PT_LogCount = logRows.Count
End Function

' Writes the collected rows to the top of the Change log tab (newest on top). Returns the count.
Public Function PT_LogFlush(ByVal wb As Object) As Long
    Dim n As Long, ws As Object, hdr As Object, outv() As Variant, i As Long, k As Long, rw As Variant
    Dim r0 As Long, c0 As Long
    If logRows Is Nothing Then Exit Function
    n = logRows.Count
    If n = 0 Then Exit Function
    Set hdr = PT_Rng(wb, "Log_Header")
    Set ws = hdr.Worksheet
    r0 = hdr.Row + 1: c0 = hdr.Column
    ReDim outv(1 To n, 1 To PT_LOG_COLS)
    For i = 1 To n
        rw = logRows(i)
        For k = 1 To PT_LOG_COLS: outv(i, k) = rw(k - 1): Next
    Next
    ws.Rows(r0 & ":" & (r0 + n - 1)).Insert -4121, 1                 ' shift down, format from below
    ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + n - 1, c0 + PT_LOG_COLS - 1)).Value = outv
    ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + n - 1, c0)).NumberFormat = "yyyy-mm-dd hh:mm"
    With ws.Range(ws.Cells(r0, c0 + PT_LOG_COLS - 1), ws.Cells(r0 + n - 1, c0 + PT_LOG_COLS - 1)).Validation
        .Delete
        .Add 3, 1, 1, "=Lst_YesNo"
    End With
    Set logRows = New Collection
    PT_LogFlush = n
End Function

' Build macros call this first. True = refuse (the user has been told which rows to check).
Public Function PT_RefuseIfUnchecked(ByVal wb As Object, ByVal macroName As String) As Boolean
    Dim n As Long, ws As Object, hdr As Object, r As Long, lastR As Long, s As String, shown As Long
    Dim cChange As Long, cChecked As Long
    wb.Application.Calculate
    n = PT_NameNum(wb, "Log_Unchecked", 0)
    If n <= 0 Then Exit Function
    Set hdr = PT_Rng(wb, "Log_Header")
    Set ws = hdr.Worksheet
    cChange = hdr.Column + 5: cChecked = hdr.Column + 10
    lastR = ws.Cells(ws.Rows.Count, cChange).End(-4162).Row
    For r = hdr.Row + 1 To lastR
        If PT_Str(ws.Cells(r, cChange).Value) <> "" And UCase(PT_Str(ws.Cells(r, cChecked).Value)) <> "YES" Then
            s = s & "  row " & r & ": " & PT_Str(ws.Cells(r, cChange).Value) & " " & _
                PT_Str(ws.Cells(r, cChange + 1).Value) & "  " & PT_Str(ws.Cells(r, hdr.Column + 3).Value) & vbCrLf
            shown = shown + 1
            If shown >= 8 Then s = s & "  ..." & vbCrLf: Exit For
        End If
    Next
    MsgBox macroName & " did not run. Nothing was changed." & vbCrLf & vbCrLf & _
           n & " rows on the Change log tab are not checked:" & vbCrLf & s & vbCrLf & _
           "Look at each one, type Yes in its Checked cell, save, and run again.", vbExclamation, macroName
    PT_RefuseIfUnchecked = True
End Function

' ---------------------------------------------------------------- files
' Output folder from Settings (blank = CAD folder next to the workbook), created if missing, ends with "\".
Public Function PT_OutputFolder(ByVal wb As Object) As String
    Dim s As String
    s = PT_NameStr(wb, "Set_OutputFolder", "")
    If s = "" Then s = wb.Path & "\CAD"
    If Right(s, 1) = "\" Then s = Left(s, Len(s) - 1)
    PT_MakeDirs s
    PT_OutputFolder = s & "\"
End Function

' File prefix from Settings. Blank = the workbook file name, e.g. "RaspberryPi4.xlsx" -> "RaspberryPi4_".
Public Function PT_Prefix(ByVal wb As Object) As String
    Dim s As String
    s = PT_NameStr(wb, "Set_Prefix", "")
    If s = "" Then
        s = wb.Name
        If InStrRev(s, ".") > 1 Then s = Left(s, InStrRev(s, ".") - 1)
        s = PT_SafeFile(s) & "_"
    End If
    PT_Prefix = s
End Function

Public Sub PT_MakeDirs(ByVal p As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If p = "" Then Exit Sub
    If fso.FolderExists(p) Then Exit Sub
    PT_MakeDirs fso.GetParentFolderName(p)
    fso.CreateFolder p
End Sub

' =====================================================================================
' SolidWorks helpers shared by the build macros. Lengths in the arguments are mm
' (board frame); conversion to metres happens here.
' =====================================================================================


' ---------------------------------------------------------------- properties
Public Function PT_GetProp(ByVal md As Object, ByVal fld As String) As String
    Dim valOut As String, resOut As String, wasRes As Boolean, linkProp As Boolean
    On Error Resume Next
    If md Is Nothing Then Exit Function
    md.Extension.CustomPropertyManager("").Get6 fld, False, valOut, resOut, wasRes, linkProp
    If resOut <> "" Then PT_GetProp = resOut Else PT_GetProp = valOut
End Function

Public Function PT_PropNum(ByVal md As Object, ByVal fld As String) As Double
    Dim s As String
    s = PT_GetProp(md, fld)
    If s <> "" And IsNumeric(s) Then PT_PropNum = CDbl(s)
End Function

Public Sub PT_SetProp(ByVal md As Object, ByVal fld As String, ByVal valS As String)
    md.Extension.CustomPropertyManager("").Add3 fld, 30, valS, 2       ' text, replace value
End Sub

' mm -> whole micrometres as text (no decimal separator, so the Windows language does not matter)
Public Function PT_Um(ByVal mm As Double) As String
    PT_Um = CStr(CLng(Round(mm * 1000#)))
End Function

Public Function PT_SafeFile(ByVal s As String) As String
    Dim i As Long, ch As String, outS As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch Like "[A-Za-z0-9_-]" Then outS = outS & ch Else outS = outS & "_"
    Next
    If Len(outS) > 60 Then outS = Left(outS, 60)
    PT_SafeFile = outS
End Function

Public Function PT_BoxText(ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double, _
                           ByVal lx As Double, ByVal wy As Double, ByVal hz As Double) As String
    PT_BoxText = "at " & Round(x0, 3) & ", " & Round(y0, 3) & ", " & Round(z0, 3) & _
                 " size " & Round(lx, 3) & " x " & Round(wy, 3) & " x " & Round(hz, 3)
End Function

' ---------------------------------------------------------------- documents
Public Sub PT_ActivateDoc(ByVal md As Object)
    Dim errs As Long
    swApp.ActivateDoc3 md.GetTitle, False, 0, errs
End Sub

' Thermal assembly path from Settings (blank = <prefix>Thermal.sldasm in the output folder).
Public Function PT_AsmPath(ByVal wb As Object) As String
    Dim s As String
    s = PT_NameStr(wb, "Set_AsmName", "")
    If s = "" Then s = PT_SafeFile(PT_Prefix(wb)) & "Thermal.sldasm"
    If InStr(s, "\") = 0 Then s = PT_OutputFolder(wb) & s
    If LCase(Right(s, 7)) <> ".sldasm" Then s = s & ".sldasm"
    PT_AsmPath = s
End Function

' The assembly if it is open or on disk, else Nothing.
Public Function PT_OpenAsm(ByVal asmPath As String) As Object
    Dim errs As Long, warns As Long, md As Object
    Set md = swApp.GetOpenDocumentByName(asmPath)
    If md Is Nothing Then
        If Dir(asmPath) <> "" Then Set md = swApp.OpenDoc6(asmPath, PT_swDocASSEMBLY, PT_swOpenSilent, "", errs, warns)
    End If
    Set PT_OpenAsm = md
End Function

' Template path from Settings, else the SolidWorks default; "" if the file is not there.
Public Function PT_Template(ByVal wb As Object, ByVal nameS As String, ByVal swPref As Long) As String
    Dim s As String
    s = PT_NameStr(wb, nameS, "")
    If s = "" Then s = swApp.GetUserPreferenceStringValue(swPref)
    If s <> "" Then
        If Dir(s) = "" Then s = ""
    End If
    PT_Template = s
End Function

' Top-level components whose part carries Role = roleS. Fills have (Key -> component) and
' wasSupp (Key -> True if it was suppressed). Suppressed ones are resolved so they can be read;
' the caller suppresses again what it does not need. Returns the last one found (for single parts).
Public Function PT_ReadTagged(ByVal asm As Object, ByVal roleS As String, ByVal have As Object, ByVal wasSupp As Object) As Object
    Dim comps As Variant, i As Long, c As Object, md As Object, keyS As String, supp As Boolean, lastC As Object
    asm.ResolveAllLightWeightComponents False
    comps = asm.GetComponents(True)
    If IsEmpty(comps) Then Exit Function
    For i = 0 To UBound(comps)
        Set c = comps(i)
        supp = (c.GetSuppression2 = PT_swSuppressed)
        If supp Then c.SetSuppression2 PT_swFullyResolved
        Set md = c.GetModelDoc2
        If PT_GetProp(md, "Role") = roleS Then
            keyS = PT_GetProp(md, "Key")
            If keyS = "" Then keyS = "(no key) " & c.Name2
            If Not have.Exists(keyS) Then
                have.Add keyS, c
                wasSupp.Add keyS, supp
                Set lastC = c
            End If
        ElseIf supp Then
            c.SetSuppression2 PT_swSuppressed
        End If
    Next
    Set PT_ReadTagged = lastC
End Function

' Suppresses tagged components whose key is not in want. Logs the ones that were active. Returns that count.
Public Function PT_SuppressUnwanted(ByVal have As Object, ByVal wasSupp As Object, ByVal want As Object, _
                                    ByVal macroName As String, ByVal doLog As Boolean) As Long
    Dim it As Variant, c As Object, md As Object, cnt As Long
    For Each it In have.Keys
        If Not want.Exists(it) Then
            Set c = have(it)
            If Not wasSupp(it) Then
                Set md = c.GetModelDoc2
                If doLog Then PT_LogAdd macroName, CStr(it), PT_GetProp(md, "Name"), PT_GetProp(md, "RefDes"), _
                                        "removed", "Suppressed in the assembly", "", ""
                cnt = cnt + 1
            End If
            c.SetSuppression2 PT_swSuppressed
        End If
    Next
    PT_SuppressUnwanted = cnt
End Function

' ---------------------------------------------------------------- features and faces
Public Function PT_FindFeature(ByVal md As Object, ByVal featName As String) As Object
    Dim f As Object
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        If f.Name = featName Then Set PT_FindFeature = f: Exit Function
        Set f = f.GetNextFeature
    Loop
End Function

' nth reference plane of the part (1 = Front, the XY plane in the standard templates)
Public Function PT_RefPlane(ByVal md As Object, ByVal nth As Long) As Object
    Dim f As Object, k As Long
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then
            k = k + 1
            If k = nth Then Set PT_RefPlane = f: Exit Function
        End If
        Set f = f.GetNextFeature
    Loop
End Function

Public Function PT_CountFeatures(ByVal md As Object, ByVal typeName2 As String) As Long
    Dim f As Object
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = typeName2 Then PT_CountFeatures = PT_CountFeatures + 1
        Set f = f.GetNextFeature
    Loop
End Function

Public Function PT_LastFeatureOfType(ByVal md As Object, ByVal typeName2 As String) As Object
    Dim f As Object, lastF As Object
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = typeName2 Then Set lastF = f
        Set f = f.GetNextFeature
    Loop
    Set PT_LastFeatureOfType = lastF
End Function

Public Sub PT_DeleteFeature(ByVal md As Object, ByVal feat As Object)
    md.ClearSelection2 True
    feat.Select2 False, 0
    md.Extension.DeleteSelection2 PT_swDeleteAbsorbed          ' also deletes its sketch
    md.ClearSelection2 True
End Sub

' model point (mm) -> sketch coordinates (m) of the active sketch
Public Function PT_ToSketch(ByVal xf As Object, ByVal x As Double, ByVal y As Double, ByVal z As Double) As Variant
    Dim mu As SldWorks.MathUtility, d(2) As Double, pt As SldWorks.MathPoint, xfT As SldWorks.MathTransform
    d(0) = x / 1000#: d(1) = y / 1000#: d(2) = z / 1000#
    Set mu = swApp.GetMathUtility
    Set xfT = xf
    Set pt = mu.CreatePoint(d)
    Set pt = pt.MultiplyTransform(xfT)
    PT_ToSketch = pt.ArrayData
End Function

' Rectangle from (x0,y0) to (x1,y1), board frame mm, on plane pl whose height is zPlane (mm).
' Uses the sketch's own transform, so the plane orientation does not matter. Returns the sketch feature.
Public Function PT_SketchRect(ByVal md As Object, ByVal pl As Object, ByVal x0 As Double, ByVal y0 As Double, _
                              ByVal x1 As Double, ByVal y1 As Double, ByVal zPlane As Double) As Object
    Dim sk As Object, xf As Object, a As Variant, b As Variant, skF As Object, tryN As Long, seg As Variant, nBefore As Long
    nBefore = PT_CountFeatures(md, "ProfileFeature")
    For tryN = 0 To 1
        md.ClearSelection2 True
        pl.Select2 False, 0
        md.SketchManager.InsertSketch True
        Set sk = md.SketchManager.ActiveSketch
        If sk Is Nothing Then Err.Raise vbObjectError + 20, , "could not start a sketch on " & pl.Name
        Set xf = sk.ModelToSketchTransform
        a = PT_ToSketch(xf, x0, y0, zPlane)
        b = PT_ToSketch(xf, x1, y1, zPlane)
        md.SketchManager.AddToDB = (tryN = 0)          ' second try: normal sketch mode
        seg = md.SketchManager.CreateCornerRectangle(a(0), a(1), 0, b(0), b(1), 0)
        md.SketchManager.AddToDB = False
        md.SketchManager.InsertSketch True
        md.ClearSelection2 True
        If PT_CountFeatures(md, "ProfileFeature") > nBefore Then Set skF = PT_LastFeatureOfType(md, "ProfileFeature"): Exit For
    Next
    If skF Is Nothing Then Err.Raise vbObjectError + 21, , "the sketch was not made (rectangle " & _
        Format(Abs(x1 - x0), "0.000") & " x " & Format(Abs(y1 - y0), "0.000") & " mm)"
    Set PT_SketchRect = skF
End Function

' Blind boss from a sketch. down = True extrudes the other way. Returns the feature.
Public Function PT_Boss(ByVal md As Object, ByVal sk As Object, ByVal depthMm As Double, ByVal down As Boolean) As Object
    Dim feat As Object
    md.ClearSelection2 True
    sk.Select2 False, 0
    Set feat = md.FeatureManager.FeatureExtrusion2(True, False, down, 0, 0, depthMm / 1000#, 0, False, False, False, False, _
                                                   0, 0, False, False, False, False, True, True, True, 0, 0, False)
    md.ClearSelection2 True
    If feat Is Nothing Then Err.Raise vbObjectError + 22, , "the extrusion failed"
    Set PT_Boss = feat
End Function

' Cut through everything, both directions. Returns the feature (Nothing if SolidWorks refused, e.g. no overlap).
Public Function PT_CutThrough(ByVal md As Object, ByVal sk As Object) As Object
    Dim feat As Object
    md.ClearSelection2 True
    sk.Select2 False, 0
    Set feat = md.FeatureManager.FeatureCut4(False, False, False, 1, 1, 0.01, 0.01, False, False, False, False, _
                                             0, 0, False, False, False, False, False, True, True, False, False, False, _
                                             0, 0, False, False)
    md.ClearSelection2 True
    Set PT_CutThrough = feat
End Function

Public Sub PT_FlipExtrude(ByVal md As Object, ByVal feat As Object)
    Dim def As Object
    Set def = feat.GetDefinition
    def.AccessSelections md, Nothing
    def.ReverseDirection = Not def.ReverseDirection
    feat.ModifyDefinition def, md, Nothing
End Sub

Public Sub PT_SetExtrudeDepth(ByVal md As Object, ByVal feat As Object, ByVal mm As Double)
    Dim def As Object
    Set def = feat.GetDefinition
    def.AccessSelections md, Nothing
    def.SetDepth True, mm / 1000#
    feat.ModifyDefinition def, md, Nothing
End Sub

Public Function PT_ExtrudeDepth(ByVal feat As Object) As Double
    Dim def As Object
    Set def = feat.GetDefinition
    PT_ExtrudeDepth = def.GetDepth(True) * 1000#
End Function

Public Sub PT_SetPlaneOffset(ByVal md As Object, ByVal feat As Object, ByVal mm As Double, ByVal flipIt As Boolean)
    Dim def As Object
    Set def = feat.GetDefinition
    def.AccessSelections md, Nothing
    def.Distance = mm / 1000#
    If flipIt Then def.ReverseDirection = Not def.ReverseDirection
    feat.ModifyDefinition def, md, Nothing
End Sub

Public Function PT_PlaneOffset(ByVal feat As Object) As Double
    Dim def As Object
    Set def = feat.GetDefinition
    PT_PlaneOffset = def.Distance * 1000#
End Function

' Z range (mm) of the solid bodies of a part: Array(zmin, zmax), or Empty
Public Function PT_BodyZRange(ByVal md As Object) As Variant
    Dim bodies As Variant, i As Long, bb As Variant, zmin As Double, zmax As Double
    bodies = md.GetBodies2(PT_swSolidBody, True)
    If IsEmpty(bodies) Then Exit Function
    zmin = 1E+30: zmax = -1E+30
    For i = 0 To UBound(bodies)
        bb = bodies(i).GetBodyBox
        If bb(2) * 1000# < zmin Then zmin = bb(2) * 1000#
        If bb(5) * 1000# > zmax Then zmax = bb(5) * 1000#
    Next
    PT_BodyZRange = Array(zmin, zmax)
End Function

' Lowest Z (mm) over the faces a feature made, or 1E+30 if it made none
Public Function PT_FeatureZMin(ByVal feat As Object) As Double
    Dim faces As Variant, i As Long, bb As Variant, zmin As Double
    zmin = 1E+30
    faces = feat.GetFaces
    If Not IsEmpty(faces) Then
        For i = 0 To UBound(faces)
            bb = faces(i).GetBox
            If bb(2) * 1000# < zmin Then zmin = bb(2) * 1000#
        Next
    End If
    PT_FeatureZMin = zmin
End Function

Public Function PT_AllFaces(ByVal md As Object) As Collection
    Dim bodies As Variant, faces As Variant, i As Long, j As Long, res As Collection
    Set res = New Collection
    bodies = md.GetBodies2(PT_swSolidBody, True)
    If Not IsEmpty(bodies) Then
        For i = 0 To UBound(bodies)
            faces = bodies(i).GetFaces
            If Not IsEmpty(faces) Then
                For j = 0 To UBound(faces): res.Add faces(j): Next
            End If
        Next
    End If
    Set PT_AllFaces = res
End Function

' Names the largest flat face facing up (or down) that lies at height zmm and whose centre is inside
' the rectangle x0..x1, y0..y1 (mm). Any other face holding that name loses it. True if named.
Public Function PT_NameFlatFace(ByVal md As Object, ByVal up As Boolean, ByVal zmm As Double, ByVal nameS As String, _
                                ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double) As Boolean
    Dim faces As Collection, fc As Object, nv As Variant, bb As Variant, best As Object, bestA As Double, a As Double
    Dim cx As Double, cy As Double
    On Error Resume Next
    Set faces = PT_AllFaces(md)
    For Each fc In faces
        If md.GetEntityName(fc) = nameS Then md.DeleteEntityName fc
    Next
    For Each fc In faces
        nv = Empty
        nv = fc.Normal
        If Not IsEmpty(nv) Then
            If (up And nv(2) > 0.99) Or ((Not up) And nv(2) < -0.99) Then
                bb = fc.GetBox
                cx = (bb(0) + bb(3)) * 500#: cy = (bb(1) + bb(4)) * 500#
                If Abs(bb(2) * 1000# - zmm) < 0.001 And Abs(bb(5) * 1000# - zmm) < 0.001 And _
                   cx >= x0 And cx <= x1 And cy >= y0 And cy <= y1 Then
                    a = fc.GetArea
                    If a > bestA Then bestA = a: Set best = fc
                End If
            End If
        End If
    Next
    If best Is Nothing Then Exit Function
    If md.GetEntityName(best) <> "" Then md.DeleteEntityName best
    PT_NameFlatFace = md.SetEntityName(best, nameS)
End Function

' ---------------------------------------------------------------- footprint shape (rotated parts)
' A part can be a rotated rectangle instead of an axis-aligned box. BuildBoard sets PT_ShapeRot / PT_ShapeL /
' PT_ShapeW before PT_AddBoxPart or PT_CheckBoxPart and clears them after (PT_ShapeClear). The L and W given to
' those functions are then the ENVELOPE (axis-aligned size) of the rotated rectangle, so every box comparison
' in the other macros keeps working. The shape is stored in the part's custom properties (Rot_mdeg, OwnL_um,
' OwnW_um) and PT_BuildBox / PT_BuildLayeredBox read it back, so a rebuild keeps the rotation.
' (the three variables are declared at the top of this module)
Public Sub PT_ShapeClear()
    PT_ShapeRot = 0: PT_ShapeL = 0: PT_ShapeW = 0
End Sub

' Axis-aligned size of a rectangle ownL x ownW turned by rotDeg.
Public Sub PT_ShapeEnvelope(ByVal rotDeg As Double, ByVal ownL As Double, ByVal ownW As Double, _
                            ByRef envL As Double, ByRef envW As Double)
    Dim ca As Double, sa As Double
    ca = Abs(Cos(rotDeg * PT_PI / 180#)): sa = Abs(Sin(rotDeg * PT_PI / 180#))
    envL = ownL * ca + ownW * sa
    envW = ownL * sa + ownW * ca
End Sub

' True when the shape is a real rotated rectangle whose envelope is the box lx x wy (within 0.02 mm).
Public Function PT_ShapeIsRotated(ByVal rotDeg As Double, ByVal ownL As Double, ByVal ownW As Double, _
                                  ByVal lx As Double, ByVal wy As Double) As Boolean
    Dim e1 As Double, e2 As Double
    If Abs(rotDeg) < 0.05 Or ownL <= 0 Or ownW <= 0 Then Exit Function
    PT_ShapeEnvelope rotDeg, ownL, ownW, e1, e2
    PT_ShapeIsRotated = (Abs(e1 - lx) <= 0.02 And Abs(e2 - wy) <= 0.02)
End Function

Public Sub PT_SetShapeProps(ByVal md As Object)
    PT_SetProp md, "Rot_mdeg", CStr(CLng(Round(PT_ShapeRot * 1000#)))
    PT_SetProp md, "OwnL_um", PT_Um(PT_ShapeL)
    PT_SetProp md, "OwnW_um", PT_Um(PT_ShapeW)
End Sub

' Shape stored in the part: True when it is a rotated rectangle that fits the box lx x wy.
Public Function PT_ShapeOfPart(ByVal md As Object, ByVal lx As Double, ByVal wy As Double, _
                               ByRef rotDeg As Double, ByRef ownL As Double, ByRef ownW As Double) As Boolean
    rotDeg = PT_PropNum(md, "Rot_mdeg") / 1000#
    ownL = PT_PropNum(md, "OwnL_um") / 1000#
    ownW = PT_PropNum(md, "OwnW_um") / 1000#
    PT_ShapeOfPart = PT_ShapeIsRotated(rotDeg, ownL, ownW, lx, wy)
End Function

' Footprint sketch of a box part, from (0,0) to (lx,wy) on plane pl: a rectangle, or the rotated rectangle stored in
' the part when there is one (its corners stay inside 0..lx and 0..wy, so the part origin is still the box corner).
Public Function PT_SketchFootprint(ByVal md As Object, ByVal pl As Object, ByVal lx As Double, ByVal wy As Double, _
                                   ByVal zPlane As Double) As Object
    Dim rotDeg As Double, ownL As Double, ownW As Double, sk As Object, xf As Object, skF As Object
    Dim ca As Double, sa As Double, k As Long, j As Long, sx(3) As Double, sy(3) As Double
    Dim dx As Double, dy As Double, px As Double, py As Double, a As Variant, nBefore As Long
    nBefore = PT_CountFeatures(md, "ProfileFeature")
    If Not PT_ShapeOfPart(md, lx, wy, rotDeg, ownL, ownW) Then
        Set PT_SketchFootprint = PT_SketchRect(md, pl, 0, 0, lx, wy, zPlane)
        Exit Function
    End If
    ca = Cos(rotDeg * PT_PI / 180#): sa = Sin(rotDeg * PT_PI / 180#)
    md.ClearSelection2 True
    pl.Select2 False, 0
    md.SketchManager.InsertSketch True
    Set sk = md.SketchManager.ActiveSketch
    If sk Is Nothing Then Err.Raise vbObjectError + 23, , "could not start a sketch on " & pl.Name
    Set xf = sk.ModelToSketchTransform
    For k = 0 To 3
        dx = ownL / 2#: dy = ownW / 2#
        If k = 0 Or k = 3 Then dx = -dx
        If k = 0 Or k = 1 Then dy = -dy
        px = lx / 2# + dx * ca - dy * sa
        py = wy / 2# + dx * sa + dy * ca
        a = PT_ToSketch(xf, px, py, zPlane)
        sx(k) = a(0): sy(k) = a(1)
    Next
    md.SketchManager.AddToDB = True
    For k = 0 To 3
        j = (k + 1) Mod 4
        md.SketchManager.CreateLine sx(k), sy(k), 0, sx(j), sy(j), 0
    Next
    md.SketchManager.AddToDB = False
    md.SketchManager.InsertSketch True
    md.ClearSelection2 True
    If PT_CountFeatures(md, "ProfileFeature") > nBefore Then Set skF = PT_LastFeatureOfType(md, "ProfileFeature")
    If skF Is Nothing Then Err.Raise vbObjectError + 24, , "the footprint sketch was not made (" & Format(lx, "0.000") & " x " & Format(wy, "0.000") & " mm, turned " & Format(rotDeg, "0.0") & " deg)"
    Set PT_SketchFootprint = skF
End Function

' ---------------------------------------------------------------- box parts
' Box from (0,0,0) to (L,W,H) in the part. An older box made by this code is deleted first, so this also resizes.
' When the part carries a rotated shape (PT_SetShapeProps) the footprint is that rotated rectangle.
Public Sub PT_BuildBox(ByVal md As Object, ByVal lx As Double, ByVal wy As Double, ByVal hz As Double, _
                       ByVal topName As String, ByVal bottomName As String)
    Dim f As Object, pl As Object, sk As Object, feat As Object, zr As Variant
    PT_DropLayers md
    Set f = PT_FindFeature(md, "PT_Box")
    If Not f Is Nothing Then PT_DeleteFeature md, f
    Set pl = PT_RefPlane(md, 1)
    If pl Is Nothing Then Err.Raise vbObjectError + 10, , "the part template has no reference plane"
    Set sk = PT_SketchFootprint(md, pl, lx, wy, 0)
    sk.Name = "PT_BoxSketch"
    Set feat = PT_Boss(md, sk, hz, False)
    feat.Name = "PT_Box"
    zr = PT_BodyZRange(md)
    If Not IsEmpty(zr) Then
        If zr(0) < -hz / 2 Then PT_FlipExtrude md, feat               ' went down: turn it up
    End If
    PT_NameBoxFaces md, hz, topName, bottomName
End Sub

' Names the top and bottom face of a box part. Raises an error when either face cannot be found or named,
' so the caller reports the part instead of building it without the faces SetupSim needs.
Private Sub PT_NameBoxFaces(ByVal md As Object, ByVal hz As Double, ByVal topName As String, ByVal bottomName As String)
    If Not PT_NameFlatFace(md, True, hz, topName, -1E+30, -1E+30, 1E+30, 1E+30) Then
        Err.Raise vbObjectError + 13, , "could not name the top face " & topName
    End If
    If Not PT_NameFlatFace(md, False, 0, bottomName, -1E+30, -1E+30, 1E+30, 1E+30) Then
        Err.Raise vbObjectError + 14, , "could not name the bottom face " & bottomName
    End If
End Sub

' Removes the junction core and top layer made by PT_BuildLayeredBox (M8) and clears their tag.
Public Sub PT_DropLayers(ByVal md As Object)
    Dim f As Object, nm As Variant, cur As String
    For Each nm In Array("PT_LayerTop", "PT_LayerCore")
        Set f = PT_FindFeature(md, CStr(nm))
        If Not f Is Nothing Then PT_DeleteFeature md, f
    Next
    cur = PT_GetProp(md, "Layers_um")
    If cur <> "" And cur <> "-" Then PT_SetProp md, "Layers_um", "-"      ' "-" = no layers
End Sub

' Box made of three stacked bodies for the two-resistor chip model (M8): bottom layer (tBot, toward the board),
' junction core, top layer (tTop). Same envelope as PT_BuildBox. Feature names: PT_Box (bottom layer),
' PT_LayerCore, PT_LayerTop. Tag Layers_um = "<tBot>|<tTop>" in micrometres. False on failure (PT_LastErr).
Public Function PT_BuildLayeredBox(ByVal md As Object, ByVal lx As Double, ByVal wy As Double, ByVal hz As Double, _
                                   ByVal tBot As Double, ByVal tTop As Double, _
                                   ByVal topName As String, ByVal bottomName As String) As Boolean
    Dim f As Object, pl As Object, sk As Object, feat As Object, zr As Variant, down As Boolean
    PT_LastErr = ""
    On Error GoTo failed
    If tBot <= 0 Or tTop <= 0 Or tBot + tTop >= hz Then Err.Raise vbObjectError + 11, , "layer thicknesses do not fit the chip height"
    PT_DropLayers md
    Set f = PT_FindFeature(md, "PT_Box")
    If Not f Is Nothing Then PT_DeleteFeature md, f
    Set pl = PT_RefPlane(md, 1)
    If pl Is Nothing Then Err.Raise vbObjectError + 10, , "the part template has no reference plane"
    Set sk = PT_SketchFootprint(md, pl, lx, wy, 0)
    sk.Name = "PT_BoxSketch"
    Set feat = PT_Boss(md, sk, tBot, False)
    feat.Name = "PT_Box"
    zr = PT_BodyZRange(md)
    If Not IsEmpty(zr) Then
        If zr(0) < -tBot / 2 Then PT_FlipExtrude md, feat: down = True
    End If
    PT_BossLayer md, pl, lx, wy, tBot, hz - tBot - tTop, "PT_LayerCore", down
    PT_BossLayer md, pl, lx, wy, hz - tTop, tTop, "PT_LayerTop", down
    PT_NameBoxFaces md, hz, topName, bottomName
    PT_SetProp md, "Layers_um", PT_Um(tBot) & "|" & PT_Um(tTop)
    PT_BuildLayeredBox = True
    Exit Function
failed:
    PT_LastErr = Err.Description
End Function

' Separate body (not merged) from height z0 to z0 + depth (mm), made by a boss with a start offset.
' Tries the extrude direction that is known to work first, the other one when the result lands in the wrong place.
Private Sub PT_BossLayer(ByVal md As Object, ByVal pl As Object, ByVal lx As Double, ByVal wy As Double, _
                         ByVal z0 As Double, ByVal depth As Double, ByVal featName As String, ByRef down As Boolean)
    Dim attempt As Long, sk As Object, feat As Object, d As Boolean, zmn As Double
    For attempt = 0 To 1
        If attempt = 0 Then d = down Else d = Not down
        Set sk = PT_SketchFootprint(md, pl, lx, wy, 0)
        md.ClearSelection2 True
        sk.Select2 False, 0
        Set feat = md.FeatureManager.FeatureExtrusion2(True, False, d, 0, 0, depth / 1000#, 0, False, False, False, False, _
                                                       0, 0, False, False, False, False, False, True, True, 3, z0 / 1000#, False)
        md.ClearSelection2 True
        If Not feat Is Nothing Then
            zmn = PT_FeatureZMin(feat)
            If Abs(zmn - z0) < 0.001 Then
                feat.Name = featName
                sk.Name = featName & "Sketch"
                down = d
                Exit Sub
            End If
            PT_DeleteFeature md, feat
        Else
            PT_DeleteFeature md, sk
        End If
    Next
    Err.Raise vbObjectError + 12, , "could not make the layer " & featName
End Sub

' Bodies of a component that carry the heat: the middle one of three stacked layers (M8), else all bodies.
Public Function PT_CoreBodies(ByVal c As Object) As Variant
    Dim b As Variant, i As Long, j As Long, zs(2) As Double, mid3 As Long, bb As Variant, one(0) As Variant
    b = PT_CompBodies(c)
    PT_CoreBodies = b
    If IsEmpty(b) Then Exit Function
    If UBound(b) <> 2 Then Exit Function
    For i = 0 To 2
        bb = b(i).GetBodyBox
        zs(i) = bb(2)
    Next
    For i = 0 To 2
        j = 0
        If zs(i) > zs(0) Then j = j + 1
        If zs(i) > zs(1) Then j = j + 1
        If zs(i) > zs(2) Then j = j + 1
        If j = 1 Then mid3 = i
    Next
    Set one(0) = b(mid3)
    PT_CoreBodies = one
End Function

Public Sub PT_TagBox(ByVal md As Object, ByVal roleS As String, ByVal keyS As String, ByVal nameS As String, ByVal refD As String, _
                     ByVal srcRow As Long, ByVal lx As Double, ByVal wy As Double, ByVal hz As Double, ByVal macroName As String)
    PT_SetProp md, "Role", roleS
    PT_SetProp md, "Key", keyS
    PT_SetProp md, "Name", nameS
    PT_SetProp md, "RefDes", refD
    PT_SetProp md, "SourceRow", CStr(srcRow)
    PT_SetProp md, "L_um", PT_Um(lx)
    PT_SetProp md, "W_um", PT_Um(wy)
    PT_SetProp md, "H_um", PT_Um(hz)
    PT_SetProp md, "BuiltBy", macroName
End Sub

Public Sub PT_PlaceAndFix(ByVal asm As Object, ByVal c As SldWorks.Component2, ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double)
    Dim mu As SldWorks.MathUtility, xf As SldWorks.MathTransform, d(15) As Double, v As Variant
    d(0) = 1: d(4) = 1: d(8) = 1
    d(9) = x0 / 1000#: d(10) = y0 / 1000#: d(11) = z0 / 1000#
    d(12) = 1
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
    PT_FixComp asm, c
End Sub

Public Sub PT_FixComp(ByVal asm As Object, ByVal c As Object)
    asm.ClearSelection2 True
    c.Select4 False, Nothing, False
    asm.FixComponent
    asm.ClearSelection2 True
End Sub

' New box part file, inserted in the assembly, placed and fixed. False on failure (reason in PT_LastErr).
Public Function PT_AddBoxPart(ByVal asm As Object, ByVal partTmpl As String, ByVal fpath As String, _
                              ByVal roleS As String, ByVal keyS As String, ByVal nameS As String, ByVal refD As String, _
                              ByVal srcRow As Long, ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double, _
                              ByVal lx As Double, ByVal wy As Double, ByVal hz As Double, _
                              ByVal topName As String, ByVal bottomName As String, ByVal macroName As String) As Boolean
    Dim md As Object, errs As Long, warns As Long, c As Object, ttl As String
    PT_LastErr = ""
    On Error GoTo failed
    If Not swApp.GetOpenDocumentByName(fpath) Is Nothing Then Err.Raise vbObjectError + 1, , "a file with this name is already open: " & fpath
    Set md = swApp.NewDocument(partTmpl, 0, 0, 0)
    If md Is Nothing Then Err.Raise vbObjectError + 2, , "SolidWorks could not make a new part from " & partTmpl
    PT_SetShapeProps md
    PT_BuildBox md, lx, wy, hz, topName, bottomName
    PT_TagBox md, roleS, keyS, nameS, refD, srcRow, lx, wy, hz, macroName
    If Not md.Extension.SaveAs(fpath, 0, PT_swSaveSilent, Nothing, errs, warns) Then Err.Raise vbObjectError + 3, , "could not save " & fpath & " (error " & errs & ")"
    ttl = md.GetTitle
    PT_ActivateDoc asm
    Set c = asm.AddComponent5(fpath, 0, "", False, "", 0, 0, 0)
    If c Is Nothing Then Err.Raise vbObjectError + 4, , "could not insert " & fpath & " into the assembly"
    PT_PlaceAndFix asm, c, x0, y0, z0
    swApp.CloseDoc ttl
    PT_AddBoxPart = True
    Exit Function
failed:
    PT_LastErr = Err.Description
    Resume addCleanup
addCleanup:
    On Error Resume Next
    If Not md Is Nothing Then swApp.CloseDoc md.GetTitle           ' drop the part, saved or not, so a rerun is not blocked
    PT_ActivateDoc asm
End Function

' Box of a component in the assembly (board frame, mm): Array(xmin, ymin, zmin, L, W, H), or Empty
Public Function PT_CompBox(ByVal c As Object) As Variant
    Dim v As Variant, k As Long, outv(5) As Double
    On Error Resume Next
    v = c.GetBox(False, False)
    On Error GoTo 0
    If IsEmpty(v) Then Exit Function
    For k = 0 To 2
        outv(k) = v(k) * 1000#
        outv(3 + k) = (v(3 + k) - v(k)) * 1000#
    Next
    PT_CompBox = outv
End Function

' Existing box part: compares its real box in the assembly with the wanted box, rebuilds and moves it
' when needed, refreshes its tags, logs what changed.
' Returns 0 no change, +1 resized, +2 moved (so 3 = both), -1 failed (reason in PT_LastErr).
Public Function PT_CheckBoxPart(ByVal asm As Object, ByVal c As Object, ByVal roleS As String, ByVal keyS As String, _
                                ByVal nameS As String, ByVal refD As String, ByVal srcRow As Long, _
                                ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double, _
                                ByVal lx As Double, ByVal wy As Double, ByVal hz As Double, _
                                ByVal topName As String, ByVal bottomName As String, _
                                ByVal tolPos As Double, ByVal tolSize As Double, ByVal macroName As String, ByVal doLog As Boolean) As Long
    Dim md As Object, bx As Variant, oldS(2) As Double, newS(2) As Double, oldP(2) As Double, newP(2) As Double
    Dim k As Long, sizeChanged As Boolean, posChanged As Boolean, ad As Variant, errs As Long, warns As Long
    Dim fields As Variant, refChanged As Boolean, res As Long
    Dim shapeChanged As Boolean, oldRot As Double, oldOwnL As Double, oldOwnW As Double
    PT_LastErr = ""
    On Error GoTo failed
    fields = Array("X min", "Y min", "Z min", "L", "W", "H")
    Set md = c.GetModelDoc2
    If md Is Nothing Then Err.Raise vbObjectError + 5, , "the part file could not be loaded"
    bx = PT_CompBox(c)
    If IsEmpty(bx) Then Err.Raise vbObjectError + 6, , "the part has no geometry"
    For k = 0 To 2
        oldP(k) = bx(k): oldS(k) = bx(3 + k)
    Next
    newS(0) = lx: newS(1) = wy: newS(2) = hz
    newP(0) = x0: newP(1) = y0: newP(2) = z0
    For k = 0 To 2
        If Abs(newS(k) - oldS(k)) > tolSize Then sizeChanged = True
        If Abs(newP(k) - oldP(k)) > tolPos Then posChanged = True
    Next
    refChanged = (PT_GetProp(md, "RefDes") <> refD)
    ' rotated footprint: parts made before this feature have no shape properties = not rotated
    oldRot = PT_PropNum(md, "Rot_mdeg") / 1000#
    oldOwnL = PT_PropNum(md, "OwnL_um") / 1000#
    oldOwnW = PT_PropNum(md, "OwnW_um") / 1000#
    If Abs(oldRot - PT_ShapeRot) > 0.05 Then shapeChanged = True
    If PT_ShapeRot <> 0 Or oldRot <> 0 Then
        If Abs(oldOwnL - PT_ShapeL) > tolSize Or Abs(oldOwnW - PT_ShapeW) > tolSize Then shapeChanged = True
    End If

    If sizeChanged Or refChanged Or shapeChanged Then
        ' rebuilt box starts at the part origin, so the component goes exactly to the wanted corner
        PT_SetShapeProps md
        PT_BuildBox md, lx, wy, hz, topName, bottomName
        PT_TagBox md, roleS, keyS, nameS, refD, srcRow, lx, wy, hz, macroName
        md.Save3 PT_swSaveSilent, errs, warns
        PT_PlaceAndFix asm, c, x0, y0, z0
    Else
        If PT_GetProp(md, "Name") <> nameS Or PT_GetProp(md, "SourceRow") <> CStr(srcRow) Then
            PT_TagBox md, roleS, keyS, nameS, refD, srcRow, lx, wy, hz, macroName
            md.Save3 PT_swSaveSilent, errs, warns
        End If
        If posChanged Then
            ' shift the component by the difference, whatever is inside the part
            ad = c.Transform2.ArrayData
            PT_PlaceAndFix asm, c, ad(9) * 1000# + (x0 - oldP(0)), ad(10) * 1000# + (y0 - oldP(1)), ad(11) * 1000# + (z0 - oldP(2))
        ElseIf Not c.IsFixed Then
            PT_FixComp asm, c
        End If
    End If

    If sizeChanged Or shapeChanged Then
        res = res + 1
        If doLog Then
            If shapeChanged Then PT_LogAdd macroName, keyS, nameS, refD, "resized", "Rotation (deg)", Round(oldRot, 3), Round(PT_ShapeRot, 3)
            For k = 0 To 2
                If Abs(newS(k) - oldS(k)) > tolSize Then PT_LogAdd macroName, keyS, nameS, refD, "resized", CStr(fields(3 + k)), Round(oldS(k), 3), newS(k)
            Next
        End If
    End If
    If posChanged Then
        res = res + 2
        If doLog Then
            For k = 0 To 2
                If Abs(newP(k) - oldP(k)) > tolPos Then PT_LogAdd macroName, keyS, nameS, refD, "moved", CStr(fields(k)), Round(oldP(k), 3), newP(k)
            Next
        End If
    End If
    PT_CheckBoxPart = res
    Exit Function
failed:
    PT_LastErr = Err.Description
    PT_CheckBoxPart = -1
    Resume checkCleanup
checkCleanup:
    On Error Resume Next
    PT_ActivateDoc asm
End Function

' ---------------------------------------------------------------- Simulation helpers
' Simulation document of the assembly (loads the Simulation add-in if needed). Nothing if not available.
Public Function PT_GetSimDoc(ByVal asm As Object) As Object
    Dim addin As Object, cw As Object, exeDir As String
    On Error Resume Next
    PT_ActivateDoc asm
    Set addin = swApp.GetAddInObject("SldWorks.Simulation")
    If addin Is Nothing Then
        exeDir = swApp.GetExecutablePath
        If Right(exeDir, 1) <> "\" Then exeDir = exeDir & "\"
        swApp.LoadAddIn exeDir & "Simulation\cosworks.dll"
        Set addin = swApp.GetAddInObject("SldWorks.Simulation")
    End If
    If addin Is Nothing Then Exit Function
    Set cw = addin.CosmosWorks
    If cw Is Nothing Then Exit Function
    Set PT_GetSimDoc = cw.ActiveDoc
End Function

Public Function PT_FindStudy(ByVal simDoc As Object, ByVal studyName As String) As Object
    Dim sm As Object, i As Long, st As Object
    Set sm = simDoc.StudyManager
    For i = 0 To sm.StudyCount - 1
        Set st = sm.GetStudy(i)
        If Not st Is Nothing Then
            If StrComp(st.Name, studyName, vbTextCompare) = 0 Then Set PT_FindStudy = st: Exit Function
        End If
    Next
End Function

' Solid bodies of a component, assembly context (Variant array or Empty)
Public Function PT_CompBodies(ByVal c As Object) As Variant
    Dim info As Variant
    PT_CompBodies = c.GetBodies3(PT_swSolidBody, info)
End Function

' Named face of a component's part, in assembly context, or Nothing
Public Function PT_AsmFace(ByVal c As Object, ByVal faceName As String) As Object
    Dim md As Object, f As Object
    On Error Resume Next
    Set md = c.GetModelDoc2
    If md Is Nothing Then Exit Function
    Set f = md.GetEntityByName(faceName, 2)
    If f Is Nothing Then Exit Function
    Set PT_AsmFace = c.GetCorrespondingEntity(f)
End Function

' Every active (not suppressed) tagged component: Collection of Array(role, key, refdes, component)
Public Function PT_ListTagged(ByVal asm As Object) As Collection
    Dim comps As Variant, i As Long, c As Object, md As Object, roleS As String, res As Collection
    Set res = New Collection
    asm.ResolveAllLightWeightComponents False
    comps = asm.GetComponents(True)
    If Not IsEmpty(comps) Then
        For i = 0 To UBound(comps)
            Set c = comps(i)
            If c.GetSuppression2 <> PT_swSuppressed Then
                Set md = c.GetModelDoc2
                roleS = PT_GetProp(md, "Role")
                If roleS <> "" Then res.Add Array(roleS, PT_GetProp(md, "Key"), PT_GetProp(md, "RefDes"), c)
            End If
        Next
    End If
    Set PT_ListTagged = res
End Function

' Collection of objects -> Variant array for the Simulation API (Empty if none)
Public Function PT_ToArray(ByVal coll As Collection) As Variant
    Dim a() As Variant, i As Long
    If coll.Count = 0 Then Exit Function
    ReDim a(coll.Count - 1)
    For i = 1 To coll.Count: Set a(i - 1) = coll(i): Next
    PT_ToArray = a
End Function

' ---------------------------------------------------------------- board rules
' Clips a part box (xmin, ymin, zmin, L, W, H in mm, 0-based array) at the PCB surface so no part
' goes into the board (through-hole pins, edge connectors). Top-side parts start at the PCB top,
' bottom-side parts end at the PCB bottom. Returns True if the box was changed.
Public Function PT_ClipToBoard(ByRef bx() As Double, ByVal pcbT As Double) As Boolean
    Dim z0 As Double, z1 As Double, changed As Boolean
    z0 = bx(2): z1 = bx(2) + bx(5)
    If (z0 + z1) / 2 >= pcbT / 2 Then
        If z0 < pcbT - 0.0005 And z1 > pcbT + 0.0005 Then bx(2) = pcbT: bx(5) = z1 - pcbT: changed = True
    Else
        If z1 > 0.0005 And z0 < -0.0005 Then bx(5) = -z0: changed = True
    End If
    PT_ClipToBoard = changed
End Function

' The build and sync macros work on the normal configuration, not on PT_HeatOnly made by SetupSim.
Public Sub PT_LeaveHeatOnly(ByVal asm As Object)
    Dim names As Variant, i As Long
    On Error Resume Next
    If asm.ConfigurationManager.ActiveConfiguration.Name <> "PT_HeatOnly" Then Exit Sub
    names = asm.GetConfigurationNames
    For i = 0 To UBound(names)
        If names(i) <> "PT_HeatOnly" Then asm.ShowConfiguration2 names(i): Exit Sub
    Next
End Sub

' ---------------------------------------------------------------- results tabs (ReadResults, RunFlow)
' Column of "<RefDes> peak (C)" on the Sim results header row; adds the "peak" and "avg" columns when missing.
Public Function PT_RefColumn(ByVal ws As Object, ByVal hdr As Object, ByVal refD As String) As Long
    Dim c As Long, lastC As Long, want As String
    want = refD & " peak (C)"
    lastC = ws.Cells(hdr.Row, ws.Columns.Count).End(-4159).Column
    For c = hdr.Column + hdr.Columns.Count To lastC
        If StrComp(PT_Str(ws.Cells(hdr.Row, c).Value), want, vbTextCompare) = 0 Then PT_RefColumn = c: Exit Function
    Next
    c = lastC + 1
    If c < hdr.Column + hdr.Columns.Count Then c = hdr.Column + hdr.Columns.Count
    ws.Cells(hdr.Row, c).Value = want
    ws.Cells(hdr.Row, c + 1).Value = refD & " avg (C)"
    ws.Range(ws.Cells(hdr.Row, c), ws.Cells(hdr.Row, c + 1)).Interior.Color = hdr.Cells(1, 1).Interior.Color
    ws.Range(ws.Cells(hdr.Row, c), ws.Cells(hdr.Row, c + 1)).Font.Bold = True
    PT_RefColumn = c
End Function

' "U1 Body, U2 Contact R, ..." for the covered chips on the Spreader tab (Spr_Chips offsets: 0 RefDes, 7 Covered, 9 TIM mode, 16 Check)
Public Function PT_TimModes(ByVal wb As Object) As String
    Dim v As Variant, r As Long, s As String
    v = PT_Rng(wb, "Spr_Chips").Value
    For r = 1 To UBound(v, 1)
        If PT_Str(v(r, 1)) <> "" And UCase(PT_Str(v(r, 8))) = "YES" And PT_Str(v(r, 17)) = "OK" Then
            s = s & IIf(s = "", "", ", ") & PT_Str(v(r, 1)) & " " & PT_Str(v(r, 10))
        End If
    Next
    PT_TimModes = s
End Function

' Thermal solution column AF (TS_SimTj): the peak temperature of each RefDes in column B, blank where there is no result.
' peak = Dictionary, RefDes (upper case) -> degrees C.
Public Sub PT_WriteSimTj(ByVal wb As Object, ByVal peak As Object)
    Dim tj As Object, refs As Variant, outv() As Variant, i As Long, rk As String, ws As Object
    Set tj = PT_Rng(wb, "TS_SimTj")
    If tj Is Nothing Then Exit Sub
    Set ws = tj.Worksheet
    refs = ws.Range(ws.Cells(tj.Row, 2), ws.Cells(tj.Row + tj.Rows.Count - 1, 2)).Value
    ReDim outv(1 To tj.Rows.Count, 1 To 1)
    For i = 1 To tj.Rows.Count
        rk = UCase(PT_Str(refs(i, 1)))
        If rk <> "" Then
            If peak.Exists(rk) Then outv(i, 1) = Round(peak(rk), 2)
        End If
    Next
    tj.Columns(1).Value = outv
End Sub

' One pass over the top-level components for several roles at once (SyncFromCAD reads four roles).
' roles = Array("PCB", "Part", ...), haves and supps = Arrays of Dictionaries in the same order (see PT_ReadTagged).
Public Sub PT_ReadRoles(ByVal asm As Object, ByVal roles As Variant, ByVal haves As Variant, ByVal supps As Variant)
    Dim comps As Variant, i As Long, k As Long, hit As Long, c As Object, md As Object, keyS As String, tagRole As String, supp As Boolean
    asm.ResolveAllLightWeightComponents False
    comps = asm.GetComponents(True)
    If IsEmpty(comps) Then Exit Sub
    For i = 0 To UBound(comps)
        Set c = comps(i)
        supp = (c.GetSuppression2 = PT_swSuppressed)
        If supp Then c.SetSuppression2 PT_swFullyResolved
        Set md = c.GetModelDoc2
        tagRole = PT_GetProp(md, "Role")
        hit = -1
        For k = 0 To UBound(roles)
            If roles(k) = tagRole Then hit = k: Exit For
        Next
        If hit >= 0 Then
            keyS = PT_GetProp(md, "Key")
            If keyS = "" Then keyS = "(no key) " & c.Name2
            If Not haves(hit).Exists(keyS) Then
                haves(hit).Add keyS, c
                supps(hit).Add keyS, supp
            End If
        ElseIf supp Then
            c.SetSuppression2 PT_swSuppressed
        End If
    Next
End Sub
' ---------------------------------------------------------------- colours (BuildBoard)
' Same palette on screen and in the workbook: BuildBoard colours each part in SolidWorks and fills its
' Name cell on CAD geometry with the same colour. The PCB is green. Part colour = palette(row index mod 12).
Public Function PT_PcbRgb() As Long
    PT_PcbRgb = RGB(16, 140, 39)
End Function

Public Function PT_PartRgb(ByVal idx As Long) As Long
    Dim p As Variant, k As Long
    p = Array(RGB(66, 133, 244), RGB(244, 160, 0), RGB(219, 68, 55), RGB(171, 71, 188), RGB(0, 172, 193), _
              RGB(255, 202, 40), RGB(141, 110, 99), RGB(236, 64, 122), RGB(92, 107, 192), RGB(255, 112, 67), _
              RGB(120, 144, 156), RGB(38, 198, 218))
    k = idx Mod 12
    If k < 0 Then k = k + 12
    PT_PartRgb = p(k)
End Function

' Sets the colour of a component in the assembly (this configuration). rgbL is a VBA RGB Long. True if accepted.
Public Function PT_ApplyColor(ByVal c As Object, ByVal rgbL As Long) As Boolean
    Dim v(8) As Double, vv As Variant
    v(0) = (rgbL And 255) / 255#
    v(1) = ((rgbL \ 256) And 255) / 255#
    v(2) = ((rgbL \ 65536) And 255) / 255#
    v(3) = 1: v(4) = 1: v(5) = 0.5: v(6) = 0.4: v(7) = 0: v(8) = 0
    vv = v
    On Error GoTo failed
    c.MaterialPropertyValues = vv
    PT_ApplyColor = c.HasMaterialPropertyValues
    If Not PT_ApplyColor Then PT_LastErr = "SolidWorks did not store the colour"
    Exit Function
failed:
    PT_LastErr = Err.Description
End Function
