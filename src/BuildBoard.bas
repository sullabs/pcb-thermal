Attribute VB_Name = "BuildBoard"
Option Explicit
' =====================================================================================
' BuildBoard  -  SolidWorks 2023 macro (milestone M3, project folder mode W6)
'
' Builds the thermal assembly: one box part for the PCB and one box part per included
' body (used position and size). Works for any board.
'
'   No thermal assembly yet  -> creates it and every part.
'   Assembly already there   -> updates it: moves and resizes parts that changed,
'                               adds new ones, suppresses parts no longer included,
'                               and writes a change record for each.
'
' Where it reads and writes (asked at the start, the last one is the default):
'   project.json picked (board folder of the page) -> reads forSolidWorks (bodies, PCB, heat) and settings,
'       writes cad.json build and changes. Refuses while cad.json has change records not ticked on the page,
'       or when the page has not saved since the last Import STEP.
'   a workbook picked -> reads the CAD geometry tab, writes the Change log tab. Refuses while the Change log
'       has unchecked rows.
'
' Parts are fixed at coordinates (no mates). Every part carries custom properties
' Role (PCB / Part), Key, Name, RefDes, SourceRow and its size in micrometres; later
' macros find parts by these properties, never by file name.
' Named faces: PCB_Top, PCB_Bottom on the PCB; Chip_<RefDes>_Top (or Top) and Bottom on parts.
'
' Needs modules PTCommon and PTJson in the same macro.
' =====================================================================================

Const MACRO_NAME As String = "BuildBoard"
Const swDefaultTemplatePart As Long = 8
Const swDefaultTemplateAssembly As Long = 9

' CAD geometry columns, as offsets from the first column of Geo_Table (column B)
Const G_NAMEUSED As Long = 3      ' E
Const G_REF As Long = 4           ' F
Const G_INC As Long = 5           ' G
Const G_USED As Long = 19         ' U..Z used Xmin Ymin Zmin L W H
Const G_KEY As Long = 31          ' AG
Const G_ROT As Long = 33          ' AI imported rotation (deg), AJ own L, AK own W (added by ImportStep)
Const G_OWNL As Long = 34
Const G_OWNW As Long = 35

' wanted parts, from the workbook or from project.json forSolidWorks
Dim wKey() As String, wName() As String, wRef() As String, wRow() As Long, wBox() As Double, nW As Long
Dim wRot() As Double, wOwnL() As Double, wOwnW() As Double       ' rotated footprint per wanted part (0 = plain box)
Dim wCut() As Double                                             ' mm cut off at the board surface
Dim wFail() As String                                            ' why the part failed ("" = built)
Dim nRotated As Long
Dim skippedRows As Collection, failMsgs As Collection, clipped As Collection
Dim skipList As Collection                                       ' Array(key, name, reason) of included bodies not built
Dim ovlList As Collection                                        ' Array(a, b, x, y, z) of overlapping boxes
Dim isUpdate As Boolean
Dim nAdded As Long, nMoved As Long, nResized As Long, nSupp As Long, nRestored As Long
Dim geoRow0 As Long, nColored As Long, nColorFail As Long     ' sheet row of the first Geo_Table row; colour counts
Dim noPartHeat As String                 ' RefDes with heat but no included part
Dim mRpt As String                       ' report path of this run
Dim mOverlaps As Collection              ' overlap texts of this run

Sub main()
    Set swApp = Application.SldWorks
    PT_Trace MACRO_NAME & ": started (VBA macro)."
    PT_LogReset
    Set skippedRows = New Collection
    Set failMsgs = New Collection
    Set clipped = New Collection
    Set skipList = New Collection
    Set ovlList = New Collection
    nAdded = 0: nMoved = 0: nResized = 0: nSupp = 0: nRestored = 0: nColored = 0: nColorFail = 0

    ' ---------- project folder (page) or workbook ----------
    Select Case PJ_PickSource(MACRO_NAME)
        Case "json"
            RunProject
            Exit Sub
        Case "xlsx"
        Case Else
            Exit Sub
    End Select

    ' ---------- workbook ----------
    Dim xl As Object, wb As Object
    Set xl = PT_GetExcel()
    Set wb = PT_PickWorkbook(xl, MACRO_NAME, "CAD geometry|Settings|Change log", PJ_WbPath)
    If wb Is Nothing Then Exit Sub
    If Not PT_NeedNames(wb, "Geo_Table|Geo_PCB_Used|Log_Header|Log_Unchecked|Set_Prefix|Set_TolPos|Set_TolSize|" & _
                            "Set_PartTemplate|Set_AsmTemplate|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
    If PT_RefuseIfUnchecked(wb, MACRO_NAME) Then Exit Sub

    Dim tolPos As Double, tolSize As Double, prefix As String, outDir As String
    tolPos = PT_NameNum(wb, "Set_TolPos", 0.05)
    tolSize = PT_NameNum(wb, "Set_TolSize", 0.05)
    prefix = PT_SafeFile(PT_Prefix(wb))

    Dim pcbL As Double, pcbW As Double, pcbT As Double, rp As Object
    Set rp = PT_Rng(wb, "Geo_PCB_Used")
    pcbL = PT_Num(rp.Cells(1, 1).Value, 0)
    pcbW = PT_Num(rp.Cells(2, 1).Value, 0)
    pcbT = PT_Num(rp.Cells(3, 1).Value, 0)
    ReadWanted wb, pcbT
    If pcbL <= 0 Or pcbW <= 0 Or pcbT <= 0 Then
        PT_Msg "BuildBoard did not run. The PCB size on CAD geometry (Used column, Length X, Width Y, Thickness Z) is missing." & vbCrLf & _
               "Run ImportStep or type an override there.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    If nW = 0 Then
        If PT_Msg("No row on CAD geometry has Include = Yes with a size and a key (" & skippedRows.Count & " included rows were skipped)." & vbCrLf & _
                  "Only the PCB would be built." & vbCrLf & vbCrLf & "Build it anyway?", vbYesNo + vbQuestion, MACRO_NAME) <> vbYes Then Exit Sub
    End If
    noPartHeat = HeatWithoutPart(wb)

    ' ---------- files and templates ----------
    Dim partTmpl As String, asmTmpl As String, asmPath As String
    outDir = PT_OutputFolder(wb)
    partTmpl = PT_Template(wb, "Set_PartTemplate", swDefaultTemplatePart)
    asmTmpl = PT_Template(wb, "Set_AsmTemplate", swDefaultTemplateAssembly)
    If partTmpl = "" Or asmTmpl = "" Then
        PT_Msg "BuildBoard did not run. The " & IIf(partTmpl = "", "part", "assembly") & " template was not found." & vbCrLf & _
               "Set its full path on the Settings tab.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    asmPath = PT_AsmPath(wb)

    ' ---------- assembly and report ----------
    If Not BuildAll(partTmpl, asmTmpl, asmPath, outDir, prefix, pcbL, pcbW, pcbT, tolPos, tolSize, "Workbook", wb.FullName) Then Exit Sub

    ' ---------- Change log ----------
    Dim nLog As Long
    On Error GoTo xlFail
    xl.ScreenUpdating = False
    nLog = PT_LogFlush(wb)
    xl.ScreenUpdating = True
    xl.Calculate
    Dim chipRows As Long
    chipRows = PT_FillChipLists(wb)
    wb.Save
    On Error GoTo 0

    Dim s As String
    s = "Written and saved to:" & vbCrLf & asmPath & vbCrLf & "(parts in the same folder)" & vbCrLf & vbCrLf
    If isUpdate Then
        s = s & UpdateLine() & vbCrLf
        If nLog > 0 Then
            s = s & nLog & " rows added to the Change log tab of " & wb.Name & ". Check them there." & vbCrLf
        Else
            s = s & "Nothing changed." & vbCrLf
        End If
    Else
        s = s & "Created the thermal assembly: PCB plus " & (nAdded - 1) & " parts." & vbCrLf
    End If
    s = s & CommonLines()
    If chipRows > 0 Then s = s & chipRows & " chip rows written to the Spreader and Sim setup tabs (they were empty). Check the Covered column." & vbCrLf
    s = s & nColored & " parts coloured in SolidWorks (PCB green). The workbook shows the same colours on its own." & _
        IIf(nColorFail > 0, " " & nColorFail & " could not be coloured.", "") & vbCrLf
    If noPartHeat <> "" Then s = s & "Heat in the register but no built part for: " & noPartHeat & ". Set RefDes and Include = Yes on CAD geometry, or they get no heat source." & vbCrLf
    s = s & mOverlaps.Count & " overlapping pairs of boxes left (listed in the report). Fix them before SetupSim." & vbCrLf & vbCrLf & _
        "Report: " & mRpt
    PT_Msg s, IIf(failMsgs.Count > 0, vbExclamation, vbInformation), MACRO_NAME
    Exit Sub

xlFail:
    xl.ScreenUpdating = True
    PT_Msg "The assembly was saved to " & asmPath & ", but writing the Change log failed: " & Err.Description & vbCrLf & _
           "Close any cell you are editing in Excel. The changes are listed in " & mRpt, vbExclamation, MACRO_NAME
End Sub

' ================================================================= project folder (page) mode
' Reads forSolidWorks (worked out by the page on every save) and the settings from project.json, builds or updates
' the assembly, and writes cad.json build (what was built, failures, overlaps) and changes.
Sub RunProject()
    Dim cad As Object, cadPath As String, impDate As String, fsDate As String
    cadPath = PJ_Dir & "\" & PJ_File("cad")
    Set cad = PJ_ReadOwn(PJ_File("cad"), "pcb-thermal/cad")
    If cad Is Nothing Then
        PT_Msg "BuildBoard did not run. Nothing was changed." & vbCrLf & vbCrLf & cadPath & " could not be read:" & vbCrLf & PT_LastErr, _
               vbExclamation, MACRO_NAME
        Exit Sub
    End If
    If PJ_Obj(cad, "import") Is Nothing Then
        PT_Msg "BuildBoard did not run. Nothing was changed." & vbCrLf & vbCrLf & "There is no import in " & cadPath & "." & vbCrLf & _
               "Run Import STEP first (button 2).", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    If PJ_Obj(PJ_Project, "forSolidWorks") Is Nothing Then
        PT_Msg "BuildBoard did not run. Nothing was changed." & vbCrLf & vbCrLf & _
               "project.json has no forSolidWorks yet (the page writes it on Save once the PCB size and k are known)." & vbCrLf & _
               "Open the page, check tab 3 (Link bodies), press Save and run again.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    impDate = PJ_Str(cad, "import/date", "")
    fsDate = PJ_Str(PJ_Project, "forSolidWorks/computedAt", "")
    If impDate <> "" And fsDate < impDate Then
        PT_Msg "BuildBoard did not run. Nothing was changed." & vbCrLf & vbCrLf & _
               "The page has not saved since the last Import STEP:" & vbCrLf & _
               "  Import STEP " & Replace(impDate, "T", " ") & vbCrLf & _
               "  project.json worked out " & Replace(fsDate, "T", " ") & vbCrLf & vbCrLf & _
               "Open the page, check tab 3 (Link bodies), press Save and run again, so the build uses the new import.", _
               vbExclamation, MACRO_NAME
        Exit Sub
    End If
    If PJ_RefuseIfUnchecked(cad, MACRO_NAME) Then Exit Sub

    Dim tolPos As Double, tolSize As Double, prefix As String, outDir As String
    tolPos = PJ_Num(PJ_Project, "settings/tolPosMm", 0.05)
    tolSize = PJ_Num(PJ_Project, "settings/tolSizeMm", 0.05)
    prefix = PT_SafeFile(PJ_Prefix())

    Dim pcbL As Double, pcbW As Double, pcbT As Double
    pcbL = PJ_Num(PJ_Project, "forSolidWorks/pcb/l", 0)
    pcbW = PJ_Num(PJ_Project, "forSolidWorks/pcb/w", 0)
    pcbT = PJ_Num(PJ_Project, "forSolidWorks/pcb/h", 0)
    ReadWantedJson cad, pcbT
    If pcbL <= 0 Or pcbW <= 0 Or pcbT <= 0 Then
        PT_Msg "BuildBoard did not run. The PCB size (L, W, thickness) is missing in project.json forSolidWorks." & vbCrLf & _
               "Check the PCB size on tab 3 of the page, save, and run again.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    If nW = 0 Then
        If PT_Msg("No body is included with a size and a key (" & skippedRows.Count & " included bodies were skipped)." & vbCrLf & _
                  "Only the PCB would be built." & vbCrLf & vbCrLf & "Build it anyway?", vbYesNo + vbQuestion, MACRO_NAME) <> vbYes Then Exit Sub
    End If
    noPartHeat = HeatWithoutPartJson()

    ' ---------- files and templates ----------
    Dim partTmpl As String, asmTmpl As String, asmPath As String
    outDir = PJ_OutputFolder()
    partTmpl = PJ_Template("partTemplate", swDefaultTemplatePart)
    asmTmpl = PJ_Template("asmTemplate", swDefaultTemplateAssembly)
    If partTmpl = "" Or asmTmpl = "" Then
        PT_Msg "BuildBoard did not run. The " & IIf(partTmpl = "", "part", "assembly") & " template was not found." & vbCrLf & _
               "Leave it blank in the project settings to use the SolidWorks default, or give its full path.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    asmPath = PJ_AsmPath()

    ' ---------- assembly and report ----------
    If Not BuildAll(partTmpl, asmTmpl, asmPath, outDir, prefix, pcbL, pcbW, pcbT, tolPos, tolSize, "Project", PJ_Dir) Then Exit Sub

    ' ---------- cad.json ----------
    Dim nLog As Long
    PJ_Put cad, "build", BuildRecord(asmPath)
    nLog = PJ_FlushChanges(cad)
    If Not PJ_SaveOwn(cad, PJ_File("cad"), MACRO_NAME) Then
        PT_Msg "The assembly was saved to " & asmPath & ", but " & cadPath & " could not be written:" & vbCrLf & PT_LastErr & vbCrLf & vbCrLf & _
               "The changes are listed in " & mRpt & ". Run Build Board again to write cad.json.", vbExclamation, MACRO_NAME
        Exit Sub
    End If

    Dim s As String
    s = "Written and saved to:" & vbCrLf & asmPath & vbCrLf & "(parts in the same folder)" & vbCrLf & _
        "and " & cadPath & " (the page reloads it on its own, tab 6 Build model)." & vbCrLf & vbCrLf
    If isUpdate Then
        s = s & UpdateLine() & vbCrLf
        If nLog > 0 Then
            s = s & nLog & " change records written to cad.json. Tick them on tab 10 of the page (Changes)." & vbCrLf
        Else
            s = s & "Nothing changed." & vbCrLf
        End If
    Else
        s = s & "Created the thermal assembly: PCB plus " & (nAdded - 1) & " parts." & vbCrLf
    End If
    s = s & CommonLines()
    s = s & nColored & " parts coloured in SolidWorks (PCB green)." & _
        IIf(nColorFail > 0, " " & nColorFail & " could not be coloured.", "") & vbCrLf
    If noPartHeat <> "" Then s = s & "Heat but no built part for: " & noPartHeat & ". Link and include their bodies on tab 3, or they get no heat source." & vbCrLf
    s = s & mOverlaps.Count & " overlapping pairs of boxes left (listed in the report and on tab 6). Fix them before the study." & vbCrLf & vbCrLf & _
        "Report: " & mRpt
    PT_Msg s, IIf(failMsgs.Count > 0, vbExclamation, vbInformation), MACRO_NAME
End Sub

' cad.json build: assembly, date, mode, counts, one record per included body, overlaps, report.
Function BuildRecord(ByVal asmPath As String) As Object
    Dim d As Object, cnts As Object, bl As Collection, b As Object, ol As Collection, o As Object, sz As Collection
    Dim i As Long, it As Variant
    Set d = PJ_Dict()
    PJ_Put d, "assembly", asmPath
    PJ_Put d, "date", PJ_Now()
    PJ_Put d, "mode", IIf(isUpdate, "update", "create")
    PJ_Put d, "inputsHash", PJ_Str(PJ_Project, "forSolidWorks/inputsHash", "")            ' the page compares it to spot an out-of-date model
    Set cnts = PJ_Dict()
    PJ_Put cnts, "wanted", nW
    PJ_Put cnts, "added", nAdded
    PJ_Put cnts, "moved", nMoved
    PJ_Put cnts, "resized", nResized
    PJ_Put cnts, "suppressed", nSupp
    PJ_Put cnts, "broughtBack", nRestored
    PJ_Put cnts, "failed", failMsgs.Count
    PJ_Put cnts, "rotated", nRotated
    PJ_Put d, "counts", cnts
    Set bl = New Collection
    For i = 0 To nW - 1
        Set b = PJ_Dict()
        PJ_Put b, "key", wKey(i)
        PJ_Put b, "name", wName(i)
        If wFail(i) <> "" Then
            PJ_Put b, "status", "failed"
            PJ_Put b, "reason", wFail(i)
        Else
            PJ_Put b, "status", "built"
        End If
        If wCut(i) > 0.0005 Then PJ_Put b, "cutAtBoardMm", Round(wCut(i), 4)
        bl.Add b
    Next
    For Each it In skipList
        Set b = PJ_Dict()
        If KeyOk(CStr(it(0))) Then PJ_Put b, "key", CStr(it(0))
        PJ_Put b, "name", CStr(it(1))
        PJ_Put b, "status", "skipped"
        PJ_Put b, "reason", CStr(it(2))
        bl.Add b
    Next
    PJ_Put d, "bodies", bl
    Set ol = New Collection
    For Each it In ovlList
        Set o = PJ_Dict()
        PJ_Put o, "a", CStr(it(0))
        PJ_Put o, "b", CStr(it(1))
        Set sz = New Collection
        sz.Add Round(CDbl(it(2)), 3): sz.Add Round(CDbl(it(3)), 3): sz.Add Round(CDbl(it(4)), 3)
        PJ_Put o, "size", sz
        ol.Add o
    Next
    PJ_Put d, "overlaps", ol
    PJ_Put d, "report", mRpt
    Set BuildRecord = d
End Function

' A body key as the schema wants it: lower case letters, digits and _ only.
Function KeyOk(ByVal keyS As String) As Boolean
    KeyOk = (keyS <> "" And Not keyS Like "*[!a-z0-9_]*")
End Function

' ================================================================= both modes
' Opens or makes the assembly, builds every part, colours, saves, finds overlaps, writes the report.
' False when SolidWorks failed (the user has been told).
Function BuildAll(ByVal partTmpl As String, ByVal asmTmpl As String, ByVal asmPath As String, ByVal outDir As String, _
                  ByVal prefix As String, ByVal pcbL As Double, ByVal pcbW As Double, ByVal pcbT As Double, _
                  ByVal tolPos As Double, ByVal tolSize As Double, ByVal srcProp As String, ByVal srcPath As String) As Boolean
    Dim asm As Object, errs As Long, warns As Long
    On Error GoTo swFail
    Set asm = PT_OpenAsm(asmPath)
    If asm Is Nothing And Dir(asmPath) <> "" Then
        PT_Msg "BuildBoard could not open " & asmPath & ". Nothing was changed.", vbExclamation, MACRO_NAME
        Exit Function
    End If
    isUpdate = Not asm Is Nothing
    If Not isUpdate Then
        Set asm = swApp.NewDocument(asmTmpl, 0, 0, 0)
        If asm Is Nothing Then
            PT_Msg "SolidWorks could not make a new assembly from" & vbCrLf & asmTmpl, vbExclamation, MACRO_NAME
            Exit Function
        End If
        If Not asm.Extension.SaveAs(asmPath, 0, PT_swSaveSilent, Nothing, errs, warns) Then
            PT_Msg "Could not save the new assembly as" & vbCrLf & asmPath & vbCrLf & "(error " & errs & "). Is the folder writable?", _
                   vbExclamation, MACRO_NAME
            Exit Function
        End If
    End If
    PT_LeaveHeatOnly asm
    PT_SetProp asm, "Role", "ThermalAssembly"
    PT_SetProp asm, srcProp, srcPath
    PT_ActivateDoc asm

    SyncAssembly asm, partTmpl, outDir, prefix, pcbL, pcbW, pcbT, tolPos, tolSize

    ColorParts asm
    PT_ActivateDoc asm
    asm.ForceRebuild3 False
    asm.GraphicsRedraw2
    asm.Save3 PT_swSaveSilent + PT_swSaveReferenced, errs, warns
    On Error GoTo 0

    Set mOverlaps = FindOverlaps(pcbL, pcbW, pcbT)
    mRpt = outDir & prefix & "BuildBoard_report.txt"
    WriteReport mRpt, asmPath, mOverlaps
    BuildAll = True
    Exit Function

swFail:
    PT_Msg "BuildBoard stopped in SolidWorks: " & Err.Description & vbCrLf & _
           "The assembly may be half done and was not saved. Close it without saving, send this message, and run again.", _
           vbExclamation, MACRO_NAME
End Function

Function UpdateLine() As String
    UpdateLine = "Updated the thermal assembly: " & nAdded & " parts added, " & nMoved & " moved, " & nResized & " resized, " & _
                 nSupp & " suppressed, " & nRestored & " brought back."
End Function

Function CommonLines() As String
    Dim s As String
    If skippedRows.Count > 0 Then s = s & skippedRows.Count & " included bodies skipped (listed in the report)." & vbCrLf
    If failMsgs.Count > 0 Then s = s & failMsgs.Count & " parts FAILED (listed in the report)." & vbCrLf
    If clipped.Count > 0 Then s = s & clipped.Count & " parts cut at the board surface so they do not go into the PCB (listed in the report)." & vbCrLf
    If nRotated > 0 Then s = s & nRotated & " of the wanted parts are rotated rectangles (turned in the board plane)." & vbCrLf
    CommonLines = s
End Function

' ================================================================= reading the wanted parts
Sub SizeWanted(ByVal cnt As Long)
    If cnt < 1 Then cnt = 1
    ReDim wKey(cnt): ReDim wName(cnt): ReDim wRef(cnt)
    ReDim wRow(cnt): ReDim wBox(cnt, 5)
    ReDim wRot(cnt): ReDim wOwnL(cnt): ReDim wOwnW(cnt)
    ReDim wCut(cnt): ReDim wFail(cnt)
    nW = 0
    nRotated = 0
End Sub

' Workbook: included rows of CAD geometry.
Sub ReadWanted(ByVal wb As Object, ByVal pcbT As Double)
    Dim rg As Object, v As Variant, r As Long, k As Long, used As Object, ok As Boolean, rowName As String, bx(5) As Double
    Dim hasRot As Boolean, rotV As Double, ownLV As Double, ownWV As Double
    Dim hOld As Double
    Set rg = PT_Rng(wb, "Geo_Table")
    geoRow0 = rg.Row
    v = rg.Value
    Set used = CreateObject("Scripting.Dictionary")
    SizeWanted UBound(v, 1)
    hasRot = (UBound(v, 2) >= 1 + G_OWNW)
    For r = 1 To UBound(v, 1)
        rowName = PT_Str(v(r, 1 + G_NAMEUSED))
        If rowName <> "" And UCase(PT_Str(v(r, 1 + G_INC))) = "YES" Then
            ok = True
            For k = 0 To 5
                If Not PT_IsNum(v(r, 1 + G_USED + k)) Then ok = False
            Next
            If ok Then
                If PT_Num(v(r, 1 + G_USED + 3), 0) <= 0 Or PT_Num(v(r, 1 + G_USED + 4), 0) <= 0 Or _
                   PT_Num(v(r, 1 + G_USED + 5), 0) <= 0 Then ok = False
            End If
            If Not ok Then
                skippedRows.Add "Row " & (rg.Row + r - 1) & " " & rowName & ": position or size missing or zero"
            ElseIf PT_Str(v(r, 1 + G_KEY)) = "" Then
                skippedRows.Add "Row " & (rg.Row + r - 1) & " " & rowName & ": no key (run SetupWorkbook)"
            Else
                wKey(nW) = PT_UniqueKey(used, PT_Str(v(r, 1 + G_KEY)))
                wName(nW) = rowName
                wRef(nW) = PT_Str(v(r, 1 + G_REF))
                wRow(nW) = rg.Row + r - 1
                For k = 0 To 5: bx(k) = CDbl(v(r, 1 + G_USED + k)): Next
                If pcbT > 0 Then
                    hOld = CDbl(v(r, 1 + G_USED + 5))
                    ' a part that sinks less than 0.05 mm is STEP rounding: cut it quietly, list only the real ones
                    If PT_ClipToBoard(bx, pcbT) Then
                        wCut(nW) = hOld - bx(5)
                        If hOld - bx(5) > 0.05 Then clipped.Add rowName & ": box cut at the board surface (" & _
                                                                Round(hOld - bx(5), 3) & " mm went into the PCB)"
                    End If
                End If
                For k = 0 To 5: wBox(nW, k) = bx(k): Next
                ' rotated footprint: used only when the box was not overridden in L or W (envelope still fits)
                If hasRot Then
                    If PT_IsNum(v(r, 1 + G_ROT)) And PT_IsNum(v(r, 1 + G_OWNL)) And PT_IsNum(v(r, 1 + G_OWNW)) Then
                        rotV = CDbl(v(r, 1 + G_ROT)): ownLV = CDbl(v(r, 1 + G_OWNL)): ownWV = CDbl(v(r, 1 + G_OWNW))
                        If PT_ShapeIsRotated(rotV, ownLV, ownWV, bx(3), bx(4)) Then
                            wRot(nW) = rotV: wOwnL(nW) = ownLV: wOwnW(nW) = ownWV
                            nRotated = nRotated + 1
                        End If
                    End If
                End If
                nW = nW + 1
            End If
        End If
    Next
End Sub

' Project: forSolidWorks.bodies (the included bodies with their used box, overrides already applied by the page).
' wRow holds the palette index: the body's place in the import (bodies typed in by hand after them), as the
' workbook's row order.
Sub ReadWantedJson(ByVal cad As Object, ByVal pcbT As Double)
    Dim bl As Object, impB As Object, it As Variant, used As Object, place As Object, k As Long, ok As Boolean
    Dim flds As Variant, nameS As String, keyS As String, bx(5) As Double, hOld As Double, nextPlace As Long
    Dim rotV As Double, ownLV As Double, ownWV As Double
    flds = Array("xMin", "yMin", "zMin", "l", "w", "h")
    Set used = CreateObject("Scripting.Dictionary")
    Set place = CreateObject("Scripting.Dictionary")
    Set impB = PJ_Obj(cad, "import/bodies")
    If Not impB Is Nothing Then
        For Each it In impB
            keyS = PJ_Str(it, "key", "")
            If keyS <> "" And Not place.Exists(keyS) Then place.Add keyS, place.Count
        Next
    End If
    nextPlace = place.Count
    geoRow0 = 0
    Set bl = PJ_Obj(PJ_Project, "forSolidWorks/bodies")
    If bl Is Nothing Then
        SizeWanted 0
        Exit Sub
    End If
    SizeWanted bl.Count
    For Each it In bl
        nameS = PJ_Str(it, "name", "")
        keyS = PJ_Str(it, "key", "")
        ok = True
        For k = 0 To 5
            If Not PJ_IsNum(it, "box/" & flds(k)) Then ok = False
        Next
        If ok Then
            If PJ_Num(it, "box/l", 0) <= 0 Or PJ_Num(it, "box/w", 0) <= 0 Or PJ_Num(it, "box/h", 0) <= 0 Then ok = False
        End If
        If Not ok Then
            skippedRows.Add nameS & " (" & keyS & "): position or size missing or zero"
            skipList.Add Array(keyS, nameS, "position or size missing or zero")
        ElseIf keyS = "" Then
            skippedRows.Add nameS & ": no key"
            skipList.Add Array(keyS, nameS, "no key")
        Else
            wKey(nW) = PT_UniqueKey(used, keyS)
            wName(nW) = nameS
            wRef(nW) = PJ_Str(it, "refDes", "")
            If place.Exists(keyS) Then
                wRow(nW) = place(keyS)
            Else
                wRow(nW) = nextPlace: nextPlace = nextPlace + 1
            End If
            For k = 0 To 5: bx(k) = PJ_Num(it, "box/" & flds(k), 0): Next
            If pcbT > 0 Then
                hOld = bx(5)
                ' a part that sinks less than 0.05 mm is STEP rounding: cut it quietly, list only the real ones
                If PT_ClipToBoard(bx, pcbT) Then
                    wCut(nW) = hOld - bx(5)
                    If hOld - bx(5) > 0.05 Then clipped.Add nameS & ": box cut at the board surface (" & _
                                                            Round(hOld - bx(5), 3) & " mm went into the PCB)"
                End If
            End If
            For k = 0 To 5: wBox(nW, k) = bx(k): Next
            ' rotated footprint: the page sends rotDeg only when the box was not overridden in L or W
            If PJ_IsNum(it, "rotDeg") And PJ_IsNum(it, "ownL") And PJ_IsNum(it, "ownW") Then
                rotV = PJ_Num(it, "rotDeg", 0): ownLV = PJ_Num(it, "ownL", 0): ownWV = PJ_Num(it, "ownW", 0)
                If PT_ShapeIsRotated(rotV, ownLV, ownWV, bx(3), bx(4)) Then
                    wRot(nW) = rotV: wOwnL(nW) = ownLV: wOwnW(nW) = ownWV
                    nRotated = nRotated + 1
                End If
            End If
            nW = nW + 1
        End If
    Next
End Sub

' ================================================================= SolidWorks side
Sub SyncAssembly(ByVal asm As Object, ByVal partTmpl As String, ByVal outDir As String, ByVal prefix As String, _
                 ByVal pcbL As Double, ByVal pcbW As Double, ByVal pcbT As Double, ByVal tolPos As Double, ByVal tolSize As Double)
    Dim have As Object, wasSupp As Object, want As Object, pcbHave As Object, pcbSupp As Object
    Dim pcbComp As Object, i As Long, res As Long, fpath As String, okAdd As Boolean
    Set have = CreateObject("Scripting.Dictionary")
    Set wasSupp = CreateObject("Scripting.Dictionary")
    Set want = CreateObject("Scripting.Dictionary")
    Set pcbHave = CreateObject("Scripting.Dictionary")
    Set pcbSupp = CreateObject("Scripting.Dictionary")
    For i = 0 To nW - 1: want.Add wKey(i), i: Next
    PT_ShapeClear

    Set pcbComp = PT_ReadTagged(asm, "PCB", pcbHave, pcbSupp)
    PT_ReadTagged asm, "Part", have, wasSupp

    ' ---- PCB
    If pcbComp Is Nothing Then
        If PT_AddBoxPart(asm, partTmpl, outDir & prefix & "PCB.sldprt", "PCB", "pcb", "PCB", "", 0, 0, 0, 0, _
                         pcbL, pcbW, pcbT, "PCB_Top", "PCB_Bottom", MACRO_NAME) Then
            nAdded = nAdded + 1
            If isUpdate Then PT_LogAdd MACRO_NAME, "pcb", "PCB", "", "added", "Box", "", PT_BoxText(0, 0, 0, pcbL, pcbW, pcbT)
        Else
            failMsgs.Add "Could not add the PCB: " & PT_LastErr
        End If
    Else
        res = PT_CheckBoxPart(asm, pcbComp, "PCB", "pcb", "PCB", "", 0, 0, 0, 0, pcbL, pcbW, pcbT, _
                              "PCB_Top", "PCB_Bottom", tolPos, tolSize, MACRO_NAME, isUpdate)
        CountResult res, "PCB"
    End If

    ' ---- parts
    For i = 0 To nW - 1
        If have.Exists(wKey(i)) Then
            If wasSupp(wKey(i)) Then
                nRestored = nRestored + 1
                If isUpdate Then PT_LogAdd MACRO_NAME, wKey(i), wName(i), wRef(i), "added", "Brought back (was suppressed)", "", ""
            End If
            PT_ShapeRot = wRot(i): PT_ShapeL = wOwnL(i): PT_ShapeW = wOwnW(i)
            res = PT_CheckBoxPart(asm, have(wKey(i)), "Part", wKey(i), wName(i), wRef(i), wRow(i), _
                                  wBox(i, 0), wBox(i, 1), wBox(i, 2), wBox(i, 3), wBox(i, 4), wBox(i, 5), _
                                  TopName(wRef(i)), "Bottom", tolPos, tolSize, MACRO_NAME, isUpdate)
            PT_ShapeClear
            CountResult res, wName(i)
            If res < 0 Then wFail(i) = PT_LastErr
        Else
            fpath = outDir & prefix & "P_" & PT_SafeFile(wKey(i)) & ".sldprt"
            PT_ShapeRot = wRot(i): PT_ShapeL = wOwnL(i): PT_ShapeW = wOwnW(i)
            okAdd = PT_AddBoxPart(asm, partTmpl, fpath, "Part", wKey(i), wName(i), wRef(i), wRow(i), _
                             wBox(i, 0), wBox(i, 1), wBox(i, 2), wBox(i, 3), wBox(i, 4), wBox(i, 5), _
                             TopName(wRef(i)), "Bottom", MACRO_NAME)
            PT_ShapeClear
            If okAdd Then
                nAdded = nAdded + 1
                If isUpdate Then PT_LogAdd MACRO_NAME, wKey(i), wName(i), wRef(i), "added", "Box", "", _
                                           PT_BoxText(wBox(i, 0), wBox(i, 1), wBox(i, 2), wBox(i, 3), wBox(i, 4), wBox(i, 5))
            Else
                failMsgs.Add "Could not add " & wName(i) & " (" & wKey(i) & "): " & PT_LastErr
                wFail(i) = PT_LastErr
            End If
        End If
    Next

    ' ---- parts no longer wanted: suppress, never delete
    nSupp = PT_SuppressUnwanted(have, wasSupp, want, MACRO_NAME, isUpdate)
End Sub

' Colours the PCB green and every wanted part with its palette colour (palette index = its row on CAD geometry,
' or its place in the import in project mode).
Sub ColorParts(ByVal asm As Object)
    Dim idx As Object, i As Long, it As Variant, rgbL As Long, ok As Boolean
    Set idx = CreateObject("Scripting.Dictionary")
    For i = 0 To nW - 1: idx(wKey(i)) = wRow(i) - geoRow0: Next
    For Each it In PT_ListTagged(asm)
        ok = False
        If it(0) = "PCB" Then
            rgbL = PT_PcbRgb(): ok = True
        ElseIf it(0) = "Part" Then
            If idx.Exists(it(1)) Then rgbL = PT_PartRgb(CLng(idx(it(1)))): ok = True
        End If
        If ok Then
            If PT_ApplyColor(it(3), rgbL) Then nColored = nColored + 1 Else nColorFail = nColorFail + 1
        End If
    Next
End Sub

' RefDes with heat in the register that no included part on CAD geometry carries. SetupSim would give them no heat source.
Function HeatWithoutPart(ByVal wb As Object) As String
    Dim rr As Object, rq As Object, refs As Variant, qs As Variant, r As Long, have As Object, i As Long, s As String, cnt As Long
    Set rr = PT_Rng(wb, "Reg_RefDes"): Set rq = PT_Rng(wb, "Reg_QUsed")
    If rr Is Nothing Or rq Is Nothing Then Exit Function
    Set have = CreateObject("Scripting.Dictionary")
    For i = 0 To nW - 1
        If wRef(i) <> "" Then have(UCase(wRef(i))) = 1
    Next
    refs = rr.Value: qs = rq.Value
    For r = 1 To UBound(refs, 1)
        If PT_Str(refs(r, 1)) <> "" And PT_Num(qs(r, 1), 0) > 0 Then
            If Not have.Exists(UCase(PT_Str(refs(r, 1)))) Then
                cnt = cnt + 1
                If cnt <= 8 Then s = s & IIf(s = "", "", ", ") & PT_Str(refs(r, 1))
            End If
        End If
    Next
    If cnt > 8 Then s = s & " and " & (cnt - 8) & " more"
    HeatWithoutPart = s
End Function

' Same from project.json: RefDes in forSolidWorks.heat above 0 W that no wanted part carries.
Function HeatWithoutPartJson() As String
    Dim heat As Object, have As Object, i As Long, rk As Variant, s As String, cnt As Long
    Set heat = PJ_Obj(PJ_Project, "forSolidWorks/heat")
    If heat Is Nothing Then Exit Function
    Set have = CreateObject("Scripting.Dictionary")
    For i = 0 To nW - 1
        If wRef(i) <> "" Then have(UCase(wRef(i))) = 1
    Next
    For Each rk In heat.Keys
        If PJ_Num(heat, CStr(rk), 0) > 0 And Not have.Exists(UCase(CStr(rk))) Then
            cnt = cnt + 1
            If cnt <= 8 Then s = s & IIf(s = "", "", ", ") & rk
        End If
    Next
    If cnt > 8 Then s = s & " and " & (cnt - 8) & " more"
    HeatWithoutPartJson = s
End Function

Sub CountResult(ByVal res As Long, ByVal what As String)
    If res < 0 Then
        failMsgs.Add "Could not update " & what & ": " & PT_LastErr
    Else
        If (res And 1) <> 0 Then nResized = nResized + 1
        If (res And 2) <> 0 Then nMoved = nMoved + 1
    End If
End Sub

Function TopName(ByVal refD As String) As String
    If refD <> "" Then TopName = "Chip_" & PT_SafeFile(refD) & "_Top" Else TopName = "Top"
End Function

' ================================================================= checks and report
' Pairs of boxes (parts and PCB) that overlap by more than 0.01 mm in X, Y and Z.
' Also fills ovlList (keys and overlap sizes) for cad.json.
Function FindOverlaps(ByVal pcbL As Double, ByVal pcbW As Double, ByVal pcbT As Double) As Collection
    Dim res As Collection, i As Long, j As Long, k As Long, ov(2) As Double, hit As Boolean
    Dim a0 As Double, a1 As Double, b0 As Double, b1 As Double, pcbBox(2) As Double
    Set res = New Collection
    Set ovlList = New Collection
    pcbBox(0) = pcbL: pcbBox(1) = pcbW: pcbBox(2) = pcbT
    For i = 0 To nW - 1
        hit = True
        For k = 0 To 2
            a0 = wBox(i, k): a1 = wBox(i, k) + wBox(i, k + 3)
            ov(k) = Min2(a1, pcbBox(k)) - Max2(a0, 0)
            If ov(k) <= 0.01 Then hit = False
        Next
        If hit Then
            res.Add wName(i) & " goes into the PCB by " & Round(ov(2), 2) & " mm (overlap " & _
                    Round(ov(0), 2) & " x " & Round(ov(1), 2) & " mm)"
            ovlList.Add Array(wKey(i), "pcb", ov(0), ov(1), ov(2))
        End If
        For j = i + 1 To nW - 1
            hit = True
            For k = 0 To 2
                a0 = wBox(i, k): a1 = wBox(i, k) + wBox(i, k + 3)
                b0 = wBox(j, k): b1 = wBox(j, k) + wBox(j, k + 3)
                ov(k) = Min2(a1, b1) - Max2(a0, b0)
                If ov(k) <= 0.01 Then hit = False
            Next
            If hit And (wRot(i) <> 0 Or wRot(j) <> 0) Then hit = OrientedOverlap(i, j)
            If hit Then
                res.Add wName(i) & " and " & wName(j) & " overlap by " & Round(ov(0), 2) & " x " & _
                        Round(ov(1), 2) & " x " & Round(ov(2), 2) & " mm"
                ovlList.Add Array(wKey(i), wKey(j), ov(0), ov(1), ov(2))
            End If
        Next
    Next
    Set FindOverlaps = res
End Function

' Footprints (rectangles, turned by wRot) of two wanted parts overlap in XY. Separating axis test, 0.01 mm margin.
' The envelope boxes overlapped already; this decides whether the real shapes do.
Function OrientedOverlap(ByVal i As Long, ByVal j As Long) As Boolean
    Const PI As Double = 3.14159265358979
    Dim cx(1) As Double, cy(1) As Double, hx(1) As Double, hy(1) As Double, an(1) As Double
    Dim m As Long, w As Long, q As Long, axX As Double, axY As Double, rr(1) As Double, d As Double
    For m = 0 To 1
        w = IIf(m = 0, i, j)
        cx(m) = wBox(w, 0) + wBox(w, 3) / 2: cy(m) = wBox(w, 1) + wBox(w, 4) / 2
        If wRot(w) <> 0 Then
            hx(m) = wOwnL(w) / 2: hy(m) = wOwnW(w) / 2: an(m) = wRot(w) * PI / 180
        Else
            hx(m) = wBox(w, 3) / 2: hy(m) = wBox(w, 4) / 2: an(m) = 0
        End If
    Next
    OrientedOverlap = True
    For m = 0 To 1
        For q = 0 To 1
            If q = 0 Then axX = Cos(an(m)): axY = Sin(an(m)) Else axX = -Sin(an(m)): axY = Cos(an(m))
            For w = 0 To 1
                rr(w) = hx(w) * Abs(Cos(an(w)) * axX + Sin(an(w)) * axY) + hy(w) * Abs(-Sin(an(w)) * axX + Cos(an(w)) * axY)
            Next
            d = Abs((cx(1) - cx(0)) * axX + (cy(1) - cy(0)) * axY)
            If d >= rr(0) + rr(1) - 0.01 Then OrientedOverlap = False: Exit Function
        Next
    Next
End Function

Sub WriteReport(ByVal rptPath As String, ByVal asmPath As String, ByVal overlaps As Collection)
    Dim fso As Object, ts As Object, it As Variant
    On Error GoTo done
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(rptPath, True)
    ts.WriteLine "BuildBoard report " & Format(Now, "yyyy-mm-dd hh:nn")
    ts.WriteLine "Assembly: " & asmPath
    ts.WriteLine IIf(isUpdate, "Mode: update", "Mode: create")
    ts.WriteLine "Rotated footprints: " & nRotated & " of " & nW & " wanted parts"
    ts.WriteLine "Parts wanted: " & nW & "  added " & nAdded & "  moved " & nMoved & "  resized " & nResized & _
                 "  suppressed " & nSupp & "  brought back " & nRestored & "  failed " & failMsgs.Count
    ts.WriteLine ""
    ts.WriteLine "FAILED (" & failMsgs.Count & ")"
    For Each it In failMsgs: ts.WriteLine "  " & it: Next
    ts.WriteLine ""
    ts.WriteLine "INCLUDED BODIES SKIPPED (" & skippedRows.Count & ")"
    For Each it In skippedRows: ts.WriteLine "  " & it: Next
    ts.WriteLine ""
    ts.WriteLine "PARTS CUT AT THE BOARD SURFACE (" & clipped.Count & ")"
    ts.WriteLine "  Through-hole and edge connectors reach into the PCB. Their box starts at the board surface instead,"
    ts.WriteLine "  so the simulation has no overlapping bodies. The imported size is kept."
    For Each it In clipped: ts.WriteLine "  " & it: Next
    ts.WriteLine ""
    ts.WriteLine "OVERLAPPING BOXES (" & overlaps.Count & ")"
    ts.WriteLine "  Usually a wrong override or two bodies for one part. Fix them before SetupSim (an override,"
    ts.WriteLine "  or leave one of them out)."
    For Each it In overlaps: ts.WriteLine "  " & it: Next
    ts.Close
done:
End Sub

Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function

Function Max2(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Max2 = a Else Max2 = b
End Function
