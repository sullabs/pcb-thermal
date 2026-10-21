Option Strict Off
Option Explicit On

Imports System
Imports Microsoft.VisualBasic

' Port of src/BuildBoard.bas (milestone M3). Same behaviour and messages as the VBA macro.
'
' Builds the thermal assembly from the CAD geometry tab: one box part for the PCB and one box part per included body
' (used position and size). No thermal assembly yet: creates it and every part. Assembly already there: updates it
' (moves and resizes parts that changed, adds new ones, suppresses parts no longer included) and logs every change.
' Parts are fixed at coordinates (no mates) and carry custom properties Role, Key, Name, RefDes, SourceRow and size.
' Named faces: PCB_Top, PCB_Bottom on the PCB; Chip_<RefDes>_Top (or Top) and Bottom on parts.
' Refuses to run while the Change log has unchecked rows.
Public Module BuildBoard

    Private Const MACRO_NAME As String = "BuildBoard"
    Private Const swDefaultTemplatePart As Integer = 8
    Private Const swDefaultTemplateAssembly As Integer = 9

    ' CAD geometry columns, as offsets from the first column of Geo_Table (column B)
    Private Const G_NAMEUSED As Integer = 3      ' E
    Private Const G_REF As Integer = 4           ' F
    Private Const G_INC As Integer = 5           ' G
    Private Const G_USED As Integer = 19         ' U..Z used Xmin Ymin Zmin L W H
    Private Const G_KEY As Integer = 31          ' AG
    Private Const G_ROT As Integer = 33          ' AI imported rotation (deg), AJ own L, AK own W (added by ImportStep)
    Private Const G_OWNL As Integer = 34
    Private Const G_OWNW As Integer = 35

    ' wanted parts, from Excel
    Private wKey() As String, wName() As String, wRef() As String, wRow() As Integer, wBox(,) As Double, nW As Integer
    Private wRot() As Double, wOwnL() As Double, wOwnW() As Double       ' rotated footprint per wanted part (0 = plain box)
    Private nRotated As Integer
    Private skippedRows As Collection, failMsgs As Collection, clipped As Collection
    Private isUpdate As Boolean
    Private nAdded As Integer, nMoved As Integer, nResized As Integer, nSupp As Integer, nRestored As Integer
    Private geoRow0 As Integer, nColored As Integer, nColorFail As Integer     ' sheet row of the first Geo_Table row; colour counts
    Private noPartHeat As String                 ' RefDes with heat but no included part

    Public Sub Main()
        PT_LogReset()
        skippedRows = New Collection()
        failMsgs = New Collection()
        clipped = New Collection()
        nAdded = 0 : nMoved = 0 : nResized = 0 : nSupp = 0 : nRestored = 0 : nColored = 0 : nColorFail = 0 : nW = 0
        isUpdate = False

        ' ---------- workbook ----------
        Dim xl As Object, wb As Object
        xl = PT_GetExcel()
        wb = PT_PickWorkbook(xl, MACRO_NAME, "CAD geometry|Settings|Change log")
        If wb Is Nothing Then Exit Sub
        If Not PT_NeedNames(wb, "Geo_Table|Geo_PCB_Used|Log_Header|Log_Unchecked|Set_Prefix|Set_TolPos|Set_TolSize|" &
                                "Set_PartTemplate|Set_AsmTemplate|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
        If PT_RefuseIfUnchecked(wb, MACRO_NAME) Then Exit Sub

        Dim tolPos As Double, tolSize As Double, prefix As String, outDir As String
        tolPos = PT_NameNum(wb, "Set_TolPos", 0.05)
        tolSize = PT_NameNum(wb, "Set_TolSize", 0.05)
        prefix = PT_SafeFile(PT_Prefix(wb))

        Dim pcbL As Double, pcbW As Double, pcbT As Double, rp As Object
        rp = PT_Rng(wb, "Geo_PCB_Used")
        pcbL = PT_Num(rp.Cells(1, 1).Value, 0)
        pcbW = PT_Num(rp.Cells(2, 1).Value, 0)
        pcbT = PT_Num(rp.Cells(3, 1).Value, 0)
        ReadWanted(wb, pcbT)
        If pcbL <= 0 Or pcbW <= 0 Or pcbT <= 0 Then
            MsgBox("BuildBoard did not run. The PCB size on CAD geometry (Used column, Length X, Width Y, Thickness Z) is missing." & vbCrLf &
                   "Run ImportStep or type an override there.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        If nW = 0 Then
            If MsgBox("No row on CAD geometry has Include = Yes with a size and a key (" & skippedRows.Count & " included rows were skipped)." & vbCrLf &
                      "Only the PCB would be built." & vbCrLf & vbCrLf & "Build it anyway?", vbYesNo + vbQuestion, MACRO_NAME) <> vbYes Then Exit Sub
        End If
        noPartHeat = HeatWithoutPart(wb)

        ' ---------- files and templates ----------
        Dim partTmpl As String, asmTmpl As String, asmPath As String
        outDir = PT_OutputFolder(wb)
        partTmpl = PT_Template(wb, "Set_PartTemplate", swDefaultTemplatePart)
        asmTmpl = PT_Template(wb, "Set_AsmTemplate", swDefaultTemplateAssembly)
        If partTmpl = "" Or asmTmpl = "" Then
            MsgBox("BuildBoard did not run. The " & IIf(partTmpl = "", "part", "assembly") & " template was not found." & vbCrLf &
                   "Set its full path on the Settings tab.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        asmPath = PT_AsmPath(wb)

        ' ---------- assembly ----------
        Dim asm As Object, errs As Integer = 0, warns As Integer = 0, rptPath As String = ""
        On Error GoTo swFail
        asm = PT_OpenAsm(asmPath)
        If asm Is Nothing And System.IO.File.Exists(asmPath) Then
            MsgBox("BuildBoard could not open " & asmPath & ". Nothing was changed.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        isUpdate = asm IsNot Nothing
        If Not isUpdate Then
            asm = swApp.NewDocument(asmTmpl, 0, 0, 0)
            If asm Is Nothing Then
                MsgBox("SolidWorks could not make a new assembly from" & vbCrLf & asmTmpl, vbExclamation, MACRO_NAME)
                Exit Sub
            End If
            If Not SwSaveAs(asm, asmPath, errs, warns) Then
                MsgBox("Could not save the new assembly as" & vbCrLf & asmPath & vbCrLf & "(error " & errs & "). Is the folder writable?",
                       vbExclamation, MACRO_NAME)
                Exit Sub
            End If
        End If
        PT_LeaveHeatOnly(asm)
        PT_SetProp(asm, "Role", "ThermalAssembly")
        PT_SetProp(asm, "Workbook", CStr(wb.FullName))
        PT_ActivateDoc(asm)

        SyncAssembly(asm, partTmpl, outDir, prefix, pcbL, pcbW, pcbT, tolPos, tolSize)

        ColorParts(asm)
        PT_ActivateDoc(asm)
        asm.ForceRebuild3(False)
        asm.GraphicsRedraw2()
        SwSave3(asm, PT_swSaveSilent + PT_swSaveReferenced, errs, warns)
        On Error GoTo 0

        ' ---------- report ----------
        Dim overlaps As Collection
        overlaps = FindOverlaps(pcbL, pcbW, pcbT)
        rptPath = outDir & prefix & "BuildBoard_report.txt"
        WriteReport(rptPath, asmPath, overlaps)

        ' ---------- Change log ----------
        Dim nLog As Integer
        On Error GoTo xlFail
        xl.ScreenUpdating = False
        nLog = PT_LogFlush(wb)
        xl.ScreenUpdating = True
        xl.Calculate()
        Dim chipRows As Integer
        chipRows = PT_FillChipLists(wb)
        wb.Save()
        On Error GoTo 0

        Dim s As String
        s = "Written and saved to:" & vbCrLf & asmPath & vbCrLf & "(parts in the same folder)" & vbCrLf & vbCrLf
        If isUpdate Then
            s = s & "Updated the thermal assembly: " & nAdded & " parts added, " & nMoved & " moved, " & nResized & " resized, " &
                nSupp & " suppressed, " & nRestored & " brought back." & vbCrLf
            If nLog > 0 Then
                s = s & nLog & " rows added to the Change log tab of " & CStr(wb.Name) & ". Check them there." & vbCrLf
            Else
                s = s & "Nothing changed." & vbCrLf
            End If
        Else
            s = s & "Created the thermal assembly: PCB plus " & (nAdded - 1) & " parts." & vbCrLf
        End If
        If skippedRows.Count > 0 Then s = s & skippedRows.Count & " included rows skipped (listed in the report)." & vbCrLf
        If failMsgs.Count > 0 Then s = s & failMsgs.Count & " parts FAILED (listed in the report)." & vbCrLf
        If clipped.Count > 0 Then s = s & clipped.Count & " parts cut at the board surface so they do not go into the PCB (listed in the report)." & vbCrLf
        If chipRows > 0 Then s = s & chipRows & " chip rows written to the Spreader and Sim setup tabs (they were empty). Check the Covered column." & vbCrLf
        s = s & nColored & " parts coloured in SolidWorks (PCB green). The workbook shows the same colours on its own." &
            IIf(nColorFail > 0, " " & nColorFail & " could not be coloured.", "") & vbCrLf
        If nRotated > 0 Then s = s & nRotated & " of the wanted parts are rotated rectangles (turned in the board plane)." & vbCrLf
        If noPartHeat <> "" Then s = s & "Heat in the register but no built part for: " & noPartHeat & ". Set RefDes and Include = Yes on CAD geometry, or they get no heat source." & vbCrLf
        s = s & overlaps.Count & " overlapping pairs of boxes left (listed in the report). Fix them before SetupSim." & vbCrLf & vbCrLf &
            "Report: " & rptPath
        MsgBox(s, IIf(failMsgs.Count > 0, vbExclamation, vbInformation), MACRO_NAME)
        Exit Sub

swFail:
        MsgBox("BuildBoard stopped in SolidWorks: " & Err.Description & vbCrLf &
               "The assembly may be half done and was not saved. Close it without saving, send this message, and run again.",
               vbExclamation, MACRO_NAME)
        Exit Sub
xlFail:
        xl.ScreenUpdating = True
        MsgBox("The assembly was saved to " & asmPath & ", but writing the Change log failed: " & Err.Description & vbCrLf &
               "Close any cell you are editing in Excel. The changes are listed in " & rptPath, vbExclamation, MACRO_NAME)
    End Sub

    ' ================================================================= Excel side
    Private Sub ReadWanted(ByVal wb As Object, ByVal pcbT As Double)
        Dim rg As Object, v(,) As Object, ok As Boolean, rowName As String, bx(5) As Double
        Dim hasRot As Boolean, rotV As Double, ownLV As Double, ownWV As Double, hOld As Double
        Dim used As New Dict()
        rg = PT_Rng(wb, "Geo_Table")
        geoRow0 = CInt(rg.Row)
        v = ToTable(rg.Value)
        nW = 0
        Dim nr As Integer = UBound(v, 1)
        ReDim wKey(nr) : ReDim wName(nr) : ReDim wRef(nr)
        ReDim wRow(nr) : ReDim wBox(nr, 5)
        ReDim wRot(nr) : ReDim wOwnL(nr) : ReDim wOwnW(nr)
        nRotated = 0
        hasRot = (UBound(v, 2) >= 1 + G_OWNW)
        For r As Integer = 1 To nr
            rowName = PT_Str(v(r, 1 + G_NAMEUSED))
            If rowName <> "" And UCase(PT_Str(v(r, 1 + G_INC))) = "YES" Then
                ok = True
                For k As Integer = 0 To 5
                    If Not PT_IsNum(v(r, 1 + G_USED + k)) Then ok = False
                Next
                If ok Then
                    If PT_Num(v(r, 1 + G_USED + 3), 0) <= 0 Or PT_Num(v(r, 1 + G_USED + 4), 0) <= 0 Or
                       PT_Num(v(r, 1 + G_USED + 5), 0) <= 0 Then ok = False
                End If
                If Not ok Then
                    skippedRows.Add("Row " & (CInt(rg.Row) + r - 1) & " " & rowName & ": position or size missing or zero")
                ElseIf PT_Str(v(r, 1 + G_KEY)) = "" Then
                    skippedRows.Add("Row " & (CInt(rg.Row) + r - 1) & " " & rowName & ": no key (run SetupWorkbook)")
                Else
                    wKey(nW) = PT_UniqueKey(used, PT_Str(v(r, 1 + G_KEY)))
                    wName(nW) = rowName
                    wRef(nW) = PT_Str(v(r, 1 + G_REF))
                    wRow(nW) = CInt(rg.Row) + r - 1
                    For k As Integer = 0 To 5
                        bx(k) = CDbl(v(r, 1 + G_USED + k))
                    Next
                    If pcbT > 0 Then
                        hOld = CDbl(v(r, 1 + G_USED + 5))
                        ' a part that sinks less than 0.05 mm is STEP rounding: cut it quietly, list only the real ones
                        If PT_ClipToBoard(bx, pcbT) Then
                            If hOld - bx(5) > 0.05 Then clipped.Add(rowName & ": box cut at the board surface (" &
                                                                    Round(hOld - bx(5), 3) & " mm went into the PCB)")
                        End If
                    End If
                    For k As Integer = 0 To 5
                        wBox(nW, k) = bx(k)
                    Next
                    wRot(nW) = 0 : wOwnL(nW) = 0 : wOwnW(nW) = 0
                    ' rotated footprint: used only when the box was not overridden in L or W (envelope still fits)
                    If hasRot Then
                        If PT_IsNum(v(r, 1 + G_ROT)) And PT_IsNum(v(r, 1 + G_OWNL)) And PT_IsNum(v(r, 1 + G_OWNW)) Then
                            rotV = CDbl(v(r, 1 + G_ROT)) : ownLV = CDbl(v(r, 1 + G_OWNL)) : ownWV = CDbl(v(r, 1 + G_OWNW))
                            If PT_ShapeIsRotated(rotV, ownLV, ownWV, bx(3), bx(4)) Then
                                wRot(nW) = rotV : wOwnL(nW) = ownLV : wOwnW(nW) = ownWV
                                nRotated = nRotated + 1
                            End If
                        End If
                    End If
                    nW = nW + 1
                End If
            End If
        Next
    End Sub

    ' ================================================================= SolidWorks side
    Private Sub SyncAssembly(ByVal asm As Object, ByVal partTmpl As String, ByVal outDir As String, ByVal prefix As String,
                             ByVal pcbL As Double, ByVal pcbW As Double, ByVal pcbT As Double, ByVal tolPos As Double, ByVal tolSize As Double)
        Dim have As New Dict(), wasSupp As New Dict(), want As New Dict(), pcbHave As New Dict(), pcbSupp As New Dict()
        Dim pcbComp As Object, res As Integer, fpath As String, okAdd As Boolean
        For i As Integer = 0 To nW - 1
            want.Add(wKey(i), i)
        Next
        PT_ShapeClear()

        pcbComp = PT_ReadTagged(asm, "PCB", pcbHave, pcbSupp)
        PT_ReadTagged(asm, "Part", have, wasSupp)

        ' ---- PCB
        If pcbComp Is Nothing Then
            If PT_AddBoxPart(asm, partTmpl, outDir & prefix & "PCB.sldprt", "PCB", "pcb", "PCB", "", 0, 0, 0, 0,
                             pcbL, pcbW, pcbT, "PCB_Top", "PCB_Bottom", MACRO_NAME) Then
                nAdded = nAdded + 1
                If isUpdate Then PT_LogAdd(MACRO_NAME, "pcb", "PCB", "", "added", "Box", "", PT_BoxText(0, 0, 0, pcbL, pcbW, pcbT))
            Else
                failMsgs.Add("Could not add the PCB: " & PT_LastErr)
            End If
        Else
            res = PT_CheckBoxPart(asm, pcbComp, "PCB", "pcb", "PCB", "", 0, 0, 0, 0, pcbL, pcbW, pcbT,
                                  "PCB_Top", "PCB_Bottom", tolPos, tolSize, MACRO_NAME, isUpdate)
            CountResult(res, "PCB")
        End If

        ' ---- parts
        For i As Integer = 0 To nW - 1
            If have.Exists(wKey(i)) Then
                If CBool(wasSupp(wKey(i))) Then
                    nRestored = nRestored + 1
                    If isUpdate Then PT_LogAdd(MACRO_NAME, wKey(i), wName(i), wRef(i), "added", "Brought back (was suppressed)", "", "")
                End If
                PT_ShapeRot = wRot(i) : PT_ShapeL = wOwnL(i) : PT_ShapeW = wOwnW(i)
                res = PT_CheckBoxPart(asm, have(wKey(i)), "Part", wKey(i), wName(i), wRef(i), wRow(i),
                                      wBox(i, 0), wBox(i, 1), wBox(i, 2), wBox(i, 3), wBox(i, 4), wBox(i, 5),
                                      TopName(wRef(i)), "Bottom", tolPos, tolSize, MACRO_NAME, isUpdate)
                PT_ShapeClear()
                CountResult(res, wName(i))
            Else
                fpath = outDir & prefix & "P_" & PT_SafeFile(wKey(i)) & ".sldprt"
                PT_ShapeRot = wRot(i) : PT_ShapeL = wOwnL(i) : PT_ShapeW = wOwnW(i)
                okAdd = PT_AddBoxPart(asm, partTmpl, fpath, "Part", wKey(i), wName(i), wRef(i), wRow(i),
                                      wBox(i, 0), wBox(i, 1), wBox(i, 2), wBox(i, 3), wBox(i, 4), wBox(i, 5),
                                      TopName(wRef(i)), "Bottom", MACRO_NAME)
                PT_ShapeClear()
                If okAdd Then
                    nAdded = nAdded + 1
                    If isUpdate Then PT_LogAdd(MACRO_NAME, wKey(i), wName(i), wRef(i), "added", "Box", "",
                                               PT_BoxText(wBox(i, 0), wBox(i, 1), wBox(i, 2), wBox(i, 3), wBox(i, 4), wBox(i, 5)))
                Else
                    failMsgs.Add("Could not add " & wName(i) & " (" & wKey(i) & "): " & PT_LastErr)
                End If
            End If
        Next

        ' ---- parts no longer wanted: suppress, never delete
        nSupp = PT_SuppressUnwanted(have, wasSupp, want, MACRO_NAME, isUpdate)
    End Sub

    ' Colours the PCB green and every wanted part with its palette colour (palette index = its row on CAD geometry).
    Private Sub ColorParts(ByVal asm As Object)
        Dim idx As New Dict(), rgbL As Integer, ok As Boolean
        For i As Integer = 0 To nW - 1
            idx(wKey(i)) = wRow(i) - geoRow0
        Next
        For Each it As Object In PT_ListTagged(asm)
            ok = False
            If CStr(it(0)) = "PCB" Then
                rgbL = PT_PcbRgb() : ok = True
            ElseIf CStr(it(0)) = "Part" Then
                If idx.Exists(it(1)) Then
                    rgbL = PT_PartRgb(CInt(idx(it(1)))) : ok = True
                End If
            End If
            If ok Then
                If PT_ApplyColor(it(3), rgbL) Then nColored = nColored + 1 Else nColorFail = nColorFail + 1
            End If
        Next
    End Sub

    ' RefDes with heat in the register that no included part on CAD geometry carries. SetupSim would give them no heat source.
    Private Function HeatWithoutPart(ByVal wb As Object) As String
        Dim rr As Object, rq As Object, refs(,) As Object, qs(,) As Object, have As New Dict(), s As String = "", cnt As Integer = 0
        rr = PT_Rng(wb, "Reg_RefDes") : rq = PT_Rng(wb, "Reg_QUsed")
        If rr Is Nothing Or rq Is Nothing Then Return ""
        For i As Integer = 0 To nW - 1
            If wRef(i) <> "" Then have(UCase(wRef(i))) = 1
        Next
        refs = ToTable(rr.Value) : qs = ToTable(rq.Value)
        For r As Integer = 1 To UBound(refs, 1)
            If PT_Str(refs(r, 1)) <> "" And PT_Num(qs(r, 1), 0) > 0 Then
                If Not have.Exists(UCase(PT_Str(refs(r, 1)))) Then
                    cnt = cnt + 1
                    If cnt <= 8 Then s = s & IIf(s = "", "", ", ") & PT_Str(refs(r, 1))
                End If
            End If
        Next
        If cnt > 8 Then s = s & " and " & (cnt - 8) & " more"
        Return s
    End Function

    Private Sub CountResult(ByVal res As Integer, ByVal what As String)
        If res < 0 Then
            failMsgs.Add("Could not update " & what & ": " & PT_LastErr)
        Else
            If (res And 1) <> 0 Then nResized = nResized + 1
            If (res And 2) <> 0 Then nMoved = nMoved + 1
        End If
    End Sub

    Private Function TopName(ByVal refD As String) As String
        If refD <> "" Then Return "Chip_" & PT_SafeFile(refD) & "_Top"
        Return "Top"
    End Function

    ' ================================================================= checks and report
    ' Pairs of boxes (parts and PCB) that overlap by more than 0.01 mm in X, Y and Z.
    Private Function FindOverlaps(ByVal pcbL As Double, ByVal pcbW As Double, ByVal pcbT As Double) As Collection
        Dim res As New Collection(), ov(2) As Double, hit As Boolean
        Dim a0 As Double, a1 As Double, b0 As Double, b1 As Double, pcbBox(2) As Double
        pcbBox(0) = pcbL : pcbBox(1) = pcbW : pcbBox(2) = pcbT
        For i As Integer = 0 To nW - 1
            hit = True
            For k As Integer = 0 To 2
                a0 = wBox(i, k) : a1 = wBox(i, k) + wBox(i, k + 3)
                ov(k) = Min2(a1, pcbBox(k)) - Max2(a0, 0)
                If ov(k) <= 0.01 Then hit = False
            Next
            If hit Then res.Add(wName(i) & " goes into the PCB by " & Round(ov(2), 2) & " mm (overlap " &
                                Round(ov(0), 2) & " x " & Round(ov(1), 2) & " mm)")
            For j As Integer = i + 1 To nW - 1
                hit = True
                For k As Integer = 0 To 2
                    a0 = wBox(i, k) : a1 = wBox(i, k) + wBox(i, k + 3)
                    b0 = wBox(j, k) : b1 = wBox(j, k) + wBox(j, k + 3)
                    ov(k) = Min2(a1, b1) - Max2(a0, b0)
                    If ov(k) <= 0.01 Then hit = False
                Next
                If hit And (wRot(i) <> 0 Or wRot(j) <> 0) Then hit = OrientedOverlap(i, j)
                If hit Then res.Add(wName(i) & " and " & wName(j) & " overlap by " & Round(ov(0), 2) & " x " &
                                    Round(ov(1), 2) & " x " & Round(ov(2), 2) & " mm")
            Next
        Next
        Return res
    End Function

    ' Footprints (rectangles, turned by wRot) of two wanted parts overlap in XY. Separating axis test, 0.01 mm margin.
    ' The envelope boxes overlapped already; this decides whether the real shapes do.
    Private Function OrientedOverlap(ByVal i As Integer, ByVal j As Integer) As Boolean
        Const PI As Double = 3.14159265358979
        Dim cx(1) As Double, cy(1) As Double, hx(1) As Double, hy(1) As Double, an(1) As Double
        Dim w As Integer, axX As Double, axY As Double, rr(1) As Double, d As Double
        For m As Integer = 0 To 1
            w = IIf(m = 0, i, j)
            cx(m) = wBox(w, 0) + wBox(w, 3) / 2 : cy(m) = wBox(w, 1) + wBox(w, 4) / 2
            If wRot(w) <> 0 Then
                hx(m) = wOwnL(w) / 2 : hy(m) = wOwnW(w) / 2 : an(m) = wRot(w) * PI / 180
            Else
                hx(m) = wBox(w, 3) / 2 : hy(m) = wBox(w, 4) / 2 : an(m) = 0
            End If
        Next
        For m As Integer = 0 To 1
            For q As Integer = 0 To 1
                If q = 0 Then
                    axX = Math.Cos(an(m)) : axY = Math.Sin(an(m))
                Else
                    axX = -Math.Sin(an(m)) : axY = Math.Cos(an(m))
                End If
                For w = 0 To 1
                    rr(w) = hx(w) * Math.Abs(Math.Cos(an(w)) * axX + Math.Sin(an(w)) * axY) + hy(w) * Math.Abs(-Math.Sin(an(w)) * axX + Math.Cos(an(w)) * axY)
                Next
                d = Math.Abs((cx(1) - cx(0)) * axX + (cy(1) - cy(0)) * axY)
                If d >= rr(0) + rr(1) - 0.01 Then Return False
            Next
        Next
        Return True
    End Function

    Private Sub WriteReport(ByVal rptPath As String, ByVal asmPath As String, ByVal overlaps As Collection)
        Try
            Using ts As System.IO.StreamWriter = NewTextFile(rptPath)
                ts.WriteLine("BuildBoard report " & NowText())
                ts.WriteLine("Assembly: " & asmPath)
                ts.WriteLine(IIf(isUpdate, "Mode: update", "Mode: create"))
                ts.WriteLine("Rotated footprints: " & nRotated & " of " & nW & " wanted parts")
                ts.WriteLine("Parts wanted: " & nW & "  added " & nAdded & "  moved " & nMoved & "  resized " & nResized &
                             "  suppressed " & nSupp & "  brought back " & nRestored & "  failed " & failMsgs.Count)
                ts.WriteLine("")
                ts.WriteLine("FAILED (" & failMsgs.Count & ")")
                For Each it As Object In failMsgs
                    ts.WriteLine("  " & CStr(it))
                Next
                ts.WriteLine("")
                ts.WriteLine("INCLUDED ROWS SKIPPED (" & skippedRows.Count & ")")
                For Each it As Object In skippedRows
                    ts.WriteLine("  " & CStr(it))
                Next
                ts.WriteLine("")
                ts.WriteLine("PARTS CUT AT THE BOARD SURFACE (" & clipped.Count & ")")
                ts.WriteLine("  Through-hole and edge connectors reach into the PCB. Their box starts at the board surface instead,")
                ts.WriteLine("  so the simulation has no overlapping bodies. CAD geometry keeps the full size.")
                For Each it As Object In clipped
                    ts.WriteLine("  " & CStr(it))
                Next
                ts.WriteLine("")
                ts.WriteLine("OVERLAPPING BOXES (" & overlaps.Count & ")")
                ts.WriteLine("  Usually a wrong override or two bodies for one part. Fix them before SetupSim (override on CAD geometry,")
                ts.WriteLine("  or Include = No for one of them).")
                For Each it As Object In overlaps
                    ts.WriteLine("  " & CStr(it))
                Next
            End Using
        Catch
        End Try
    End Sub

    Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
        If a < b Then Return a
        Return b
    End Function

    Private Function Max2(ByVal a As Double, ByVal b As Double) As Double
        If a > b Then Return a
        Return b
    End Function

End Module
