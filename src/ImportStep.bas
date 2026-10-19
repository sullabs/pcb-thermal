Attribute VB_Name = "ImportStep"
Option Explicit
' =====================================================================================
' ImportStep  -  SolidWorks 2023 macro
'
' Reads a PCB assembly (a STEP file, or the model already open in SolidWorks) and
' writes every body's size, position and rotation. Works for any board.
' Needs modules PTCommon and PTJson in the same macro.
'
' Where it writes (asked at the start, the last one is the default):
'   project.json picked (board folder of the page) -> cad.json in that folder: import, importPrevious, changes.
'       Settings come from project.json (settings.import, tolerances). project.json itself is never written:
'       the page keeps include, RefDes and overrides by body key, and new bodies start not included.
'   a workbook picked (copy of PCB_thermal_TEMPLATE.xlsx) -> tabs STEP import, CAD geometry, Change log (below).
'
' What is imported (setting Imp_TopOnly on the STEP import tab):
'   No  = EVERY solid body gets its own row (default; small parts are not lost inside bigger ones)
'   Yes = one row per top-level component (old behaviour, bodies of a component are merged into one box)
'
' Board frame written to the sheet (mm):
'   origin at the PCB corner on its bottom face, X along the longest board edge,
'   Y along the other edge, Z out of the side that carries most of the parts.
'   If the board is turned in the model file (not square to the model axes), the frame is turned with it,
'   so board X and Y follow the board edges.
'
' Rotation of each body:
'   The shape of every body is measured from its surface points. If the body is turned in the board plane
'   (for example a part at 30 degrees), the row also gets Rot (deg) and the size of the turned rectangle
'   (Own L, Own W). Xmin, Ymin, L, W stay the axis-aligned envelope of that turned rectangle, so every
'   other macro keeps working. BuildBoard builds the turned rectangle, not the big envelope box.
'
' Include: new rows get Include = No. Set Yes on CAD geometry for the parts you want built.
' The first run on a workbook adds the Rot columns (STEP import N:P, CAD geometry AI:AK) and grows the
' body tables when the model has more bodies than rows (Excel moves all formulas that point to them).
'
' Settings (STEP import tab, names Imp_PcbName, Imp_MinSize, Imp_TopOnly):
'   text a body name must contain to be the PCB           (default PCB)
'   skip bodies whose L and W are both below this (mm)    (default 0 = keep all)
'
' Change report: before writing, the last import is copied to the hidden tab
' 'STEP import (previous)'. Every body added, removed, moved, resized or renamed since
' then gets a row on the Change log tab (tolerances Set_TolPos, Set_TolSize).
' When the import level changes (components -> bodies) the report is skipped once: every row is new.
'
' Your inputs on 'CAD geometry' (name override, RefDes, Include, overrides) follow
' their body by key, so re-importing a changed STEP does not shift them to other rows.
' =====================================================================================

Const MACRO_NAME As String = "ImportStep"
Const MAX_ITEMS As Long = 5000
Const swSolidBody As Long = 0
Const swSheetBody As Long = 1
Const swDocPART As Long = 1
Const swDocASSEMBLY As Long = 2
Const PI As Double = 3.14159265358979
Const NBINS As Long = 256

' CAD geometry columns, as offsets from the first column of Geo_Table (column B)
Const G_NAME As Long = 1          ' C imported name
Const G_OVNAME As Long = 2        ' D name override
Const G_REF As Long = 4           ' F RefDes
Const G_INC As Long = 5           ' G Include
Const G_IMP As Long = 6           ' H..M imported box
Const G_OVR As Long = 13          ' O..T overrides
Const G_ROT As Long = 33          ' AI imported rotation (deg), AJ own L, AK own W
Const G_ROTCOLS As Long = 3

Dim nm() As String
Dim bx() As Double          ' (item, 0..5) = xmin ymin zmin xmax ymax zmax, model space, mm
Dim rbx() As Double         ' same, but in the part's own space (before the component transform), mm
Dim srcComp() As Object     ' owning component (Nothing for a part document)
Dim srcBody() As Object     ' the body (Nothing when the item is a whole component)
Dim srcMode() As Long       ' how to turn part-space points into model space: 0 none, 1 matrix, 2 transposed matrix
Dim srcXf() As Double       ' (item, 0..12) component transform: 9 rotation, 3 translation (m), scale
Dim n As Long
Dim msgs As Collection
Dim msgLvl As Collection    ' level of each message: info or warning (cad.json messages)
Dim capRows As Long         ' body rows available on CAD geometry
' the model in board coordinates (Measure), read by both writers
Dim mOut() As Variant       ' rows (0 = PCB): 0 #, 1 name, 2..4 min, 5..7 size, 8 side, 9 key, 10 rot, 11 own L, 12 own W
Dim mCnt As Long            ' rows in mOut
Dim mFr(8) As Double        ' board frame: board X, Y, Z as unit vectors in model space
Dim mPcbL As Double, mPcbW As Double, mPcbT As Double
Dim mTurned As Long         ' bodies turned in the board plane
Dim nSheetBodies As Long    ' surface bodies seen (not imported)
Dim itemsCapped As Boolean

Sub main()
    Set swApp = Application.SldWorks
    PT_Trace MACRO_NAME & ": started (VBA macro)."
    Set msgs = New Collection
    Set msgLvl = New Collection
    PT_LogReset
    nSheetBodies = 0: itemsCapped = False

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
    Dim xl As Object, wb As Object, wsI As Object, wsG As Object
    Set xl = PT_GetExcel()
    Set wb = PT_PickWorkbook(xl, MACRO_NAME, "STEP import|CAD geometry", PJ_WbPath)
    If wb Is Nothing Then Exit Sub
    If Not PT_NeedNames(wb, "Imp_Table|Imp_PcbName|Imp_MinSize|Imp_TopOnly|Imp_Model|Imp_Time|Imp_Frame|Imp_Messages|" & _
                            "ImpPrev_Table|ImpPrev_Model|ImpPrev_Time|Geo_Table|Log_Header|Set_TolPos|Set_TolSize", MACRO_NAME) Then Exit Sub

    Dim upg As String
    upg = UpgradeWorkbook(wb, xl)
    If Left(upg, 4) = "ERR:" Then
        PT_Msg "ImportStep did not run. The workbook could not be prepared for rotation columns:" & vbCrLf & Mid(upg, 5) & vbCrLf & _
               "Close any cell you are editing in Excel and run again.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    If upg <> "" Then Note "info", upg

    Dim rImp As Object, rGeo As Object, rMsg As Object, rRot As Object
    Set rImp = PT_Rng(wb, "Imp_Table")
    Set rGeo = PT_Rng(wb, "Geo_Table")
    Set rMsg = PT_Rng(wb, "Imp_Messages")
    Set rRot = PT_Rng(wb, "Imp_Rot")
    Set wsI = rImp.Worksheet
    Set wsG = rGeo.Worksheet

    Dim pcbKey As String, minSize As Double, perBody As Boolean, tolPos As Double, tolSize As Double
    pcbKey = PT_NameStr(wb, "Imp_PcbName", "PCB")
    minSize = PT_NameNum(wb, "Imp_MinSize", 0)
    perBody = (UCase(PT_NameStr(wb, "Imp_TopOnly", "Yes")) = "NO")
    tolPos = PT_NameNum(wb, "Set_TolPos", 0.05)
    tolSize = PT_NameNum(wb, "Set_TolSize", 0.05)

    ' ---------- model ----------
    Dim doc As Object, srcName As String
    Set doc = GetModel(srcName, "")
    If doc Is Nothing Then Exit Sub
    If Not Measure(doc, srcName, pcbKey, minSize, perBody, xl) Then Exit Sub

    ' ---------- room for the bodies: grow CAD geometry and STEP import when needed ----------
    Dim added As Long, k As Long
    added = EnsureCapacity(wb, xl, mCnt - 1)
    If added < 0 Then
        PT_Msg "ImportStep stopped while making room for " & (mCnt - 1) & " bodies on the CAD geometry tab:" & vbCrLf & PT_LastErr & vbCrLf & _
               "Nothing else was written. Close any cell you are editing in Excel and run again.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    If added > 0 Then Note "info", "CAD geometry and STEP import were extended by " & added & " body rows (Excel moved every formula that points to them)."
    Set rImp = PT_Rng(wb, "Imp_Table")
    Set rGeo = PT_Rng(wb, "Geo_Table")
    Set rRot = PT_Rng(wb, "Imp_Rot")
    Set wsI = rImp.Worksheet
    Set wsG = rGeo.Worksheet
    capRows = BodyCapacity(rGeo)
    If mCnt - 1 > capRows Or mCnt > rImp.Rows.Count Then
        PT_Msg CStr(mCnt - 1) & " bodies to import; CAD geometry holds " & capRows & "." & vbCrLf & "Nothing was written.", vbExclamation, MACRO_NAME
        Exit Sub
    End If

    ' ---------- save current CAD geometry inputs, keyed by body key ----------
    ' rec: 0 imported name, 1 name override, 2 RefDes, 3 Include, 4..9 overrides,
    '      10..15 imported box, used to match by position when names differ
    Dim saved As Object: Set saved = CreateObject("Scripting.Dictionary")
    Dim refByKey As Object: Set refByKey = CreateObject("Scripting.Dictionary")
    Dim geoKeys As Object: Set geoKeys = CreateObject("Scripting.Dictionary")
    Dim coll As Collection
    Dim old As Variant, rec As Variant, key As String, r As Long
    old = rGeo.Resize(capRows, G_OVR + 6).Value        ' 1-based, column 1 = #
    For r = 1 To UBound(old, 1)
        If IsError(old(r, 1 + G_NAME)) Then
        ElseIf PT_Str(old(r, 1 + G_NAME)) <> "" Then
            rec = Array(CStr(old(r, 1 + G_NAME)), old(r, 1 + G_OVNAME), old(r, 1 + G_REF), old(r, 1 + G_INC), _
                        old(r, 1 + G_OVR), old(r, 2 + G_OVR), old(r, 3 + G_OVR), old(r, 4 + G_OVR), old(r, 5 + G_OVR), old(r, 6 + G_OVR), _
                        old(r, 1 + G_IMP), old(r, 2 + G_IMP), old(r, 3 + G_IMP), old(r, 4 + G_IMP), old(r, 5 + G_IMP), old(r, 6 + G_IMP))
            key = PT_UniqueKey(geoKeys, PT_Key(CStr(old(r, 1 + G_NAME))))
            If Not saved.Exists(key) Then saved.Add key, New Collection
            Set coll = saved(key)
            coll.Add rec
            If Not refByKey.Exists(key) Then refByKey.Add key, PT_Str(old(r, 1 + G_REF))
        End If
    Next

    ' ---------- change report: last import vs this one ----------
    Dim oldImp As Variant, oldModel As Variant, oldTime As Variant
    oldImp = rImp.Value
    oldModel = PT_Rng(wb, "Imp_Model").Value
    oldTime = PT_Rng(wb, "Imp_Time").Value
    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")
    Dim hadOld As Boolean, newLevel As String, oldLevel As String, levelChanged As Boolean, hadRows As Boolean
    newLevel = IIf(perBody, "bodies", "components")
    oldLevel = GetLevelUsed(wb)
    For r = 2 To UBound(oldImp, 1)
        If PT_Str(oldImp(r, 2)) <> "" Then hadRows = True: Exit For
    Next
    If oldLevel = "" And hadRows Then oldLevel = "components"       ' workbooks from before this version
    levelChanged = (hadRows And oldLevel <> newLevel)
    If levelChanged Then
        Note "warning", "Import level changed from " & oldLevel & " to " & newLevel & ": the change report was skipped once, because nearly every row is new. " & _
                        "Inputs on CAD geometry were kept by name or position where the body still exists."
    Else
        hadOld = CompareImports(oldImp, mOut, mCnt, refByKey, tolPos, tolSize, counts)
    End If

    ' ---------- write ----------
    On Error GoTo writeFail
    xl.ScreenUpdating = False
    PT_Rng(wb, "ImpPrev_Table").Value = oldImp
    PT_Rng(wb, "ImpPrev_Model").Value = oldModel
    PT_Rng(wb, "ImpPrev_Time").Value = oldTime

    rImp.ClearContents
    rMsg.ClearContents
    If Not rRot Is Nothing Then rRot.ClearContents
    Dim outW() As Variant, outR() As Variant
    ReDim outW(1 To mCnt, 1 To rImp.Columns.Count)
    For r = 0 To mCnt - 1
        For k = 0 To 9
            If k < rImp.Columns.Count Then outW(r + 1, k + 1) = mOut(r, k)
        Next
    Next
    rImp.Resize(mCnt, rImp.Columns.Count).Value = outW
    If Not rRot Is Nothing Then
        ReDim outR(1 To mCnt, 1 To G_ROTCOLS)
        For r = 0 To mCnt - 1
            For k = 1 To G_ROTCOLS: outR(r + 1, k) = mOut(r, 9 + k): Next
        Next
        rRot.Resize(mCnt, G_ROTCOLS).Value = outR
    End If
    PT_Rng(wb, "Imp_Model").Value = srcName
    PT_Rng(wb, "Imp_Time").Value = Format(Now, "yyyy-mm-dd hh:nn")
    PT_Rng(wb, "Imp_Frame").Value = FrameText(mFr) & " Origin at the PCB corner on its bottom face. Units mm."
    SetLevelUsed wb, newLevel

    Dim g0 As Long, gc As Long
    g0 = rGeo.Row: gc = rGeo.Column
    wsG.Cells(g0, gc + G_OVNAME).Resize(capRows, 1).ClearContents
    wsG.Cells(g0, gc + G_REF).Resize(capRows, 2).ClearContents
    wsG.Cells(g0, gc + G_OVR).Resize(capRows, 6).ClearContents

    Dim assigned() As Boolean, byName As Long, byBox As Long, hit As Boolean
    Dim kk As Variant, t As Long, s As String
    ReDim assigned(mCnt - 1)
    Dim isNewRow() As Boolean
    ReDim isNewRow(mCnt - 1)
    ' pass 1: same key
    For r = 1 To mCnt - 1
        key = CStr(mOut(r, 9))
        If saved.Exists(key) Then
            Set coll = saved(key)
            If coll.Count > 0 Then
                rec = coll(1): coll.Remove 1
                PutInputs wsG, g0 + r - 1, gc, rec
                assigned(r) = True: byName = byName + 1
            End If
        End If
    Next
    ' pass 2: same box (position and size within 0.05 mm), for names that changed
    For r = 1 To mCnt - 1
        If Not assigned(r) Then
            hit = False
            For Each kk In saved.Keys
                Set coll = saved(kk)
                For t = 1 To coll.Count
                    rec = coll(t)
                    If SameBox(rec, mOut, r) Then
                        coll.Remove t
                        PutInputs wsG, g0 + r - 1, gc, rec
                        assigned(r) = True: byBox = byBox + 1: hit = True
                        Exit For
                    End If
                Next
                If hit Then Exit For
            Next
            If Not hit Then
                wsG.Cells(g0 + r - 1, gc + G_INC).Value = "No"         ' new body: off until you switch it on
                isNewRow(r) = True
            End If
        End If
    Next
    Dim restored As Long: restored = byName + byBox
    If byBox > 0 Then Note "info", byBox & " rows kept their inputs by matching position and size, because the body name changed."

    ' pass 3: empty RefDes cells get the register RefDes whose key equals the body name
    ' (ECAD exports name each body by its designator, e.g. "U1"). Typed RefDes cells are never changed.
    Dim nAuto As Long
    Dim nOn As Long
    nAuto = AutoLinkRefDes(wb, wsG, g0, gc, mOut, mCnt, isNewRow, nOn)
    If nAuto > 0 Then Note "info", nAuto & " RefDes filled on CAD geometry because the body name equals a register RefDes. Check column F."
    If nOn > 0 Then Note "info", nOn & " bodies linked to a register RefDes for the first time were set to Include = Yes (column G)."

    ' inputs whose body is no longer in the model
    For Each kk In saved.Keys
        Set coll = saved(kk)
        For t = 1 To coll.Count
            rec = coll(t)
            If HasInput(rec) Then
                s = "Not in this model, inputs not applied: " & rec(0) & "  [name " & rec(1) & ", RefDes " & rec(2) & _
                    ", Include " & rec(3) & ", overrides " & rec(4) & "/" & rec(5) & "/" & rec(6) & "/" & rec(7) & "/" & rec(8) & "/" & rec(9) & "]"
                Note "warning", s
            End If
        Next
    Next

    Dim nLog As Long
    nLog = PT_LogFlush(wb)
    If Not hadOld And Not levelChanged Then Note "info", "First import into this workbook: nothing to compare, no Change log rows."
    If mTurned > 0 Then Note "info", mTurned & " bodies are turned in the board plane (Rot column on STEP import). BuildBoard builds them as turned rectangles."
    If mCnt > 1 Then Note "info", (mCnt - 1) & " body rows have Include = No unless you had set Yes before. Set Yes on CAD geometry (column G) for the parts to build."

    If msgs.Count = 0 Then Note "info", "No warnings."
    For t = 1 To msgs.Count
        If t <= rMsg.Rows.Count Then rMsg.Cells(t, 1).Value = msgs(t)
    Next
    xl.ScreenUpdating = True
    xl.Calculate
    Dim chipRows As Long
    chipRows = PT_FillChipLists(wb)
    If chipRows > 0 Then Note "info", chipRows & " chip rows written to the Spreader and Sim setup tabs (they were empty)."
    wb.Save

    PT_Msg "Written and saved to:" & vbCrLf & wb.FullName & vbCrLf & "(tabs STEP import, CAD geometry, Change log)" & vbCrLf & vbCrLf & _
           "Imported " & (mCnt - 1) & " bodies plus the PCB (" & Round(mPcbL, 2) & " x " & _
           Round(mPcbW, 2) & " x " & Round(mPcbT, 2) & " mm), level: " & newLevel & "." & vbCrLf & _
           mTurned & " bodies turned in the board plane." & vbCrLf & _
           restored & " rows kept their CAD geometry inputs (" & byName & " by key, " & byBox & " by position)." & vbCrLf & vbCrLf & _
           "Changes since the last import: " & ChangeSummary(counts, nLog, hadOld, levelChanged, _
               " rows added to the Change log tab. Check them there (type Yes under Checked).") & vbCrLf & vbCrLf & _
           "Messages:" & vbCrLf & JoinMsgs("(more on the STEP import tab)"), vbInformation, MACRO_NAME
    Exit Sub
writeFail:
    xl.ScreenUpdating = True
    xl.StatusBar = False
    PT_Msg "Stopped while writing to Excel: " & Err.Description & vbCrLf & _
           "The workbook was not saved. Close any cell you are editing in Excel, close the workbook without saving, and run again.", _
           vbExclamation, MACRO_NAME
End Sub

Sub Note(ByVal lvl As String, ByVal txt As String)
    msgs.Add txt
    msgLvl.Add lvl
End Sub

Sub ShowStatus(ByVal xl As Object, ByVal txt As String)
    If xl Is Nothing Then Exit Sub
    On Error Resume Next
    If txt = "" Then xl.StatusBar = False Else xl.StatusBar = txt
End Sub

' ---------------------------------------------------------------- the model in board coordinates
' Reads the bodies of doc, finds the PCB and the board frame, and measures every body in that frame (mm).
' Fills mOut, mCnt, mFr, mPcbL, mPcbW, mPcbT and mTurned. Both writers (workbook, cad.json) use them.
' False when the model has no solid body (the user has been told).
Function Measure(ByVal doc As Object, ByVal srcName As String, ByVal pcbKey As String, ByVal minSize As Double, _
                 ByVal perBody As Boolean, ByVal xl As Object) As Boolean
    ShowStatus xl, "ImportStep: reading the bodies of " & srcName & " ..."
    CollectItems doc, perBody
    If n = 0 Then
        ShowStatus xl, ""
        PT_Msg "No solid bodies found in " & srcName & ". Nothing was written.", vbExclamation, MACRO_NAME
        Exit Function
    End If
    If itemsCapped Then Note "warning", "The model has more than " & MAX_ITEMS & " bodies; only the first " & MAX_ITEMS & " were read."
    If nSheetBodies > 0 Then Note "info", nSheetBodies & " surface (sheet) bodies were not imported. Only solid bodies are."

    ' ---------- PCB ----------
    Dim pcb As Long, i As Long, k As Long, best As Double, v As Double
    pcb = -1: best = -1
    For i = 0 To n - 1
        If InStr(1, nm(i), pcbKey, vbTextCompare) > 0 Then
            v = BoxVol(i): If v > best Then best = v: pcb = i
        End If
    Next
    If pcb < 0 Then
        For i = 0 To n - 1
            v = Flatness(i): If v > best Then best = v: pcb = i
        Next
        Note "warning", "No body name contains '" & pcbKey & "'. PCB picked by shape: " & nm(pcb) & ". Check it is the board, or change the PCB name setting."
    End If

    ' ---------- board frame: axes ----------
    Dim d(2) As Double, nAx As Long, xAx As Long, midN As Double, up As Double, dn As Double, c As Double
    For k = 0 To 2: d(k) = bx(pcb, k + 3) - bx(pcb, k): Next
    nAx = 0: xAx = 0
    For k = 1 To 2
        If d(k) < d(nAx) Then nAx = k
        If d(k) > d(xAx) Then xAx = k
    Next
    midN = (bx(pcb, nAx) + bx(pcb, nAx + 3)) / 2
    For i = 0 To n - 1
        If i <> pcb Then
            c = (bx(i, nAx) + bx(i, nAx + 3)) / 2
            If c >= midN Then up = up + BoxVol(i) Else dn = dn + BoxVol(i)
        End If
    Next
    Dim ex(2) As Double, ey(2) As Double, ez(2) As Double
    ex(xAx) = 1
    ez(nAx) = IIf(up >= dn, 1, -1)
    ey(0) = ez(1) * ex(2) - ez(2) * ex(1)
    ey(1) = ez(2) * ex(0) - ez(0) * ex(2)
    ey(2) = ez(0) * ex(1) - ez(1) * ex(0)

    Dim fr(8) As Double                       ' frame: board X, Y, Z as unit vectors in model space
    For k = 0 To 2: fr(k) = ex(k): fr(3 + k) = ey(k): fr(6 + k) = ez(k): Next

    ' ---------- board frame: turn it with the PCB if the PCB is not square to the model axes ----------
    Dim sh(10) As Double, shPcb(10) As Double, havePcb As Boolean, turn As Double
    Dim pPts() As Double, pN As Long                ' PCB surface points, read once (the board is the biggest tessellation)
    pN = ItemPoints(pcb, pPts)
    havePcb = (pN >= 9)
    If havePcb Then havePcb = Footprint(pPts, pN, fr, sh)
    If havePcb Then
        If Abs(sh(6)) >= 0.3 Then
            turn = sh(6)
            RotateFrame fr, turn
            havePcb = Footprint(pPts, pN, fr, sh)
        End If
        If havePcb Then
            If sh(8) > sh(7) * 1.0001 Then            ' long edge ended up on Y: turn 90 degrees
                RotateFrame fr, 90
                turn = turn + 90
                havePcb = Footprint(pPts, pN, fr, sh)
            End If
        End If
    End If
    Erase pPts
    For k = 0 To 10: shPcb(k) = sh(k): Next
    If turn <> 0 Then Note "info", "The board is turned " & Round(turn, 2) & " degrees about its normal in the model file. The board frame was turned with it."
    If Not havePcb Then Note "warning", "The PCB surface points could not be read. The board frame follows the model axes."

    ' ---------- shape of every item in the board frame ----------
    Dim shp() As Double, nNoPts As Long, nTurned As Long
    ReDim shp(n - 1, 10)
    For i = 0 To n - 1
        If i Mod 5 = 0 Then ShowStatus xl, "ImportStep: measuring body " & (i + 1) & " of " & n & " ..."
        If i = pcb And havePcb Then
            For k = 0 To 10: shp(i, k) = shPcb(k): Next
        ElseIf ItemShape(i, fr, sh) Then
            For k = 0 To 10: shp(i, k) = sh(k): Next
        Else
            BoxShape i, fr, sh
            For k = 0 To 10: shp(i, k) = sh(k): Next
            nNoPts = nNoPts + 1
        End If
        If Abs(shp(i, 6)) >= 0.3 Then nTurned = nTurned + 1
    Next
    ShowStatus xl, ""
    If nNoPts > 0 Then Note "info", nNoPts & " bodies gave no surface points; their plain bounding box was used (no rotation)."

    Dim bmin() As Double, bmax() As Double
    ReDim bmin(n - 1, 2): ReDim bmax(n - 1, 2)
    For i = 0 To n - 1
        For k = 0 To 2: bmin(i, k) = shp(i, k): bmax(i, k) = shp(i, 3 + k): Next
    Next
    Dim org(2) As Double
    For k = 0 To 2: org(k) = bmin(pcb, k): Next
    Dim tPcb As Double: tPcb = bmax(pcb, 2) - bmin(pcb, 2)

    Dim fill As Double, bodyVol As Double, boxV As Double
    bodyVol = ItemVolume(pcb)
    boxV = (bmax(pcb, 0) - bmin(pcb, 0)) * (bmax(pcb, 1) - bmin(pcb, 1)) * tPcb
    If bodyVol > 0 And boxV > 0 Then
        fill = bodyVol / boxV
        If fill < 0.9 Then Note "warning", "PCB fills only " & Format(fill, "0%") & " of its box: the board has large cutouts or is not rectangular. " & _
            "Check the PCB L and W and override them if needed."
    End If

    ' ---------- rows: PCB first, then components that pass the size filter ----------
    Dim order() As Long, cnt As Long, skipped As Long, L As Double, W As Double
    ReDim order(n - 1)
    order(0) = pcb: cnt = 1
    For i = 0 To n - 1
        If i <> pcb Then
            L = bmax(i, 0) - bmin(i, 0): W = bmax(i, 1) - bmin(i, 1)
            If L < minSize And W < minSize Then
                skipped = skipped + 1
            Else
                order(cnt) = i: cnt = cnt + 1
            End If
        End If
    Next
    If skipped > 0 Then Note "info", skipped & " bodies smaller than " & minSize & " mm in both L and W were skipped (size filter setting)."

    ' mOut columns: 0 #, 1 name, 2..4 min, 5..7 size, 8 side, 9 key, 10 rot, 11 own L, 12 own W
    Dim r As Long, j As Long, usedKeys As Object, rd As Double
    Set usedKeys = CreateObject("Scripting.Dictionary")
    ReDim mOut(0 To cnt - 1, 0 To 12)
    For r = 0 To cnt - 1
        j = order(r)
        mOut(r, 0) = IIf(r = 0, "PCB", r)
        mOut(r, 1) = nm(j)
        For k = 0 To 2
            mOut(r, 2 + k) = Round(bmin(j, k) - org(k), 3)
            mOut(r, 5 + k) = Round(bmax(j, k) - bmin(j, k), 3)
        Next
        If r = 0 Then
            mOut(r, 8) = "Board"
            mOut(r, 9) = PT_UniqueKey(usedKeys, "pcb")
        Else
            If ((bmin(j, 2) + bmax(j, 2)) / 2 - org(2)) >= tPcb / 2 Then mOut(r, 8) = "Top" Else mOut(r, 8) = "Bottom"
            mOut(r, 9) = PT_UniqueKey(usedKeys, PT_Key(nm(j)))
        End If
        rd = shp(j, 6)
        If r = 0 Or Abs(rd) < 0.3 Then
            mOut(r, 10) = 0: mOut(r, 11) = mOut(r, 5): mOut(r, 12) = mOut(r, 6)
        Else
            mOut(r, 10) = Round(rd, 3): mOut(r, 11) = Round(shp(j, 7), 3): mOut(r, 12) = Round(shp(j, 8), 3)
        End If
    Next
    mCnt = cnt
    For k = 0 To 8: mFr(k) = fr(k): Next
    mPcbL = bmax(pcb, 0) - bmin(pcb, 0)
    mPcbW = bmax(pcb, 1) - bmin(pcb, 1)
    mPcbT = tPcb
    mTurned = nTurned
    Measure = True
End Function

' ================================================================= project folder (page) mode
' Reads the import settings from project.json and writes cad.json: import (this run), importPrevious (the run
' before) and change records for what came, went, moved, changed size or was renamed. project.json is never written:
' the page keeps the user's choices (include, RefDes, overrides) by body key, and new bodies start not included.
Sub RunProject()
    Dim cad As Object, cadPath As String, pcbKey As String, minSize As Double, perBody As Boolean
    Dim tolPos As Double, tolSize As Double, stepF As String, doc As Object, srcName As String
    cadPath = PJ_Dir & "\cad.json"
    Set cad = PJ_ReadOwn("cad.json", "pcb-thermal/cad")
    If cad Is Nothing Then
        PT_Msg "ImportStep did not run. Nothing was changed." & vbCrLf & vbCrLf & cadPath & " could not be read:" & vbCrLf & PT_LastErr & vbCrLf & vbCrLf & _
               "Move that file out of the folder (a copy of each day is in backup\) and run again.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    pcbKey = PJ_Str(PJ_Project, "settings/import/pcbNameContains", "PCB")
    minSize = PJ_Num(PJ_Project, "settings/import/minSizeMm", 0)
    perBody = Not PJ_Bool(PJ_Project, "settings/import/topLevelOnly", False)
    tolPos = PJ_Num(PJ_Project, "settings/tolPosMm", 0.05)
    tolSize = PJ_Num(PJ_Project, "settings/tolSizeMm", 0.05)
    stepF = PJ_Str(PJ_Project, "board/stepFile", "")
    If stepF <> "" And InStr(stepF, ":") = 0 And Left$(stepF, 2) <> "\\" Then stepF = PJ_Dir & "\" & stepF

    ' ---------- model ----------
    Set doc = GetModel(srcName, stepF)
    If doc Is Nothing Then Exit Sub
    If Not Measure(doc, srcName, pcbKey, minSize, perBody, Nothing) Then Exit Sub

    ' ---------- keys and change report ----------
    Dim prevImp As Object, hadOld As Boolean, levelChanged As Boolean, newLevel As String, oldLevel As String, nKept As Long
    Dim counts As Object
    Set counts = CreateObject("Scripting.Dictionary")
    newLevel = IIf(perBody, "bodies", "components")
    Set prevImp = PJ_Obj(cad, "import")
    If Not prevImp Is Nothing Then
        hadOld = True
        If PJ_Has(prevImp, "settingsUsed/topLevelOnly") Then
            oldLevel = IIf(PJ_Bool(prevImp, "settingsUsed/topLevelOnly", False), "components", "bodies")
            levelChanged = (oldLevel <> newLevel)
        End If
    End If
    nKept = KeepKeysByBox(prevImp, tolPos, tolSize)
    If nKept > 0 Then Note "info", nKept & " bodies kept the key of the earlier body with the same position and size, because their name changed. " & _
                                   "Their choices on the page stay with them."
    If levelChanged Then
        Note "warning", "Import level changed from " & oldLevel & " to " & newLevel & ": the change report was skipped once, because nearly every body is new."
    ElseIf hadOld Then
        CompareProject prevImp, tolPos, tolSize, counts
    End If

    ' ---------- bodies new to the project, and choices whose body is gone ----------
    Dim choices As Object, inNew As Object, r As Long, nNew As Long, kk As Variant, refD As String
    Set choices = PJ_Obj(PJ_Project, "bodies")
    Set inNew = CreateObject("Scripting.Dictionary")
    For r = 1 To mCnt - 1
        inNew(CStr(mOut(r, 9))) = 1
        If choices Is Nothing Then
            nNew = nNew + 1
        ElseIf Not choices.Exists(CStr(mOut(r, 9))) Then
            nNew = nNew + 1
        End If
    Next
    If Not choices Is Nothing Then
        For Each kk In choices.Keys
            If Not inNew.Exists(CStr(kk)) And Not PJ_Has(choices(kk), "manual") Then
                refD = PJ_Str(choices(kk), "refDes", "")
                If refD <> "" Or PJ_Has(choices(kk), "override") Or PJ_Has(choices(kk), "nameOverride") Then
                    Note "warning", "Not in this model: body " & kk & IIf(refD <> "", " (RefDes " & refD & ")", "") & _
                                    ". Its choices on the page are kept but not used."
                End If
            End If
        Next
    End If
    If Not hadOld Then Note "info", "First import into this project: nothing to compare, no change records."
    If mTurned > 0 Then Note "info", mTurned & " bodies are turned in the board plane. Build Board builds them as turned rectangles."
    If nNew > 0 Then Note "info", nNew & " bodies are new to this project and start not included. Include and link them on tab 3 of the page."
    If msgs.Count = 0 Then Note "info", "No warnings."

    ' ---------- cad.json ----------
    Dim impD As Object, su As Object, pc As Object, bl As Collection, b As Object, ml As Collection, t As Long
    Set impD = PJ_Dict()
    PJ_Put impD, "model", srcName
    PJ_Put impD, "date", PJ_Now()
    PJ_Put impD, "frame", FrameText(mFr) & " Origin at the PCB corner on its bottom face. Units mm."
    Set su = PJ_Dict()
    PJ_Put su, "pcbNameContains", pcbKey
    PJ_Put su, "minSizeMm", minSize
    PJ_Put su, "topLevelOnly", Not perBody
    PJ_Put impD, "settingsUsed", su
    Set pc = PJ_Dict()
    PJ_Put pc, "name", CStr(mOut(0, 1))
    PJ_Put pc, "box", RowBox(0)
    PJ_Put impD, "pcb", pc
    Set bl = New Collection
    For r = 1 To mCnt - 1
        Set b = PJ_Dict()
        PJ_Put b, "key", CStr(mOut(r, 9))
        PJ_Put b, "name", CStr(mOut(r, 1))
        PJ_Put b, "box", RowBox(r)
        PJ_Put b, "side", CStr(mOut(r, 8))
        If Abs(CDbl(mOut(r, 10))) >= 0.3 Then
            PJ_Put b, "rotDeg", CDbl(mOut(r, 10))
            PJ_Put b, "ownL", PJ_Pos(CDbl(mOut(r, 11)))
            PJ_Put b, "ownW", PJ_Pos(CDbl(mOut(r, 12)))
        End If
        bl.Add b
    Next
    PJ_Put impD, "bodies", bl
    Set ml = New Collection
    For t = 1 To msgs.Count
        ml.Add PJ_Message(msgLvl(t), msgs(t))
    Next
    PJ_Put impD, "messages", ml
    If Not prevImp Is Nothing Then PJ_Put cad, "importPrevious", prevImp
    PJ_Put cad, "import", impD

    Dim nLog As Long
    nLog = PJ_FlushChanges(cad)
    If Not PJ_SaveOwn(cad, "cad.json", MACRO_NAME) Then
        PT_Msg "ImportStep read the model but could not write" & vbCrLf & cadPath & vbCrLf & PT_LastErr & vbCrLf & vbCrLf & _
               "Nothing else was changed. Close any program that has the file open and run again.", vbExclamation, MACRO_NAME
        Exit Sub
    End If

    PT_Msg "Written to:" & vbCrLf & cadPath & vbCrLf & "The page reloads it on its own (tab 2 Import geometry)." & vbCrLf & vbCrLf & _
           "Imported " & (mCnt - 1) & " bodies plus the PCB (" & Round(mPcbL, 2) & " x " & Round(mPcbW, 2) & " x " & _
           Round(mPcbT, 2) & " mm), level: " & newLevel & "." & vbCrLf & _
           mTurned & " bodies turned in the board plane." & vbCrLf & _
           nNew & " bodies new to this project (not included until you include them on tab 3)." & vbCrLf & vbCrLf & _
           "Changes since the last import: " & ChangeSummary(counts, nLog, hadOld, levelChanged, _
               " change records written to cad.json. Tick them on tab 10 of the page (Changes) before Build Board.") & vbCrLf & vbCrLf & _
           "Messages:" & vbCrLf & JoinMsgs("(all messages are on tab 2 of the page)"), vbInformation, MACRO_NAME
End Sub

' Box of row r of mOut as a JSON object.
Function RowBox(ByVal r As Long) As Object
    Set RowBox = PJ_Box(CDbl(mOut(r, 2)), CDbl(mOut(r, 3)), CDbl(mOut(r, 4)), CDbl(mOut(r, 5)), CDbl(mOut(r, 6)), CDbl(mOut(r, 7)))
End Function

' A body whose name changed since the last import gets the key of the old body with the same box (within the
' tolerances), when no body of this import uses that key. The page keeps the user's choices by key, so they stay
' with the body. Returns the number of keys kept this way.
Function KeepKeysByBox(ByVal prevImp As Object, ByVal tolPos As Double, ByVal tolSize As Double) As Long
    Dim oldB As Object, it As Variant, used As Object, oldKeys As Object, freeOld As Collection
    Dim r As Long, j As Long, oldKey As String
    If prevImp Is Nothing Then Exit Function
    Set oldB = PJ_Obj(prevImp, "bodies")
    If oldB Is Nothing Then Exit Function
    Set used = CreateObject("Scripting.Dictionary")
    Set oldKeys = CreateObject("Scripting.Dictionary")
    For r = 0 To mCnt - 1: used(CStr(mOut(r, 9))) = 1: Next
    Set freeOld = New Collection
    For Each it In oldB
        If IsObject(it) Then
            oldKey = PJ_Str(it, "key", "")
            If oldKey <> "" Then
                oldKeys(oldKey) = 1
                If Not used.Exists(oldKey) Then freeOld.Add it
            End If
        End If
    Next
    If freeOld.Count = 0 Then Exit Function
    For r = 1 To mCnt - 1
        If Not oldKeys.Exists(CStr(mOut(r, 9))) Then
            For j = 1 To freeOld.Count
                If BoxMatches(freeOld(j), r, tolPos, tolSize) Then
                    mOut(r, 9) = PJ_Str(freeOld(j), "key", "")
                    freeOld.Remove j
                    KeepKeysByBox = KeepKeysByBox + 1
                    Exit For
                End If
            Next
        End If
        If freeOld.Count = 0 Then Exit For
    Next
End Function

' True when the box of an imported body (cad.json) equals row r of mOut within the tolerances.
Function BoxMatches(ByVal ob As Object, ByVal r As Long, ByVal tolPos As Double, ByVal tolSize As Double) As Boolean
    Dim flds As Variant, k As Long, tol As Double
    flds = Array("xMin", "yMin", "zMin", "l", "w", "h")
    For k = 0 To 5
        If Not PJ_IsNum(ob, "box/" & flds(k)) Then Exit Function
        If k < 3 Then tol = tolPos Else tol = tolSize
        If Abs(PJ_Num(ob, "box/" & flds(k), 0) - CDbl(mOut(r, 2 + k))) > tol Then Exit Function
    Next
    BoxMatches = True
End Function

' Change records (PT_LogAdd) for this import against the last one in cad.json. PJ_FlushChanges writes them.
Sub CompareProject(ByVal prevImp As Object, ByVal tolPos As Double, ByVal tolSize As Double, ByVal counts As Object)
    Dim oldByKey As Object, used As Object, oldB As Object, it As Variant, r As Long, keyS As String
    Dim ob As Object, nameS As String, oldName As String
    Set oldByKey = CreateObject("Scripting.Dictionary")
    Set used = CreateObject("Scripting.Dictionary")
    If Not PJ_Obj(prevImp, "pcb") Is Nothing Then oldByKey.Add "pcb", PJ_Obj(prevImp, "pcb")
    Set oldB = PJ_Obj(prevImp, "bodies")
    If Not oldB Is Nothing Then
        For Each it In oldB
            If IsObject(it) Then
                keyS = PJ_Str(it, "key", "")
                If keyS <> "" And Not oldByKey.Exists(keyS) Then oldByKey.Add keyS, it
            End If
        Next
    End If
    For r = 0 To mCnt - 1
        keyS = CStr(mOut(r, 9))
        If r = 0 Then nameS = "PCB" Else nameS = CStr(mOut(r, 1))
        If oldByKey.Exists(keyS) Then
            used(keyS) = 1
            Set ob = oldByKey(keyS)
            CompareOne keyS, nameS, BodyRefDes(keyS), ob, r, tolPos, tolSize, counts
            oldName = PJ_Str(ob, "name", "")
            If r > 0 And oldName <> "" And oldName <> nameS Then
                PT_LogAdd MACRO_NAME, keyS, nameS, BodyRefDes(keyS), "renamed", "Name", oldName, nameS
                Bump counts, "renamed"
            End If
        ElseIf r > 0 Then
            PT_LogAdd MACRO_NAME, keyS, nameS, "", "added", "Box", "", _
                      BoxText(mOut(r, 2), mOut(r, 3), mOut(r, 4), mOut(r, 5), mOut(r, 6), mOut(r, 7))
            Bump counts, "added"
        End If
    Next
    For Each it In oldByKey.Keys
        If Not used.Exists(it) And it <> "pcb" Then
            Set ob = oldByKey(it)
            PT_LogAdd MACRO_NAME, CStr(it), PJ_Str(ob, "name", ""), BodyRefDes(CStr(it)), "removed", "Box", _
                      BoxText(PJ_Num(ob, "box/xMin", 0), PJ_Num(ob, "box/yMin", 0), PJ_Num(ob, "box/zMin", 0), _
                              PJ_Num(ob, "box/l", 0), PJ_Num(ob, "box/w", 0), PJ_Num(ob, "box/h", 0)), ""
            Bump counts, "removed"
        End If
    Next
End Sub

Sub CompareOne(ByVal keyS As String, ByVal nameS As String, ByVal refD As String, ByVal ob As Object, ByVal r As Long, _
               ByVal tolPos As Double, ByVal tolSize As Double, ByVal counts As Object)
    Dim flds As Variant, labels As Variant, k As Long, a As Double, b As Double
    flds = Array("xMin", "yMin", "zMin", "l", "w", "h")
    labels = Array("X min", "Y min", "Z min", "L", "W", "H")
    For k = 0 To 5
        If PJ_IsNum(ob, "box/" & flds(k)) Then
            a = PJ_Num(ob, "box/" & flds(k), 0)
            b = CDbl(mOut(r, 2 + k))
            If k < 3 Then
                If Abs(b - a) > tolPos Then
                    PT_LogAdd MACRO_NAME, keyS, nameS, refD, "moved", CStr(labels(k)), a, b
                    Bump counts, "moved"
                End If
            Else
                If Abs(b - a) > tolSize Then
                    PT_LogAdd MACRO_NAME, keyS, nameS, refD, "resized", CStr(labels(k)), a, b
                    Bump counts, "resized"
                End If
            End If
        End If
    Next
End Sub

' RefDes the page links to a body key ("" when none).
Function BodyRefDes(ByVal keyS As String) As String
    BodyRefDes = PJ_Str(PJ_Project, "bodies/" & keyS & "/refDes", "")
End Function

' ---------------------------------------------------------------- change report
' Compares the previous STEP import table (oldImp, 1-based, columns B..K) with the new rows.
' Adds Change log rows through PTCommon. Returns False when there was no previous import.
Function CompareImports(ByRef oldImp As Variant, ByRef outv() As Variant, ByVal cnt As Long, ByVal refByKey As Object, _
                        ByVal tolPos As Double, ByVal tolSize As Double, ByVal counts As Object) As Boolean
    Dim oldRow As Object, used As Object, r As Long, s As String, k0 As String, o As Long
    Dim newOnly As Collection, oldOnly As Collection, it As Variant, it2 As Variant, found As Boolean, hasRows As Boolean
    Set oldRow = CreateObject("Scripting.Dictionary")
    Set used = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(oldImp, 1)
        s = PT_Str(oldImp(r, 2))
        If s <> "" Then
            If r = 1 Then
                k0 = "pcb"
            Else
                k0 = PT_Str(oldImp(r, 10))
                If k0 = "" Then k0 = PT_Key(s)
                hasRows = True
            End If
            If Not oldRow.Exists(k0) Then oldRow.Add k0, r
        End If
    Next
    If Not hasRows Then Exit Function
    CompareImports = True

    Set newOnly = New Collection
    For r = 0 To cnt - 1
        k0 = CStr(outv(r, 9))
        If oldRow.Exists(k0) Then
            used.Add k0, 1
            CompareBox k0, IIf(r = 0, "PCB", CStr(outv(r, 1))), RefOf(refByKey, k0), oldImp, oldRow(k0), outv, r, tolPos, tolSize, counts
        Else
            newOnly.Add r
        End If
    Next
    Set oldOnly = New Collection
    For Each it In oldRow.Keys
        If Not used.Exists(it) Then oldOnly.Add CStr(it)
    Next

    ' same box, different key = renamed
    For Each it In newOnly
        r = it
        found = False
        For Each it2 In oldOnly
            o = oldRow(it2)
            If BoxClose(oldImp, o, outv, r, tolPos, tolSize) Then
                PT_LogAdd MACRO_NAME, CStr(outv(r, 9)), CStr(outv(r, 1)), RefOf(refByKey, CStr(it2)), "renamed", "Name", _
                          PT_Str(oldImp(o, 2)), CStr(outv(r, 1))
                Bump counts, "renamed"
                RemoveFrom oldOnly, CStr(it2)
                found = True
                Exit For
            End If
        Next
        If Not found Then
            PT_LogAdd MACRO_NAME, CStr(outv(r, 9)), CStr(outv(r, 1)), "", "added", "Box", "", BoxText(outv(r, 2), outv(r, 3), outv(r, 4), outv(r, 5), outv(r, 6), outv(r, 7))
            Bump counts, "added"
        End If
    Next
    For Each it In oldOnly
        o = oldRow(it)
        PT_LogAdd MACRO_NAME, CStr(it), PT_Str(oldImp(o, 2)), RefOf(refByKey, CStr(it)), "removed", "Box", _
                  BoxText(oldImp(o, 3), oldImp(o, 4), oldImp(o, 5), oldImp(o, 6), oldImp(o, 7), oldImp(o, 8)), ""
        Bump counts, "removed"
    Next
End Function

Sub CompareBox(ByVal key As String, ByVal nameS As String, ByVal refD As String, ByRef oldImp As Variant, ByVal o As Long, _
               ByRef outv() As Variant, ByVal r As Long, ByVal tolPos As Double, ByVal tolSize As Double, ByVal counts As Object)
    Dim k As Long, fields As Variant, a As Double, b As Double
    fields = Array("X min", "Y min", "Z min", "L", "W", "H")
    For k = 0 To 5
        a = PT_Num(oldImp(o, 3 + k), 0)
        b = PT_Num(outv(r, 2 + k), 0)
        If k < 3 Then
            If Abs(b - a) > tolPos Then
                PT_LogAdd MACRO_NAME, key, nameS, refD, "moved", CStr(fields(k)), a, b
                Bump counts, "moved"
            End If
        Else
            If Abs(b - a) > tolSize Then
                PT_LogAdd MACRO_NAME, key, nameS, refD, "resized", CStr(fields(k)), a, b
                Bump counts, "resized"
            End If
        End If
    Next
End Sub

Function BoxClose(ByRef oldImp As Variant, ByVal o As Long, ByRef outv() As Variant, ByVal r As Long, _
                  ByVal tolPos As Double, ByVal tolSize As Double) As Boolean
    Dim k As Long, tol As Double
    For k = 0 To 5
        If Not PT_IsNum(oldImp(o, 3 + k)) Then Exit Function
        If k < 3 Then tol = tolPos Else tol = tolSize
        If Abs(CDbl(oldImp(o, 3 + k)) - CDbl(outv(r, 2 + k))) > tol Then Exit Function
    Next
    BoxClose = True
End Function

Function BoxText(ByVal x0 As Variant, ByVal y0 As Variant, ByVal z0 As Variant, ByVal lx As Variant, ByVal wy As Variant, ByVal hz As Variant) As String
    BoxText = "at " & PT_Num(x0, 0) & ", " & PT_Num(y0, 0) & ", " & PT_Num(z0, 0) & _
              " size " & PT_Num(lx, 0) & " x " & PT_Num(wy, 0) & " x " & PT_Num(hz, 0)
End Function

Function RefOf(ByVal refByKey As Object, ByVal key As String) As String
    If refByKey.Exists(key) Then RefOf = refByKey(key)
End Function

Sub Bump(ByVal counts As Object, ByVal what As String)
    If counts.Exists(what) Then counts(what) = counts(what) + 1 Else counts.Add what, 1
End Sub

Sub RemoveFrom(ByVal coll As Collection, ByVal what As String)
    Dim i As Long
    For i = 1 To coll.Count
        If coll(i) = what Then coll.Remove i: Exit Sub
    Next
End Sub

' ---------------------------------------------------------------- model access
' The model to import: the one open in SolidWorks, or a STEP file (the picker starts at stepDefault when given).
Function GetModel(ByRef srcName As String, ByVal stepDefault As String) As Object
    Dim act As Object, a As VbMsgBoxResult
    Set act = swApp.ActiveDoc
    If Not act Is Nothing Then
        If act.GetType = swDocASSEMBLY Or act.GetType = swDocPART Then
            a = PT_Msg("Use the model open in SolidWorks?" & vbCrLf & act.GetTitle & vbCrLf & vbCrLf & _
                       "No = pick a STEP file instead.", vbYesNoCancel + vbQuestion, MACRO_NAME)
            If a = vbCancel Then Exit Function
            If a = vbYes Then
                srcName = act.GetTitle
                Set GetModel = act
                Exit Function
            End If
        End If
    End If
    Dim f As String, opt As Long, cfg As String, disp As String
    If stepDefault <> "" Then
        If Dir(stepDefault) = "" Then stepDefault = ""
    End If
    f = swApp.GetOpenFileName("Pick the PCB STEP file", stepDefault, "STEP files (*.step;*.stp)|*.step;*.stp|All files (*.*)|*.*||", opt, cfg, disp)
    If f = "" Then Exit Function
    Dim impData As Object, errs As Long
    Set impData = swApp.GetImportFileData(f)
    Set GetModel = swApp.LoadFile4(f, "r", impData, errs)
    If GetModel Is Nothing Then
        PT_Msg "SolidWorks could not open " & f & " (error " & errs & ").", vbExclamation, MACRO_NAME
        Exit Function
    End If
    srcName = f
End Function

' ---------------------------------------------------------------- collecting the bodies
Sub CollectItems(ByVal doc As Object, ByVal perBody As Boolean)
    n = 0
    ReDim nm(MAX_ITEMS - 1): ReDim bx(MAX_ITEMS - 1, 5): ReDim rbx(MAX_ITEMS - 1, 5)
    ReDim srcComp(MAX_ITEMS - 1): ReDim srcBody(MAX_ITEMS - 1)
    ReDim srcMode(MAX_ITEMS - 1): ReDim srcXf(MAX_ITEMS - 1, 12)
    Dim i As Long, v As Variant, c As Object
    If doc.GetType = swDocASSEMBLY Then
        doc.ResolveAllLightWeightComponents False
        Dim comps As Variant, ch As Variant
        comps = doc.GetComponents(Not perBody)
        If IsEmpty(comps) Then Exit Sub
        ' per-component mode: a STEP often wraps everything in one sub-assembly, step into it
        If Not perBody Then
            Do While UBound(comps) = 0
                ch = comps(0).GetChildren
                If IsEmpty(ch) Then Exit Do
                If UBound(ch) < 0 Then Exit Do
                comps = ch
            Loop
        End If
        For i = 0 To UBound(comps)
            Set c = comps(i)
            If Not c.IsSuppressed Then
                If perBody Then
                    If IsLeaf(c) Then AddCompBodies c
                Else
                    If Not c.IsHidden(True) Then AddWholeComp c
                End If
            End If
        Next
    Else
        Dim bodies As Variant, one(5) As Double, xd(12) As Double, k As Long
        For k = 0 To 12: xd(k) = 0: Next
        xd(0) = 1: xd(4) = 1: xd(8) = 1: xd(12) = 1
        bodies = doc.GetBodies2(swSolidBody, True)
        If IsEmpty(bodies) Then Exit Sub
        For i = 0 To UBound(bodies)
            v = bodies(i).GetBodyBox
            If Not IsEmpty(v) Then
                For k = 0 To 5: one(k) = v(k) * 1000#: Next
                AddItem bodies(i).Name, one, one, Nothing, bodies(i), 0, xd
            End If
        Next
    End If
End Sub

Function IsLeaf(ByVal c As Object) As Boolean
    Dim ch As Variant
    ch = c.GetChildren
    If IsEmpty(ch) Then IsLeaf = True: Exit Function
    IsLeaf = (UBound(ch) < 0)
End Function

' One item per solid body of a component. A component with one body keeps the component name (so keys and
' inputs from earlier imports still match); with several bodies the body name is put before the instance suffix.
Sub AddCompBodies(ByVal c As Object)
    Dim bodies As Variant, info As Variant, info2 As Variant, sheetB As Variant, k As Long, q As Long
    Dim xd(12) As Double, spMode As Long, ab(5) As Double, rb(5) As Double, bb As Variant, lbl As String, nb As Long
    On Error Resume Next
    bodies = c.GetBodies3(swSolidBody, info)
    sheetB = c.GetBodies3(swSheetBody, info2)
    On Error GoTo 0
    If Not IsEmpty(sheetB) Then nSheetBodies = nSheetBodies + UBound(sheetB) + 1
    If IsEmpty(bodies) Then Exit Sub
    nb = UBound(bodies) + 1
    ReadXf c, xd
    spMode = PickMode(c, bodies, xd)
    For k = 0 To nb - 1
        bb = Empty
        On Error Resume Next
        bb = bodies(k).GetBodyBox
        On Error GoTo 0
        If Not IsEmpty(bb) Then
            For q = 0 To 5: rb(q) = bb(q) * 1000#: Next
            BoxXf bb, xd, spMode, ab, True
            If nb = 1 Then lbl = c.Name2 Else lbl = BodyLabel(c.Name2, bodies(k).Name, k)
            AddItem lbl, ab, rb, c, bodies(k), spMode, xd
        End If
    Next
End Sub

' One item for a whole component (setting Imp_TopOnly = Yes): its box in the assembly.
Sub AddWholeComp(ByVal c As Object)
    Dim v As Variant, bodies As Variant, info As Variant, k As Long, q As Long, xd(12) As Double, spMode As Long
    Dim ab(5) As Double, rb(5) As Double, bb As Variant, haveRaw As Boolean
    v = c.GetBox(False, False)
    If IsEmpty(v) Then Exit Sub
    For q = 0 To 5: ab(q) = v(q) * 1000#: Next
    ReadXf c, xd
    On Error Resume Next
    bodies = c.GetBodies3(swSolidBody, info)
    On Error GoTo 0
    For q = 0 To 2: rb(q) = 1E+30: rb(3 + q) = -1E+30: Next
    If Not IsEmpty(bodies) Then
        spMode = PickMode(c, bodies, xd)
        For k = 0 To UBound(bodies)
            bb = Empty
            On Error Resume Next
            bb = bodies(k).GetBodyBox
            On Error GoTo 0
            If Not IsEmpty(bb) Then
                haveRaw = True
                For q = 0 To 2
                    If bb(q) * 1000# < rb(q) Then rb(q) = bb(q) * 1000#
                    If bb(3 + q) * 1000# > rb(3 + q) Then rb(3 + q) = bb(3 + q) * 1000#
                Next
            End If
        Next
    End If
    If Not haveRaw Then
        For q = 0 To 5: rb(q) = ab(q): Next
        spMode = 0
    End If
    AddItem c.Name2, ab, rb, c, Nothing, spMode, xd
End Sub

' "<path>/<part>.STEP-1" + body "Imported12" -> "<path>/<part>.STEP_Imported12-1"
Function BodyLabel(ByVal compName As String, ByVal bodyName As String, ByVal k As Long) As String
    Dim re As Object, mt As Object, headS As String, tailS As String
    If bodyName = "" Then bodyName = "body" & (k + 1)
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(.*?)(-\d+|<\d+>)$"
    If re.Test(compName) Then
        Set mt = re.Execute(compName)
        headS = mt(0).SubMatches(0): tailS = mt(0).SubMatches(1)
    Else
        headS = compName: tailS = ""
    End If
    BodyLabel = headS & "_" & bodyName & tailS
End Function

Sub AddItem(ByVal lbl As String, ByRef asmBox() As Double, ByRef rawBox() As Double, ByVal comp As Object, ByVal body As Object, _
            ByVal spMode As Long, ByRef xd() As Double)
    If n >= MAX_ITEMS Then itemsCapped = True: Exit Sub
    Dim k As Long
    nm(n) = lbl
    For k = 0 To 5: bx(n, k) = asmBox(k): rbx(n, k) = rawBox(k): Next
    For k = 0 To 12: srcXf(n, k) = xd(k): Next
    Set srcComp(n) = comp
    Set srcBody(n) = body
    srcMode(n) = spMode
    n = n + 1
End Sub

' Component transform as 13 numbers: 9 rotation, 3 translation (metres), scale. Identity when unreadable.
Sub ReadXf(ByVal c As Object, ByRef xd() As Double)
    Dim t As Object, a As Variant, k As Long
    For k = 0 To 12: xd(k) = 0: Next
    xd(0) = 1: xd(4) = 1: xd(8) = 1: xd(12) = 1
    On Error Resume Next
    Set t = c.Transform2
    If t Is Nothing Then Exit Sub
    a = t.ArrayData
    If IsEmpty(a) Then Exit Sub
    If UBound(a) < 12 Then Exit Sub
    For k = 0 To 12: xd(k) = CDbl(a(k)): Next
End Sub

' Which way turns a body's own-space points into model space: 0 none, 1 matrix, 2 transposed matrix.
' Decided by comparing the transformed body boxes with the component box in the assembly (GetBox).
Function PickMode(ByVal c As Object, ByVal bodies As Variant, ByRef xd() As Double) As Long
    Dim cb As Variant, md As Long, k As Long, q As Long, bb As Variant, ub(5) As Double, e As Double, bestE As Double, bestM As Long
    On Error Resume Next
    cb = c.GetBox(False, False)
    On Error GoTo 0
    If IsEmpty(cb) Then Exit Function
    bestE = 1E+30
    For md = 0 To 2
        For q = 0 To 2: ub(q) = 1E+30: ub(3 + q) = -1E+30: Next
        For k = 0 To UBound(bodies)
            bb = Empty
            On Error Resume Next
            bb = bodies(k).GetBodyBox
            On Error GoTo 0
            If Not IsEmpty(bb) Then BoxXf bb, xd, md, ub, False
        Next
        e = 0
        For q = 0 To 5
            If Abs(ub(q) - cb(q) * 1000#) > e Then e = Abs(ub(q) - cb(q) * 1000#)
        Next
        If e < bestE - 0.001 Then bestE = e: bestM = md
    Next
    PickMode = bestM
End Function

' Box (metres, 6 numbers) -> corners turned by the component transform -> axis-aligned box in mm.
' reset = True starts the result from nothing, False adds to it.
Sub BoxXf(ByVal bb As Variant, ByRef xd() As Double, ByVal spMode As Long, ByRef outB() As Double, ByVal reset As Boolean)
    Dim a As Long, b As Long, c As Long, q As Long, x As Double, y As Double, z As Double, ox As Double, oy As Double, oz As Double
    If reset Then
        For q = 0 To 2: outB(q) = 1E+30: outB(3 + q) = -1E+30: Next
    End If
    For a = 0 To 1: For b = 0 To 1: For c = 0 To 1
        x = bb(0 + 3 * a) * 1000#: y = bb(1 + 3 * b) * 1000#: z = bb(2 + 3 * c) * 1000#
        XfPoint spMode, xd, x, y, z, ox, oy, oz
        If ox < outB(0) Then outB(0) = ox
        If oy < outB(1) Then outB(1) = oy
        If oz < outB(2) Then outB(2) = oz
        If ox > outB(3) Then outB(3) = ox
        If oy > outB(4) Then outB(4) = oy
        If oz > outB(5) Then outB(5) = oz
    Next: Next: Next
End Sub

' Point in mm, own space -> model space (mm)
Sub XfPoint(ByVal spMode As Long, ByRef xd() As Double, ByVal x As Double, ByVal y As Double, ByVal z As Double, _
            ByRef ox As Double, ByRef oy As Double, ByRef oz As Double)
    Dim sc As Double
    If spMode = 0 Then
        ox = x: oy = y: oz = z
        Exit Sub
    End If
    sc = xd(12): If sc = 0 Then sc = 1
    If spMode = 1 Then
        ox = (xd(0) * x + xd(1) * y + xd(2) * z) * sc + xd(9) * 1000#
        oy = (xd(3) * x + xd(4) * y + xd(5) * z) * sc + xd(10) * 1000#
        oz = (xd(6) * x + xd(7) * y + xd(8) * z) * sc + xd(11) * 1000#
    Else
        ox = (xd(0) * x + xd(3) * y + xd(6) * z) * sc + xd(9) * 1000#
        oy = (xd(1) * x + xd(4) * y + xd(7) * z) * sc + xd(10) * 1000#
        oz = (xd(2) * x + xd(5) * y + xd(8) * z) * sc + xd(11) * 1000#
    End If
End Sub

Function ItemVolume(ByVal i As Long) As Double
    Dim bodies As Variant, info As Variant, vol As Double, leaves As Collection, lc As Variant
    On Error Resume Next
    If Not srcBody(i) Is Nothing Then
        vol = BodiesVolume(Array(srcBody(i)))
    ElseIf Not srcComp(i) Is Nothing Then
        Set leaves = New Collection
        LeafComps srcComp(i), leaves
        For Each lc In leaves
            bodies = Empty
            bodies = lc.GetBodies3(swSolidBody, info)
            vol = vol + BodiesVolume(bodies)
        Next
    End If
    ItemVolume = vol * 1000000000#                           ' m3 -> mm3
End Function

Function BodiesVolume(ByVal bodies As Variant) As Double
    Dim k As Long, mp As Variant, vol As Double
    On Error Resume Next
    If IsEmpty(bodies) Then Exit Function
    For k = 0 To UBound(bodies)
        mp = Empty
        mp = bodies(k).GetMassProperties(1#)
        If Not IsEmpty(mp) Then vol = vol + mp(3)
    Next
    BodiesVolume = vol
End Function

' The part components under a component (the component itself when it is a part). A component that is a
' sub-assembly has no bodies of its own, so its shape comes from these. Suppressed and hidden ones are left out.
Sub LeafComps(ByVal c As Object, ByVal leaves As Collection)
    Dim ch As Variant, k As Long
    If c.IsSuppressed Then Exit Sub
    If IsLeaf(c) Then
        leaves.Add c
        Exit Sub
    End If
    ch = c.GetChildren
    For k = 0 To UBound(ch)
        If Not ch(k).IsHidden(True) Then LeafComps ch(k), leaves
    Next
End Sub

' ---------------------------------------------------------------- surface points of a body
' Points (mm, model space) of an item: 3 numbers per point, returned count is the number of NUMBERS used.
' A whole component is read part by part, each with its own transform (it can be a sub-assembly).
Function ItemPoints(ByVal i As Long, ByRef pts() As Double) As Long
    Dim np As Long, k As Long, q As Long, bodies As Variant, info As Variant, rb(5) As Double, ab(5) As Double
    Dim xd(12) As Double, md As Long, start As Long, npPrev As Long, cb As Variant, bb As Variant
    Dim leaves As Collection, lc As Variant
    ReDim pts(3 * 4096 - 1)
    On Error GoTo failed
    If Not srcBody(i) Is Nothing Then
        np = AppendBodyPts(srcBody(i), pts, np)
        If np < 9 Then Exit Function
        For q = 0 To 5: rb(q) = rbx(i, q): ab(q) = bx(i, q): Next
        For q = 0 To 12: xd(q) = srcXf(i, q): Next
        FixPts pts, 0, np, rb, ab, srcMode(i), xd
    ElseIf Not srcComp(i) Is Nothing Then
        Set leaves = New Collection
        LeafComps srcComp(i), leaves
        For Each lc In leaves
            bodies = Empty
            On Error Resume Next
            bodies = lc.GetBodies3(swSolidBody, info)
            On Error GoTo failed
            If Not IsEmpty(bodies) Then
                cb = Empty
                On Error Resume Next
                cb = lc.GetBox(False, False)
                On Error GoTo failed
                ReadXf lc, xd
                md = PickMode(lc, bodies, xd)
                For q = 0 To 2: rb(q) = 1E+30: rb(3 + q) = -1E+30: Next
                start = np
                For k = 0 To UBound(bodies)
                    bb = Empty
                    On Error Resume Next
                    bb = bodies(k).GetBodyBox
                    On Error GoTo failed
                    If Not IsEmpty(bb) Then
                        For q = 0 To 2
                            If bb(q) * 1000# < rb(q) Then rb(q) = bb(q) * 1000#
                            If bb(3 + q) * 1000# > rb(3 + q) Then rb(3 + q) = bb(3 + q) * 1000#
                        Next
                    End If
                    npPrev = np
                    np = AppendBodyPts(bodies(k), pts, np)
                    If np = npPrev Then Exit Function           ' a body without points would leave a hole in the shape
                Next
                If IsEmpty(cb) Then
                    BoxXf Array(rb(0) / 1000#, rb(1) / 1000#, rb(2) / 1000#, rb(3) / 1000#, rb(4) / 1000#, rb(5) / 1000#), xd, md, ab, True
                Else
                    For q = 0 To 5: ab(q) = cb(q) * 1000#: Next
                End If
                FixPts pts, start, np, rb, ab, md, xd
            End If
        Next
    End If
    If np < 9 Then Exit Function
    ItemPoints = np
    Exit Function
failed:
    ItemPoints = 0
End Function

' Puts the points pts(start .. np-1) of one part into model space (mm). rawBox: the part's body box in its own
' space (mm), asmBox: its box in the assembly (mm), md and xd: how to turn own-space points (see PickMode, ReadXf).
Sub FixPts(ByRef pts() As Double, ByVal start As Long, ByVal np As Long, ByRef rawBox() As Double, _
           ByRef asmBox() As Double, ByVal md As Long, ByRef xd() As Double)
    Dim rb(5) As Double, dRaw As Double, dAsm As Double, sc As Double, ox As Double, oy As Double, oz As Double
    Dim k As Long, q As Long
    For q = 0 To 2: rb(q) = 1E+30: rb(3 + q) = -1E+30: Next
    For k = start To np - 3 Step 3
        For q = 0 To 2
            If pts(k + q) < rb(q) Then rb(q) = pts(k + q)
            If pts(k + q) > rb(3 + q) Then rb(3 + q) = pts(k + q)
        Next
    Next
    ' unit guard: the points must have about the size of the body box (mm). A factor of 1000 means the API returned
    ' millimetres or metres where metres were expected: put it right instead of building the wrong shape.
    dRaw = (rb(3) - rb(0)) + (rb(4) - rb(1)) + (rb(5) - rb(2))
    dAsm = (rawBox(3) - rawBox(0)) + (rawBox(4) - rawBox(1)) + (rawBox(5) - rawBox(2))
    sc = 1#
    If dRaw > 0 And dAsm > 0 Then
        If dRaw / dAsm > 500 And dRaw / dAsm < 2000 Then sc = 0.001
        If dRaw / dAsm < 0.002 And dRaw / dAsm > 0.0005 Then sc = 1000#
    End If
    If sc <> 1# Then
        For k = start To np - 1: pts(k) = pts(k) * sc: Next
        For q = 0 To 5: rb(q) = rb(q) * sc: Next
    End If
    If md = 0 Then Exit Sub
    ' are the points in the part's own space (turn them) or already in the assembly (leave them)?
    dRaw = 0: dAsm = 0
    For q = 0 To 5
        If Abs(rb(q) - rawBox(q)) > dRaw Then dRaw = Abs(rb(q) - rawBox(q))
        If Abs(rb(q) - asmBox(q)) > dAsm Then dAsm = Abs(rb(q) - asmBox(q))
    Next
    If dAsm + 0.0001 < dRaw Then Exit Sub
    For k = start To np - 3 Step 3
        XfPoint md, xd, pts(k), pts(k + 1), pts(k + 2), ox, oy, oz
        pts(k) = ox: pts(k + 1) = oy: pts(k + 2) = oz
    Next
End Sub

Sub GrowPts(ByRef pts() As Double, ByVal need As Long)
    Dim cap As Long
    cap = UBound(pts) + 1
    If need <= cap Then Exit Sub
    Do While cap < need: cap = cap * 2: Loop
    ReDim Preserve pts(cap - 1)
End Sub

' Appends the surface points of a body (mm, the body's own space). Two ways: face triangles (when EVERY face has them),
' else the body tessellation object. Sparse sources such as the vertex list are not used: they miss curved surfaces.
' Returns the new count of numbers (unchanged when no complete point set could be read).
Function AppendBodyPts(ByVal body As Object, ByRef pts() As Double, ByVal np As Long) As Long
    Dim faces As Variant, f As Long, t As Variant, q As Long, m As Long, start As Long, nEmpty As Long
    Dim tess As Object, nv As Long, p As Variant
    On Error Resume Next
    start = np
    faces = body.GetFaces
    If Not IsEmpty(faces) Then
        For f = 0 To UBound(faces)
            t = Empty
            t = faces(f).GetTessTriangles(True)
            If IsEmpty(t) Then
                nEmpty = nEmpty + 1
            Else
                m = UBound(t) + 1
                GrowPts pts, np + m
                For q = 0 To m - 3 Step 3
                    pts(np) = CDbl(t(q)) * 1000#: pts(np + 1) = CDbl(t(q + 1)) * 1000#: pts(np + 2) = CDbl(t(q + 2)) * 1000#
                    np = np + 3
                Next
            End If
        Next
    End If
    If nEmpty > 0 Then np = start                       ' some faces had no triangles: the set is incomplete, use the tessellation object
    If np = start And Not IsEmpty(faces) Then
        Set tess = Nothing
        Set tess = body.GetTessellation(faces)
        If Not tess Is Nothing Then
            tess.NeedFaceFacetMap = False
            tess.NeedVertexParams = False
            tess.NeedVertexNormal = False
            If tess.Tessellate Then
                nv = tess.GetVertexCount
                GrowPts pts, np + 3 * nv
                For q = 0 To nv - 1
                    p = Empty
                    p = tess.GetVertexPoint(q)
                    If Not IsEmpty(p) Then
                        pts(np) = CDbl(p(0)) * 1000#: pts(np + 1) = CDbl(p(1)) * 1000#: pts(np + 2) = CDbl(p(2)) * 1000#
                        np = np + 3
                    End If
                Next
            End If
        End If
    End If
    AppendBodyPts = np
End Function

' ---------------------------------------------------------------- shape in the board frame
' sh(0..2) min and sh(3..5) max of the envelope (board X, Y, Z), sh(6) rotation (deg, -45..45, 0 = square to the frame),
' sh(7) own L, sh(8) own W (the turned rectangle), sh(9) sh(10) centre X, Y. False when no points could be read.
Function ItemShape(ByVal i As Long, ByRef fr() As Double, ByRef sh() As Double) As Boolean
    Dim pts() As Double, np As Long
    np = ItemPoints(i, pts)
    If np < 9 Then Exit Function
    ItemShape = Footprint(pts, np, fr, sh)
End Function

' Same result layout from the plain box (corners projected on the frame), no rotation.
Sub BoxShape(ByVal i As Long, ByRef fr() As Double, ByRef sh() As Double)
    Dim a As Long, b As Long, c As Long, k As Long, p(2) As Double, q(2) As Double
    For k = 0 To 2: sh(k) = 1E+30: sh(3 + k) = -1E+30: Next
    For a = 0 To 1: For b = 0 To 1: For c = 0 To 1
        p(0) = bx(i, 0 + 3 * a): p(1) = bx(i, 1 + 3 * b): p(2) = bx(i, 2 + 3 * c)
        For k = 0 To 2
            q(k) = p(0) * fr(3 * k) + p(1) * fr(3 * k + 1) + p(2) * fr(3 * k + 2)
            If q(k) < sh(k) Then sh(k) = q(k)
            If q(k) > sh(3 + k) Then sh(3 + k) = q(k)
        Next
    Next: Next: Next
    sh(6) = 0: sh(7) = sh(3) - sh(0): sh(8) = sh(4) - sh(1)
    sh(9) = (sh(0) + sh(3)) / 2: sh(10) = (sh(1) + sh(4)) / 2
End Sub

' Smallest turned rectangle round the points, seen from +Z of the frame fr (rows: X, Y, Z unit vectors).
Function Footprint(ByRef pts() As Double, ByVal np As Long, ByRef fr() As Double, ByRef sh() As Double) As Boolean
    Dim q As Long, u As Double, v As Double, w As Double, bi As Long
    Dim umin As Double, umax As Double, vmin As Double, vmax As Double, wmin As Double, wmax As Double
    umin = 1E+30: vmin = 1E+30: wmin = 1E+30: umax = -1E+30: vmax = -1E+30: wmax = -1E+30
    For q = 0 To np - 3 Step 3
        u = pts(q) * fr(0) + pts(q + 1) * fr(1) + pts(q + 2) * fr(2)
        v = pts(q) * fr(3) + pts(q + 1) * fr(4) + pts(q + 2) * fr(5)
        w = pts(q) * fr(6) + pts(q + 1) * fr(7) + pts(q + 2) * fr(8)
        If u < umin Then umin = u
        If u > umax Then umax = u
        If v < vmin Then vmin = v
        If v > vmax Then vmax = v
        If w < wmin Then wmin = w
        If w > wmax Then wmax = w
    Next
    sh(0) = umin: sh(1) = vmin: sh(2) = wmin: sh(3) = umax: sh(4) = vmax: sh(5) = wmax
    sh(6) = 0: sh(7) = umax - umin: sh(8) = vmax - vmin
    sh(9) = (umin + umax) / 2: sh(10) = (vmin + vmax) / 2
    Footprint = True
    Dim spanU As Double, spanV As Double
    spanU = umax - umin: spanV = vmax - vmin
    If spanU < 0.001 Or spanV < 0.001 Then Exit Function          ' flat sliver: nothing to turn

    ' keep, per bin, the lowest and highest point (a few hundred points are enough for the hull)
    Dim aMinV(NBINS - 1) As Double, aMinU(NBINS - 1) As Double, aMaxV(NBINS - 1) As Double, aMaxU(NBINS - 1) As Double
    Dim bMinU(NBINS - 1) As Double, bMinV(NBINS - 1) As Double, bMaxU(NBINS - 1) As Double, bMaxV(NBINS - 1) As Double
    For bi = 0 To NBINS - 1
        aMinV(bi) = 1E+30: aMaxV(bi) = -1E+30: bMinU(bi) = 1E+30: bMaxU(bi) = -1E+30
    Next
    For q = 0 To np - 3 Step 3
        u = pts(q) * fr(0) + pts(q + 1) * fr(1) + pts(q + 2) * fr(2)
        v = pts(q) * fr(3) + pts(q + 1) * fr(4) + pts(q + 2) * fr(5)
        bi = CLng((u - umin) / spanU * (NBINS - 1))
        If bi < 0 Then bi = 0
        If bi > NBINS - 1 Then bi = NBINS - 1
        If v < aMinV(bi) Then aMinV(bi) = v: aMinU(bi) = u
        If v > aMaxV(bi) Then aMaxV(bi) = v: aMaxU(bi) = u
        bi = CLng((v - vmin) / spanV * (NBINS - 1))
        If bi < 0 Then bi = 0
        If bi > NBINS - 1 Then bi = NBINS - 1
        If u < bMinU(bi) Then bMinU(bi) = u: bMinV(bi) = v
        If u > bMaxU(bi) Then bMaxU(bi) = u: bMaxV(bi) = v
    Next
    Dim ru(4 * NBINS - 1) As Double, rv(4 * NBINS - 1) As Double, nr As Long
    For bi = 0 To NBINS - 1
        If aMinV(bi) < 1E+29 Then ru(nr) = aMinU(bi): rv(nr) = aMinV(bi): nr = nr + 1
        If aMaxV(bi) > -1E+29 Then ru(nr) = aMaxU(bi): rv(nr) = aMaxV(bi): nr = nr + 1
        If bMinU(bi) < 1E+29 Then ru(nr) = bMinU(bi): rv(nr) = bMinV(bi): nr = nr + 1
        If bMaxU(bi) > -1E+29 Then ru(nr) = bMaxU(bi): rv(nr) = bMaxV(bi): nr = nr + 1
    Next
    If nr < 4 Then Exit Function

    ' angle with the smallest rectangle: coarse steps, then fine steps round the best one
    Dim ang As Double, ar As Double, bestAr As Double, bestAng As Double
    Dim u0 As Double, u1 As Double, v0 As Double, v1 As Double
    bestAr = 1E+300
    For ang = 0 To 89.5 Step 0.5
        ExtentsAt ru, rv, nr, ang, u0, u1, v0, v1
        ar = (u1 - u0) * (v1 - v0)
        If ar < bestAr - 0.000000001 Then bestAr = ar: bestAng = ang
    Next
    Dim ang0 As Double
    ang0 = bestAng
    For ang = ang0 - 0.5 To ang0 + 0.5 Step 0.02
        ExtentsAt ru, rv, nr, ang, u0, u1, v0, v1
        ar = (u1 - u0) * (v1 - v0)
        If ar < bestAr - 0.000000001 Then bestAr = ar: bestAng = ang
    Next
    ExtentsAt ru, rv, nr, bestAng, u0, u1, v0, v1

    Dim lr As Double, wr As Double, cu As Double, cv As Double, cx As Double, cy As Double, phi As Double, tmp As Double
    lr = u1 - u0: wr = v1 - v0
    cu = (u0 + u1) / 2: cv = (v0 + v1) / 2
    phi = bestAng
    Do While phi > 45
        phi = phi - 90: tmp = lr: lr = wr: wr = tmp
    Loop
    Do While phi <= -45
        phi = phi + 90: tmp = lr: lr = wr: wr = tmp
    Loop
    cx = cu * Cos(bestAng * PI / 180#) - cv * Sin(bestAng * PI / 180#)
    cy = cu * Sin(bestAng * PI / 180#) + cv * Cos(bestAng * PI / 180#)

    ' worth it only when the turned rectangle is clearly smaller than the envelope box
    Dim envL As Double, envW As Double
    envL = lr * Abs(Cos(phi * PI / 180#)) + wr * Abs(Sin(phi * PI / 180#))
    envW = lr * Abs(Sin(phi * PI / 180#)) + wr * Abs(Cos(phi * PI / 180#))
    If Abs(phi) < 0.3 Then Exit Function
    If (spanU * spanV) < 0.000001 Then Exit Function
    If 1# - (lr * wr) / (spanU * spanV) < 0.03 Then Exit Function
    If envL > spanU * 1.5 + 0.1 Or envW > spanV * 1.5 + 0.1 Then Exit Function
    sh(0) = cx - envL / 2: sh(3) = cx + envL / 2
    sh(1) = cy - envW / 2: sh(4) = cy + envW / 2
    sh(6) = phi: sh(7) = lr: sh(8) = wr
    sh(9) = cx: sh(10) = cy
End Function

Sub ExtentsAt(ByRef ru() As Double, ByRef rv() As Double, ByVal nr As Long, ByVal ang As Double, _
              ByRef u0 As Double, ByRef u1 As Double, ByRef v0 As Double, ByRef v1 As Double)
    Dim ca As Double, sa As Double, q As Long, a As Double, b As Double
    ca = Cos(ang * PI / 180#): sa = Sin(ang * PI / 180#)
    u0 = 1E+30: v0 = 1E+30: u1 = -1E+30: v1 = -1E+30
    For q = 0 To nr - 1
        a = ru(q) * ca + rv(q) * sa
        b = -ru(q) * sa + rv(q) * ca
        If a < u0 Then u0 = a
        If a > u1 Then u1 = a
        If b < v0 Then v0 = b
        If b > v1 Then v1 = b
    Next
End Sub

' Turns the frame (board X and Y) by deg about board Z: x' = cos x + sin y, y' = -sin x + cos y
Sub RotateFrame(ByRef fr() As Double, ByVal deg As Double)
    Dim ca As Double, sa As Double, k As Long, x As Double, y As Double
    ca = Cos(deg * PI / 180#): sa = Sin(deg * PI / 180#)
    If Abs(ca) < 0.000000000001 Then ca = 0
    If Abs(sa) < 0.000000000001 Then sa = 0
    For k = 0 To 2
        x = fr(k): y = fr(3 + k)
        fr(k) = ca * x + sa * y
        fr(3 + k) = -sa * x + ca * y
    Next
End Sub

' ---------------------------------------------------------------- geometry helpers
Function BoxVol(ByVal i As Long) As Double
    BoxVol = (bx(i, 3) - bx(i, 0)) * (bx(i, 4) - bx(i, 1)) * (bx(i, 5) - bx(i, 2))
End Function

Function Flatness(ByVal i As Long) As Double
    Dim a As Double, b As Double, c As Double, t As Double
    a = bx(i, 3) - bx(i, 0): b = bx(i, 4) - bx(i, 1): c = bx(i, 5) - bx(i, 2)
    If a > b Then t = a: a = b: b = t
    If b > c Then t = b: b = c: c = t
    If a > b Then t = a: a = b: b = t
    If a < 0.000001 Then a = 0.000001
    Flatness = b * c / a
End Function

' ---------------------------------------------------------------- matching and text
Sub PutInputs(ByVal ws As Object, ByVal g As Long, ByVal gc As Long, ByVal rec As Variant)
    Dim k As Long
    ws.Cells(g, gc + G_OVNAME).Value = rec(1)
    ws.Cells(g, gc + G_REF).Value = rec(2)
    ws.Cells(g, gc + G_INC).Value = rec(3)
    For k = 0 To 5: ws.Cells(g, gc + G_OVR + k).Value = rec(4 + k): Next
End Sub

Function SameBox(ByVal rec As Variant, ByRef outv() As Variant, ByVal r As Long) As Boolean
    Dim k As Long, a As Variant
    For k = 0 To 5
        a = rec(10 + k)
        If IsEmpty(a) Or IsError(a) Then Exit Function
        If Not IsNumeric(a) Then Exit Function
        If Abs(CDbl(a) - CDbl(outv(r, 2 + k))) > 0.05 Then Exit Function
    Next
    SameBox = True
End Function

Function HasInput(ByVal rec As Variant) As Boolean
    Dim k As Long
    For k = 1 To 9
        If k <> 3 Then                       ' Include alone is not worth reporting
            If PT_Str(rec(k)) <> "" Then HasInput = True: Exit Function
        End If
    Next
End Function

' Fills empty RefDes cells (CAD geometry column F) for bodies whose name key equals the key of a register
' RefDes. A RefDes already used on CAD geometry (typed, restored or manual row) is not given out again.
' Returns the number of cells filled.
Function AutoLinkRefDes(ByVal wb As Object, ByVal wsG As Object, ByVal g0 As Long, ByVal gc As Long, _
                        ByRef outv As Variant, ByVal cnt As Long, ByRef isNewRow() As Boolean, ByRef nOn As Long) As Long
    Dim rReg As Object, rGeoT As Object, regV As Variant, geoV As Variant
    Dim regByKey As Object, usedRef As Object, i As Long, rk As String, bk As String, refD As String
    Set rReg = PT_Rng(wb, "Reg_RefDes")
    Set rGeoT = PT_Rng(wb, "Geo_Table")
    If rReg Is Nothing Or rGeoT Is Nothing Then Exit Function
    Set regByKey = CreateObject("Scripting.Dictionary")
    Set usedRef = CreateObject("Scripting.Dictionary")
    regV = rReg.Value
    For i = 1 To UBound(regV, 1)
        refD = PT_Str(regV(i, 1))
        If refD <> "" Then
            rk = NormKey(refD)
            If rk <> "" And Not regByKey.Exists(rk) Then regByKey.Add rk, refD
        End If
    Next
    If regByKey.Count = 0 Then Exit Function
    geoV = rGeoT.Columns(1 + G_REF).Value
    For i = 1 To UBound(geoV, 1)
        refD = UCase(PT_Str(geoV(i, 1)))
        If refD <> "" And Not usedRef.Exists(refD) Then usedRef.Add refD, 1
    Next
    For i = 1 To cnt - 1
        If PT_Str(wsG.Cells(g0 + i - 1, gc + G_REF).Value) = "" Then
            bk = NormKey(CStr(outv(i, 1)))
            If regByKey.Exists(bk) Then
                refD = regByKey(bk)
                If Not usedRef.Exists(UCase(refD)) Then
                    wsG.Cells(g0 + i - 1, gc + G_REF).Value = refD
                    usedRef.Add UCase(refD), 1
                    AutoLinkRefDes = AutoLinkRefDes + 1
                    If UCase(PT_Str(wsG.Cells(g0 + i - 1, gc + G_INC).Value)) <> "YES" Then
                        wsG.Cells(g0 + i - 1, gc + G_INC).Value = "Yes"
                        nOn = nOn + 1
                    End If
                End If
            End If
        End If
    Next
End Function

Function JoinMsgs(ByVal moreWhere As String) As String
    Dim t As Long, s As String
    For t = 1 To msgs.Count
        If t > 8 Then s = s & moreWhere: Exit For
        s = s & "- " & msgs(t) & vbCrLf
    Next
    JoinMsgs = s
End Function

' ---------------------------------------------------------------- workbook preparation
' First run on a workbook: adds the rotation columns and makes the import level "every solid body".
' Returns "" when nothing was needed, a message when done, "ERR:<reason>" when it failed.
Function UpgradeWorkbook(ByVal wb As Object, ByVal xl As Object) As String
    Dim rImp As Object, rGeo As Object, wsI As Object, wsG As Object, rTop As Object
    Dim r0 As Long, c0 As Long, cRot As Long, cG As Long, hdrI As Long, hdrG As Long, cap As Long
    Dim lastImp As Long, lastGeo As Long, stepName As String, f As String, k As Long, colL As String, calcWas As Long
    Dim msgOut As String
    If Not PT_Rng(wb, "Imp_Rot") Is Nothing Then Exit Function
    Set rImp = PT_Rng(wb, "Imp_Table")
    Set rGeo = PT_Rng(wb, "Geo_Table")
    Set wsI = rImp.Worksheet
    Set wsG = rGeo.Worksheet
    calcWas = xl.Calculation
    On Error GoTo fail
    xl.Calculation = -4135                           ' manual
    xl.ScreenUpdating = False
    stepName = wsI.Name

    ' ---- STEP import: Rot (deg), Own L, Own W right of the settings column
    r0 = rImp.Row: c0 = rImp.Column: hdrI = r0 - 1
    cRot = c0 + 12                                   ' column N when the table starts in B
    lastImp = r0 + rImp.Rows.Count - 1
    wsI.Cells(hdrI, cRot).Value = "Rot (deg)"
    wsI.Cells(hdrI, cRot + 1).Value = "Own L"
    wsI.Cells(hdrI, cRot + 2).Value = "Own W"
    wsI.Cells(hdrI, c0 + 8).Copy
    wsI.Range(wsI.Cells(hdrI, cRot), wsI.Cells(hdrI, cRot + 2)).PasteSpecial -4122
    wsI.Cells(r0, c0 + 8).Copy
    wsI.Range(wsI.Cells(r0, cRot), wsI.Cells(lastImp, cRot + 2)).PasteSpecial -4122
    xl.CutCopyMode = False
    wsI.Range(wsI.Cells(1, cRot), wsI.Cells(1, cRot + 2)).EntireColumn.ColumnWidth = 10
    AddNote wsI.Cells(hdrI, cRot), "Rotation of the body in the board plane, degrees, counter-clockwise seen from the top (Z). " & _
        "0 = square to the board edges. Written by ImportStep from the body surface. Xmin, Ymin, L, W stay the axis-aligned envelope of the turned rectangle."
    AddNote wsI.Cells(hdrI, cRot + 1), "Length of the turned rectangle along its own X direction (mm). Equals L when Rot is 0."
    AddNote wsI.Cells(hdrI, cRot + 2), "Width of the turned rectangle along its own Y direction (mm). Equals W when Rot is 0."
    wb.Names.Add "Imp_Rot", "='" & stepName & "'!" & wsI.Range(wsI.Cells(r0, cRot), wsI.Cells(lastImp, cRot + 2)).Address

    ' ---- CAD geometry: AI:AK mirror the three columns
    cG = rGeo.Column + G_ROT
    hdrG = rGeo.Row - 1
    cap = BodyCapacity(rGeo)
    lastGeo = rGeo.Row + rGeo.Rows.Count - 1
    wsG.Cells(hdrG, cG).Value = "Rot (deg)"
    wsG.Cells(hdrG, cG + 1).Value = "Own L"
    wsG.Cells(hdrG, cG + 2).Value = "Own W"
    wsG.Cells(hdrG, rGeo.Column + 12).Copy
    wsG.Range(wsG.Cells(hdrG, cG), wsG.Cells(hdrG, cG + 2)).PasteSpecial -4122
    wsG.Cells(rGeo.Row, rGeo.Column + G_IMP).Copy
    wsG.Range(wsG.Cells(rGeo.Row, cG), wsG.Cells(rGeo.Row + cap - 1, cG + 2)).PasteSpecial -4122
    xl.CutCopyMode = False
    wsG.Range(wsG.Cells(1, cG), wsG.Cells(1, cG + 2)).EntireColumn.ColumnWidth = 10
    For k = 0 To 2
        colL = Split(wsI.Cells(1, cRot + k).Address(True, False), "$")(0)
        f = "=IF(ISNUMBER('" & stepName & "'!$" & colL & (r0 + 1) & "),'" & stepName & "'!$" & colL & (r0 + 1) & ","""")"
        wsG.Range(wsG.Cells(rGeo.Row, cG + k), wsG.Cells(rGeo.Row + cap - 1, cG + k)).Formula = f
    Next
    AddNote wsG.Cells(hdrG, cG), "Imported rotation of the body in the board plane (degrees, from STEP import). BuildBoard builds a turned rectangle when this is not 0 " & _
        "and L and W (Used) still equal the envelope of the turned rectangle. If you override L or W, the part is built as a plain box."
    AddNote wsG.Cells(hdrG, cG + 1), "Imported length of the turned rectangle along its own X direction (mm)."
    AddNote wsG.Cells(hdrG, cG + 2), "Imported width of the turned rectangle along its own Y direction (mm)."
    wb.Names("Geo_Table").RefersTo = "='" & wsG.Name & "'!" & wsG.Range(wsG.Cells(rGeo.Row, rGeo.Column), wsG.Cells(lastGeo, cG + 2)).Address

    ' ---- default import level: every solid body (was: one row per component)
    Set rTop = PT_Rng(wb, "Imp_TopOnly")
    msgOut = "Workbook prepared once: Rot columns added (STEP import N:P, CAD geometry AI:AK)."
    If Not rTop Is Nothing Then
        If UCase(PT_Str(rTop.Cells(1, 1).Value)) = "YES" Then
            rTop.Cells(1, 1).Value = "No"
            msgOut = msgOut & " Setting 'one row per top-level component' changed to No, so every solid body is imported."
        End If
        rTop.Cells(1, 1).Offset(0, -1).Value = "One row per top-level component? (No = every solid body)"
        AddNote rTop.Cells(1, 1), "No (default) = every solid body gets its own row, so small parts are not lost inside a bigger component. " & _
            "Yes = one row per top-level component (bodies of a component are merged into one box). Read by ImportStep."
    End If
    xl.Calculation = calcWas
    xl.ScreenUpdating = True
    UpgradeWorkbook = msgOut
    Exit Function
fail:
    UpgradeWorkbook = "ERR:" & Err.Description
    On Error Resume Next
    xl.CutCopyMode = False
    xl.Calculation = calcWas
    xl.ScreenUpdating = True
End Function

Sub AddNote(ByVal cel As Object, ByVal txt As String)
    On Error Resume Next
    cel.Comment.Delete
    cel.AddComment txt
End Sub

' Number of body rows on CAD geometry: rows from the top of Geo_Table whose # cell is a number
' (manual rows below them have text there).
Function BodyCapacity(ByVal rGeo As Object) As Long
    Dim r As Long
    For r = 1 To rGeo.Rows.Count
        If Not PT_IsNum(rGeo.Cells(r, 1).Value) Then Exit For
        BodyCapacity = r
    Next
End Function

' Makes room for need bodies. Inserts rows into CAD geometry before its manual rows (Excel then moves every formula,
' range and name that points there), copies the last body row into them, and lengthens the STEP import tables.
' Returns the number of body rows added (0 = nothing to do), or -1 on failure (PT_LastErr).
Function EnsureCapacity(ByVal wb As Object, ByVal xl As Object, ByVal need As Long) As Long
    Dim rGeo As Object, rImp As Object, rPrev As Object, wsG As Object, cur As Long, newCap As Long, nAdd As Long
    Dim first As Long, c0 As Long, k As Long, calcWas As Long, target As Long, lastA As Long, lastB As Long, cnt As Object
    Set rGeo = PT_Rng(wb, "Geo_Table")
    Set rImp = PT_Rng(wb, "Imp_Table")
    Set rPrev = PT_Rng(wb, "ImpPrev_Table")
    cur = BodyCapacity(rGeo)
    newCap = cur
    If need > cur Then newCap = (need \ 50 + 1) * 50
    target = rImp.Row + newCap                                   ' PCB row + newCap body rows
    lastA = rImp.Row + rImp.Rows.Count - 1
    lastB = rPrev.Row + rPrev.Rows.Count - 1
    If newCap = cur And lastA >= target And lastB >= target And lastA = lastB Then Exit Function
    If lastA > target Then target = lastA
    If lastB > target Then target = lastB
    calcWas = xl.Calculation
    PT_LastErr = ""
    On Error GoTo fail
    xl.Calculation = -4135
    xl.ScreenUpdating = False
    nAdd = newCap - cur
    If nAdd > 0 Then
        Set wsG = rGeo.Worksheet
        first = rGeo.Row + cur
        c0 = rGeo.Column
        wsG.Rows(first & ":" & (first + nAdd - 1)).Insert -4121
        wsG.Rows(first - 1).Copy
        wsG.Rows(first & ":" & (first + nAdd - 1)).PasteSpecial -4104
        xl.CutCopyMode = False
        wsG.Cells(first, c0 + G_OVNAME).Resize(nAdd, 1).ClearContents          ' inputs must not be copied
        wsG.Cells(first, c0 + G_REF).Resize(nAdd, 2).ClearContents
        wsG.Cells(first, c0 + G_OVR).Resize(nAdd, 6).ClearContents
        For k = 0 To nAdd - 1: wsG.Cells(first + k, c0).Value = cur + k + 1: Next
        ' The CHECKS count of imported bodies ends exactly at the old last body row, so Excel does not stretch it.
        Set cnt = wsG.Range("H1:H25").Find("Bodies imported from STEP", , -4163, 1)
        If Not cnt Is Nothing Then cnt.Offset(0, 5).Formula = "=COUNTIF(INDEX(Geo_Table,0,2),""?*"")"
    End If
    ResizeName wb, xl, "Imp_Table", target
    ResizeName wb, xl, "ImpPrev_Table", target
    ResizeName wb, xl, "Imp_Rot", target
    xl.Calculation = calcWas
    xl.ScreenUpdating = True
    EnsureCapacity = nAdd
    Exit Function
fail:
    PT_LastErr = Err.Description
    EnsureCapacity = -1
    On Error Resume Next
    xl.CutCopyMode = False
    xl.Calculation = calcWas
    xl.ScreenUpdating = True
End Function

' Lengthens a named table so it ends at newLastRow (formats of its last row are copied down). Never shortens.
Sub ResizeName(ByVal wb As Object, ByVal xl As Object, ByVal nameS As String, ByVal newLastRow As Long)
    Dim r As Object, ws As Object, lastR As Long, c1 As Long
    Set r = PT_Rng(wb, nameS)
    If r Is Nothing Then Exit Sub
    Set ws = r.Worksheet
    lastR = r.Row + r.Rows.Count - 1
    If lastR >= newLastRow Then Exit Sub
    c1 = r.Column + r.Columns.Count - 1
    On Error Resume Next
    ws.Range(ws.Cells(lastR, r.Column), ws.Cells(lastR, c1)).Copy
    ws.Range(ws.Cells(lastR + 1, r.Column), ws.Cells(newLastRow, c1)).PasteSpecial -4122
    xl.CutCopyMode = False
    On Error GoTo 0
    wb.Names(nameS).RefersTo = "='" & ws.Name & "'!" & ws.Range(ws.Cells(r.Row, r.Column), ws.Cells(newLastRow, c1)).Address
End Sub

' Import level of the last run ("bodies" or "components"), kept in the workbook name Imp_Level. "" = unknown.
Function GetLevelUsed(ByVal wb As Object) As String
    Dim s As String
    On Error Resume Next
    s = wb.Names("Imp_Level").RefersTo
    On Error GoTo 0
    s = Replace(Replace(s, "=", ""), """", "")
    GetLevelUsed = s
End Function

Sub SetLevelUsed(ByVal wb As Object, ByVal lvl As String)
    wb.Names.Add "Imp_Level", "=""" & lvl & """"
End Sub

' ---------------------------------------------------------------- texts
Function FrameText(ByRef fr() As Double) As String
    FrameText = "Board X = model " & AxisText(fr, 0) & ", Board Y = model " & AxisText(fr, 3) & ", Board Z = model " & AxisText(fr, 6) & "."
End Function

Function AxisText(ByRef fr() As Double, ByVal o As Long) As String
    Dim k As Long, big As Long
    For k = 1 To 2
        If Abs(fr(o + k)) > Abs(fr(o + big)) Then big = k
    Next
    If Abs(fr(o + big)) > 0.9999 Then
        AxisText = IIf(fr(o + big) > 0, "+", "-") & Mid("XYZ", big + 1, 1)
    Else
        AxisText = "direction (" & Format(fr(o), "0.000") & ", " & Format(fr(o + 1), "0.000") & ", " & Format(fr(o + 2), "0.000") & ")"
    End If
End Function

' tailS follows the number of change rows, e.g. " rows added to the Change log tab. Check them there."
Function ChangeSummary(ByVal counts As Object, ByVal nLog As Long, ByVal hadOld As Boolean, ByVal levelChanged As Boolean, _
                       ByVal tailS As String) As String
    Dim s As String, it As Variant
    If levelChanged Then ChangeSummary = "not compared (import level changed, see messages).": Exit Function
    If Not hadOld Then ChangeSummary = "none (first import).": Exit Function
    If nLog = 0 Then ChangeSummary = "none.": Exit Function
    For Each it In counts.Keys
        s = s & counts(it) & " " & it & ", "
    Next
    ChangeSummary = Left(s, Len(s) - 2) & "." & vbCrLf & nLog & tailS
End Function

