Option Strict Off
Option Explicit On

Imports System
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports SW = SolidWorks.Interop.sldworks

' Port of src/ImportStep.bas. Same behaviour and messages as the VBA macro.
'
' Reads a PCB assembly (a STEP file, or the model already open in SolidWorks) and writes every body's size, position
' and rotation into the 'STEP import' tab of the workbook. Works for any board.
' What is imported (setting Imp_TopOnly): No = every solid body gets its own row (default); Yes = one row per top-level
' component. Board frame (mm): origin at the PCB corner on its bottom face, X along the longest board edge, Y along the
' other edge, Z out of the side that carries most of the parts; turned with the board when the board is turned in the model.
' Rotation of each body in the board plane: Rot (deg), Own L, Own W. Xmin, Ymin, L, W stay the axis-aligned envelope.
' Change report: the last import is copied to 'STEP import (previous)'; every body added, removed, moved, resized or
' renamed gets a row on the Change log tab. Inputs on 'CAD geometry' follow their body by key.
Public Module ImportStep

    Private Const MACRO_NAME As String = "ImportStep"
    Private Const MAX_ITEMS As Integer = 5000
    Private Const swSolidBody As Integer = 0
    Private Const swSheetBody As Integer = 1
    Private Const swDocPART As Integer = 1
    Private Const swDocASSEMBLY As Integer = 2
    Private Const PI As Double = 3.14159265358979
    Private Const NBINS As Integer = 256

    ' CAD geometry columns, as offsets from the first column of Geo_Table (column B)
    Private Const G_NAME As Integer = 1          ' C imported name
    Private Const G_OVNAME As Integer = 2        ' D name override
    Private Const G_REF As Integer = 4           ' F RefDes
    Private Const G_INC As Integer = 5           ' G Include
    Private Const G_IMP As Integer = 6           ' H..M imported box
    Private Const G_OVR As Integer = 13          ' O..T overrides
    Private Const G_ROT As Integer = 33          ' AI imported rotation (deg), AJ own L, AK own W
    Private Const G_ROTCOLS As Integer = 3

    Private nm() As String
    Private bx(,) As Double          ' (item, 0..5) = xmin ymin zmin xmax ymax zmax, model space, mm
    Private rbx(,) As Double         ' same, but in the part's own space (before the component transform), mm
    Private srcComp() As Object      ' owning component (Nothing for a part document)
    Private srcBody() As Object      ' the body (Nothing when the item is a whole component)
    Private srcMode() As Integer     ' how to turn part-space points into model space: 0 none, 1 matrix, 2 transposed matrix
    Private srcXf(,) As Double       ' (item, 0..12) component transform: 9 rotation, 3 translation (m), scale
    Private n As Integer
    Private msgs As Collection
    Private capRows As Integer       ' body rows available on CAD geometry
    Private nSheetBodies As Integer  ' surface bodies seen (not imported)
    Private itemsCapped As Boolean

    Public Sub Main()
        msgs = New Collection()
        PT_LogReset()
        nSheetBodies = 0 : itemsCapped = False : n = 0

        ' ---------- workbook ----------
        Dim xl As Object, wb As Object, wsI As Object, wsG As Object
        xl = PT_GetExcel()
        wb = PT_PickWorkbook(xl, MACRO_NAME, "STEP import|CAD geometry")
        If wb Is Nothing Then Exit Sub
        If Not PT_NeedNames(wb, "Imp_Table|Imp_PcbName|Imp_MinSize|Imp_TopOnly|Imp_Model|Imp_Time|Imp_Frame|Imp_Messages|" &
                                "ImpPrev_Table|ImpPrev_Model|ImpPrev_Time|Geo_Table|Log_Header|Set_TolPos|Set_TolSize", MACRO_NAME) Then Exit Sub

        Dim upg As String
        upg = UpgradeWorkbook(wb, xl)
        If Left(upg, 4) = "ERR:" Then
            MsgBox("ImportStep did not run. The workbook could not be prepared for rotation columns:" & vbCrLf & Mid(upg, 5) & vbCrLf &
                   "Close any cell you are editing in Excel and run again.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        If upg <> "" Then msgs.Add(upg)

        Dim rImp As Object, rGeo As Object, rMsg As Object, rRot As Object
        rImp = PT_Rng(wb, "Imp_Table")
        rGeo = PT_Rng(wb, "Geo_Table")
        rMsg = PT_Rng(wb, "Imp_Messages")
        rRot = PT_Rng(wb, "Imp_Rot")
        wsI = rImp.Worksheet
        wsG = rGeo.Worksheet

        Dim pcbKey As String, minSize As Double, perBody As Boolean, tolPos As Double, tolSize As Double
        pcbKey = PT_NameStr(wb, "Imp_PcbName", "PCB")
        minSize = PT_NameNum(wb, "Imp_MinSize", 0)
        perBody = (UCase(PT_NameStr(wb, "Imp_TopOnly", "Yes")) = "NO")
        tolPos = PT_NameNum(wb, "Set_TolPos", 0.05)
        tolSize = PT_NameNum(wb, "Set_TolSize", 0.05)

        ' ---------- model ----------
        Dim doc As Object, srcName As String = ""
        doc = GetModel(srcName)
        If doc Is Nothing Then Exit Sub
        xl.StatusBar = "ImportStep: reading the bodies of " & srcName & " ..."
        Dim swT As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()
        CollectItems(doc, perBody)
        Log.Write("ImportStep: " & n & " items collected in " & (swT.ElapsedMilliseconds / 1000.0).ToString("0.0") & " s")
        If n = 0 Then
            xl.StatusBar = False
            MsgBox("No solid bodies found in " & srcName & ". Nothing was written.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        If itemsCapped Then msgs.Add("The model has more than " & MAX_ITEMS & " bodies; only the first " & MAX_ITEMS & " were read.")
        If nSheetBodies > 0 Then msgs.Add(nSheetBodies & " surface (sheet) bodies were not imported. Only solid bodies are.")

        ' ---------- PCB ----------
        Dim pcb As Integer, i As Integer, k As Integer, best As Double, v As Double
        pcb = -1 : best = -1
        For i = 0 To n - 1
            If InStr(1, nm(i), pcbKey, vbTextCompare) > 0 Then
                v = BoxVol(i)
                If v > best Then
                    best = v
                    pcb = i
                End If
            End If
        Next
        If pcb < 0 Then
            For i = 0 To n - 1
                v = Flatness(i)
                If v > best Then
                    best = v
                    pcb = i
                End If
            Next
            msgs.Add("No body name contains '" & pcbKey & "'. PCB picked by shape: " & nm(pcb) & ". Check it is the board, or change the PCB name setting.")
        End If

        ' ---------- board frame: axes ----------
        Dim d(2) As Double, nAx As Integer, xAx As Integer, midN As Double, up As Double, dn As Double, c As Double
        For k = 0 To 2
            d(k) = bx(pcb, k + 3) - bx(pcb, k)
        Next
        nAx = 0 : xAx = 0
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
        For k = 0 To 2
            fr(k) = ex(k) : fr(3 + k) = ey(k) : fr(6 + k) = ez(k)
        Next

        ' ---------- board frame: turn it with the PCB if the PCB is not square to the model axes ----------
        Dim sh(10) As Double, shPcb(10) As Double, havePcb As Boolean, turn As Double
        Dim pPts() As Double = Nothing, pN As Integer         ' PCB surface points, read once (the board is the biggest tessellation)
        swT.Restart()
        pN = ItemPoints(pcb, pPts)
        Log.Write("ImportStep: PCB points " & pN \ 3 & " in " & (swT.ElapsedMilliseconds / 1000.0).ToString("0.0") & " s")
        havePcb = (pN >= 9)
        If havePcb Then havePcb = Footprint(pPts, pN, fr, sh)
        If havePcb Then
            If Math.Abs(sh(6)) >= 0.3 Then
                turn = sh(6)
                RotateFrame(fr, turn)
                havePcb = Footprint(pPts, pN, fr, sh)
            End If
            If havePcb Then
                If sh(8) > sh(7) * 1.0001 Then            ' long edge ended up on Y: turn 90 degrees
                    RotateFrame(fr, 90)
                    turn = turn + 90
                    havePcb = Footprint(pPts, pN, fr, sh)
                End If
            End If
        End If
        pPts = Nothing
        For k = 0 To 10
            shPcb(k) = sh(k)
        Next
        If turn <> 0 Then msgs.Add("The board is turned " & Round(turn, 2) & " degrees about its normal in the model file. The board frame was turned with it.")
        If Not havePcb Then msgs.Add("The PCB surface points could not be read. The board frame follows the model axes.")

        ' ---------- shape of every item in the board frame ----------
        Dim shp(,) As Double, nNoPts As Integer, nTurned As Integer
        ReDim shp(n - 1, 10)
        swT.Restart() : tPoints = 0
        For i = 0 To n - 1
            If i Mod 5 = 0 Then xl.StatusBar = "ImportStep: measuring body " & (i + 1) & " of " & n & " ..."
            If i = pcb And havePcb Then
                For k = 0 To 10
                    shp(i, k) = shPcb(k)
                Next
            ElseIf ItemShape(i, fr, sh) Then
                For k = 0 To 10
                    shp(i, k) = sh(k)
                Next
            Else
                BoxShape(i, fr, sh)
                For k = 0 To 10
                    shp(i, k) = sh(k)
                Next
                nNoPts = nNoPts + 1
            End If
            If Math.Abs(shp(i, 6)) >= 0.3 Then nTurned = nTurned + 1
        Next
        xl.StatusBar = False
        Log.Write("ImportStep: shapes of " & n & " items in " & (swT.ElapsedMilliseconds / 1000.0).ToString("0.0") & " s (points " & tPoints.ToString("0.0") & " s)")
        If nNoPts > 0 Then msgs.Add(nNoPts & " bodies gave no surface points; their plain bounding box was used (no rotation).")

        Dim bmin(,) As Double, bmax(,) As Double
        ReDim bmin(n - 1, 2) : ReDim bmax(n - 1, 2)
        For i = 0 To n - 1
            For k = 0 To 2
                bmin(i, k) = shp(i, k) : bmax(i, k) = shp(i, 3 + k)
            Next
        Next
        Dim org(2) As Double
        For k = 0 To 2
            org(k) = bmin(pcb, k)
        Next
        Dim tPcb As Double = bmax(pcb, 2) - bmin(pcb, 2)

        Dim fill As Double, bodyVol As Double, boxV As Double
        bodyVol = ItemVolume(pcb)
        boxV = (bmax(pcb, 0) - bmin(pcb, 0)) * (bmax(pcb, 1) - bmin(pcb, 1)) * tPcb
        If bodyVol > 0 And boxV > 0 Then
            fill = bodyVol / boxV
            If fill < 0.9 Then msgs.Add("PCB fills only " & Format(fill, "0%") & " of its box: the board has large cutouts or is not rectangular. " &
                "Check L and W on CAD geometry and override if needed.")
        End If

        ' ---------- rows: PCB first, then components that pass the size filter ----------
        Dim order() As Integer, cnt As Integer, skipped As Integer, L As Double, W As Double
        ReDim order(n - 1)
        order(0) = pcb : cnt = 1
        For i = 0 To n - 1
            If i <> pcb Then
                L = bmax(i, 0) - bmin(i, 0) : W = bmax(i, 1) - bmin(i, 1)
                If L < minSize And W < minSize Then
                    skipped = skipped + 1
                Else
                    order(cnt) = i : cnt = cnt + 1
                End If
            End If
        Next
        If skipped > 0 Then msgs.Add(skipped & " bodies smaller than " & minSize & " mm in both L and W were skipped (size filter setting).")

        ' ---------- room for the bodies: grow CAD geometry and STEP import when needed ----------
        Dim added As Integer
        added = EnsureCapacity(wb, xl, cnt - 1)
        If added < 0 Then
            MsgBox("ImportStep stopped while making room for " & (cnt - 1) & " bodies on the CAD geometry tab:" & vbCrLf & PT_LastErr & vbCrLf &
                   "Nothing else was written. Close any cell you are editing in Excel and run again.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        If added > 0 Then msgs.Add("CAD geometry and STEP import were extended by " & added & " body rows (Excel moved every formula that points to them).")
        rImp = PT_Rng(wb, "Imp_Table")
        rGeo = PT_Rng(wb, "Geo_Table")
        rRot = PT_Rng(wb, "Imp_Rot")
        wsI = rImp.Worksheet
        wsG = rGeo.Worksheet
        capRows = BodyCapacity(rGeo)
        If cnt - 1 > capRows Or cnt > CInt(rImp.Rows.Count) Then
            MsgBox(CStr(cnt - 1) & " bodies to import; CAD geometry holds " & capRows & "." & vbCrLf & "Nothing was written.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If

        ' outv columns: 0 #, 1 name, 2..4 min, 5..7 size, 8 side, 9 key, 10 rot, 11 own L, 12 own W
        Dim outv(,) As Object, r As Integer, j As Integer, usedKeys As New Dict(), rd As Double
        ReDim outv(cnt - 1, 12)
        For r = 0 To cnt - 1
            j = order(r)
            outv(r, 0) = IIf(r = 0, "PCB", r)
            outv(r, 1) = nm(j)
            For k = 0 To 2
                outv(r, 2 + k) = Round(bmin(j, k) - org(k), 3)
                outv(r, 5 + k) = Round(bmax(j, k) - bmin(j, k), 3)
            Next
            If r = 0 Then
                outv(r, 8) = "Board"
                outv(r, 9) = PT_UniqueKey(usedKeys, "pcb")
            Else
                If ((bmin(j, 2) + bmax(j, 2)) / 2 - org(2)) >= tPcb / 2 Then outv(r, 8) = "Top" Else outv(r, 8) = "Bottom"
                outv(r, 9) = PT_UniqueKey(usedKeys, PT_Key(nm(j)))
            End If
            rd = shp(j, 6)
            If r = 0 Or Math.Abs(rd) < 0.3 Then
                outv(r, 10) = 0 : outv(r, 11) = outv(r, 5) : outv(r, 12) = outv(r, 6)
            Else
                outv(r, 10) = Round(rd, 3) : outv(r, 11) = Round(shp(j, 7), 3) : outv(r, 12) = Round(shp(j, 8), 3)
            End If
        Next

        ' ---------- save current CAD geometry inputs, keyed by body key ----------
        ' rec: 0 imported name, 1 name override, 2 RefDes, 3 Include, 4..9 overrides,
        '      10..15 imported box, used to match by position when names differ
        Dim saved As New Dict(), refByKey As New Dict(), geoKeys As New Dict()
        Dim coll As Collection
        Dim old(,) As Object, rec As Object(), key As String
        old = ToTable(rGeo.Resize(capRows, G_OVR + 6).Value)        ' 1-based, column 1 = #
        For r = 1 To UBound(old, 1)
            If IsError(old(r, 1 + G_NAME)) Then
            ElseIf PT_Str(old(r, 1 + G_NAME)) <> "" Then
                rec = Arr(CStr(old(r, 1 + G_NAME)), old(r, 1 + G_OVNAME), old(r, 1 + G_REF), old(r, 1 + G_INC),
                          old(r, 1 + G_OVR), old(r, 2 + G_OVR), old(r, 3 + G_OVR), old(r, 4 + G_OVR), old(r, 5 + G_OVR), old(r, 6 + G_OVR),
                          old(r, 1 + G_IMP), old(r, 2 + G_IMP), old(r, 3 + G_IMP), old(r, 4 + G_IMP), old(r, 5 + G_IMP), old(r, 6 + G_IMP))
                key = PT_UniqueKey(geoKeys, PT_Key(CStr(old(r, 1 + G_NAME))))
                If Not saved.Exists(key) Then saved.Add(key, New Collection())
                coll = saved(key)
                coll.Add(rec)
                If Not refByKey.Exists(key) Then refByKey.Add(key, PT_Str(old(r, 1 + G_REF)))
            End If
        Next

        ' ---------- change report: last import vs this one ----------
        Dim oldImp(,) As Object, oldModel As Object, oldTime As Object
        oldImp = ToTable(rImp.Value)
        oldModel = PT_Rng(wb, "Imp_Model").Value
        oldTime = PT_Rng(wb, "Imp_Time").Value
        Dim counts As New Dict()
        Dim hadOld As Boolean, newLevel As String, oldLevel As String, levelChanged As Boolean, hadRows As Boolean
        newLevel = IIf(perBody, "bodies", "components")
        oldLevel = GetLevelUsed(wb)
        For r = 2 To UBound(oldImp, 1)
            If PT_Str(oldImp(r, 2)) <> "" Then
                hadRows = True
                Exit For
            End If
        Next
        If oldLevel = "" And hadRows Then oldLevel = "components"       ' workbooks from before this version
        levelChanged = (hadRows And oldLevel <> newLevel)
        If levelChanged Then
            msgs.Add("Import level changed from " & oldLevel & " to " & newLevel & ": the change report was skipped once, because nearly every row is new. " &
                     "Inputs on CAD geometry were kept by name or position where the body still exists.")
        Else
            hadOld = CompareImports(oldImp, outv, cnt, refByKey, tolPos, tolSize, counts)
        End If

        ' ---------- write ----------
        On Error GoTo writeFail
        xl.ScreenUpdating = False
        PT_Rng(wb, "ImpPrev_Table").Value = oldImp
        PT_Rng(wb, "ImpPrev_Model").Value = oldModel
        PT_Rng(wb, "ImpPrev_Time").Value = oldTime

        rImp.ClearContents()
        rMsg.ClearContents()
        If rRot IsNot Nothing Then rRot.ClearContents()
        Dim outW(,) As Object, outR(,) As Object
        Dim impCols As Integer = CInt(rImp.Columns.Count)
        outW = NewTable(cnt, impCols)
        For r = 0 To cnt - 1
            For k = 0 To 9
                If k < impCols Then outW(r + 1, k + 1) = outv(r, k)
            Next
        Next
        rImp.Resize(cnt, impCols).Value = outW
        If rRot IsNot Nothing Then
            outR = NewTable(cnt, G_ROTCOLS)
            For r = 0 To cnt - 1
                For k = 1 To G_ROTCOLS
                    outR(r + 1, k) = outv(r, 9 + k)
                Next
            Next
            rRot.Resize(cnt, G_ROTCOLS).Value = outR
        End If
        PT_Rng(wb, "Imp_Model").Value = srcName
        PT_Rng(wb, "Imp_Time").Value = NowText()
        PT_Rng(wb, "Imp_Frame").Value = FrameText(fr) & " Origin at the PCB corner on its bottom face. Units mm."
        SetLevelUsed(wb, newLevel)

        Dim g0 As Integer, gc As Integer
        g0 = CInt(rGeo.Row) : gc = CInt(rGeo.Column)
        wsG.Cells(g0, gc + G_OVNAME).Resize(capRows, 1).ClearContents()
        wsG.Cells(g0, gc + G_REF).Resize(capRows, 2).ClearContents()
        wsG.Cells(g0, gc + G_OVR).Resize(capRows, 6).ClearContents()

        Dim assigned() As Boolean, byName As Integer, byBox As Integer, hit As Boolean
        Dim t As Integer, s As String
        ReDim assigned(cnt - 1)
        Dim isNewRow() As Boolean
        ReDim isNewRow(cnt - 1)
        ' pass 1: same key
        For r = 1 To cnt - 1
            key = CStr(outv(r, 9))
            If saved.Exists(key) Then
                coll = saved(key)
                If coll.Count > 0 Then
                    rec = coll(1) : coll.Remove(1)
                    PutInputs(wsG, g0 + r - 1, gc, rec)
                    assigned(r) = True : byName = byName + 1
                End If
            End If
        Next
        ' pass 2: same box (position and size within 0.05 mm), for names that changed
        For r = 1 To cnt - 1
            If Not assigned(r) Then
                hit = False
                For Each kk As Object In saved.Keys
                    coll = saved(kk)
                    For t = 1 To coll.Count
                        rec = coll(t)
                        If SameBox(rec, outv, r) Then
                            coll.Remove(t)
                            PutInputs(wsG, g0 + r - 1, gc, rec)
                            assigned(r) = True : byBox = byBox + 1 : hit = True
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
        Dim restored As Integer = byName + byBox
        If byBox > 0 Then msgs.Add(byBox & " rows kept their inputs by matching position and size, because the body name changed.")

        ' pass 3: empty RefDes cells get the register RefDes whose key equals the body name
        ' (ECAD exports name each body by its designator, e.g. "U1"). Typed RefDes cells are never changed.
        Dim nAuto As Integer, nOn As Integer
        nAuto = AutoLinkRefDes(wb, wsG, g0, gc, outv, cnt, isNewRow, nOn)
        If nAuto > 0 Then msgs.Add(nAuto & " RefDes filled on CAD geometry because the body name equals a register RefDes. Check column F.")
        If nOn > 0 Then msgs.Add(nOn & " bodies linked to a register RefDes for the first time were set to Include = Yes (column G).")

        ' inputs whose body is no longer in the model
        For Each kk As Object In saved.Keys
            coll = saved(kk)
            For t = 1 To coll.Count
                rec = coll(t)
                If HasInput(rec) Then
                    s = "Not in this model, inputs not applied: " & CStr(rec(0)) & "  [name " & CStr(rec(1)) & ", RefDes " & CStr(rec(2)) &
                        ", Include " & CStr(rec(3)) & ", overrides " & CStr(rec(4)) & "/" & CStr(rec(5)) & "/" & CStr(rec(6)) & "/" &
                        CStr(rec(7)) & "/" & CStr(rec(8)) & "/" & CStr(rec(9)) & "]"
                    msgs.Add(s)
                End If
            Next
        Next

        Dim nLog As Integer
        nLog = PT_LogFlush(wb)
        If Not hadOld And Not levelChanged Then msgs.Add("First import into this workbook: nothing to compare, no Change log rows.")
        If nTurned > 0 Then msgs.Add(nTurned & " bodies are turned in the board plane (Rot column on STEP import). BuildBoard builds them as turned rectangles.")
        If cnt > 1 Then msgs.Add((cnt - 1) & " body rows have Include = No unless you had set Yes before. Set Yes on CAD geometry (column G) for the parts to build.")

        If msgs.Count = 0 Then msgs.Add("No warnings.")
        For t = 1 To msgs.Count
            If t <= CInt(rMsg.Rows.Count) Then rMsg.Cells(t, 1).Value = msgs(t)
        Next
        xl.ScreenUpdating = True
        xl.Calculate()
        Dim chipRows As Integer
        chipRows = PT_FillChipLists(wb)
        If chipRows > 0 Then msgs.Add(chipRows & " chip rows written to the Spreader and Sim setup tabs (they were empty).")
        wb.Save()

        MsgBox("Written and saved to:" & vbCrLf & CStr(wb.FullName) & vbCrLf & "(tabs STEP import, CAD geometry, Change log)" & vbCrLf & vbCrLf &
               "Imported " & (cnt - 1) & " bodies plus the PCB (" & Round(bmax(pcb, 0) - bmin(pcb, 0), 2) & " x " &
               Round(bmax(pcb, 1) - bmin(pcb, 1), 2) & " x " & Round(tPcb, 2) & " mm), level: " & newLevel & "." & vbCrLf &
               nTurned & " bodies turned in the board plane." & vbCrLf &
               restored & " rows kept their CAD geometry inputs (" & byName & " by key, " & byBox & " by position)." & vbCrLf & vbCrLf &
               "Changes since the last import: " & ChangeSummary(counts, nLog, hadOld, levelChanged) & vbCrLf & vbCrLf &
               "Messages:" & vbCrLf & JoinMsgs(), vbInformation, MACRO_NAME)
        Exit Sub
writeFail:
        xl.ScreenUpdating = True
        xl.StatusBar = False
        MsgBox("Stopped while writing to Excel: " & Err.Description & vbCrLf &
               "The workbook was not saved. Close any cell you are editing in Excel, close the workbook without saving, and run again.",
               vbExclamation, MACRO_NAME)
    End Sub

    ' ---------------------------------------------------------------- change report
    ' Compares the previous STEP import table (oldImp, 1-based, columns B..K) with the new rows.
    ' Adds Change log rows through PTCommon. Returns False when there was no previous import.
    Private Function CompareImports(ByVal oldImp(,) As Object, ByVal outv(,) As Object, ByVal cnt As Integer, ByVal refByKey As Dict,
                                    ByVal tolPos As Double, ByVal tolSize As Double, ByVal counts As Dict) As Boolean
        Dim oldRow As New Dict(), used As New Dict(), r As Integer, s As String, k0 As String, o As Integer
        Dim newOnly As Collection, oldOnly As Collection, found As Boolean, hasRows As Boolean
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
                If Not oldRow.Exists(k0) Then oldRow.Add(k0, r)
            End If
        Next
        If Not hasRows Then Return False

        newOnly = New Collection()
        For r = 0 To cnt - 1
            k0 = CStr(outv(r, 9))
            If oldRow.Exists(k0) Then
                used.Add(k0, 1)
                CompareBox(k0, IIf(r = 0, "PCB", CStr(outv(r, 1))), RefOf(refByKey, k0), oldImp, CInt(oldRow(k0)), outv, r, tolPos, tolSize, counts)
            Else
                newOnly.Add(r)
            End If
        Next
        oldOnly = New Collection()
        For Each it As Object In oldRow.Keys
            If Not used.Exists(it) Then oldOnly.Add(CStr(it))
        Next

        ' same box, different key = renamed
        For Each it As Object In newOnly
            r = CInt(it)
            found = False
            For Each it2 As Object In oldOnly
                o = CInt(oldRow(it2))
                If BoxClose(oldImp, o, outv, r, tolPos, tolSize) Then
                    PT_LogAdd(MACRO_NAME, CStr(outv(r, 9)), CStr(outv(r, 1)), RefOf(refByKey, CStr(it2)), "renamed", "Name",
                              PT_Str(oldImp(o, 2)), CStr(outv(r, 1)))
                    Bump(counts, "renamed")
                    RemoveFrom(oldOnly, CStr(it2))
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                PT_LogAdd(MACRO_NAME, CStr(outv(r, 9)), CStr(outv(r, 1)), "", "added", "Box", "", BoxText(outv(r, 2), outv(r, 3), outv(r, 4), outv(r, 5), outv(r, 6), outv(r, 7)))
                Bump(counts, "added")
            End If
        Next
        For Each it As Object In oldOnly
            o = CInt(oldRow(it))
            PT_LogAdd(MACRO_NAME, CStr(it), PT_Str(oldImp(o, 2)), RefOf(refByKey, CStr(it)), "removed", "Box",
                      BoxText(oldImp(o, 3), oldImp(o, 4), oldImp(o, 5), oldImp(o, 6), oldImp(o, 7), oldImp(o, 8)), "")
            Bump(counts, "removed")
        Next
        Return True
    End Function

    Private Sub CompareBox(ByVal key As String, ByVal nameS As String, ByVal refD As String, ByVal oldImp(,) As Object, ByVal o As Integer,
                           ByVal outv(,) As Object, ByVal r As Integer, ByVal tolPos As Double, ByVal tolSize As Double, ByVal counts As Dict)
        Dim fields As String() = {"X min", "Y min", "Z min", "L", "W", "H"}, a As Double, b As Double
        For k As Integer = 0 To 5
            a = PT_Num(oldImp(o, 3 + k), 0)
            b = PT_Num(outv(r, 2 + k), 0)
            If k < 3 Then
                If Math.Abs(b - a) > tolPos Then
                    PT_LogAdd(MACRO_NAME, key, nameS, refD, "moved", fields(k), a, b)
                    Bump(counts, "moved")
                End If
            Else
                If Math.Abs(b - a) > tolSize Then
                    PT_LogAdd(MACRO_NAME, key, nameS, refD, "resized", fields(k), a, b)
                    Bump(counts, "resized")
                End If
            End If
        Next
    End Sub

    Private Function BoxClose(ByVal oldImp(,) As Object, ByVal o As Integer, ByVal outv(,) As Object, ByVal r As Integer,
                              ByVal tolPos As Double, ByVal tolSize As Double) As Boolean
        Dim tol As Double
        For k As Integer = 0 To 5
            If Not PT_IsNum(oldImp(o, 3 + k)) Then Return False
            If k < 3 Then tol = tolPos Else tol = tolSize
            If Math.Abs(CDbl(oldImp(o, 3 + k)) - CDbl(outv(r, 2 + k))) > tol Then Return False
        Next
        Return True
    End Function

    Private Function BoxText(ByVal x0 As Object, ByVal y0 As Object, ByVal z0 As Object, ByVal lx As Object, ByVal wy As Object, ByVal hz As Object) As String
        Return "at " & PT_Num(x0, 0) & ", " & PT_Num(y0, 0) & ", " & PT_Num(z0, 0) &
               " size " & PT_Num(lx, 0) & " x " & PT_Num(wy, 0) & " x " & PT_Num(hz, 0)
    End Function

    Private Function RefOf(ByVal refByKey As Dict, ByVal key As String) As String
        If refByKey.Exists(key) Then Return CStr(refByKey(key))
        Return ""
    End Function

    Private Sub Bump(ByVal counts As Dict, ByVal what As String)
        If counts.Exists(what) Then counts(what) = CInt(counts(what)) + 1 Else counts.Add(what, 1)
    End Sub

    Private Sub RemoveFrom(ByVal coll As Collection, ByVal what As String)
        For i As Integer = 1 To coll.Count
            If CStr(coll(i)) = what Then
                coll.Remove(i)
                Exit Sub
            End If
        Next
    End Sub

    ' ---------------------------------------------------------------- model access
    Private Function GetModel(ByRef srcName As String) As Object
        Dim act As Object, a As MsgBoxResult
        act = swApp.ActiveDoc
        If act IsNot Nothing Then
            If SwDocType(act) = swDocASSEMBLY Or SwDocType(act) = swDocPART Then
                a = MsgBox("Use the model open in SolidWorks?" & vbCrLf & CStr(act.GetTitle) & vbCrLf & vbCrLf &
                           "No = pick a STEP file instead.", vbYesNoCancel + vbQuestion, MACRO_NAME)
                If a = vbCancel Then Return Nothing
                If a = vbYes Then
                    srcName = CStr(act.GetTitle)
                    Return act
                End If
            End If
        End If
        Dim f As String
        f = SwGetOpenFileName("Pick the PCB STEP file", "", "STEP files (*.step;*.stp)|*.step;*.stp|All files (*.*)|*.*||")
        If f = "" Then Return Nothing
        Dim errs As Integer = 0, md As Object
        md = SwLoadFile4(f, errs)
        If md Is Nothing Then
            MsgBox("SolidWorks could not open " & f & " (error " & errs & ").", vbExclamation, MACRO_NAME)
            Return Nothing
        End If
        srcName = f
        Return md
    End Function

    ' ---------------------------------------------------------------- collecting the bodies
    Private Sub CollectItems(ByVal doc As Object, ByVal perBody As Boolean)
        n = 0
        ReDim nm(MAX_ITEMS - 1) : ReDim bx(MAX_ITEMS - 1, 5) : ReDim rbx(MAX_ITEMS - 1, 5)
        ReDim srcComp(MAX_ITEMS - 1) : ReDim srcBody(MAX_ITEMS - 1)
        ReDim srcMode(MAX_ITEMS - 1) : ReDim srcXf(MAX_ITEMS - 1, 12)
        Dim v As Object, c As Object
        If SwDocType(doc) = swDocASSEMBLY Then
            doc.ResolveAllLightWeightComponents(False)
            Dim comps As Object, ch As Object
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
            For i As Integer = 0 To UBound(comps)
                c = comps(i)
                If Not CBool(c.IsSuppressed) Then
                    If perBody Then
                        If IsLeaf(c) Then AddCompBodies(c)
                    Else
                        If Not CBool(c.IsHidden(True)) Then AddWholeComp(c)
                    End If
                End If
            Next
        Else
            Dim bodies As Object, one(5) As Double, xd(12) As Double
            xd(0) = 1 : xd(4) = 1 : xd(8) = 1 : xd(12) = 1
            bodies = doc.GetBodies2(swSolidBody, True)
            If IsEmpty(bodies) Then Exit Sub
            For i As Integer = 0 To UBound(bodies)
                v = bodies(i).GetBodyBox
                If Not IsEmpty(v) Then
                    For k As Integer = 0 To 5
                        one(k) = v(k) * 1000.0
                    Next
                    AddItem(CStr(bodies(i).Name), one, one, Nothing, bodies(i), 0, xd)
                End If
            Next
        End If
    End Sub

    Private Function IsLeaf(ByVal c As Object) As Boolean
        Dim ch As Object = c.GetChildren
        If IsEmpty(ch) Then Return True
        Return (UBound(ch) < 0)
    End Function

    ' One item per solid body of a component. A component with one body keeps the component name (so keys and inputs
    ' from earlier imports still match); with several bodies the body name is put before the instance suffix.
    Private Sub AddCompBodies(ByVal c As Object)
        Dim bodies As Object = Nothing, info As Object = Nothing, info2 As Object = Nothing, sheetB As Object = Nothing
        Dim xd(12) As Double, spMode As Integer, ab(5) As Double, rb(5) As Double, bb As Object, lbl As String, nb As Integer
        Try
            bodies = c.GetBodies3(swSolidBody, info)
        Catch
        End Try
        Try
            sheetB = c.GetBodies3(swSheetBody, info2)
        Catch
        End Try
        If Not IsEmpty(sheetB) Then nSheetBodies = nSheetBodies + UBound(sheetB) + 1
        If IsEmpty(bodies) Then Exit Sub
        nb = UBound(bodies) + 1
        ReadXf(c, xd)
        spMode = PickMode(c, bodies, xd)
        For k As Integer = 0 To nb - 1
            bb = Nothing
            Try
                bb = bodies(k).GetBodyBox
            Catch
            End Try
            If Not IsEmpty(bb) Then
                For q As Integer = 0 To 5
                    rb(q) = bb(q) * 1000.0
                Next
                BoxXf(bb, xd, spMode, ab, True)
                If nb = 1 Then lbl = CStr(c.Name2) Else lbl = BodyLabel(CStr(c.Name2), CStr(bodies(k).Name), k)
                AddItem(lbl, ab, rb, c, bodies(k), spMode, xd)
            End If
        Next
    End Sub

    ' One item for a whole component (setting Imp_TopOnly = Yes): its box in the assembly.
    Private Sub AddWholeComp(ByVal c As Object)
        Dim v As Object, bodies As Object = Nothing, info As Object = Nothing, xd(12) As Double, spMode As Integer
        Dim ab(5) As Double, rb(5) As Double, bb As Object, haveRaw As Boolean
        v = c.GetBox(False, False)
        If IsEmpty(v) Then Exit Sub
        For q As Integer = 0 To 5
            ab(q) = v(q) * 1000.0
        Next
        ReadXf(c, xd)
        Try
            bodies = c.GetBodies3(swSolidBody, info)
        Catch
        End Try
        For q As Integer = 0 To 2
            rb(q) = 1.0E+30 : rb(3 + q) = -1.0E+30
        Next
        If Not IsEmpty(bodies) Then
            spMode = PickMode(c, bodies, xd)
            For k As Integer = 0 To UBound(bodies)
                bb = Nothing
                Try
                    bb = bodies(k).GetBodyBox
                Catch
                End Try
                If Not IsEmpty(bb) Then
                    haveRaw = True
                    For q As Integer = 0 To 2
                        If bb(q) * 1000.0 < rb(q) Then rb(q) = bb(q) * 1000.0
                        If bb(3 + q) * 1000.0 > rb(3 + q) Then rb(3 + q) = bb(3 + q) * 1000.0
                    Next
                End If
            Next
        End If
        If Not haveRaw Then
            For q As Integer = 0 To 5
                rb(q) = ab(q)
            Next
            spMode = 0
        End If
        AddItem(CStr(c.Name2), ab, rb, c, Nothing, spMode, xd)
    End Sub

    ' "<path>/<part>.STEP-1" + body "Imported12" -> "<path>/<part>.STEP_Imported12-1"
    Private Function BodyLabel(ByVal compName As String, ByVal bodyName As String, ByVal k As Integer) As String
        Dim headS As String, tailS As String
        If bodyName = "" Then bodyName = "body" & (k + 1)
        Dim m As Match = Regex.Match(compName, "^(.*?)(-\d+|<\d+>)$")
        If m.Success Then
            headS = m.Groups(1).Value : tailS = m.Groups(2).Value
        Else
            headS = compName : tailS = ""
        End If
        Return headS & "_" & bodyName & tailS
    End Function

    Private Sub AddItem(ByVal lbl As String, ByVal asmBox() As Double, ByVal rawBox() As Double, ByVal comp As Object, ByVal body As Object,
                        ByVal spMode As Integer, ByVal xd() As Double)
        If n >= MAX_ITEMS Then
            itemsCapped = True
            Exit Sub
        End If
        nm(n) = lbl
        For k As Integer = 0 To 5
            bx(n, k) = asmBox(k) : rbx(n, k) = rawBox(k)
        Next
        For k As Integer = 0 To 12
            srcXf(n, k) = xd(k)
        Next
        srcComp(n) = comp
        srcBody(n) = body
        srcMode(n) = spMode
        n = n + 1
    End Sub

    ' Component transform as 13 numbers: 9 rotation, 3 translation (metres), scale. Identity when unreadable.
    Private Sub ReadXf(ByVal c As Object, ByVal xd() As Double)
        For k As Integer = 0 To 12
            xd(k) = 0
        Next
        xd(0) = 1 : xd(4) = 1 : xd(8) = 1 : xd(12) = 1
        Try
            Dim t As Object = c.Transform2
            If t Is Nothing Then Exit Sub
            Dim a As Object = t.ArrayData
            If IsEmpty(a) Then Exit Sub
            If UBound(a) < 12 Then Exit Sub
            For k As Integer = 0 To 12
                xd(k) = CDbl(a(k))
            Next
        Catch
        End Try
    End Sub

    ' Which way turns a body's own-space points into model space: 0 none, 1 matrix, 2 transposed matrix.
    ' Decided by comparing the transformed body boxes with the component box in the assembly (GetBox).
    Private Function PickMode(ByVal c As Object, ByVal bodies As Object, ByVal xd() As Double) As Integer
        Dim cb As Object = Nothing, bb As Object, ub(5) As Double, e As Double, bestE As Double, bestM As Integer = 0
        Try
            cb = c.GetBox(False, False)
        Catch
        End Try
        If IsEmpty(cb) Then Return 0
        bestE = 1.0E+30
        For md As Integer = 0 To 2
            For q As Integer = 0 To 2
                ub(q) = 1.0E+30 : ub(3 + q) = -1.0E+30
            Next
            For k As Integer = 0 To UBound(bodies)
                bb = Nothing
                Try
                    bb = bodies(k).GetBodyBox
                Catch
                End Try
                If Not IsEmpty(bb) Then BoxXf(bb, xd, md, ub, False)
            Next
            e = 0
            For q As Integer = 0 To 5
                If Math.Abs(ub(q) - cb(q) * 1000.0) > e Then e = Math.Abs(ub(q) - cb(q) * 1000.0)
            Next
            If e < bestE - 0.001 Then
                bestE = e
                bestM = md
            End If
        Next
        Return bestM
    End Function

    ' Box (metres, 6 numbers) -> corners turned by the component transform -> axis-aligned box in mm.
    ' reset = True starts the result from nothing, False adds to it.
    Private Sub BoxXf(ByVal bb As Object, ByVal xd() As Double, ByVal spMode As Integer, ByVal outB() As Double, ByVal reset As Boolean)
        Dim x As Double, y As Double, z As Double, ox As Double, oy As Double, oz As Double
        If reset Then
            For q As Integer = 0 To 2
                outB(q) = 1.0E+30 : outB(3 + q) = -1.0E+30
            Next
        End If
        For a As Integer = 0 To 1
            For b As Integer = 0 To 1
                For c As Integer = 0 To 1
                    x = bb(0 + 3 * a) * 1000.0 : y = bb(1 + 3 * b) * 1000.0 : z = bb(2 + 3 * c) * 1000.0
                    XfPoint(spMode, xd, x, y, z, ox, oy, oz)
                    If ox < outB(0) Then outB(0) = ox
                    If oy < outB(1) Then outB(1) = oy
                    If oz < outB(2) Then outB(2) = oz
                    If ox > outB(3) Then outB(3) = ox
                    If oy > outB(4) Then outB(4) = oy
                    If oz > outB(5) Then outB(5) = oz
                Next
            Next
        Next
    End Sub

    ' Point in mm, own space -> model space (mm)
    Private Sub XfPoint(ByVal spMode As Integer, ByVal xd() As Double, ByVal x As Double, ByVal y As Double, ByVal z As Double,
                        ByRef ox As Double, ByRef oy As Double, ByRef oz As Double)
        Dim sc As Double
        If spMode = 0 Then
            ox = x : oy = y : oz = z
            Exit Sub
        End If
        sc = xd(12)
        If sc = 0 Then sc = 1
        If spMode = 1 Then
            ox = (xd(0) * x + xd(1) * y + xd(2) * z) * sc + xd(9) * 1000.0
            oy = (xd(3) * x + xd(4) * y + xd(5) * z) * sc + xd(10) * 1000.0
            oz = (xd(6) * x + xd(7) * y + xd(8) * z) * sc + xd(11) * 1000.0
        Else
            ox = (xd(0) * x + xd(3) * y + xd(6) * z) * sc + xd(9) * 1000.0
            oy = (xd(1) * x + xd(4) * y + xd(7) * z) * sc + xd(10) * 1000.0
            oz = (xd(2) * x + xd(5) * y + xd(8) * z) * sc + xd(11) * 1000.0
        End If
    End Sub

    Private Function ItemVolume(ByVal i As Integer) As Double
        Dim bodies As Object = Nothing, info As Object = Nothing, mp As Object, vol As Double = 0
        Dim leaves As New Collections.Generic.List(Of Object)
        Try
            If srcBody(i) IsNot Nothing Then
                leaves.Add(Nothing)
            ElseIf srcComp(i) IsNot Nothing Then
                LeafComps(srcComp(i), leaves)
            End If
            For Each lc As Object In leaves
                bodies = Nothing
                If lc Is Nothing Then
                    bodies = Arr(srcBody(i))
                Else
                    Try
                        bodies = lc.GetBodies3(swSolidBody, info)
                    Catch
                    End Try
                End If
                If Not IsEmpty(bodies) Then
                    For k As Integer = 0 To UBound(bodies)
                        mp = Nothing
                        Try
                            mp = bodies(k).GetMassProperties(1.0)
                        Catch
                        End Try
                        If Not IsEmpty(mp) Then vol = vol + CDbl(mp(3))
                    Next
                End If
            Next
        Catch
        End Try
        Return vol * 1000000000.0                           ' m3 -> mm3
    End Function

    ' The part components under a component (the component itself when it is a part). A component that is a
    ' sub-assembly has no bodies of its own, so its shape comes from these. Suppressed and hidden ones are left out.
    Private Sub LeafComps(ByVal c As Object, ByVal leaves As Collections.Generic.List(Of Object))
        If CBool(c.IsSuppressed) Then Exit Sub
        If IsLeaf(c) Then
            leaves.Add(c)
            Exit Sub
        End If
        Dim ch As Object = c.GetChildren
        For k As Integer = 0 To UBound(ch)
            If Not CBool(ch(k).IsHidden(True)) Then LeafComps(ch(k), leaves)
        Next
    End Sub

    ' ---------------------------------------------------------------- surface points of a body
    ' Points (mm, model space) of an item: 3 numbers per point, returned count is the number of NUMBERS used.
    ' A whole component is read part by part, each with its own transform (it can be a sub-assembly).
    Private Function ItemPoints(ByVal i As Integer, ByRef pts() As Double) As Integer
        Dim np As Integer = 0, bodies As Object = Nothing, info As Object = Nothing, rb(5) As Double, ab(5) As Double
        Dim xd(12) As Double, md As Integer, start As Integer, npPrev As Integer, cb As Object, bb As Object
        ReDim pts(3 * 4096 - 1)
        Try
            If srcBody(i) IsNot Nothing Then
                np = AppendBodyPts(srcBody(i), pts, np)
                If np < 9 Then Return 0
                For q As Integer = 0 To 5
                    rb(q) = rbx(i, q) : ab(q) = bx(i, q)
                Next
                For q As Integer = 0 To 12
                    xd(q) = srcXf(i, q)
                Next
                FixPts(pts, 0, np, rb, ab, srcMode(i), xd)
            ElseIf srcComp(i) IsNot Nothing Then
                Dim leaves As New Collections.Generic.List(Of Object)
                LeafComps(srcComp(i), leaves)
                For Each lc As Object In leaves
                    bodies = Nothing
                    Try
                        bodies = lc.GetBodies3(swSolidBody, info)
                    Catch
                    End Try
                    If Not IsEmpty(bodies) Then
                        cb = Nothing
                        Try
                            cb = lc.GetBox(False, False)
                        Catch
                        End Try
                        ReadXf(lc, xd)
                        md = PickMode(lc, bodies, xd)
                        For q As Integer = 0 To 2
                            rb(q) = 1.0E+30 : rb(3 + q) = -1.0E+30
                        Next
                        start = np
                        For k As Integer = 0 To UBound(bodies)
                            bb = Nothing
                            Try
                                bb = bodies(k).GetBodyBox
                            Catch
                            End Try
                            If Not IsEmpty(bb) Then
                                For q As Integer = 0 To 2
                                    If bb(q) * 1000.0 < rb(q) Then rb(q) = bb(q) * 1000.0
                                    If bb(3 + q) * 1000.0 > rb(3 + q) Then rb(3 + q) = bb(3 + q) * 1000.0
                                Next
                            End If
                            npPrev = np
                            np = AppendBodyPts(bodies(k), pts, np)
                            If np = npPrev Then Return 0           ' a body without points would leave a hole in the shape
                        Next
                        If IsEmpty(cb) Then
                            Dim rbm(5) As Double
                            For q As Integer = 0 To 5
                                rbm(q) = rb(q) / 1000.0
                            Next
                            BoxXf(rbm, xd, md, ab, True)
                        Else
                            For q As Integer = 0 To 5
                                ab(q) = cb(q) * 1000.0
                            Next
                        End If
                        FixPts(pts, start, np, rb, ab, md, xd)
                    End If
                Next
            End If
        Catch
            Return 0
        End Try
        If np < 9 Then Return 0
        Return np
    End Function

    ' Puts the points pts(start .. np-1) of one part into model space (mm). rawBox: the part's body box in its own
    ' space (mm), asmBox: its box in the assembly (mm), md and xd: how to turn own-space points (see PickMode, ReadXf).
    Private Sub FixPts(ByVal pts() As Double, ByVal start As Integer, ByVal np As Integer, ByVal rawBox() As Double,
                       ByVal asmBox() As Double, ByVal md As Integer, ByVal xd() As Double)
        Dim rb(5) As Double, dRaw As Double, dAsm As Double, sc As Double, ox As Double, oy As Double, oz As Double
        For q As Integer = 0 To 2
            rb(q) = 1.0E+30 : rb(3 + q) = -1.0E+30
        Next
        For k As Integer = start To np - 3 Step 3
            For q As Integer = 0 To 2
                If pts(k + q) < rb(q) Then rb(q) = pts(k + q)
                If pts(k + q) > rb(3 + q) Then rb(3 + q) = pts(k + q)
            Next
        Next
        ' unit guard: the points must have about the size of the body box (mm). A factor of 1000 means the API returned
        ' millimetres or metres where metres were expected: put it right instead of building the wrong shape.
        dRaw = (rb(3) - rb(0)) + (rb(4) - rb(1)) + (rb(5) - rb(2))
        dAsm = (rawBox(3) - rawBox(0)) + (rawBox(4) - rawBox(1)) + (rawBox(5) - rawBox(2))
        sc = 1.0
        If dRaw > 0 And dAsm > 0 Then
            If dRaw / dAsm > 500 And dRaw / dAsm < 2000 Then sc = 0.001
            If dRaw / dAsm < 0.002 And dRaw / dAsm > 0.0005 Then sc = 1000.0
        End If
        If sc <> 1.0 Then
            For k As Integer = start To np - 1
                pts(k) = pts(k) * sc
            Next
            For q As Integer = 0 To 5
                rb(q) = rb(q) * sc
            Next
        End If
        If md = 0 Then Exit Sub
        ' are the points in the part's own space (turn them) or already in the assembly (leave them)?
        dRaw = 0 : dAsm = 0
        For q As Integer = 0 To 5
            If Math.Abs(rb(q) - rawBox(q)) > dRaw Then dRaw = Math.Abs(rb(q) - rawBox(q))
            If Math.Abs(rb(q) - asmBox(q)) > dAsm Then dAsm = Math.Abs(rb(q) - asmBox(q))
        Next
        If dAsm + 0.0001 < dRaw Then Exit Sub
        For k As Integer = start To np - 3 Step 3
            XfPoint(md, xd, pts(k), pts(k + 1), pts(k + 2), ox, oy, oz)
            pts(k) = ox : pts(k + 1) = oy : pts(k + 2) = oz
        Next
    End Sub

    Private Sub GrowPts(ByRef pts() As Double, ByVal need As Integer)
        Dim cap As Integer = UBound(pts) + 1
        If need <= cap Then Exit Sub
        Do While cap < need
            cap = cap * 2
        Loop
        ReDim Preserve pts(cap - 1)
    End Sub

    ' Appends the surface points of a body (mm, the body's own space). Two ways: face triangles (when EVERY face has them),
    ' else the body tessellation object. Sparse sources such as the vertex list are not used: they miss curved surfaces.
    ' Returns the new count of numbers (unchanged when no complete point set could be read).
    ' Typed SolidWorks calls here (early binding): this loop makes thousands of calls on a big board.
    Private Function AppendBodyPts(ByVal body As Object, ByRef pts() As Double, ByVal np As Integer) As Integer
        Dim faces As Object = Nothing, t As Object, m As Integer, start As Integer, nEmpty As Integer = 0
        Dim tess As SW.Tessellation, nv As Integer, p As Double()
        Dim b2 As SW.Body2 = TryCast(body, SW.Body2)
        If b2 Is Nothing Then Return np
        start = np
        Try
            faces = b2.GetFaces()
        Catch
        End Try
        If Not IsEmpty(faces) Then
            Dim fa As Object() = CType(faces, Object())
            For f As Integer = 0 To fa.Length - 1
                t = Nothing
                Try
                    t = CType(fa(f), SW.Face2).GetTessTriangles(True)
                Catch
                End Try
                If IsEmpty(t) Then
                    nEmpty = nEmpty + 1
                Else
                    Dim ta As Array = CType(t, Array)
                    m = ta.Length
                    GrowPts(pts, np + m)
                    Dim tf As Single() = TryCast(t, Single())
                    Dim td As Double() = TryCast(t, Double())
                    For q As Integer = 0 To m - 3 Step 3
                        If tf IsNot Nothing Then
                            pts(np) = CDbl(tf(q)) * 1000.0 : pts(np + 1) = CDbl(tf(q + 1)) * 1000.0 : pts(np + 2) = CDbl(tf(q + 2)) * 1000.0
                        ElseIf td IsNot Nothing Then
                            pts(np) = td(q) * 1000.0 : pts(np + 1) = td(q + 1) * 1000.0 : pts(np + 2) = td(q + 2) * 1000.0
                        Else
                            pts(np) = CDbl(ta.GetValue(q)) * 1000.0 : pts(np + 1) = CDbl(ta.GetValue(q + 1)) * 1000.0 : pts(np + 2) = CDbl(ta.GetValue(q + 2)) * 1000.0
                        End If
                        np = np + 3
                    Next
                End If
            Next
        End If
        lastPtsRoute = "faces " & If(faces Is Nothing, 0, UBound(faces) + 1) & ", without triangles " & nEmpty
        If nEmpty > 0 Then np = start                       ' some faces had no triangles: the set is incomplete, use the tessellation object
        If np = start And Not IsEmpty(faces) Then
            lastPtsRoute &= ", tessellation object"
            tess = Nothing
            Try
                tess = CType(b2.GetTessellation(faces), SW.Tessellation)
            Catch
            End Try
            If tess IsNot Nothing Then
                Try
                    tess.NeedFaceFacetMap = False
                    tess.NeedVertexParams = False
                    tess.NeedVertexNormal = False
                    If tess.Tessellate() Then
                        nv = tess.GetVertexCount()
                        GrowPts(pts, np + 3 * nv)
                        For q As Integer = 0 To nv - 1
                            p = Nothing
                            Try
                                p = CType(tess.GetVertexPoint(q), Double())
                            Catch
                            End Try
                            If p IsNot Nothing Then
                                pts(np) = CDbl(p(0)) * 1000.0 : pts(np + 1) = CDbl(p(1)) * 1000.0 : pts(np + 2) = CDbl(p(2)) * 1000.0
                                np = np + 3
                            End If
                        Next
                    End If
                Catch
                End Try
            End If
        End If
        Return np
    End Function

    ' ---------------------------------------------------------------- shape in the board frame
    ' sh(0..2) min and sh(3..5) max of the envelope (board X, Y, Z), sh(6) rotation (deg, -45..45, 0 = square to the frame),
    ' sh(7) own L, sh(8) own W (the turned rectangle), sh(9) sh(10) centre X, Y. False when no points could be read.
    Private tPoints As Double          ' seconds spent reading surface points (log only)
    Private lastPtsRoute As String = "" ' how the last body's points were read (log only)

    Private Function ItemShape(ByVal i As Integer, ByVal fr() As Double, ByVal sh() As Double) As Boolean
        Dim pts() As Double = Nothing, np As Integer
        Dim t0 As Long = Diagnostics.Stopwatch.GetTimestamp()
        np = ItemPoints(i, pts)
        Dim dt As Double = (Diagnostics.Stopwatch.GetTimestamp() - t0) / CDbl(Diagnostics.Stopwatch.Frequency)
        tPoints += dt
        If dt > 5 Then Log.Write("ImportStep: " & nm(i) & ": " & np \ 3 & " points in " & dt.ToString("0.0") & " s (" & lastPtsRoute & ")")
        If np < 9 Then Return False
        Return Footprint(pts, np, fr, sh)
    End Function

    ' Same result layout from the plain box (corners projected on the frame), no rotation.
    Private Sub BoxShape(ByVal i As Integer, ByVal fr() As Double, ByVal sh() As Double)
        Dim p(2) As Double, q(2) As Double
        For k As Integer = 0 To 2
            sh(k) = 1.0E+30 : sh(3 + k) = -1.0E+30
        Next
        For a As Integer = 0 To 1
            For b As Integer = 0 To 1
                For c As Integer = 0 To 1
                    p(0) = bx(i, 0 + 3 * a) : p(1) = bx(i, 1 + 3 * b) : p(2) = bx(i, 2 + 3 * c)
                    For k As Integer = 0 To 2
                        q(k) = p(0) * fr(3 * k) + p(1) * fr(3 * k + 1) + p(2) * fr(3 * k + 2)
                        If q(k) < sh(k) Then sh(k) = q(k)
                        If q(k) > sh(3 + k) Then sh(3 + k) = q(k)
                    Next
                Next
            Next
        Next
        sh(6) = 0 : sh(7) = sh(3) - sh(0) : sh(8) = sh(4) - sh(1)
        sh(9) = (sh(0) + sh(3)) / 2 : sh(10) = (sh(1) + sh(4)) / 2
    End Sub

    ' Smallest turned rectangle round the points, seen from +Z of the frame fr (rows: X, Y, Z unit vectors).
    Private Function Footprint(ByVal pts() As Double, ByVal np As Integer, ByVal fr() As Double, ByVal sh() As Double) As Boolean
        Dim u As Double, v As Double, w As Double, bi As Integer
        Dim umin As Double, umax As Double, vmin As Double, vmax As Double, wmin As Double, wmax As Double
        umin = 1.0E+30 : vmin = 1.0E+30 : wmin = 1.0E+30 : umax = -1.0E+30 : vmax = -1.0E+30 : wmax = -1.0E+30
        For q As Integer = 0 To np - 3 Step 3
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
        sh(0) = umin : sh(1) = vmin : sh(2) = wmin : sh(3) = umax : sh(4) = vmax : sh(5) = wmax
        sh(6) = 0 : sh(7) = umax - umin : sh(8) = vmax - vmin
        sh(9) = (umin + umax) / 2 : sh(10) = (vmin + vmax) / 2
        Dim spanU As Double, spanV As Double
        spanU = umax - umin : spanV = vmax - vmin
        If spanU < 0.001 Or spanV < 0.001 Then Return True          ' flat sliver: nothing to turn

        ' keep, per bin, the lowest and highest point (a few hundred points are enough for the hull)
        Dim aMinV(NBINS - 1) As Double, aMinU(NBINS - 1) As Double, aMaxV(NBINS - 1) As Double, aMaxU(NBINS - 1) As Double
        Dim bMinU(NBINS - 1) As Double, bMinV(NBINS - 1) As Double, bMaxU(NBINS - 1) As Double, bMaxV(NBINS - 1) As Double
        For bi = 0 To NBINS - 1
            aMinV(bi) = 1.0E+30 : aMaxV(bi) = -1.0E+30 : bMinU(bi) = 1.0E+30 : bMaxU(bi) = -1.0E+30
        Next
        For q As Integer = 0 To np - 3 Step 3
            u = pts(q) * fr(0) + pts(q + 1) * fr(1) + pts(q + 2) * fr(2)
            v = pts(q) * fr(3) + pts(q + 1) * fr(4) + pts(q + 2) * fr(5)
            bi = CInt((u - umin) / spanU * (NBINS - 1))
            If bi < 0 Then bi = 0
            If bi > NBINS - 1 Then bi = NBINS - 1
            If v < aMinV(bi) Then
                aMinV(bi) = v : aMinU(bi) = u
            End If
            If v > aMaxV(bi) Then
                aMaxV(bi) = v : aMaxU(bi) = u
            End If
            bi = CInt((v - vmin) / spanV * (NBINS - 1))
            If bi < 0 Then bi = 0
            If bi > NBINS - 1 Then bi = NBINS - 1
            If u < bMinU(bi) Then
                bMinU(bi) = u : bMinV(bi) = v
            End If
            If u > bMaxU(bi) Then
                bMaxU(bi) = u : bMaxV(bi) = v
            End If
        Next
        Dim ru(4 * NBINS - 1) As Double, rv(4 * NBINS - 1) As Double, nr As Integer = 0
        For bi = 0 To NBINS - 1
            If aMinV(bi) < 1.0E+29 Then
                ru(nr) = aMinU(bi) : rv(nr) = aMinV(bi) : nr = nr + 1
            End If
            If aMaxV(bi) > -1.0E+29 Then
                ru(nr) = aMaxU(bi) : rv(nr) = aMaxV(bi) : nr = nr + 1
            End If
            If bMinU(bi) < 1.0E+29 Then
                ru(nr) = bMinU(bi) : rv(nr) = bMinV(bi) : nr = nr + 1
            End If
            If bMaxU(bi) > -1.0E+29 Then
                ru(nr) = bMaxU(bi) : rv(nr) = bMaxV(bi) : nr = nr + 1
            End If
        Next
        If nr < 4 Then Return True

        ' angle with the smallest rectangle: coarse steps, then fine steps round the best one
        Dim ang As Double, ar As Double, bestAr As Double, bestAng As Double
        Dim u0 As Double, u1 As Double, v0 As Double, v1 As Double
        bestAr = 1.0E+300
        For ang = 0 To 89.5 Step 0.5
            ExtentsAt(ru, rv, nr, ang, u0, u1, v0, v1)
            ar = (u1 - u0) * (v1 - v0)
            If ar < bestAr - 0.000000001 Then
                bestAr = ar : bestAng = ang
            End If
        Next
        Dim ang0 As Double = bestAng
        For ang = ang0 - 0.5 To ang0 + 0.5 Step 0.02
            ExtentsAt(ru, rv, nr, ang, u0, u1, v0, v1)
            ar = (u1 - u0) * (v1 - v0)
            If ar < bestAr - 0.000000001 Then
                bestAr = ar : bestAng = ang
            End If
        Next
        ExtentsAt(ru, rv, nr, bestAng, u0, u1, v0, v1)

        Dim lr As Double, wr As Double, cu As Double, cv As Double, cx As Double, cy As Double, phi As Double, tmp As Double
        lr = u1 - u0 : wr = v1 - v0
        cu = (u0 + u1) / 2 : cv = (v0 + v1) / 2
        phi = bestAng
        Do While phi > 45
            phi = phi - 90 : tmp = lr : lr = wr : wr = tmp
        Loop
        Do While phi <= -45
            phi = phi + 90 : tmp = lr : lr = wr : wr = tmp
        Loop
        cx = cu * Math.Cos(bestAng * PI / 180.0) - cv * Math.Sin(bestAng * PI / 180.0)
        cy = cu * Math.Sin(bestAng * PI / 180.0) + cv * Math.Cos(bestAng * PI / 180.0)

        ' worth it only when the turned rectangle is clearly smaller than the envelope box
        Dim envL As Double, envW As Double
        envL = lr * Math.Abs(Math.Cos(phi * PI / 180.0)) + wr * Math.Abs(Math.Sin(phi * PI / 180.0))
        envW = lr * Math.Abs(Math.Sin(phi * PI / 180.0)) + wr * Math.Abs(Math.Cos(phi * PI / 180.0))
        If Math.Abs(phi) < 0.3 Then Return True
        If (spanU * spanV) < 0.000001 Then Return True
        If 1.0 - (lr * wr) / (spanU * spanV) < 0.03 Then Return True
        If envL > spanU * 1.5 + 0.1 Or envW > spanV * 1.5 + 0.1 Then Return True
        sh(0) = cx - envL / 2 : sh(3) = cx + envL / 2
        sh(1) = cy - envW / 2 : sh(4) = cy + envW / 2
        sh(6) = phi : sh(7) = lr : sh(8) = wr
        sh(9) = cx : sh(10) = cy
        Return True
    End Function

    Private Sub ExtentsAt(ByVal ru() As Double, ByVal rv() As Double, ByVal nr As Integer, ByVal ang As Double,
                          ByRef u0 As Double, ByRef u1 As Double, ByRef v0 As Double, ByRef v1 As Double)
        Dim ca As Double, sa As Double, a As Double, b As Double
        ca = Math.Cos(ang * PI / 180.0) : sa = Math.Sin(ang * PI / 180.0)
        u0 = 1.0E+30 : v0 = 1.0E+30 : u1 = -1.0E+30 : v1 = -1.0E+30
        For q As Integer = 0 To nr - 1
            a = ru(q) * ca + rv(q) * sa
            b = -ru(q) * sa + rv(q) * ca
            If a < u0 Then u0 = a
            If a > u1 Then u1 = a
            If b < v0 Then v0 = b
            If b > v1 Then v1 = b
        Next
    End Sub

    ' Turns the frame (board X and Y) by deg about board Z: x' = cos x + sin y, y' = -sin x + cos y
    Private Sub RotateFrame(ByVal fr() As Double, ByVal deg As Double)
        Dim ca As Double, sa As Double, x As Double, y As Double
        ca = Math.Cos(deg * PI / 180.0) : sa = Math.Sin(deg * PI / 180.0)
        If Math.Abs(ca) < 0.000000000001 Then ca = 0
        If Math.Abs(sa) < 0.000000000001 Then sa = 0
        For k As Integer = 0 To 2
            x = fr(k) : y = fr(3 + k)
            fr(k) = ca * x + sa * y
            fr(3 + k) = -sa * x + ca * y
        Next
    End Sub

    ' ---------------------------------------------------------------- geometry helpers
    Private Function BoxVol(ByVal i As Integer) As Double
        Return (bx(i, 3) - bx(i, 0)) * (bx(i, 4) - bx(i, 1)) * (bx(i, 5) - bx(i, 2))
    End Function

    Private Function Flatness(ByVal i As Integer) As Double
        Dim a As Double, b As Double, c As Double, t As Double
        a = bx(i, 3) - bx(i, 0) : b = bx(i, 4) - bx(i, 1) : c = bx(i, 5) - bx(i, 2)
        If a > b Then
            t = a : a = b : b = t
        End If
        If b > c Then
            t = b : b = c : c = t
        End If
        If a > b Then
            t = a : a = b : b = t
        End If
        If a < 0.000001 Then a = 0.000001
        Return b * c / a
    End Function

    ' ---------------------------------------------------------------- matching and text
    Private Sub PutInputs(ByVal ws As Object, ByVal g As Integer, ByVal gc As Integer, ByVal rec As Object())
        ws.Cells(g, gc + G_OVNAME).Value = rec(1)
        ws.Cells(g, gc + G_REF).Value = rec(2)
        ws.Cells(g, gc + G_INC).Value = rec(3)
        For k As Integer = 0 To 5
            ws.Cells(g, gc + G_OVR + k).Value = rec(4 + k)
        Next
    End Sub

    Private Function SameBox(ByVal rec As Object(), ByVal outv(,) As Object, ByVal r As Integer) As Boolean
        Dim a As Object
        For k As Integer = 0 To 5
            a = rec(10 + k)
            If IsEmpty(a) Or IsError(a) Then Return False
            If Not IsNumeric(a) Then Return False
            If Math.Abs(CDbl(a) - CDbl(outv(r, 2 + k))) > 0.05 Then Return False
        Next
        Return True
    End Function

    Private Function HasInput(ByVal rec As Object()) As Boolean
        For k As Integer = 1 To 9
            If k <> 3 Then                       ' Include alone is not worth reporting
                If PT_Str(rec(k)) <> "" Then Return True
            End If
        Next
        Return False
    End Function

    ' Fills empty RefDes cells (CAD geometry column F) for bodies whose name key equals the key of a register RefDes.
    ' A RefDes already used on CAD geometry (typed, restored or manual row) is not given out again.
    ' Returns the number of cells filled.
    Private Function AutoLinkRefDes(ByVal wb As Object, ByVal wsG As Object, ByVal g0 As Integer, ByVal gc As Integer,
                                    ByVal outv(,) As Object, ByVal cnt As Integer, ByVal isNewRow() As Boolean, ByRef nOn As Integer) As Integer
        Dim rReg As Object, rGeoT As Object, regV(,) As Object, geoV(,) As Object
        Dim regByKey As New Dict(), usedRef As New Dict(), rk As String, bk As String, refD As String, nFilled As Integer = 0
        rReg = PT_Rng(wb, "Reg_RefDes")
        rGeoT = PT_Rng(wb, "Geo_Table")
        If rReg Is Nothing Or rGeoT Is Nothing Then Return 0
        regV = ToTable(rReg.Value)
        For i As Integer = 1 To UBound(regV, 1)
            refD = PT_Str(regV(i, 1))
            If refD <> "" Then
                rk = NormKey(refD)
                If rk <> "" And Not regByKey.Exists(rk) Then regByKey.Add(rk, refD)
            End If
        Next
        If regByKey.Count = 0 Then Return 0
        geoV = ToTable(rGeoT.Columns(1 + G_REF).Value)
        For i As Integer = 1 To UBound(geoV, 1)
            refD = UCase(PT_Str(geoV(i, 1)))
            If refD <> "" And Not usedRef.Exists(refD) Then usedRef.Add(refD, 1)
        Next
        For i As Integer = 1 To cnt - 1
            If PT_Str(wsG.Cells(g0 + i - 1, gc + G_REF).Value) = "" Then
                bk = NormKey(CStr(outv(i, 1)))
                If regByKey.Exists(bk) Then
                    refD = CStr(regByKey(bk))
                    If Not usedRef.Exists(UCase(refD)) Then
                        wsG.Cells(g0 + i - 1, gc + G_REF).Value = refD
                        usedRef.Add(UCase(refD), 1)
                        nFilled = nFilled + 1
                        If UCase(PT_Str(wsG.Cells(g0 + i - 1, gc + G_INC).Value)) <> "YES" Then
                            wsG.Cells(g0 + i - 1, gc + G_INC).Value = "Yes"
                            nOn = nOn + 1
                        End If
                    End If
                End If
            End If
        Next
        Return nFilled
    End Function

    Private Function JoinMsgs() As String
        Dim s As String = ""
        For t As Integer = 1 To msgs.Count
            If t > 8 Then
                s = s & "(more on the STEP import tab)"
                Exit For
            End If
            s = s & "- " & CStr(msgs(t)) & vbCrLf
        Next
        Return s
    End Function

    ' ---------------------------------------------------------------- workbook preparation
    ' First run on a workbook: adds the rotation columns and makes the import level "every solid body".
    ' Returns "" when nothing was needed, a message when done, "ERR:<reason>" when it failed.
    Private Function UpgradeWorkbook(ByVal wb As Object, ByVal xl As Object) As String
        Dim rImp As Object, rGeo As Object, wsI As Object, wsG As Object, rTop As Object
        Dim r0 As Integer, c0 As Integer, cRot As Integer, cG As Integer, hdrI As Integer, hdrG As Integer, cap As Integer
        Dim lastImp As Integer, lastGeo As Integer, stepName As String, f As String, colL As String, calcWas As Object = Nothing
        Dim msgOut As String
        If PT_Rng(wb, "Imp_Rot") IsNot Nothing Then Return ""
        rImp = PT_Rng(wb, "Imp_Table")
        rGeo = PT_Rng(wb, "Geo_Table")
        wsI = rImp.Worksheet
        wsG = rGeo.Worksheet
        calcWas = xl.Calculation
        On Error GoTo fail
        xl.Calculation = -4135                           ' manual
        xl.ScreenUpdating = False
        stepName = CStr(wsI.Name)

        ' ---- STEP import: Rot (deg), Own L, Own W right of the settings column
        r0 = CInt(rImp.Row) : c0 = CInt(rImp.Column) : hdrI = r0 - 1
        cRot = c0 + 12                                   ' column N when the table starts in B
        lastImp = r0 + CInt(rImp.Rows.Count) - 1
        wsI.Cells(hdrI, cRot).Value = "Rot (deg)"
        wsI.Cells(hdrI, cRot + 1).Value = "Own L"
        wsI.Cells(hdrI, cRot + 2).Value = "Own W"
        wsI.Cells(hdrI, c0 + 8).Copy()
        wsI.Range(wsI.Cells(hdrI, cRot), wsI.Cells(hdrI, cRot + 2)).PasteSpecial(-4122)
        wsI.Cells(r0, c0 + 8).Copy()
        wsI.Range(wsI.Cells(r0, cRot), wsI.Cells(lastImp, cRot + 2)).PasteSpecial(-4122)
        xl.CutCopyMode = False
        wsI.Range(wsI.Cells(1, cRot), wsI.Cells(1, cRot + 2)).EntireColumn.ColumnWidth = 10
        AddNote(wsI.Cells(hdrI, cRot), "Rotation of the body in the board plane, degrees, counter-clockwise seen from the top (Z). " &
            "0 = square to the board edges. Written by ImportStep from the body surface. Xmin, Ymin, L, W stay the axis-aligned envelope of the turned rectangle.")
        AddNote(wsI.Cells(hdrI, cRot + 1), "Length of the turned rectangle along its own X direction (mm). Equals L when Rot is 0.")
        AddNote(wsI.Cells(hdrI, cRot + 2), "Width of the turned rectangle along its own Y direction (mm). Equals W when Rot is 0.")
        wb.Names.Add("Imp_Rot", "='" & stepName & "'!" & CStr(wsI.Range(wsI.Cells(r0, cRot), wsI.Cells(lastImp, cRot + 2)).Address))

        ' ---- CAD geometry: AI:AK mirror the three columns
        cG = CInt(rGeo.Column) + G_ROT
        hdrG = CInt(rGeo.Row) - 1
        cap = BodyCapacity(rGeo)
        lastGeo = CInt(rGeo.Row) + CInt(rGeo.Rows.Count) - 1
        wsG.Cells(hdrG, cG).Value = "Rot (deg)"
        wsG.Cells(hdrG, cG + 1).Value = "Own L"
        wsG.Cells(hdrG, cG + 2).Value = "Own W"
        wsG.Cells(hdrG, rGeo.Column + 12).Copy()
        wsG.Range(wsG.Cells(hdrG, cG), wsG.Cells(hdrG, cG + 2)).PasteSpecial(-4122)
        wsG.Cells(rGeo.Row, rGeo.Column + G_IMP).Copy()
        wsG.Range(wsG.Cells(rGeo.Row, cG), wsG.Cells(rGeo.Row + cap - 1, cG + 2)).PasteSpecial(-4122)
        xl.CutCopyMode = False
        wsG.Range(wsG.Cells(1, cG), wsG.Cells(1, cG + 2)).EntireColumn.ColumnWidth = 10
        For k As Integer = 0 To 2
            colL = Split(CStr(wsI.Cells(1, cRot + k).Address(True, False)), "$")(0)
            f = "=IF(ISNUMBER('" & stepName & "'!$" & colL & (r0 + 1) & "),'" & stepName & "'!$" & colL & (r0 + 1) & ","""")"
            wsG.Range(wsG.Cells(rGeo.Row, cG + k), wsG.Cells(rGeo.Row + cap - 1, cG + k)).Formula = f
        Next
        AddNote(wsG.Cells(hdrG, cG), "Imported rotation of the body in the board plane (degrees, from STEP import). BuildBoard builds a turned rectangle when this is not 0 " &
            "and L and W (Used) still equal the envelope of the turned rectangle. If you override L or W, the part is built as a plain box.")
        AddNote(wsG.Cells(hdrG, cG + 1), "Imported length of the turned rectangle along its own X direction (mm).")
        AddNote(wsG.Cells(hdrG, cG + 2), "Imported width of the turned rectangle along its own Y direction (mm).")
        wb.Names("Geo_Table").RefersTo = "='" & CStr(wsG.Name) & "'!" & CStr(wsG.Range(wsG.Cells(rGeo.Row, rGeo.Column), wsG.Cells(lastGeo, cG + 2)).Address)

        ' ---- default import level: every solid body (was: one row per component)
        rTop = PT_Rng(wb, "Imp_TopOnly")
        msgOut = "Workbook prepared once: Rot columns added (STEP import N:P, CAD geometry AI:AK)."
        If rTop IsNot Nothing Then
            If UCase(PT_Str(rTop.Cells(1, 1).Value)) = "YES" Then
                rTop.Cells(1, 1).Value = "No"
                msgOut = msgOut & " Setting 'one row per top-level component' changed to No, so every solid body is imported."
            End If
            rTop.Cells(1, 1).Offset(0, -1).Value = "One row per top-level component? (No = every solid body)"
            AddNote(rTop.Cells(1, 1), "No (default) = every solid body gets its own row, so small parts are not lost inside a bigger component. " &
                "Yes = one row per top-level component (bodies of a component are merged into one box). Read by ImportStep.")
        End If
        xl.Calculation = calcWas
        xl.ScreenUpdating = True
        Return msgOut
fail:
        Dim errText As String = Err.Description
        On Error Resume Next
        xl.CutCopyMode = False
        xl.Calculation = calcWas
        xl.ScreenUpdating = True
        Return "ERR:" & errText
    End Function

    Private Sub AddNote(ByVal cel As Object, ByVal txt As String)
        Try
            cel.Comment.Delete()
        Catch
        End Try
        Try
            cel.AddComment(txt)
        Catch
        End Try
    End Sub

    ' Number of body rows on CAD geometry: rows from the top of Geo_Table whose # cell is a number
    ' (manual rows below them have text there).
    Private Function BodyCapacity(ByVal rGeo As Object) As Integer
        Dim cap As Integer = 0
        For r As Integer = 1 To CInt(rGeo.Rows.Count)
            If Not PT_IsNum(rGeo.Cells(r, 1).Value) Then Exit For
            cap = r
        Next
        Return cap
    End Function

    ' Makes room for need bodies. Inserts rows into CAD geometry before its manual rows (Excel then moves every formula,
    ' range and name that points there), copies the last body row into them, and lengthens the STEP import tables.
    ' Returns the number of body rows added (0 = nothing to do), or -1 on failure (PT_LastErr).
    Private Function EnsureCapacity(ByVal wb As Object, ByVal xl As Object, ByVal need As Integer) As Integer
        Dim rGeo As Object, rImp As Object, rPrev As Object, wsG As Object, cur As Integer, newCap As Integer, nAdd As Integer
        Dim first As Integer, c0 As Integer, calcWas As Object = Nothing, target As Integer, lastA As Integer, lastB As Integer, cnt As Object
        rGeo = PT_Rng(wb, "Geo_Table")
        rImp = PT_Rng(wb, "Imp_Table")
        rPrev = PT_Rng(wb, "ImpPrev_Table")
        cur = BodyCapacity(rGeo)
        newCap = cur
        If need > cur Then newCap = (need \ 50 + 1) * 50
        target = CInt(rImp.Row) + newCap                                   ' PCB row + newCap body rows
        lastA = CInt(rImp.Row) + CInt(rImp.Rows.Count) - 1
        lastB = CInt(rPrev.Row) + CInt(rPrev.Rows.Count) - 1
        If newCap = cur And lastA >= target And lastB >= target And lastA = lastB Then Return 0
        If lastA > target Then target = lastA
        If lastB > target Then target = lastB
        calcWas = xl.Calculation
        PT_LastErr = ""
        On Error GoTo fail
        xl.Calculation = -4135
        xl.ScreenUpdating = False
        nAdd = newCap - cur
        If nAdd > 0 Then
            wsG = rGeo.Worksheet
            first = CInt(rGeo.Row) + cur
            c0 = CInt(rGeo.Column)
            wsG.Rows(first & ":" & (first + nAdd - 1)).Insert(-4121)
            wsG.Rows(first - 1).Copy()
            wsG.Rows(first & ":" & (first + nAdd - 1)).PasteSpecial(-4104)
            xl.CutCopyMode = False
            wsG.Cells(first, c0 + G_OVNAME).Resize(nAdd, 1).ClearContents()          ' inputs must not be copied
            wsG.Cells(first, c0 + G_REF).Resize(nAdd, 2).ClearContents()
            wsG.Cells(first, c0 + G_OVR).Resize(nAdd, 6).ClearContents()
            For k As Integer = 0 To nAdd - 1
                wsG.Cells(first + k, c0).Value = cur + k + 1
            Next
            ' The CHECKS count of imported bodies ends exactly at the old last body row, so Excel does not stretch it.
            cnt = wsG.Range("H1:H25").Find("Bodies imported from STEP", , -4163, 1)
            If cnt IsNot Nothing Then cnt.Offset(0, 5).Formula = "=COUNTIF(INDEX(Geo_Table,0,2),""?*"")"
        End If
        ResizeName(wb, xl, "Imp_Table", target)
        ResizeName(wb, xl, "ImpPrev_Table", target)
        ResizeName(wb, xl, "Imp_Rot", target)
        xl.Calculation = calcWas
        xl.ScreenUpdating = True
        Return nAdd
fail:
        PT_LastErr = Err.Description
        On Error Resume Next
        xl.CutCopyMode = False
        xl.Calculation = calcWas
        xl.ScreenUpdating = True
        Return -1
    End Function

    ' Lengthens a named table so it ends at newLastRow (formats of its last row are copied down). Never shortens.
    Private Sub ResizeName(ByVal wb As Object, ByVal xl As Object, ByVal nameS As String, ByVal newLastRow As Integer)
        Dim r As Object, ws As Object, lastR As Integer, c1 As Integer
        r = PT_Rng(wb, nameS)
        If r Is Nothing Then Exit Sub
        ws = r.Worksheet
        lastR = CInt(r.Row) + CInt(r.Rows.Count) - 1
        If lastR >= newLastRow Then Exit Sub
        c1 = CInt(r.Column) + CInt(r.Columns.Count) - 1
        Try
            ws.Range(ws.Cells(lastR, r.Column), ws.Cells(lastR, c1)).Copy()
            ws.Range(ws.Cells(lastR + 1, r.Column), ws.Cells(newLastRow, c1)).PasteSpecial(-4122)
            xl.CutCopyMode = False
        Catch
        End Try
        wb.Names(nameS).RefersTo = "='" & CStr(ws.Name) & "'!" & CStr(ws.Range(ws.Cells(r.Row, r.Column), ws.Cells(newLastRow, c1)).Address)
    End Sub

    ' Import level of the last run ("bodies" or "components"), kept in the workbook name Imp_Level. "" = unknown.
    Private Function GetLevelUsed(ByVal wb As Object) As String
        Dim s As String = ""
        Try
            s = CStr(wb.Names("Imp_Level").RefersTo)
        Catch
        End Try
        Return Replace(Replace(s, "=", ""), """", "")
    End Function

    Private Sub SetLevelUsed(ByVal wb As Object, ByVal lvl As String)
        wb.Names.Add("Imp_Level", "=""" & lvl & """")
    End Sub

    ' ---------------------------------------------------------------- texts
    Private Function FrameText(ByVal fr() As Double) As String
        Return "Board X = model " & AxisText(fr, 0) & ", Board Y = model " & AxisText(fr, 3) & ", Board Z = model " & AxisText(fr, 6) & "."
    End Function

    Private Function AxisText(ByVal fr() As Double, ByVal o As Integer) As String
        Dim big As Integer = 0
        For k As Integer = 1 To 2
            If Math.Abs(fr(o + k)) > Math.Abs(fr(o + big)) Then big = k
        Next
        If Math.Abs(fr(o + big)) > 0.9999 Then
            Return IIf(fr(o + big) > 0, "+", "-") & Mid("XYZ", big + 1, 1)
        End If
        Return "direction (" & Format(fr(o), "0.000") & ", " & Format(fr(o + 1), "0.000") & ", " & Format(fr(o + 2), "0.000") & ")"
    End Function

    Private Function ChangeSummary(ByVal counts As Dict, ByVal nLog As Integer, ByVal hadOld As Boolean, ByVal levelChanged As Boolean) As String
        Dim s As String = ""
        If levelChanged Then Return "not compared (import level changed, see messages)."
        If Not hadOld Then Return "none (first import)."
        If nLog = 0 Then Return "none."
        For Each it As Object In counts.Keys
            s = s & CStr(counts(it)) & " " & CStr(it) & ", "
        Next
        Return Left(s, Len(s) - 2) & "." & vbCrLf & nLog & " rows added to the Change log tab. Check them there (type Yes under Checked)."
    End Function

End Module
