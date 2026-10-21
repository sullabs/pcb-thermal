Option Strict Off
Option Explicit On

Imports System
Imports Microsoft.VisualBasic

' Port of src/BuildSpreader.bas (milestone M4). Same behaviour and messages as the VBA macro.
'
' Completes the heat spreader part(s) and builds the TIM parts in the thermal assembly, from the Spreader tab (top plate)
' and the Spreader bottom tab (plate under the PCB). The user draws the plate (a part with "Spreader" in its name, sketch
' named Outline on the Front Plane); the macro adds PT_Under, PT_Plate, Clear_<Key> cuts and Ped_<RefDes> pedestals,
' names faces Top, Underside, Contact_<RefDes>, and builds TIM parts <prefix>TIM_<RefDes>.sldprt (TIM mode Body).
' With no top spreader part it offers a rectangle the size of the PCB. Update runs change features in place, rebuild what
' moved or resized, add new ones, suppress unwanted ones, and log every change. Refuses to run while the Change log has
' unchecked rows.
Public Module BuildSpreader

    Private Const MACRO_NAME As String = "BuildSpreader"
    Private Const swDefaultTemplatePart As Integer = 8
    Private Const swUnSuppressFeature As Integer = 1
    Private Const swSuppressFeature As Integer = 0
    Private Const swAllConfiguration As Integer = 2

    ' Spreader tab chip table offsets (Spr_Chips, column B = 0)
    Private Const C_REF As Integer = 0, C_KEY As Integer = 1, C_X As Integer = 2, C_Y As Integer = 3, C_L As Integer = 4, C_W As Integer = 5
    Private Const C_TOP As Integer = 6, C_COV As Integer = 7, C_TIMMODE As Integer = 9, C_TIMT As Integer = 13, C_PED As Integer = 14, C_CHECK As Integer = 16
    ' clearance cut table offsets (Spr_Cuts, column B = 0)
    Private Const K_KEY As Integer = 1, K_NAME As Integer = 2, K_X As Integer = 3, K_Y As Integer = 4, K_L As Integer = 5, K_W As Integer = 6

    ' wanted chips (of the side being built; for the bottom side chBox(,4) is the chip bottom face)
    Private chRef() As String, chKey() As String, chBox(,) As Double, chPed() As Double, chTim() As Double, chBody() As Boolean, nCh As Integer
    ' wanted cuts
    Private cuKey() As String, cuName() As String, cuBox(,) As Double, nCu As Integer
    ' TIM parts wanted, both sides
    Private tRef(199) As String, tBox(199, 4) As Double, tTim(199) As Double, tBot(199) As Boolean, nT As Integer
    Private skipped As Collection, failMsgs As Collection, notes As Collection
    Private doLog As Boolean
    Private bottomSide As Boolean                                ' the side being built
    Private wsSide As Object                                     ' Spreader bottom tab (when bottomSide)
    Private rptBody As String, resMsg As String
    Private pOx As Double, pOy As Double, pOz As Double          ' spreader part origin in board coordinates, mm
    Private nFeatAdded As Integer, nFeatChanged As Integer, nFeatSupp As Integer, nTimAdded As Integer, nTimChanged As Integer, nTimSupp As Integer

    Public Sub Main()
        PT_LogReset()
        skipped = New Collection() : failMsgs = New Collection() : notes = New Collection()
        nFeatAdded = 0 : nFeatChanged = 0 : nFeatSupp = 0 : nTimAdded = 0 : nTimChanged = 0 : nTimSupp = 0
        nT = 0 : rptBody = "" : resMsg = "" : bottomSide = False : doLog = False

        ' ---------- workbook ----------
        Dim xl As Object, wb As Object
        xl = PT_GetExcel()
        wb = PT_PickWorkbook(xl, MACRO_NAME, "Spreader|CAD geometry|Settings|Change log")
        If wb Is Nothing Then Exit Sub
        If Not PT_NeedNames(wb, "Spr_Chips|Spr_Cuts|Spr_ZUnder|Spr_Thickness|Geo_PCB_Used|Log_Header|Log_Unchecked|" &
                                "Set_Prefix|Set_TolPos|Set_TolSize|Set_PartTemplate|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
        If PT_RefuseIfUnchecked(wb, MACRO_NAME) Then Exit Sub

        Dim pcbL As Double, pcbW As Double, rp As Object, tolPos As Double, tolSize As Double
        xl.Calculate()
        rp = PT_Rng(wb, "Geo_PCB_Used")
        pcbL = PT_Num(rp.Cells(1, 1).Value, 0)
        pcbW = PT_Num(rp.Cells(2, 1).Value, 0)
        tolPos = PT_NameNum(wb, "Set_TolPos", 0.05)
        tolSize = PT_NameNum(wb, "Set_TolSize", 0.05)

        ' ---------- assembly and files ----------
        Dim asm As Object, asmPath As String, partTmpl As String, outDir As String, prefix As String, rptPath As String = ""
        outDir = PT_OutputFolder(wb)
        prefix = PT_SafeFile(PT_Prefix(wb))
        partTmpl = PT_Template(wb, "Set_PartTemplate", swDefaultTemplatePart)
        If partTmpl = "" Then
            MsgBox("BuildSpreader did not run. The part template was not found. Set its full path on the Settings tab.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        asmPath = PT_AsmPath(wb)
        On Error GoTo swFail
        asm = PT_OpenAsm(asmPath)
        If asm Is Nothing Then
            On Error GoTo 0
            MsgBox("BuildSpreader did not run. The thermal assembly is not there yet:" & vbCrLf & asmPath & vbCrLf &
                   "Run BuildBoard first.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        PT_ActivateDoc(asm)
        PT_LeaveHeatOnly(asm)
        asm.ResolveAllLightWeightComponents(False)

        ' ---------- the two sides ----------
        Dim res As Integer, nBuilt As Integer, errs As Integer = 0, warns As Integer = 0
        For pass As Integer = 0 To 1
            bottomSide = (pass = 1)
            wsSide = Nothing
            If bottomSide Then
                wsSide = PT_Sheet(wb, "Spreader bottom")
                If wsSide Is Nothing Then Exit For
            End If
            res = DoSide(wb, asm, partTmpl, outDir, prefix, tolPos, tolSize, pcbL, pcbW)
            If res < 0 Then Exit Sub
            nBuilt = nBuilt + res
        Next
        If nBuilt = 0 Then
            MsgBox("BuildSpreader did nothing: no spreader part and no covered chips." & vbCrLf &
                   "Mark chips Covered = Yes on the Spreader tab (top) or the Spreader bottom tab, and draw the plate outline.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If

        ' ---------- TIM parts (both sides) ----------
        SyncTims(asm, partTmpl, outDir, prefix, tolPos, tolSize)

        PT_ActivateDoc(asm)
        asm.ForceRebuild3(False)
        SwSave3(asm, PT_swSaveSilent + PT_swSaveReferenced, errs, warns)
        On Error GoTo 0

        ' ---------- report, Change log ----------
        Dim nLog As Integer
        rptPath = outDir & prefix & "BuildSpreader_report.txt"
        WriteReport(rptPath, asmPath)
        On Error GoTo xlFail
        xl.ScreenUpdating = False
        nLog = PT_LogFlush(wb)
        xl.ScreenUpdating = True
        xl.Calculate()
        wb.Save()
        On Error GoTo 0

        Dim s As String
        s = "Written and saved to the spreader part(s) and the thermal assembly:" & vbCrLf & asmPath & vbCrLf & vbCrLf & resMsg &
            "Cuts and pedestals: " & nFeatAdded & " added, " & nFeatChanged & " rebuilt, " & nFeatSupp & " suppressed." & vbCrLf &
            "TIM parts: " & nTimAdded & " added, " & nTimChanged & " changed, " & nTimSupp & " suppressed." & vbCrLf
        If nLog > 0 Then s = s & nLog & " rows added to the Change log tab. Check them there." & vbCrLf
        If skipped.Count > 0 Then s = s & skipped.Count & " covered chips skipped (see the Check column on the Spreader tabs)." & vbCrLf
        If notes.Count > 0 Then s = s & notes.Count & " notes (listed in the report)." & vbCrLf
        If failMsgs.Count > 0 Then s = s & failMsgs.Count & " problems (listed in the report)." & vbCrLf
        s = s & vbCrLf & "Report: " & rptPath
        MsgBox(s, IIf(failMsgs.Count > 0, vbExclamation, vbInformation), MACRO_NAME)
        Exit Sub

swFail:
        If Err.Number >= vbObjectError + 34 And Err.Number <= vbObjectError + 36 Then
            MsgBox("BuildSpreader did not run (" & IIf(bottomSide, "bottom", "top") & " spreader): " & Err.Description, vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        MsgBox("BuildSpreader stopped in SolidWorks (" & IIf(bottomSide, "bottom", "top") & " spreader): " & Err.Description & vbCrLf &
               "Nothing was saved. Close the spreader part and the assembly without saving, send this message, and run again.",
               vbExclamation, MACRO_NAME)
        Exit Sub
xlFail:
        xl.ScreenUpdating = True
        MsgBox("The spreader was saved, but writing the Change log failed: " & Err.Description & vbCrLf &
               "Close any cell you are editing in Excel. The changes are listed in " & rptPath, vbExclamation, MACRO_NAME)
    End Sub

    ' Builds or updates one plate (top, or bottom when bottomSide). Returns 1 built, 0 nothing to do, -1 stop the macro.
    Private Function DoSide(ByVal wb As Object, ByVal asm As Object, ByVal partTmpl As String, ByVal outDir As String, ByVal prefix As String,
                            ByVal tolPos As Double, ByVal tolSize As Double, ByVal pcbL As Double, ByVal pcbW As Double) As Integer
        Dim zu As Double, tPl As Double, zuP As Double, sideTxt As String
        Dim sprComp As Object, md As Object = Nothing, ans As MsgBoxResult, newRev As Boolean = False
        Dim sprPath As String = "", sprName As String = "", sprMsg As String = "", nFound As Integer, errs As Integer = 0, warns As Integer = 0
        sideTxt = IIf(bottomSide, "Bottom", "Top")
        zu = SNum(wb, "Spr_ZUnder", 0)
        tPl = SNum(wb, "Spr_Thickness", 0)
        ReadChips(wb)
        ReadCuts(wb)
        sprComp = FindSpreaderComp(asm, nFound, bottomSide)
        If nFound > 1 Then
            MsgBox("BuildSpreader did not run. The assembly has " & nFound & " " & LCase(sideTxt) & " spreader parts that are not suppressed." & vbCrLf &
                   "Select the one to complete in the assembly (click it in the tree) and run again.", vbExclamation, MACRO_NAME)
            Return -1
        End If
        If sprComp Is Nothing And nCh = 0 Then Return 0                 ' nothing for this side
        If bottomSide And nCh = 0 And zu = 0 Then                       ' a bottom part exists but no chip is covered: leave it as it is
            notes.Add("Bottom: no chip is covered on the Spreader bottom tab, so the bottom spreader part was left unchanged.")
            Return 0
        End If
        If tPl <= 0 Or (Not bottomSide And zu <= 0) Or (bottomSide And nCh = 0 And zu = 0) Then
            MsgBox("BuildSpreader did not run (" & LCase(sideTxt) & " spreader). On its tab the plate thickness (" & tPl & ") must be above 0 and " &
                   "the face height (" & zu & ") must be set." & vbCrLf &
                   "Mark at least one chip Covered = Yes, or set the face height to Manual.", vbExclamation, MACRO_NAME)
            Return -1
        End If
        pOx = 0 : pOy = 0 : pOz = 0
        If sprComp Is Nothing Then
            If bottomSide Then
                notes.Add("Bottom: " & nCh & " chips are covered but there is no spreader part with Spreader and Bottom in its name. Draw its outline and run again.")
                Return 0
            End If
            ans = MsgBox("The thermal assembly has no spreader part yet." & vbCrLf & vbCrLf &
                         "To draw your own: in the assembly choose Insert > Component > New Part, put 'Spreader' in its name, " &
                         "pick the Front Plane (or a face parallel to it), sketch the outline, exit the sketch, name the sketch Outline, " &
                         "save the part as a file, then run this macro again." & vbCrLf & vbCrLf &
                         "Yes = make a rectangle the size of the PCB now (its Outline sketch can be edited)" & vbCrLf &
                         "No = stop", vbYesNo + vbQuestion, MACRO_NAME)
            If ans <> vbYes Then Return -1
            If pcbL <= 0 Or pcbW <= 0 Then
                MsgBox("BuildSpreader did not run. CAD geometry has no PCB size.", vbExclamation, MACRO_NAME)
                Return -1
            End If
            newRev = True
        End If

        If newRev Then
            doLog = False
            zuP = zu
            sprPath = outDir & prefix & "Spreader_r" & Format(NextRev(outDir, prefix), "00") & ".sldprt"
            md = swApp.NewDocument(partTmpl, 0, 0, 0)
            If md Is Nothing Then Err.Raise(vbObjectError + 30, , "SolidWorks could not make a new part from " & partTmpl)
            MakeBase(md, zu, tPl, pcbL, pcbW)
            sprName = Mid(sprPath, InStrRev(sprPath, "\") + 1)
            PT_SetProp(md, "Role", "Spreader")
            PT_SetProp(md, "Key", LCase(Left(sprName, Len(sprName) - 7)))
            PT_SetProp(md, "Name", "Spreader " & sprName)
            PT_SetProp(md, "BuiltBy", MACRO_NAME)
            If Not SwSaveAs(md, sprPath, errs, warns) Then
                Err.Raise(vbObjectError + 31, , "could not save " & sprPath & " (error " & errs & ")")
            End If
            sprMsg = "New rectangular spreader made from the PCB size."
        Else
            If CBool(sprComp.IsVirtual) Then Err.Raise(vbObjectError + 34, , "the spreader is a virtual part inside the assembly. " &
                                      "Right-click it in the tree, choose Save Part (in External File), and run again.")
            If Not ReadOffset(sprComp) Then Err.Raise(vbObjectError + 35, , "the spreader part is turned relative to the assembly. " &
                                      "Make it again on the Front Plane (or a face parallel to it) so its X, Y and Z match the board.")
            sprComp.SetSuppression2(PT_swFullyResolved)          ' a suppressed or lightweight part has no bodies to edit or name
            md = sprComp.GetModelDoc2
            If md Is Nothing Then Err.Raise(vbObjectError + 37, , "the spreader part could not be opened")
            sprPath = CStr(md.GetPathName)
            If sprPath = "" Then Err.Raise(vbObjectError + 34, , "the spreader part has no file yet. Save it first.")
            sprName = Mid(sprPath, InStrRev(sprPath, "\") + 1)
            zuP = zu - pOz
            doLog = PT_FindFeature(md, "PT_Plate") IsNot Nothing
            PT_ActivateDoc(md)
            md.FeatureManager.EditRollback(1, "")                ' swMoveRollbackBarToEnd: features below a rollback bar are not built
            If doLog And IsEmpty(md.GetBodies2(PT_swSolidBody, True)) Then   ' a part loaded only as a component has no bodies to read: open its own window
                Dim md2 As Object
                md2 = SwOpenDoc6(sprPath, PT_swDocPART, PT_swOpenSilent, errs, warns)
                If md2 IsNot Nothing Then
                    md = md2
                    PT_ActivateDoc(md)
                End If
            End If
            If Not doLog Then
                PrepareOutline(md)
                MakeUserPlate(md, zuP, tPl)
                If PT_GetProp(md, "Role") = "" Then PT_SetProp(md, "Role", "Spreader")
                If PT_GetProp(md, "Key") = "" Then PT_SetProp(md, "Key", LCase(Left(sprName, Len(sprName) - 7)))
                If PT_GetProp(md, "Name") = "" Then PT_SetProp(md, "Name", "Spreader " & sprName)
                PT_SetProp(md, "Side", sideTxt)
                PT_SetProp(md, "BuiltBy", MACRO_NAME)
                sprMsg = sideTxt & " spreader completed from your Outline sketch."
            Else
                SyncPlate(md, zuP, tPl, tolPos, tolSize)
                sprMsg = sideTxt & " spreader updated."
            End If
        End If

        CheckOrigin(md)
        SyncCuts(md, zuP, tolPos)
        SyncPedestals(md, zuP, tolPos, tolSize)
        md.ForceRebuild3(False)
        NameSpreaderFaces(md, zuP, tPl)
        SwSave3(md, PT_swSaveSilent, errs, warns)

        ' ---------- spreader in the assembly ----------
        PT_ActivateDoc(asm)
        If newRev Then
            Dim c As Object, ttl As String
            ttl = CStr(md.GetTitle)
            c = asm.AddComponent5(sprPath, 0, "", False, "", 0.0, 0.0, 0.0)
            If c Is Nothing Then Err.Raise(vbObjectError + 33, , "could not insert " & sprPath & " into the assembly")
            PT_PlaceAndFix(asm, c, 0, 0, 0)
            swApp.CloseDoc(ttl)
            PT_ActivateDoc(asm)
        Else
            If Not CBool(sprComp.IsFixed) Then PT_FixComp(asm, sprComp)
        End If

        ' ---------- remember the TIM parts and the report lines of this side ----------
        For i As Integer = 0 To nCh - 1
            If chBody(i) And chTim(i) > 0.0005 And nT <= UBound(tRef) Then
                tRef(nT) = chRef(i)
                tBox(nT, 0) = chBox(i, 0) : tBox(nT, 1) = chBox(i, 1) : tBox(nT, 2) = chBox(i, 2) : tBox(nT, 3) = chBox(i, 3) : tBox(nT, 4) = chBox(i, 4)
                tTim(nT) = chTim(i) : tBot(nT) = bottomSide
                nT = nT + 1
            End If
        Next
        rptBody = rptBody & UCase(sideTxt) & " SPREADER: " & sprPath & vbCrLf &
                  "  plate face toward the board at Z " & zu & " mm, plate " & tPl & " mm, part origin at (" & Round(pOx, 3) & ", " & Round(pOy, 3) & ", " & Round(pOz, 3) & ") mm" & vbCrLf &
                  "  chips (" & nCh & "):" & vbCrLf
        For i As Integer = 0 To nCh - 1
            rptBody = rptBody & "    " & chRef(i) & ": face Z " & chBox(i, 4) & ", TIM " & chTim(i) & IIf(chBody(i), " (body)", " (contact resistance)") &
                      ", pedestal " & chPed(i) & " mm" & vbCrLf
        Next
        rptBody = rptBody & "  clearance cuts (" & nCu & "):" & vbCrLf
        For i As Integer = 0 To nCu - 1
            rptBody = rptBody & "    " & cuName(i) & vbCrLf
        Next
        rptBody = rptBody & vbCrLf
        resMsg = resMsg & sprMsg & " (" & sprName & ") Face at Z " & zu & " mm, plate " & tPl & " mm." & vbCrLf
        Return 1
    End Function

    ' Named range of the side being built: workbook name for the top tab, tab-level name for the bottom tab.
    Private Function SRng(ByVal wb As Object, ByVal nameS As String) As Object
        Try
            If bottomSide Then Return wsSide.Names(nameS).RefersToRange
            Return wb.Names(nameS).RefersToRange
        Catch
            Return Nothing
        End Try
    End Function

    Private Function SNum(ByVal wb As Object, ByVal nameS As String, ByVal dflt As Double) As Double
        Dim r As Object = SRng(wb, nameS)
        If r Is Nothing Then Return dflt
        Return PT_Num(r.Cells(1, 1).Value, dflt)
    End Function

    ' ================================================================= Excel side
    Private Sub ReadChips(ByVal wb As Object)
        Dim v(,) As Object, refD As String, seen As New Dict(), ck As String
        v = ToTable(SRng(wb, "Spr_Chips").Value)
        nCh = 0
        Dim nr As Integer = UBound(v, 1)
        ReDim chRef(nr) : ReDim chKey(nr) : ReDim chBox(nr, 4)
        ReDim chPed(nr) : ReDim chTim(nr) : ReDim chBody(nr)
        For r As Integer = 1 To nr
            refD = PT_Str(v(r, 1 + C_REF))
            If refD <> "" And UCase(PT_Str(v(r, 1 + C_COV))) = "YES" Then
                ck = PT_Str(v(r, 1 + C_CHECK))
                If ck <> "OK" Then
                    skipped.Add(refD & ": " & ck)
                ElseIf seen.Exists(UCase(refD)) Then
                    skipped.Add(refD & ": listed twice, second row ignored")
                Else
                    seen.Add(UCase(refD), 1)
                    chRef(nCh) = refD
                    chKey(nCh) = PT_Str(v(r, 1 + C_KEY))
                    chBox(nCh, 0) = PT_Num(v(r, 1 + C_X), 0)
                    chBox(nCh, 1) = PT_Num(v(r, 1 + C_Y), 0)
                    chBox(nCh, 2) = PT_Num(v(r, 1 + C_L), 0)
                    chBox(nCh, 3) = PT_Num(v(r, 1 + C_W), 0)
                    chBox(nCh, 4) = PT_Num(v(r, 1 + C_TOP), 0)
                    chPed(nCh) = PT_Num(v(r, 1 + C_PED), 0)
                    chTim(nCh) = PT_Num(v(r, 1 + C_TIMT), 0)
                    chBody(nCh) = (UCase(PT_Str(v(r, 1 + C_TIMMODE))) = "BODY")
                    nCh = nCh + 1
                End If
            End If
        Next
    End Sub

    Private Sub ReadCuts(ByVal wb As Object)
        Dim v(,) As Object, keyS As String, seen As New Dict()
        v = ToTable(SRng(wb, "Spr_Cuts").Value)
        nCu = 0
        Dim nr As Integer = UBound(v, 1)
        ReDim cuKey(nr) : ReDim cuName(nr) : ReDim cuBox(nr, 3)
        For r As Integer = 1 To nr
            keyS = PT_Str(v(r, 1 + K_KEY))
            If keyS <> "" And PT_IsNum(v(r, 1 + K_X)) And PT_IsNum(v(r, 1 + K_L)) Then
                If Not seen.Exists(keyS) Then
                    seen.Add(keyS, 1)
                    cuKey(nCu) = keyS
                    cuName(nCu) = PT_Str(v(r, 1 + K_NAME))
                    cuBox(nCu, 0) = PT_Num(v(r, 1 + K_X), 0)
                    cuBox(nCu, 1) = PT_Num(v(r, 1 + K_Y), 0)
                    cuBox(nCu, 2) = PT_Num(v(r, 1 + K_L), 0)
                    cuBox(nCu, 3) = PT_Num(v(r, 1 + K_W), 0)
                    nCu = nCu + 1
                End If
            End If
        Next
    End Sub

    ' ================================================================= spreader part
    ' Plane, outline and plate of a new spreader. Checks where SolidWorks put them and flips if needed.
    Private Sub MakeBase(ByVal md As Object, ByVal zu As Double, ByVal tPl As Double, ByVal pcbL As Double, ByVal pcbW As Double)
        Dim front As Object, pl As Object, sk As Object, feat As Object, zr As Double()
        front = PT_RefPlane(md, 1)
        If front Is Nothing Then Err.Raise(vbObjectError + 40, , "the part template has no reference plane")
        Dim nBefore As Integer = CountPlanes(md)
        md.ClearSelection2(True)
        front.Select2(False, 0)
        md.CreatePlaneAtOffset3(zu / 1000.0, False, True)
        md.ClearSelection2(True)
        If CountPlanes(md) <= nBefore Then Err.Raise(vbObjectError + 41, , "the underside plane was not made")
        pl = PT_LastFeatureOfType(md, "RefPlane")
        pl.Name = "PT_Under"
        sk = PT_SketchRect(md, pl, 0, 0, pcbL, pcbW, zu)
        sk.Name = "Outline"
        feat = PT_Boss(md, sk, tPl, False)
        feat.Name = "PT_Plate"

        zr = PT_BodyZRange(md)
        If zr Is Nothing Then Err.Raise(vbObjectError + 42, , "the plate has no body")
        If (zr(0) + zr(1)) / 2 < 0 Then                      ' plane went below the board: turn it over
            PT_SetPlaneOffset(md, pl, zu, True)
            md.EditRebuild3()
            zr = PT_BodyZRange(md)
        End If
        If Math.Abs(zr(0) - zu) > 0.001 Then                 ' plate went down from the plane: turn it up
            PT_FlipExtrude(md, feat)
            md.EditRebuild3()
            zr = PT_BodyZRange(md)
        End If
        If Math.Abs(zr(0) - zu) > 0.001 Or Math.Abs(zr(1) - (zu + tPl)) > 0.001 Then
            Err.Raise(vbObjectError + 43, , "the plate ended up at Z " & Round(zr(0), 3) & " to " & Round(zr(1), 3) &
                      " mm, expected " & zu & " to " & (zu + tPl))
        End If
    End Sub

    Private Sub SyncPlate(ByVal md As Object, ByVal zu As Double, ByVal tPl As Double, ByVal tolPos As Double, ByVal tolSize As Double)
        Dim pl As Object, feat As Object, oldZ As Double, oldT As Double, byOffset As Boolean, refit As Boolean = False, sideTxt As String
        sideTxt = IIf(bottomSide, "Bottom spreader", "Spreader")
        pl = PT_FindFeature(md, "PT_Under")
        feat = PT_FindFeature(md, "PT_Plate")
        oldT = PT_ExtrudeDepth(feat)
        byOffset = PlateStartsByOffset(md, feat)             ' plate made from the user's Outline (start offset) or the PCB rectangle
        If byOffset Then oldZ = PlaneZ(pl) Else oldZ = PT_PlaneOffset(pl)
        If Math.Abs(oldZ - zu) > tolPos Then
            If byOffset Then
                SetPlaneZ(md, pl, zu)
                refit = True
            Else
                PT_SetPlaneOffset(md, pl, zu, False)
            End If
            PT_LogAdd(MACRO_NAME, "PT_Under", sideTxt & " plate face height", "", "moved", "Z", Round(oldZ, 3), zu)
            nFeatChanged = nFeatChanged + 1
        End If
        If Math.Abs(oldT - tPl) > tolSize Then
            If byOffset Then refit = True Else PT_SetExtrudeDepth(md, feat, tPl)
            PT_LogAdd(MACRO_NAME, "PT_Plate", sideTxt & " plate thickness", "", "resized", "Thickness", Round(oldT, 3), tPl)
            nFeatChanged = nFeatChanged + 1
        End If
        If refit Then FitPlate(md, feat, zu, tPl)
        md.EditRebuild3()
    End Sub

    Private Sub SyncCuts(ByVal md As Object, ByVal zu As Double, ByVal tolPos As Double)
        Dim want As New Dict(), featName As String, p As String
        For i As Integer = 0 To nCu - 1
            featName = "Clear_" & PT_SafeFile(cuKey(i))
            want.Add(featName, 1)
            p = PT_Um(cuBox(i, 0)) & "|" & PT_Um(cuBox(i, 1)) & "|" & PT_Um(cuBox(i, 2)) & "|" & PT_Um(cuBox(i, 3))
            SyncOne(md, featName, p, cuKey(i), "Cut for " & cuName(i), "", zu,
                    cuBox(i, 0) - pOx, cuBox(i, 1) - pOy, cuBox(i, 0) - pOx + cuBox(i, 2), cuBox(i, 1) - pOy + cuBox(i, 3), 0, False)
        Next
        SuppressOthers(md, "Clear_", want)
    End Sub

    Private Sub SyncPedestals(ByVal md As Object, ByVal zu As Double, ByVal tolPos As Double, ByVal tolSize As Double)
        Dim want As New Dict(), featName As String, p As String
        For i As Integer = 0 To nCh - 1
            If chPed(i) > 0.0005 Then
                featName = "Ped_" & PT_SafeFile(chRef(i))
                want.Add(featName, 1)
                p = PT_Um(chBox(i, 0)) & "|" & PT_Um(chBox(i, 1)) & "|" & PT_Um(chBox(i, 2)) & "|" & PT_Um(chBox(i, 3)) & "|" & PT_Um(chPed(i))
                SyncOne(md, featName, p, featName, "Pedestal for " & chRef(i), chRef(i), zu,
                        chBox(i, 0) - pOx, chBox(i, 1) - pOy, chBox(i, 0) - pOx + chBox(i, 2), chBox(i, 1) - pOy + chBox(i, 3), chPed(i), True)
            Else
                notes.Add(chRef(i) & ": no pedestal (gap filled by the TIM). Its contact is the Underside face over the chip.")
            End If
        Next
        SuppressOthers(md, "Ped_", want)
    End Sub

    ' One macro feature (cut or pedestal): leave it, bring it back, rebuild it, or make it.
    Private Sub SyncOne(ByVal md As Object, ByVal featName As String, ByVal params As String, ByVal logKey As String,
                        ByVal logName As String, ByVal refD As String, ByVal zu As Double,
                        ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double,
                        ByVal pedH As Double, ByVal isPed As Boolean)
        Dim f As Object, oldP As String, pl As Object, sk As Object, feat As Object, zmin As Double, replaced As Boolean = False, pedMin As Double
        Dim emptyNames As Object = Nothing
        On Error GoTo failed
        f = PT_FindFeature(md, featName)
        oldP = PT_GetProp(md, "PT_" & featName)
        If f IsNot Nothing Then
            If oldP = params Then
                If CBool(f.IsSuppressed) Then
                    f.SetSuppression2(swUnSuppressFeature, swAllConfiguration, emptyNames)
                    If doLog Then PT_LogAdd(MACRO_NAME, logKey, logName, refD, "added", "Brought back (was suppressed)", "", "")
                    nFeatChanged = nFeatChanged + 1
                End If
                Exit Sub
            End If
            PT_DeleteFeature(md, f)
            replaced = True
        End If

        pl = PT_FindFeature(md, "PT_Under")
        sk = PT_SketchRect(md, pl, x0, y0, x1, y1, zu)
        sk.Name = featName & "_S"
        If isPed Then
            feat = PT_Boss(md, sk, pedH, Not bottomSide)     ' top plate: pedestal goes down; bottom plate: up
            md.EditRebuild3()
            If SolidCount(md) > 1 Then                        ' not joined to the plate: the chip is outside the outline
                PT_DeleteFeature(md, feat)
                notes.Add(refD & ": no pedestal, the chip is outside the plate outline. Redraw the Outline over it or set Covered = No.")
                Exit Sub
            End If
            pedMin = IIf(bottomSide, zu, zu - pedH)             ' lowest Z of the pedestal
            zmin = PT_FeatureZMin(feat)
            If Math.Abs(zmin - pedMin) > 0.002 Then              ' went the wrong way: turn it
                PT_FlipExtrude(md, feat)
                md.EditRebuild3()
                zmin = PT_FeatureZMin(feat)
                If Math.Abs(zmin - pedMin) > 0.002 Then Err.Raise(vbObjectError + 50, , "pedestal lowest point is at Z " & Round(zmin, 3) &
                                                                  ", expected " & Round(pedMin, 3))
            End If
        Else
            feat = PT_CutThrough(md, sk)
            If feat Is Nothing Then                            ' the tall part is outside the plate outline: nothing to cut
                PT_DeleteFeature(md, sk)
                notes.Add(logName & ": skipped, the part is outside the plate outline.")
                Exit Sub
            End If
        End If
        feat.Name = featName
        PT_SetProp(md, "PT_" & featName, params)
        If replaced Then
            nFeatChanged = nFeatChanged + 1
            PT_LogAdd(MACRO_NAME, logKey, logName, refD, IIf(SameSize(oldP, params), "moved", "resized"), "Box", ParamText(oldP), ParamText(params))
        Else
            nFeatAdded = nFeatAdded + 1
            If doLog Then PT_LogAdd(MACRO_NAME, logKey, logName, refD, "added", "Box", "", ParamText(params))
        End If
        Exit Sub
failed:
        failMsgs.Add(featName & ": " & Err.Description)
    End Sub

    ' Suppresses macro features with this prefix that are not wanted any more.
    Private Sub SuppressOthers(ByVal md As Object, ByVal prefixS As String, ByVal want As Dict)
        Dim f As Object, emptyNames As Object = Nothing, nm As String
        f = md.FirstFeature
        Do While f IsNot Nothing
            nm = CStr(f.Name)
            If Left(nm, Len(prefixS)) = prefixS And Right(nm, 2) <> "_S" Then
                If Not want.Exists(nm) And Not CBool(f.IsSuppressed) Then
                    f.SetSuppression2(swSuppressFeature, swAllConfiguration, emptyNames)
                    PT_LogAdd(MACRO_NAME, nm, IIf(prefixS = "Ped_", "Pedestal", "Cut"), IIf(prefixS = "Ped_", Mid(nm, 5), ""),
                              "removed", "Suppressed in the spreader", "", "")
                    nFeatSupp = nFeatSupp + 1
                End If
            End If
            f = f.GetNextFeature
        Loop
    End Sub

    Private Sub NameSpreaderFaces(ByVal md As Object, ByVal zu As Double, ByVal tPl As Double)
        Dim big As Double = 1.0E+30
        ' Top = outer face, Underside = face toward the board. Top plate: outer face up. Bottom plate: outer face down.
        If bottomSide Then
            If Not PT_NameFlatFace(md, False, zu - tPl, "Top", -big, -big, big, big) Then failMsgs.Add("Could not find the plate outer face to name it Top")
            If Not PT_NameFlatFace(md, True, zu, "Underside", -big, -big, big, big) Then failMsgs.Add("Could not find the plate face toward the board to name it Underside")
        Else
            If Not PT_NameFlatFace(md, True, zu + tPl, "Top", -big, -big, big, big) Then
                failMsgs.Add("Could not find the plate top face to name it Top (faces seen: " & PT_AllFaces(md).Count & ", part " & CStr(md.GetTitle) &
                             ", active " & CStr(swApp.ActiveDoc.GetTitle) & ")")
            End If
            If Not PT_NameFlatFace(md, False, zu, "Underside", -big, -big, big, big) Then failMsgs.Add("Could not find the plate underside to name it Underside")
        End If
        For i As Integer = 0 To nCh - 1
            If chPed(i) > 0.0005 And PT_FindFeature(md, "Ped_" & PT_SafeFile(chRef(i))) IsNot Nothing Then
                If Not PT_NameFlatFace(md, bottomSide, IIf(bottomSide, zu + chPed(i), zu - chPed(i)), "Contact_" & PT_SafeFile(chRef(i)),
                                       chBox(i, 0) - pOx, chBox(i, 1) - pOy, chBox(i, 0) - pOx + chBox(i, 2), chBox(i, 1) - pOy + chBox(i, 3)) Then
                    failMsgs.Add("Could not find the pedestal bottom of " & chRef(i) & " to name it Contact_" & chRef(i))
                End If
            End If
        Next
    End Sub

    ' ================================================================= TIM parts
    Private Sub SyncTims(ByVal asm As Object, ByVal partTmpl As String, ByVal outDir As String, ByVal prefix As String,
                         ByVal tolPos As Double, ByVal tolSize As Double)
        Dim have As New Dict(), wasSupp As New Dict(), want As New Dict(), keyS As String, res As Integer, fpath As String
        Dim topN As String, botN As String, nameS As String, z0 As Double
        PT_ReadTagged(asm, "TIM", have, wasSupp)
        ' Face names: _Top is the face toward the plate, _Bottom the face on the chip, on both sides of the board.
        ' Top plate: the TIM sits on the chip top (z0 = chip top). Bottom plate: it hangs under the chip bottom (z0 = chip bottom - TIM).
        For i As Integer = 0 To nT - 1
            keyS = "tim_" & LCase(PT_SafeFile(tRef(i)))
            If Not want.Exists(keyS) Then
                want.Add(keyS, 1)
                nameS = "TIM " & tRef(i)
                If tBot(i) Then
                    z0 = tBox(i, 4) - tTim(i)
                    topN = "TIM_" & PT_SafeFile(tRef(i)) & "_Bottom"
                    botN = "TIM_" & PT_SafeFile(tRef(i)) & "_Top"
                Else
                    z0 = tBox(i, 4)
                    topN = "TIM_" & PT_SafeFile(tRef(i)) & "_Top"
                    botN = "TIM_" & PT_SafeFile(tRef(i)) & "_Bottom"
                End If
                If have.Exists(keyS) Then
                    If CBool(wasSupp(keyS)) And doLog Then PT_LogAdd(MACRO_NAME, keyS, nameS, tRef(i), "added", "Brought back (was suppressed)", "", "")
                    res = PT_CheckBoxPart(asm, have(keyS), "TIM", keyS, nameS, tRef(i), 0,
                                          tBox(i, 0), tBox(i, 1), z0, tBox(i, 2), tBox(i, 3), tTim(i),
                                          topN, botN, tolPos, tolSize, MACRO_NAME, True)
                    If res < 0 Then failMsgs.Add("Could not update " & nameS & ": " & PT_LastErr)
                    If res > 0 Then nTimChanged = nTimChanged + 1
                Else
                    fpath = outDir & prefix & "TIM_" & PT_SafeFile(tRef(i)) & ".sldprt"
                    If PT_AddBoxPart(asm, partTmpl, fpath, "TIM", keyS, nameS, tRef(i), 0,
                                     tBox(i, 0), tBox(i, 1), z0, tBox(i, 2), tBox(i, 3), tTim(i),
                                     topN, botN, MACRO_NAME) Then
                        nTimAdded = nTimAdded + 1
                        If doLog Then PT_LogAdd(MACRO_NAME, keyS, nameS, tRef(i), "added", "Box", "",
                                                PT_BoxText(tBox(i, 0), tBox(i, 1), z0, tBox(i, 2), tBox(i, 3), tTim(i)))
                    Else
                        failMsgs.Add("Could not add " & nameS & ": " & PT_LastErr)
                    End If
                End If
            End If
        Next
        nTimSupp = PT_SuppressUnwanted(have, wasSupp, want, MACRO_NAME, True)
    End Sub

    ' ================================================================= helpers
    Private Function CountPlanes(ByVal md As Object) As Integer
        Return PT_CountFeatures(md, "RefPlane")
    End Function

    Private Function NextRev(ByVal outDir As String, ByVal prefix As String) As Integer
        Dim cnt As Integer = 0, stem As String, numS As String, f As String
        stem = LCase(prefix & "Spreader_r")
        If System.IO.Directory.Exists(outDir) Then
            For Each full As String In System.IO.Directory.GetFiles(outDir, prefix & "Spreader_r*.sldprt")
                f = System.IO.Path.GetFileName(full)
                numS = Mid(f, Len(stem) + 1)
                numS = Left(numS, InStr(numS & ".", ".") - 1)
                If IsNumeric(numS) Then
                    If CInt(numS) > cnt Then cnt = CInt(numS)
                End If
            Next
        End If
        Return cnt + 1
    End Function

    ' "x|y|L|W|h" in micrometres -> readable mm text
    Private Function ParamText(ByVal p As String) As String
        Dim a As String(), s As String = ""
        If p = "" Then Return ""
        a = Split(p, "|")
        For i As Integer = 0 To UBound(a)
            If IsNumeric(a(i)) Then s = s & IIf(i = 0, "", IIf(i = 2, " size ", IIf(i = 4, " h ", ", "))) & CDbl(a(i)) / 1000.0
        Next
        Return "at " & s
    End Function

    ' same size (L, W and height), only the position differs
    Private Function SameSize(ByVal p1 As String, ByVal p2 As String) As Boolean
        Dim a As String(), b As String()
        If p1 = "" Then Return False
        a = Split(p1, "|") : b = Split(p2, "|")
        If UBound(a) <> UBound(b) Then Return False
        For i As Integer = 2 To UBound(a)
            If a(i) <> b(i) Then Return False
        Next
        Return True
    End Function

    Private Sub WriteReport(ByVal rptPath As String, ByVal asmPath As String)
        Try
            Using ts As System.IO.StreamWriter = NewTextFile(rptPath)
                ts.WriteLine("BuildSpreader report " & NowText())
                ts.WriteLine("Assembly: " & asmPath)
                ts.WriteLine("Cuts and pedestals: added " & nFeatAdded & ", rebuilt " & nFeatChanged & ", suppressed " & nFeatSupp)
                ts.WriteLine("TIM parts: added " & nTimAdded & ", changed " & nTimChanged & ", suppressed " & nTimSupp)
                ts.WriteLine("")
                ts.Write(rptBody)
                ts.WriteLine("PROBLEMS (" & failMsgs.Count & ")")
                For Each it As Object In failMsgs
                    ts.WriteLine("  " & CStr(it))
                Next
                ts.WriteLine("")
                ts.WriteLine("COVERED CHIPS SKIPPED (" & skipped.Count & ")")
                For Each it As Object In skipped
                    ts.WriteLine("  " & CStr(it))
                Next
                ts.WriteLine("")
                ts.WriteLine("NOTES")
                For Each it As Object In notes
                    ts.WriteLine("  " & CStr(it))
                Next
            End Using
        Catch
        End Try
    End Sub

    ' ================================================================= user-sketched spreader
    ' The spreader part in the assembly that has "spreader" in its name and is not suppressed.
    ' A component selected in the assembly wins. nFound = how many matched.
    Private Function FindSpreaderComp(ByVal asm As Object, ByRef nFound As Integer, ByVal wantBottom As Boolean) As Object
        Dim comps As Object, c As Object, found As Object = Nothing, sel As Object, nm As String
        nFound = 0
        Try
            sel = asm.SelectionManager
            If CInt(sel.GetSelectedObjectCount2(-1)) = 1 Then
                c = sel.GetSelectedObjectsComponent4(1, -1)
                If c IsNot Nothing Then
                    nm = LCase(CStr(c.Name2))
                    If InStr(nm, "spreader") > 0 And (InStr(nm, "bottom") > 0) = wantBottom Then
                        nFound = 1
                        Return c
                    End If
                End If
            End If
        Catch
        End Try
        comps = asm.GetComponents(True)
        If IsEmpty(comps) Then Return Nothing
        For i As Integer = 0 To UBound(comps)
            c = comps(i)
            If c.GetSuppression2 <> PT_swSuppressed Then
                nm = LCase(CStr(c.Name2))
                If InStr(nm, "spreader") > 0 And (InStr(nm, "bottom") > 0) = wantBottom Then
                    nFound = nFound + 1
                    found = c
                End If
            End If
        Next
        Return found
    End Function

    ' Position of the spreader part in the assembly (mm). False if it is turned (rotation is not the identity).
    Private Function ReadOffset(ByVal c As Object) As Boolean
        Dim d As Object = c.Transform2.ArrayData
        For k As Integer = 0 To 8
            If Math.Abs(d(k) - IIf(k Mod 4 = 0, 1.0, 0.0)) > 0.000001 Then Return False
        Next
        pOx = d(9) * 1000.0 : pOy = d(10) * 1000.0 : pOz = d(11) * 1000.0
        Return True
    End Function

    ' The plate outline is the sketch named Outline. One unnamed sketch is taken as the outline.
    Private Sub PrepareOutline(ByVal md As Object)
        Dim f As Object, cnt As Integer = 0, lastS As Object = Nothing
        If PT_FindFeature(md, "Outline") IsNot Nothing Then
            CheckOutlinePlane(md)
            Exit Sub
        End If
        f = md.FirstFeature
        Do While f IsNot Nothing
            If CStr(f.GetTypeName2) = "ProfileFeature" Then
                cnt = cnt + 1
                lastS = f
            End If
            f = f.GetNextFeature
        Loop
        If cnt = 1 Then
            lastS.Name = "Outline"
            notes.Add("The only sketch in the spreader part was renamed Outline.")
            CheckOutlinePlane(md)
            Exit Sub
        End If
        If cnt = 0 Then Err.Raise(vbObjectError + 36, , "the spreader part has no sketch. Sketch the plate outline on the Front Plane, exit the sketch and name it Outline.")
        Err.Raise(vbObjectError + 36, , "the spreader part has " & cnt & " sketches. Rename the one with the plate outline to Outline.")
    End Sub

    Private Function SolidCount(ByVal md As Object) As Integer
        Dim b As Object = md.GetBodies2(PT_swSolidBody, True)
        If IsEmpty(b) Then Return 0
        Return UBound(b) + 1
    End Function

    ' Plate extrude: starts startMm from the sketch plane, depthMm thick.
    Private Sub PlateDef(ByVal md As Object, ByVal feat As Object, ByVal startMm As Double, ByVal depthMm As Double,
                         ByVal revDir As Boolean, ByVal revOff As Boolean)
        Dim def As Object = feat.GetDefinition
        def.AccessSelections(md, NullDisp)
        def.FromType = 3                                       ' swStartOffset
        def.FromOffsetDistance = startMm / 1000.0
        def.FromOffsetReverse = revOff
        def.ReverseDirection = revDir
        def.SetDepth(True, depthMm / 1000.0)
        feat.ModifyDefinition(def, md, NullDisp)
    End Sub

    Private Function PlateStartsByOffset(ByVal md As Object, ByVal feat As Object) As Boolean
        Dim def As Object = feat.GetDefinition                 ' read only: AccessSelections would leave the part rolled back
        Return CInt(def.FromType) = 3
    End Function

    ' The board lies in the Front Plane of the thermal assembly (board X, Y in the plane, Z up), whatever way the STEP
    ' had it. The Outline must be on a plane parallel to that, not on the Top or Right plane.
    Private Sub CheckOutlinePlane(ByVal md As Object)
        Dim sk As Object, d As Object
        sk = PT_FindFeature(md, "Outline")
        If sk Is Nothing Then Exit Sub
        d = sk.GetSpecificFeature2.ModelToSketchTransform.ArrayData
        If Math.Abs(CDbl(d(8))) < 0.9999 Then
            Err.Raise(vbObjectError + 36, , "the Outline sketch is not parallel to the board. In the thermal assembly the board lies in the " &
                      "Front Plane (Z is up through the board), not the Top Plane. Draw the Outline on the Front Plane or on a face parallel to it.")
        End If
    End Sub

    ' Underside plane at zuP (board Z of the plate face toward the board, in the part frame), then the plate:
    ' the Outline sketch (any shape) extruded tPl away from the board, starting zuP from its plane.
    Private Sub MakeUserPlate(ByVal md As Object, ByVal zuP As Double, ByVal tPl As Double)
        Dim front As Object, pl As Object, sk As Object, feat As Object, nBefore As Integer
        sk = PT_FindFeature(md, "Outline")
        front = PT_RefPlane(md, 1)
        If front Is Nothing Then Err.Raise(vbObjectError + 40, , "the spreader part has no reference plane")
        nBefore = CountPlanes(md)
        md.ClearSelection2(True)
        front.Select2(False, 0)
        md.CreatePlaneAtOffset3(Math.Abs(zuP) / 1000.0, (zuP < 0), True)
        md.ClearSelection2(True)
        If CountPlanes(md) <= nBefore Then Err.Raise(vbObjectError + 41, , "the underside plane was not made")
        pl = PT_LastFeatureOfType(md, "RefPlane")
        pl.Name = "PT_Under"
        If Math.Abs(PlaneZ(pl) - zuP) > 0.001 Then SetPlaneZ(md, pl, zuP)
        feat = PT_Boss(md, sk, tPl, False)
        feat.Name = "PT_Plate"
        FitPlate(md, feat, zuP, tPl)
    End Sub

    ' Puts the plate at its Z range: top plate zuP to zuP + tPl, bottom plate zuP - tPl to zuP. Tries the direction flips.
    Private Sub FitPlate(ByVal md As Object, ByVal feat As Object, ByVal zuP As Double, ByVal tPl As Double)
        Dim zr As Double() = Nothing, ok As Boolean = False, zlo As Double, zhi As Double
        zlo = IIf(bottomSide, zuP - tPl, zuP)
        zhi = IIf(bottomSide, zuP, zuP + tPl)
        For k As Integer = 0 To 3
            PlateDef(md, feat, Math.Abs(zuP), tPl, ((k And 1) <> 0), ((k And 2) <> 0))
            md.EditRebuild3()
            zr = FeatZRange(feat)
            If zr IsNot Nothing Then
                If Math.Abs(zr(0) - zlo) < 0.001 And Math.Abs(zr(1) - zhi) < 0.001 Then
                    ok = True
                    Exit For
                End If
            End If
        Next
        If ok Then Exit Sub
        If zr Is Nothing Then Err.Raise(vbObjectError + 42, , "the Outline sketch made no solid. Is it a closed outline?")
        Err.Raise(vbObjectError + 43, , "the plate ended up at Z " & Round(zr(0), 3) & " to " & Round(zr(1), 3) &
                  " mm in the part, expected " & Round(zlo, 3) & " to " & Round(zhi, 3) & ". Was the Outline drawn on the Front Plane?")
    End Sub

    ' Z (mm, part frame) of a reference plane
    Private Function PlaneZ(ByVal pl As Object) As Double
        Dim d As Object = pl.GetSpecificFeature2.Transform.ArrayData
        Return CDbl(d(11)) * 1000.0
    End Function

    ' Moves the underside plane to Z zmm (either sign): sets the distance, turns the direction if it went the wrong way.
    Private Sub SetPlaneZ(ByVal md As Object, ByVal pl As Object, ByVal zmm As Double)
        PT_SetPlaneOffset(md, pl, Math.Abs(zmm), False)
        md.EditRebuild3()
        If Math.Abs(PlaneZ(pl) - zmm) <= 0.001 Then Exit Sub
        PT_SetPlaneOffset(md, pl, Math.Abs(zmm), True)
        md.EditRebuild3()
        If Math.Abs(PlaneZ(pl) - zmm) > 0.001 Then Err.Raise(vbObjectError + 44, , "could not put the underside plane at Z " & Round(zmm, 3) & " mm (it is at " & Round(PlaneZ(pl), 3) & ")")
    End Sub

    ' Lowest and highest Z (mm) over the faces a feature made: {zmin, zmax}, or Nothing
    Private Function FeatZRange(ByVal feat As Object) As Double()
        Dim faces As Object = feat.GetFaces, bb As Object, zmin As Double, zmax As Double
        If Not IsArray(faces) Then Return Nothing
        zmin = 1.0E+30 : zmax = -1.0E+30
        For i As Integer = 0 To UBound(faces)
            bb = faces(i).GetBox
            If bb(2) * 1000.0 < zmin Then zmin = bb(2) * 1000.0
            If bb(5) * 1000.0 > zmax Then zmax = bb(5) * 1000.0
        Next
        Return New Double() {zmin, zmax}
    End Function

    ' The cut and pedestal boxes are in the part frame. If the part has moved in the assembly since the last run, they are
    ' stale: forget their stored sizes so SyncOne rebuilds them in the right place.
    Private Sub CheckOrigin(ByVal md As Object)
        Dim cur As String, oldO As String, f As Object, nm As String, cnt As Integer = 0
        cur = PT_Um(pOx) & "|" & PT_Um(pOy) & "|" & PT_Um(pOz)
        oldO = PT_GetProp(md, "PT_Origin")
        If oldO <> "" And oldO <> cur Then
            f = md.FirstFeature
            Do While f IsNot Nothing
                nm = CStr(f.Name)
                If (Left(nm, 4) = "Ped_" Or Left(nm, 6) = "Clear_") And Right(nm, 2) <> "_S" Then
                    PT_SetProp(md, "PT_" & nm, "moved")
                    cnt = cnt + 1
                End If
                f = f.GetNextFeature
            Loop
            If cnt > 0 Then notes.Add("The spreader part has moved in the assembly: " & cnt & " pedestals and cuts are rebuilt in their new place.")
        End If
        PT_SetProp(md, "PT_Origin", cur)
    End Sub

End Module
