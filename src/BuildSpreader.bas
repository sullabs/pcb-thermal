Attribute VB_Name = "BuildSpreader"
Option Explicit
' =====================================================================================
' BuildSpreader  -  SolidWorks 2023 macro (milestone M4)
'
' Completes the heat spreader part(s) and builds the TIM parts in the thermal assembly,
' from the Spreader tab (top plate) and the Spreader bottom tab (plate under the PCB).
' Works for any board. Run BuildBoard first.
'
' YOU draw the plate: in the thermal assembly insert a new part with "Spreader" in its name
' ("Spreader" and "Bottom" for the plate under the PCB), on the Front Plane (or a face parallel
' to it), sketch any closed outline, name the sketch Outline, save the part as a file.
' The macro then adds to that part, in board coordinates:
'   PT_Under        plane at the plate face toward the board (Spr_ZUnder of that tab)
'   PT_Plate        your Outline extruded by the plate thickness, away from the board
'   Clear_<Key>     one cut per tall part (tab section C), through all
'   Ped_<RefDes>    one pedestal per covered chip with pedestal > 0, toward the chip
' Named faces: Top (outer face), Underside (face toward the board), Contact_<RefDes> (pedestal ends).
' TIM parts (<prefix>TIM_<RefDes>.sldprt): box between chip and plate, TIM mode Body only.
' With no top spreader part it offers a rectangle the size of the PCB (<prefix>Spreader_rNN.sldprt).
'
' Update: run again after changing the tabs or the outline. The plane height and plate
' thickness are changed in place; a cut or pedestal whose size or position changed is
' rebuilt; new ones are added at the end of the tree; ones no longer wanted are
' suppressed, never deleted. Features you added yourself (holes, fillets...) and the
' Outline sketch are never touched. Every change goes to the Change log tab.
'
' Refuses to run while the Change log has unchecked rows.
' Needs module PTCommon in the same macro, and a copy of PCB_thermal_TEMPLATE.xlsx.
' =====================================================================================

Const MACRO_NAME As String = "BuildSpreader"
Const swDefaultTemplatePart As Long = 8
Const swUnSuppressFeature As Long = 1
Const swSuppressFeature As Long = 0
Const swAllConfiguration As Long = 2

' Spreader tab chip table offsets (Spr_Chips, column B = 0)
Const C_REF As Long = 0, C_KEY As Long = 1, C_X As Long = 2, C_Y As Long = 3, C_L As Long = 4, C_W As Long = 5
Const C_TOP As Long = 6, C_COV As Long = 7, C_TIMMODE As Long = 9, C_TIMT As Long = 13, C_PED As Long = 14, C_CHECK As Long = 16
' clearance cut table offsets (Spr_Cuts, column B = 0)
Const K_KEY As Long = 1, K_NAME As Long = 2, K_X As Long = 3, K_Y As Long = 4, K_L As Long = 5, K_W As Long = 6

' wanted chips (of the side being built; for the bottom side chBox(,4) is the chip bottom face)
Dim chRef() As String, chKey() As String, chBox() As Double, chPed() As Double, chTim() As Double, chBody() As Boolean, nCh As Long
' wanted cuts
Dim cuKey() As String, cuName() As String, cuBox() As Double, nCu As Long
' TIM parts wanted, both sides
Dim tRef(0 To 199) As String, tBox(0 To 199, 0 To 4) As Double, tTim(0 To 199) As Double, tBot(0 To 199) As Boolean, nT As Long
Dim skipped As Collection, failMsgs As Collection, notes As Collection
Dim doLog As Boolean
Dim bottomSide As Boolean                                ' the side being built
Dim wsSide As Object                                     ' Spreader bottom tab (when bottomSide)
Dim rptBody As String, resMsg As String
Dim pOx As Double, pOy As Double, pOz As Double          ' spreader part origin in board coordinates, mm (0 when it sits at the assembly origin)
Dim nFeatAdded As Long, nFeatChanged As Long, nFeatSupp As Long, nTimAdded As Long, nTimChanged As Long, nTimSupp As Long

Sub main()
    Set swApp = Application.SldWorks
    PT_Trace MACRO_NAME & ": started (VBA macro)."
    PT_LogReset
    Set skipped = New Collection: Set failMsgs = New Collection: Set notes = New Collection
    nFeatAdded = 0: nFeatChanged = 0: nFeatSupp = 0: nTimAdded = 0: nTimChanged = 0: nTimSupp = 0
    nT = 0: rptBody = "": resMsg = ""

    ' ---------- workbook ----------
    Dim xl As Object, wb As Object
    Set xl = PT_GetExcel()
    Set wb = PT_PickWorkbook(xl, MACRO_NAME, "Spreader|CAD geometry|Settings|Change log")
    If wb Is Nothing Then Exit Sub
    If Not PT_NeedNames(wb, "Spr_Chips|Spr_Cuts|Spr_ZUnder|Spr_Thickness|Geo_PCB_Used|Log_Header|Log_Unchecked|" & _
                            "Set_Prefix|Set_TolPos|Set_TolSize|Set_PartTemplate|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
    If PT_RefuseIfUnchecked(wb, MACRO_NAME) Then Exit Sub

    Dim pcbL As Double, pcbW As Double, rp As Object, tolPos As Double, tolSize As Double
    xl.Calculate
    Set rp = PT_Rng(wb, "Geo_PCB_Used")
    pcbL = PT_Num(rp.Cells(1, 1).Value, 0)
    pcbW = PT_Num(rp.Cells(2, 1).Value, 0)
    tolPos = PT_NameNum(wb, "Set_TolPos", 0.05)
    tolSize = PT_NameNum(wb, "Set_TolSize", 0.05)

    ' ---------- assembly and files ----------
    Dim asm As Object, asmPath As String, partTmpl As String, outDir As String, prefix As String
    outDir = PT_OutputFolder(wb)
    prefix = PT_SafeFile(PT_Prefix(wb))
    partTmpl = PT_Template(wb, "Set_PartTemplate", swDefaultTemplatePart)
    If partTmpl = "" Then
        PT_Msg "BuildSpreader did not run. The part template was not found. Set its full path on the Settings tab.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    asmPath = PT_AsmPath(wb)
    On Error GoTo swFail
    Set asm = PT_OpenAsm(asmPath)
    If asm Is Nothing Then
        On Error GoTo 0
        PT_Msg "BuildSpreader did not run. The thermal assembly is not there yet:" & vbCrLf & asmPath & vbCrLf & _
               "Run BuildBoard first.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    PT_ActivateDoc asm
    PT_LeaveHeatOnly asm
    asm.ResolveAllLightWeightComponents False

    ' ---------- the two sides ----------
    Dim pass As Long, res As Long, nBuilt As Long, errs As Long, warns As Long
    For pass = 0 To 1
        bottomSide = (pass = 1)
        Set wsSide = Nothing
        If bottomSide Then
            Set wsSide = PT_Sheet(wb, "Spreader bottom")
            If wsSide Is Nothing Then Exit For
        End If
        res = DoSide(wb, asm, partTmpl, outDir, prefix, tolPos, tolSize, pcbL, pcbW)
        If res < 0 Then Exit Sub
        nBuilt = nBuilt + res
    Next
    If nBuilt = 0 Then
        PT_Msg "BuildSpreader did nothing: no spreader part and no covered chips." & vbCrLf & _
               "Mark chips Covered = Yes on the Spreader tab (top) or the Spreader bottom tab, and draw the plate outline.", vbExclamation, MACRO_NAME
        Exit Sub
    End If

    ' ---------- TIM parts (both sides) ----------
    SyncTims asm, partTmpl, outDir, prefix, tolPos, tolSize

    PT_ActivateDoc asm
    asm.ForceRebuild3 False
    asm.Save3 PT_swSaveSilent + PT_swSaveReferenced, errs, warns
    On Error GoTo 0

    ' ---------- report, Change log ----------
    Dim rptPath As String, nLog As Long
    rptPath = outDir & prefix & "BuildSpreader_report.txt"
    WriteReport rptPath, asmPath
    On Error GoTo xlFail
    xl.ScreenUpdating = False
    nLog = PT_LogFlush(wb)
    xl.ScreenUpdating = True
    xl.Calculate
    wb.Save
    On Error GoTo 0

    Dim s As String
    s = "Written and saved to the spreader part(s) and the thermal assembly:" & vbCrLf & asmPath & vbCrLf & vbCrLf & resMsg & _
        "Cuts and pedestals: " & nFeatAdded & " added, " & nFeatChanged & " rebuilt, " & nFeatSupp & " suppressed." & vbCrLf & _
        "TIM parts: " & nTimAdded & " added, " & nTimChanged & " changed, " & nTimSupp & " suppressed." & vbCrLf
    If nLog > 0 Then s = s & nLog & " rows added to the Change log tab. Check them there." & vbCrLf
    If skipped.Count > 0 Then s = s & skipped.Count & " covered chips skipped (see the Check column on the Spreader tabs)." & vbCrLf
    If notes.Count > 0 Then s = s & notes.Count & " notes (listed in the report)." & vbCrLf
    If failMsgs.Count > 0 Then s = s & failMsgs.Count & " problems (listed in the report)." & vbCrLf
    s = s & vbCrLf & "Report: " & rptPath
    PT_Msg s, IIf(failMsgs.Count > 0, vbExclamation, vbInformation), MACRO_NAME
    Exit Sub

swFail:
    If Err.Number >= vbObjectError + 34 And Err.Number <= vbObjectError + 36 Then
        PT_Msg "BuildSpreader did not run (" & IIf(bottomSide, "bottom", "top") & " spreader): " & Err.Description, vbExclamation, MACRO_NAME
        Exit Sub
    End If
    PT_Msg "BuildSpreader stopped in SolidWorks (" & IIf(bottomSide, "bottom", "top") & " spreader): " & Err.Description & vbCrLf & _
           "Nothing was saved. Close the spreader part and the assembly without saving, send this message, and run again.", _
           vbExclamation, MACRO_NAME
    Exit Sub
xlFail:
    xl.ScreenUpdating = True
    PT_Msg "The spreader was saved, but writing the Change log failed: " & Err.Description & vbCrLf & _
           "Close any cell you are editing in Excel. The changes are listed in " & rptPath, vbExclamation, MACRO_NAME
End Sub

' Builds or updates one plate (top, or bottom when bottomSide). Returns 1 built, 0 nothing to do, -1 stop the macro.
Function DoSide(ByVal wb As Object, ByVal asm As Object, ByVal partTmpl As String, ByVal outDir As String, ByVal prefix As String, _
                ByVal tolPos As Double, ByVal tolSize As Double, ByVal pcbL As Double, ByVal pcbW As Double) As Long
    Dim zu As Double, tPl As Double, zuP As Double, sideTxt As String, i As Long
    Dim sprComp As Object, md As Object, ans As VbMsgBoxResult, newRev As Boolean
    Dim sprPath As String, sprName As String, sprMsg As String, nFound As Long, errs As Long, warns As Long
    sideTxt = IIf(bottomSide, "Bottom", "Top")
    zu = SNum(wb, "Spr_ZUnder", 0)
    tPl = SNum(wb, "Spr_Thickness", 0)
    ReadChips wb
    ReadCuts wb
    Set sprComp = FindSpreaderComp(asm, nFound, bottomSide)
    If nFound > 1 Then
        PT_Msg "BuildSpreader did not run. The assembly has " & nFound & " " & LCase(sideTxt) & " spreader parts that are not suppressed." & vbCrLf & _
               "Select the one to complete in the assembly (click it in the tree) and run again.", vbExclamation, MACRO_NAME
        DoSide = -1
        Exit Function
    End If
    If sprComp Is Nothing And nCh = 0 Then Exit Function                 ' nothing for this side
    If bottomSide And nCh = 0 And zu = 0 Then                           ' a bottom part exists but no chip is covered: leave it as it is
        notes.Add "Bottom: no chip is covered on the Spreader bottom tab, so the bottom spreader part was left unchanged."
        Exit Function
    End If
    If tPl <= 0 Or (Not bottomSide And zu <= 0) Or (bottomSide And nCh = 0 And zu = 0) Then
        PT_Msg "BuildSpreader did not run (" & LCase(sideTxt) & " spreader). On its tab the plate thickness (" & tPl & ") must be above 0 and " & _
               "the face height (" & zu & ") must be set." & vbCrLf & _
               "Mark at least one chip Covered = Yes, or set the face height to Manual.", vbExclamation, MACRO_NAME
        DoSide = -1
        Exit Function
    End If
    pOx = 0: pOy = 0: pOz = 0
    If sprComp Is Nothing Then
        If bottomSide Then
            notes.Add "Bottom: " & nCh & " chips are covered but there is no spreader part with Spreader and Bottom in its name. Draw its outline and run again."
            Exit Function
        End If
        ans = PT_Msg("The thermal assembly has no spreader part yet." & vbCrLf & vbCrLf & _
                     "To draw your own: in the assembly choose Insert > Component > New Part, put 'Spreader' in its name, " & _
                     "pick the Front Plane (or a face parallel to it), sketch the outline, exit the sketch, name the sketch Outline, " & _
                     "save the part as a file, then run this macro again." & vbCrLf & vbCrLf & _
                     "Yes = make a rectangle the size of the PCB now (its Outline sketch can be edited)" & vbCrLf & _
                     "No = stop", vbYesNo + vbQuestion, MACRO_NAME)
        If ans <> vbYes Then DoSide = -1: Exit Function
        If pcbL <= 0 Or pcbW <= 0 Then
            PT_Msg "BuildSpreader did not run. CAD geometry has no PCB size.", vbExclamation, MACRO_NAME
            DoSide = -1
            Exit Function
        End If
        newRev = True
    End If

    If newRev Then
        doLog = False
        zuP = zu
        sprPath = outDir & prefix & "Spreader_r" & Format(NextRev(outDir, prefix), "00") & ".sldprt"
        Set md = swApp.NewDocument(partTmpl, 0, 0, 0)
        If md Is Nothing Then Err.Raise vbObjectError + 30, , "SolidWorks could not make a new part from " & partTmpl
        MakeBase md, zu, tPl, pcbL, pcbW
        sprName = Mid(sprPath, InStrRev(sprPath, "\") + 1)
        PT_SetProp md, "Role", "Spreader"
        PT_SetProp md, "Key", LCase(Left(sprName, Len(sprName) - 7))
        PT_SetProp md, "Name", "Spreader " & sprName
        PT_SetProp md, "BuiltBy", MACRO_NAME
        If Not md.Extension.SaveAs(sprPath, 0, PT_swSaveSilent, Nothing, errs, warns) Then
            Err.Raise vbObjectError + 31, , "could not save " & sprPath & " (error " & errs & ")"
        End If
        sprMsg = "New rectangular spreader made from the PCB size."
    Else
        If sprComp.IsVirtual Then Err.Raise vbObjectError + 34, , "the spreader is a virtual part inside the assembly. " & _
                                  "Right-click it in the tree, choose Save Part (in External File), and run again."
        If Not ReadOffset(sprComp) Then Err.Raise vbObjectError + 35, , "the spreader part is turned relative to the assembly. " & _
                                  "Make it again on the Front Plane (or a face parallel to it) so its X, Y and Z match the board."
        sprComp.SetSuppression2 PT_swFullyResolved          ' a suppressed or lightweight part has no bodies to edit or name
        Set md = sprComp.GetModelDoc2
        If md Is Nothing Then Err.Raise vbObjectError + 37, , "the spreader part could not be opened"
        sprPath = md.GetPathName
        If sprPath = "" Then Err.Raise vbObjectError + 34, , "the spreader part has no file yet. Save it first."
        sprName = Mid(sprPath, InStrRev(sprPath, "\") + 1)
        zuP = zu - pOz
        doLog = Not (PT_FindFeature(md, "PT_Plate") Is Nothing)
        PT_ActivateDoc md
        md.FeatureManager.EditRollback 1, ""                ' swMoveRollbackBarToEnd: features below a rollback bar are not built
        If doLog And IsEmpty(md.GetBodies2(PT_swSolidBody, True)) Then   ' a part loaded only as a component has no bodies to read: open its own window
            Dim md2 As Object
            Set md2 = swApp.OpenDoc6(sprPath, PT_swDocPART, PT_swOpenSilent, "", errs, warns)
            If Not md2 Is Nothing Then Set md = md2: PT_ActivateDoc md
        End If
        If Not doLog Then
            PrepareOutline md
            MakeUserPlate md, zuP, tPl
            If PT_GetProp(md, "Role") = "" Then PT_SetProp md, "Role", "Spreader"
            If PT_GetProp(md, "Key") = "" Then PT_SetProp md, "Key", LCase(Left(sprName, Len(sprName) - 7))
            If PT_GetProp(md, "Name") = "" Then PT_SetProp md, "Name", "Spreader " & sprName
            PT_SetProp md, "Side", sideTxt
            PT_SetProp md, "BuiltBy", MACRO_NAME
            sprMsg = sideTxt & " spreader completed from your Outline sketch."
        Else
            SyncPlate md, zuP, tPl, tolPos, tolSize
            sprMsg = sideTxt & " spreader updated."
        End If
    End If

    CheckOrigin md
    SyncCuts md, zuP, tolPos
    SyncPedestals md, zuP, tolPos, tolSize
    md.ForceRebuild3 False
    NameSpreaderFaces md, zuP, tPl
    md.Save3 PT_swSaveSilent, errs, warns

    ' ---------- spreader in the assembly ----------
    PT_ActivateDoc asm
    If newRev Then
        Dim c As Object, ttl As String
        ttl = md.GetTitle
        Set c = asm.AddComponent5(sprPath, 0, "", False, "", 0, 0, 0)
        If c Is Nothing Then Err.Raise vbObjectError + 33, , "could not insert " & sprPath & " into the assembly"
        PT_PlaceAndFix asm, c, 0, 0, 0
        swApp.CloseDoc ttl
        PT_ActivateDoc asm
    Else
        If Not sprComp.IsFixed Then PT_FixComp asm, sprComp
    End If

    ' ---------- remember the TIM parts and the report lines of this side ----------
    For i = 0 To nCh - 1
        If chBody(i) And chTim(i) > 0.0005 And nT <= UBound(tRef) Then
            tRef(nT) = chRef(i)
            tBox(nT, 0) = chBox(i, 0): tBox(nT, 1) = chBox(i, 1): tBox(nT, 2) = chBox(i, 2): tBox(nT, 3) = chBox(i, 3): tBox(nT, 4) = chBox(i, 4)
            tTim(nT) = chTim(i): tBot(nT) = bottomSide
            nT = nT + 1
        End If
    Next
    rptBody = rptBody & UCase(sideTxt) & " SPREADER: " & sprPath & vbCrLf & _
              "  plate face toward the board at Z " & zu & " mm, plate " & tPl & " mm, part origin at (" & Round(pOx, 3) & ", " & Round(pOy, 3) & ", " & Round(pOz, 3) & ") mm" & vbCrLf & _
              "  chips (" & nCh & "):" & vbCrLf
    For i = 0 To nCh - 1
        rptBody = rptBody & "    " & chRef(i) & ": face Z " & chBox(i, 4) & ", TIM " & chTim(i) & IIf(chBody(i), " (body)", " (contact resistance)") & _
                  ", pedestal " & chPed(i) & " mm" & vbCrLf
    Next
    rptBody = rptBody & "  clearance cuts (" & nCu & "):" & vbCrLf
    For i = 0 To nCu - 1: rptBody = rptBody & "    " & cuName(i) & vbCrLf: Next
    rptBody = rptBody & vbCrLf
    resMsg = resMsg & sprMsg & " (" & sprName & ") Face at Z " & zu & " mm, plate " & tPl & " mm." & vbCrLf
    DoSide = 1
End Function

' Named range of the side being built: workbook name for the top tab, tab-level name for the bottom tab.
Function SRng(ByVal wb As Object, ByVal nameS As String) As Object
    On Error Resume Next
    If bottomSide Then Set SRng = wsSide.Names(nameS).RefersToRange Else Set SRng = wb.Names(nameS).RefersToRange
End Function

Function SNum(ByVal wb As Object, ByVal nameS As String, ByVal dflt As Double) As Double
    Dim r As Object
    Set r = SRng(wb, nameS)
    If r Is Nothing Then SNum = dflt Else SNum = PT_Num(r.Cells(1, 1).Value, dflt)
End Function

' ================================================================= Excel side
Sub ReadChips(ByVal wb As Object)
    Dim v As Variant, r As Long, k As Long, refD As String, seen As Object, ck As String
    Set seen = CreateObject("Scripting.Dictionary")
    v = SRng(wb, "Spr_Chips").Value
    nCh = 0
    ReDim chRef(UBound(v, 1)): ReDim chKey(UBound(v, 1)): ReDim chBox(UBound(v, 1), 4)
    ReDim chPed(UBound(v, 1)): ReDim chTim(UBound(v, 1)): ReDim chBody(UBound(v, 1))
    For r = 1 To UBound(v, 1)
        refD = PT_Str(v(r, 1 + C_REF))
        If refD <> "" And UCase(PT_Str(v(r, 1 + C_COV))) = "YES" Then
            ck = PT_Str(v(r, 1 + C_CHECK))
            If ck <> "OK" Then
                skipped.Add refD & ": " & ck
            ElseIf seen.Exists(UCase(refD)) Then
                skipped.Add refD & ": listed twice, second row ignored"
            Else
                seen.Add UCase(refD), 1
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

Sub ReadCuts(ByVal wb As Object)
    Dim v As Variant, r As Long, keyS As String, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    v = SRng(wb, "Spr_Cuts").Value
    nCu = 0
    ReDim cuKey(UBound(v, 1)): ReDim cuName(UBound(v, 1)): ReDim cuBox(UBound(v, 1), 3)
    For r = 1 To UBound(v, 1)
        keyS = PT_Str(v(r, 1 + K_KEY))
        If keyS <> "" And PT_IsNum(v(r, 1 + K_X)) And PT_IsNum(v(r, 1 + K_L)) Then
            If Not seen.Exists(keyS) Then
                seen.Add keyS, 1
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
Sub MakeBase(ByVal md As Object, ByVal zu As Double, ByVal tPl As Double, ByVal pcbL As Double, ByVal pcbW As Double)
    Dim front As Object, pl As Object, sk As Object, feat As Object, zr As Variant
    Set front = PT_RefPlane(md, 1)
    If front Is Nothing Then Err.Raise vbObjectError + 40, , "the part template has no reference plane"
    Dim nBefore As Long
    nBefore = CountPlanes(md)
    md.ClearSelection2 True
    front.Select2 False, 0
    md.CreatePlaneAtOffset3 zu / 1000#, False, True
    md.ClearSelection2 True
    If CountPlanes(md) <= nBefore Then Err.Raise vbObjectError + 41, , "the underside plane was not made"
    Set pl = PT_LastFeatureOfType(md, "RefPlane")
    pl.Name = "PT_Under"
    Set sk = PT_SketchRect(md, pl, 0, 0, pcbL, pcbW, zu)
    sk.Name = "Outline"
    Set feat = PT_Boss(md, sk, tPl, False)
    feat.Name = "PT_Plate"

    zr = PT_BodyZRange(md)
    If IsEmpty(zr) Then Err.Raise vbObjectError + 42, , "the plate has no body"
    If (zr(0) + zr(1)) / 2 < 0 Then                      ' plane went below the board: turn it over
        PT_SetPlaneOffset md, pl, zu, True
        md.EditRebuild3
        zr = PT_BodyZRange(md)
    End If
    If Abs(zr(0) - zu) > 0.001 Then                      ' plate went down from the plane: turn it up
        PT_FlipExtrude md, feat
        md.EditRebuild3
        zr = PT_BodyZRange(md)
    End If
    If Abs(zr(0) - zu) > 0.001 Or Abs(zr(1) - (zu + tPl)) > 0.001 Then
        Err.Raise vbObjectError + 43, , "the plate ended up at Z " & Round(zr(0), 3) & " to " & Round(zr(1), 3) & _
                  " mm, expected " & zu & " to " & (zu + tPl)
    End If
End Sub

Sub SyncPlate(ByVal md As Object, ByVal zu As Double, ByVal tPl As Double, ByVal tolPos As Double, ByVal tolSize As Double)
    Dim pl As Object, feat As Object, oldZ As Double, oldT As Double, byOffset As Boolean, refit As Boolean, sideTxt As String
    sideTxt = IIf(bottomSide, "Bottom spreader", "Spreader")
    Set pl = PT_FindFeature(md, "PT_Under")
    Set feat = PT_FindFeature(md, "PT_Plate")
    oldT = PT_ExtrudeDepth(feat)
    byOffset = PlateStartsByOffset(md, feat)             ' plate made from the user's Outline (start offset) or the PCB rectangle
    If byOffset Then oldZ = PlaneZ(pl) Else oldZ = PT_PlaneOffset(pl)
    If Abs(oldZ - zu) > tolPos Then
        If byOffset Then
            SetPlaneZ md, pl, zu
            refit = True
        Else
            PT_SetPlaneOffset md, pl, zu, False
        End If
        PT_LogAdd MACRO_NAME, "PT_Under", sideTxt & " plate face height", "", "moved", "Z", Round(oldZ, 3), zu
        nFeatChanged = nFeatChanged + 1
    End If
    If Abs(oldT - tPl) > tolSize Then
        If byOffset Then refit = True Else PT_SetExtrudeDepth md, feat, tPl
        PT_LogAdd MACRO_NAME, "PT_Plate", sideTxt & " plate thickness", "", "resized", "Thickness", Round(oldT, 3), tPl
        nFeatChanged = nFeatChanged + 1
    End If
    If refit Then FitPlate md, feat, zu, tPl
    md.EditRebuild3
End Sub

Sub SyncCuts(ByVal md As Object, ByVal zu As Double, ByVal tolPos As Double)
    Dim i As Long, want As Object, featName As String, p As String
    Set want = CreateObject("Scripting.Dictionary")
    For i = 0 To nCu - 1
        featName = "Clear_" & PT_SafeFile(cuKey(i))
        want.Add featName, 1
        p = PT_Um(cuBox(i, 0)) & "|" & PT_Um(cuBox(i, 1)) & "|" & PT_Um(cuBox(i, 2)) & "|" & PT_Um(cuBox(i, 3))
        SyncOne md, featName, p, cuKey(i), "Cut for " & cuName(i), "", zu, _
                cuBox(i, 0) - pOx, cuBox(i, 1) - pOy, cuBox(i, 0) - pOx + cuBox(i, 2), cuBox(i, 1) - pOy + cuBox(i, 3), 0, False
    Next
    SuppressOthers md, "Clear_", want
End Sub

Sub SyncPedestals(ByVal md As Object, ByVal zu As Double, ByVal tolPos As Double, ByVal tolSize As Double)
    Dim i As Long, want As Object, featName As String, p As String
    Set want = CreateObject("Scripting.Dictionary")
    For i = 0 To nCh - 1
        If chPed(i) > 0.0005 Then
            featName = "Ped_" & PT_SafeFile(chRef(i))
            want.Add featName, 1
            p = PT_Um(chBox(i, 0)) & "|" & PT_Um(chBox(i, 1)) & "|" & PT_Um(chBox(i, 2)) & "|" & PT_Um(chBox(i, 3)) & "|" & PT_Um(chPed(i))
            SyncOne md, featName, p, featName, "Pedestal for " & chRef(i), chRef(i), zu, _
                    chBox(i, 0) - pOx, chBox(i, 1) - pOy, chBox(i, 0) - pOx + chBox(i, 2), chBox(i, 1) - pOy + chBox(i, 3), chPed(i), True
        Else
            notes.Add chRef(i) & ": no pedestal (gap filled by the TIM). Its contact is the Underside face over the chip."
        End If
    Next
    SuppressOthers md, "Ped_", want
End Sub

' One macro feature (cut or pedestal): leave it, bring it back, rebuild it, or make it.
Sub SyncOne(ByVal md As Object, ByVal featName As String, ByVal params As String, ByVal logKey As String, _
            ByVal logName As String, ByVal refD As String, ByVal zu As Double, _
            ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
            ByVal pedH As Double, ByVal isPed As Boolean)
    Dim f As Object, oldP As String, pl As Object, sk As Object, feat As Object, zmin As Double, replaced As Boolean, pedMin As Double
    Dim emptyNames As Variant
    On Error GoTo failed
    Set f = PT_FindFeature(md, featName)
    oldP = PT_GetProp(md, "PT_" & featName)
    If Not f Is Nothing Then
        If oldP = params Then
            If f.IsSuppressed Then
                f.SetSuppression2 swUnSuppressFeature, swAllConfiguration, emptyNames
                If doLog Then PT_LogAdd MACRO_NAME, logKey, logName, refD, "added", "Brought back (was suppressed)", "", ""
                nFeatChanged = nFeatChanged + 1
            End If
            Exit Sub
        End If
        PT_DeleteFeature md, f
        replaced = True
    End If

    Set pl = PT_FindFeature(md, "PT_Under")
    Set sk = PT_SketchRect(md, pl, x0, y0, x1, y1, zu)
    sk.Name = featName & "_S"
    If isPed Then
        Set feat = PT_Boss(md, sk, pedH, Not bottomSide)     ' top plate: pedestal goes down; bottom plate: up
        md.EditRebuild3
        If SolidCount(md) > 1 Then                        ' not joined to the plate: the chip is outside the outline
            PT_DeleteFeature md, feat
            notes.Add refD & ": no pedestal, the chip is outside the plate outline. Redraw the Outline over it or set Covered = No."
            Exit Sub
        End If
        pedMin = IIf(bottomSide, zu, zu - pedH)             ' lowest Z of the pedestal
        zmin = PT_FeatureZMin(feat)
        If Abs(zmin - pedMin) > 0.002 Then                  ' went the wrong way: turn it
            PT_FlipExtrude md, feat
            md.EditRebuild3
            zmin = PT_FeatureZMin(feat)
            If Abs(zmin - pedMin) > 0.002 Then Err.Raise vbObjectError + 50, , "pedestal lowest point is at Z " & Round(zmin, 3) & _
                                                          ", expected " & Round(pedMin, 3)
        End If
    Else
        Set feat = PT_CutThrough(md, sk)
        If feat Is Nothing Then                            ' the tall part is outside the plate outline: nothing to cut
            PT_DeleteFeature md, sk
            notes.Add logName & ": skipped, the part is outside the plate outline."
            Exit Sub
        End If
    End If
    feat.Name = featName
    PT_SetProp md, "PT_" & featName, params
    If replaced Then
        nFeatChanged = nFeatChanged + 1
        PT_LogAdd MACRO_NAME, logKey, logName, refD, IIf(SameSize(oldP, params), "moved", "resized"), "Box", ParamText(oldP), ParamText(params)
    Else
        nFeatAdded = nFeatAdded + 1
        If doLog Then PT_LogAdd MACRO_NAME, logKey, logName, refD, "added", "Box", "", ParamText(params)
    End If
    Exit Sub
failed:
    failMsgs.Add featName & ": " & Err.Description
End Sub

' Suppresses macro features with this prefix that are not wanted any more.
Sub SuppressOthers(ByVal md As Object, ByVal prefixS As String, ByVal want As Object)
    Dim f As Object, emptyNames As Variant, nm As String
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        nm = f.Name
        If Left(nm, Len(prefixS)) = prefixS And Right(nm, 2) <> "_S" Then
            If Not want.Exists(nm) And Not f.IsSuppressed Then
                f.SetSuppression2 swSuppressFeature, swAllConfiguration, emptyNames
                PT_LogAdd MACRO_NAME, nm, IIf(prefixS = "Ped_", "Pedestal", "Cut"), IIf(prefixS = "Ped_", Mid(nm, 5), ""), _
                          "removed", "Suppressed in the spreader", "", ""
                nFeatSupp = nFeatSupp + 1
            End If
        End If
        Set f = f.GetNextFeature
    Loop
End Sub

Sub NameSpreaderFaces(ByVal md As Object, ByVal zu As Double, ByVal tPl As Double)
    Dim i As Long, big As Double
    big = 1E+30
    ' Top = outer face, Underside = face toward the board. Top plate: outer face up. Bottom plate: outer face down.
    If bottomSide Then
        If Not PT_NameFlatFace(md, False, zu - tPl, "Top", -big, -big, big, big) Then failMsgs.Add "Could not find the plate outer face to name it Top"
        If Not PT_NameFlatFace(md, True, zu, "Underside", -big, -big, big, big) Then failMsgs.Add "Could not find the plate face toward the board to name it Underside"
    Else
        If Not PT_NameFlatFace(md, True, zu + tPl, "Top", -big, -big, big, big) Then
            failMsgs.Add "Could not find the plate top face to name it Top (faces seen: " & PT_AllFaces(md).Count & ", part " & md.GetTitle & ", active " & swApp.ActiveDoc.GetTitle & ")"
        End If
        If Not PT_NameFlatFace(md, False, zu, "Underside", -big, -big, big, big) Then failMsgs.Add "Could not find the plate underside to name it Underside"
    End If
    For i = 0 To nCh - 1
        If chPed(i) > 0.0005 And Not PT_FindFeature(md, "Ped_" & PT_SafeFile(chRef(i))) Is Nothing Then
            If Not PT_NameFlatFace(md, bottomSide, IIf(bottomSide, zu + chPed(i), zu - chPed(i)), "Contact_" & PT_SafeFile(chRef(i)), _
                                   chBox(i, 0) - pOx, chBox(i, 1) - pOy, chBox(i, 0) - pOx + chBox(i, 2), chBox(i, 1) - pOy + chBox(i, 3)) Then
                failMsgs.Add "Could not find the pedestal bottom of " & chRef(i) & " to name it Contact_" & chRef(i)
            End If
        End If
    Next
End Sub

' ================================================================= TIM parts
Sub SyncTims(ByVal asm As Object, ByVal partTmpl As String, ByVal outDir As String, ByVal prefix As String, _
             ByVal tolPos As Double, ByVal tolSize As Double)
    Dim have As Object, wasSupp As Object, want As Object, i As Long, keyS As String, res As Long, fpath As String
    Dim topN As String, botN As String, nameS As String, z0 As Double
    Set have = CreateObject("Scripting.Dictionary")
    Set wasSupp = CreateObject("Scripting.Dictionary")
    Set want = CreateObject("Scripting.Dictionary")
    PT_ReadTagged asm, "TIM", have, wasSupp
    ' Face names: _Top is the face toward the plate, _Bottom the face on the chip, on both sides of the board.
    ' Top plate: the TIM sits on the chip top (z0 = chip top). Bottom plate: it hangs under the chip bottom (z0 = chip bottom - TIM).
    For i = 0 To nT - 1
        keyS = "tim_" & LCase(PT_SafeFile(tRef(i)))
        If Not want.Exists(keyS) Then
            want.Add keyS, 1
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
                If wasSupp(keyS) And doLog Then PT_LogAdd MACRO_NAME, keyS, nameS, tRef(i), "added", "Brought back (was suppressed)", "", ""
                res = PT_CheckBoxPart(asm, have(keyS), "TIM", keyS, nameS, tRef(i), 0, _
                                      tBox(i, 0), tBox(i, 1), z0, tBox(i, 2), tBox(i, 3), tTim(i), _
                                      topN, botN, tolPos, tolSize, MACRO_NAME, True)
                If res < 0 Then failMsgs.Add "Could not update " & nameS & ": " & PT_LastErr
                If res > 0 Then nTimChanged = nTimChanged + 1
            Else
                fpath = outDir & prefix & "TIM_" & PT_SafeFile(tRef(i)) & ".sldprt"
                If PT_AddBoxPart(asm, partTmpl, fpath, "TIM", keyS, nameS, tRef(i), 0, _
                                 tBox(i, 0), tBox(i, 1), z0, tBox(i, 2), tBox(i, 3), tTim(i), _
                                 topN, botN, MACRO_NAME) Then
                    nTimAdded = nTimAdded + 1
                    If doLog Then PT_LogAdd MACRO_NAME, keyS, nameS, tRef(i), "added", "Box", "", _
                                            PT_BoxText(tBox(i, 0), tBox(i, 1), z0, tBox(i, 2), tBox(i, 3), tTim(i))
                Else
                    failMsgs.Add "Could not add " & nameS & ": " & PT_LastErr
                End If
            End If
        End If
    Next
    nTimSupp = PT_SuppressUnwanted(have, wasSupp, want, MACRO_NAME, True)
End Sub

' ================================================================= helpers
Function CountPlanes(ByVal md As Object) As Long
    Dim f As Object, n As Long
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then n = n + 1
        Set f = f.GetNextFeature
    Loop
    CountPlanes = n
End Function

Function NextRev(ByVal outDir As String, ByVal prefix As String) As Long
    Dim f As String, n As Long, stem As String, numS As String
    stem = LCase(prefix & "Spreader_r")
    f = Dir(outDir & prefix & "Spreader_r*.sldprt")
    Do While f <> ""
        numS = Mid(f, Len(stem) + 1)
        numS = Left(numS, InStr(numS & ".", ".") - 1)
        If IsNumeric(numS) Then
            If CLng(numS) > n Then n = CLng(numS)
        End If
        f = Dir
    Loop
    NextRev = n + 1
End Function

' "x|y|L|W|h" in micrometres -> readable mm text
Function ParamText(ByVal p As String) As String
    Dim a As Variant, i As Long, s As String
    If p = "" Then Exit Function
    a = Split(p, "|")
    For i = 0 To UBound(a)
        If IsNumeric(a(i)) Then s = s & IIf(i = 0, "", IIf(i = 2, " size ", IIf(i = 4, " h ", ", "))) & CDbl(a(i)) / 1000#
    Next
    ParamText = "at " & s
End Function

' same size (L, W and height), only the position differs
Function SameSize(ByVal p1 As String, ByVal p2 As String) As Boolean
    Dim a As Variant, b As Variant, i As Long
    If p1 = "" Then Exit Function
    a = Split(p1, "|"): b = Split(p2, "|")
    If UBound(a) <> UBound(b) Then Exit Function
    For i = 2 To UBound(a)
        If a(i) <> b(i) Then Exit Function
    Next
    SameSize = True
End Function

Sub WriteReport(ByVal rptPath As String, ByVal asmPath As String)
    Dim fso As Object, ts As Object, it As Variant
    On Error GoTo done
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(rptPath, True)
    ts.WriteLine "BuildSpreader report " & Format(Now, "yyyy-mm-dd hh:nn")
    ts.WriteLine "Assembly: " & asmPath
    ts.WriteLine "Cuts and pedestals: added " & nFeatAdded & ", rebuilt " & nFeatChanged & ", suppressed " & nFeatSupp
    ts.WriteLine "TIM parts: added " & nTimAdded & ", changed " & nTimChanged & ", suppressed " & nTimSupp
    ts.WriteLine ""
    ts.Write rptBody
    ts.WriteLine "PROBLEMS (" & failMsgs.Count & ")"
    For Each it In failMsgs: ts.WriteLine "  " & it: Next
    ts.WriteLine ""
    ts.WriteLine "COVERED CHIPS SKIPPED (" & skipped.Count & ")"
    For Each it In skipped: ts.WriteLine "  " & it: Next
    ts.WriteLine ""
    ts.WriteLine "NOTES"
    For Each it In notes: ts.WriteLine "  " & it: Next
    ts.Close
done:
End Sub

' ================================================================= user-sketched spreader
' The spreader part in the assembly that has "spreader" in its name and is not suppressed.
' A component selected in the assembly wins. nFound = how many matched.
Function FindSpreaderComp(ByVal asm As Object, ByRef nFound As Long, ByVal wantBottom As Boolean) As Object
    Dim comps As Variant, i As Long, c As Object, found As Object, sel As Object, nm As String
    nFound = 0
    On Error Resume Next
    Set sel = asm.SelectionManager
    If sel.GetSelectedObjectCount2(-1) = 1 Then
        Set c = sel.GetSelectedObjectsComponent4(1, -1)
        If Not c Is Nothing Then
            nm = LCase(c.Name2)
            If InStr(nm, "spreader") > 0 And (InStr(nm, "bottom") > 0) = wantBottom Then
                nFound = 1
                Set FindSpreaderComp = c
                Exit Function
            End If
        End If
    End If
    On Error GoTo 0
    comps = asm.GetComponents(True)
    If IsEmpty(comps) Then Exit Function
    For i = 0 To UBound(comps)
        Set c = comps(i)
        If c.GetSuppression2 <> PT_swSuppressed Then
            nm = LCase(c.Name2)
            If InStr(nm, "spreader") > 0 And (InStr(nm, "bottom") > 0) = wantBottom Then nFound = nFound + 1: Set found = c
        End If
    Next
    Set FindSpreaderComp = found
End Function

' Position of the spreader part in the assembly (mm). False if it is turned (rotation is not the identity).
Function ReadOffset(ByVal c As Object) As Boolean
    Dim d As Variant, k As Long
    d = c.Transform2.ArrayData
    For k = 0 To 8
        If Abs(d(k) - IIf(k Mod 4 = 0, 1#, 0#)) > 0.000001 Then Exit Function
    Next
    pOx = d(9) * 1000#: pOy = d(10) * 1000#: pOz = d(11) * 1000#
    ReadOffset = True
End Function

' The plate outline is the sketch named Outline. One unnamed sketch is taken as the outline.
Sub PrepareOutline(ByVal md As Object)
    Dim f As Object, n As Long, lastS As Object
    If Not PT_FindFeature(md, "Outline") Is Nothing Then CheckOutlinePlane md: Exit Sub
    Set f = md.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "ProfileFeature" Then n = n + 1: Set lastS = f
        Set f = f.GetNextFeature
    Loop
    If n = 1 Then
        lastS.Name = "Outline"
        notes.Add "The only sketch in the spreader part was renamed Outline."
        CheckOutlinePlane md
        Exit Sub
    End If
    If n = 0 Then Err.Raise vbObjectError + 36, , "the spreader part has no sketch. Sketch the plate outline on the Front Plane, exit the sketch and name it Outline."
    Err.Raise vbObjectError + 36, , "the spreader part has " & n & " sketches. Rename the one with the plate outline to Outline."
End Sub

Function SolidCount(ByVal md As Object) As Long
    Dim b As Variant
    b = md.GetBodies2(PT_swSolidBody, True)
    If IsEmpty(b) Then Exit Function
    SolidCount = UBound(b) + 1
End Function

' Underside plane at zuP, then the plate: the Outline sketch (any shape) extruded tPl, starting zuP above its plane.

' Plate extrude: starts startMm from the sketch plane, depthMm thick.
Sub PlateDef(ByVal md As Object, ByVal feat As Object, ByVal startMm As Double, ByVal depthMm As Double, _
             ByVal revDir As Boolean, ByVal revOff As Boolean)
    Dim def As Object
    Set def = feat.GetDefinition
    def.AccessSelections md, Nothing
    def.FromType = 3                                       ' swStartOffset
    def.FromOffsetDistance = startMm / 1000#
    def.FromOffsetReverse = revOff
    def.ReverseDirection = revDir
    def.SetDepth True, depthMm / 1000#
    feat.ModifyDefinition def, md, Nothing
End Sub

Function PlateStartsByOffset(ByVal md As Object, ByVal feat As Object) As Boolean
    Dim def As Object
    Set def = feat.GetDefinition                         ' read only: AccessSelections would leave the part rolled back
    PlateStartsByOffset = (def.FromType = 3)
End Function



' The board lies in the Front Plane of the thermal assembly (board X, Y in the plane, Z up), whatever way the STEP
' had it. The Outline must be on a plane parallel to that, not on the Top or Right plane.
Sub CheckOutlinePlane(ByVal md As Object)
    Dim sk As Object, d As Variant
    Set sk = PT_FindFeature(md, "Outline")
    If sk Is Nothing Then Exit Sub
    d = sk.GetSpecificFeature2.ModelToSketchTransform.ArrayData
    If Abs(d(8)) < 0.9999 Then
        Err.Raise vbObjectError + 36, , "the Outline sketch is not parallel to the board. In the thermal assembly the board lies in the " & _
                  "Front Plane (Z is up through the board), not the Top Plane. Draw the Outline on the Front Plane or on a face parallel to it."
    End If
End Sub

' Underside plane at zuP (board Z of the plate face toward the board, in the part frame), then the plate:
' the Outline sketch (any shape) extruded tPl away from the board, starting zuP from its plane.
Sub MakeUserPlate(ByVal md As Object, ByVal zuP As Double, ByVal tPl As Double)
    Dim front As Object, pl As Object, sk As Object, feat As Object, nBefore As Long
    Set sk = PT_FindFeature(md, "Outline")
    Set front = PT_RefPlane(md, 1)
    If front Is Nothing Then Err.Raise vbObjectError + 40, , "the spreader part has no reference plane"
    nBefore = CountPlanes(md)
    md.ClearSelection2 True
    front.Select2 False, 0
    md.CreatePlaneAtOffset3 Abs(zuP) / 1000#, (zuP < 0), True
    md.ClearSelection2 True
    If CountPlanes(md) <= nBefore Then Err.Raise vbObjectError + 41, , "the underside plane was not made"
    Set pl = PT_LastFeatureOfType(md, "RefPlane")
    pl.Name = "PT_Under"
    If Abs(PlaneZ(pl) - zuP) > 0.001 Then SetPlaneZ md, pl, zuP
    Set feat = PT_Boss(md, sk, tPl, False)
    feat.Name = "PT_Plate"
    FitPlate md, feat, zuP, tPl
End Sub

' Puts the plate at its Z range: top plate zuP to zuP + tPl, bottom plate zuP - tPl to zuP. Tries the direction flips.
Sub FitPlate(ByVal md As Object, ByVal feat As Object, ByVal zuP As Double, ByVal tPl As Double)
    Dim k As Long, zr As Variant, ok As Boolean, zlo As Double, zhi As Double
    zlo = IIf(bottomSide, zuP - tPl, zuP)
    zhi = IIf(bottomSide, zuP, zuP + tPl)
    For k = 0 To 3
        PlateDef md, feat, Abs(zuP), tPl, ((k And 1) <> 0), ((k And 2) <> 0)
        md.EditRebuild3
        zr = FeatZRange(feat)
        If Not IsEmpty(zr) Then
            If Abs(zr(0) - zlo) < 0.001 And Abs(zr(1) - zhi) < 0.001 Then ok = True: Exit For
        End If
    Next
    If ok Then Exit Sub
    If IsEmpty(zr) Then Err.Raise vbObjectError + 42, , "the Outline sketch made no solid. Is it a closed outline?"
    Err.Raise vbObjectError + 43, , "the plate ended up at Z " & Round(zr(0), 3) & " to " & Round(zr(1), 3) & _
              " mm in the part, expected " & Round(zlo, 3) & " to " & Round(zhi, 3) & ". Was the Outline drawn on the Front Plane?"
End Sub

' Z (mm, part frame) of a reference plane
Function PlaneZ(ByVal pl As Object) As Double
    Dim d As Variant
    d = pl.GetSpecificFeature2.Transform.ArrayData
    PlaneZ = d(11) * 1000#
End Function

' Moves the underside plane to Z zmm (either sign): sets the distance, turns the direction if it went the wrong way.
Sub SetPlaneZ(ByVal md As Object, ByVal pl As Object, ByVal zmm As Double)
    PT_SetPlaneOffset md, pl, Abs(zmm), False
    md.EditRebuild3
    If Abs(PlaneZ(pl) - zmm) <= 0.001 Then Exit Sub
    PT_SetPlaneOffset md, pl, Abs(zmm), True
    md.EditRebuild3
    If Abs(PlaneZ(pl) - zmm) > 0.001 Then Err.Raise vbObjectError + 44, , "could not put the underside plane at Z " & Round(zmm, 3) & " mm (it is at " & Round(PlaneZ(pl), 3) & ")"
End Sub

' Lowest and highest Z (mm) over the faces a feature made: Array(zmin, zmax), or Empty
Function FeatZRange(ByVal feat As Object) As Variant
    Dim faces As Variant, i As Long, bb As Variant, zmin As Double, zmax As Double
    faces = feat.GetFaces
    If Not IsArray(faces) Then Exit Function
    zmin = 1E+30: zmax = -1E+30
    For i = 0 To UBound(faces)
        bb = faces(i).GetBox
        If bb(2) * 1000# < zmin Then zmin = bb(2) * 1000#
        If bb(5) * 1000# > zmax Then zmax = bb(5) * 1000#
    Next
    FeatZRange = Array(zmin, zmax)
End Function

' The cut and pedestal boxes are in the part frame. If the part has moved in the assembly since the last run, they are
' stale: forget their stored sizes so SyncOne rebuilds them in the right place.
Sub CheckOrigin(ByVal md As Object)
    Dim cur As String, oldO As String, f As Object, nm As String, n As Long
    cur = PT_Um(pOx) & "|" & PT_Um(pOy) & "|" & PT_Um(pOz)
    oldO = PT_GetProp(md, "PT_Origin")
    If oldO <> "" And oldO <> cur Then
        Set f = md.FirstFeature
        Do While Not f Is Nothing
            nm = f.Name
            If (Left(nm, 4) = "Ped_" Or Left(nm, 6) = "Clear_") And Right(nm, 2) <> "_S" Then
                PT_SetProp md, "PT_" & nm, "moved"
                n = n + 1
            End If
            Set f = f.GetNextFeature
        Loop
        If n > 0 Then notes.Add "The spreader part has moved in the assembly: " & n & " pedestals and cuts are rebuilt in their new place."
    End If
    PT_SetProp md, "PT_Origin", cur
End Sub
