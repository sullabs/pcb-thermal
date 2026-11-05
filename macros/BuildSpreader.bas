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
' Project folder (page, W7): pick <prefix>_project.json at the start. The plates come from forSolidWorks plates
' (face height, thickness, covered chips with TIM and pedestal, clearance cuts), worked out by the page on Save.
' The outcome goes to <prefix>_cad.json spreaderBuild, the change records to its changes. Refuses while cad.json
' has change records not ticked on the page.
' Ansys page (RunAnsys, the Build Solution button): the plates come from study.json solution.forSolidWorks of a board
' folder of ansys\page.html, in the board STEP's frame. The assembly is <folder>\solidworks\<board>_solution.sldasm,
' made from the board STEP on the first run; the plate parts are Spreader_Top / Spreader_Bottom there. The outcome goes
' to solution_cad.json build. No cuts: tall parts in the way are the page's warning, cut them by hand.
' Needs modules PTCommon and PTJson in the same macro. Old workbooks: a copy of PCB_thermal_TEMPLATE.xlsx.
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
Dim jsonMode As Boolean                                  ' project folder mode: the plate of the side comes from jPlate
Dim jPlate As Object                                     ' forSolidWorks plates top or bottom of the side being built
Dim sideRec(0 To 1) As Object                            ' cad.json spreaderBuild top and bottom of this run
Dim ansysMode As Boolean                                 ' Ansys page mode (RunAnsys): jPlate is a plate of study.json solution.forSolidWorks

Sub main()
    Set swApp = Application.SldWorks
    PT_Trace MACRO_NAME & ": started (VBA macro)."
    PT_LogReset
    Set skipped = New Collection: Set failMsgs = New Collection: Set notes = New Collection
    nFeatAdded = 0: nFeatChanged = 0: nFeatSupp = 0: nTimAdded = 0: nTimChanged = 0: nTimSupp = 0
    nT = 0: rptBody = "": resMsg = ""
    jsonMode = False: ansysMode = False: Set jPlate = Nothing: Set sideRec(0) = Nothing: Set sideRec(1) = Nothing

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
    Set wb = PT_PickWorkbook(xl, MACRO_NAME, "Spreader|CAD geometry|Settings|Change log", PJ_WbPath)
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

' ================================================================= project folder (page) mode
' Reads forSolidWorks plates (worked out by the page on every save), builds or updates the plates and TIM parts,
' and writes cad.json spreaderBuild (per plate: file, pedestals, TIM parts, cuts, problems) and changes.
Sub RunProject()
    Dim cad As Object, cadPath As String, fsDate As String, bDate As String
    jsonMode = True
    cadPath = PJ_Dir & "\" & PJ_File("cad")
    Set cad = PJ_ReadOwn(PJ_File("cad"), "pcb-thermal/cad")
    If cad Is Nothing Then
        PT_Msg "BuildSpreader did not run. Nothing was changed." & vbCrLf & vbCrLf & cadPath & " could not be read:" & vbCrLf & PT_LastErr, _
               vbExclamation, MACRO_NAME
        Exit Sub
    End If
    If PJ_Obj(cad, "build") Is Nothing Then
        PT_Msg "BuildSpreader did not run. Nothing was changed." & vbCrLf & vbCrLf & "There is no build in " & cadPath & "." & vbCrLf & _
               "Run Build Board first (button 3).", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    If PJ_Obj(PJ_Project, "forSolidWorks/plates") Is Nothing Then
        PT_Msg "BuildSpreader did not run. Nothing was changed." & vbCrLf & vbCrLf & _
               "The project has no enabled spreader plate. Add one on tab 7 of the page, cover the chips, save and run again.", _
               vbExclamation, MACRO_NAME
        Exit Sub
    End If
    fsDate = PJ_Str(PJ_Project, "forSolidWorks/computedAt", "")
    bDate = PJ_Str(cad, "build/date", "")
    If bDate <> "" And fsDate < bDate Then
        If PT_Msg("The page has not saved since the last Build Board (" & Replace(bDate, "T", " ") & ")." & vbCrLf & _
                  "The plate is built from what the page worked out on its last Save (" & Replace(fsDate, "T", " ") & ")." & vbCrLf & vbCrLf & _
                  "Build it anyway?", vbYesNo + vbQuestion, MACRO_NAME) <> vbYes Then Exit Sub
    End If
    If PJ_RefuseIfUnchecked(cad, MACRO_NAME) Then Exit Sub

    Dim pcbL As Double, pcbW As Double, tolPos As Double, tolSize As Double
    pcbL = PJ_Num(PJ_Project, "forSolidWorks/pcb/l", 0)
    pcbW = PJ_Num(PJ_Project, "forSolidWorks/pcb/w", 0)
    tolPos = PJ_Num(PJ_Project, "settings/tolPosMm", 0.05)
    tolSize = PJ_Num(PJ_Project, "settings/tolSizeMm", 0.05)

    ' ---------- assembly and files ----------
    Dim asm As Object, asmPath As String, partTmpl As String, outDir As String, prefix As String
    outDir = PJ_OutputFolder()
    prefix = PT_SafeFile(PJ_Prefix())
    partTmpl = PJ_Template("partTemplate", swDefaultTemplatePart)
    If partTmpl = "" Then
        PT_Msg "BuildSpreader did not run. The part template was not found. Leave it blank in the project settings to use the SolidWorks default.", _
               vbExclamation, MACRO_NAME
        Exit Sub
    End If
    asmPath = PJ_AsmPath()
    On Error GoTo swFail
    Set asm = PT_OpenAsm(asmPath)
    If asm Is Nothing Then
        On Error GoTo 0
        PT_Msg "BuildSpreader did not run. The thermal assembly is not there:" & vbCrLf & asmPath & vbCrLf & _
               "Run Build Board first.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    PT_ActivateDoc asm
    PT_LeaveHeatOnly asm
    asm.ResolveAllLightWeightComponents False

    ' ---------- the two sides ----------
    Dim pass As Long, res As Long, nBuilt As Long, errs As Long, warns As Long
    For pass = 0 To 1
        bottomSide = (pass = 1)
        Set jPlate = PJ_Obj(PJ_Project, "forSolidWorks/plates/" & IIf(bottomSide, "bottom", "top"))
        If Not jPlate Is Nothing Then
            res = DoSide(Nothing, asm, partTmpl, outDir, prefix, tolPos, tolSize, pcbL, pcbW)
            If res < 0 Then Exit Sub
            nBuilt = nBuilt + res
        End If
    Next
    If nBuilt = 0 Then
        On Error GoTo 0
        PT_Msg "BuildSpreader did nothing: no spreader part and no covered chips." & vbCrLf & _
               "Tick Covered for the chips on tab 7 of the page, save, and draw the plate outline (or let the macro make a rectangle).", _
               vbExclamation, MACRO_NAME
        Exit Sub
    End If

    ' ---------- TIM parts (both sides) ----------
    SyncTims asm, partTmpl, outDir, prefix, tolPos, tolSize
    PT_ActivateDoc asm
    asm.ForceRebuild3 False
    asm.Save3 PT_swSaveSilent + PT_swSaveReferenced, errs, warns
    On Error GoTo 0

    ' ---------- report, cad.json ----------
    Dim rptPath As String, nLog As Long, sb As Object, k As Long
    rptPath = outDir & prefix & "BuildSpreader_report.txt"
    WriteReport rptPath, asmPath
    Set sb = PJ_Obj(cad, "spreaderBuild")
    If sb Is Nothing Then Set sb = PJ_Dict()
    For k = 0 To 1
        If Not sideRec(k) Is Nothing Then
            PJ_Put sideRec(k), "report", rptPath
            PJ_Put sb, IIf(k = 1, "bottom", "top"), sideRec(k)
        End If
    Next
    PJ_Put cad, "spreaderBuild", sb
    nLog = PJ_FlushChanges(cad)
    If Not PJ_SaveOwn(cad, PJ_File("cad"), MACRO_NAME) Then
        PT_Msg "The spreader was saved, but " & cadPath & " could not be written:" & vbCrLf & PT_LastErr & vbCrLf & vbCrLf & _
               "The changes are listed in " & rptPath & ". Run Build Spreader again to write cad.json.", vbExclamation, MACRO_NAME
        Exit Sub
    End If

    Dim s As String
    s = "Written and saved to the spreader part(s) and the thermal assembly:" & vbCrLf & asmPath & vbCrLf & _
        "and " & cadPath & " (the page reloads it on its own, tab 7 Spreader)." & vbCrLf & vbCrLf & resMsg & _
        "Cuts and pedestals: " & nFeatAdded & " added, " & nFeatChanged & " rebuilt, " & nFeatSupp & " suppressed." & vbCrLf & _
        "TIM parts: " & nTimAdded & " added, " & nTimChanged & " changed, " & nTimSupp & " suppressed." & vbCrLf
    If nLog > 0 Then s = s & nLog & " change records written to cad.json. Tick them on tab 10 of the page." & vbCrLf
    If skipped.Count > 0 Then s = s & skipped.Count & " covered chips skipped (listed in the report)." & vbCrLf
    If notes.Count > 0 Then s = s & notes.Count & " notes (listed in the report)." & vbCrLf
    If failMsgs.Count > 0 Then s = s & failMsgs.Count & " problems (listed in the report and on tab 7)." & vbCrLf
    s = s & vbCrLf & "Report: " & rptPath
    PT_Msg s, IIf(failMsgs.Count > 0, vbExclamation, vbInformation), MACRO_NAME
    Exit Sub

swFail:
    If Err.Number >= vbObjectError + 34 And Err.Number <= vbObjectError + 36 Then
        PT_Msg "BuildSpreader did not run (" & IIf(bottomSide, "bottom", "top") & " spreader): " & Err.Description, vbExclamation, MACRO_NAME
        Exit Sub
    End If
    PT_Msg "BuildSpreader stopped in SolidWorks (" & IIf(bottomSide, "bottom", "top") & " spreader): " & Err.Description & vbCrLf & _
           "Nothing was saved and cad.json was not changed. Close the spreader part and the assembly without saving, send this message, and run again.", _
           vbExclamation, MACRO_NAME
End Sub

' ================================================================= Ansys page mode (Build Solution button)
' study.json solution.forSolidWorks plates[]: side, tMm, undersideZMm, outlineMm [x0, y0, x1, y1] and chips[] (refDes,
' boxMm [x0, y0, x1, y1], faceZMm, timTMm, pedestalMm), worked out by the page in the board STEP's frame (mm).
' Makes <folder>\solidworks\<board>_solution.sldasm from the board STEP once, then builds or updates the plates and the
' TIM parts in it as the project mode does: the Outline sketch and features you added are never touched.
Public Sub RunAnsys()
    Const ME_NAME As String = "BuildSolution"
    Set swApp = Application.SldWorks
    PT_Trace ME_NAME & ": started (VBA macro)."
    PT_LogReset
    PT_ShapeClear
    Set skipped = New Collection: Set failMsgs = New Collection: Set notes = New Collection
    nFeatAdded = 0: nFeatChanged = 0: nFeatSupp = 0: nTimAdded = 0: nTimChanged = 0: nTimSupp = 0
    nT = 0: nCu = 0: rptBody = "": resMsg = ""
    jsonMode = True: ansysMode = True: Set jPlate = Nothing: Set sideRec(0) = Nothing: Set sideRec(1) = Nothing
    If Not PJ_PickStudy(ME_NAME) Then Exit Sub

    Dim fs As Object, plates As Object, stepPath As String
    If Not PJ_Bool(PJ_Project, "solution/on", False) Then
        PT_Msg ME_NAME & " did not run. The thermal solution is off on the page (step 4). Turn it on, tick the parts and try again.", vbExclamation, ME_NAME
        Exit Sub
    End If
    Set fs = PJ_Obj(PJ_Project, "solution/forSolidWorks")
    If Not fs Is Nothing Then Set plates = PJ_Obj(fs, "plates")
    If plates Is Nothing Then
        PT_Msg ME_NAME & " did not run. study.json has no plate yet." & vbCrLf & vbCrLf & _
               "On the page, tick the parts under the plate (step 4). If it says board.json has no STEP frame, press Read heights from the STEP (step 1).", _
               vbExclamation, ME_NAME
        Exit Sub
    End If
    If plates.Count = 0 Then
        PT_Msg ME_NAME & " did not run. No ticked part has a body in the STEP, so there is no plate to build.", vbExclamation, ME_NAME
        Exit Sub
    End If
    stepPath = PJ_Dir & "\" & PJ_Str(fs, "stepFile", "")

    Dim asm As Object, asmPath As String, partTmpl As String, outDir As String
    asmPath = PJ_SolutionAsm()
    outDir = PJ_Dir & "\solidworks"
    PT_MakeDirs outDir
    outDir = outDir & "\"
    partTmpl = swApp.GetUserPreferenceStringValue(swDefaultTemplatePart)
    If partTmpl <> "" Then
        If Dir(partTmpl) = "" Then partTmpl = ""
    End If
    If partTmpl = "" Then
        PT_Msg ME_NAME & " did not run. The SolidWorks default part template was not found (Tools > Options > Default Templates).", vbExclamation, ME_NAME
        Exit Sub
    End If

    Dim pl As Variant, res As Long, nBuilt As Long, errs As Long, warns As Long
    On Error GoTo swFail
    Set asm = PT_OpenAsm(asmPath)
    If asm Is Nothing Then Set asm = NewSolutionAsm(stepPath, asmPath)
    PT_ActivateDoc asm
    asm.ResolveAllLightWeightComponents False
    CheckStepFrame asm, fs
    For Each pl In plates
        Set jPlate = pl
        bottomSide = (PJ_Str(jPlate, "side", "top") = "bottom")
        res = DoSide(Nothing, asm, partTmpl, outDir, "", 0.01, 0.01, 0, 0)
        If res < 0 Then Exit Sub
        nBuilt = nBuilt + res
    Next
    SyncTims asm, partTmpl, outDir, "", 0.01, 0.01
    PT_ActivateDoc asm
    asm.ForceRebuild3 False
    asm.Save3 PT_swSaveSilent + PT_swSaveReferenced, errs, warns
    On Error GoTo 0

    Dim rptPath As String, sc As Object, b As Object, k As Long
    rptPath = outDir & "BuildSolution_report.txt"
    WriteReport rptPath, asmPath
    Set sc = PJ_ReadOwn("solution_cad.json", "pcb-thermal-ansys/solution-cad")
    If sc Is Nothing Then
        PT_Msg "The plates were saved, but " & PJ_Dir & "\solution_cad.json could not be read, so it was left as it is:" & vbCrLf & PT_LastErr, vbExclamation, ME_NAME
        Exit Sub
    End If
    Set b = PJ_Dict()
    PJ_Put b, "date", PJ_Now()
    PJ_Put b, "assembly", asmPath
    PJ_Put b, "pageSavedAt", PJ_Str(fs, "computedAt", "")
    For k = 0 To 1
        If Not sideRec(k) Is Nothing Then PJ_Put b, IIf(k = 1, "bottom", "top"), sideRec(k)
    Next
    PJ_Put b, "notes", notes
    PJ_Put b, "report", rptPath
    PJ_Put sc, "build", b
    If Not PJ_SaveOwn(sc, "solution_cad.json", ME_NAME) Then
        PT_Msg "The plates were saved, but solution_cad.json could not be written:" & vbCrLf & PT_LastErr, vbExclamation, ME_NAME
        Exit Sub
    End If

    Dim s As String
    s = "Built and saved:" & vbCrLf & asmPath & vbCrLf & vbCrLf & resMsg & _
        "Pedestals: " & nFeatAdded & " added, " & nFeatChanged & " rebuilt, " & nFeatSupp & " suppressed." & vbCrLf & _
        "TIM parts: " & nTimAdded & " added, " & nTimChanged & " changed, " & nTimSupp & " suppressed." & vbCrLf
    If skipped.Count > 0 Then s = s & skipped.Count & " ticked parts skipped (listed in the report)." & vbCrLf
    If notes.Count > 0 Then s = s & notes.Count & " notes (listed in the report)." & vbCrLf
    If failMsgs.Count > 0 Then s = s & failMsgs.Count & " problems (listed in the report)." & vbCrLf
    s = s & vbCrLf & "Finish the plates here (holes, shape, cuts around parts in the way), then press Send Solution." & vbCrLf & _
        "A change of Lift, thickness or TIM on the page: let it save, then Build Solution again; your own features stay." & vbCrLf & vbCrLf & "Report: " & rptPath
    PT_Msg s, IIf(failMsgs.Count > 0, vbExclamation, vbInformation), ME_NAME
    Exit Sub

swFail:
    If Err.Number >= vbObjectError + 34 And Err.Number <= vbObjectError + 36 Then
        PT_Msg ME_NAME & " did not run (" & IIf(bottomSide, "bottom", "top") & " plate): " & Err.Description, vbExclamation, ME_NAME
        Exit Sub
    End If
    PT_Msg ME_NAME & " stopped in SolidWorks (" & IIf(bottomSide, "bottom", "top") & " plate): " & Err.Description & vbCrLf & _
           "Nothing was saved and solution_cad.json was not changed. Close the plate part and the assembly without saving, send this message, and run again.", _
           vbExclamation, ME_NAME
End Sub

' The board STEP opened as an assembly and saved as asmPath (its parts next to it).
Function NewSolutionAsm(ByVal stepPath As String, ByVal asmPath As String) As Object
    Dim impData As Object, errs As Long, warns As Long, md As Object
    If Right(stepPath, 1) = "\" Or Dir(stepPath) = "" Then Err.Raise vbObjectError + 34, , "the board STEP is not in the board folder: " & stepPath & _
                                                                   ". Put it there and press Read heights from the STEP on the page."
    Set impData = swApp.GetImportFileData(stepPath)
    Set md = swApp.LoadFile4(stepPath, "r", impData, errs)
    If md Is Nothing Then Err.Raise vbObjectError + 34, , "SolidWorks could not open " & stepPath & " (error " & errs & ")"
    If md.GetType <> PT_swDocASSEMBLY Then Err.Raise vbObjectError + 34, , "the board STEP opened as one part, not an assembly of parts. Export it again from Altium as an assembly."
    If Not md.Extension.SaveAs(asmPath, 0, PT_swSaveSilent + PT_swSaveReferenced, Nothing, errs, warns) Then
        Err.Raise vbObjectError + 34, , "could not save " & asmPath & " (error " & errs & ")"
    End If
    notes.Add "New assembly made from " & stepPath & "."
    Set NewSolutionAsm = md
End Function

' The page placed the plates by the board's top face in the STEP (frame/topZMm). SolidWorks keeps the STEP's
' coordinates on import; this checks it on the biggest flat component at any level, the board (of two the same size,
' the thinner: a STEP often wraps everything in one sub-assembly).
Sub CheckStepFrame(ByVal asm As Object, ByVal fs As Object)
    Dim comps As Variant, i As Long, bx As Variant, best As Variant, bestA As Double, a As Double, want As Double
    If Not PJ_IsNum(fs, "frame/topZMm") Then Exit Sub
    want = PJ_Num(fs, "frame/topZMm", 0)
    comps = asm.GetComponents(False)
    If IsEmpty(comps) Then Exit Sub
    For i = 0 To UBound(comps)
        If PT_GetProp(comps(i).GetModelDoc2, "Role") = "" Then
            bx = PT_CompBox(comps(i))
            If Not IsEmpty(bx) Then
                a = bx(3) * bx(4)
                If bx(5) < 0.25 * Sqr(a) Then
                    If a > bestA * 1.01 Then
                        bestA = a: best = bx
                    ElseIf a > bestA * 0.99 Then
                        If bx(5) < best(5) Then bestA = a: best = bx
                    End If
                End If
            End If
        End If
    Next
    If bestA = 0 Then Exit Sub
    If Abs(best(2) + best(5) - want) > 0.05 Then
        notes.Add "The board's top face is at Z " & Round(best(2) + best(5), 3) & " mm here, the page expects " & Round(want, 3) & _
                  " mm. The STEP may have been moved on import: check that the plates sit on the parts."
    End If
End Sub

' Ticked chips of a study.json plate (Ansys page mode).
Sub ReadChipsAnsys()
    Dim cl As Object, it As Variant, bx As Object, refD As String, n As Long, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Set cl = PJ_Obj(jPlate, "chips")
    nCh = 0
    If Not cl Is Nothing Then n = cl.Count
    ReDim chRef(n): ReDim chKey(n): ReDim chBox(n, 4): ReDim chPed(n): ReDim chTim(n): ReDim chBody(n)
    If n = 0 Then Exit Sub
    For Each it In cl
        refD = PJ_Str(it, "refDes", "")
        Set bx = PJ_Obj(it, "boxMm")
        If refD = "" Then
        ElseIf bx Is Nothing Then
            skipped.Add refD & ": no box in study.json (save the page again)"
        ElseIf bx.Count <> 4 Or seen.Exists(UCase(refD)) Then
            skipped.Add refD & ": listed twice or with a broken box, ignored"
        Else
            seen.Add UCase(refD), 1
            chRef(nCh) = refD
            chKey(nCh) = refD
            chBox(nCh, 0) = bx(1)
            chBox(nCh, 1) = bx(2)
            chBox(nCh, 2) = bx(3) - bx(1)
            chBox(nCh, 3) = bx(4) - bx(2)
            chBox(nCh, 4) = PJ_Num(it, "faceZMm", 0)
            chPed(nCh) = PJ_Num(it, "pedestalMm", 0)
            chTim(nCh) = PJ_Num(it, "timTMm", 0)
            chBody(nCh) = True
            nCh = nCh + 1
        End If
    Next
End Sub

' First plate of the Ansys page: the Outline sketch is the page's rectangle (outlineMm), on the Front Plane, so you can
' edit it like one you drew; then the underside plane and the plate, as for a user-drawn outline.
Sub MakeAnsysBase(ByVal md As Object, ByVal zu As Double, ByVal tPl As Double)
    Dim front As Object, sk As Object, o As Object
    Set o = PJ_Obj(jPlate, "outlineMm")
    If o Is Nothing Then Err.Raise vbObjectError + 34, , "study.json has no plate outline (save the page again)"
    Set front = PT_RefPlane(md, 1)
    If front Is Nothing Then Err.Raise vbObjectError + 40, , "the part template has no reference plane"
    Set sk = PT_SketchRect(md, front, o(1), o(2), o(3), o(4), 0)
    sk.Name = "Outline"
    MakeUserPlate md, zu, tPl
End Sub

' Covered chips of jPlate. The chip box comes from forSolidWorks bodies (by key); the face toward the plate is
' the body top for the top plate and the body bottom for the bottom plate.
Sub ReadChipsJson()
    Dim cl As Object, bl As Object, it As Variant, bx As Object, b As Variant, keyS As String, refD As String, n As Long, seen As Object
    If ansysMode Then ReadChipsAnsys: Exit Sub
    Set seen = CreateObject("Scripting.Dictionary")
    Set cl = PJ_Obj(jPlate, "chips")
    Set bl = PJ_Obj(PJ_Project, "forSolidWorks/bodies")
    nCh = 0
    If Not cl Is Nothing Then n = cl.Count
    ReDim chRef(n): ReDim chKey(n): ReDim chBox(n, 4): ReDim chPed(n): ReDim chTim(n): ReDim chBody(n)
    If n = 0 Then Exit Sub
    For Each it In cl
        refD = PJ_Str(it, "refDes", "")
        keyS = PJ_Str(it, "key", "")
        Set bx = Nothing
        If Not bl Is Nothing Then
            For Each b In bl
                If PJ_Str(b, "key", "") = keyS Then Set bx = PJ_Obj(b, "box"): Exit For
            Next
        End If
        If refD = "" Then
        ElseIf bx Is Nothing Then
            skipped.Add refD & ": its body " & keyS & " is not in the model (include it on tab 3 and save)"
        ElseIf seen.Exists(UCase(refD)) Then
            skipped.Add refD & ": listed twice, second one ignored"
        Else
            seen.Add UCase(refD), 1
            chRef(nCh) = refD
            chKey(nCh) = keyS
            chBox(nCh, 0) = PJ_Num(bx, "xMin", 0)
            chBox(nCh, 1) = PJ_Num(bx, "yMin", 0)
            chBox(nCh, 2) = PJ_Num(bx, "l", 0)
            chBox(nCh, 3) = PJ_Num(bx, "w", 0)
            If bottomSide Then chBox(nCh, 4) = PJ_Num(bx, "zMin", 0) Else chBox(nCh, 4) = PJ_Num(bx, "zMin", 0) + PJ_Num(bx, "h", 0)
            chPed(nCh) = PJ_Num(it, "pedestalMm", 0)
            chTim(nCh) = PJ_Num(it, "timTMm", 0)
            chBody(nCh) = (UCase(PJ_Str(it, "timMode", "Body")) = "BODY")
            nCh = nCh + 1
        End If
    Next
End Sub

' Clearance cuts of jPlate: tall parts the plate does not cover, already grown by the XY clearance on the page.
Sub ReadCutsJson()
    Dim cl As Object, it As Variant, keyS As String, n As Long, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Set cl = PJ_Obj(jPlate, "cuts")
    nCu = 0
    If Not cl Is Nothing Then n = cl.Count
    ReDim cuKey(n): ReDim cuName(n): ReDim cuBox(n, 3)
    If n = 0 Then Exit Sub
    For Each it In cl
        keyS = PJ_Str(it, "key", "")
        If keyS <> "" And Not seen.Exists(keyS) Then
            seen.Add keyS, 1
            cuKey(nCu) = keyS
            cuName(nCu) = PJ_Str(it, "name", keyS)
            cuBox(nCu, 0) = PJ_Num(it, "xMin", 0)
            cuBox(nCu, 1) = PJ_Num(it, "yMin", 0)
            cuBox(nCu, 2) = PJ_Num(it, "l", 0)
            cuBox(nCu, 3) = PJ_Num(it, "w", 0)
            nCu = nCu + 1
        End If
    Next
End Sub

' cad.json spreaderBuild record of the side just built. fail0 = problems before this side started.
Function SideRecord(ByVal sprPath As String, ByVal prefix As String, ByVal fail0 As Long) As Object
    Dim d As Object, tims As Collection, fl As Collection, i As Long, nPed As Long
    Set d = PJ_Dict()
    Set tims = New Collection
    Set fl = New Collection
    PJ_Put d, "date", PJ_Now()
    PJ_Put d, "file", sprPath
    For i = 0 To nCh - 1
        If chPed(i) > 0.0005 Then nPed = nPed + 1
        If chBody(i) And chTim(i) > 0.0005 Then tims.Add prefix & "TIM_" & PT_SafeFile(chRef(i)) & ".sldprt"
    Next
    PJ_Put d, "pedestals", nPed
    PJ_Put d, "tims", tims
    PJ_Put d, "cuts", nCu
    For i = fail0 + 1 To failMsgs.Count
        fl.Add CStr(failMsgs(i))
    Next
    PJ_Put d, "failures", fl
    Set SideRecord = d
End Function

' Builds or updates one plate (top, or bottom when bottomSide). Returns 1 built, 0 nothing to do, -1 stop the macro.
Function DoSide(ByVal wb As Object, ByVal asm As Object, ByVal partTmpl As String, ByVal outDir As String, ByVal prefix As String, _
                ByVal tolPos As Double, ByVal tolSize As Double, ByVal pcbL As Double, ByVal pcbW As Double) As Long
    Dim zu As Double, tPl As Double, zuP As Double, sideTxt As String, i As Long
    Dim sprComp As Object, md As Object, ans As VbMsgBoxResult, newRev As Boolean
    Dim sprPath As String, sprName As String, sprMsg As String, nFound As Long, errs As Long, warns As Long
    sideTxt = IIf(bottomSide, "Bottom", "Top")
    Dim fail0 As Long
    fail0 = failMsgs.Count
    If ansysMode Then
        zu = PJ_Num(jPlate, "undersideZMm", 0)
        tPl = PJ_Num(jPlate, "tMm", 0)
        ReadChipsJson
        ReadCutsJson
    ElseIf jsonMode Then
        zu = PJ_Num(jPlate, "faceZMm", 0)
        tPl = PJ_Num(jPlate, "thicknessMm", 0)
        ReadChipsJson
        ReadCutsJson
    Else
        zu = SNum(wb, "Spr_ZUnder", 0)
        tPl = SNum(wb, "Spr_Thickness", 0)
        ReadChips wb
        ReadCuts wb
    End If
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
    If tPl <= 0 Or (Not ansysMode And ((Not bottomSide And zu <= 0) Or (bottomSide And nCh = 0 And zu = 0))) Then
        PT_Msg "BuildSpreader did not run (" & LCase(sideTxt) & " spreader). On its tab the plate thickness (" & tPl & ") must be above 0 and " & _
               "the face height (" & zu & ") must be set." & vbCrLf & _
               "Mark at least one chip Covered = Yes, or set the face height to Manual.", vbExclamation, MACRO_NAME
        DoSide = -1
        Exit Function
    End If
    pOx = 0: pOy = 0: pOz = 0
    If sprComp Is Nothing And ansysMode Then
        newRev = True
    ElseIf sprComp Is Nothing Then
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
        If ansysMode Then                                   ' Spreader_Top / Spreader_Bottom; a stale file of that name is kept
            sprPath = outDir & "Spreader_" & sideTxt & ".sldprt"
            If Dir(sprPath) <> "" Then sprPath = outDir & "Spreader_" & sideTxt & "_" & Format(Now, "yyyymmdd_hhnnss") & ".sldprt"
        Else
            sprPath = outDir & prefix & "Spreader_r" & Format(NextRev(outDir, prefix), "00") & ".sldprt"
        End If
        Set md = swApp.NewDocument(partTmpl, 0, 0, 0)
        If md Is Nothing Then Err.Raise vbObjectError + 30, , "SolidWorks could not make a new part from " & partTmpl
        If ansysMode Then MakeAnsysBase md, zu, tPl Else MakeBase md, zu, tPl, pcbL, pcbW
        sprName = Mid(sprPath, InStrRev(sprPath, "\") + 1)
        PT_SetProp md, "Side", sideTxt
        PT_SetProp md, "Role", "Spreader"
        PT_SetProp md, "Key", LCase(Left(sprName, Len(sprName) - 7))
        PT_SetProp md, "Name", "Spreader " & sprName
        PT_SetProp md, "BuiltBy", MACRO_NAME
        If Not md.Extension.SaveAs(sprPath, 0, PT_swSaveSilent, Nothing, errs, warns) Then
            Err.Raise vbObjectError + 31, , "could not save " & sprPath & " (error " & errs & ")"
        End If
        sprMsg = IIf(ansysMode, "New " & LCase(sideTxt) & " plate made around the ticked parts.", "New rectangular spreader made from the PCB size.")
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
    If jsonMode Then Set sideRec(IIf(bottomSide, 1, 0)) = SideRecord(sprPath, prefix, fail0)
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
