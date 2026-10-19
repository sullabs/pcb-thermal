Attribute VB_Name = "SetupSim"
Option Explicit
' =====================================================================================
' SetupSim  -  SolidWorks 2023 macro (milestone M6, SolidWorks Simulation thermal)
'
' Makes a new steady-state thermal study on the thermal assembly from the Sim setup
' tab, meshes it and runs it. Every run is a new study (PT_yyyymmdd_hhmm); older
' studies are never changed. Works for any board. Then run ReadResults.
'
'   Materials   PCB orthotropic (k in-plane / through-plane from CAD geometry),
'               chips as calibrated blocks (k z = H / (theta jc x A), k xy from
'               Sim setup, blank = k z), parts without heat with one generic k,
'               spreader and TIMs from the Spreader tab.
'               Chips with Chip model = Two-resistor (needs theta jc and theta jb) are
'               made of three stacked bodies: bottom layer (theta jb to the board),
'               junction core (carries the heat), top layer (theta jc to the top).
'               Layer k = layer thickness / (theta x footprint area). The part file
'               of the chip is rebuilt in three layers when needed (M8).
'   Heat        each RefDes gets the Q used of the Component register (summed over
'               its rows), spread over the bodies with that RefDes (the junction core
'               only, for a two-resistor chip).
'   Contacts    global bonded; thermal resistance sets for part-to-PCB, chip-to-TIM
'               and TIM-to-spreader when their value on Sim setup is above 0; chips
'               in TIM mode "Contact R" get the TIM R of the Spreader tab.
'   Surfaces    convection and radiation to ambient per group (Sim setup table):
'               spreader top, spreader other, PCB top, PCB bottom and sides, parts
'               (faces touching another body are left out).
'   Mesh        global size, finer control on chips and TIMs. High quality.
'
' "Include parts without heat" = No runs the study in configuration PT_HeatOnly,
' where parts without heat are suppressed.
' If anything is missing (a RefDes with heat but no body, a named face...), the study
' is made but not meshed or run, and the message says what to fix.
'
' Refuses to run while the Change log has unchecked rows.
' Project folder (page, W7): pick <prefix>_project.json at the start. Inputs come from forSolidWorks (heat, chip k,
' PCB k, plates) and simSetup (mesh, surfaces, contacts); the new run, with a copy of those inputs, is added to
' <prefix>_results.json for Read Results and tab 9 of the page.
' Needs modules PTCommon and PTJson in the same macro. Old workbooks: a copy of PCB_thermal_TEMPLATE.xlsx.
' =====================================================================================

Const MACRO_NAME As String = "SetupSim"
Const swsThermal As Long = 3
Const swsSteadyState As Long = 1
Const swsFFEPlus As Long = 2
Const swsUnitSI As Long = 0
Const swsMM As Long = 0
Const swsBondedThermal As Long = 1
Const swsResistanceSet As Long = 0
Const swsResistTotal As Long = 0
Const swsResistDistributed As Long = 1
Const swsSurfaceToAmbient As Long = 0
Const swsHighQuality As Long = 1
Const CFG_HEATONLY As String = "PT_HeatOnly"

' Sim chips table offsets (Sim_Chips, column B = 0)
Const S_REF As Long = 0, S_MODEL As Long = 1, S_JC As Long = 2, S_JB As Long = 3, S_KZ As Long = 6, S_KXY As Long = 7
' Spreader chip table offsets (Spr_Chips, column B = 0)
Const P_REF As Long = 0, P_COV As Long = 7, P_TIMMODE As Long = 9, P_TIMK As Long = 11, P_PED As Long = 14, P_TIMR As Long = 15, P_CHECK As Long = 16

Dim problems As Collection, notes As Collection
Dim tag As Collection                     ' Array(role, key, refdes, component)
Dim pbar As Object                       ' SolidWorks progress bar (Progress, ProgressEnd)
Dim tagIdx As Object                      ' lower-case component name (full and last segment) -> position in tag; built on first use
Dim heatQ As Object                       ' RefDes (upper) -> W
Dim chipKz As Object, chipKxy As Object   ' RefDes (upper) -> W/m K
Dim timK As Object, timR As Object, timBody As Object, pedH As Object   ' RefDes (upper) -> values, covered chips only
Dim pcbT As Double
Dim kSprBot As Double, densSprBot As Double   ' bottom plate (Spreader bottom tab); same as the top plate when that tab is missing
Dim twoRes As Object                     ' RefDes (upper) -> Array(theta jc, theta jb, layer mm, area mm2); two-resistor chips only
Dim layerFrac As Double, kCore As Double  ' outer layer thickness as a fraction of chip height; k of the junction core
' Inputs of the run, from the workbook (ReadWorkbook) or the project folder (ReadProjectJson)
Dim jsonMode As Boolean
Dim amb As Double, meshG As Double, meshF As Double, inclNonHeat As Boolean, kParts As Double
Dim kPcbXY As Double, kPcbZ As Double, kSpr As Double, densSpr As Double, totalReg As Double
Dim rPart As Double, rChipTim As Double, rTimSpr As Double           ' contact resistances, K m2/W
Dim surfH(4) As Double, surfE(4) As Double, surfN(4) As String      ' spreader top, spreader other, PCB top, PCB bottom, parts

Sub main()
    Set swApp = Application.SldWorks
    PT_Trace MACRO_NAME & ": started (VBA macro)."
    Set problems = New Collection: Set notes = New Collection
    jsonMode = False

    ' ---------- project folder (page) or workbook ----------
    Dim xl As Object, wb As Object, res As Object, resPath As String
    Select Case PJ_PickSource(MACRO_NAME)
        Case "json"
            jsonMode = True
        Case "xlsx"
        Case Else
            Exit Sub
    End Select
    If jsonMode Then
        resPath = PJ_Dir & "\" & PJ_File("results")
        Set res = PJ_ReadOwn(PJ_File("results"), "pcb-thermal/results")
        If res Is Nothing Then
            PT_Msg "SetupSim did not run. Nothing was changed." & vbCrLf & vbCrLf & resPath & " could not be read:" & vbCrLf & PT_LastErr, _
                   vbExclamation, MACRO_NAME
            Exit Sub
        End If
        If Not ReadProjectJson() Then Exit Sub
    Else
        Set xl = PT_GetExcel()
        Set wb = PT_PickWorkbook(xl, MACRO_NAME, "Sim setup|Spreader|CAD geometry|Component register|Change log", PJ_WbPath)
        If wb Is Nothing Then Exit Sub
        If Not ReadWorkbook(xl, wb) Then Exit Sub
    End If

    ' ---------- assembly, configuration ----------
    Dim asm As Object, asmPath As String
    If jsonMode Then asmPath = PJ_AsmPath() Else asmPath = PT_AsmPath(wb)
    Set asm = PT_OpenAsm(asmPath)
    If asm Is Nothing Then
        PT_Msg "SetupSim did not run. No thermal assembly at" & vbCrLf & asmPath & vbCrLf & "Run BuildBoard first.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    PT_ActivateDoc asm
    On Error GoTo swFail
    Progress "SetupSim: preparing the assembly", 5
    SetConfiguration asm, inclNonHeat
    Set tag = PT_ListTagged(asm)
    Set tagIdx = Nothing
    If EnsureLayers() Then
        asm.EditRebuild3
        PT_ActivateDoc asm
        Set tag = PT_ListTagged(asm)
        Set tagIdx = Nothing
    End If
    ' Bare board study: no spreader in the assembly (not built yet, or suppressed with its TIMs).
    ' The chip rows of the spreader are then ignored and the chip tops lose heat to air.
    Dim itB As Variant, hasSpr As Boolean, sprFiles As Collection, sprNm As String
    Set sprFiles = New Collection
    For Each itB In tag
        If itB(0) = "Spreader" Then
            hasSpr = True
            sprNm = itB(3).GetPathName
            sprFiles.Add Mid(sprNm, InStrRev(sprNm, "\") + 1)
        End If
    Next
    If Not hasSpr Then
        timK.RemoveAll: timR.RemoveAll: timBody.RemoveAll: pedH.RemoveAll
        notes.Add "Bare board study: no spreader in the assembly, so the spreader chips were ignored."
    End If

    ' ---------- Simulation ----------
    Dim simDoc As Object, study As Object, runName As String, errC As Long, k As Long
    Set simDoc = PT_GetSimDoc(asm)
    If simDoc Is Nothing Then
        ProgressEnd
        PT_Msg "SetupSim did not run. SolidWorks Simulation is not available." & vbCrLf & _
               "Turn it on in Tools > Add-Ins > SOLIDWORKS Simulation, then run again.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    runName = "PT_" & Format(Now, "yyyymmdd_hhnn")
    Do While Not PT_FindStudy(simDoc, runName) Is Nothing
        k = k + 1
        runName = "PT_" & Format(Now, "yyyymmdd_hhnn") & "_" & k
    Loop
    Set study = simDoc.StudyManager.CreateNewStudy3(runName, swsThermal, 0, errC)
    If study Is Nothing Then Err.Raise vbObjectError + 60, , "could not create the study (error " & errC & ")"
    study.ThermalStudyOptions.SolutionType = swsSteadyState
    ' SolverType is left at its default (Automatic). Setting it to FFEPlus crashed SolidWorks 2023 SP5.

    Progress "SetupSim: materials", 20
    ApplyMaterials study, kParts, kPcbXY, kPcbZ, kSpr, densSpr
    Dim applied As Double
    Progress "SetupSim: heat, contacts, surfaces", 40
    applied = ApplyHeat(study)
    ApplyContacts study
    ApplySurfaces study, amb

    ' ---------- mesh and run, only when nothing is missing ----------
    Dim ran As Boolean, runErr As Long
    If problems.Count = 0 Then
        Progress "SetupSim: meshing", 70
        MeshStudy study, meshG, meshF
        If problems.Count = 0 Then
            Progress "SetupSim: solving. SolidWorks is busy for several minutes, do not close it", 85
            runErr = study.RunAnalysis
            If runErr <> 0 Then
                problems.Add "The analysis stopped with error " & runErr & " (swsRunAnalysisError_e; 24 = solver failed, " & _
                             "16 = thermal conductivity missing, 19/20 = material missing)."
            Else
                ran = True
            End If
        End If
    End If
    ProgressEnd
    asm.Save3 PT_swSaveSilent, errC, k
    On Error GoTo 0

    ' ---------- report ----------
    Dim rptPath As String, s As String, it As Variant, n As Long
    If jsonMode Then
        rptPath = PJ_OutputFolder() & PT_SafeFile(PJ_Prefix()) & "SetupSim_" & runName & ".txt"
    Else
        rptPath = PT_OutputFolder(wb) & PT_SafeFile(PT_Prefix(wb)) & "SetupSim_" & runName & ".txt"
    End If
    WriteReport rptPath, asmPath, runName, applied, totalReg, ran

    ' ---------- run name: results file (page) or Sim setup (workbook) ----------
    If jsonMode Then
        Dim runs As Object, rr As Object
        Set rr = NewRun(runName, asmPath, sprFiles, ran, rptPath, study)
        Set runs = PJ_Obj(res, "runs")
        If runs Is Nothing Then Set runs = New Collection
        runs.Add rr
        PJ_Put res, "runs", runs
        If Not PJ_SaveOwn(res, PJ_File("results"), MACRO_NAME) Then
            PT_Msg "The study " & runName & " was made, but " & resPath & " could not be written:" & vbCrLf & PT_LastErr & vbCrLf & _
                   "Read Results will not find the run. Run SetupSim again.", vbExclamation, MACRO_NAME
            Exit Sub
        End If
    Else
        On Error GoTo xlFail
        xl.ScreenUpdating = False
        PT_Rng(wb, "Sim_RunName").Value = runName
        xl.ScreenUpdating = True
        wb.Save
        On Error GoTo 0
    End If

    Dim whereS As String
    If jsonMode Then whereS = "(run written to " & resPath & ", tab 9 of the page)" Else whereS = "(run name written to Sim setup, " & wb.Name & ")"
    s = "Study " & runName & " made in" & vbCrLf & asmPath & vbCrLf & _
        whereS & vbCrLf & vbCrLf & _
        "Heat applied " & Round(applied, 3) & " W, " & IIf(jsonMode, "page", "register") & " total " & Round(totalReg, 3) & " W." & vbCrLf & _
        IIf(hasSpr, "", "Bare board study (no spreader). Compare it with the bare board hand check (tab 5)." & vbCrLf)
    If ran Then
        s = s & "Meshed and solved. Next: run Read Results (button 8)." & vbCrLf
    ElseIf problems.Count > 0 Then
        s = s & "NOT meshed or run. Fix these and run SetupSim again:" & vbCrLf
        For Each it In problems
            n = n + 1
            If n > 6 Then s = s & "(more in the report)" & vbCrLf: Exit For
            s = s & "- " & it & vbCrLf
        Next
    End If
    s = s & vbCrLf & "Report: " & rptPath
    PT_Msg s, IIf(ran, vbInformation, vbExclamation), MACRO_NAME
    Exit Sub

swFail:
    ProgressEnd
    PT_Msg "SetupSim stopped in SolidWorks: " & Err.Description & vbCrLf & _
           "A half-made study may be in the Simulation tree: delete it, send this message, and run again.", vbExclamation, MACRO_NAME
    Exit Sub
xlFail:
    xl.ScreenUpdating = True
    PT_Msg "The study was made, but writing the run name to Excel failed: " & Err.Description, vbExclamation, MACRO_NAME
End Sub

' ================================================================= inputs
' Workbook: the named cells of Sim setup, Spreader, CAD geometry and the register. False = stop (the user was told).
Function ReadWorkbook(ByVal xl As Object, ByVal wb As Object) As Boolean
    If Not PT_NeedNames(wb, "Sim_Type|Sim_RunName|Sim_Ambient|Sim_MeshGlobal|Sim_MeshFine|Sim_IncludeNonHeat|Sim_kParts|" & _
                            "Sim_Surfaces|Sim_Contacts|Sim_Chips|Spr_Chips|Spr_k|Spr_Density|Geo_kXY|Geo_kZ|Geo_PCB_Used|" & _
                            "Reg_RefDes|Reg_QUsed|Reg_TotalQ|Log_Header|Log_Unchecked|Set_Prefix|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Function
    If PT_RefuseIfUnchecked(wb, MACRO_NAME) Then Exit Function
    xl.Calculate
    If UCase(PT_NameStr(wb, "Sim_Type", "Thermal")) <> "THERMAL" Then
        PT_Msg "SetupSim did not run. Study type on Sim setup is " & PT_NameStr(wb, "Sim_Type", "") & _
               ". This macro does SolidWorks Simulation thermal only; use RunFlow for Flow Simulation.", vbExclamation, MACRO_NAME
        Exit Function
    End If
    amb = PT_NameNum(wb, "Sim_Ambient", 25)
    meshG = PT_NameNum(wb, "Sim_MeshGlobal", 3)
    meshF = PT_NameNum(wb, "Sim_MeshFine", 0.5)
    inclNonHeat = (UCase(PT_NameStr(wb, "Sim_IncludeNonHeat", "Yes")) <> "NO")
    kParts = PT_NameNum(wb, "Sim_kParts", 0.5)
    layerFrac = PT_NameNum(wb, "Sim_LayerFrac", 0.1)
    kCore = PT_NameNum(wb, "Sim_kCore", 100)
    kPcbXY = PT_NameNum(wb, "Geo_kXY", 0)
    kPcbZ = PT_NameNum(wb, "Geo_kZ", 0)
    kSpr = PT_NameNum(wb, "Spr_k", 0)
    densSpr = PT_NameNum(wb, "Spr_Density", 2700)
    kSprBot = kSpr: densSprBot = densSpr
    Dim wsBot As Object, rB As Object, rv As Variant, sv As Variant, g As Long
    Set wsBot = PT_Sheet(wb, "Spreader bottom")
    If Not wsBot Is Nothing Then
        On Error Resume Next
        Set rB = wsBot.Names("Spr_k").RefersToRange
        If Not rB Is Nothing Then kSprBot = PT_Num(rB.Cells(1, 1).Value, kSpr)
        Set rB = Nothing
        Set rB = wsBot.Names("Spr_Density").RefersToRange
        If Not rB Is Nothing Then densSprBot = PT_Num(rB.Cells(1, 1).Value, densSpr)
        On Error GoTo 0
    End If
    pcbT = PT_Num(PT_Rng(wb, "Geo_PCB_Used").Cells(3, 1).Value, 0)
    totalReg = PT_NameNum(wb, "Reg_TotalQ", 0)
    sv = PT_Rng(wb, "Sim_Surfaces").Value
    For g = 0 To 4
        surfH(g) = PT_Num(sv(g + 1, 2), 0): surfE(g) = PT_Num(sv(g + 1, 3), 0): surfN(g) = PT_Str(sv(g + 1, 1))
    Next
    If Not InputsOk("on Sim setup", "CAD geometry D19:D20") Then Exit Function
    ReadHeat wb
    ReadChips wb
    ReadTims wb
    ReadWorkbook = True
End Function

' Project folder: forSolidWorks (worked out by the page on Save) and simSetup. False = stop (the user was told).
Function ReadProjectJson() As Boolean
    Dim cad As Object, ss As Object, g As Long, grp As Variant, nm As Variant
    Set cad = PJ_ReadOwn(PJ_File("cad"), "pcb-thermal/cad")
    If cad Is Nothing Then
        PT_Msg "SetupSim did not run. " & PJ_File("cad") & " could not be read:" & vbCrLf & PT_LastErr, vbExclamation, MACRO_NAME
        Exit Function
    End If
    If PJ_Obj(PJ_Project, "forSolidWorks") Is Nothing Or PJ_Obj(PJ_Project, "simSetup") Is Nothing Then
        PT_Msg "SetupSim did not run. The project has no simulation setup or no worked out numbers yet." & vbCrLf & _
               "Open the page, check tab 8 (Simulation setup), press Save and run again.", vbExclamation, MACRO_NAME
        Exit Function
    End If
    If PJ_RefuseIfUnchecked(cad, MACRO_NAME) Then Exit Function
    Set ss = PJ_Obj(PJ_Project, "simSetup")
    If UCase(PJ_Str(ss, "studyType", "Thermal")) <> "THERMAL" Then
        PT_Msg "SetupSim did not run. The study type on tab 8 of the page is " & PJ_Str(ss, "studyType", "") & "." & vbCrLf & _
               "This button makes SolidWorks Simulation thermal studies. For Flow use Run Flow Simulation (button 7).", vbExclamation, MACRO_NAME
        Exit Function
    End If
    If PJ_Str(cad, "build/inputsHash", "") <> "" And PJ_Str(cad, "build/inputsHash", "") <> PJ_Str(PJ_Project, "forSolidWorks/inputsHash", "") Then
        If PT_Msg("The PCB or the bodies changed on the page after the last Build Board." & vbCrLf & _
                  "The study would use the old model. Run Build Board first to update it." & vbCrLf & vbCrLf & _
                  "Make the study on the old model anyway?", vbYesNo + vbQuestion, MACRO_NAME) <> vbYes Then Exit Function
    End If
    amb = PJ_Num(PJ_Project, "forSolidWorks/ambientC", 25)
    meshG = PJ_Num(ss, "meshGlobalMm", 3)
    meshF = PJ_Num(ss, "meshFineMm", 0.5)
    inclNonHeat = PJ_Bool(ss, "includePartsWithoutHeat", True)
    kParts = PJ_Num(ss, "kPartsWithoutHeat", 0.5)
    layerFrac = PJ_Num(ss, "layerFrac", 0.1)
    kCore = PJ_Num(ss, "kCore", 100)
    kPcbXY = PJ_Num(PJ_Project, "forSolidWorks/pcb/kxy", 0)
    kPcbZ = PJ_Num(PJ_Project, "forSolidWorks/pcb/kz", 0)
    pcbT = PJ_Num(PJ_Project, "forSolidWorks/pcb/h", 0)
    ' plate material; the bottom plate falls back to the top one, a board without plates gets aluminium (unused)
    kSpr = PJ_Num(PJ_Project, "forSolidWorks/plates/top/k", PJ_Num(PJ_Project, "forSolidWorks/plates/bottom/k", 167))
    densSpr = PJ_Num(PJ_Project, "forSolidWorks/plates/top/densityKgM3", PJ_Num(PJ_Project, "forSolidWorks/plates/bottom/densityKgM3", 2700))
    kSprBot = PJ_Num(PJ_Project, "forSolidWorks/plates/bottom/k", kSpr)
    densSprBot = PJ_Num(PJ_Project, "forSolidWorks/plates/bottom/densityKgM3", densSpr)
    rPart = PJ_Num(ss, "contacts/partToPcb", 0)
    rChipTim = PJ_Num(ss, "contacts/chipToTim", 0)
    rTimSpr = PJ_Num(ss, "contacts/timToSpreader", 0)
    grp = Array("spreaderTop", "spreaderOther", "pcbTop", "pcbBottom", "parts")
    nm = Array("Spreader top", "Spreader other", "PCB top", "PCB bottom and sides", "Parts")
    For g = 0 To 4
        surfH(g) = PJ_Num(ss, "surfaces/" & grp(g) & "/h", 0)
        surfE(g) = PJ_Num(ss, "surfaces/" & grp(g) & "/emissivity", 0)
        surfN(g) = nm(g)
    Next
    If Not InputsOk("on tab 8 of the page", "tab 3 PCB stack") Then Exit Function
    ReadHeatJson
    ReadChipsJson
    ReadTimsJson
    ReadProjectJson = True
End Function

' The checks both sources share. whereS: where the user fixes them.
Function InputsOk(ByVal whereS As String, ByVal pcbWhere As String) As Boolean
    If layerFrac <= 0 Or layerFrac >= 0.45 Or kCore <= 0 Then
        PT_Msg "SetupSim did not run. " & whereS & " the outer layer thickness must be above 0 and below 0.45 (share of the chip height) " & _
               "and the core k above 0.", vbExclamation, MACRO_NAME
        Exit Function
    End If
    If meshG <= 0 Or meshF <= 0 Or kParts <= 0 Or kPcbXY <= 0 Or kPcbZ <= 0 Or kSpr <= 0 Then
        PT_Msg "SetupSim did not run. Mesh sizes and k of parts without heat (" & whereS & "), PCB k (" & pcbWhere & ") and spreader k " & _
               "must all be above 0.", vbExclamation, MACRO_NAME
        Exit Function
    End If
    InputsOk = True
End Function

Sub ReadHeatJson()
    Dim h As Object, it As Variant
    Set heatQ = CreateObject("Scripting.Dictionary")
    totalReg = 0
    Set h = PJ_Obj(PJ_Project, "forSolidWorks/heat")
    If h Is Nothing Then Exit Sub
    For Each it In h.Keys
        If PJ_Num(h, CStr(it), 0) > 0 And Not heatQ.Exists(UCase(CStr(it))) Then
            heatQ.Add UCase(CStr(it)), PJ_Num(h, CStr(it), 0)
            totalReg = totalReg + PJ_Num(h, CStr(it), 0)
        End If
    Next
End Sub

Sub ReadChipsJson()
    Dim cl As Object, it As Variant, key As String, kz As Double, kxy As Double, thJc As Double, thJb As Double
    Set chipKz = CreateObject("Scripting.Dictionary")
    Set chipKxy = CreateObject("Scripting.Dictionary")
    Set twoRes = CreateObject("Scripting.Dictionary")
    Set cl = PJ_Obj(PJ_Project, "forSolidWorks/chips")
    If cl Is Nothing Then Exit Sub
    For Each it In cl
        key = UCase(PJ_Str(it, "refDes", ""))
        If key <> "" And Not chipKz.Exists(key) Then
            If UCase(PJ_Str(it, "model", "Block")) = "TWO-RESISTOR" Then
                thJc = PJ_Num(it, "rjcCW", 0)
                thJb = PJ_Num(it, "rjbCW", 0)
                If thJc > 0 And thJb > 0 And heatQ.Exists(key) Then
                    If Not twoRes.Exists(key) Then twoRes.Add key, Array(thJc, thJb, 0#, 0#)
                Else
                    notes.Add key & ": two-resistor needs Rjc, Rjb and heat; built as a calibrated block."
                End If
            End If
            kz = PJ_Num(it, "kz", 0)
            kxy = PJ_Num(it, "kxy", 0)
            If kz > 0 Then
                If kxy <= 0 Then kxy = kz
                chipKz.Add key, kz
                chipKxy.Add key, kxy
            End If
        End If
    Next
End Sub

' Covered chips of both plates: TIM k, TIM R (contact resistance mode), body or not, pedestal height.
Sub ReadTimsJson()
    Dim pass As Long, cl As Object, it As Variant, key As String
    Set timK = CreateObject("Scripting.Dictionary"): Set timR = CreateObject("Scripting.Dictionary")
    Set timBody = CreateObject("Scripting.Dictionary"): Set pedH = CreateObject("Scripting.Dictionary")
    For pass = 0 To 1
        Set cl = PJ_Obj(PJ_Project, "forSolidWorks/plates/" & IIf(pass = 0, "top", "bottom") & "/chips")
        If Not cl Is Nothing Then
            For Each it In cl
                key = UCase(PJ_Str(it, "refDes", ""))
                If key <> "" And Not timK.Exists(key) Then
                    timK.Add key, PJ_Num(it, "timK", 0)
                    timR.Add key, PJ_Num(it, "timRCW", 0)
                    timBody.Add key, (UCase(PJ_Str(it, "timMode", "Body")) = "BODY")
                    pedH.Add key, PJ_Num(it, "pedestalMm", 0)
                End If
            Next
        End If
    Next
End Sub

' results.json run for a new study, with a copy of the inputs it was made from (design 3.3).
Function NewRun(ByVal runName As String, ByVal asmPath As String, ByVal sprFiles As Collection, ByVal ran As Boolean, _
                ByVal rptPath As String, ByVal study As Object) As Object
    Dim d As Object, msh As Object, inp As Object, nl As Collection, it As Variant, nodes As Long, elems As Long
    Set d = PJ_Dict()
    PJ_Put d, "name", runName
    PJ_Put d, "date", PJ_Now()
    PJ_Put d, "studyType", "Thermal"
    PJ_Put d, "study", runName
    PJ_Put d, "status", IIf(ran, "solved", "set up")
    PJ_Put d, "assembly", asmPath
    If sprFiles.Count > 0 Then PJ_Put d, "spreaderFiles", sprFiles
    Set msh = PJ_Dict()
    PJ_Put msh, "globalMm", meshG
    PJ_Put msh, "fineMm", meshF
    If ran Then
        On Error Resume Next
        nodes = study.Mesh.NodeCount
        elems = study.Mesh.ElementCount
        On Error GoTo 0
        If nodes > 0 Then PJ_Put msh, "nodes", nodes
        If elems > 0 Then PJ_Put msh, "elements", elems
    End If
    PJ_Put d, "mesh", msh
    Set inp = PJ_Dict()
    PJ_Put inp, "forSolidWorks", PJ_Obj(PJ_Project, "forSolidWorks")
    PJ_Put inp, "simSetup", PJ_Obj(PJ_Project, "simSetup")
    PJ_Put d, "inputs", inp
    PJ_Put d, "report", rptPath
    Set nl = New Collection
    For Each it In problems: nl.Add "Problem: " & CStr(it): Next
    For Each it In notes: nl.Add CStr(it): Next
    If nl.Count > 0 Then PJ_Put d, "notes", nl
    Set NewRun = d
End Function

' ================================================================= Excel side
Sub ReadHeat(ByVal wb As Object)
    Dim refs As Variant, qs As Variant, r As Long, key As String
    Set heatQ = CreateObject("Scripting.Dictionary")
    refs = PT_Rng(wb, "Reg_RefDes").Value
    qs = PT_Rng(wb, "Reg_QUsed").Value
    For r = 1 To UBound(refs, 1)
        key = UCase(PT_Str(refs(r, 1)))
        If key <> "" And PT_Num(qs(r, 1), 0) > 0 Then
            If heatQ.Exists(key) Then heatQ(key) = heatQ(key) + PT_Num(qs(r, 1), 0) Else heatQ.Add key, PT_Num(qs(r, 1), 0)
        End If
    Next
End Sub

Sub ReadChips(ByVal wb As Object)
    Dim v As Variant, r As Long, key As String, kz As Double, kxy As Double, thJc As Double, thJb As Double
    Set chipKz = CreateObject("Scripting.Dictionary")
    Set chipKxy = CreateObject("Scripting.Dictionary")
    Set twoRes = CreateObject("Scripting.Dictionary")
    v = PT_Rng(wb, "Sim_Chips").Value
    For r = 1 To UBound(v, 1)
        key = UCase(PT_Str(v(r, 1 + S_REF)))
        If key <> "" And Not chipKz.Exists(key) Then
            If UCase(PT_Str(v(r, 1 + S_MODEL))) = "TWO-RESISTOR" Then
                thJc = PT_Num(v(r, 1 + S_JC), 0)
                thJb = PT_Num(v(r, 1 + S_JB), 0)
                If thJc > 0 And thJb > 0 And heatQ.Exists(key) Then
                    If Not twoRes.Exists(key) Then twoRes.Add key, Array(thJc, thJb, 0#, 0#)
                Else
                    notes.Add key & ": two-resistor needs theta jc, theta jb and heat; built as a calibrated block."
                End If
            End If
            kz = PT_Num(v(r, 1 + S_KZ), 0)
            kxy = PT_Num(v(r, 1 + S_KXY), 0)
            If kz > 0 Then
                If kxy <= 0 Then kxy = kz
                chipKz.Add key, kz
                chipKxy.Add key, kxy
            End If
        End If
    Next
End Sub

Sub ReadTims(ByVal wb As Object)
    Dim v As Variant, r As Long, key As String, pass As Long, rg As Object, wsB As Object
    Set timK = CreateObject("Scripting.Dictionary"): Set timR = CreateObject("Scripting.Dictionary")
    Set timBody = CreateObject("Scripting.Dictionary"): Set pedH = CreateObject("Scripting.Dictionary")
    For pass = 0 To 1                                   ' 0 = Spreader tab (top plate), 1 = Spreader bottom tab
        Set rg = Nothing
        If pass = 0 Then
            Set rg = PT_Rng(wb, "Spr_Chips")
        Else
            Set wsB = PT_Sheet(wb, "Spreader bottom")
            If Not wsB Is Nothing Then
                On Error Resume Next
                Set rg = wsB.Names("Spr_Chips").RefersToRange
                On Error GoTo 0
            End If
        End If
        If Not rg Is Nothing Then
            v = rg.Value
            For r = 1 To UBound(v, 1)
                key = UCase(PT_Str(v(r, 1 + P_REF)))
                If key <> "" And UCase(PT_Str(v(r, 1 + P_COV))) = "YES" And PT_Str(v(r, 1 + P_CHECK)) = "OK" And Not timK.Exists(key) Then
                    timK.Add key, PT_Num(v(r, 1 + P_TIMK), 0)
                    timR.Add key, PT_Num(v(r, 1 + P_TIMR), 0)
                    timBody.Add key, (UCase(PT_Str(v(r, 1 + P_TIMMODE))) = "BODY")
                    pedH.Add key, PT_Num(v(r, 1 + P_PED), 0)
                End If
            Next
        End If
    Next
End Sub

' ================================================================= two-resistor chips (M8)
' Makes the part file of every two-resistor chip a stack of three bodies (bottom layer, junction core, top layer),
' and turns a chip that is no longer two-resistor back into one box. True if any part file was changed.
Function EnsureLayers() As Boolean
    Dim it As Variant, c As Object, md As Object, refU As String, bx As Variant, lx As Double, wy As Double, hz As Double
    Dim tt As Double, have As String, want As String, errC As Long, warns As Long, p As Variant, changed As Boolean
    For Each it In tag
        If it(0) = "Part" And it(2) <> "" Then
            refU = UCase(it(2))
            Set c = it(3)
            Set md = c.GetModelDoc2
            have = PT_GetProp(md, "Layers_um")
            If twoRes.Exists(refU) Or (have <> "" And have <> "-") Then
                lx = PT_PropNum(md, "L_um") / 1000#
                wy = PT_PropNum(md, "W_um") / 1000#
                hz = PT_PropNum(md, "H_um") / 1000#
                If lx <= 0 Or wy <= 0 Or hz <= 0 Then
                    bx = PT_CompBox(c)
                    If Not IsEmpty(bx) Then lx = bx(3): wy = bx(4): hz = bx(5)
                End If
                If lx > 0 And wy > 0 And hz > 0 Then
                    If twoRes.Exists(refU) Then
                        tt = Round(layerFrac * hz, 3)
                        want = PT_Um(tt) & "|" & PT_Um(tt)
                        If have <> want Then
                            If PT_BuildLayeredBox(md, lx, wy, hz, tt, tt, TopFace(it(2)), "Bottom") Then
                                md.Save3 PT_swSaveSilent, errC, warns
                                changed = True
                            Else
                                notes.Add refU & ": the three layers could not be made (" & PT_LastErr & "); built as a calibrated block."
                                PT_BuildBox md, lx, wy, hz, TopFace(it(2)), "Bottom"
                                md.Save3 PT_swSaveSilent, errC, warns
                                changed = True
                                twoRes.Remove refU
                            End If
                        End If
                        If twoRes.Exists(refU) Then
                            p = twoRes(refU)
                            twoRes(refU) = Array(p(0), p(1), tt, lx * wy)
                        End If
                    Else
                        PT_BuildBox md, lx, wy, hz, TopFace(it(2)), "Bottom"
                        PT_SetProp md, "Layers_um", "-"
                        md.Save3 PT_swSaveSilent, errC, warns
                        changed = True
                        notes.Add refU & ": layers removed, the chip model is Block again."
                    End If
                End If
            End If
        End If
    Next
    EnsureLayers = changed
End Function

' ================================================================= configuration
Sub SetConfiguration(ByVal asm As Object, ByVal inclNonHeat As Boolean)
    Dim names As Variant, i As Long, hasCfg As Boolean, cur As String, lst As Collection, it As Variant
    names = asm.GetConfigurationNames
    For i = 0 To UBound(names)
        If names(i) = CFG_HEATONLY Then hasCfg = True
    Next
    cur = asm.ConfigurationManager.ActiveConfiguration.Name
    If inclNonHeat Then
        If cur = CFG_HEATONLY Then
            For i = 0 To UBound(names)
                If names(i) <> CFG_HEATONLY Then asm.ShowConfiguration2 names(i): Exit For
            Next
        End If
        Exit Sub
    End If
    If Not hasCfg Then asm.ConfigurationManager.AddConfiguration2 CFG_HEATONLY, "Made by SetupSim: parts without heat suppressed", "", 0, "", "", True
    asm.ShowConfiguration2 CFG_HEATONLY
    ' in this configuration: parts with heat active, parts without heat suppressed
    Set lst = AllTagged(asm)
    For Each it In lst
        If it(0) = "Part" Then
            If heatQ.Exists(UCase(it(2))) Then it(3).SetSuppression2 PT_swFullyResolved Else it(3).SetSuppression2 PT_swSuppressed
        End If
    Next
    notes.Add "Configuration " & CFG_HEATONLY & " is active: parts without heat are suppressed there."
End Sub

' like PT_ListTagged but includes suppressed components (their tags read by resolving them for a moment)
Function AllTagged(ByVal asm As Object) As Collection
    Dim comps As Variant, i As Long, c As Object, md As Object, res As Collection, supp As Boolean
    Set res = New Collection
    comps = asm.GetComponents(True)
    If Not IsEmpty(comps) Then
        For i = 0 To UBound(comps)
            Set c = comps(i)
            supp = (c.GetSuppression2 = PT_swSuppressed)
            If supp Then c.SetSuppression2 PT_swFullyResolved
            Set md = c.GetModelDoc2
            res.Add Array(PT_GetProp(md, "Role"), PT_GetProp(md, "Key"), PT_GetProp(md, "RefDes"), c)
            If supp Then c.SetSuppression2 PT_swSuppressed
        Next
    End If
    Set AllTagged = res
End Function

' ================================================================= materials
Sub ApplyMaterials(ByVal study As Object, ByVal kParts As Double, ByVal kPcbXY As Double, ByVal kPcbZ As Double, _
                   ByVal kSpr As Double, ByVal densSpr As Double)
    Dim sm As Object, i As Long, j As Long, errC As Long, scomp As Object, body As Object, info As Variant
    Dim refU As String, roleS As String, c As Object
    Set sm = study.SolidManager
    For i = 0 To sm.ComponentCount - 1
        Set scomp = sm.GetComponentAt(i, errC)
        If Not scomp Is Nothing Then
            info = TagOfSimComp(scomp.ComponentName)
            If IsEmpty(info) Then
                roleS = "": refU = "": Set c = Nothing
            Else
                roleS = info(0): refU = UCase(info(2)): Set c = info(3)
            End If
            For j = 0 To scomp.SolidBodyCount - 1
                Set body = scomp.GetSolidBodyAt(j, errC)
                If Not body Is Nothing Then
                    Select Case roleS
                        Case "PCB"
                            SetMat body, "PT_PCB", kPcbXY, kPcbZ, 1850, c, scomp.ComponentName
                        Case "Spreader"
                            If InStr(LCase(scomp.ComponentName), "bottom") > 0 Then
                                SetMat body, "PT_SpreaderBottom", kSprBot, kSprBot, densSprBot, Nothing, scomp.ComponentName
                            Else
                                SetMat body, "PT_Spreader", kSpr, kSpr, densSpr, Nothing, scomp.ComponentName
                            End If
                        Case "TIM"
                            If timK.Exists(refU) Then
                                If timK(refU) > 0 Then
                                    SetMat body, "PT_TIM_" & refU, timK(refU), timK(refU), 2500, Nothing, scomp.ComponentName
                                Else
                                    problems.Add "TIM k of " & refU & " is missing on the Spreader tab"
                                End If
                            Else
                                problems.Add "TIM part " & scomp.ComponentName & " has no covered chip row on the Spreader tab"
                            End If
                        Case "Part"
                            If twoRes.Exists(refU) And scomp.SolidBodyCount = 3 Then
                                SetLayerMat body, refU, j, scomp.ComponentName
                            ElseIf chipKz.Exists(refU) Then
                                SetMat body, "PT_Chip_" & refU, chipKxy(refU), chipKz(refU), 2000, c, scomp.ComponentName
                            Else
                                If heatQ.Exists(refU) Then notes.Add refU & ": no calibrated k on Sim setup, generic k used"
                                SetMat body, "PT_Part", kParts, kParts, 2000, Nothing, scomp.ComponentName
                            End If
                        Case Else
                            notes.Add scomp.ComponentName & " is not tagged by the macros: generic k used"
                            SetMat body, "PT_Part", kParts, kParts, 2000, Nothing, scomp.ComponentName
                    End Select
                End If
            Next
        End If
    Next
End Sub

' Material of one body of a two-resistor chip. Layer by body name (feature names PT_Box, PT_LayerCore, PT_LayerTop),
' by body order when the names are something else (bottom, core, top).
Sub SetLayerMat(ByVal body As Object, ByVal refU As String, ByVal idx As Long, ByVal compName As String)
    Dim p As Variant, nm As String, layer As Long, k As Double
    p = twoRes(refU)
    nm = UCase(body.SolidBodyName)
    If InStr(nm, "LAYERCORE") > 0 Then
        layer = 1
    ElseIf InStr(nm, "LAYERTOP") > 0 Then
        layer = 2
    ElseIf InStr(nm, "PT_BOX") > 0 Then
        layer = 0
    Else
        layer = idx
        notes.Add refU & ": body name """ & body.SolidBodyName & """ not recognised, layer taken from body order."
    End If
    Select Case layer
        Case 0
            k = (p(2) / 1000#) / (p(1) * p(3) * 0.000001)
            SetMat body, "PT_Chip_" & refU & "_Bottom", k, k, 2000, Nothing, compName
            notes.Add refU & ": bottom layer " & Round(p(2), 3) & " mm, k " & Round(k, 3) & " W/m K (theta jb " & p(1) & " C/W)."
        Case 1
            SetMat body, "PT_Chip_" & refU & "_Core", kCore, kCore, 2000, Nothing, compName
        Case Else
            k = (p(2) / 1000#) / (p(0) * p(3) * 0.000001)
            SetMat body, "PT_Chip_" & refU & "_Top", k, k, 2000, Nothing, compName
            notes.Add refU & ": top layer " & Round(p(2), 3) & " mm, k " & Round(k, 3) & " W/m K (theta jc " & p(0) & " C/W)."
    End Select
End Sub

' Tag of the assembly component whose name matches a Simulation component name, or Empty
Function TagOfSimComp(ByVal simName As String) As Variant
    Dim it As Variant, nm As String, i As Long, kLast As String
    If tagIdx Is Nothing Then
        ' one pass over the components instead of one per Simulation component (each Name2 is a COM call)
        Set tagIdx = CreateObject("Scripting.Dictionary")
        For Each it In tag
            i = i + 1
            nm = LCase(it(3).Name2)
            If Not tagIdx.Exists(nm) Then tagIdx.Add nm, i
            kLast = Mid(nm, InStrRev(nm, "/") + 1)
            If Not tagIdx.Exists(kLast) Then tagIdx.Add kLast, i
        Next
    End If
    nm = LCase(simName)
    If Not tagIdx.Exists(nm) Then Exit Function
    TagOfSimComp = tag(tagIdx(nm))
End Function

' Isotropic when kxy = kz or no component is given; otherwise orthotropic with Z normal to the part's Front plane.
Sub SetMat(ByVal body As Object, ByVal matName As String, ByVal kxy As Double, ByVal kz As Double, ByVal dens As Double, _
           ByVal c As Object, ByVal compName As String)
    Dim mat As Object, errC As Long, pl As Object
    Set mat = body.GetDefaultMaterial
    If mat Is Nothing Then problems.Add "No material object for " & compName: Exit Sub
    mat.MaterialUnits = swsUnitSI
    mat.MaterialName = matName
    If c Is Nothing Or Abs(kxy - kz) < 0.000001 Then
        mat.ModelType = 0
        mat.SetPropertyByName2 "KX", kxy, False
    Else
        Set pl = CompFrontPlane(c)
        mat.ModelType = 1
        mat.SetPropertyByName2 "KX", kxy, False
        mat.SetPropertyByName2 "KY", kxy, False
        mat.SetPropertyByName2 "KZ", kz, False
        errC = -1
        If Not pl Is Nothing Then mat.SetReferencePlane pl, errC
        If errC <> 0 Then
            ' fall back: isotropic with the through-thickness value (conservative)
            mat.ModelType = 0
            mat.SetPropertyByName2 "KX", kz, False
            notes.Add compName & ": orthotropic direction could not be set, isotropic k = " & Round(kz, 3) & " used"
        End If
    End If
    mat.SetPropertyByName2 "DENS", dens, False
    mat.SetPropertyByName2 "C", 900, False
    errC = body.SetSolidBodyMaterial(mat)
    If errC <> 0 Then problems.Add "Material for " & compName & " was refused (error " & errC & ")"
End Sub

' first reference plane of a component (Front = XY), assembly context
Function CompFrontPlane(ByVal c As Object) As Object
    Dim f As Object
    On Error Resume Next
    Set f = c.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then Set CompFrontPlane = f: Exit Function
        Set f = f.GetNextFeature
    Loop
End Function

' ================================================================= heat
Function ApplyHeat(ByVal study As Object) As Double
    Dim lbc As Object, it As Variant, refU As Variant, bodies As Collection, b As Variant, i As Long
    Dim hp As Object, errC As Long, q As Double, total As Double, arr As Variant
    Set lbc = study.LoadsAndRestraintsManager
    For Each refU In heatQ.Keys
        q = heatQ(refU)
        Set bodies = New Collection
        For Each it In tag
            If it(0) = "Part" And UCase(it(2)) = refU Then
                If twoRes.Exists(refU) Then b = PT_CoreBodies(it(3)) Else b = PT_CompBodies(it(3))
                If Not IsEmpty(b) Then
                    For i = 0 To UBound(b): bodies.Add b(i): Next
                End If
            End If
        Next
        If bodies.Count = 0 Then
            problems.Add refU & " has " & Round(q, 3) & " W in the register but no body with that RefDes in the assembly " & _
                         "(set RefDes and Include on CAD geometry, run BuildBoard)"
        Else
            arr = PT_ToArray(bodies)
            Set hp = lbc.AddHeatPower((arr), errC)
            If hp Is Nothing Or errC <> 0 Then
                problems.Add "Heat power on " & refU & " was refused (error " & errC & ")"
            Else
                hp.HeatPowerBeginEdit
                hp.Unit = swsUnitSI
                hp.HPValue = q / bodies.Count          ' the value is applied to each body
                errC = hp.HeatPowerEndEdit
                If errC <> 0 Then problems.Add "Heat power on " & refU & " could not be set (error " & errC & ")" Else total = total + q
            End If
        End If
    Next
    ApplyHeat = total
End Function

' ================================================================= contacts
Sub ApplyContacts(ByVal study As Object)
    Dim cm As Object
    Dim it As Variant, c As Object, bx As Variant, refU As String, fA As Object, fB As Object, spr As Object, tims As Object
    Dim sprBot As Object, sprUse As Object, chipBot As Boolean
    Set cm = study.ContactManager
    cm.SetGlobalContact swsBondedThermal, 0

    Dim pcb As Object
    Set tims = CreateObject("Scripting.Dictionary")
    For Each it In tag
        If it(0) = "PCB" Then Set pcb = it(3)
        If it(0) = "Spreader" Then
            If InStr(LCase(it(3).Name2), "bottom") > 0 Then Set sprBot = it(3) Else Set spr = it(3)
        End If
        If it(0) = "TIM" Then
            If Not tims.Exists(UCase(it(2))) Then tims.Add UCase(it(2)), it(3)
        End If
    Next

    ' part to PCB
    If rPart > 0 And Not pcb Is Nothing Then
        For Each it In tag
            If it(0) = "Part" Then
                Set c = it(3)
                bx = PT_CompBox(c)
                If Not IsEmpty(bx) Then
                    If Abs(bx(2) - pcbT) < 0.01 Then
                        ResistanceSet cm, PT_AsmFace(c, "Bottom"), PT_AsmFace(pcb, "PCB_Top"), rPart, swsResistDistributed, it(1) & " to PCB"
                    ElseIf Abs(bx(2) + bx(5)) < 0.01 Then
                        ResistanceSet cm, PT_AsmFace(c, TopFace(it(2))), PT_AsmFace(pcb, "PCB_Bottom"), rPart, swsResistDistributed, it(1) & " to PCB"
                    End If
                End If
            End If
        Next
    End If

    ' chips under the spreader
    For Each it In tag
        If it(0) = "Part" And it(2) <> "" Then
            refU = UCase(it(2))
            If timK.Exists(refU) Then
                ' a chip hanging under the PCB touches the bottom plate with its Bottom face; a chip on top touches the top plate with its top face
                bx = PT_CompBox(it(3))
                chipBot = False
                If Not IsEmpty(bx) Then chipBot = (Abs(bx(2) + bx(5)) < 0.01)
                If chipBot Then
                    Set fA = PT_AsmFace(it(3), "Bottom")
                    Set sprUse = sprBot
                Else
                    Set fA = PT_AsmFace(it(3), TopFace(it(2)))
                    Set sprUse = spr
                End If
                If timBody(refU) Then
                    If Not tims.Exists(refU) Then
                        problems.Add refU & ": TIM mode Body but no TIM part in the assembly (run BuildSpreader)"
                    Else
                        If rChipTim > 0 Then ResistanceSet cm, fA, PT_AsmFace(tims(refU), "TIM_" & PT_SafeFile(it(2)) & "_Bottom"), _
                                                            rChipTim, swsResistDistributed, refU & " chip to TIM"
                        If rTimSpr > 0 Then ResistanceSet cm, PT_AsmFace(tims(refU), "TIM_" & PT_SafeFile(it(2)) & "_Top"), _
                                                          SpreaderContact(sprUse, it(2)), rTimSpr, swsResistDistributed, refU & " TIM to spreader"
                    End If
                Else
                    ' Contact R: the TIM is a resistance between chip top and spreader
                    If timR(refU) > 0 Then
                        ResistanceSet cm, fA, SpreaderContact(sprUse, it(2)), timR(refU), swsResistTotal, refU & " TIM as contact resistance"
                    Else
                        problems.Add refU & ": TIM mode Contact R but TIM R on the Spreader tab is empty"
                    End If
                End If
            End If
        End If
    Next
End Sub

Function TopFace(ByVal refD As String) As String
    If refD <> "" Then TopFace = "Chip_" & PT_SafeFile(refD) & "_Top" Else TopFace = "Top"
End Function

Function SpreaderContact(ByVal spr As Object, ByVal refD As String) As Object
    If spr Is Nothing Then Exit Function
    If pedH.Exists(UCase(refD)) Then
        If pedH(UCase(refD)) > 0.0005 Then Set SpreaderContact = PT_AsmFace(spr, "Contact_" & PT_SafeFile(refD)): Exit Function
    End If
    Set SpreaderContact = PT_AsmFace(spr, "Underside")
End Function

Sub ResistanceSet(ByVal cm As Object, ByVal fSrc As Object, ByVal fTgt As Object, ByVal rValue As Double, _
                  ByVal rType As Long, ByVal what As String)
    Dim cs As Object, errC As Long, a1 As Variant, a2 As Variant
    If fSrc Is Nothing Or fTgt Is Nothing Then
        problems.Add "Contact resistance " & what & ": a named face is missing (run BuildBoard / BuildSpreader again)"
        Exit Sub
    End If
    a1 = Array(fSrc): a2 = Array(fTgt)
    Set cs = cm.CreateContactSet2(swsResistanceSet, 0, (a1), (a2), errC)
    If cs Is Nothing Or errC <> 0 Then
        problems.Add "Contact resistance " & what & " was refused (error " & errC & ")"
        Exit Sub
    End If
    cs.ContactSetBeginEdit
    cs.ResistanceType = rType
    cs.ResistanceUnit = swsUnitSI
    cs.ResistanceValue = rValue
    errC = cs.ContactSetEndEdit
    If errC <> 0 Then problems.Add "Contact resistance " & what & " could not be set (error " & errC & ")"
End Sub

' ================================================================= convection and radiation
Sub ApplySurfaces(ByVal study As Object, ByVal ambC As Double)
    Dim g As Long, groups(4) As Collection, it As Variant, c As Object, md As Object
    Dim faces As Collection, fc As Object, nm As String, refU As String, bx As Variant, onTop As Boolean, onBottom As Boolean
    For g = 0 To 4: Set groups(g) = New Collection: Next
    For Each it In tag
        Set c = it(3)
        Set md = c.GetModelDoc2
        Set faces = PT_AllFaces(md)
        refU = UCase(it(2))
        Select Case it(0)
            Case "Spreader"
                For Each fc In faces
                    nm = md.GetEntityName(fc)
                    If nm = "Top" Then
                        groups(0).Add c.GetCorrespondingEntity(fc)
                    ElseIf Left(nm, 8) <> "Contact_" Then
                        groups(1).Add c.GetCorrespondingEntity(fc)
                    End If
                Next
            Case "PCB"
                For Each fc In faces
                    If md.GetEntityName(fc) = "PCB_Top" Then groups(2).Add c.GetCorrespondingEntity(fc) Else groups(3).Add c.GetCorrespondingEntity(fc)
                Next
            Case "Part"
                bx = PT_CompBox(c)
                onTop = False: onBottom = False
                If Not IsEmpty(bx) Then
                    onTop = (Abs(bx(2) - pcbT) < 0.01)              ' sits on the PCB top: its bottom touches the board
                    onBottom = (Abs(bx(2) + bx(5)) < 0.01)          ' hangs under the PCB: its top touches the board
                End If
                For Each fc In faces
                    nm = md.GetEntityName(fc)
                    If twoRes.Exists(refU) And InnerFace(fc, bx) Then
                        ' face between two layers of the chip: touches the next layer
                    ElseIf nm = "Bottom" And onTop Then
                    ElseIf nm = TopFace(it(2)) And (onBottom Or timK.Exists(refU)) Then
                    Else
                        groups(4).Add c.GetCorrespondingEntity(fc)
                    End If
                Next
            Case "TIM"
                For Each fc In faces
                    nm = md.GetEntityName(fc)
                    If Right(nm, 4) <> "_Top" And Right(nm, 7) <> "_Bottom" Then groups(4).Add c.GetCorrespondingEntity(fc)
                Next
        End Select
    Next
    For g = 0 To 4
        If groups(g).Count > 0 Then SurfaceLoads study, groups(g), surfH(g), surfE(g), ambC, surfN(g)
    Next
    notes.Add "PCB top and bottom faces get convection over their whole area, including where parts sit (small error)."
End Sub

' True for a flat horizontal face of a chip part that is neither its bottom nor its top (part frame: z from 0 to H).
Function InnerFace(ByVal fc As Object, ByVal bx As Variant) As Boolean
    Dim nv As Variant, bb As Variant, z As Double
    If IsEmpty(bx) Then Exit Function
    On Error Resume Next
    nv = fc.Normal
    bb = fc.GetBox
    On Error GoTo 0
    If IsEmpty(nv) Or IsEmpty(bb) Then Exit Function
    If Abs(nv(2)) < 0.99 Then Exit Function
    If Abs(bb(5) - bb(2)) * 1000# > 0.001 Then Exit Function
    z = bb(2) * 1000#
    InnerFace = (z > 0.001 And z < bx(5) - 0.001)
End Function

Sub SurfaceLoads(ByVal study As Object, ByVal faces As Collection, ByVal h As Double, ByVal eps As Double, _
                 ByVal ambC As Double, ByVal groupName As String)
    Dim lbc As Object, arr As Variant, cv As Object, rd As Object, errC As Long
    Set lbc = study.LoadsAndRestraintsManager
    arr = PT_ToArray(faces)
    If h > 0 Then
        Set cv = lbc.AddConvection((arr), errC)
        If cv Is Nothing Or errC <> 0 Then
            problems.Add "Convection on " & groupName & " was refused (error " & errC & ")"
        Else
            cv.ConvectionBeginEdit
            cv.Unit = swsUnitSI
            cv.ConvectionCoefficient = h
            cv.BulkAmbientTemperature = ambC + 273.15          ' SI units: kelvin
            errC = cv.ConvectionEndEdit
            If errC <> 0 Then problems.Add "Convection on " & groupName & " could not be set (error " & errC & ")"
        End If
    End If
    If eps > 0 Then
        Set rd = lbc.AddRadiation(swsSurfaceToAmbient, (arr), errC)
        If rd Is Nothing Or errC <> 0 Then
            problems.Add "Radiation on " & groupName & " was refused (error " & errC & ")"
        Else
            rd.RadiationBeginEdit
            rd.Unit = swsUnitSI
            rd.Emmisivity = eps
            rd.ViewFactor = 1
            rd.AmbientTemperature = ambC + 273.15
            errC = rd.RadiationEndEdit
            If errC <> 0 Then problems.Add "Radiation on " & groupName & " could not be set (error " & errC & ")"
        End If
    End If
End Sub

' ================================================================= mesh
Sub MeshStudy(ByVal study As Object, ByVal meshG As Double, ByVal meshF As Double)
    Dim msh As Object, fine As Collection, it As Variant, mc As Object, errC As Long, arr As Variant
    Dim el As Double, tl As Double, tol As Double
    Set msh = study.Mesh
    msh.Quality = swsHighQuality
    Set fine = New Collection
    For Each it In tag
        If it(0) = "TIM" Then fine.Add it(3)
        If it(0) = "Part" Then
            If heatQ.Exists(UCase(it(2))) Then fine.Add it(3)
        End If
    Next
    If fine.Count > 0 Then
        arr = PT_ToArray(fine)
        Set mc = msh.ApplyMeshControl((arr), errC)
        If mc Is Nothing Or errC <> 0 Then
            notes.Add "Fine mesh control on chips and TIMs was refused (error " & errC & "); global size used"
        Else
            mc.MeshControlBeginEdit
            mc.Units = swsMM
            mc.ElementSize = meshF
            mc.MeshControlEndEdit
        End If
    End If
    msh.GetDefaultElementSizeAndTolerance swsMM, el, tl
    If el > 0 Then tol = tl * meshG / el Else tol = meshG * 0.05
    errC = study.CreateMesh(swsMM, meshG, tol)
    If errC <> 0 Then problems.Add "Meshing failed (error " & errC & "). Try a smaller global or fine mesh size on Sim setup."
End Sub

' ================================================================= report
Sub WriteReport(ByVal rptPath As String, ByVal asmPath As String, ByVal runName As String, ByVal applied As Double, _
                ByVal totalReg As Double, ByVal ran As Boolean)
    Dim fso As Object, ts As Object, it As Variant, refU As Variant
    On Error GoTo done
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(rptPath, True)
    ts.WriteLine "SetupSim report " & Format(Now, "yyyy-mm-dd hh:nn")
    ts.WriteLine "Assembly: " & asmPath
    ts.WriteLine "Study: " & runName & IIf(ran, "  (meshed and solved)", "  (NOT solved)")
    ts.WriteLine "Heat applied " & applied & " W, register total " & totalReg & " W"
    ts.WriteLine ""
    ts.WriteLine "HEAT PER REFDES"
    For Each refU In heatQ.Keys: ts.WriteLine "  " & refU & ": " & heatQ(refU) & " W": Next
    ts.WriteLine ""
    ts.WriteLine "PROBLEMS (" & problems.Count & ")"
    For Each it In problems: ts.WriteLine "  " & it: Next
    ts.WriteLine ""
    ts.WriteLine "NOTES"
    For Each it In notes: ts.WriteLine "  " & it: Next
    ts.Close
done:
End Sub

' ================================================================= progress
' SolidWorks status bar progress with a title, so a run of several minutes does not look frozen.
Sub Progress(ByVal txt As String, ByVal pos As Long)
    Dim pb As SldWorks.UserProgressBar
    On Error Resume Next
    If pbar Is Nothing Then
        swApp.GetUserProgressBar pb
        Set pbar = pb
        If Not pbar Is Nothing Then pbar.Start 0, 100, txt
    Else
        pbar.UpdateTitle txt
    End If
    If Not pbar Is Nothing Then pbar.UpdateProgress pos
End Sub

Sub ProgressEnd()
    On Error Resume Next
    If Not pbar Is Nothing Then pbar.End
    Set pbar = Nothing
End Sub
