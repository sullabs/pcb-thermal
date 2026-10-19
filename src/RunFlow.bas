Attribute VB_Name = "RunFlow"
Option Explicit
' =====================================================================================
' RunFlow  -  SolidWorks 2023 macro (milestone M7, Flow Simulation)
'
' Works on a Flow Simulation project that the user made once by hand on the thermal
' assembly. The macro does not build the project. It:
'   1. sets the heat of every Flow volume heat source from the Component register
'      (the source name must contain the RefDes, for example "U1" or "Q U1"),
'   2. sets the initial temperature to the ambient on Sim setup,
'   3. optionally solves,
'   4. reads the goals and writes them like ReadResults does:
'        Sim results tab     one new row (peak and average per RefDes, spreader, PCB)
'        Thermal solution    column AF (Sim Tj) = peak per RefDes
'   5. writes a report CAD\<prefix>RunFlow_<run>.txt
'
' Goal names: a temperature goal whose name contains a RefDes is that chip's goal.
' Add "max", "min" or "avg" to the name (no word means max). A goal whose name contains
' "spreader" or "pcb" is read as that part's max or min.
'
' Project folder (page, W7): pick <prefix>_project.json at the start. Heat per RefDes and the ambient come from
' forSolidWorks, the study type from simSetup (must be Flow); the run (FL_date_time) with a copy of the inputs goes to
' <prefix>_results.json, tab 9 of the page.
' Needs modules PTCommon and PTJson in the same macro. Old workbooks: a copy of PCB_thermal_TEMPLATE.xlsx.
' Needs SOLIDWORKS Flow Simulation turned on (Tools > Add-Ins) and its API SDK installed.
' =====================================================================================

Const MACRO_NAME As String = "RunFlow"
Const FLOW_PRODUCT As String = "Flow Simulation"
Const FLOW_VERSION As String = "2023"
Const FLOW_SRV_CLSID As String = "new:{478B60C3-62E1-4643-A9E0-8A448E7C3E68}"

' nikFeaturesTypes_e
Const nikVolumeSource As Long = 4
' nikParametersTypes_e
Const nikTemperature As Long = 10
Const nikInitialSolidTemperature As Long = 26
Const nikEnvRadiativeTemperature As Long = 30
Const nikHeatGenerationRate As Long = 21
Const nikVolumeSourceType As Long = 43
' nikVolumeSourceTypes_e
Const nikVolumeSourceType_HeatGenerationRate As Long = 1
' nikParameterDependencyTypes_e
Const nikDependencyConstant As Long = 1

Dim notes As Collection
Dim lines As Collection
Dim reWords As Object
Dim jsonMode As Boolean
Dim rptDir As String                     ' output folder and file prefix of the report
Dim regTotal As Double                   ' heat of all parts (W), for the messages

Sub main()
    Set swApp = Application.SldWorks
    PT_Trace MACRO_NAME & ": started (VBA macro)."
    Set notes = New Collection
    Set lines = New Collection

    jsonMode = False: regTotal = 0
    Dim xl As Object, wb As Object, rf As Object, rfPath As String, cad As Object
    Select Case PJ_PickSource(MACRO_NAME)
        Case "json"
            jsonMode = True
        Case "xlsx"
        Case Else
            Exit Sub
    End Select

    ' ---------- heat per RefDes: forSolidWorks (page) or the register (workbook) ----------
    Dim heat As Object, rv As Variant, qv As Variant, i As Long, rk As String, totalQ As Double, hj As Object, hk0 As Variant
    Set heat = CreateObject("Scripting.Dictionary")
    If jsonMode Then
        rfPath = PJ_Dir & "\" & PJ_File("results")
        Set rf = PJ_ReadOwn(PJ_File("results"), "pcb-thermal/results")
        Set cad = PJ_ReadOwn(PJ_File("cad"), "pcb-thermal/cad")
        If rf Is Nothing Or cad Is Nothing Then
            PT_Msg "RunFlow did not run. The results or cad file of the project could not be read:" & vbCrLf & PT_LastErr, vbExclamation, MACRO_NAME
            Exit Sub
        End If
        If UCase(PJ_Str(PJ_Project, "simSetup/studyType", "Thermal")) <> "FLOW" Then
            PT_Msg "RunFlow did not run. The study type on tab 8 of the page is " & PJ_Str(PJ_Project, "simSetup/studyType", "Thermal") & "." & vbCrLf & _
                   "Set it to Flow, save, and run again.", vbExclamation, MACRO_NAME
            Exit Sub
        End If
        If PJ_RefuseIfUnchecked(cad, MACRO_NAME) Then Exit Sub
        Set hj = PJ_Obj(PJ_Project, "forSolidWorks/heat")
        If Not hj Is Nothing Then
            For Each hk0 In hj.Keys
                rk = UCase(CStr(hk0))
                If Not heat.Exists(rk) Then heat.Add rk, 0#
                heat(rk) = heat(rk) + PJ_Num(hj, CStr(hk0), 0)
            Next
        End If
        rptDir = PJ_OutputFolder() & PT_SafeFile(PJ_Prefix())
    Else
        Set xl = PT_GetExcel()
        Set wb = PT_PickWorkbook(xl, MACRO_NAME, "Sim setup|Sim results|Thermal solution|Component register", PJ_WbPath)
        If wb Is Nothing Then Exit Sub
        If Not PT_NeedNames(wb, "Sim_Type|Sim_Ambient|Sim_Surfaces|Res_Header|TS_SimTj|Reg_RefDes|Reg_QUsed|" & _
                                "Reg_TotalQ|Spr_Chips|Set_Prefix|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
        xl.Calculate
        If UCase(PT_NameStr(wb, "Sim_Type", "Thermal")) <> "FLOW" Then
            PT_Msg "RunFlow did not run. Study type on Sim setup is " & PT_NameStr(wb, "Sim_Type", "") & "." & vbCrLf & _
                   "Set it to Flow and run again.", vbExclamation, MACRO_NAME
            Exit Sub
        End If
        If PT_RefuseIfUnchecked(wb, MACRO_NAME) Then Exit Sub
        rv = PT_Rng(wb, "Reg_RefDes").Value
        qv = PT_Rng(wb, "Reg_QUsed").Value
        For i = 1 To UBound(rv, 1)
            rk = UCase(PT_Str(rv(i, 1)))
            If rk <> "" Then
                If Not heat.Exists(rk) Then heat.Add rk, 0#
                heat(rk) = heat(rk) + PT_Num(qv(i, 1), 0)
            End If
        Next
        rptDir = PT_OutputFolder(wb) & PT_SafeFile(PT_Prefix(wb))
    End If
    For Each hk0 In heat.Keys: regTotal = regTotal + heat(hk0): Next
    If heat.Count = 0 Then
        PT_Msg "RunFlow did not run. No part has heat (" & IIf(jsonMode, "tab 1 of the page", "Component register") & ").", vbExclamation, MACRO_NAME
        Exit Sub
    End If

    ' ---------- assembly and Flow project ----------
    Dim asm As Object, asmPath As String
    If jsonMode Then asmPath = PJ_AsmPath() Else asmPath = PT_AsmPath(wb)
    Set asm = PT_OpenAsm(asmPath)
    If asm Is Nothing Then
        PT_Msg "RunFlow did not run. No thermal assembly at" & vbCrLf & asmPath, vbExclamation, MACRO_NAME
        Exit Sub
    End If
    PT_ActivateDoc asm

    Dim why As String, api As Object, ia As Object, md As Object, fda As Object
    Set api = OpenFlowApi(why)
    If api Is Nothing Then
        PT_Msg "RunFlow did not run. " & why, vbExclamation, MACRO_NAME
        Exit Sub
    End If
    On Error GoTo apiFail
    Set ia = AttachFlow(api)
    If ia Is Nothing Then
        PT_Msg "RunFlow did not run. Could not attach to SolidWorks Flow Simulation." & vbCrLf & _
               "Turn on SOLIDWORKS Flow Simulation in Tools > Add-Ins and run again.", vbExclamation, MACRO_NAME
        GoTo done
    End If
    Set md = ia.ActiveDocument
    If md Is Nothing Then
        PT_Msg "RunFlow did not run. Flow Simulation sees no active document." & vbCrLf & _
               "Open the thermal assembly and turn on Flow Simulation.", vbExclamation, MACRO_NAME
        GoTo done
    End If
    If StrComp(CStr(md.GetPathName), asmPath, vbTextCompare) <> 0 Then
        PT_Msg "RunFlow did not run. The active document is" & vbCrLf & CStr(md.GetPathName) & vbCrLf & _
               "but the thermal assembly is" & vbCrLf & asmPath & vbCrLf & "Make the thermal assembly the active window.", _
               vbExclamation, MACRO_NAME
        GoTo done
    End If
    Set fda = md.GetActiveProject
    If fda Is Nothing Then
        PT_Msg "RunFlow did not run. The thermal assembly has no Flow Simulation project." & vbCrLf & _
               "Create one once (Flow Simulation > Wizard), add a volume heat source per chip and temperature goals," & vbCrLf & _
               "then run again. See docs\HOWTO_M7.md.", vbExclamation, MACRO_NAME
        GoTo done
    End If

    ' ---------- heat sources ----------
    Dim en As Object, ft As Object, tk As Long, refU As String, q As Double, nSet As Long, applied As Double
    Dim seen As Object, vr As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    Set en = fda.EnumFeatures
    If Not en Is Nothing Then
        en.Reset
        Set ft = en.Next
        Do While Not ft Is Nothing
            tk = -1
            vr = ft.Type
            If IsNumeric(vr) Then tk = CLng(vr)
            If tk = nikVolumeSource Then
                refU = FindRef(CStr(ft.Name), heat)
                If refU = "" Then
                    notes.Add "Volume source '" & CStr(ft.Name) & "' has no RefDes in its name: left as it is."
                ElseIf seen.Exists(refU) Then
                    notes.Add "Volume source '" & CStr(ft.Name) & "': " & refU & " already has a source. Left as it is."
                Else
                    seen.Add refU, 1
                    q = heat(refU)
                    If SetSourceHeat(ft, fda, q) Then
                        nSet = nSet + 1: applied = applied + q
                        lines.Add "source " & CStr(ft.Name) & " = " & Format(q, "0.000") & " W (" & refU & ")"
                    End If
                End If
            End If
            Set ft = en.Next
        Loop
    End If
    Dim hk As Variant
    For Each hk In heat.Keys
        If heat(hk) > 0 And Not seen.Exists(CStr(hk)) Then
            notes.Add CStr(hk) & " has " & Format(heat(hk), "0.00") & " W of heat but no Flow volume source with " & CStr(hk) & " in its name."
        End If
    Next

    ' ---------- ambient ----------
    Dim amb As Double, nAmb As Long
    Dim hasAmb As Boolean
    If jsonMode Then
        hasAmb = PJ_IsNum(PJ_Project, "forSolidWorks/ambientC")
        If hasAmb Then amb = PJ_Num(PJ_Project, "forSolidWorks/ambientC", 25)
    ElseIf PT_IsNum(PT_Rng(wb, "Sim_Ambient").Cells(1, 1).Value) Then
        hasAmb = True
        amb = CDbl(PT_Rng(wb, "Sim_Ambient").Cells(1, 1).Value)
    End If
    If hasAmb Then
        nAmb = SetInitialTemperature(fda, amb + 273.15)
        If nAmb > 0 Then
            lines.Add "initial temperature = " & Format(amb, "0.0") & " C (" & nAmb & " parameters)"
        Else
            notes.Add "Initial temperature was not set: no temperature parameter found in the initial conditions."
        End If
    Else
        notes.Add "Ambient is not a number: initial temperature left as it is."
    End If

    If nSet = 0 Then
        PT_Msg "RunFlow stopped. No Flow volume heat source was updated." & vbCrLf & vbCrLf & ListNotes(), vbExclamation, MACRO_NAME
        GoTo done
    End If

    ' ---------- solve or read ----------
    Dim runName As String, ans As Long, solved As Boolean, rh As Object
    runName = "FL_" & Format(Now, "yyyymmdd_hhnn")
    ans = PT_Msg(nSet & " heat sources set, " & Format(applied, "0.00") & " W of " & Format(regTotal, "0.00") & " W of heat on the parts." & _
                 vbCrLf & vbCrLf & "Yes = solve now (can take many minutes) and read the goals" & vbCrLf & _
                 "No = do not solve, read the goals of the last solved run" & vbCrLf & "Cancel = stop here", _
                 vbYesNoCancel + vbQuestion, MACRO_NAME)
    If ans = vbCancel Then
        md.Save
        WriteReport runName, CStr(fda.Name), applied, False
        PT_Msg "Sources updated. Nothing solved, nothing written to " & IIf(jsonMode, "the results file.", "Excel.") & vbCrLf & "Report: " & ReportPath(runName), vbInformation, MACRO_NAME
        GoTo done
    End If
    If ans = vbYes Then
        If Not IsTrue(fda.Solve2(True, True, True, False)) Then
            PT_Msg "RunFlow stopped. The Flow solver did not finish." & vbCrLf & ListNotes() & _
                   "Check the project in SolidWorks (Flow Simulation > Run).", vbExclamation, MACRO_NAME
            GoTo done
        End If
        solved = True
    End If
    Set rh = fda.LoadResults(True)
    If rh Is Nothing Then
        PT_Msg "RunFlow stopped. The project has no results to read. Solve it first (Yes at the question).", vbExclamation, MACRO_NAME
        GoTo done
    End If

    ' ---------- goals ----------
    Dim peak As Object, avg As Object, gres As Object, ge As Object, gl As Object
    Dim gname As String, gv As Double, gk As String, gref As String, grp As String
    Dim sprMax As Double, sprMin As Double, pcbMax As Double, pcbMin As Double
    Dim hSprMax As Boolean, hSprMin As Boolean, hPcbMax As Boolean, hPcbMin As Boolean, nGoals As Long
    Set peak = CreateObject("Scripting.Dictionary")
    Set avg = CreateObject("Scripting.Dictionary")
    Set gres = rh.GetGoalsCalculationResults2
    If Not gres Is Nothing Then Set ge = gres.GetGoalsEnum
    If ge Is Nothing Then
        PT_Msg "RunFlow stopped. The results have no goals. Add temperature goals to the Flow project.", vbExclamation, MACRO_NAME
        GoTo done
    End If
    ge.Reset
    Set gl = ge.Next
    Do While Not gl Is Nothing
        gname = CStr(gl.GetGoalName)
        gv = CDbl(gl.GetLastCalculatedValue) - 273.15
        gk = GoalKind(gname)
        gref = FindRef(gname, heat)
        grp = ""
        If gref = "" Then grp = FindGroup(gname)
        nGoals = nGoals + 1
        If (gref <> "" Or grp <> "") And gv < -270 Then
            notes.Add "Goal '" & gname & "' has no valid temperature (0 K): not stored. Check that its body is meshed and solid."
            lines.Add "goal " & gname & " = no value (0 K): skipped"
        ElseIf gref <> "" Then
            If gk = "avg" Then
                avg(gref) = gv
            ElseIf gk = "max" Then
                peak(gref) = gv
            Else
                lines.Add "goal " & gname & " = " & Format(gv, "0.00") & " C (min, not stored)"
            End If
            lines.Add "goal " & gname & " = " & Format(gv, "0.00") & " C -> " & gref & " " & gk
        ElseIf grp = "SPREADER" Then
            If gk = "min" Then
                sprMin = gv: hSprMin = True
            ElseIf gk = "max" Then
                sprMax = gv: hSprMax = True
            End If
            lines.Add "goal " & gname & " = " & Format(gv, "0.00") & " C -> spreader " & gk
        ElseIf grp = "PCB" Then
            If gk = "min" Then
                pcbMin = gv: hPcbMin = True
            ElseIf gk = "max" Then
                pcbMax = gv: hPcbMax = True
            End If
            lines.Add "goal " & gname & " = " & Format(gv, "0.00") & " C -> pcb " & gk
        Else
            lines.Add "goal " & gname & " = " & Format(gv, "0.000") & " (not a RefDes, spreader or PCB goal: ignored)"
        End If
        Set gl = ge.Next
    Loop
    If peak.Count = 0 And avg.Count = 0 Then
        PT_Msg "RunFlow stopped. No goal name contains a RefDes. Nothing was written." & vbCrLf & _
               "Rename the goals, for example 'U1 max' (see docs\HOWTO_M7.md).", vbExclamation, MACRO_NAME
        GoTo done
    End If
    md.Save
    On Error GoTo 0

    ' ---------- results file (page) ----------
    If jsonMode Then
        Dim jr As Object, runs As Object
        Set jr = FlowRun(runName, CStr(fda.Name), solved, asmPath, peak, avg, hSprMax And hSprMin, sprMin, sprMax, hPcbMax And hPcbMin, pcbMin, pcbMax)
        WriteReport runName, CStr(fda.Name), applied, True
        PJ_Put jr, "report", ReportPath(runName)
        Set runs = PJ_Obj(rf, "runs")
        If runs Is Nothing Then Set runs = New Collection
        runs.Add jr
        PJ_Put rf, "runs", runs
        If Not PJ_SaveOwn(rf, PJ_File("results"), MACRO_NAME) Then
            PT_Msg "RunFlow read the goals, but " & rfPath & " could not be written:" & vbCrLf & PT_LastErr, vbExclamation, MACRO_NAME
            GoTo done
        End If
        PT_Msg "Written to:" & vbCrLf & rfPath & vbCrLf & "(run " & runName & ", tab 9 of the page)" & vbCrLf & vbCrLf & _
               "Flow project " & CStr(fda.Name) & ": " & nSet & " heat sources set (" & Format(applied, "0.00") & " W), " & _
               peak.Count & " chips with a peak goal, " & nGoals & " goals read." & vbCrLf & _
               "Report: " & ReportPath(runName) & vbCrLf & ListNotes(), vbInformation, MACRO_NAME
        GoTo done
    End If

    ' ---------- Excel ----------
    Dim hdr As Object, ws As Object, r As Long, c As Long, lastR As Long, sv As Variant, refK As Variant
    On Error GoTo xlFail
    xl.ScreenUpdating = False
    Set hdr = PT_Rng(wb, "Res_Header")
    Set ws = hdr.Worksheet
    lastR = ws.Cells(ws.Rows.Count, hdr.Column).End(-4162).Row
    If lastR < hdr.Row Then lastR = hdr.Row
    r = lastR + 1
    sv = PT_Rng(wb, "Sim_Surfaces").Value
    ws.Cells(r, hdr.Column).Value = runName
    ws.Cells(r, hdr.Column + 1).Value = Format(Now, "yyyy-mm-dd hh:nn")
    ws.Cells(r, hdr.Column + 2).Value = "Flow " & CStr(fda.Name) & IIf(solved, "", " (last run)")
    ws.Cells(r, hdr.Column + 6).Value = PT_Num(sv(1, 2), 0)
    ws.Cells(r, hdr.Column + 7).Value = PT_Num(sv(3, 2), 0)
    ws.Cells(r, hdr.Column + 8).Value = PT_TimModes(wb)
    If hSprMax Then ws.Cells(r, hdr.Column + 9).Value = Round(sprMax, 2)
    If hSprMin Then ws.Cells(r, hdr.Column + 10).Value = Round(sprMin, 2)
    If hPcbMax Then ws.Cells(r, hdr.Column + 11).Value = Round(pcbMax, 2)
    If hPcbMin Then ws.Cells(r, hdr.Column + 12).Value = Round(pcbMin, 2)
    For Each refK In peak.Keys
        c = PT_RefColumn(ws, hdr, CStr(refK))
        ws.Cells(r, c).Value = Round(peak(refK), 2)
        If avg.Exists(refK) Then ws.Cells(r, c + 1).Value = Round(avg(refK), 2)
    Next
    For Each refK In avg.Keys
        If Not peak.Exists(refK) Then
            c = PT_RefColumn(ws, hdr, CStr(refK))
            ws.Cells(r, c + 1).Value = Round(avg(refK), 2)
        End If
    Next

    PT_WriteSimTj wb, peak
    xl.ScreenUpdating = True
    xl.Calculate
    wb.Save
    On Error GoTo 0
    WriteReport runName, CStr(fda.Name), applied, True

    Dim s As String
    s = "Written and saved to:" & vbCrLf & wb.FullName & vbCrLf & "(Sim results row " & r & ", Thermal solution column AF)" & vbCrLf & vbCrLf & _
        "Flow project " & CStr(fda.Name) & ": " & nSet & " heat sources set (" & Format(applied, "0.00") & " W), " & _
        peak.Count & " chips with a peak goal, " & nGoals & " goals read." & vbCrLf & _
        "Report: " & ReportPath(runName) & vbCrLf & ListNotes()
    PT_Msg s, vbInformation, MACRO_NAME
    GoTo done

apiFail:
    PT_Msg "RunFlow stopped in the Flow Simulation API: " & Err.Description & vbCrLf & _
           "Nothing was written to Excel. Check that Flow Simulation is on and the project opens in SolidWorks.", vbExclamation, MACRO_NAME
    GoTo done
xlFail:
    xl.ScreenUpdating = True
    PT_Msg "RunFlow stopped while writing to Excel: " & Err.Description & vbCrLf & _
           "The workbook was not saved. Close any cell you are editing, close the workbook without saving, and run again.", _
           vbExclamation, MACRO_NAME
done:
    On Error Resume Next
    Set gl = Nothing: Set ge = Nothing: Set gres = Nothing: Set rh = Nothing
    Set ft = Nothing: Set en = Nothing: Set fda = Nothing: Set md = Nothing: Set ia = Nothing
    If Not api Is Nothing Then api.UnloadProductAPI
    Set api = Nothing
End Sub

' ---------------------------------------------------------------- Flow API
' Loads the Flow API object. Tries the direct class first, then the out-of-process server
' (needed when the macro host is 32-bit). Returns Nothing and says why.
Function OpenFlowApi(ByRef why As String) As Object
    Dim api As Object, srv As Object, rv As Variant, ok As Boolean, errNum As Long, errTxt As String
    On Error Resume Next
    Set api = CreateObject("NIKCommonApi2.BaseApiObject")
    If api Is Nothing Then
        Err.Clear
        Set srv = GetObject(FLOW_SRV_CLSID)
        If Not srv Is Nothing Then Set api = srv.LoadNIKCommonApiLibrary
    End If
    If api Is Nothing Then
        why = "The Flow Simulation API is not available (" & Err.Description & ")." & vbCrLf & _
              "Install the API SDK: run 'flow simulation api sdk x64.exe' from the api folder of the Flow Simulation install."
        Exit Function
    End If
    Err.Clear
    rv = api.LoadProductAPI2(FLOW_PRODUCT, FLOW_VERSION)
    errNum = Err.Number: errTxt = Err.Description
    Err.Clear
    ok = IsTrue(rv)
    If Not ok Then
        why = "The Flow API could not load " & FLOW_PRODUCT & " " & FLOW_VERSION & " (error " & errNum & ": " & errTxt & ")."
        Set api = Nothing
        Exit Function
    End If
    Set OpenFlowApi = api
End Function

' The Flow API hands back its Boolean results as a raw 1 (not -1), so "Not value" is wrong. Test non-zero instead.
Function IsTrue(ByVal v As Variant) As Boolean
    On Error Resume Next
    IsTrue = (CLng(v) <> 0)
End Function

Function AttachFlow(ByVal api As Object) As Object
    Dim ia As Object
    On Error Resume Next
    Set ia = api.Attach2RunningObject2(CLng(swApp.GetProcessID))
    If ia Is Nothing Then
        Err.Clear
        Set ia = api.Attach2RunningObject
    End If
    Set AttachFlow = ia
End Function

' Volume source: type = Heat generation rate, value = q (W). True when both were accepted.
Function SetSourceHeat(ByVal ft As Object, ByVal fda As Object, ByVal q As Double) As Boolean
    Dim pf As Object, ep As Object, par As Object, rc As Long, gotQ As Boolean, pass As Long, pt As Long
    Set pf = ft.GetInterface("IParametrizedFeature")
    If pf Is Nothing Then
        notes.Add "Source '" & CStr(ft.Name) & "': no parameters available."
        Exit Function
    End If
    For pass = 1 To 2
        Set ep = pf.EnumParameters
        If ep Is Nothing Then Exit For
        ep.Reset
        Set par = ep.Next
        Do While Not par Is Nothing
            pt = CLng(par.Type)
            If pass = 1 And pt = nikVolumeSourceType Then
                rc = CLng(par.SetLongValue(nikVolumeSourceType_HeatGenerationRate))
                If rc <> 0 Then notes.Add "Source '" & CStr(ft.Name) & "': could not set the type to heat generation rate (code " & rc & ")."
            ElseIf pass = 2 And pt = nikHeatGenerationRate Then
                If CLng(par.DependencyType) <> nikDependencyConstant Then
                    notes.Add "Source '" & CStr(ft.Name) & "': heat is a formula or table in the project. Change it to a constant. Left as it is."
                Else
                    rc = CLng(par.SetValue(q))
                    If rc = 0 Then
                        gotQ = True
                    Else
                        notes.Add "Source '" & CStr(ft.Name) & "': heat value refused (code " & rc & ")."
                    End If
                End If
            End If
            Set par = ep.Next
        Loop
    Next
    If Not gotQ Then Exit Function
    If Not IsTrue(ft.Rebuild(fda)) Then notes.Add "Source '" & CStr(ft.Name) & "' did not rebuild: " & CStr(fda.GetLastRebuildError)
    SetSourceHeat = True
End Function

' Sets fluid temperature, initial solid temperature and radiation environment temperature (kelvin). Returns how many.
Function SetInitialTemperature(ByVal fda As Object, ByVal kelvin As Double) As Long
    Dim ic As Object, pf As Object, ep As Object, par As Object, rc As Long, n As Long, pt As Long
    Set ic = fda.GlobalInitialConditions
    If ic Is Nothing Then Exit Function
    Set pf = ic.GetInterface("IParametrizedFeature")
    If pf Is Nothing Then Exit Function
    Set ep = pf.EnumParameters
    If ep Is Nothing Then Exit Function
    ep.Reset
    Set par = ep.Next
    Do While Not par Is Nothing
        pt = CLng(par.Type)
        If pt = nikTemperature Or pt = nikInitialSolidTemperature Or pt = nikEnvRadiativeTemperature Then
            If CLng(par.DependencyType) = nikDependencyConstant Then
                rc = CLng(par.SetValue(kelvin))
                If rc = 0 Then n = n + 1
            End If
        End If
        Set par = ep.Next
    Loop
    SetInitialTemperature = n
End Function

' ---------------------------------------------------------------- names
' Upper-case words of a text (letters and digits only)
Function Words(ByVal txt As String) As Variant
    If reWords Is Nothing Then
        Set reWords = CreateObject("VBScript.RegExp")
        reWords.Global = True
        reWords.Pattern = "[^A-Za-z0-9]+"
    End If
    Words = Split(Trim(reWords.Replace(UCase(txt), " ")), " ")
End Function

' First word of txt that is a RefDes in refs, else ""
Function FindRef(ByVal txt As String, ByVal refs As Object) As String
    Dim w As Variant, i As Long
    w = Words(txt)
    For i = 0 To UBound(w)
        If CStr(w(i)) <> "" Then
            If refs.Exists(CStr(w(i))) Then FindRef = CStr(w(i)): Exit Function
        End If
    Next
End Function

' "SPREADER" or "PCB" when the goal name has that word, else ""
Function FindGroup(ByVal txt As String) As String
    Dim w As Variant, i As Long
    w = Words(txt)
    For i = 0 To UBound(w)
        If CStr(w(i)) = "SPREADER" Or CStr(w(i)) = "PCB" Then FindGroup = CStr(w(i)): Exit Function
    Next
End Function

' "min", "avg" or "max" (default) from the words of a goal name
Function GoalKind(ByVal txt As String) As String
    Dim w As Variant, i As Long
    GoalKind = "max"
    w = Words(txt)
    For i = 0 To UBound(w)
        Select Case CStr(w(i))
            Case "MIN", "MINIMUM": GoalKind = "min": Exit Function
            Case "AVG", "AV", "MEAN", "AVERAGE": GoalKind = "avg": Exit Function
        End Select
    Next
End Function

Function ListNotes() As String
    Dim q As Variant, n As Long, s As String
    For Each q In notes
        n = n + 1
        If n > 6 Then s = s & "- ... more in the report" & vbCrLf: Exit For
        s = s & "- " & q & vbCrLf
    Next
    ListNotes = s
End Function

' ---------------------------------------------------------------- report
Function ReportPath(ByVal runName As String) As String
    ReportPath = rptDir & "RunFlow_" & runName & ".txt"
End Function

Sub WriteReport(ByVal runName As String, ByVal projName As String, ByVal applied As Double, ByVal readGoals As Boolean)
    Dim fso As Object, ts As Object, q As Variant
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(ReportPath(runName), True)
    ts.WriteLine "RunFlow " & runName & "  " & Format(Now, "yyyy-mm-dd hh:nn")
    ts.WriteLine "Flow project: " & projName
    ts.WriteLine "Heat applied: " & Format(applied, "0.000") & " W of " & Format(regTotal, "0.000") & " W of heat on the parts"
    ts.WriteLine "Goals read: " & IIf(readGoals, "yes", "no")
    ts.WriteLine ""
    For Each q In lines: ts.WriteLine q: Next
    ts.WriteLine ""
    For Each q In notes: ts.WriteLine "NOTE " & q: Next
    ts.Close
End Sub

' ---------------------------------------------------------------- results file (page)
' results.json run of this Flow solve: goals per RefDes (a part needs its max goal), PCB and spreader goals, and a
' copy of the inputs. RefDes are written as on the page.
Function FlowRun(ByVal runName As String, ByVal projName As String, ByVal solved As Boolean, ByVal asmPath As String, _
                 ByVal peak As Object, ByVal avg As Object, ByVal hasSpr As Boolean, ByVal sprMin As Double, ByVal sprMax As Double, _
                 ByVal hasPcb As Boolean, ByVal pcbMin As Double, ByVal pcbMax As Double) As Object
    Dim d As Object, inp As Object, parts As Object, p As Object, mm As Object, refK As Variant, refS As String
    Dim refNames As Object, pl As Object, it As Variant, nl As Collection
    Set refNames = CreateObject("Scripting.Dictionary")
    Set pl = PJ_Obj(PJ_Project, "parts")
    If Not pl Is Nothing Then
        For Each it In pl
            refS = PJ_Str(it, "refDes", "")
            If refS <> "" And Not refNames.Exists(UCase(refS)) Then refNames.Add UCase(refS), refS
        Next
    End If
    Set d = PJ_Dict()
    PJ_Put d, "name", runName
    PJ_Put d, "date", PJ_Now()
    PJ_Put d, "studyType", "Flow"
    PJ_Put d, "study", "Flow " & projName & IIf(solved, "", " (last run)")
    PJ_Put d, "status", "read"
    PJ_Put d, "assembly", asmPath
    Set inp = PJ_Dict()
    PJ_Put inp, "forSolidWorks", PJ_Obj(PJ_Project, "forSolidWorks")
    PJ_Put inp, "simSetup", PJ_Obj(PJ_Project, "simSetup")
    PJ_Put d, "inputs", inp
    If hasSpr Then
        Set mm = PJ_Dict(): PJ_Put mm, "maxC", Round(sprMax, 2): PJ_Put mm, "minC", Round(sprMin, 2)
        PJ_Put d, "spreader", mm
    End If
    If hasPcb Then
        Set mm = PJ_Dict(): PJ_Put mm, "maxC", Round(pcbMax, 2): PJ_Put mm, "minC", Round(pcbMin, 2)
        PJ_Put d, "pcb", mm
    End If
    Set parts = PJ_Dict()
    For Each refK In peak.Keys
        If refNames.Exists(CStr(refK)) Then refS = refNames(CStr(refK)) Else refS = CStr(refK)
        Set p = PJ_Dict()
        PJ_Put p, "peakC", Round(peak(refK), 2)
        If avg.Exists(refK) Then PJ_Put p, "avgC", Round(avg(refK), 2)
        PJ_Put parts, refS, p
    Next
    PJ_Put d, "parts", parts
    Set nl = New Collection
    For Each it In avg.Keys
        If Not peak.Exists(it) Then nl.Add CStr(it) & ": only an average goal (" & Format(avg(it), "0.0") & " C). Add a max goal to see it on the page."
    Next
    For Each it In notes: nl.Add CStr(it): Next
    If nl.Count > 0 Then PJ_Put d, "notes", nl
    Set FlowRun = d
End Function
