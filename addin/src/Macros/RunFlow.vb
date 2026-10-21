Option Strict Off
Option Explicit On

Imports System
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic

' Port of src/RunFlow.bas (milestone M7, Flow Simulation). Same behaviour and messages as the VBA macro.
'
' Works on a Flow Simulation project that the user made once by hand on the thermal assembly. It:
'   1. sets the heat of every Flow volume heat source from the Component register (the source name must contain the RefDes),
'   2. sets the initial temperature to the ambient on Sim setup,
'   3. optionally solves,
'   4. reads the goals and writes them like ReadResults does (Sim results row, Thermal solution column AF),
'   5. writes a report CAD\<prefix>RunFlow_<run>.txt
' Goal names: a temperature goal whose name contains a RefDes is that chip's goal; "max", "min" or "avg" in the name
' (no word means max). A goal whose name contains "spreader" or "pcb" is read as that part's max or min.
Public Module RunFlow

    Private Const MACRO_NAME As String = "RunFlow"
    Private Const FLOW_PRODUCT As String = "Flow Simulation"
    Private Const FLOW_VERSION As String = "2023"
    Private Const FLOW_SRV_CLSID As String = "new:{478B60C3-62E1-4643-A9E0-8A448E7C3E68}"

    ' nikFeaturesTypes_e
    Private Const nikVolumeSource As Integer = 4
    ' nikParametersTypes_e
    Private Const nikTemperature As Integer = 10
    Private Const nikInitialSolidTemperature As Integer = 26
    Private Const nikEnvRadiativeTemperature As Integer = 30
    Private Const nikHeatGenerationRate As Integer = 21
    Private Const nikVolumeSourceType As Integer = 43
    ' nikVolumeSourceTypes_e
    Private Const nikVolumeSourceType_HeatGenerationRate As Integer = 1
    ' nikParameterDependencyTypes_e
    Private Const nikDependencyConstant As Integer = 1

    Private notes As Collection
    Private lines As Collection
    Private ReadOnly reWords As New Regex("[^A-Za-z0-9]+")

    Public Sub Main()
        notes = New Collection()
        lines = New Collection()
        Dim api As Object = Nothing, ia As Object = Nothing, md As Object = Nothing, fda As Object = Nothing
        Dim en As Object = Nothing, ft As Object = Nothing, rh As Object = Nothing, gres As Object = Nothing, ge As Object = Nothing, gl As Object = Nothing

        Dim xl As Object, wb As Object
        xl = PT_GetExcel()
        wb = PT_PickWorkbook(xl, MACRO_NAME, "Sim setup|Sim results|Thermal solution|Component register")
        If wb Is Nothing Then Exit Sub
        If Not PT_NeedNames(wb, "Sim_Type|Sim_Ambient|Sim_Surfaces|Res_Header|TS_SimTj|Reg_RefDes|Reg_QUsed|" &
                                "Reg_TotalQ|Spr_Chips|Set_Prefix|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
        xl.Calculate()
        If UCase(PT_NameStr(wb, "Sim_Type", "Thermal")) <> "FLOW" Then
            MsgBox("RunFlow did not run. Study type on Sim setup is " & PT_NameStr(wb, "Sim_Type", "") & "." & vbCrLf &
                   "Set it to Flow and run again.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        If PT_RefuseIfUnchecked(wb, MACRO_NAME) Then Exit Sub

        ' ---------- heat per RefDes from the register ----------
        Dim heat As New Dict(), rv(,) As Object, qv(,) As Object, rk As String
        rv = ToTable(PT_Rng(wb, "Reg_RefDes").Value)
        qv = ToTable(PT_Rng(wb, "Reg_QUsed").Value)
        For i As Integer = 1 To UBound(rv, 1)
            rk = UCase(PT_Str(rv(i, 1)))
            If rk <> "" Then
                If Not heat.Exists(rk) Then heat.Add(rk, 0.0)
                heat(rk) = CDbl(heat(rk)) + PT_Num(qv(i, 1), 0)
            End If
        Next
        If heat.Count = 0 Then
            MsgBox("RunFlow did not run. The Component register has no RefDes.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If

        ' ---------- assembly and Flow project ----------
        Dim asm As Object, asmPath As String
        asmPath = PT_AsmPath(wb)
        asm = PT_OpenAsm(asmPath)
        If asm Is Nothing Then
            MsgBox("RunFlow did not run. No thermal assembly at" & vbCrLf & asmPath, vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        PT_ActivateDoc(asm)

        Dim why As String = ""
        api = OpenFlowApi(why)
        If api Is Nothing Then
            MsgBox("RunFlow did not run. " & why, vbExclamation, MACRO_NAME)
            Exit Sub
        End If

        Dim tk As Integer, refU As String, q As Double, nSet As Integer = 0, applied As Double = 0
        Dim seen As New Dict(), vr As Object
        Dim amb As Double, nAmb As Integer
        Dim runName As String = "", ans As MsgBoxResult, solved As Boolean = False
        Dim peak As New Dict(), avg As New Dict()
        Dim gname As String, gv As Double, gk As String, gref As String, grp As String
        Dim sprMax As Double, sprMin As Double, pcbMax As Double, pcbMin As Double
        Dim hSprMax As Boolean, hSprMin As Boolean, hPcbMax As Boolean, hPcbMin As Boolean, nGoals As Integer = 0
        Dim hdr As Object, ws As Object, r As Integer, c As Integer, lastR As Integer, sv(,) As Object
        Dim s As String

        On Error GoTo apiFail
        ia = AttachFlow(api)
        If ia Is Nothing Then
            MsgBox("RunFlow did not run. Could not attach to SolidWorks Flow Simulation." & vbCrLf &
                   "Turn on SOLIDWORKS Flow Simulation in Tools > Add-Ins and run again.", vbExclamation, MACRO_NAME)
            GoTo done
        End If
        md = ia.ActiveDocument
        If md Is Nothing Then
            MsgBox("RunFlow did not run. Flow Simulation sees no active document." & vbCrLf &
                   "Open the thermal assembly and turn on Flow Simulation.", vbExclamation, MACRO_NAME)
            GoTo done
        End If
        If StrComp(CStr(md.GetPathName), asmPath, vbTextCompare) <> 0 Then
            MsgBox("RunFlow did not run. The active document is" & vbCrLf & CStr(md.GetPathName) & vbCrLf &
                   "but the thermal assembly is" & vbCrLf & asmPath & vbCrLf & "Make the thermal assembly the active window.",
                   vbExclamation, MACRO_NAME)
            GoTo done
        End If
        fda = md.GetActiveProject
        If fda Is Nothing Then
            MsgBox("RunFlow did not run. The thermal assembly has no Flow Simulation project." & vbCrLf &
                   "Create one once (Flow Simulation > Wizard), add a volume heat source per chip and temperature goals," & vbCrLf &
                   "then run again. See docs\HOWTO_M7.md.", vbExclamation, MACRO_NAME)
            GoTo done
        End If

        ' ---------- heat sources ----------
        en = fda.EnumFeatures
        If en IsNot Nothing Then
            en.Reset()
            ft = en.Next
            Do While ft IsNot Nothing
                tk = -1
                vr = ft.Type
                If IsNumeric(vr) Then tk = CInt(vr)
                If tk = nikVolumeSource Then
                    refU = FindRef(CStr(ft.Name), heat)
                    If refU = "" Then
                        notes.Add("Volume source '" & CStr(ft.Name) & "' has no RefDes in its name: left as it is.")
                    ElseIf seen.Exists(refU) Then
                        notes.Add("Volume source '" & CStr(ft.Name) & "': " & refU & " already has a source. Left as it is.")
                    Else
                        seen.Add(refU, 1)
                        q = CDbl(heat(refU))
                        If SetSourceHeat(ft, fda, q) Then
                            nSet = nSet + 1 : applied = applied + q
                            lines.Add("source " & CStr(ft.Name) & " = " & Format(q, "0.000") & " W (" & refU & ")")
                        End If
                    End If
                End If
                ft = en.Next
            Loop
        End If
        For Each hk As Object In heat.Keys
            If CDbl(heat(hk)) > 0 And Not seen.Exists(CStr(hk)) Then
                notes.Add(CStr(hk) & " has " & Format(CDbl(heat(hk)), "0.00") & " W in the register but no Flow volume source with " & CStr(hk) & " in its name.")
            End If
        Next

        ' ---------- ambient ----------
        If PT_IsNum(PT_Rng(wb, "Sim_Ambient").Cells(1, 1).Value) Then
            amb = CDbl(PT_Rng(wb, "Sim_Ambient").Cells(1, 1).Value)
            nAmb = SetInitialTemperature(fda, amb + 273.15)
            If nAmb > 0 Then
                lines.Add("initial temperature = " & Format(amb, "0.0") & " C (" & nAmb & " parameters)")
            Else
                notes.Add("Initial temperature was not set: no temperature parameter found in the initial conditions.")
            End If
        Else
            notes.Add("Ambient on Sim setup is not a number: initial temperature left as it is.")
        End If

        If nSet = 0 Then
            MsgBox("RunFlow stopped. No Flow volume heat source was updated." & vbCrLf & vbCrLf & ListNotes(), vbExclamation, MACRO_NAME)
            GoTo done
        End If

        ' ---------- solve or read ----------
        runName = "FL_" & NowStamp()
        ans = MsgBox(nSet & " heat sources set, " & Format(applied, "0.00") & " W of " & Format(PT_NameNum(wb, "Reg_TotalQ", 0), "0.00") & " W in the register." &
                     vbCrLf & vbCrLf & "Yes = solve now (can take many minutes) and read the goals" & vbCrLf &
                     "No = do not solve, read the goals of the last solved run" & vbCrLf & "Cancel = stop here",
                     vbYesNoCancel + vbQuestion, MACRO_NAME)
        If ans = vbCancel Then
            md.Save()
            WriteReport(wb, runName, CStr(fda.Name), applied, False)
            MsgBox("Sources updated. Nothing solved, nothing written to Excel." & vbCrLf & "Report: " & ReportPath(wb, runName), vbInformation, MACRO_NAME)
            GoTo done
        End If
        If ans = vbYes Then
            If Not IsTrue(fda.Solve2(True, True, True, False)) Then
                MsgBox("RunFlow stopped. The Flow solver did not finish." & vbCrLf & ListNotes() &
                       "Check the project in SolidWorks (Flow Simulation > Run).", vbExclamation, MACRO_NAME)
                GoTo done
            End If
            solved = True
        End If
        rh = fda.LoadResults(True)
        If rh Is Nothing Then
            MsgBox("RunFlow stopped. The project has no results to read. Solve it first (Yes at the question).", vbExclamation, MACRO_NAME)
            GoTo done
        End If

        ' ---------- goals ----------
        gres = rh.GetGoalsCalculationResults2
        If gres IsNot Nothing Then ge = gres.GetGoalsEnum
        If ge Is Nothing Then
            MsgBox("RunFlow stopped. The results have no goals. Add temperature goals to the Flow project.", vbExclamation, MACRO_NAME)
            GoTo done
        End If
        ge.Reset()
        gl = ge.Next
        Do While gl IsNot Nothing
            gname = CStr(gl.GetGoalName)
            gv = CDbl(gl.GetLastCalculatedValue) - 273.15
            gk = GoalKind(gname)
            gref = FindRef(gname, heat)
            grp = ""
            If gref = "" Then grp = FindGroup(gname)
            nGoals = nGoals + 1
            If (gref <> "" Or grp <> "") And gv < -270 Then
                notes.Add("Goal '" & gname & "' has no valid temperature (0 K): not stored. Check that its body is meshed and solid.")
                lines.Add("goal " & gname & " = no value (0 K): skipped")
            ElseIf gref <> "" Then
                If gk = "avg" Then
                    avg(gref) = gv
                ElseIf gk = "max" Then
                    peak(gref) = gv
                Else
                    lines.Add("goal " & gname & " = " & Format(gv, "0.00") & " C (min, not stored)")
                End If
                lines.Add("goal " & gname & " = " & Format(gv, "0.00") & " C -> " & gref & " " & gk)
            ElseIf grp = "SPREADER" Then
                If gk = "min" Then
                    sprMin = gv : hSprMin = True
                ElseIf gk = "max" Then
                    sprMax = gv : hSprMax = True
                End If
                lines.Add("goal " & gname & " = " & Format(gv, "0.00") & " C -> spreader " & gk)
            ElseIf grp = "PCB" Then
                If gk = "min" Then
                    pcbMin = gv : hPcbMin = True
                ElseIf gk = "max" Then
                    pcbMax = gv : hPcbMax = True
                End If
                lines.Add("goal " & gname & " = " & Format(gv, "0.00") & " C -> pcb " & gk)
            Else
                lines.Add("goal " & gname & " = " & Format(gv, "0.000") & " (not a RefDes, spreader or PCB goal: ignored)")
            End If
            gl = ge.Next
        Loop
        If peak.Count = 0 And avg.Count = 0 Then
            MsgBox("RunFlow stopped. No goal name contains a RefDes. Nothing was written to Excel." & vbCrLf &
                   "Rename the goals, for example 'U1 max' (see docs\HOWTO_M7.md).", vbExclamation, MACRO_NAME)
            GoTo done
        End If
        md.Save()
        On Error GoTo 0

        ' ---------- Excel ----------
        On Error GoTo xlFail
        xl.ScreenUpdating = False
        hdr = PT_Rng(wb, "Res_Header")
        ws = hdr.Worksheet
        lastR = CInt(ws.Cells(ws.Rows.Count, hdr.Column).End(-4162).Row)
        If lastR < CInt(hdr.Row) Then lastR = CInt(hdr.Row)
        r = lastR + 1
        sv = ToTable(PT_Rng(wb, "Sim_Surfaces").Value)
        ws.Cells(r, hdr.Column).Value = runName
        ws.Cells(r, hdr.Column + 1).Value = NowText()
        ws.Cells(r, hdr.Column + 2).Value = "Flow " & CStr(fda.Name) & IIf(solved, "", " (last run)")
        ws.Cells(r, hdr.Column + 6).Value = PT_Num(sv(1, 2), 0)
        ws.Cells(r, hdr.Column + 7).Value = PT_Num(sv(3, 2), 0)
        ws.Cells(r, hdr.Column + 8).Value = PT_TimModes(wb)
        If hSprMax Then ws.Cells(r, hdr.Column + 9).Value = Round(sprMax, 2)
        If hSprMin Then ws.Cells(r, hdr.Column + 10).Value = Round(sprMin, 2)
        If hPcbMax Then ws.Cells(r, hdr.Column + 11).Value = Round(pcbMax, 2)
        If hPcbMin Then ws.Cells(r, hdr.Column + 12).Value = Round(pcbMin, 2)
        For Each refK As Object In peak.Keys
            c = PT_RefColumn(ws, hdr, CStr(refK))
            ws.Cells(r, c).Value = Round(CDbl(peak(refK)), 2)
            If avg.Exists(refK) Then ws.Cells(r, c + 1).Value = Round(CDbl(avg(refK)), 2)
        Next
        For Each refK As Object In avg.Keys
            If Not peak.Exists(refK) Then
                c = PT_RefColumn(ws, hdr, CStr(refK))
                ws.Cells(r, c + 1).Value = Round(CDbl(avg(refK)), 2)
            End If
        Next

        PT_WriteSimTj(wb, peak)
        xl.ScreenUpdating = True
        xl.Calculate()
        wb.Save()
        On Error GoTo 0
        WriteReport(wb, runName, CStr(fda.Name), applied, True)

        s = "Written and saved to:" & vbCrLf & CStr(wb.FullName) & vbCrLf & "(Sim results row " & r & ", Thermal solution column AF)" & vbCrLf & vbCrLf &
            "Flow project " & CStr(fda.Name) & ": " & nSet & " heat sources set (" & Format(applied, "0.00") & " W), " &
            peak.Count & " chips with a peak goal, " & nGoals & " goals read." & vbCrLf &
            "Report: " & ReportPath(wb, runName) & vbCrLf & ListNotes()
        MsgBox(s, vbInformation, MACRO_NAME)
        GoTo done

apiFail:
        MsgBox("RunFlow stopped in the Flow Simulation API: " & Err.Description & vbCrLf &
               "Nothing was written to Excel. Check that Flow Simulation is on and the project opens in SolidWorks.", vbExclamation, MACRO_NAME)
        GoTo done
xlFail:
        xl.ScreenUpdating = True
        MsgBox("RunFlow stopped while writing to Excel: " & Err.Description & vbCrLf &
               "The workbook was not saved. Close any cell you are editing, close the workbook without saving, and run again.",
               vbExclamation, MACRO_NAME)
done:
        On Error Resume Next
        gl = Nothing : ge = Nothing : gres = Nothing : rh = Nothing
        ft = Nothing : en = Nothing : fda = Nothing : md = Nothing : ia = Nothing
        If api IsNot Nothing Then api.UnloadProductAPI()
        api = Nothing
    End Sub

    ' ---------------------------------------------------------------- Flow API
    ' Loads the Flow API object. Tries the direct class first, then the out-of-process server. Returns Nothing and says why.
    Private Function OpenFlowApi(ByRef why As String) As Object
        Dim api As Object = Nothing, srv As Object = Nothing, rv As Object = Nothing, errTxt As String = ""
        Try
            api = CreateObject("NIKCommonApi2.BaseApiObject")
        Catch ex As Exception
            errTxt = ex.Message
        End Try
        If api Is Nothing Then
            Try
                srv = GetObject(FLOW_SRV_CLSID)
                If srv IsNot Nothing Then api = srv.LoadNIKCommonApiLibrary
            Catch ex As Exception
                errTxt = ex.Message
            End Try
        End If
        If api Is Nothing Then
            why = "The Flow Simulation API is not available (" & errTxt & ")." & vbCrLf &
                  "Install the API SDK: run 'flow simulation api sdk x64.exe' from the api folder of the Flow Simulation install."
            Return Nothing
        End If
        errTxt = ""
        Try
            rv = api.LoadProductAPI2(FLOW_PRODUCT, FLOW_VERSION)
        Catch ex As Exception
            errTxt = ex.Message
        End Try
        If Not IsTrue(rv) Then
            why = "The Flow API could not load " & FLOW_PRODUCT & " " & FLOW_VERSION & " (" & errTxt & ")."
            Return Nothing
        End If
        Return api
    End Function

    ' The Flow API hands back its Boolean results as a raw 1 (not -1), so "Not value" is wrong. Test non-zero instead.
    Private Function IsTrue(ByVal v As Object) As Boolean
        Try
            Return CInt(v) <> 0
        Catch
            Return False
        End Try
    End Function

    Private Function AttachFlow(ByVal api As Object) As Object
        Dim ia As Object = Nothing
        Try
            ia = api.Attach2RunningObject2(CInt(swApp.GetProcessID))
        Catch
        End Try
        If ia Is Nothing Then
            Try
                ia = api.Attach2RunningObject
            Catch
            End Try
        End If
        Return ia
    End Function

    ' Volume source: type = Heat generation rate, value = q (W). True when both were accepted.
    Private Function SetSourceHeat(ByVal ft As Object, ByVal fda As Object, ByVal q As Double) As Boolean
        Dim pf As Object, ep As Object, par As Object, rc As Integer, gotQ As Boolean = False, pt As Integer
        pf = ft.GetInterface("IParametrizedFeature")
        If pf Is Nothing Then
            notes.Add("Source '" & CStr(ft.Name) & "': no parameters available.")
            Return False
        End If
        For pass As Integer = 1 To 2
            ep = pf.EnumParameters
            If ep Is Nothing Then Exit For
            ep.Reset()
            par = ep.Next
            Do While par IsNot Nothing
                pt = CInt(par.Type)
                If pass = 1 And pt = nikVolumeSourceType Then
                    rc = CInt(par.SetLongValue(nikVolumeSourceType_HeatGenerationRate))
                    If rc <> 0 Then notes.Add("Source '" & CStr(ft.Name) & "': could not set the type to heat generation rate (code " & rc & ").")
                ElseIf pass = 2 And pt = nikHeatGenerationRate Then
                    If CInt(par.DependencyType) <> nikDependencyConstant Then
                        notes.Add("Source '" & CStr(ft.Name) & "': heat is a formula or table in the project. Change it to a constant. Left as it is.")
                    Else
                        rc = CInt(par.SetValue(q))
                        If rc = 0 Then
                            gotQ = True
                        Else
                            notes.Add("Source '" & CStr(ft.Name) & "': heat value refused (code " & rc & ").")
                        End If
                    End If
                End If
                par = ep.Next
            Loop
        Next
        If Not gotQ Then Return False
        If Not IsTrue(ft.Rebuild(fda)) Then notes.Add("Source '" & CStr(ft.Name) & "' did not rebuild: " & CStr(fda.GetLastRebuildError))
        Return True
    End Function

    ' Sets fluid temperature, initial solid temperature and radiation environment temperature (kelvin). Returns how many.
    Private Function SetInitialTemperature(ByVal fda As Object, ByVal kelvin As Double) As Integer
        Dim ic As Object, pf As Object, ep As Object, par As Object, rc As Integer, cnt As Integer = 0, pt As Integer
        ic = fda.GlobalInitialConditions
        If ic Is Nothing Then Return 0
        pf = ic.GetInterface("IParametrizedFeature")
        If pf Is Nothing Then Return 0
        ep = pf.EnumParameters
        If ep Is Nothing Then Return 0
        ep.Reset()
        par = ep.Next
        Do While par IsNot Nothing
            pt = CInt(par.Type)
            If pt = nikTemperature Or pt = nikInitialSolidTemperature Or pt = nikEnvRadiativeTemperature Then
                If CInt(par.DependencyType) = nikDependencyConstant Then
                    rc = CInt(par.SetValue(kelvin))
                    If rc = 0 Then cnt = cnt + 1
                End If
            End If
            par = ep.Next
        Loop
        Return cnt
    End Function

    ' ---------------------------------------------------------------- names
    ' Upper-case words of a text (letters and digits only)
    Private Function Words(ByVal txt As String) As String()
        Return Split(Trim(reWords.Replace(UCase(txt), " ")), " ")
    End Function

    ' First word of txt that is a RefDes in refs, else ""
    Private Function FindRef(ByVal txt As String, ByVal refs As Dict) As String
        For Each w As String In Words(txt)
            If w <> "" Then
                If refs.Exists(w) Then Return w
            End If
        Next
        Return ""
    End Function

    ' "SPREADER" or "PCB" when the goal name has that word, else ""
    Private Function FindGroup(ByVal txt As String) As String
        For Each w As String In Words(txt)
            If w = "SPREADER" Or w = "PCB" Then Return w
        Next
        Return ""
    End Function

    ' "min", "avg" or "max" (default) from the words of a goal name
    Private Function GoalKind(ByVal txt As String) As String
        For Each w As String In Words(txt)
            Select Case w
                Case "MIN", "MINIMUM" : Return "min"
                Case "AVG", "AV", "MEAN", "AVERAGE" : Return "avg"
            End Select
        Next
        Return "max"
    End Function

    Private Function ListNotes() As String
        Dim cnt As Integer = 0, s As String = ""
        For Each q As Object In notes
            cnt = cnt + 1
            If cnt > 6 Then
                s = s & "- ... more in the report" & vbCrLf
                Exit For
            End If
            s = s & "- " & CStr(q) & vbCrLf
        Next
        Return s
    End Function

    ' ---------------------------------------------------------------- report
    Private Function ReportPath(ByVal wb As Object, ByVal runName As String) As String
        Return PT_OutputFolder(wb) & PT_SafeFile(PT_Prefix(wb)) & "RunFlow_" & runName & ".txt"
    End Function

    Private Sub WriteReport(ByVal wb As Object, ByVal runName As String, ByVal projName As String, ByVal applied As Double, ByVal readGoals As Boolean)
        Try
            Using ts As System.IO.StreamWriter = NewTextFile(ReportPath(wb, runName))
                ts.WriteLine("RunFlow " & runName & "  " & NowText())
                ts.WriteLine("Flow project: " & projName)
                ts.WriteLine("Heat applied: " & Format(applied, "0.000") & " W of " & Format(PT_NameNum(wb, "Reg_TotalQ", 0), "0.000") & " W in the register")
                ts.WriteLine("Goals read: " & IIf(readGoals, "yes", "no"))
                ts.WriteLine("")
                For Each q As Object In lines
                    ts.WriteLine(CStr(q))
                Next
                ts.WriteLine("")
                For Each q As Object In notes
                    ts.WriteLine("NOTE " & CStr(q))
                Next
            End Using
        Catch
        End Try
    End Sub

End Module
