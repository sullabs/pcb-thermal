Attribute VB_Name = "ReadResults"
Option Explicit
' =====================================================================================
' ReadResults  -  SolidWorks 2023 macro (milestone M6)
'
' Reads the solved study named on Sim setup (Sim_RunName, written by SetupSim) and
' writes the results to the workbook. Works for any board.
'
'   Sim results tab     one new row: run, date, inputs snapshot, spreader and PCB
'                       max/min, energy balance, plot file, and per RefDes the peak
'                       and average temperature (columns added when a RefDes is new).
'   Thermal solution    column AF (Sim Tj) = peak temperature of each RefDes, so
'                       column AG shows hand minus simulation.
'   Plot                picture of the first result plot, CAD\results\<run>.png
'
' Project folder (page, W7): pick <prefix>_project.json at the start. The run is the newest thermal run of
' <prefix>_results.json not read yet (Run Thermal Study added it); its temperatures, energy balance and plot go into
' that run, status "read". Tab 9 of the page shows it.
' Needs modules PTCommon and PTJson in the same macro. Old workbooks: a copy of PCB_thermal_TEMPLATE.xlsx.
' =====================================================================================

Const MACRO_NAME As String = "ReadResults"
Const swsTEMP As Long = 0
Const swsCelsius As Long = 2
Const swsUnitSI As Long = 0

Dim notes As Collection
Dim jsonMode As Boolean

Sub main()
    Set swApp = Application.SldWorks
    PT_Trace MACRO_NAME & ": started (VBA macro)."
    Set notes = New Collection
    jsonMode = False

    ' ---------- project folder (page) or workbook ----------
    Dim xl As Object, wb As Object, rf As Object, rfPath As String, jRun As Object, runName As String, qReg As Double
    Select Case PJ_PickSource(MACRO_NAME)
        Case "json"
            jsonMode = True
        Case "xlsx"
        Case Else
            Exit Sub
    End Select
    If jsonMode Then
        rfPath = PJ_Dir & "\" & PJ_File("results")
        Set rf = PJ_ReadOwn(PJ_File("results"), "pcb-thermal/results")
        If rf Is Nothing Then
            PT_Msg "ReadResults did not run. " & rfPath & " could not be read:" & vbCrLf & PT_LastErr, vbExclamation, MACRO_NAME
            Exit Sub
        End If
        Set jRun = RunToRead(rf)
        If jRun Is Nothing Then
            PT_Msg "ReadResults did not run. " & rfPath & " has no thermal run waiting to be read." & vbCrLf & _
                   "Run the study first (Run Thermal Study, button 6).", vbExclamation, MACRO_NAME
            Exit Sub
        End If
        runName = PJ_Str(jRun, "name", "")
        qReg = HeatTotal(PJ_Obj(jRun, "inputs/forSolidWorks/heat"))
        If qReg <= 0 Then qReg = HeatTotal(PJ_Obj(PJ_Project, "forSolidWorks/heat"))
    Else
        Set xl = PT_GetExcel()
        Set wb = PT_PickWorkbook(xl, MACRO_NAME, "Sim setup|Sim results|Thermal solution|Component register", PJ_WbPath)
        If wb Is Nothing Then Exit Sub
        If Not PT_NeedNames(wb, "Sim_RunName|Sim_MeshGlobal|Sim_MeshFine|Sim_Surfaces|Res_Header|TS_SimTj|Reg_RefDes|Reg_QUsed|" & _
                                "Reg_TotalQ|Spr_Chips|Set_Prefix|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
        xl.Calculate
        runName = PT_NameStr(wb, "Sim_RunName", "")
        qReg = PT_NameNum(wb, "Reg_TotalQ", 0)
    End If
    If runName = "" Then
        PT_Msg "ReadResults did not run. Sim setup has no run name: run SetupSim first.", vbExclamation, MACRO_NAME
        Exit Sub
    End If

    ' ---------- study ----------
    Dim asm As Object, asmPath As String, simDoc As Object, study As Object, res As Object
    If jsonMode Then asmPath = PJ_AsmPath() Else asmPath = PT_AsmPath(wb)
    Set asm = PT_OpenAsm(asmPath)
    If asm Is Nothing Then
        PT_Msg "ReadResults did not run. No thermal assembly at" & vbCrLf & asmPath, vbExclamation, MACRO_NAME
        Exit Sub
    End If
    Set simDoc = PT_GetSimDoc(asm)
    If simDoc Is Nothing Then
        PT_Msg "ReadResults did not run. Turn on SOLIDWORKS Simulation (Tools > Add-Ins) and run again.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    Set study = PT_FindStudy(simDoc, runName)
    If study Is Nothing Then
        PT_Msg "ReadResults did not run. The assembly has no study named " & runName & "." & vbCrLf & _
               IIf(jsonMode, "Run Run Thermal Study again (button 6).", "Run SetupSim, or type an existing study name in the run name cell of Sim setup."), vbExclamation, MACRO_NAME
        Exit Sub
    End If
    On Error Resume Next
    study.ActivateConfiguration
    On Error GoTo swFail
    Set res = study.Results
    If res Is Nothing Then
        PT_Msg "ReadResults did not run. Study " & runName & " has no results: it was not solved." & vbCrLf & _
               "Check the SetupSim report, fix what it lists and run SetupSim again.", vbExclamation, MACRO_NAME
        Exit Sub
    End If

    ' ---------- temperatures ----------
    Dim tag As Collection, it As Variant, refs As Object, refU As Variant, t As Variant
    Dim peak As Object, avg As Object, sprMax As Double, sprMin As Double, pcbMax As Double, pcbMin As Double
    Dim sprFile As String, pcbBodies As Collection, sprBodies As Collection, sprTop As Collection, sprBot As Collection, sprNm As String
    Set tag = PT_ListTagged(asm)
    Set refs = CreateObject("Scripting.Dictionary")
    Set pcbBodies = New Collection: Set sprBodies = New Collection: Set sprTop = New Collection: Set sprBot = New Collection
    For Each it In tag
        Select Case it(0)
            Case "Part"
                If it(2) <> "" Then
                    If Not refs.Exists(UCase(it(2))) Then refs.Add UCase(it(2)), New Collection
                    AddCoreBodies refs(UCase(it(2))), it(3)
                End If
            Case "PCB"
                AddBodies pcbBodies, it(3)
            Case "Spreader"
                AddBodies sprBodies, it(3)
                sprNm = Mid(it(3).GetPathName, InStrRev(it(3).GetPathName, "\") + 1)
                If InStr(LCase(sprNm), "bottom") > 0 Then AddBodies sprBot, it(3) Else AddBodies sprTop, it(3)
                If InStr(sprFile & "; ", sprNm & "; ") = 0 Then sprFile = sprFile & IIf(sprFile = "", "", "; ") & sprNm
        End Select
    Next
    Set peak = CreateObject("Scripting.Dictionary")
    Set avg = CreateObject("Scripting.Dictionary")
    For Each refU In refs.Keys
        t = TempStats(res, refs(refU))
        If Not IsEmpty(t) Then peak.Add refU, t(1): avg.Add refU, t(2)
    Next
    t = TempStats(res, sprBodies)
    If Not IsEmpty(t) Then sprMin = t(0): sprMax = t(1)
    If sprTop.Count > 0 And sprBot.Count > 0 Then
        t = TempStats(res, sprTop)
        If Not IsEmpty(t) Then notes.Add "Top spreader " & Format(t(0), "0.0") & " to " & Format(t(1), "0.0") & " C."
        t = TempStats(res, sprBot)
        If Not IsEmpty(t) Then notes.Add "Bottom spreader " & Format(t(0), "0.0") & " to " & Format(t(1), "0.0") & " C."
    End If
    t = TempStats(res, pcbBodies)
    If Not IsEmpty(t) Then pcbMin = t(0): pcbMax = t(1)

    ' ---------- energy balance: heat in vs heat out of the whole model ----------
    Dim hp As Variant, errC As Long, balance As Double, hasBal As Boolean, qIn As Double, qOut As Double
    If pcbBodies.Count > 0 Then
        hp = res.GetHeatPowerOrEnergy2(False, PT_ToArray(pcbBodies), swsUnitSI, 1, 1, errC)
        If errC = 0 And Not IsEmpty(hp) Then
            ' hp = selection in, out, net, then whole model in, out, net. The gross in and out include exchange with the
            ' ambient, so they differ in a healthy run. The net flow out of the model must equal the heat applied
            ' (register total), so the imbalance is |register total - |net|| as a percentage of the register total.
            qIn = Abs(hp(3)): qOut = Abs(hp(4))
            notes.Add "Heat flow, whole model: in " & Format(qIn, "0.000") & " W, out " & Format(qOut, "0.000") & " W, net " & Format(hp(5), "0.000") & " W."
            If qReg > 0.000001 Then
                balance = Round(Abs(qReg - Abs(hp(5))) / qReg * 100#, 2): hasBal = True
            ElseIf qOut > 0.000001 Then
                balance = Round(Abs(qIn - qOut) / qOut * 100#, 2): hasBal = True
            End If
        Else
            notes.Add "Energy balance could not be read (error " & errC & ")."
        End If
    End If

    ' ---------- plot picture ----------
    Dim plotPath As String, pn As Variant, errs As Long, warns As Long
    If jsonMode Then plotPath = PJ_OutputFolder() & "results\" Else plotPath = PT_OutputFolder(wb) & "results\"
    PT_MakeDirs Left(plotPath, Len(plotPath) - 1)
    plotPath = plotPath & runName & ".png"
    On Error Resume Next
    pn = res.GetPlotNames
    If Not IsEmpty(pn) Then
        res.ActivatePlot ResultPlotName(pn)
        asm.ViewZoomtofit2
        If Not asm.Extension.SaveAs(plotPath, 0, PT_swSaveSilent, Nothing, errs, warns) Then plotPath = ""
    Else
        plotPath = ""
    End If
    If plotPath = "" Then notes.Add "No plot picture saved (no result plot, or the picture could not be written)."
    On Error GoTo 0

    ' ---------- results file (page) ----------
    If jsonMode Then
        WriteRunJson jRun, peak, avg, sprBodies.Count > 0, sprMin, sprMax, pcbBodies.Count > 0, pcbMin, pcbMax, hasBal, balance, plotPath
        If Not PJ_SaveOwn(rf, PJ_File("results"), MACRO_NAME) Then
            PT_Msg "ReadResults read the study, but " & rfPath & " could not be written:" & vbCrLf & PT_LastErr, vbExclamation, MACRO_NAME
            Exit Sub
        End If
        PT_Msg "Written to:" & vbCrLf & rfPath & vbCrLf & "(run " & runName & ", tab 9 of the page)" & vbCrLf & vbCrLf & SummaryText(runName, peak.Count, _
               sprTop.Count > 0 And sprBot.Count > 0, sprBodies.Count > 0, sprMin, sprMax, pcbMin, pcbMax, hasBal, balance, plotPath), vbInformation, MACRO_NAME
        Exit Sub
    End If

    ' ---------- Excel ----------
    Dim hdr As Object, ws As Object, r As Long, c As Long, lastR As Long, sv As Variant
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
    ws.Cells(r, hdr.Column + 2).Value = runName
    ws.Cells(r, hdr.Column + 3).Value = sprFile
    ws.Cells(r, hdr.Column + 4).Value = PT_NameNum(wb, "Sim_MeshGlobal", 0)
    ws.Cells(r, hdr.Column + 5).Value = PT_NameNum(wb, "Sim_MeshFine", 0)
    ws.Cells(r, hdr.Column + 6).Value = PT_Num(sv(1, 2), 0)
    ws.Cells(r, hdr.Column + 7).Value = PT_Num(sv(3, 2), 0)
    ws.Cells(r, hdr.Column + 8).Value = PT_TimModes(wb)
    If sprBodies.Count > 0 Then ws.Cells(r, hdr.Column + 9).Value = Round(sprMax, 2): ws.Cells(r, hdr.Column + 10).Value = Round(sprMin, 2)
    If pcbBodies.Count > 0 Then ws.Cells(r, hdr.Column + 11).Value = Round(pcbMax, 2): ws.Cells(r, hdr.Column + 12).Value = Round(pcbMin, 2)
    If hasBal Then ws.Cells(r, hdr.Column + 13).Value = balance
    ws.Cells(r, hdr.Column + 14).Value = plotPath
    For Each refU In peak.Keys
        c = PT_RefColumn(ws, hdr, CStr(refU))
        ws.Cells(r, c).Value = Round(peak(refU), 2)
        ws.Cells(r, c + 1).Value = Round(avg(refU), 2)
    Next

    ' Thermal solution AF = Sim Tj (peak) per RefDes; rows without a result are cleared
    PT_WriteSimTj wb, peak
    xl.ScreenUpdating = True
    xl.Calculate
    wb.Save
    On Error GoTo 0

    Dim s As String, q As Variant, n As Long
    s = "Written and saved to:" & vbCrLf & wb.FullName & vbCrLf & "(Sim results row " & r & ", Thermal solution column AF)" & vbCrLf & vbCrLf & _
        "Study " & runName & ": " & peak.Count & " parts with a RefDes read." & vbCrLf & _
        IIf(sprTop.Count > 0 And sprBot.Count > 0, "Spreaders (both plates) ", "Spreader ") & IIf(sprBodies.Count > 0, Round(sprMin, 1) & " to " & Round(sprMax, 1) & " C", "none") & ", PCB " & _
        Round(pcbMin, 1) & " to " & Round(pcbMax, 1) & " C." & vbCrLf & _
        "Energy imbalance " & IIf(hasBal, balance & " %", "not read") & " (should be within 2 %)." & vbCrLf
    If plotPath <> "" Then s = s & "Plot: " & plotPath & vbCrLf
    For Each q In notes
        n = n + 1
        If n > 5 Then Exit For
        s = s & "- " & q & vbCrLf
    Next
    PT_Msg s, vbInformation, MACRO_NAME
    Exit Sub

swFail:
    PT_Msg "ReadResults stopped in SolidWorks: " & Err.Description & vbCrLf & "Nothing was written.", vbExclamation, MACRO_NAME
    Exit Sub
xlFail:
    xl.ScreenUpdating = True
    PT_Msg "ReadResults stopped while writing to Excel: " & Err.Description & vbCrLf & _
           "The workbook was not saved. Close any cell you are editing, close the workbook without saving, and run again.", _
           vbExclamation, MACRO_NAME
End Sub

Sub AddBodies(ByVal coll As Collection, ByVal c As Object)
    Dim b As Variant, i As Long
    b = PT_CompBodies(c)
    If IsEmpty(b) Then Exit Sub
    For i = 0 To UBound(b): coll.Add b(i): Next
End Sub

' A chip made of three layers (two-resistor model, M8) reports the junction core; any other part, all its bodies.
Sub AddCoreBodies(ByVal coll As Collection, ByVal c As Object)
    Dim b As Variant, i As Long
    b = PT_CoreBodies(c)
    If IsEmpty(b) Then Exit Sub
    For i = 0 To UBound(b): coll.Add b(i): Next
End Sub

' Array(min, max, mean) of nodal temperatures (C) on some bodies, or Empty
Function TempStats(ByVal res As Object, ByVal bodies As Collection) As Variant
    Dim v As Variant, errC As Long, i As Long, n As Long, tmin As Double, tmax As Double, tsum As Double, tv As Double
    If bodies.Count = 0 Then Exit Function
    v = res.GetThermalForEntities(swsTEMP, 1, Nothing, PT_ToArray(bodies), swsCelsius, errC)
    If errC <> 0 Or IsEmpty(v) Then
        notes.Add "Temperatures could not be read for some bodies (error " & errC & ")."
        Exit Function
    End If
    tmin = 1E+30: tmax = -1E+30
    ' the array is node, temperature, node, temperature, ...
    For i = 1 To UBound(v) Step 2
        tv = CDbl(v(i))
        If tv < tmin Then tmin = tv
        If tv > tmax Then tmax = tv
        tsum = tsum + tv
        n = n + 1
    Next
    If n = 0 Then Exit Function
    TempStats = Array(tmin, tmax, tsum / n)
End Function

' ================================================================= project folder (page) mode
' The newest thermal run of results.json that was made but not read yet.
Function RunToRead(ByVal rf As Object) As Object
    Dim runs As Object, it As Variant
    Set runs = PJ_Obj(rf, "runs")
    If runs Is Nothing Then Exit Function
    For Each it In runs
        If PJ_Str(it, "studyType", "") = "Thermal" And PJ_Str(it, "status", "") <> "read" Then Set RunToRead = it
    Next
End Function

' Sum of the heat per RefDes (W) of a forSolidWorks heat object, 0 when there is none.
Function HeatTotal(ByVal h As Object) As Double
    Dim it As Variant
    If h Is Nothing Then Exit Function
    For Each it In h.Keys
        HeatTotal = HeatTotal + PJ_Num(h, CStr(it), 0)
    Next
End Function

' Puts what was read into the run: status read, PCB and spreader min and max, energy balance, per RefDes peak and
' average (RefDes written as on the page), the plot picture and the notes.
Sub WriteRunJson(ByVal jRun As Object, ByVal peak As Object, ByVal avg As Object, ByVal hasSpr As Boolean, ByVal sprMin As Double, _
                 ByVal sprMax As Double, ByVal hasPcb As Boolean, ByVal pcbMin As Double, ByVal pcbMax As Double, ByVal hasBal As Boolean, _
                 ByVal balance As Double, ByVal plotPath As String)
    Dim refNames As Object, pl As Object, it As Variant, refU As Variant, refS As String, parts As Object, p As Object, mm As Object
    Dim nl As Collection, old As Object
    Set refNames = CreateObject("Scripting.Dictionary")
    Set pl = PJ_Obj(PJ_Project, "parts")
    If Not pl Is Nothing Then
        For Each it In pl
            refS = PJ_Str(it, "refDes", "")
            If refS <> "" And Not refNames.Exists(UCase(refS)) Then refNames.Add UCase(refS), refS
        Next
    End If
    PJ_Put jRun, "status", "read"
    If hasSpr Then
        Set mm = PJ_Dict(): PJ_Put mm, "maxC", Round(sprMax, 2): PJ_Put mm, "minC", Round(sprMin, 2)
        PJ_Put jRun, "spreader", mm
    End If
    If hasPcb Then
        Set mm = PJ_Dict(): PJ_Put mm, "maxC", Round(pcbMax, 2): PJ_Put mm, "minC", Round(pcbMin, 2)
        PJ_Put jRun, "pcb", mm
    End If
    If hasBal Then PJ_Put jRun, "energyBalancePct", balance
    Set parts = PJ_Dict()
    For Each refU In peak.Keys
        If refNames.Exists(CStr(refU)) Then refS = refNames(CStr(refU)) Else refS = CStr(refU)
        Set p = PJ_Dict()
        PJ_Put p, "peakC", Round(peak(refU), 2)
        If avg.Exists(refU) Then PJ_Put p, "avgC", Round(avg(refU), 2)
        PJ_Put parts, refS, p
    Next
    PJ_Put jRun, "parts", parts
    If plotPath <> "" Then PJ_Put jRun, "plotFile", plotPath
    Set nl = New Collection
    Set old = PJ_Obj(jRun, "notes")
    If Not old Is Nothing Then
        For Each it In old: nl.Add CStr(it): Next
    End If
    For Each it In notes: nl.Add CStr(it): Next
    If nl.Count > 0 Then PJ_Put jRun, "notes", nl
End Sub

Function SummaryText(ByVal runName As String, ByVal nRead As Long, ByVal bothPlates As Boolean, ByVal hasSpr As Boolean, _
                     ByVal sprMin As Double, ByVal sprMax As Double, ByVal pcbMin As Double, ByVal pcbMax As Double, _
                     ByVal hasBal As Boolean, ByVal balance As Double, ByVal plotPath As String) As String
    Dim s As String, q As Variant, n As Long
    s = "Study " & runName & ": " & nRead & " parts with a RefDes read." & vbCrLf & _
        IIf(bothPlates, "Spreaders (both plates) ", "Spreader ") & IIf(hasSpr, Round(sprMin, 1) & " to " & Round(sprMax, 1) & " C", "none") & ", PCB " & _
        Round(pcbMin, 1) & " to " & Round(pcbMax, 1) & " C." & vbCrLf & _
        "Energy imbalance " & IIf(hasBal, balance & " %", "not read") & " (should be within 2 %)." & vbCrLf
    If plotPath <> "" Then s = s & "Plot: " & plotPath & vbCrLf
    For Each q In notes
        n = n + 1
        If n > 5 Then Exit For
        s = s & "- " & q & vbCrLf
    Next
    SummaryText = s
End Function

' The temperature plot to save as the run picture. GetPlotNames also lists the mesh quality plot (Quality1),
' sometimes first: take a Thermal plot, else the first plot that is not a mesh quality plot.
Private Function ResultPlotName(ByVal names As Variant) As String
    Dim i As Long, nm As String
    For i = LBound(names) To UBound(names)
        nm = CStr(names(i))
        If LCase(Left(nm, 7)) = "thermal" Then ResultPlotName = nm: Exit Function
    Next
    For i = LBound(names) To UBound(names)
        nm = CStr(names(i))
        If LCase(Left(nm, 7)) <> "quality" Then ResultPlotName = nm: Exit Function
    Next
    ResultPlotName = CStr(names(LBound(names)))
End Function
