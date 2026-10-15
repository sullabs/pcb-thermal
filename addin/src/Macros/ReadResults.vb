Option Strict Off
Option Explicit On

Imports System
Imports Microsoft.VisualBasic
Imports CW = SolidWorks.Interop.cosworks

' Port of src/ReadResults.bas (milestone M6). Same behaviour and messages as the VBA macro.
'
' Reads the solved study named on Sim setup (Sim_RunName, written by SetupSim) and writes the results:
'   Sim results tab     one new row: run, date, inputs snapshot, spreader and PCB max/min, energy balance, plot file,
'                       and per RefDes the peak and average temperature (columns added when a RefDes is new).
'   Thermal solution    column AF (Sim Tj) = peak temperature of each RefDes.
'   Plot                picture of the first result plot, CAD\results\<run>.png
Public Module ReadResults

    Private Const MACRO_NAME As String = "ReadResults"
    Private Const swsTEMP As Integer = 0
    Private Const swsCelsius As Integer = 2
    Private Const swsUnitSI As Integer = 0

    Private notes As Collection

    Public Sub Main()
        notes = New Collection()

        Dim xl As Object, wb As Object
        xl = PT_GetExcel()
        wb = PT_PickWorkbook(xl, MACRO_NAME, "Sim setup|Sim results|Thermal solution|Component register")
        If wb Is Nothing Then Exit Sub
        If Not PT_NeedNames(wb, "Sim_RunName|Sim_MeshGlobal|Sim_MeshFine|Sim_Surfaces|Res_Header|TS_SimTj|Reg_RefDes|Reg_QUsed|" &
                                "Reg_TotalQ|Spr_Chips|Set_Prefix|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
        xl.Calculate()

        Dim runName As String
        runName = PT_NameStr(wb, "Sim_RunName", "")
        If runName = "" Then
            MsgBox("ReadResults did not run. Sim setup has no run name: run SetupSim first.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If

        ' ---------- study ----------
        Dim asm As Object, asmPath As String, simDoc As Object, study As Object, res As Object
        asmPath = PT_AsmPath(wb)
        asm = PT_OpenAsm(asmPath)
        If asm Is Nothing Then
            MsgBox("ReadResults did not run. No thermal assembly at" & vbCrLf & asmPath, vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        simDoc = PT_GetSimDoc(asm)
        If simDoc Is Nothing Then
            MsgBox("ReadResults did not run. Turn on SOLIDWORKS Simulation (Tools > Add-Ins) and run again.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        study = PT_FindStudy(simDoc, runName)
        If study Is Nothing Then
            MsgBox("ReadResults did not run. The assembly has no study named " & runName & "." & vbCrLf &
                   "Run SetupSim, or type an existing study name in the run name cell of Sim setup.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        On Error Resume Next
        study.ActivateConfiguration()
        On Error GoTo swFail
        res = study.Results
        If res Is Nothing Then
            MsgBox("ReadResults did not run. Study " & runName & " has no results: it was not solved." & vbCrLf &
                   "Check the SetupSim report, fix what it lists and run SetupSim again.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If

        ' ---------- temperatures ----------
        Dim tag As Collection, refs As New Dict(), t As Double()
        Dim peak As New Dict(), avg As New Dict(), sprMax As Double, sprMin As Double, pcbMax As Double, pcbMin As Double
        Dim sprFile As String = "", pcbBodies As New Collection(), sprBodies As New Collection(), sprTop As New Collection(), sprBot As New Collection(), sprNm As String
        tag = PT_ListTagged(asm)
        For Each it As Object In tag
            Select Case CStr(it(0))
                Case "Part"
                    If CStr(it(2)) <> "" Then
                        If Not refs.Exists(UCase(CStr(it(2)))) Then refs.Add(UCase(CStr(it(2))), New Collection())
                        AddCoreBodies(refs(UCase(CStr(it(2)))), it(3))
                    End If
                Case "PCB"
                    AddBodies(pcbBodies, it(3))
                Case "Spreader"
                    AddBodies(sprBodies, it(3))
                    sprNm = Mid(CStr(it(3).GetPathName), InStrRev(CStr(it(3).GetPathName), "\") + 1)
                    If InStr(LCase(sprNm), "bottom") > 0 Then AddBodies(sprBot, it(3)) Else AddBodies(sprTop, it(3))
                    If InStr(sprFile & "; ", sprNm & "; ") = 0 Then sprFile = sprFile & IIf(sprFile = "", "", "; ") & sprNm
            End Select
        Next
        For Each refU As Object In refs.Keys
            t = TempStats(res, refs(refU))
            If t IsNot Nothing Then
                peak.Add(refU, t(1)) : avg.Add(refU, t(2))
            End If
        Next
        t = TempStats(res, sprBodies)
        If t IsNot Nothing Then
            sprMin = t(0) : sprMax = t(1)
        End If
        If sprTop.Count > 0 And sprBot.Count > 0 Then
            t = TempStats(res, sprTop)
            If t IsNot Nothing Then notes.Add("Top spreader " & Format(t(0), "0.0") & " to " & Format(t(1), "0.0") & " C.")
            t = TempStats(res, sprBot)
            If t IsNot Nothing Then notes.Add("Bottom spreader " & Format(t(0), "0.0") & " to " & Format(t(1), "0.0") & " C.")
        End If
        t = TempStats(res, pcbBodies)
        If t IsNot Nothing Then
            pcbMin = t(0) : pcbMax = t(1)
        End If

        ' ---------- energy balance: heat in vs heat out of the whole model ----------
        Dim hp As Object, errC As Integer = 0, balance As Double, hasBal As Boolean = False, qIn As Double, qOut As Double, qReg As Double
        If pcbBodies.Count > 0 Then
            hp = CType(res, CW.ICWResults).GetHeatPowerOrEnergy2(False, PT_ToArray(pcbBodies), swsUnitSI, 1, 1, errC)
            If errC = 0 And Not IsEmpty(hp) Then
                ' hp = selection in, out, net, then whole model in, out, net. The net flow out of the model must equal the heat
                ' applied (register total), so the imbalance is |register total - |net|| as a percentage of the register total.
                qIn = Math.Abs(CDbl(hp(3))) : qOut = Math.Abs(CDbl(hp(4)))
                notes.Add("Heat flow, whole model: in " & Format(qIn, "0.000") & " W, out " & Format(qOut, "0.000") & " W, net " & Format(CDbl(hp(5)), "0.000") & " W.")
                qReg = PT_NameNum(wb, "Reg_TotalQ", 0)
                If qReg > 0.000001 Then
                    balance = Round(Math.Abs(qReg - Math.Abs(CDbl(hp(5)))) / qReg * 100.0, 2) : hasBal = True
                ElseIf qOut > 0.000001 Then
                    balance = Round(Math.Abs(qIn - qOut) / qOut * 100.0, 2) : hasBal = True
                End If
            Else
                notes.Add("Energy balance could not be read (error " & errC & ").")
            End If
        End If

        ' ---------- plot picture ----------
        Dim plotPath As String, pn As Object, errs As Integer = 0, warns As Integer = 0
        plotPath = PT_OutputFolder(wb) & "results\"
        PT_MakeDirs(Left(plotPath, Len(plotPath) - 1))
        plotPath = plotPath & runName & ".png"
        On Error Resume Next
        pn = Nothing
        pn = res.GetPlotNames
        If Not IsEmpty(pn) Then
            res.ActivatePlot(CStr(pn(0)))
            asm.ViewZoomtofit2()
            If Not SwSaveAs(asm, plotPath, errs, warns) Then plotPath = ""
        Else
            plotPath = ""
        End If
        If plotPath = "" Then notes.Add("No plot picture saved (no result plot, or the picture could not be written).")
        On Error GoTo 0

        ' ---------- Excel ----------
        Dim hdr As Object, ws As Object, r As Integer, c As Integer, lastR As Integer, sv(,) As Object
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
        ws.Cells(r, hdr.Column + 2).Value = runName
        ws.Cells(r, hdr.Column + 3).Value = sprFile
        ws.Cells(r, hdr.Column + 4).Value = PT_NameNum(wb, "Sim_MeshGlobal", 0)
        ws.Cells(r, hdr.Column + 5).Value = PT_NameNum(wb, "Sim_MeshFine", 0)
        ws.Cells(r, hdr.Column + 6).Value = PT_Num(sv(1, 2), 0)
        ws.Cells(r, hdr.Column + 7).Value = PT_Num(sv(3, 2), 0)
        ws.Cells(r, hdr.Column + 8).Value = PT_TimModes(wb)
        If sprBodies.Count > 0 Then
            ws.Cells(r, hdr.Column + 9).Value = Round(sprMax, 2) : ws.Cells(r, hdr.Column + 10).Value = Round(sprMin, 2)
        End If
        If pcbBodies.Count > 0 Then
            ws.Cells(r, hdr.Column + 11).Value = Round(pcbMax, 2) : ws.Cells(r, hdr.Column + 12).Value = Round(pcbMin, 2)
        End If
        If hasBal Then ws.Cells(r, hdr.Column + 13).Value = balance
        ws.Cells(r, hdr.Column + 14).Value = plotPath
        For Each refU As Object In peak.Keys
            c = PT_RefColumn(ws, hdr, CStr(refU))
            ws.Cells(r, c).Value = Round(CDbl(peak(refU)), 2)
            ws.Cells(r, c + 1).Value = Round(CDbl(avg(refU)), 2)
        Next

        ' Thermal solution AF = Sim Tj (peak) per RefDes; rows without a result are cleared
        PT_WriteSimTj(wb, peak)
        xl.ScreenUpdating = True
        xl.Calculate()
        wb.Save()
        On Error GoTo 0

        Dim s As String, cnt As Integer = 0
        s = "Written and saved to:" & vbCrLf & CStr(wb.FullName) & vbCrLf & "(Sim results row " & r & ", Thermal solution column AF)" & vbCrLf & vbCrLf &
            "Study " & runName & ": " & peak.Count & " parts with a RefDes read." & vbCrLf &
            IIf(sprTop.Count > 0 And sprBot.Count > 0, "Spreaders (both plates) ", "Spreader ") & IIf(sprBodies.Count > 0, Round(sprMin, 1) & " to " & Round(sprMax, 1) & " C", "none") & ", PCB " &
            Round(pcbMin, 1) & " to " & Round(pcbMax, 1) & " C." & vbCrLf &
            "Energy imbalance " & IIf(hasBal, balance & " %", "not read") & " (should be within 2 %)." & vbCrLf
        If plotPath <> "" Then s = s & "Plot: " & plotPath & vbCrLf
        For Each q As Object In notes
            cnt = cnt + 1
            If cnt > 5 Then Exit For
            s = s & "- " & CStr(q) & vbCrLf
        Next
        MsgBox(s, vbInformation, MACRO_NAME)
        Exit Sub

swFail:
        MsgBox("ReadResults stopped in SolidWorks: " & Err.Description & vbCrLf & "Nothing was written.", vbExclamation, MACRO_NAME)
        Exit Sub
xlFail:
        xl.ScreenUpdating = True
        MsgBox("ReadResults stopped while writing to Excel: " & Err.Description & vbCrLf &
               "The workbook was not saved. Close any cell you are editing, close the workbook without saving, and run again.",
               vbExclamation, MACRO_NAME)
    End Sub

    Private Sub AddBodies(ByVal coll As Collection, ByVal c As Object)
        Dim b As Object = PT_CompBodies(c)
        If IsEmpty(b) Then Exit Sub
        For i As Integer = 0 To UBound(b)
            coll.Add(b(i))
        Next
    End Sub

    ' A chip made of three layers (two-resistor model, M8) reports the junction core; any other part, all its bodies.
    Private Sub AddCoreBodies(ByVal coll As Collection, ByVal c As Object)
        Dim b As Object = PT_CoreBodies(c)
        If IsEmpty(b) Then Exit Sub
        For i As Integer = 0 To UBound(b)
            coll.Add(b(i))
        Next
    End Sub

    ' {min, max, mean} of nodal temperatures (C) on some bodies, or Nothing
    Private Function TempStats(ByVal res As Object, ByVal bodies As Collection) As Double()
        Dim v As Object, errC As Integer = 0, cnt As Integer = 0, tmin As Double, tmax As Double, tsum As Double = 0, tv As Double
        If bodies.Count = 0 Then Return Nothing
        v = CType(res, CW.ICWResults).GetThermalForEntities(swsTEMP, 1, Nothing, PT_ToArray(bodies), swsCelsius, errC)
        If errC <> 0 Or IsEmpty(v) Then
            notes.Add("Temperatures could not be read for some bodies (error " & errC & ").")
            Return Nothing
        End If
        Dim va As Array = CType(v, Array)
        tmin = 1.0E+30 : tmax = -1.0E+30
        ' the array is node, temperature, node, temperature, ...
        For i As Integer = 1 To va.Length - 1 Step 2
            tv = CDbl(va.GetValue(i))
            If tv < tmin Then tmin = tv
            If tv > tmax Then tmax = tv
            tsum = tsum + tv
            cnt = cnt + 1
        Next
        If cnt = 0 Then Return Nothing
        Return New Double() {tmin, tmax, tsum / cnt}
    End Function

End Module
