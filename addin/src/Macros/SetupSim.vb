Option Strict Off
Option Explicit On

Imports System
Imports Microsoft.VisualBasic
Imports SW = SolidWorks.Interop.sldworks
Imports CW = SolidWorks.Interop.cosworks

' Port of src/SetupSim.bas (milestone M6, SolidWorks Simulation thermal). Same behaviour and messages as the VBA macro.
'
' Makes a new steady-state thermal study on the thermal assembly from the Sim setup tab, meshes it and runs it. Every run
' is a new study (PT_yyyymmdd_hhmm); older studies are never changed. Then run ReadResults.
'   Materials   PCB orthotropic, chips as calibrated blocks (or two-resistor stacks, M8), parts without heat one k,
'               spreader and TIMs from the Spreader tab.
'   Heat        each RefDes gets the Q used of the Component register, spread over its bodies (core only for two-resistor).
'   Contacts    global bonded; thermal resistance sets where Sim setup asks for them; Contact R TIM mode.
'   Surfaces    convection and radiation to ambient per group.
'   Mesh        global size, finer control on chips and TIMs. High quality.
' If anything is missing the study is made but not meshed or run, and the message says what to fix.
' Refuses to run while the Change log has unchecked rows.
Public Module SetupSim

    Private Const MACRO_NAME As String = "SetupSim"
    Private Const swsThermal As Integer = 3
    Private Const swsSteadyState As Integer = 1
    Private Const swsUnitSI As Integer = 0
    Private Const swsMM As Integer = 0
    Private Const swsBondedThermal As Integer = 1
    Private Const swsResistanceSet As Integer = 0
    Private Const swsResistTotal As Integer = 0
    Private Const swsResistDistributed As Integer = 1
    Private Const swsSurfaceToAmbient As Integer = 0
    Private Const swsHighQuality As Integer = 1
    Private Const CFG_HEATONLY As String = "PT_HeatOnly"

    ' Sim chips table offsets (Sim_Chips, column B = 0)
    Private Const S_REF As Integer = 0, S_MODEL As Integer = 1, S_JC As Integer = 2, S_JB As Integer = 3, S_KZ As Integer = 6, S_KXY As Integer = 7
    ' Spreader chip table offsets (Spr_Chips, column B = 0)
    Private Const P_REF As Integer = 0, P_COV As Integer = 7, P_TIMMODE As Integer = 9, P_TIMK As Integer = 11, P_PED As Integer = 14, P_TIMR As Integer = 15, P_CHECK As Integer = 16

    Private problems As Collection, notes As Collection
    Private tag As Collection                     ' {role, key, refdes, component}
    Private pbar As SW.UserProgressBar            ' SolidWorks progress bar
    Private tagIdx As Dict                        ' lower-case component name (full and last segment) -> position in tag
    Private heatQ As Dict                         ' RefDes (upper) -> W
    Private chipKz As Dict, chipKxy As Dict       ' RefDes (upper) -> W/m K
    Private timK As Dict, timR As Dict, timBody As Dict, pedH As Dict   ' RefDes (upper) -> values, covered chips only
    Private pcbT As Double
    Private kSprBot As Double, densSprBot As Double   ' bottom plate (Spreader bottom tab); same as the top plate when that tab is missing
    Private twoRes As Dict                        ' RefDes (upper) -> {theta jc, theta jb, layer mm, area mm2}; two-resistor chips only
    Private layerFrac As Double, kCore As Double  ' outer layer thickness as a fraction of chip height; k of the junction core

    Public Sub Main()
        problems = New Collection() : notes = New Collection()
        pbar = Nothing : tagIdx = Nothing

        ' ---------- workbook ----------
        Dim xl As Object, wb As Object
        xl = PT_GetExcel()
        wb = PT_PickWorkbook(xl, MACRO_NAME, "Sim setup|Spreader|CAD geometry|Component register|Change log")
        If wb Is Nothing Then Exit Sub
        If Not PT_NeedNames(wb, "Sim_Type|Sim_RunName|Sim_Ambient|Sim_MeshGlobal|Sim_MeshFine|Sim_IncludeNonHeat|Sim_kParts|" &
                                "Sim_Surfaces|Sim_Contacts|Sim_Chips|Spr_Chips|Spr_k|Spr_Density|Geo_kXY|Geo_kZ|Geo_PCB_Used|" &
                                "Reg_RefDes|Reg_QUsed|Reg_TotalQ|Log_Header|Log_Unchecked|Set_Prefix|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
        If PT_RefuseIfUnchecked(wb, MACRO_NAME) Then Exit Sub
        xl.Calculate()
        If UCase(PT_NameStr(wb, "Sim_Type", "Thermal")) <> "THERMAL" Then
            MsgBox("SetupSim did not run. Study type on Sim setup is " & PT_NameStr(wb, "Sim_Type", "") &
                   ". This macro does SolidWorks Simulation thermal only; Flow Simulation comes later (M7).", vbExclamation, MACRO_NAME)
            Exit Sub
        End If

        Dim amb As Double, meshG As Double, meshF As Double, inclNonHeat As Boolean, kParts As Double
        Dim kPcbXY As Double, kPcbZ As Double, kSpr As Double, densSpr As Double, totalReg As Double
        amb = PT_NameNum(wb, "Sim_Ambient", 25)
        meshG = PT_NameNum(wb, "Sim_MeshGlobal", 3)
        meshF = PT_NameNum(wb, "Sim_MeshFine", 0.5)
        inclNonHeat = (UCase(PT_NameStr(wb, "Sim_IncludeNonHeat", "Yes")) <> "NO")
        kParts = PT_NameNum(wb, "Sim_kParts", 0.5)
        layerFrac = PT_NameNum(wb, "Sim_LayerFrac", 0.1)
        kCore = PT_NameNum(wb, "Sim_kCore", 100)
        If layerFrac <= 0 Or layerFrac >= 0.45 Or kCore <= 0 Then
            MsgBox("SetupSim did not run. On Sim setup the outer layer thickness must be above 0 and below 0.45 (fraction of the chip height) " &
                   "and the core k above 0.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        kPcbXY = PT_NameNum(wb, "Geo_kXY", 0)
        kPcbZ = PT_NameNum(wb, "Geo_kZ", 0)
        kSpr = PT_NameNum(wb, "Spr_k", 0)
        densSpr = PT_NameNum(wb, "Spr_Density", 2700)
        kSprBot = kSpr : densSprBot = densSpr
        Dim wsBot As Object, rB As Object
        wsBot = PT_Sheet(wb, "Spreader bottom")
        If wsBot IsNot Nothing Then
            rB = TabName(wsBot, "Spr_k")
            If rB IsNot Nothing Then kSprBot = PT_Num(rB.Cells(1, 1).Value, kSpr)
            rB = TabName(wsBot, "Spr_Density")
            If rB IsNot Nothing Then densSprBot = PT_Num(rB.Cells(1, 1).Value, densSpr)
        End If
        pcbT = PT_Num(PT_Rng(wb, "Geo_PCB_Used").Cells(3, 1).Value, 0)
        totalReg = PT_NameNum(wb, "Reg_TotalQ", 0)
        If meshG <= 0 Or meshF <= 0 Or kParts <= 0 Or kPcbXY <= 0 Or kPcbZ <= 0 Or kSpr <= 0 Then
            MsgBox("SetupSim did not run. Mesh sizes, k of parts without heat, PCB k (CAD geometry D19:D20) and spreader k " &
                   "must all be above 0.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        ReadHeat(wb)
        ReadChips(wb)
        ReadTims(wb)

        ' ---------- assembly, configuration ----------
        Dim asm As Object, asmPath As String
        asmPath = PT_AsmPath(wb)
        asm = PT_OpenAsm(asmPath)
        If asm Is Nothing Then
            MsgBox("SetupSim did not run. No thermal assembly at" & vbCrLf & asmPath & vbCrLf & "Run BuildBoard first.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        PT_ActivateDoc(asm)
        Dim hasSpr As Boolean = False, simDoc As Object, study As Object, runName As String = "", errC As Integer = 0, k As Integer = 0
        Dim applied As Double, ran As Boolean = False, runErr As Integer
        On Error GoTo swFail
        Progress("SetupSim: preparing the assembly", 5)
        SetConfiguration(asm, inclNonHeat)
        tag = PT_ListTagged(asm)
        tagIdx = Nothing
        If EnsureLayers() Then
            asm.EditRebuild3()
            PT_ActivateDoc(asm)
            tag = PT_ListTagged(asm)
            tagIdx = Nothing
        End If
        ' Bare board study: no spreader in the assembly (not built yet, or suppressed with its TIMs).
        ' The chip rows of the Spreader tab are then ignored and the chip tops lose heat to air.
        For Each itB As Object In tag
            If CStr(itB(0)) = "Spreader" Then hasSpr = True
        Next
        If Not hasSpr Then
            timK.RemoveAll() : timR.RemoveAll() : timBody.RemoveAll() : pedH.RemoveAll()
            notes.Add("Bare board study: no spreader in the assembly, so the Spreader tab chip rows were ignored.")
        End If

        ' ---------- Simulation ----------
        simDoc = PT_GetSimDoc(asm)
        If simDoc Is Nothing Then
            ProgressEnd()
            MsgBox("SetupSim did not run. SolidWorks Simulation is not available." & vbCrLf &
                   "Turn it on in Tools > Add-Ins > SOLIDWORKS Simulation, then run again.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        runName = "PT_" & NowStamp()
        Do While PT_FindStudy(simDoc, runName) IsNot Nothing
            k = k + 1
            runName = "PT_" & NowStamp() & "_" & k
        Loop
        study = CType(simDoc.StudyManager, CW.ICWStudyManager).CreateNewStudy3(runName, swsThermal, 0, errC)
        If study Is Nothing Then Err.Raise(vbObjectError + 60, , "could not create the study (error " & errC & ")")
        study.ThermalStudyOptions.SolutionType = swsSteadyState
        ' SolverType is left at its default (Automatic). Setting it to FFEPlus crashed SolidWorks 2023 SP5.

        Progress("SetupSim: materials", 20)
        ApplyMaterials(study, kParts, kPcbXY, kPcbZ, kSpr, densSpr)
        Progress("SetupSim: heat, contacts, surfaces", 40)
        applied = ApplyHeat(study)
        ApplyContacts(study, wb)
        ApplySurfaces(study, wb, amb)

        ' ---------- mesh and run, only when nothing is missing ----------
        If problems.Count = 0 Then
            Progress("SetupSim: meshing", 70)
            MeshStudy(study, meshG, meshF)
            If problems.Count = 0 Then
                Progress("SetupSim: solving. SolidWorks is busy for several minutes, do not close it", 85)
                runErr = CInt(study.RunAnalysis)
                If runErr <> 0 Then
                    problems.Add("The analysis stopped with error " & runErr & " (swsRunAnalysisError_e; 24 = solver failed, " &
                                 "16 = thermal conductivity missing, 19/20 = material missing).")
                Else
                    ran = True
                End If
            End If
        End If
        ProgressEnd()
        SwSave3(asm, PT_swSaveSilent, errC, k)
        On Error GoTo 0

        ' ---------- Excel ----------
        On Error GoTo xlFail
        xl.ScreenUpdating = False
        PT_Rng(wb, "Sim_RunName").Value = runName
        xl.ScreenUpdating = True
        wb.Save()
        On Error GoTo 0

        ' ---------- report ----------
        Dim rptPath As String, s As String, cnt As Integer = 0
        rptPath = PT_OutputFolder(wb) & PT_SafeFile(PT_Prefix(wb)) & "SetupSim_" & runName & ".txt"
        WriteReport(rptPath, asmPath, runName, applied, totalReg, ran)
        s = "Study " & runName & " made in" & vbCrLf & asmPath & vbCrLf & "(run name written to Sim setup, " & CStr(wb.Name) & ")" & vbCrLf & vbCrLf &
            "Heat applied " & Round(applied, 3) & " W, register total " & Round(totalReg, 3) & " W." & vbCrLf &
            IIf(hasSpr, "", "Bare board study (no spreader). Compare it on the Bare board network tab." & vbCrLf)
        If ran Then
            s = s & "Meshed and solved. Next: run ReadResults." & vbCrLf
        ElseIf problems.Count > 0 Then
            s = s & "NOT meshed or run. Fix these and run SetupSim again:" & vbCrLf
            For Each it As Object In problems
                cnt = cnt + 1
                If cnt > 6 Then
                    s = s & "(more in the report)" & vbCrLf
                    Exit For
                End If
                s = s & "- " & CStr(it) & vbCrLf
            Next
        End If
        s = s & vbCrLf & "Report: " & rptPath
        MsgBox(s, IIf(ran, vbInformation, vbExclamation), MACRO_NAME)
        Exit Sub

swFail:
        ProgressEnd()
        MsgBox("SetupSim stopped in SolidWorks: " & Err.Description & vbCrLf &
               "A half-made study may be in the Simulation tree: delete it, send this message, and run again.", vbExclamation, MACRO_NAME)
        Exit Sub
xlFail:
        xl.ScreenUpdating = True
        MsgBox("The study was made, but writing the run name to Excel failed: " & Err.Description, vbExclamation, MACRO_NAME)
    End Sub

    ' Tab-level name on a sheet, or Nothing
    Private Function TabName(ByVal ws As Object, ByVal nameS As String) As Object
        Try
            Return ws.Names(nameS).RefersToRange
        Catch
            Return Nothing
        End Try
    End Function

    ' ================================================================= Excel side
    Private Sub ReadHeat(ByVal wb As Object)
        Dim refs(,) As Object, qs(,) As Object, key As String
        heatQ = New Dict()
        refs = ToTable(PT_Rng(wb, "Reg_RefDes").Value)
        qs = ToTable(PT_Rng(wb, "Reg_QUsed").Value)
        For r As Integer = 1 To UBound(refs, 1)
            key = UCase(PT_Str(refs(r, 1)))
            If key <> "" And PT_Num(qs(r, 1), 0) > 0 Then
                If heatQ.Exists(key) Then heatQ(key) = CDbl(heatQ(key)) + PT_Num(qs(r, 1), 0) Else heatQ.Add(key, PT_Num(qs(r, 1), 0))
            End If
        Next
    End Sub

    Private Sub ReadChips(ByVal wb As Object)
        Dim v(,) As Object, key As String, kz As Double, kxy As Double, thJc As Double, thJb As Double
        chipKz = New Dict()
        chipKxy = New Dict()
        twoRes = New Dict()
        v = ToTable(PT_Rng(wb, "Sim_Chips").Value)
        For r As Integer = 1 To UBound(v, 1)
            key = UCase(PT_Str(v(r, 1 + S_REF)))
            If key <> "" And Not chipKz.Exists(key) Then
                If UCase(PT_Str(v(r, 1 + S_MODEL))) = "TWO-RESISTOR" Then
                    thJc = PT_Num(v(r, 1 + S_JC), 0)
                    thJb = PT_Num(v(r, 1 + S_JB), 0)
                    If thJc > 0 And thJb > 0 And heatQ.Exists(key) Then
                        If Not twoRes.Exists(key) Then twoRes.Add(key, New Double() {thJc, thJb, 0.0, 0.0})
                    Else
                        notes.Add(key & ": two-resistor needs theta jc, theta jb and heat; built as a calibrated block.")
                    End If
                End If
                kz = PT_Num(v(r, 1 + S_KZ), 0)
                kxy = PT_Num(v(r, 1 + S_KXY), 0)
                If kz > 0 Then
                    If kxy <= 0 Then kxy = kz
                    chipKz.Add(key, kz)
                    chipKxy.Add(key, kxy)
                End If
            End If
        Next
    End Sub

    Private Sub ReadTims(ByVal wb As Object)
        Dim v(,) As Object, key As String, rg As Object, wsB As Object
        timK = New Dict() : timR = New Dict()
        timBody = New Dict() : pedH = New Dict()
        For pass As Integer = 0 To 1                       ' 0 = Spreader tab (top plate), 1 = Spreader bottom tab
            rg = Nothing
            If pass = 0 Then
                rg = PT_Rng(wb, "Spr_Chips")
            Else
                wsB = PT_Sheet(wb, "Spreader bottom")
                If wsB IsNot Nothing Then rg = TabName(wsB, "Spr_Chips")
            End If
            If rg IsNot Nothing Then
                v = ToTable(rg.Value)
                For r As Integer = 1 To UBound(v, 1)
                    key = UCase(PT_Str(v(r, 1 + P_REF)))
                    If key <> "" And UCase(PT_Str(v(r, 1 + P_COV))) = "YES" And PT_Str(v(r, 1 + P_CHECK)) = "OK" And Not timK.Exists(key) Then
                        timK.Add(key, PT_Num(v(r, 1 + P_TIMK), 0))
                        timR.Add(key, PT_Num(v(r, 1 + P_TIMR), 0))
                        timBody.Add(key, (UCase(PT_Str(v(r, 1 + P_TIMMODE))) = "BODY"))
                        pedH.Add(key, PT_Num(v(r, 1 + P_PED), 0))
                    End If
                Next
            End If
        Next
    End Sub

    ' ================================================================= two-resistor chips (M8)
    ' Makes the part file of every two-resistor chip a stack of three bodies (bottom layer, junction core, top layer),
    ' and turns a chip that is no longer two-resistor back into one box. True if any part file was changed.
    Private Function EnsureLayers() As Boolean
        Dim c As Object, md As Object, refU As String, bx As Double(), lx As Double, wy As Double, hz As Double
        Dim tt As Double, have As String, want As String, errC As Integer = 0, warns As Integer = 0, p As Double(), changed As Boolean = False
        For Each it As Object In tag
            If CStr(it(0)) = "Part" And CStr(it(2)) <> "" Then
                refU = UCase(CStr(it(2)))
                c = it(3)
                md = c.GetModelDoc2
                have = PT_GetProp(md, "Layers_um")
                If twoRes.Exists(refU) Or (have <> "" And have <> "-") Then
                    lx = PT_PropNum(md, "L_um") / 1000.0
                    wy = PT_PropNum(md, "W_um") / 1000.0
                    hz = PT_PropNum(md, "H_um") / 1000.0
                    If lx <= 0 Or wy <= 0 Or hz <= 0 Then
                        bx = PT_CompBox(c)
                        If bx IsNot Nothing Then
                            lx = bx(3) : wy = bx(4) : hz = bx(5)
                        End If
                    End If
                    If lx > 0 And wy > 0 And hz > 0 Then
                        If twoRes.Exists(refU) Then
                            tt = Round(layerFrac * hz, 3)
                            want = PT_Um(tt) & "|" & PT_Um(tt)
                            If have <> want Then
                                If PT_BuildLayeredBox(md, lx, wy, hz, tt, tt, TopFace(CStr(it(2))), "Bottom") Then
                                    SwSave3(md, PT_swSaveSilent, errC, warns)
                                    changed = True
                                Else
                                    notes.Add(refU & ": the three layers could not be made (" & PT_LastErr & "); built as a calibrated block.")
                                    PT_BuildBox(md, lx, wy, hz, TopFace(CStr(it(2))), "Bottom")
                                    SwSave3(md, PT_swSaveSilent, errC, warns)
                                    changed = True
                                    twoRes.Remove(refU)
                                End If
                            End If
                            If twoRes.Exists(refU) Then
                                p = twoRes(refU)
                                twoRes(refU) = New Double() {p(0), p(1), tt, lx * wy}
                            End If
                        Else
                            PT_BuildBox(md, lx, wy, hz, TopFace(CStr(it(2))), "Bottom")
                            PT_SetProp(md, "Layers_um", "-")
                            SwSave3(md, PT_swSaveSilent, errC, warns)
                            changed = True
                            notes.Add(refU & ": layers removed, the chip model is Block again.")
                        End If
                    End If
                End If
            End If
        Next
        Return changed
    End Function

    ' ================================================================= configuration
    Private Sub SetConfiguration(ByVal asm As Object, ByVal inclNonHeat As Boolean)
        Dim names As Object, hasCfg As Boolean = False, cur As String
        names = asm.GetConfigurationNames
        For i As Integer = 0 To UBound(names)
            If CStr(names(i)) = CFG_HEATONLY Then hasCfg = True
        Next
        cur = CStr(asm.ConfigurationManager.ActiveConfiguration.Name)
        If inclNonHeat Then
            If cur = CFG_HEATONLY Then
                For i As Integer = 0 To UBound(names)
                    If CStr(names(i)) <> CFG_HEATONLY Then
                        asm.ShowConfiguration2(names(i))
                        Exit For
                    End If
                Next
            End If
            Exit Sub
        End If
        If Not hasCfg Then asm.ConfigurationManager.AddConfiguration2(CFG_HEATONLY, "Made by SetupSim: parts without heat suppressed", "", 0, "", "", True)
        asm.ShowConfiguration2(CFG_HEATONLY)
        ' in this configuration: parts with heat active, parts without heat suppressed
        For Each it As Object In AllTagged(asm)
            If CStr(it(0)) = "Part" Then
                If heatQ.Exists(UCase(CStr(it(2)))) Then it(3).SetSuppression2(PT_swFullyResolved) Else it(3).SetSuppression2(PT_swSuppressed)
            End If
        Next
        notes.Add("Configuration " & CFG_HEATONLY & " is active: parts without heat are suppressed there.")
    End Sub

    ' like PT_ListTagged but includes suppressed components (their tags read by resolving them for a moment)
    Private Function AllTagged(ByVal asm As Object) As Collection
        Dim comps As Object, c As Object, md As Object, res As New Collection(), supp As Boolean
        comps = asm.GetComponents(True)
        If Not IsEmpty(comps) Then
            For i As Integer = 0 To UBound(comps)
                c = comps(i)
                supp = (c.GetSuppression2 = PT_swSuppressed)
                If supp Then c.SetSuppression2(PT_swFullyResolved)
                md = c.GetModelDoc2
                res.Add(Arr(PT_GetProp(md, "Role"), PT_GetProp(md, "Key"), PT_GetProp(md, "RefDes"), c))
                If supp Then c.SetSuppression2(PT_swSuppressed)
            Next
        End If
        Return res
    End Function

    ' ================================================================= materials
    Private Sub ApplyMaterials(ByVal study As Object, ByVal kParts As Double, ByVal kPcbXY As Double, ByVal kPcbZ As Double,
                               ByVal kSpr As Double, ByVal densSpr As Double)
        Dim sm As Object, errC As Integer = 0, scomp As Object, body As Object, info As Object()
        Dim refU As String, roleS As String, c As Object, compName As String
        sm = study.SolidManager
        For i As Integer = 0 To CInt(sm.ComponentCount) - 1
            scomp = CType(sm, CW.ICWSolidManager).GetComponentAt(i, errC)
            If scomp IsNot Nothing Then
                compName = CStr(scomp.ComponentName)
                info = TagOfSimComp(compName)
                If info Is Nothing Then
                    roleS = "" : refU = "" : c = Nothing
                Else
                    roleS = CStr(info(0)) : refU = UCase(CStr(info(2))) : c = info(3)
                End If
                For j As Integer = 0 To CInt(scomp.SolidBodyCount) - 1
                    body = CType(scomp, CW.ICWSolidComponent).GetSolidBodyAt(j, errC)
                    If body IsNot Nothing Then
                        Select Case roleS
                            Case "PCB"
                                SetMat(body, "PT_PCB", kPcbXY, kPcbZ, 1850, c, compName)
                            Case "Spreader"
                                If InStr(LCase(compName), "bottom") > 0 Then
                                    SetMat(body, "PT_SpreaderBottom", kSprBot, kSprBot, densSprBot, Nothing, compName)
                                Else
                                    SetMat(body, "PT_Spreader", kSpr, kSpr, densSpr, Nothing, compName)
                                End If
                            Case "TIM"
                                If timK.Exists(refU) Then
                                    If CDbl(timK(refU)) > 0 Then
                                        SetMat(body, "PT_TIM_" & refU, CDbl(timK(refU)), CDbl(timK(refU)), 2500, Nothing, compName)
                                    Else
                                        problems.Add("TIM k of " & refU & " is missing on the Spreader tab")
                                    End If
                                Else
                                    problems.Add("TIM part " & compName & " has no covered chip row on the Spreader tab")
                                End If
                            Case "Part"
                                If twoRes.Exists(refU) And CInt(scomp.SolidBodyCount) = 3 Then
                                    SetLayerMat(body, refU, j, compName)
                                ElseIf chipKz.Exists(refU) Then
                                    SetMat(body, "PT_Chip_" & refU, CDbl(chipKxy(refU)), CDbl(chipKz(refU)), 2000, c, compName)
                                Else
                                    If heatQ.Exists(refU) Then notes.Add(refU & ": no calibrated k on Sim setup, generic k used")
                                    SetMat(body, "PT_Part", kParts, kParts, 2000, Nothing, compName)
                                End If
                            Case Else
                                notes.Add(compName & " is not tagged by the macros: generic k used")
                                SetMat(body, "PT_Part", kParts, kParts, 2000, Nothing, compName)
                        End Select
                    End If
                Next
            End If
        Next
    End Sub

    ' Material of one body of a two-resistor chip. Layer by body name (feature names PT_Box, PT_LayerCore, PT_LayerTop),
    ' by body order when the names are something else (bottom, core, top).
    Private Sub SetLayerMat(ByVal body As Object, ByVal refU As String, ByVal idx As Integer, ByVal compName As String)
        Dim p As Double(), nm As String, layer As Integer, k As Double
        p = twoRes(refU)
        nm = UCase(CStr(body.SolidBodyName))
        If InStr(nm, "LAYERCORE") > 0 Then
            layer = 1
        ElseIf InStr(nm, "LAYERTOP") > 0 Then
            layer = 2
        ElseIf InStr(nm, "PT_BOX") > 0 Then
            layer = 0
        Else
            layer = idx
            notes.Add(refU & ": body name """ & CStr(body.SolidBodyName) & """ not recognised, layer taken from body order.")
        End If
        Select Case layer
            Case 0
                k = (p(2) / 1000.0) / (p(1) * p(3) * 0.000001)
                SetMat(body, "PT_Chip_" & refU & "_Bottom", k, k, 2000, Nothing, compName)
                notes.Add(refU & ": bottom layer " & Round(p(2), 3) & " mm, k " & Round(k, 3) & " W/m K (theta jb " & p(1) & " C/W).")
            Case 1
                SetMat(body, "PT_Chip_" & refU & "_Core", kCore, kCore, 2000, Nothing, compName)
            Case Else
                k = (p(2) / 1000.0) / (p(0) * p(3) * 0.000001)
                SetMat(body, "PT_Chip_" & refU & "_Top", k, k, 2000, Nothing, compName)
                notes.Add(refU & ": top layer " & Round(p(2), 3) & " mm, k " & Round(k, 3) & " W/m K (theta jc " & p(0) & " C/W).")
        End Select
    End Sub

    ' Tag of the assembly component whose name matches a Simulation component name, or Nothing
    Private Function TagOfSimComp(ByVal simName As String) As Object()
        Dim nm As String, i As Integer = 0, kLast As String
        If tagIdx Is Nothing Then
            ' one pass over the components instead of one per Simulation component (each Name2 is a COM call)
            tagIdx = New Dict()
            For Each it As Object In tag
                i = i + 1
                nm = LCase(CStr(it(3).Name2))
                If Not tagIdx.Exists(nm) Then tagIdx.Add(nm, i)
                kLast = Mid(nm, InStrRev(nm, "/") + 1)
                If Not tagIdx.Exists(kLast) Then tagIdx.Add(kLast, i)
            Next
        End If
        nm = LCase(simName)
        If Not tagIdx.Exists(nm) Then Return Nothing
        Return CType(tag(CInt(tagIdx(nm))), Object())
    End Function

    ' Isotropic when kxy = kz or no component is given; otherwise orthotropic with Z normal to the part's Front plane.
    Private Sub SetMat(ByVal body As Object, ByVal matName As String, ByVal kxy As Double, ByVal kz As Double, ByVal dens As Double,
                       ByVal c As Object, ByVal compName As String)
        Dim mat As Object, errC As Integer, pl As Object
        mat = body.GetDefaultMaterial
        If mat Is Nothing Then
            problems.Add("No material object for " & compName)
            Exit Sub
        End If
        mat.MaterialUnits = swsUnitSI
        mat.MaterialName = matName
        If c Is Nothing Or Math.Abs(kxy - kz) < 0.000001 Then
            mat.ModelType = 0
            mat.SetPropertyByName2("KX", kxy, False)
        Else
            pl = CompFrontPlane(c)
            mat.ModelType = 1
            mat.SetPropertyByName2("KX", kxy, False)
            mat.SetPropertyByName2("KY", kxy, False)
            mat.SetPropertyByName2("KZ", kz, False)
            errC = -1
            If pl IsNot Nothing Then CType(mat, CW.ICWMaterial).SetReferencePlane(pl, errC)
            If errC <> 0 Then
                ' fall back: isotropic with the through-thickness value (conservative)
                mat.ModelType = 0
                mat.SetPropertyByName2("KX", kz, False)
                notes.Add(compName & ": orthotropic direction could not be set, isotropic k = " & Round(kz, 3) & " used")
            End If
        End If
        mat.SetPropertyByName2("DENS", dens, False)
        mat.SetPropertyByName2("C", 900, False)
        errC = CInt(body.SetSolidBodyMaterial(mat))
        If errC <> 0 Then problems.Add("Material for " & compName & " was refused (error " & errC & ")")
    End Sub

    ' first reference plane of a component (Front = XY), assembly context
    Private Function CompFrontPlane(ByVal c As Object) As Object
        Try
            Dim f As Object = c.FirstFeature
            Do While f IsNot Nothing
                If CStr(f.GetTypeName2) = "RefPlane" Then Return f
                f = f.GetNextFeature
            Loop
        Catch
        End Try
        Return Nothing
    End Function

    ' ================================================================= heat
    Private Function ApplyHeat(ByVal study As Object) As Double
        Dim lbc As Object, bodies As Collection, b As Object
        Dim hp As Object, errC As Integer = 0, q As Double, total As Double = 0, arr As Object
        lbc = study.LoadsAndRestraintsManager
        For Each refU As Object In heatQ.Keys
            q = CDbl(heatQ(refU))
            bodies = New Collection()
            For Each it As Object In tag
                If CStr(it(0)) = "Part" And UCase(CStr(it(2))) = CStr(refU) Then
                    If twoRes.Exists(refU) Then b = PT_CoreBodies(it(3)) Else b = PT_CompBodies(it(3))
                    If Not IsEmpty(b) Then
                        For i As Integer = 0 To UBound(b)
                            bodies.Add(b(i))
                        Next
                    End If
                End If
            Next
            If bodies.Count = 0 Then
                problems.Add(CStr(refU) & " has " & Round(q, 3) & " W in the register but no body with that RefDes in the assembly " &
                             "(set RefDes and Include on CAD geometry, run BuildBoard)")
            Else
                arr = PT_ToArray(bodies)
                hp = CType(lbc, CW.ICWLoadsAndRestraintsManager).AddHeatPower(arr, errC)
                If hp Is Nothing Or errC <> 0 Then
                    problems.Add("Heat power on " & CStr(refU) & " was refused (error " & errC & ")")
                Else
                    hp.HeatPowerBeginEdit()
                    hp.Unit = swsUnitSI
                    hp.HPValue = q / bodies.Count          ' the value is applied to each body
                    errC = CInt(hp.HeatPowerEndEdit)
                    If errC <> 0 Then problems.Add("Heat power on " & CStr(refU) & " could not be set (error " & errC & ")") Else total = total + q
                End If
            End If
        Next
        Return total
    End Function

    ' ================================================================= contacts
    Private Sub ApplyContacts(ByVal study As Object, ByVal wb As Object)
        Dim cm As Object, rv(,) As Object, rPart As Double, rChipTim As Double, rTimSpr As Double
        Dim c As Object, bx As Double(), refU As String, fA As Object, spr As Object = Nothing, tims As New Dict()
        Dim sprBot As Object = Nothing, sprUse As Object, chipBot As Boolean
        cm = study.ContactManager
        cm.SetGlobalContact(swsBondedThermal, 0)
        rv = ToTable(PT_Rng(wb, "Sim_Contacts").Value)
        rPart = PT_Num(rv(1, 2), 0) : rChipTim = PT_Num(rv(2, 2), 0) : rTimSpr = PT_Num(rv(3, 2), 0)

        Dim pcb As Object = Nothing
        For Each it As Object In tag
            If CStr(it(0)) = "PCB" Then pcb = it(3)
            If CStr(it(0)) = "Spreader" Then
                If InStr(LCase(CStr(it(3).Name2)), "bottom") > 0 Then sprBot = it(3) Else spr = it(3)
            End If
            If CStr(it(0)) = "TIM" Then
                If Not tims.Exists(UCase(CStr(it(2)))) Then tims.Add(UCase(CStr(it(2))), it(3))
            End If
        Next

        ' part to PCB
        If rPart > 0 And pcb IsNot Nothing Then
            For Each it As Object In tag
                If CStr(it(0)) = "Part" Then
                    c = it(3)
                    bx = PT_CompBox(c)
                    If bx IsNot Nothing Then
                        If Math.Abs(bx(2) - pcbT) < 0.01 Then
                            ResistanceSet(cm, PT_AsmFace(c, "Bottom"), PT_AsmFace(pcb, "PCB_Top"), rPart, swsResistDistributed, CStr(it(1)) & " to PCB")
                        ElseIf Math.Abs(bx(2) + bx(5)) < 0.01 Then
                            ResistanceSet(cm, PT_AsmFace(c, TopFace(CStr(it(2)))), PT_AsmFace(pcb, "PCB_Bottom"), rPart, swsResistDistributed, CStr(it(1)) & " to PCB")
                        End If
                    End If
                End If
            Next
        End If

        ' chips under the spreader
        For Each it As Object In tag
            If CStr(it(0)) = "Part" And CStr(it(2)) <> "" Then
                refU = UCase(CStr(it(2)))
                If timK.Exists(refU) Then
                    ' a chip hanging under the PCB touches the bottom plate with its Bottom face; a chip on top touches the top plate with its top face
                    bx = PT_CompBox(it(3))
                    chipBot = False
                    If bx IsNot Nothing Then chipBot = (Math.Abs(bx(2) + bx(5)) < 0.01)
                    If chipBot Then
                        fA = PT_AsmFace(it(3), "Bottom")
                        sprUse = sprBot
                    Else
                        fA = PT_AsmFace(it(3), TopFace(CStr(it(2))))
                        sprUse = spr
                    End If
                    If CBool(timBody(refU)) Then
                        If Not tims.Exists(refU) Then
                            problems.Add(refU & ": TIM mode Body but no TIM part in the assembly (run BuildSpreader)")
                        Else
                            If rChipTim > 0 Then ResistanceSet(cm, fA, PT_AsmFace(tims(refU), "TIM_" & PT_SafeFile(CStr(it(2))) & "_Bottom"),
                                                               rChipTim, swsResistDistributed, refU & " chip to TIM")
                            If rTimSpr > 0 Then ResistanceSet(cm, PT_AsmFace(tims(refU), "TIM_" & PT_SafeFile(CStr(it(2))) & "_Top"),
                                                              SpreaderContact(sprUse, CStr(it(2))), rTimSpr, swsResistDistributed, refU & " TIM to spreader")
                        End If
                    Else
                        ' Contact R: the TIM is a resistance between chip top and spreader
                        If CDbl(timR(refU)) > 0 Then
                            ResistanceSet(cm, fA, SpreaderContact(sprUse, CStr(it(2))), CDbl(timR(refU)), swsResistTotal, refU & " TIM as contact resistance")
                        Else
                            problems.Add(refU & ": TIM mode Contact R but TIM R on the Spreader tab is empty")
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Function TopFace(ByVal refD As String) As String
        If refD <> "" Then Return "Chip_" & PT_SafeFile(refD) & "_Top"
        Return "Top"
    End Function

    Private Function SpreaderContact(ByVal spr As Object, ByVal refD As String) As Object
        If spr Is Nothing Then Return Nothing
        If pedH.Exists(UCase(refD)) Then
            If CDbl(pedH(UCase(refD))) > 0.0005 Then Return PT_AsmFace(spr, "Contact_" & PT_SafeFile(refD))
        End If
        Return PT_AsmFace(spr, "Underside")
    End Function

    Private Sub ResistanceSet(ByVal cm As Object, ByVal fSrc As Object, ByVal fTgt As Object, ByVal rValue As Double,
                              ByVal rType As Integer, ByVal what As String)
        Dim cs As Object, errC As Integer = 0
        If fSrc Is Nothing Or fTgt Is Nothing Then
            problems.Add("Contact resistance " & what & ": a named face is missing (run BuildBoard / BuildSpreader again)")
            Exit Sub
        End If
        cs = CType(cm, CW.ICWContactManager).CreateContactSet2(swsResistanceSet, 0, New Object() {fSrc}, New Object() {fTgt}, errC)
        If cs Is Nothing Or errC <> 0 Then
            problems.Add("Contact resistance " & what & " was refused (error " & errC & ")")
            Exit Sub
        End If
        cs.ContactSetBeginEdit()
        cs.ResistanceType = rType
        cs.ResistanceUnit = swsUnitSI
        cs.ResistanceValue = rValue
        errC = CInt(cs.ContactSetEndEdit)
        If errC <> 0 Then problems.Add("Contact resistance " & what & " could not be set (error " & errC & ")")
    End Sub

    ' ================================================================= convection and radiation
    Private Sub ApplySurfaces(ByVal study As Object, ByVal wb As Object, ByVal amb As Double)
        Dim sv(,) As Object, groups(4) As Collection, c As Object, md As Object
        Dim faces As Collection, nm As String, refU As String, bx As Double(), onTop As Boolean, onBottom As Boolean
        For g As Integer = 0 To 4
            groups(g) = New Collection()
        Next
        For Each it As Object In tag
            c = it(3)
            md = c.GetModelDoc2
            faces = PT_AllFaces(md)
            refU = UCase(CStr(it(2)))
            Select Case CStr(it(0))
                Case "Spreader"
                    For Each fc As Object In faces
                        nm = CStr(md.GetEntityName(fc))
                        If nm = "Top" Then
                            groups(0).Add(c.GetCorrespondingEntity(fc))
                        ElseIf Left(nm, 8) <> "Contact_" Then
                            groups(1).Add(c.GetCorrespondingEntity(fc))
                        End If
                    Next
                Case "PCB"
                    For Each fc As Object In faces
                        If CStr(md.GetEntityName(fc)) = "PCB_Top" Then groups(2).Add(c.GetCorrespondingEntity(fc)) Else groups(3).Add(c.GetCorrespondingEntity(fc))
                    Next
                Case "Part"
                    bx = PT_CompBox(c)
                    onTop = False : onBottom = False
                    If bx IsNot Nothing Then
                        onTop = (Math.Abs(bx(2) - pcbT) < 0.01)              ' sits on the PCB top: its bottom touches the board
                        onBottom = (Math.Abs(bx(2) + bx(5)) < 0.01)          ' hangs under the PCB: its top touches the board
                    End If
                    For Each fc As Object In faces
                        nm = CStr(md.GetEntityName(fc))
                        If twoRes.Exists(refU) And InnerFace(fc, bx) Then
                            ' face between two layers of the chip: touches the next layer
                        ElseIf nm = "Bottom" And onTop Then
                        ElseIf nm = TopFace(CStr(it(2))) And (onBottom Or timK.Exists(refU)) Then
                        Else
                            groups(4).Add(c.GetCorrespondingEntity(fc))
                        End If
                    Next
                Case "TIM"
                    For Each fc As Object In faces
                        nm = CStr(md.GetEntityName(fc))
                        If Right(nm, 4) <> "_Top" And Right(nm, 7) <> "_Bottom" Then groups(4).Add(c.GetCorrespondingEntity(fc))
                    Next
            End Select
        Next
        sv = ToTable(PT_Rng(wb, "Sim_Surfaces").Value)
        For g As Integer = 0 To 4
            If groups(g).Count > 0 Then SurfaceLoads(study, groups(g), PT_Num(sv(g + 1, 2), 0), PT_Num(sv(g + 1, 3), 0), amb, PT_Str(sv(g + 1, 1)))
        Next
        notes.Add("PCB top and bottom faces get convection over their whole area, including where parts sit (small error).")
    End Sub

    ' True for a flat horizontal face of a chip part that is neither its bottom nor its top (part frame: z from 0 to H).
    Private Function InnerFace(ByVal fc As Object, ByVal bx As Double()) As Boolean
        Dim nv As Object = Nothing, bb As Object = Nothing, z As Double
        If bx Is Nothing Then Return False
        Try
            nv = fc.Normal
        Catch
        End Try
        Try
            bb = fc.GetBox
        Catch
        End Try
        If IsEmpty(nv) Or IsEmpty(bb) Then Return False
        If Math.Abs(CDbl(nv(2))) < 0.99 Then Return False
        If Math.Abs(CDbl(bb(5)) - CDbl(bb(2))) * 1000.0 > 0.001 Then Return False
        z = CDbl(bb(2)) * 1000.0
        Return (z > 0.001 And z < bx(5) - 0.001)
    End Function

    Private Sub SurfaceLoads(ByVal study As Object, ByVal faces As Collection, ByVal h As Double, ByVal eps As Double,
                             ByVal amb As Double, ByVal groupName As String)
        Dim lbc As Object, arr As Object, cv As Object, rd As Object, errC As Integer = 0
        lbc = study.LoadsAndRestraintsManager
        arr = PT_ToArray(faces)
        If h > 0 Then
            cv = CType(lbc, CW.ICWLoadsAndRestraintsManager).AddConvection(arr, errC)
            If cv Is Nothing Or errC <> 0 Then
                problems.Add("Convection on " & groupName & " was refused (error " & errC & ")")
            Else
                cv.ConvectionBeginEdit()
                cv.Unit = swsUnitSI
                cv.ConvectionCoefficient = h
                cv.BulkAmbientTemperature = amb + 273.15          ' SI units: kelvin
                errC = CInt(cv.ConvectionEndEdit)
                If errC <> 0 Then problems.Add("Convection on " & groupName & " could not be set (error " & errC & ")")
            End If
        End If
        If eps > 0 Then
            errC = 0
            rd = CType(lbc, CW.ICWLoadsAndRestraintsManager).AddRadiation(swsSurfaceToAmbient, arr, errC)
            If rd Is Nothing Or errC <> 0 Then
                problems.Add("Radiation on " & groupName & " was refused (error " & errC & ")")
            Else
                rd.RadiationBeginEdit()
                rd.Unit = swsUnitSI
                rd.Emmisivity = eps
                rd.ViewFactor = 1
                rd.AmbientTemperature = amb + 273.15
                errC = CInt(rd.RadiationEndEdit)
                If errC <> 0 Then problems.Add("Radiation on " & groupName & " could not be set (error " & errC & ")")
            End If
        End If
    End Sub

    ' ================================================================= mesh
    Private Sub MeshStudy(ByVal study As Object, ByVal meshG As Double, ByVal meshF As Double)
        Dim msh As Object, fine As New Collection(), mc As Object, errC As Integer = 0, arr As Object
        Dim el As Double, tl As Double, tol As Double
        msh = study.Mesh
        msh.Quality = swsHighQuality
        For Each it As Object In tag
            If CStr(it(0)) = "TIM" Then fine.Add(it(3))
            If CStr(it(0)) = "Part" Then
                If heatQ.Exists(UCase(CStr(it(2)))) Then fine.Add(it(3))
            End If
        Next
        If fine.Count > 0 Then
            arr = PT_ToArray(fine)
            mc = CType(msh, CW.ICWMesh).ApplyMeshControl(arr, errC)
            If mc Is Nothing Or errC <> 0 Then
                notes.Add("Fine mesh control on chips and TIMs was refused (error " & errC & "); global size used")
            Else
                mc.MeshControlBeginEdit()
                mc.Units = swsMM
                mc.ElementSize = meshF
                mc.MeshControlEndEdit()
            End If
        End If
        CType(msh, CW.ICWMesh).GetDefaultElementSizeAndTolerance(swsMM, el, tl)
        If el > 0 Then tol = tl * meshG / el Else tol = meshG * 0.05
        errC = CInt(study.CreateMesh(swsMM, meshG, tol))
        If errC <> 0 Then problems.Add("Meshing failed (error " & errC & "). Try a smaller global or fine mesh size on Sim setup.")
    End Sub

    ' ================================================================= report
    Private Sub WriteReport(ByVal rptPath As String, ByVal asmPath As String, ByVal runName As String, ByVal applied As Double,
                            ByVal totalReg As Double, ByVal ran As Boolean)
        Try
            Using ts As System.IO.StreamWriter = NewTextFile(rptPath)
                ts.WriteLine("SetupSim report " & NowText())
                ts.WriteLine("Assembly: " & asmPath)
                ts.WriteLine("Study: " & runName & IIf(ran, "  (meshed and solved)", "  (NOT solved)"))
                ts.WriteLine("Heat applied " & applied & " W, register total " & totalReg & " W")
                ts.WriteLine("")
                ts.WriteLine("HEAT PER REFDES")
                For Each refU As Object In heatQ.Keys
                    ts.WriteLine("  " & CStr(refU) & ": " & CDbl(heatQ(refU)) & " W")
                Next
                ts.WriteLine("")
                ts.WriteLine("PROBLEMS (" & problems.Count & ")")
                For Each it As Object In problems
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

    ' ================================================================= progress
    ' SolidWorks status bar progress with a title, so a run of several minutes does not look frozen.
    Private Sub Progress(ByVal txt As String, ByVal pos As Integer)
        Try
            If pbar Is Nothing Then
                Dim pb As SW.UserProgressBar = Nothing
                CType(swApp, SW.SldWorks).GetUserProgressBar(pb)
                pbar = pb
                If pbar IsNot Nothing Then pbar.Start(0, 100, txt)
            Else
                pbar.UpdateTitle(txt)
            End If
            If pbar IsNot Nothing Then pbar.UpdateProgress(pos)
        Catch
        End Try
    End Sub

    Private Sub ProgressEnd()
        Try
            If pbar IsNot Nothing Then pbar.End()
        Catch
        End Try
        pbar = Nothing
    End Sub

End Module
