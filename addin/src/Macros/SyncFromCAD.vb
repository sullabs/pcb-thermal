Option Strict Off
Option Explicit On

Imports System
Imports Microsoft.VisualBasic

' Port of src/SyncFromCAD.bas (milestone M5). Same behaviour and messages as the VBA macro.
'
' Reads the thermal assembly and brings what you changed in SolidWorks back to Excel. Does not change any SolidWorks file.
'   PCB and parts   real box vs the Used values on CAD geometry; a difference is written as an override (CAD wins),
'                   logged "override from CAD", or "conflict" when it replaced a value you had typed.
'   Spreader        area, exposed area, volume, mass, thickness, holes -> AS BUILT block; pedestal heights and TIM
'                   thicknesses as found -> Spreader columns U and V; differences logged "excel-owned value differs".
'   Checks          logged "check failed": missing contact face, TIM not touching, cut missing, interference.
Public Module SyncFromCAD

    Private Const MACRO_NAME As String = "SyncFromCAD"
    Private Const swSelFACES As Integer = 2

    ' CAD geometry offsets from column B
    Private Const G_NAMEUSED As Integer = 3, G_REF As Integer = 4, G_OVR As Integer = 13, G_USED As Integer = 19, G_KEY As Integer = 31
    ' Spreader chip table offsets from column B
    Private Const C_REF As Integer = 0, C_X As Integer = 2, C_Y As Integer = 3, C_L As Integer = 4, C_W As Integer = 5, C_TOP As Integer = 6
    Private Const C_COV As Integer = 7, C_TIMMODE As Integer = 9, C_TIMT As Integer = 13, C_PED As Integer = 14, C_CHECK As Integer = 16
    Private Const C_PEDFOUND As Integer = 19, C_TIMFOUND As Integer = 20
    ' clearance cut offsets from column B
    Private Const K_KEY As Integer = 1, K_NAME As Integer = 2

    Private nOverride As Integer, nConflict As Integer, nDiffers As Integer, nFailed As Integer, nRepeat As Integer
    Private openFails As Dict                  ' "key|text" of check failed rows still unchecked on the Change log
    Private tolPos As Double, tolSize As Double
    Private notes As Collection
    Private bottomSide As Boolean              ' the plate being synced is the one under the PCB
    Private wsSide As Object                   ' Spreader bottom tab (its Spr_ and As_ names are tab-level)

    Public Sub Main()
        PT_LogReset()
        notes = New Collection()
        nOverride = 0 : nConflict = 0 : nDiffers = 0 : nFailed = 0 : nRepeat = 0 : bottomSide = False

        Dim xl As Object, wb As Object
        xl = PT_GetExcel()
        wb = PT_PickWorkbook(xl, MACRO_NAME, "CAD geometry|Spreader|Settings|Change log")
        If wb Is Nothing Then Exit Sub
        If Not PT_NeedNames(wb, "Geo_Table|Geo_PCB_Used|Geo_PCB_Override|Spr_Chips|Spr_Cuts|Spr_ZUnder|Spr_Thickness|Spr_Density|" &
                                "As_Area|As_TopArea|As_ExposedArea|As_Volume|As_Mass|As_Thickness|As_Holes|As_LastSync|" &
                                "Log_Header|Set_TolPos|Set_TolSize|Set_Prefix|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
        openFails = PT_OpenLogKeys(wb, "check failed")
        tolPos = PT_NameNum(wb, "Set_TolPos", 0.05)
        tolSize = PT_NameNum(wb, "Set_TolSize", 0.05)
        xl.Calculate()

        Dim asm As Object, asmPath As String
        asmPath = PT_AsmPath(wb)
        asm = PT_OpenAsm(asmPath)
        If asm Is Nothing Then
            MsgBox("SyncFromCAD did not run. No thermal assembly at" & vbCrLf & asmPath & vbCrLf & "Run BuildBoard first.", vbExclamation, MACRO_NAME)
            Exit Sub
        End If
        PT_ActivateDoc(asm)
        PT_LeaveHeatOnly(asm)

        ' ---------- read everything from SolidWorks first ----------
        Dim pcbHave As New Dict(), pcbSupp As New Dict(), partHave As New Dict(), partSupp As New Dict()
        Dim sprHave As New Dict(), sprSupp As New Dict(), timHave As New Dict(), timSupp As New Dict()
        On Error GoTo swFail
        PT_ReadRoles(asm, New String() {"PCB", "Part", "Spreader", "TIM"}, New Dict() {pcbHave, partHave, sprHave, timHave},
                     New Dict() {pcbSupp, partSupp, sprSupp, timSupp})
        RestoreSuppression(pcbHave, pcbSupp)
        RestoreSuppression(partHave, partSupp)
        RestoreSuppression(sprHave, sprSupp)
        RestoreSuppression(timHave, timSupp)
        On Error GoTo 0

        ' ---------- write to Excel ----------
        On Error GoTo xlFail
        xl.ScreenUpdating = False
        SyncPcb(wb, pcbHave, pcbSupp)
        SyncParts(wb, partHave, partSupp)
        Dim cSpr As Object, nSpr As Integer = 0
        For side As Integer = 0 To 1
            bottomSide = (side = 1)
            wsSide = Nothing
            cSpr = PickSpreader(sprHave, sprSupp, bottomSide)
            If cSpr Is Nothing Then
                If bottomSide And nSpr = 0 Then notes.Add("No active spreader in the assembly: spreader checks skipped.")
            Else
                If bottomSide Then wsSide = PT_Sheet(wb, "Spreader bottom")
                If bottomSide And wsSide Is Nothing Then
                    notes.Add("A bottom spreader is in the assembly but the workbook has no Spreader bottom tab: bottom plate skipped.")
                Else
                    SyncSpreader(wb, cSpr, timHave, timSupp)
                    SRng(wb, "As_LastSync").Value = NowText()
                    nSpr = nSpr + 1
                End If
            End If
        Next
        CheckInterference(asm)
        PT_Rng(wb, "As_LastSync").Value = NowText()
        Dim nLog As Integer
        nLog = PT_LogFlush(wb)
        xl.ScreenUpdating = True
        xl.Calculate()
        wb.Save()
        On Error GoTo 0

        Dim s As String, k As Integer = 0
        s = "Written and saved to:" & vbCrLf & CStr(wb.FullName) & vbCrLf & "(tabs CAD geometry, Spreader, Change log)" & vbCrLf & vbCrLf &
            "Read from " & asmPath & vbCrLf & vbCrLf &
            nOverride & " overrides written from CAD, " & nConflict & " of them replaced a value you had typed (conflict)." & vbCrLf &
            nDiffers & " spreader values differ from the Spreader tab design." & vbCrLf &
            nFailed & " checks failed" & IIf(nRepeat > 0, " (" & nRepeat & " already unchecked on the Change log, not added again)", "") & "." & vbCrLf
        If nLog > 0 Then s = s & nLog & " rows added to the Change log tab. Check them there." & vbCrLf Else s = s & "Nothing to report." & vbCrLf
        For Each it As Object In notes
            k = k + 1
            If k > 6 Then
                s = s & "(more notes on the Change log tab)" & vbCrLf
                Exit For
            End If
            s = s & "- " & CStr(it) & vbCrLf
        Next
        MsgBox(s, IIf(nFailed > 0 Or nConflict > 0, vbExclamation, vbInformation), MACRO_NAME)
        Exit Sub

swFail:
        MsgBox("SyncFromCAD stopped while reading SolidWorks: " & Err.Description & vbCrLf & "Nothing was written.", vbExclamation, MACRO_NAME)
        Exit Sub
xlFail:
        xl.ScreenUpdating = True
        MsgBox("SyncFromCAD stopped while writing to Excel: " & Err.Description & vbCrLf &
               "The workbook was not saved. Close any cell you are editing, close the workbook without saving, and run again.",
               vbExclamation, MACRO_NAME)
    End Sub

    ' ================================================================= PCB and parts
    Private Sub SyncPcb(ByVal wb As Object, ByVal have As Dict, ByVal wasSupp As Dict)
        Dim c As Object = Nothing, bx As Double(), usedR As Object, ovrR As Object
        Dim fields As String() = {"Length X", "Width Y", "Thickness Z"}
        For Each it As Object In have.Keys
            If Not CBool(wasSupp(it)) Then c = have(it)
        Next
        If c Is Nothing Then
            CheckFailed("pcb", "PCB", "", "No active PCB part in the thermal assembly")
            Exit Sub
        End If
        bx = PT_CompBox(c)
        If bx Is Nothing Then Exit Sub
        If Math.Abs(bx(0)) > tolPos Or Math.Abs(bx(1)) > tolPos Or Math.Abs(bx(2)) > tolPos Then
            CheckFailed("pcb", "PCB", "", "The PCB corner is not at the origin any more (" & Round(bx(0), 3) & ", " &
                        Round(bx(1), 3) & ", " & Round(bx(2), 3) & "). Every part is placed from that corner: move it back.")
        End If
        usedR = PT_Rng(wb, "Geo_PCB_Used")
        ovrR = PT_Rng(wb, "Geo_PCB_Override")
        For k As Integer = 0 To 2
            WriteOverride(ovrR.Cells(1 + k, 1), PT_Num(usedR.Cells(1 + k, 1).Value, 0), bx(3 + k), tolSize, "pcb", "PCB", "", fields(k))
        Next
    End Sub

    Private Sub SyncParts(ByVal wb As Object, ByVal have As Dict, ByVal wasSupp As Dict)
        Dim rg As Object, v(,) As Object, r As Integer, rowOf As New Dict(), c As Object, bx As Double()
        Dim g As Integer, tol As Double, nameS As String, refD As String
        Dim fields As String() = {"X min", "Y min", "Z min", "L", "W", "H"}
        Dim usedB(5) As Double, clipB(5) As Double, pcbT As Double
        pcbT = PT_Num(PT_Rng(wb, "Geo_PCB_Used").Cells(3, 1).Value, 0)
        rg = PT_Rng(wb, "Geo_Table")
        v = ToTable(rg.Value)
        For r = 1 To UBound(v, 1)
            If PT_Str(v(r, 1 + G_KEY)) <> "" Then
                If Not rowOf.Exists(PT_Str(v(r, 1 + G_KEY))) Then rowOf.Add(PT_Str(v(r, 1 + G_KEY)), r)
            End If
        Next
        For Each it As Object In have.Keys
            If Not CBool(wasSupp(it)) Then
                c = have(it)
                If Not rowOf.Exists(it) Then
                    notes.Add("Part " & CStr(it) & " is in the assembly but has no row on CAD geometry. Run BuildBoard to suppress it.")
                Else
                    r = CInt(rowOf(it))
                    g = CInt(rg.Row) + r - 1
                    nameS = PT_Str(v(r, 1 + G_NAMEUSED))
                    refD = PT_Str(v(r, 1 + G_REF))
                    bx = PT_CompBox(c)
                    If bx IsNot Nothing Then
                        ' BuildBoard cuts boxes at the board surface: compare with the same cut, write back only real edits
                        For k As Integer = 0 To 5
                            usedB(k) = PT_Num(v(r, 1 + G_USED + k), 0) : clipB(k) = usedB(k)
                        Next
                        PT_ClipToBoard(clipB, pcbT)
                        For k As Integer = 0 To 5
                            If k < 3 Then tol = tolPos Else tol = tolSize
                            If Math.Abs(bx(k) - clipB(k)) > tol Then
                                WriteOverride(rg.Worksheet.Cells(g, rg.Column + G_OVR + k), usedB(k), bx(k), tol,
                                              CStr(it), nameS, refD, fields(k))
                            End If
                        Next
                    End If
                End If
            End If
        Next
    End Sub

    ' CAD wins: writes the CAD value into the override cell when it differs from the used value.
    Private Sub WriteOverride(ByVal cell As Object, ByVal usedV As Double, ByVal cadV As Double, ByVal tol As Double,
                              ByVal keyS As String, ByVal nameS As String, ByVal refD As String, ByVal field As String)
        Dim oldOvr As Object
        cadV = Round(cadV, 3)
        If Math.Abs(cadV - usedV) <= tol Then Exit Sub
        oldOvr = cell.Value
        If PT_IsNum(oldOvr) Then
            If Math.Abs(CDbl(oldOvr) - cadV) > tol Then
                PT_LogAdd(MACRO_NAME, keyS, nameS, refD, "conflict", field, CDbl(oldOvr), cadV)
                nConflict = nConflict + 1
            End If
        Else
            PT_LogAdd(MACRO_NAME, keyS, nameS, refD, "override from CAD", field, usedV, cadV)
        End If
        cell.Value = cadV
        nOverride = nOverride + 1
    End Sub

    ' ================================================================= spreader
    ' The active spreader part of one side: the name holds "bottom" for the plate under the PCB.
    Private Function PickSpreader(ByVal sprHave As Dict, ByVal sprSupp As Dict, ByVal wantBottom As Boolean) As Object
        Dim res As Object = Nothing
        For Each it As Object In sprHave.Keys
            If Not CBool(sprSupp(it)) Then
                If (InStr(LCase(CStr(sprHave(it).Name2)), "bottom") > 0) = wantBottom Then res = sprHave(it)
            End If
        Next
        Return res
    End Function

    ' Range of a Spr_ or As_ name for the plate being synced (bottom plate: names live on the Spreader bottom tab)
    Private Function SRng(ByVal wb As Object, ByVal nameS As String) As Object
        If bottomSide Then Return wsSide.Names(nameS).RefersToRange
        Return wb.Names(nameS).RefersToRange
    End Function

    Private Function SNum(ByVal wb As Object, ByVal nameS As String) As Double
        Return PT_Num(SRng(wb, nameS).Cells(1, 1).Value, 0)
    End Function

    Private Sub SyncSpreader(ByVal wb As Object, ByVal c As Object, ByVal timHave As Dict, ByVal timSupp As Dict)
        Dim md As Object, ad As Object, tz As Double
        Dim fTop As Object, fUnder As Object, zTop As Double, zUnder As Double, thick As Double
        Dim mp As Object, vol As Double, surf As Double, topArea As Double, upArea As Double, contactArea As Double, holes As Integer
        Dim zuX As Double, tX As Double, dens As Double, plate As String
        plate = IIf(bottomSide, "Bottom spreader", "Spreader")
        md = c.GetModelDoc2
        ad = c.Transform2.ArrayData
        tz = CDbl(ad(11)) * 1000.0
        zuX = SNum(wb, "Spr_ZUnder")
        tX = SNum(wb, "Spr_Thickness")
        dens = SNum(wb, "Spr_Density")

        fTop = FaceByName(md, "Top")
        fUnder = FaceByName(md, "Underside")
        If fTop Is Nothing Or fUnder Is Nothing Then
            CheckFailed(IIf(bottomSide, "spreader_bottom", "spreader"), plate, "", "The " & LCase(plate) & " has no face named " & IIf(fTop Is Nothing, "Top", "Underside") &
                        ". Run BuildSpreader (Yes = update) to name the faces again.")
            Exit Sub
        End If
        zTop = FaceZ(fTop) + tz
        zUnder = FaceZ(fUnder) + tz
        thick = Math.Abs(zTop - zUnder)
        topArea = CDbl(fTop.GetArea) * 1000000.0
        upArea = UpFacingArea(md, bottomSide)
        holes = ClosedCircles(fTop)
        mp = md.Extension.CreateMassProperty
        vol = CDbl(mp.Volume) * 1000000000.0                     ' m3 -> mm3
        surf = CDbl(mp.SurfaceArea) * 1000000.0                  ' m2 -> mm2

        ' Excel-owned numbers
        If Math.Abs(zUnder - zuX) > tolPos Then Differs("PT_Under", plate & " underside height", "", "Z", zuX, zUnder)
        If Math.Abs(thick - tX) > tolSize Then Differs("PT_Plate", plate & " plate thickness", "", "Thickness", tX, thick)

        ' chips, pedestals and TIMs
        contactArea = SyncChips(wb, md, tz, zUnder, timHave, timSupp)

        ' clearance cuts still there
        Dim vc(,) As Object, fName As String, f As Object
        vc = ToTable(SRng(wb, "Spr_Cuts").Value)
        For r As Integer = 1 To UBound(vc, 1)
            If PT_Str(vc(r, 1 + K_KEY)) <> "" Then
                fName = "Clear_" & PT_SafeFile(PT_Str(vc(r, 1 + K_KEY)))
                f = PT_FindFeature(md, fName)
                If f Is Nothing Then
                    CheckFailed(PT_Str(vc(r, 1 + K_KEY)), PT_Str(vc(r, 1 + K_NAME)), "", "Clearance cut " & fName & " is missing. Run BuildSpreader.")
                ElseIf CBool(f.IsSuppressed) Then
                    CheckFailed(PT_Str(vc(r, 1 + K_KEY)), PT_Str(vc(r, 1 + K_NAME)), "", "Clearance cut " & fName & " is suppressed but the part is included.")
                End If
            End If
        Next

        SRng(wb, "As_Area").Value = Round(topArea, 1)
        SRng(wb, "As_TopArea").Value = Round(upArea, 1)
        SRng(wb, "As_ExposedArea").Value = Round(surf - contactArea, 1)
        SRng(wb, "As_Volume").Value = Round(vol, 1)
        SRng(wb, "As_Mass").Value = Round(vol * 0.000000001 * dens * 1000.0, 2)        ' mm3 -> m3, kg -> g
        SRng(wb, "As_Thickness").Value = Round(thick, 3)
        SRng(wb, "As_Holes").Value = holes
    End Sub

    ' Pedestal heights and TIM thicknesses as found, written to the Spreader tab; contact checks.
    ' Returns the area (mm2) of spreader faces that touch a TIM or chip (not exposed to air).
    Private Function SyncChips(ByVal wb As Object, ByVal md As Object, ByVal tz As Double, ByVal zUnder As Double,
                               ByVal timHave As Dict, ByVal timSupp As Dict) As Double
        Dim rg As Object, v(,) As Object, refD As String, pedX As Double, timX As Double, isBody As Boolean
        Dim fC As Object, zC As Double, pedF As Double, keyS As String, bx As Double(), area As Double = 0
        Dim zChipSide As Double, zPlateSide As Double
        Dim chipTop As Double, cx0 As Double, cy0 As Double, cl As Double, cw As Double, ws As Object, g As Integer
        rg = SRng(wb, "Spr_Chips")
        ws = rg.Worksheet
        v = ToTable(rg.Value)
        For r As Integer = 1 To UBound(v, 1)
            refD = PT_Str(v(r, 1 + C_REF))
            g = CInt(rg.Row) + r - 1
            ws.Cells(g, rg.Column + C_PEDFOUND).ClearContents()
            ws.Cells(g, rg.Column + C_TIMFOUND).ClearContents()
            If refD <> "" And UCase(PT_Str(v(r, 1 + C_COV))) = "YES" And PT_Str(v(r, 1 + C_CHECK)) = "OK" Then
                pedX = PT_Num(v(r, 1 + C_PED), 0)
                timX = PT_Num(v(r, 1 + C_TIMT), 0)
                isBody = (UCase(PT_Str(v(r, 1 + C_TIMMODE))) = "BODY")
                chipTop = PT_Num(v(r, 1 + C_TOP), 0)
                cx0 = PT_Num(v(r, 1 + C_X), 0) : cy0 = PT_Num(v(r, 1 + C_Y), 0)
                cl = PT_Num(v(r, 1 + C_L), 0) : cw = PT_Num(v(r, 1 + C_W), 0)

                ' contact face: pedestal bottom, or the underside when there is no pedestal
                If pedX > 0.0005 Then
                    fC = FaceByName(md, "Contact_" & PT_SafeFile(refD))
                    If fC Is Nothing Then
                        CheckFailed("Ped_" & refD, "Pedestal for " & refD, refD, "No face Contact_" & refD & ": the pedestal is missing or suppressed.")
                        zC = zUnder : pedF = 0
                    Else
                        zC = FaceZ(fC) + tz
                        pedF = IIf(bottomSide, zC - zUnder, zUnder - zC)
                        area = area + CDbl(fC.GetArea) * 1000000.0
                    End If
                Else
                    zC = zUnder : pedF = 0
                    area = area + cl * cw
                End If
                ws.Cells(g, rg.Column + C_PEDFOUND).Value = Round(pedF, 3)
                If Math.Abs(pedF - pedX) > tolSize Then Differs("Ped_" & refD, "Pedestal for " & refD, refD, "Pedestal h", pedX, pedF)

                ' TIM body between chip top and contact
                If isBody Then
                    keyS = "tim_" & LCase(PT_SafeFile(refD))
                    If Not timHave.Exists(keyS) Then
                        CheckFailed(keyS, "TIM " & refD, refD, "No TIM part for " & refD & " in the assembly. Run BuildSpreader.")
                    ElseIf CBool(timSupp(keyS)) Then
                        CheckFailed(keyS, "TIM " & refD, refD, "The TIM part for " & refD & " is suppressed.")
                    Else
                        bx = PT_CompBox(timHave(keyS))
                        If bx IsNot Nothing Then
                            ws.Cells(g, rg.Column + C_TIMFOUND).Value = Round(bx(5), 3)
                            If Math.Abs(bx(5) - timX) > tolSize Then Differs(keyS, "TIM " & refD, refD, "TIM t", timX, bx(5))
                            If bottomSide Then
                                zChipSide = bx(2) + bx(5) : zPlateSide = bx(2)
                            Else
                                zChipSide = bx(2) : zPlateSide = bx(2) + bx(5)
                            End If
                            If Math.Abs(zChipSide - chipTop) > 0.01 Then CheckFailed(keyS, "TIM " & refD, refD,
                                "TIM face on the chip is at Z " & Round(zChipSide, 3) & ", the chip face is at " & Round(chipTop, 3))
                            If Math.Abs(zPlateSide - zC) > 0.01 Then CheckFailed(keyS, "TIM " & refD, refD,
                                "TIM face on the spreader is at Z " & Round(zPlateSide, 3) & ", the spreader is at " & Round(zC, 3))
                        End If
                    End If
                ElseIf Math.Abs(CDbl(IIf(bottomSide, chipTop - zC, zC - chipTop)) - timX) > 0.01 Then
                    CheckFailed("Ped_" & refD, "Pedestal for " & refD, refD, "Gap between chip top and spreader is " &
                                Round(Math.Abs(zC - chipTop), 3) & " mm, the contact resistance assumes " & Round(timX, 3))
                End If
            End If
        Next
        Return area
    End Function

    ' ================================================================= interference
    Private Sub CheckInterference(ByVal asm As Object)
        Dim im As Object, arr As Object, comps As Object, rA As String, rB As String, nA As String, nB As String
        Try
            im = asm.InterferenceDetectionManager
            im.TreatCoincidenceAsInterference = False
            im.TreatSubAssembliesAsComponents = True
            im.IncludeMultibodyPartInterferences = False
            im.MakeInterferingPartsTransparent = False
            arr = im.GetInterferences
            If Not IsEmpty(arr) Then
                For i As Integer = 0 To UBound(arr)
                    comps = arr(i).Components
                    If UBound(comps) >= 1 Then
                        rA = PT_GetProp(comps(0).GetModelDoc2, "Role") : rB = PT_GetProp(comps(1).GetModelDoc2, "Role")
                        nA = CompLabel(comps(0)) : nB = CompLabel(comps(1))
                        If (rA = "PCB" And rB = "Part") Or (rA = "Part" And rB = "PCB") Then
                            ' through-hole and edge connectors: expected
                        Else
                            CheckFailed("interference", nA & " / " & nB, "", "Interference " & nA & " with " & nB & ", " &
                                        Round(CDbl(arr(i).Volume) * 1000000000.0, 3) & " mm3")
                        End If
                    End If
                Next
            End If
            im.Done()
        Catch ex As Exception
            notes.Add("Interference check could not run: " & ex.Message)
        End Try
    End Sub

    Private Function CompLabel(ByVal c As Object) As String
        Dim s As String = PT_GetProp(c.GetModelDoc2, "Name")
        If s = "" Then s = CStr(c.Name2)
        Return s
    End Function

    ' ================================================================= helpers
    Private Sub RestoreSuppression(ByVal have As Dict, ByVal wasSupp As Dict)
        For Each it As Object In have.Keys
            If CBool(wasSupp(it)) Then have(it).SetSuppression2(PT_swSuppressed)
        Next
    End Sub

    Private Sub CheckFailed(ByVal keyS As String, ByVal nameS As String, ByVal refD As String, ByVal what As String)
        nFailed = nFailed + 1
        If openFails.Exists(keyS & "|" & what) Then
            nRepeat = nRepeat + 1                       ' same failure is already waiting for a Yes: do not add a duplicate row
        Else
            PT_LogAdd(MACRO_NAME, keyS, nameS, refD, "check failed", what, "", "")
            openFails.Add(keyS & "|" & what, 1)
        End If
    End Sub

    Private Sub Differs(ByVal keyS As String, ByVal nameS As String, ByVal refD As String, ByVal field As String,
                        ByVal excelV As Double, ByVal cadV As Double)
        PT_LogAdd(MACRO_NAME, keyS, nameS, refD, "excel-owned value differs", field, Round(excelV, 3), Round(cadV, 3))
        nDiffers = nDiffers + 1
    End Sub

    Private Function FaceByName(ByVal md As Object, ByVal nameS As String) As Object
        Try
            Return md.GetEntityByName(nameS, swSelFACES)
        Catch
            Return Nothing
        End Try
    End Function

    ' Z (mm, part frame) of a horizontal flat face
    Private Function FaceZ(ByVal fc As Object) As Double
        Dim bb As Object = fc.GetBox
        Return (CDbl(bb(2)) + CDbl(bb(5))) * 500.0
    End Function

    Private Function UpFacingArea(ByVal md As Object, ByVal facingDown As Boolean) As Double
        Dim nv As Object, a As Double = 0
        For Each fc As Object In PT_AllFaces(md)
            nv = Nothing
            Try
                nv = fc.Normal
            Catch
            End Try
            If Not IsEmpty(nv) Then
                If (facingDown And nv(2) < -0.99) Or (Not facingDown And nv(2) > 0.99) Then a = a + CDbl(fc.GetArea)
            End If
        Next
        Return a * 1000000.0
    End Function

    ' full circles on a face = round through holes
    Private Function ClosedCircles(ByVal fc As Object) As Integer
        Dim edges As Object, crv As Object, cnt As Integer = 0
        edges = fc.GetEdges
        If IsEmpty(edges) Then Return 0
        For i As Integer = 0 To UBound(edges)
            crv = edges(i).GetCurve
            If crv IsNot Nothing Then
                If CBool(crv.IsCircle) Then
                    If edges(i).GetStartVertex Is Nothing Then cnt = cnt + 1
                End If
            End If
        Next
        Return cnt
    End Function

End Module
