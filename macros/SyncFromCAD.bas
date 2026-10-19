Attribute VB_Name = "SyncFromCAD"
Option Explicit
' =====================================================================================
' SyncFromCAD  -  SolidWorks 2023 macro (milestone M5)
'
' Reads the thermal assembly and brings what you changed in SolidWorks back to Excel.
' Works for any board. Does not change any SolidWorks file.
'
'   PCB and parts   real box in the assembly vs the Used values on CAD geometry.
'                   A difference is written as an override (CAD wins) and logged
'                   "override from CAD". If you had typed a different override,
'                   it is replaced and logged "conflict" with your old value.
'   Spreader        area, exposed area, volume, mass, thickness, holes -> Spreader tab
'                   AS BUILT block; pedestal heights and TIM thicknesses as found ->
'                   Spreader columns U and V. Differences from the Excel design are
'                   logged "excel-owned value differs" (Excel is not changed).
'   Checks          logged "check failed": covered chip without contact face, TIM not
'                   touching chip and spreader, clearance cut missing or suppressed,
'                   interference between bodies (parts going into the PCB are ignored).
'
' Project folder (page, W7): pick <prefix>_project.json at the start. Nothing is written to the project file (the
' page owns it): what differs goes to <prefix>_cad.json asBuilt (boxes read back, the plates as built, failed
' checks) and change records ("differs in SolidWorks", "check failed"). Tab 10 of the page offers to take the
' SolidWorks sizes.
' Needs modules PTCommon and PTJson in the same macro. Old workbooks: a copy of PCB_thermal_TEMPLATE.xlsx.
' =====================================================================================

Const MACRO_NAME As String = "SyncFromCAD"
Const swSelFACES As Long = 2

' CAD geometry offsets from column B
Const G_NAMEUSED As Long = 3, G_REF As Long = 4, G_OVR As Long = 13, G_USED As Long = 19, G_KEY As Long = 31
' Spreader chip table offsets from column B
Const C_REF As Long = 0, C_X As Long = 2, C_Y As Long = 3, C_L As Long = 4, C_W As Long = 5, C_TOP As Long = 6
Const C_COV As Long = 7, C_TIMMODE As Long = 9, C_TIMT As Long = 13, C_PED As Long = 14, C_CHECK As Long = 16
Const C_PEDFOUND As Long = 19, C_TIMFOUND As Long = 20
' clearance cut offsets from column B
Const K_KEY As Long = 1, K_NAME As Long = 2

Dim nOverride As Long, nConflict As Long, nDiffers As Long, nFailed As Long, nRepeat As Long
Dim openFails As Object                  ' "key|text" of check failed rows still unchecked on the Change log
Dim tolPos As Double, tolSize As Double
Dim notes As Collection
Dim bottomSide As Boolean                ' the plate being synced is the one under the PCB
Dim wsSide As Object                     ' Spreader bottom tab (its Spr_ and As_ names are tab-level)
Dim failList As Collection               ' project folder mode: texts of the failed checks, for cad.json asBuilt

Sub main()
    Set swApp = Application.SldWorks
    PT_Trace MACRO_NAME & ": started (VBA macro)."
    PT_LogReset
    Set notes = New Collection
    Set failList = New Collection
    nOverride = 0: nConflict = 0: nDiffers = 0: nFailed = 0: nRepeat = 0

    ' ---------- project folder (page) or workbook ----------
    Select Case PJ_PickSource(MACRO_NAME)
        Case "json"
            RunProject
            Exit Sub
        Case "xlsx"
        Case Else
            Exit Sub
    End Select

    Dim xl As Object, wb As Object
    Set xl = PT_GetExcel()
    Set wb = PT_PickWorkbook(xl, MACRO_NAME, "CAD geometry|Spreader|Settings|Change log", PJ_WbPath)
    If wb Is Nothing Then Exit Sub
    If Not PT_NeedNames(wb, "Geo_Table|Geo_PCB_Used|Geo_PCB_Override|Spr_Chips|Spr_Cuts|Spr_ZUnder|Spr_Thickness|Spr_Density|" & _
                            "As_Area|As_TopArea|As_ExposedArea|As_Volume|As_Mass|As_Thickness|As_Holes|As_LastSync|" & _
                            "Log_Header|Set_TolPos|Set_TolSize|Set_Prefix|Set_OutputFolder|Set_AsmName", MACRO_NAME) Then Exit Sub
    Set openFails = PT_OpenLogKeys(wb, "check failed")
    tolPos = PT_NameNum(wb, "Set_TolPos", 0.05)
    tolSize = PT_NameNum(wb, "Set_TolSize", 0.05)
    xl.Calculate

    Dim asm As Object, asmPath As String
    asmPath = PT_AsmPath(wb)
    Set asm = PT_OpenAsm(asmPath)
    If asm Is Nothing Then
        PT_Msg "SyncFromCAD did not run. No thermal assembly at" & vbCrLf & asmPath & vbCrLf & "Run BuildBoard first.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    PT_ActivateDoc asm
    PT_LeaveHeatOnly asm

    ' ---------- read everything from SolidWorks first ----------
    Dim pcbHave As Object, pcbSupp As Object, partHave As Object, partSupp As Object
    Dim sprHave As Object, sprSupp As Object, timHave As Object, timSupp As Object
    Set pcbHave = CreateObject("Scripting.Dictionary"): Set pcbSupp = CreateObject("Scripting.Dictionary")
    Set partHave = CreateObject("Scripting.Dictionary"): Set partSupp = CreateObject("Scripting.Dictionary")
    Set sprHave = CreateObject("Scripting.Dictionary"): Set sprSupp = CreateObject("Scripting.Dictionary")
    Set timHave = CreateObject("Scripting.Dictionary"): Set timSupp = CreateObject("Scripting.Dictionary")
    On Error GoTo swFail
    PT_ReadRoles asm, Array("PCB", "Part", "Spreader", "TIM"), Array(pcbHave, partHave, sprHave, timHave), _
                 Array(pcbSupp, partSupp, sprSupp, timSupp)
    RestoreSuppression pcbHave, pcbSupp
    RestoreSuppression partHave, partSupp
    RestoreSuppression sprHave, sprSupp
    RestoreSuppression timHave, timSupp
    On Error GoTo 0

    ' ---------- write to Excel ----------
    On Error GoTo xlFail
    xl.ScreenUpdating = False
    SyncPcb wb, pcbHave, pcbSupp
    SyncParts wb, partHave, partSupp
    Dim side As Long, cSpr As Object, nSpr As Long
    For side = 0 To 1
        bottomSide = (side = 1)
        Set wsSide = Nothing
        Set cSpr = PickSpreader(sprHave, sprSupp, bottomSide)
        If cSpr Is Nothing Then
            If bottomSide And nSpr = 0 Then notes.Add "No active spreader in the assembly: spreader checks skipped."
        Else
            If bottomSide Then Set wsSide = PT_Sheet(wb, "Spreader bottom")
            If bottomSide And wsSide Is Nothing Then
                notes.Add "A bottom spreader is in the assembly but the workbook has no Spreader bottom tab: bottom plate skipped."
            Else
                SyncSpreader wb, cSpr, timHave, timSupp
                SRng(wb, "As_LastSync").Value = Format(Now, "yyyy-mm-dd hh:nn")
                nSpr = nSpr + 1
            End If
        End If
    Next
    CheckInterference asm
    PT_Rng(wb, "As_LastSync").Value = Format(Now, "yyyy-mm-dd hh:nn")
    Dim nLog As Long
    nLog = PT_LogFlush(wb)
    xl.ScreenUpdating = True
    xl.Calculate
    wb.Save
    On Error GoTo 0

    Dim s As String, it As Variant, k As Long
    s = "Written and saved to:" & vbCrLf & wb.FullName & vbCrLf & "(tabs CAD geometry, Spreader, Change log)" & vbCrLf & vbCrLf & _
        "Read from " & asmPath & vbCrLf & vbCrLf & _
        nOverride & " overrides written from CAD, " & nConflict & " of them replaced a value you had typed (conflict)." & vbCrLf & _
        nDiffers & " spreader values differ from the Spreader tab design." & vbCrLf & _
        nFailed & " checks failed" & IIf(nRepeat > 0, " (" & nRepeat & " already unchecked on the Change log, not added again)", "") & "." & vbCrLf
    If nLog > 0 Then s = s & nLog & " rows added to the Change log tab. Check them there." & vbCrLf Else s = s & "Nothing to report." & vbCrLf
    For Each it In notes
        k = k + 1
        If k > 6 Then s = s & "(more notes on the Change log tab)" & vbCrLf: Exit For
        s = s & "- " & it & vbCrLf
    Next
    PT_Msg s, IIf(nFailed > 0 Or nConflict > 0, vbExclamation, vbInformation), MACRO_NAME
    Exit Sub

swFail:
    PT_Msg "SyncFromCAD stopped while reading SolidWorks: " & Err.Description & vbCrLf & "Nothing was written.", vbExclamation, MACRO_NAME
    Exit Sub
xlFail:
    xl.ScreenUpdating = True
    PT_Msg "SyncFromCAD stopped while writing to Excel: " & Err.Description & vbCrLf & _
           "The workbook was not saved. Close any cell you are editing, close the workbook without saving, and run again.", _
           vbExclamation, MACRO_NAME
End Sub

' ================================================================= project folder (page) mode
' Compares the assembly with forSolidWorks (what the page asked for on its last Save) and writes cad.json asBuilt
' and change records. The project file is never written: the page takes the SolidWorks sizes on tab 10.
Sub RunProject()
    Dim cad As Object, cadPath As String, ch As Object, it As Variant
    cadPath = PJ_Dir & "\" & PJ_File("cad")
    Set cad = PJ_ReadOwn(PJ_File("cad"), "pcb-thermal/cad")
    If cad Is Nothing Then
        PT_Msg "SyncFromCAD did not run. Nothing was changed." & vbCrLf & vbCrLf & cadPath & " could not be read:" & vbCrLf & PT_LastErr, _
               vbExclamation, MACRO_NAME
        Exit Sub
    End If
    If PJ_Obj(cad, "build") Is Nothing Or PJ_Obj(PJ_Project, "forSolidWorks") Is Nothing Then
        PT_Msg "SyncFromCAD did not run. There is no build in " & cadPath & " yet, or the page has not saved its numbers." & vbCrLf & _
               "Run Build Board first (button 3).", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    ' failed checks still waiting for a tick on the page: not added again
    Set openFails = CreateObject("Scripting.Dictionary")
    Set ch = PJ_Obj(cad, "changes")
    If Not ch Is Nothing Then
        For Each it In ch
            If PJ_Str(it, "change", "") = "check failed" And Not PJ_Has(PJ_Project, "checks/" & PJ_Str(it, "id", "")) Then
                If Not openFails.Exists(PJ_Str(it, "key", "") & "|" & PJ_Str(it, "field", "")) Then openFails.Add PJ_Str(it, "key", "") & "|" & PJ_Str(it, "field", ""), 1
            End If
        Next
    End If
    tolPos = PJ_Num(PJ_Project, "settings/tolPosMm", 0.05)
    tolSize = PJ_Num(PJ_Project, "settings/tolSizeMm", 0.05)

    Dim asm As Object, asmPath As String
    asmPath = PJ_AsmPath()
    Set asm = PT_OpenAsm(asmPath)
    If asm Is Nothing Then
        PT_Msg "SyncFromCAD did not run. No thermal assembly at" & vbCrLf & asmPath & vbCrLf & "Run Build Board first.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    PT_ActivateDoc asm
    PT_LeaveHeatOnly asm

    Dim pcbHave As Object, pcbSupp As Object, partHave As Object, partSupp As Object
    Dim sprHave As Object, sprSupp As Object, timHave As Object, timSupp As Object
    Set pcbHave = CreateObject("Scripting.Dictionary"): Set pcbSupp = CreateObject("Scripting.Dictionary")
    Set partHave = CreateObject("Scripting.Dictionary"): Set partSupp = CreateObject("Scripting.Dictionary")
    Set sprHave = CreateObject("Scripting.Dictionary"): Set sprSupp = CreateObject("Scripting.Dictionary")
    Set timHave = CreateObject("Scripting.Dictionary"): Set timSupp = CreateObject("Scripting.Dictionary")
    Dim ab As Object, bodiesOut As Collection, plates As Object, side As Long, cSpr As Object, pl As Object, plOut As Object
    Set ab = PJ_Dict()
    Set bodiesOut = New Collection
    Set plates = PJ_Dict()
    On Error GoTo swFail
    PT_ReadRoles asm, Array("PCB", "Part", "Spreader", "TIM"), Array(pcbHave, partHave, sprHave, timHave), _
                 Array(pcbSupp, partSupp, sprSupp, timSupp)
    RestoreSuppression pcbHave, pcbSupp
    RestoreSuppression partHave, partSupp
    RestoreSuppression sprHave, sprSupp
    RestoreSuppression timHave, timSupp
    SyncPcbJson pcbHave, pcbSupp
    SyncPartsJson partHave, partSupp, bodiesOut
    For side = 0 To 1
        bottomSide = (side = 1)
        Set pl = PJ_Obj(PJ_Project, "forSolidWorks/plates/" & IIf(bottomSide, "bottom", "top"))
        Set cSpr = PickSpreader(sprHave, sprSupp, bottomSide)
        If cSpr Is Nothing Then
            If Not pl Is Nothing Then notes.Add IIf(bottomSide, "Bottom", "Top") & " plate is enabled on the page but there is no active spreader part for it: run Build Spreader."
        ElseIf pl Is Nothing Then
            notes.Add "A " & IIf(bottomSide, "bottom", "top") & " spreader is in the assembly but its plate is not enabled on the page: it was not checked."
        Else
            Set plOut = SyncSpreaderJson(cSpr, pl, timHave, timSupp)
            If Not plOut Is Nothing Then PJ_Put plates, IIf(bottomSide, "bottom", "top"), plOut
        End If
    Next
    CheckInterference asm
    On Error GoTo 0

    ' ---------- cad.json ----------
    Dim fl As Collection, nLog As Long
    Set fl = New Collection
    For Each it In failList: fl.Add CStr(it): Next
    PJ_Put ab, "date", PJ_Now()
    If plates.Count > 0 Then PJ_Put ab, "plates", plates
    PJ_Put ab, "bodies", bodiesOut
    PJ_Put ab, "failedChecks", fl
    PJ_Put cad, "asBuilt", ab
    nLog = PJ_FlushChanges(cad)
    If Not PJ_SaveOwn(cad, PJ_File("cad"), MACRO_NAME) Then
        PT_Msg "SyncFromCAD read the assembly, but " & cadPath & " could not be written:" & vbCrLf & PT_LastErr, vbExclamation, MACRO_NAME
        Exit Sub
    End If

    Dim s As String, k As Long
    s = "Written to:" & vbCrLf & cadPath & vbCrLf & "(the page reloads it on its own: tab 7 Spreader and tab 10 Changes)" & vbCrLf & vbCrLf & _
        "Read from " & asmPath & vbCrLf & vbCrLf & _
        bodiesOut.Count & " bodies differ from the page." & IIf(bodiesOut.Count > 0, " Take their SolidWorks sizes on tab 10, or run Build Board to put the page's sizes back.", "") & vbCrLf & _
        nDiffers & " spreader values differ from tab 7." & vbCrLf & _
        nFailed & " checks failed" & IIf(nRepeat > 0, " (" & nRepeat & " already waiting for a tick, not added again)", "") & "." & vbCrLf
    If nLog > 0 Then s = s & nLog & " change records written. Tick them on tab 10 of the page." & vbCrLf Else s = s & "Nothing to report." & vbCrLf
    For Each it In notes
        k = k + 1
        If k > 6 Then s = s & "(more notes in the log)" & vbCrLf: Exit For
        s = s & "- " & it & vbCrLf
        PT_Trace MACRO_NAME & ": " & it
    Next
    PT_Msg s, IIf(nFailed > 0 Or bodiesOut.Count > 0, vbExclamation, vbInformation), MACRO_NAME
    Exit Sub

swFail:
    PT_Msg "SyncFromCAD stopped while reading SolidWorks: " & Err.Description & vbCrLf & "Nothing was written.", vbExclamation, MACRO_NAME
End Sub

Sub SyncPcbJson(ByVal have As Object, ByVal wasSupp As Object)
    Dim c As Object, bx As Variant, k As Long, fields As Variant, want As Double, it As Variant
    fields = Array("l", "w", "h")
    For Each it In have.Keys
        If Not wasSupp(it) Then Set c = have(it)
    Next
    If c Is Nothing Then
        CheckFailed "pcb", "PCB", "", "No active PCB part in the thermal assembly"
        Exit Sub
    End If
    bx = PT_CompBox(c)
    If IsEmpty(bx) Then Exit Sub
    If Abs(bx(0)) > tolPos Or Abs(bx(1)) > tolPos Or Abs(bx(2)) > tolPos Then
        CheckFailed "pcb", "PCB", "", "The PCB corner is not at the origin any more (" & Round(bx(0), 3) & ", " & _
                    Round(bx(1), 3) & ", " & Round(bx(2), 3) & "). Every part is placed from that corner: move it back."
    End If
    For k = 0 To 2
        want = PJ_Num(PJ_Project, "forSolidWorks/pcb/" & fields(k), 0)
        If Abs(bx(3 + k) - want) > tolSize Then
            PT_LogAdd MACRO_NAME, "pcb", "PCB", "", "differs in SolidWorks", "PCB " & UCase(fields(k)), Round(want, 3), Round(bx(3 + k), 3)
            notes.Add "PCB " & UCase(fields(k)) & " is " & Round(bx(3 + k), 3) & " mm in SolidWorks, " & Round(want, 3) & " mm on the page (tab 3, PCB size)."
        End If
    Next
End Sub

' Each built part against the box the page asked for (cut at the board surface, as Build Board builds it).
' A body that differs goes to asBuilt bodies with the box read from SolidWorks.
Sub SyncPartsJson(ByVal have As Object, ByVal wasSupp As Object, ByVal out As Collection)
    Dim bl As Object, it As Variant, b As Variant, wantB As Object, keyS As String, bx As Variant, k As Long, diff As Boolean
    Dim usedB(5) As Double, pcbT As Double, fields As Variant, tol As Double, rec As Object, nameS As String
    fields = Array("xMin", "yMin", "zMin", "l", "w", "h")
    pcbT = PJ_Num(PJ_Project, "forSolidWorks/pcb/h", 0)
    Set bl = PJ_Obj(PJ_Project, "forSolidWorks/bodies")
    For Each it In have.Keys
        If Not wasSupp(it) Then
            keyS = CStr(it)
            Set wantB = Nothing
            If Not bl Is Nothing Then
                For Each b In bl
                    If PJ_Str(b, "key", "") = keyS Then Set wantB = b: Exit For
                Next
            End If
            If wantB Is Nothing Then
                notes.Add "Part " & keyS & " is in the assembly but not included on the page. Run Build Board to suppress it."
            Else
                nameS = PJ_Str(wantB, "name", keyS)
                bx = PT_CompBox(have(it))
                If Not IsEmpty(bx) And PJ_Num(wantB, "rotDeg", 0) = 0 Then
                    For k = 0 To 5: usedB(k) = PJ_Num(wantB, "box/" & fields(k), 0): Next
                    PT_ClipToBoard usedB, pcbT
                    diff = False
                    For k = 0 To 5
                        If k < 3 Then tol = tolPos Else tol = tolSize
                        If Abs(bx(k) - usedB(k)) > tol Then
                            diff = True
                            PT_LogAdd MACRO_NAME, keyS, nameS, PJ_Str(wantB, "refDes", ""), "differs in SolidWorks", CStr(fields(k)), Round(usedB(k), 3), Round(bx(k), 3)
                        End If
                    Next
                    If diff Then
                        Set rec = PJ_Dict()
                        PJ_Put rec, "key", keyS
                        PJ_Put rec, "box", PJ_Box(bx(0), bx(1), bx(2), bx(3), bx(4), bx(5))
                        out.Add rec
                    End If
                End If
            End If
        End If
    Next
End Sub

' One plate against forSolidWorks plates top or bottom (pl). Returns the cad.json asBuilt plate record.
Function SyncSpreaderJson(ByVal c As Object, ByVal pl As Object, ByVal timHave As Object, ByVal timSupp As Object) As Object
    Dim md As Object, ad As Variant, tz As Double, plate As String
    Dim fTop As Object, fUnder As Object, zTop As Double, zUnder As Double, thick As Double
    Dim mp As Object, vol As Double, surf As Double, topArea As Double, upArea As Double, contactArea As Double, holes As Long
    Dim zuX As Double, tX As Double, dens As Double, d As Object, chipsOut As Collection
    plate = IIf(bottomSide, "Bottom spreader", "Spreader")
    Set md = c.GetModelDoc2
    ad = c.Transform2.ArrayData
    tz = ad(11) * 1000#
    zuX = PJ_Num(pl, "faceZMm", 0)
    tX = PJ_Num(pl, "thicknessMm", 0)
    dens = PJ_Num(pl, "densityKgM3", 2700)

    Set fTop = FaceByName(md, "Top")
    Set fUnder = FaceByName(md, "Underside")
    If fTop Is Nothing Or fUnder Is Nothing Then
        CheckFailed IIf(bottomSide, "spreader_bottom", "spreader"), plate, "", "The " & LCase(plate) & " has no face named " & IIf(fTop Is Nothing, "Top", "Underside") & _
                    ". Run Build Spreader again to name the faces."
        Exit Function
    End If
    zTop = FaceZ(fTop) + tz
    zUnder = FaceZ(fUnder) + tz
    thick = Abs(zTop - zUnder)
    topArea = fTop.GetArea * 1000000#
    upArea = UpFacingArea(md, bottomSide)
    holes = ClosedCircles(fTop)
    Set mp = md.Extension.CreateMassProperty
    vol = mp.Volume * 1000000000#
    surf = mp.SurfaceArea * 1000000#
    If Abs(zUnder - zuX) > tolPos Then Differs "PT_Under", plate & " underside height", "", "Z", zuX, zUnder
    If Abs(thick - tX) > tolSize Then Differs "PT_Plate", plate & " plate thickness", "", "Thickness", tX, thick

    Set chipsOut = New Collection
    contactArea = SyncChipsJson(md, pl, tz, zUnder, timHave, timSupp, chipsOut)

    Dim cl As Object, it As Variant, fName As String, f As Object
    Set cl = PJ_Obj(pl, "cuts")
    If Not cl Is Nothing Then
        For Each it In cl
            fName = "Clear_" & PT_SafeFile(PJ_Str(it, "key", ""))
            Set f = PT_FindFeature(md, fName)
            If f Is Nothing Then
                CheckFailed PJ_Str(it, "key", ""), PJ_Str(it, "name", ""), "", "Clearance cut " & fName & " is missing. Run Build Spreader."
            ElseIf f.IsSuppressed Then
                CheckFailed PJ_Str(it, "key", ""), PJ_Str(it, "name", ""), "", "Clearance cut " & fName & " is suppressed but the part is included."
            End If
        Next
    End If

    Set d = PJ_Dict()
    PJ_Put d, "outlineAreaMm2", Round(topArea, 1)
    PJ_Put d, "topAreaMm2", Round(upArea, 1)
    PJ_Put d, "exposedAreaMm2", Round(PJ_Pos(surf - contactArea), 1)
    PJ_Put d, "volumeMm3", Round(vol, 1)
    PJ_Put d, "massG", Round(vol * 0.000000001 * dens * 1000#, 2)
    PJ_Put d, "thicknessMm", Round(thick, 3)
    PJ_Put d, "holes", holes
    PJ_Put d, "chips", chipsOut
    Set SyncSpreaderJson = d
End Function

' Pedestal heights and TIM thicknesses as found, and the contact checks, for the covered chips of pl.
' Returns the area (mm2) of plate faces that touch a TIM or chip (not exposed to air).
Function SyncChipsJson(ByVal md As Object, ByVal pl As Object, ByVal tz As Double, ByVal zUnder As Double, _
                       ByVal timHave As Object, ByVal timSupp As Object, ByVal out As Collection) As Double
    Dim cl As Object, bl As Object, it As Variant, b As Variant, bx As Object, refD As String, keyS As String
    Dim pedX As Double, timX As Double, isBody As Boolean, chipTop As Double, cl0 As Double, cw0 As Double
    Dim fC As Object, zC As Double, pedF As Double, area As Double, tb As Variant, timKey As String, timF As Double
    Dim zChipSide As Double, zPlateSide As Double, rec As Object
    Set cl = PJ_Obj(pl, "chips")
    Set bl = PJ_Obj(PJ_Project, "forSolidWorks/bodies")
    If cl Is Nothing Then Exit Function
    For Each it In cl
        refD = PJ_Str(it, "refDes", "")
        keyS = PJ_Str(it, "key", "")
        Set bx = Nothing
        If Not bl Is Nothing Then
            For Each b In bl
                If PJ_Str(b, "key", "") = keyS Then Set bx = PJ_Obj(b, "box"): Exit For
            Next
        End If
        If refD <> "" And Not bx Is Nothing Then
            pedX = PJ_Num(it, "pedestalMm", 0)
            timX = PJ_Num(it, "timTMm", 0)
            isBody = (UCase(PJ_Str(it, "timMode", "Body")) = "BODY")
            If bottomSide Then chipTop = PJ_Num(bx, "zMin", 0) Else chipTop = PJ_Num(bx, "zMin", 0) + PJ_Num(bx, "h", 0)
            cl0 = PJ_Num(bx, "l", 0): cw0 = PJ_Num(bx, "w", 0)
            timF = -1
            If pedX > 0.0005 Then
                Set fC = FaceByName(md, "Contact_" & PT_SafeFile(refD))
                If fC Is Nothing Then
                    CheckFailed "Ped_" & refD, "Pedestal for " & refD, refD, "No face Contact_" & refD & ": the pedestal is missing or suppressed."
                    zC = zUnder: pedF = 0
                Else
                    zC = FaceZ(fC) + tz
                    pedF = IIf(bottomSide, zC - zUnder, zUnder - zC)
                    area = area + fC.GetArea * 1000000#
                End If
            Else
                zC = zUnder: pedF = 0
                area = area + cl0 * cw0
            End If
            If Abs(pedF - pedX) > tolSize Then Differs "Ped_" & refD, "Pedestal for " & refD, refD, "Pedestal h", pedX, pedF
            If isBody Then
                timKey = "tim_" & LCase(PT_SafeFile(refD))
                If Not timHave.Exists(timKey) Then
                    CheckFailed timKey, "TIM " & refD, refD, "No TIM part for " & refD & " in the assembly. Run Build Spreader."
                ElseIf timSupp(timKey) Then
                    CheckFailed timKey, "TIM " & refD, refD, "The TIM part for " & refD & " is suppressed."
                Else
                    tb = PT_CompBox(timHave(timKey))
                    If Not IsEmpty(tb) Then
                        timF = tb(5)
                        If Abs(tb(5) - timX) > tolSize Then Differs timKey, "TIM " & refD, refD, "TIM t", timX, tb(5)
                        If bottomSide Then
                            zChipSide = tb(2) + tb(5): zPlateSide = tb(2)
                        Else
                            zChipSide = tb(2): zPlateSide = tb(2) + tb(5)
                        End If
                        If Abs(zChipSide - chipTop) > 0.01 Then CheckFailed timKey, "TIM " & refD, refD, _
                            "TIM face on the chip is at Z " & Round(zChipSide, 3) & ", the chip face is at " & Round(chipTop, 3)
                        If Abs(zPlateSide - zC) > 0.01 Then CheckFailed timKey, "TIM " & refD, refD, _
                            "TIM face on the spreader is at Z " & Round(zPlateSide, 3) & ", the spreader is at " & Round(zC, 3)
                    End If
                End If
            ElseIf Abs(IIf(bottomSide, chipTop - zC, zC - chipTop) - timX) > 0.01 Then
                CheckFailed "Ped_" & refD, "Pedestal for " & refD, refD, "Gap between chip top and spreader is " & _
                            Round(Abs(zC - chipTop), 3) & " mm, the contact resistance assumes " & Round(timX, 3)
            End If
            Set rec = PJ_Dict()
            PJ_Put rec, "refDes", refD
            PJ_Put rec, "pedestalMm", Round(IIf(pedF < 0, 0, pedF), 3)
            If timF >= 0 Then PJ_Put rec, "timTMm", Round(timF, 3)
            out.Add rec
        End If
    Next
    SyncChipsJson = area
End Function

' ================================================================= PCB and parts
Sub SyncPcb(ByVal wb As Object, ByVal have As Object, ByVal wasSupp As Object)
    Dim c As Object, bx As Variant, usedR As Object, ovrR As Object, k As Long, fields As Variant, it As Variant
    fields = Array("Length X", "Width Y", "Thickness Z")
    For Each it In have.Keys
        If Not wasSupp(it) Then Set c = have(it)
    Next
    If c Is Nothing Then
        CheckFailed "pcb", "PCB", "", "No active PCB part in the thermal assembly"
        Exit Sub
    End If
    bx = PT_CompBox(c)
    If IsEmpty(bx) Then Exit Sub
    If Abs(bx(0)) > tolPos Or Abs(bx(1)) > tolPos Or Abs(bx(2)) > tolPos Then
        CheckFailed "pcb", "PCB", "", "The PCB corner is not at the origin any more (" & Round(bx(0), 3) & ", " & _
                    Round(bx(1), 3) & ", " & Round(bx(2), 3) & "). Every part is placed from that corner: move it back."
    End If
    Set usedR = PT_Rng(wb, "Geo_PCB_Used")
    Set ovrR = PT_Rng(wb, "Geo_PCB_Override")
    For k = 0 To 2
        WriteOverride ovrR.Cells(1 + k, 1), PT_Num(usedR.Cells(1 + k, 1).Value, 0), bx(3 + k), tolSize, "pcb", "PCB", "", CStr(fields(k))
    Next
End Sub

Sub SyncParts(ByVal wb As Object, ByVal have As Object, ByVal wasSupp As Object)
    Dim rg As Object, v As Variant, r As Long, rowOf As Object, it As Variant, c As Object, bx As Variant
    Dim k As Long, g As Long, fields As Variant, tol As Double, nameS As String, refD As String
    Dim usedB(5) As Double, clipB(5) As Double, pcbT As Double
    fields = Array("X min", "Y min", "Z min", "L", "W", "H")
    pcbT = PT_Num(PT_Rng(wb, "Geo_PCB_Used").Cells(3, 1).Value, 0)
    Set rg = PT_Rng(wb, "Geo_Table")
    v = rg.Value
    Set rowOf = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(v, 1)
        If PT_Str(v(r, 1 + G_KEY)) <> "" Then
            If Not rowOf.Exists(PT_Str(v(r, 1 + G_KEY))) Then rowOf.Add PT_Str(v(r, 1 + G_KEY)), r
        End If
    Next
    For Each it In have.Keys
        If Not wasSupp(it) Then
            Set c = have(it)
            If Not rowOf.Exists(it) Then
                notes.Add "Part " & it & " is in the assembly but has no row on CAD geometry. Run BuildBoard to suppress it."
            Else
                r = rowOf(it)
                g = rg.Row + r - 1
                nameS = PT_Str(v(r, 1 + G_NAMEUSED))
                refD = PT_Str(v(r, 1 + G_REF))
                bx = PT_CompBox(c)
                If Not IsEmpty(bx) Then
                    ' BuildBoard cuts boxes at the board surface: compare with the same cut, write back only real edits
                    For k = 0 To 5: usedB(k) = PT_Num(v(r, 1 + G_USED + k), 0): clipB(k) = usedB(k): Next
                    PT_ClipToBoard clipB, pcbT
                    For k = 0 To 5
                        If k < 3 Then tol = tolPos Else tol = tolSize
                        If Abs(bx(k) - clipB(k)) > tol Then
                            WriteOverride rg.Worksheet.Cells(g, rg.Column + G_OVR + k), usedB(k), bx(k), tol, _
                                          CStr(it), nameS, refD, CStr(fields(k))
                        End If
                    Next
                End If
            End If
        End If
    Next
End Sub

' CAD wins: writes the CAD value into the override cell when it differs from the used value.
Sub WriteOverride(ByVal cell As Object, ByVal usedV As Double, ByVal cadV As Double, ByVal tol As Double, _
                  ByVal keyS As String, ByVal nameS As String, ByVal refD As String, ByVal field As String)
    Dim oldOvr As Variant
    cadV = Round(cadV, 3)
    If Abs(cadV - usedV) <= tol Then Exit Sub
    oldOvr = cell.Value
    If PT_IsNum(oldOvr) Then
        If Abs(CDbl(oldOvr) - cadV) > tol Then
            PT_LogAdd MACRO_NAME, keyS, nameS, refD, "conflict", field, CDbl(oldOvr), cadV
            nConflict = nConflict + 1
        End If
    Else
        PT_LogAdd MACRO_NAME, keyS, nameS, refD, "override from CAD", field, usedV, cadV
    End If
    cell.Value = cadV
    nOverride = nOverride + 1
End Sub

' ================================================================= spreader
' The active spreader part of one side: the name holds "bottom" for the plate under the PCB.
Function PickSpreader(ByVal sprHave As Object, ByVal sprSupp As Object, ByVal wantBottom As Boolean) As Object
    Dim it As Variant
    For Each it In sprHave.Keys
        If Not sprSupp(it) Then
            If (InStr(LCase(sprHave(it).Name2), "bottom") > 0) = wantBottom Then Set PickSpreader = sprHave(it)
        End If
    Next
End Function

' Range of a Spr_ or As_ name for the plate being synced (bottom plate: names live on the Spreader bottom tab)
Function SRng(ByVal wb As Object, ByVal nameS As String) As Object
    If bottomSide Then Set SRng = wsSide.Names(nameS).RefersToRange Else Set SRng = wb.Names(nameS).RefersToRange
End Function

Function SNum(ByVal wb As Object, ByVal nameS As String) As Double
    SNum = PT_Num(SRng(wb, nameS).Cells(1, 1).Value, 0)
End Function

Sub SyncSpreader(ByVal wb As Object, ByVal c As Object, ByVal timHave As Object, ByVal timSupp As Object)
    Dim md As Object, ad As Variant, tz As Double
    Dim fTop As Object, fUnder As Object, zTop As Double, zUnder As Double, thick As Double
    Dim mp As Object, vol As Double, surf As Double, topArea As Double, upArea As Double, contactArea As Double, holes As Long
    Dim zuX As Double, tX As Double, dens As Double, plate As String
    plate = IIf(bottomSide, "Bottom spreader", "Spreader")
    Set md = c.GetModelDoc2
    ad = c.Transform2.ArrayData
    tz = ad(11) * 1000#
    zuX = SNum(wb, "Spr_ZUnder")
    tX = SNum(wb, "Spr_Thickness")
    dens = SNum(wb, "Spr_Density")

    Set fTop = FaceByName(md, "Top")
    Set fUnder = FaceByName(md, "Underside")
    If fTop Is Nothing Or fUnder Is Nothing Then
        CheckFailed IIf(bottomSide, "spreader_bottom", "spreader"), plate, "", "The " & LCase(plate) & " has no face named " & IIf(fTop Is Nothing, "Top", "Underside") & _
                    ". Run BuildSpreader (Yes = update) to name the faces again."
        Exit Sub
    End If
    zTop = FaceZ(fTop) + tz
    zUnder = FaceZ(fUnder) + tz
    thick = Abs(zTop - zUnder)
    topArea = fTop.GetArea * 1000000#
    upArea = UpFacingArea(md, bottomSide)
    holes = ClosedCircles(fTop)
    Set mp = md.Extension.CreateMassProperty
    vol = mp.Volume * 1000000000#                     ' m3 -> mm3
    surf = mp.SurfaceArea * 1000000#                  ' m2 -> mm2

    ' Excel-owned numbers
    If Abs(zUnder - zuX) > tolPos Then Differs "PT_Under", plate & " underside height", "", "Z", zuX, zUnder
    If Abs(thick - tX) > tolSize Then Differs "PT_Plate", plate & " plate thickness", "", "Thickness", tX, thick

    ' chips, pedestals and TIMs
    contactArea = SyncChips(wb, md, tz, zUnder, timHave, timSupp)

    ' clearance cuts still there
    Dim vc As Variant, r As Long, fName As String, f As Object
    vc = SRng(wb, "Spr_Cuts").Value
    For r = 1 To UBound(vc, 1)
        If PT_Str(vc(r, 1 + K_KEY)) <> "" Then
            fName = "Clear_" & PT_SafeFile(PT_Str(vc(r, 1 + K_KEY)))
            Set f = PT_FindFeature(md, fName)
            If f Is Nothing Then
                CheckFailed PT_Str(vc(r, 1 + K_KEY)), PT_Str(vc(r, 1 + K_NAME)), "", "Clearance cut " & fName & " is missing. Run BuildSpreader."
            ElseIf f.IsSuppressed Then
                CheckFailed PT_Str(vc(r, 1 + K_KEY)), PT_Str(vc(r, 1 + K_NAME)), "", "Clearance cut " & fName & " is suppressed but the part is included."
            End If
        End If
    Next

    SRng(wb, "As_Area").Value = Round(topArea, 1)
    SRng(wb, "As_TopArea").Value = Round(upArea, 1)
    SRng(wb, "As_ExposedArea").Value = Round(surf - contactArea, 1)
    SRng(wb, "As_Volume").Value = Round(vol, 1)
    SRng(wb, "As_Mass").Value = Round(vol * 0.000000001 * dens * 1000#, 2)        ' mm3 -> m3, kg -> g
    SRng(wb, "As_Thickness").Value = Round(thick, 3)
    SRng(wb, "As_Holes").Value = holes
End Sub

' Pedestal heights and TIM thicknesses as found, written to the Spreader tab; contact checks.
' Returns the area (mm2) of spreader faces that touch a TIM or chip (not exposed to air).
Function SyncChips(ByVal wb As Object, ByVal md As Object, ByVal tz As Double, ByVal zUnder As Double, _
                   ByVal timHave As Object, ByVal timSupp As Object) As Double
    Dim rg As Object, v As Variant, r As Long, refD As String, pedX As Double, timX As Double, isBody As Boolean
    Dim fC As Object, zC As Double, pedF As Double, keyS As String, bx As Variant, area As Double
    Dim zChipSide As Double, zPlateSide As Double
    Dim chipTop As Double, cx0 As Double, cy0 As Double, cl As Double, cw As Double, ws As Object, g As Long
    Set rg = SRng(wb, "Spr_Chips")
    Set ws = rg.Worksheet
    v = rg.Value
    For r = 1 To UBound(v, 1)
        refD = PT_Str(v(r, 1 + C_REF))
        g = rg.Row + r - 1
        ws.Cells(g, rg.Column + C_PEDFOUND).ClearContents
        ws.Cells(g, rg.Column + C_TIMFOUND).ClearContents
        If refD <> "" And UCase(PT_Str(v(r, 1 + C_COV))) = "YES" And PT_Str(v(r, 1 + C_CHECK)) = "OK" Then
            pedX = PT_Num(v(r, 1 + C_PED), 0)
            timX = PT_Num(v(r, 1 + C_TIMT), 0)
            isBody = (UCase(PT_Str(v(r, 1 + C_TIMMODE))) = "BODY")
            chipTop = PT_Num(v(r, 1 + C_TOP), 0)
            cx0 = PT_Num(v(r, 1 + C_X), 0): cy0 = PT_Num(v(r, 1 + C_Y), 0)
            cl = PT_Num(v(r, 1 + C_L), 0): cw = PT_Num(v(r, 1 + C_W), 0)

            ' contact face: pedestal bottom, or the underside when there is no pedestal
            If pedX > 0.0005 Then
                Set fC = FaceByName(md, "Contact_" & PT_SafeFile(refD))
                If fC Is Nothing Then
                    CheckFailed "Ped_" & refD, "Pedestal for " & refD, refD, "No face Contact_" & refD & ": the pedestal is missing or suppressed."
                    zC = zUnder: pedF = 0
                Else
                    zC = FaceZ(fC) + tz
                    pedF = IIf(bottomSide, zC - zUnder, zUnder - zC)
                    area = area + fC.GetArea * 1000000#
                End If
            Else
                zC = zUnder: pedF = 0
                area = area + cl * cw
            End If
            ws.Cells(g, rg.Column + C_PEDFOUND).Value = Round(pedF, 3)
            If Abs(pedF - pedX) > tolSize Then Differs "Ped_" & refD, "Pedestal for " & refD, refD, "Pedestal h", pedX, pedF

            ' TIM body between chip top and contact
            If isBody Then
                keyS = "tim_" & LCase(PT_SafeFile(refD))
                If Not timHave.Exists(keyS) Then
                    CheckFailed keyS, "TIM " & refD, refD, "No TIM part for " & refD & " in the assembly. Run BuildSpreader."
                ElseIf timSupp(keyS) Then
                    CheckFailed keyS, "TIM " & refD, refD, "The TIM part for " & refD & " is suppressed."
                Else
                    bx = PT_CompBox(timHave(keyS))
                    If Not IsEmpty(bx) Then
                        ws.Cells(g, rg.Column + C_TIMFOUND).Value = Round(bx(5), 3)
                        If Abs(bx(5) - timX) > tolSize Then Differs keyS, "TIM " & refD, refD, "TIM t", timX, bx(5)
                        If bottomSide Then
                            zChipSide = bx(2) + bx(5): zPlateSide = bx(2)
                        Else
                            zChipSide = bx(2): zPlateSide = bx(2) + bx(5)
                        End If
                        If Abs(zChipSide - chipTop) > 0.01 Then CheckFailed keyS, "TIM " & refD, refD, _
                            "TIM face on the chip is at Z " & Round(zChipSide, 3) & ", the chip face is at " & Round(chipTop, 3)
                        If Abs(zPlateSide - zC) > 0.01 Then CheckFailed keyS, "TIM " & refD, refD, _
                            "TIM face on the spreader is at Z " & Round(zPlateSide, 3) & ", the spreader is at " & Round(zC, 3)
                    End If
                End If
            ElseIf Abs(IIf(bottomSide, chipTop - zC, zC - chipTop) - timX) > 0.01 Then
                CheckFailed "Ped_" & refD, "Pedestal for " & refD, refD, "Gap between chip top and spreader is " & _
                            Round(Abs(zC - chipTop), 3) & " mm, the contact resistance assumes " & Round(timX, 3)
            End If
        End If
    Next
    SyncChips = area
End Function

' ================================================================= interference
Sub CheckInterference(ByVal asm As Object)
    Dim im As Object, arr As Variant, i As Long, comps As Variant, rA As String, rB As String, nA As String, nB As String
    On Error GoTo skip
    Set im = asm.InterferenceDetectionManager
    im.TreatCoincidenceAsInterference = False
    im.TreatSubAssembliesAsComponents = True
    im.IncludeMultibodyPartInterferences = False
    im.MakeInterferingPartsTransparent = False
    arr = im.GetInterferences
    If Not IsEmpty(arr) Then
        For i = 0 To UBound(arr)
            comps = arr(i).Components
            If UBound(comps) >= 1 Then
                rA = PT_GetProp(comps(0).GetModelDoc2, "Role"): rB = PT_GetProp(comps(1).GetModelDoc2, "Role")
                nA = CompLabel(comps(0)): nB = CompLabel(comps(1))
                If (rA = "PCB" And rB = "Part") Or (rA = "Part" And rB = "PCB") Then
                    ' through-hole and edge connectors: expected
                Else
                    CheckFailed "interference", nA & " / " & nB, "", "Interference " & nA & " with " & nB & ", " & _
                                Round(arr(i).Volume * 1000000000#, 3) & " mm3"
                End If
            End If
        Next
    End If
    im.Done
    Exit Sub
skip:
    notes.Add "Interference check could not run: " & Err.Description
End Sub

Function CompLabel(ByVal c As Object) As String
    Dim s As String
    s = PT_GetProp(c.GetModelDoc2, "Name")
    If s = "" Then s = c.Name2
    CompLabel = s
End Function

' ================================================================= helpers
Sub RestoreSuppression(ByVal have As Object, ByVal wasSupp As Object)
    Dim it As Variant
    For Each it In have.Keys
        If wasSupp(it) Then have(it).SetSuppression2 PT_swSuppressed
    Next
End Sub

Sub CheckFailed(ByVal keyS As String, ByVal nameS As String, ByVal refD As String, ByVal what As String)
    nFailed = nFailed + 1
    failList.Add nameS & ": " & what
    If openFails.Exists(keyS & "|" & what) Then
        nRepeat = nRepeat + 1                       ' same failure is already waiting for a Yes: do not add a duplicate row
    Else
        PT_LogAdd MACRO_NAME, keyS, nameS, refD, "check failed", what, "", ""
        openFails.Add keyS & "|" & what, 1
    End If
End Sub

Sub Differs(ByVal keyS As String, ByVal nameS As String, ByVal refD As String, ByVal field As String, _
            ByVal excelV As Double, ByVal cadV As Double)
    PT_LogAdd MACRO_NAME, keyS, nameS, refD, "excel-owned value differs", field, Round(excelV, 3), Round(cadV, 3)
    nDiffers = nDiffers + 1
End Sub

Function FaceByName(ByVal md As Object, ByVal nameS As String) As Object
    On Error Resume Next
    Set FaceByName = md.GetEntityByName(nameS, swSelFACES)
End Function

' Z (mm, part frame) of a horizontal flat face
Function FaceZ(ByVal fc As Object) As Double
    Dim bb As Variant
    bb = fc.GetBox
    FaceZ = (bb(2) + bb(5)) * 500#
End Function

Function UpFacingArea(ByVal md As Object, ByVal facingDown As Boolean) As Double
    Dim faces As Collection, fc As Object, nv As Variant, a As Double
    Set faces = PT_AllFaces(md)
    For Each fc In faces
        nv = Empty
        On Error Resume Next
        nv = fc.Normal
        On Error GoTo 0
        If Not IsEmpty(nv) Then
            If (facingDown And nv(2) < -0.99) Or (Not facingDown And nv(2) > 0.99) Then a = a + fc.GetArea
        End If
    Next
    UpFacingArea = a * 1000000#
End Function

' full circles on a face = round through holes
Function ClosedCircles(ByVal fc As Object) As Long
    Dim edges As Variant, i As Long, crv As Object, n As Long
    edges = fc.GetEdges
    If IsEmpty(edges) Then Exit Function
    For i = 0 To UBound(edges)
        Set crv = edges(i).GetCurve
        If Not crv Is Nothing Then
            If crv.IsCircle Then
                If edges(i).GetStartVertex Is Nothing Then n = n + 1
            End If
        End If
    Next
    ClosedCircles = n
End Function
