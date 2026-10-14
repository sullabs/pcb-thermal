Attribute VB_Name = "SetupWorkbook"
Option Explicit
' =====================================================================================
' SetupWorkbook  -  SolidWorks 2023 macro (milestone M1)
'
' Adds what the other macros need to a pcb_thermal_survey workbook:
'   tabs Settings, Change log, Spreader, Sim setup, Sim results,
'   hidden tab STEP import (previous),
'   CAD geometry columns AF (Last change), AG (Key, hidden), AH (Cut #, hidden),
'   STEP import column K (Key),
'   a workbook name for every cell a macro reads or writes.
'
' Only needed for a workbook made before PCB_thermal_TEMPLATE.xlsx. On a copy of the template
' (it has the name Tpl_Version) it only fills empty chip lists and stops.
' Safe to run again at any time: labels, formulas, dropdowns and names are rewritten,
' values you typed in blue cells are never overwritten.
' Works for any board: chip lists are filled from the RefDes column of CAD geometry.
' Needs module PTCommon in the same macro.
' =====================================================================================

Const GEO_FIRST As Long = 26      ' CAD geometry body rows
Const GEO_IMP_LAST As Long = 85   ' last row mirrored from STEP import
Const GEO_LAST As Long = 95       ' last manual row
Const IMP_FIRST As Long = 8       ' STEP import: PCB row
Const IMP_LAST As Long = 88
Const CHIP_FIRST As Long = 20     ' Spreader chip rows
Const CHIP_LAST As Long = 79
Const CUT_FIRST As Long = 85      ' Spreader clearance cut rows
Const CUT_LAST As Long = 144
Const SCHIP_FIRST As Long = 29    ' Sim setup chip rows
Const SCHIP_LAST As Long = 88

Dim stIn As Object, stCalc As Object, stHead As Object
Dim addedTabs As String
Dim chipsFilled As Long

Sub main()
    Set swApp = Application.SldWorks
    addedTabs = "": chipsFilled = 0
    Dim xl As Object, wb As Object, reg As Object
    Set xl = PT_GetExcel()
    Set wb = PT_PickWorkbook(xl, "SetupWorkbook", "Component register|CAD geometry|STEP import|Thermal solution")
    If wb Is Nothing Then Exit Sub

    ' a copy of PCB_thermal_TEMPLATE.xlsx is already set up: only fill empty chip lists
    If Not PT_Rng(wb, "Tpl_Version") Is Nothing Then
        On Error GoTo fail
        chipsFilled = PT_FillChipLists(wb)
        If chipsFilled > 0 Then wb.Save
        On Error GoTo 0
        MsgBox "This workbook comes from PCB_thermal_TEMPLATE (version " & PT_NameStr(wb, "Tpl_Version", "?") & ")." & vbCrLf & _
               "It is already set up. Nothing else to do." & vbCrLf & vbCrLf & _
               IIf(chipsFilled > 0, chipsFilled & " chip rows written to the Spreader and Sim setup tabs of " & wb.Name & ".", _
                   "Chip rows: not changed (already filled, or no RefDes with Include = Yes on CAD geometry yet)."), _
               vbInformation, "SetupWorkbook"
        Exit Sub
    End If

    On Error GoTo fail
    xl.ScreenUpdating = False
    Set reg = PT_Sheet(wb, "Component register")
    Set stIn = reg.Range("B8")
    Set stCalc = reg.Range("S8")
    Set stHead = reg.Range("B7")

    MakeSettings wb
    MakeLog wb
    ExtendImport wb
    MakePrevImport wb          ' before the STEP import names exist, so the copy gets none
    NameExisting wb
    MakeSpreader wb
    MakeSimSetup wb
    MakeSimResults wb
    ExtendGeometry wb
    xl.Calculate
    chipsFilled = PT_FillChipLists(wb)
    PT_Sheet(wb, "Settings").Activate

    xl.ScreenUpdating = True
    xl.CalculateFull
    wb.Save

    Dim unchecked As Double
    unchecked = PT_NameNum(wb, "Log_Unchecked", -1)
    MsgBox "Written and saved to:" & vbCrLf & wb.FullName & vbCrLf & vbCrLf & _
           "Tabs added: " & IIf(addedTabs = "", "none (all were there, labels and formulas refreshed)", Mid(addedTabs, 3)) & vbCrLf & _
           "Workbook names defined: " & wb.Names.Count & vbCrLf & _
           "Chip rows filled from CAD geometry: " & chipsFilled & vbCrLf & _
           "Unchecked Change log rows: " & unchecked & vbCrLf & vbCrLf & _
           "Next: check the Settings tab (templates, prefix) and the Spreader tab (Covered column).", _
           vbInformation, "SetupWorkbook"
    Exit Sub
fail:
    xl.ScreenUpdating = True
    MsgBox "SetupWorkbook stopped: " & Err.Description & vbCrLf & _
           "The workbook was not saved. Close it without saving, then run again." & vbCrLf & _
           "If it happens again, send this message.", vbExclamation, "SetupWorkbook"
End Sub

' ================================================================= tabs
Sub MakeSettings(ByVal wb As Object)
    Dim ws As Object, dfltPrefix As String
    Set ws = GetOrAddSheet(wb, "Settings")
    SetTitle ws, "B2", "Settings"
    ws.Range("B3").Value = "Blue cells are inputs. Every macro reads this tab. A blank template means the SolidWorks default."
    ws.Range("B4:D4").Value = Array("Setting", "Value", "Note")
    Head ws.Range("B4:D4")
    dfltPrefix = PrefixFromModel(wb)
    SetRow wb, ws, 5, "Part template (.prtdot)", "", "Set_PartTemplate", "Full path. Blank = SolidWorks default part template."
    SetRow wb, ws, 6, "Assembly template (.asmdot)", "", "Set_AsmTemplate", "Full path. Blank = SolidWorks default assembly template."
    SetRow wb, ws, 7, "Output folder", "", "Set_OutputFolder", "Blank = folder CAD next to this workbook."
    SetRow wb, ws, 8, "File name prefix", dfltPrefix, "Set_Prefix", "Put in front of every file the macros make."
    SetRow wb, ws, 9, "Position tolerance (mm)", 0.05, "Set_TolPos", "Smaller moves are not reported."
    SetRow wb, ws, 10, "Size tolerance (mm)", 0.05, "Set_TolSize", "Smaller size changes are not reported."
    SetRow wb, ws, 11, "Thermal assembly file name", "", "Set_AsmName", "Blank = <prefix>Thermal.sldasm"

    ' lists for dropdowns
    ws.Range("F3").Value = "Lists for dropdowns (do not edit)"
    ws.Range("F4:K4").Value = Array("Yes/No", "Height mode", "TIM mode", "Underside", "Chip model", "Study type")
    Head ws.Range("F4:K4")
    ws.Range("F5:K5").Value = Array("Yes", "Pedestal", "Body", "Auto", "Block", "Thermal")
    ws.Range("F6:K6").Value = Array("No", "Thick TIM", "Contact R", "Manual", "Two-resistor", "Flow")
    AddName wb, "Lst_YesNo", ws, "F5:F6"
    AddName wb, "Lst_HeightMode", ws, "G5:G6"
    AddName wb, "Lst_TIMMode", ws, "H5:H6"
    AddName wb, "Lst_UnderMode", ws, "I5:I6"
    AddName wb, "Lst_ChipModel", ws, "J5:J6"
    AddName wb, "Lst_SimType", ws, "K5:K6"
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 45
    ws.Columns("D").ColumnWidth = 50
    ws.Columns("F:K").ColumnWidth = 13
End Sub

Sub MakeLog(ByVal wb As Object)
    Dim ws As Object, isNew As Boolean
    isNew = Not PT_HasSheet(wb, PT_LOG_SHEET)
    Set ws = GetOrAddSheet(wb, PT_LOG_SHEET)
    SetTitle ws, "B2", "Change log"
    ws.Range("B3").Value = "Newest on top. Macros add a row for every change they find or make. " & _
                           "Look at each row and type Yes under Checked. Build macros do not run while a row is unchecked."
    ws.Range("B4").Value = "Rows not checked"
    ' starts at the header row so rows inserted at the top stay inside the range; minus 1 for the header
    Calc ws, "C4", "=COUNTIFS(G7:G5000,""?*"",L7:L5000,""<>Yes"")-1"
    ws.Range("C4").Font.Bold = True
    AddName wb, "Log_Unchecked", ws, "C4"
    ws.Range("B7:L7").Value = Array("Date", "Macro", "Key", "Name", "RefDes", "Change", "Field", "Old", "New", "Delta", "Checked")
    Head ws.Range("B7:L7")
    AddName wb, "Log_Header", ws, "B7:L7"
    If isNew Then
        ListDV ws.Range("L8:L1000"), "Lst_YesNo"
        ws.Range("L8:L1000").Interior.Color = stIn.Interior.Color
    End If
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 40
    ws.Columns("F").ColumnWidth = 8
    ws.Columns("G").ColumnWidth = 22
    ws.Columns("H:K").ColumnWidth = 12
    ws.Columns("L").ColumnWidth = 9
End Sub

Sub ExtendImport(ByVal wb As Object)
    ' Column K = identity key of each imported body, same rule as ImportStep uses.
    Dim ws As Object, r As Long, s As String, used As Object
    Set ws = PT_Sheet(wb, "STEP import")
    Set used = CreateObject("Scripting.Dictionary")
    ws.Range("K7").Value = "Key"
    Head ws.Range("K7")
    For r = IMP_FIRST To IMP_LAST
        s = PT_Str(ws.Range("C" & r).Value)
        If s = "" Then
            ws.Range("K" & r).ClearContents
        ElseIf r = IMP_FIRST Then
            ws.Range("K" & r).Value = PT_UniqueKey(used, "pcb")
        Else
            ws.Range("K" & r).Value = PT_UniqueKey(used, PT_Key(s))
        End If
    Next
    ws.Columns("K").ColumnWidth = 24
End Sub

Sub MakePrevImport(ByVal wb As Object)
    Dim ws As Object, src As Object
    Set src = PT_Sheet(wb, "STEP import")
    If Not PT_HasSheet(wb, "STEP import (previous)") Then
        src.Copy , src
        Set ws = wb.Worksheets(src.Index + 1)
        ws.Name = "STEP import (previous)"
        addedTabs = addedTabs & ", STEP import (previous) (hidden)"
    End If
    Set ws = PT_Sheet(wb, "STEP import (previous)")
    ws.Range("B2").Value = "STEP import (previous)"
    ws.Range("B3").Value = "Copy of the STEP import tab before the last import. ImportStep compares the two. Do not edit."
    src.Activate
    ws.Visible = 0
End Sub

Sub NameExisting(ByVal wb As Object)
    Dim ws As Object
    Set ws = PT_Sheet(wb, "Component register")
    AddName wb, "Reg_Table", ws, "B8:AF53"
    AddName wb, "Reg_RefDes", ws, "B8:B53"
    AddName wb, "Reg_QUsed", ws, "V8:V53"
    AddName wb, "Reg_ThetaJC", ws, "R8:R53"
    AddName wb, "Reg_TotalQ", ws, "D4"
    AddName wb, "Reg_Ambient", ws, "D5"

    Set ws = PT_Sheet(wb, "CAD geometry")
    AddName wb, "Geo_Table", ws, "B" & GEO_FIRST & ":AH" & GEO_LAST
    AddName wb, "Geo_PCB_Imported", ws, "C7:C9"
    AddName wb, "Geo_PCB_Override", ws, "D7:D9"
    AddName wb, "Geo_PCB_Used", ws, "E7:E9"
    AddName wb, "Geo_kXY", ws, "D19"
    AddName wb, "Geo_kZ", ws, "D20"
    AddName wb, "Geo_Checks", ws, "M6:M12"

    Set ws = PT_Sheet(wb, "STEP import")
    AddName wb, "Imp_Settings", ws, "M3:M5"
    AddName wb, "Imp_PcbName", ws, "M3"
    AddName wb, "Imp_MinSize", ws, "M4"
    AddName wb, "Imp_TopOnly", ws, "M5"
    AddName wb, "Imp_Model", ws, "D4"
    AddName wb, "Imp_Time", ws, "D5"
    AddName wb, "Imp_Frame", ws, "D6"
    AddName wb, "Imp_Table", ws, "B" & IMP_FIRST & ":K" & IMP_LAST
    AddName wb, "Imp_Messages", ws, "L8:L200"

    Set ws = PT_Sheet(wb, "STEP import (previous)")
    AddName wb, "ImpPrev_Model", ws, "D4"
    AddName wb, "ImpPrev_Time", ws, "D5"
    AddName wb, "ImpPrev_Table", ws, "B" & IMP_FIRST & ":K" & IMP_LAST

    Set ws = PT_Sheet(wb, "Thermal solution")
    AddName wb, "TS_SimTj", ws, "AF16:AF68"
End Sub

Sub MakeSpreader(ByVal wb As Object)
    Dim ws As Object, g As String, r As Long, sq As String
    Set ws = GetOrAddSheet(wb, "Spreader")
    g = "'CAD geometry'!"
    sq = ChrW(178)
    SetTitle ws, "B2", "Spreader"
    ws.Range("B3").Value = "Blue cells are inputs. BuildSpreader builds from this tab. The AS BUILT block is written by SyncFromCAD. " & _
                           "Heights are from the PCB bottom face, mm."

    ' A. plate and clearances
    ws.Range("B4:D4").Value = Array("A. PLATE AND CLEARANCES", "Value", "Unit / note")
    Head ws.Range("B4:D4")
    SetRow wb, ws, 5, "Plate material", "Aluminium 6061", "Spr_Material", "Name used for the SolidWorks material"
    SetRow wb, ws, 6, "Plate k", 167, "Spr_k", "W/m K"
    SetRow wb, ws, 7, "Plate density", 2700, "Spr_Density", "kg/m3"
    SetRow wb, ws, 8, "Plate thickness", 1.5, "Spr_Thickness", "mm"
    SetRow wb, ws, 9, "Clearance around tall parts, XY", 1, "Spr_ClearXY", "mm"
    SetRow wb, ws, 10, "Clearance to tall parts, Z", 1, "Spr_ClearZ", "mm. Parts whose top is above (underside - this) get a cut."
    SetRow wb, ws, 11, "Smallest pedestal", 0.3, "Spr_MinPedestal", "mm. Below this the TIM fills the whole gap."
    SetRow wb, ws, 12, "Underside height mode", "Auto", "Spr_UnderMode", "Auto = tallest covered chip + its TIM. Manual = next row."
    ListDV ws.Range("C12"), "Lst_UnderMode"
    SetRow wb, ws, 13, "Underside height if Manual", 5, "Spr_UnderManual", "mm"
    ws.Range("B14").Value = "Underside height used"
    Calc ws, "C14", "=IF(C12=""Manual"",N(C13),MAX(S" & CHIP_FIRST & ":S" & CHIP_LAST & "))"
    ws.Range("D14").Value = "mm"
    AddName wb, "Spr_ZUnder", ws, "C14"
    ws.Range("B15").Value = "Covered chips"
    Calc ws, "C15", "=COUNTIF(I" & CHIP_FIRST & ":I" & CHIP_LAST & ",""Yes"")"
    AddName wb, "Spr_CoveredCount", ws, "C15"
    ws.Range("B16").Value = "Clearance cuts"
    Calc ws, "C16", "=COUNT(J" & CUT_FIRST & ":J" & CUT_LAST & ")"
    AddName wb, "Spr_CutCount", ws, "C16"

    ' D. as built (SyncFromCAD writes the values)
    ws.Range("F4:H4").Value = Array("D. AS BUILT (SyncFromCAD writes)", "Value", "Unit")
    Head ws.Range("F4:H4")
    AsBuiltRow wb, ws, 5, "Outline area", "As_Area", "mm" & sq
    AsBuiltRow wb, ws, 6, "Top face area", "As_TopArea", "mm" & sq
    AsBuiltRow wb, ws, 7, "Exposed area (all faces not touching a part)", "As_ExposedArea", "mm" & sq
    AsBuiltRow wb, ws, 8, "Volume", "As_Volume", "mm3"
    AsBuiltRow wb, ws, 9, "Mass", "As_Mass", "g"
    AsBuiltRow wb, ws, 10, "Plate thickness", "As_Thickness", "mm"
    AsBuiltRow wb, ws, 11, "Holes (circular, through)", "As_Holes", "count"
    AsBuiltRow wb, ws, 12, "Last sync", "As_LastSync", ""

    ' B. chips under the spreader
    ws.Range("B18").Value = "B. CHIPS UNDER THE SPREADER  (type a RefDes from CAD geometry in column B, one row per chip)"
    ws.Range("B18").Font.Bold = True
    ws.Range("B19:W19").Value = Array("RefDes", "Key", "X min (mm)", "Y min (mm)", "L (mm)", "W (mm)", "Top Z (mm)", _
        "Covered", "Height mode", "TIM mode", "TIM t nominal (mm)", "TIM k (W/m K)", "Gap (mm)", "TIM t used (mm)", _
        "Pedestal h (mm)", "TIM R (K/W)", "Check", "Z need (helper)", "CAD geometry row (helper)", _
        "Pedestal as built (mm)", "TIM t as built (mm)", "As built vs design")
    Head ws.Range("B19:W19")
    r = CHIP_FIRST
    InpRange ws, "B" & r & ":B" & CHIP_LAST
    InpRange ws, "I" & r & ":M" & CHIP_LAST
    ListDV ws.Range("I" & r & ":I" & CHIP_LAST), "Lst_YesNo"
    ListDV ws.Range("J" & r & ":J" & CHIP_LAST), "Lst_HeightMode"
    ListDV ws.Range("K" & r & ":K" & CHIP_LAST), "Lst_TIMMode"
    CalcCol ws, "T", r, CHIP_LAST, "=IF($B20="""","""",IFERROR(MATCH($B20," & g & "$F$26:$F$95,0),""""))"
    CalcCol ws, "C", r, CHIP_LAST, "=IF($T20="""","""",INDEX(" & g & "$AG$26:$AG$95,$T20))"
    CalcCol ws, "D", r, CHIP_LAST, "=IF($T20="""","""",INDEX(" & g & "$U$26:$U$95,$T20))"
    CalcCol ws, "E", r, CHIP_LAST, "=IF($T20="""","""",INDEX(" & g & "$V$26:$V$95,$T20))"
    CalcCol ws, "F", r, CHIP_LAST, "=IF($T20="""","""",INDEX(" & g & "$X$26:$X$95,$T20))"
    CalcCol ws, "G", r, CHIP_LAST, "=IF($T20="""","""",INDEX(" & g & "$Y$26:$Y$95,$T20))"
    CalcCol ws, "H", r, CHIP_LAST, "=IF($T20="""","""",IFERROR(INDEX(" & g & "$W$26:$W$95,$T20)+INDEX(" & g & "$Z$26:$Z$95,$T20),""""))"
    CalcCol ws, "N", r, CHIP_LAST, "=IF(OR($I20<>""Yes"",NOT(ISNUMBER($H20))),"""",Spr_ZUnder-$H20)"
    CalcCol ws, "O", r, CHIP_LAST, "=IF($N20="""","""",IF($J20=""Thick TIM"",$N20,IF($N20-N($L20)<Spr_MinPedestal,$N20,N($L20))))"
    CalcCol ws, "P", r, CHIP_LAST, "=IF($N20="""","""",$N20-$O20)"
    CalcCol ws, "Q", r, CHIP_LAST, "=IF(OR($O20="""",N($M20)<=0,N($F20)*N($G20)<=0),"""",$O20*1000/($M20*$F20*$G20))"
    CalcCol ws, "R", r, CHIP_LAST, "=IF($B20="""","""",IF($T20="""",""RefDes not on CAD geometry""," & _
        "IF(INDEX(" & g & "$G$26:$G$95,$T20)<>""Yes"",""body not included""," & _
        "IF($I20<>""Yes"",""not covered""," & _
        "IF(NOT(ISNUMBER($N20)),""missing geometry""," & _
        "IF($N20<-0.0005,""chip is above the underside""," & _
        "IF(AND($K20=""Body"",$O20<0.01),""TIM too thin for a body: use Contact R""," & _
        "IF(AND($K20=""Contact R"",$O20>0.3),""big gap for a contact resistance: use Body"",""OK""))))))))"
    CalcCol ws, "S", r, CHIP_LAST, "=IF(AND($I20=""Yes"",ISNUMBER($H20)),$H20+N($L20),0)"
    CalcCol ws, "W", r, CHIP_LAST, "=IF(OR(NOT(ISNUMBER($U20)),NOT(ISNUMBER($P20))),""""," & _
        "IF(OR(ABS($U20-$P20)>Set_TolSize,ABS(N($V20)-$O20)>Set_TolSize),""differs"",""""))"
    ApplyLook ws.Range("U" & r & ":V" & CHIP_LAST), stCalc
    ws.Range("D" & r & ":H" & CHIP_LAST).NumberFormat = "0.00"
    ws.Range("N" & r & ":P" & CHIP_LAST).NumberFormat = "0.000"
    ws.Range("Q" & r & ":Q" & CHIP_LAST).NumberFormat = "0.00"
    AddName wb, "Spr_ChipHeader", ws, "B19:W19"
    AddName wb, "Spr_Chips", ws, "B" & CHIP_FIRST & ":W" & CHIP_LAST

    ' C. clearance cuts (all formulas)
    ws.Range("B83").Value = "C. CLEARANCE CUTS  (automatic: included parts that are not covered and reach above the underside minus the Z clearance)"
    ws.Range("B83").Font.Bold = True
    ws.Range("B84:J84").Value = Array("#", "Key", "Name", "X min (mm)", "Y min (mm)", "L (mm)", "W (mm)", "Top Z (mm)", "CAD geometry row (helper)")
    Head ws.Range("B84:J84")
    For r = CUT_FIRST To CUT_LAST
        ws.Range("B" & r).Value = r - CUT_FIRST + 1
    Next
    r = CUT_FIRST
    CalcCol ws, "J", r, CUT_LAST, "=IFERROR(MATCH($B85," & g & "$AH$26:$AH$95,0),"""")"
    CalcCol ws, "C", r, CUT_LAST, "=IF($J85="""","""",INDEX(" & g & "$AG$26:$AG$95,$J85))"
    CalcCol ws, "D", r, CUT_LAST, "=IF($J85="""","""",INDEX(" & g & "$E$26:$E$95,$J85))"
    CalcCol ws, "E", r, CUT_LAST, "=IF($J85="""","""",INDEX(" & g & "$U$26:$U$95,$J85)-Spr_ClearXY)"
    CalcCol ws, "F", r, CUT_LAST, "=IF($J85="""","""",INDEX(" & g & "$V$26:$V$95,$J85)-Spr_ClearXY)"
    CalcCol ws, "G", r, CUT_LAST, "=IF($J85="""","""",INDEX(" & g & "$X$26:$X$95,$J85)+2*Spr_ClearXY)"
    CalcCol ws, "H", r, CUT_LAST, "=IF($J85="""","""",INDEX(" & g & "$Y$26:$Y$95,$J85)+2*Spr_ClearXY)"
    CalcCol ws, "I", r, CUT_LAST, "=IF($J85="""","""",INDEX(" & g & "$W$26:$W$95,$J85)+INDEX(" & g & "$Z$26:$Z$95,$J85))"
    ws.Range("E" & r & ":I" & CUT_LAST).NumberFormat = "0.00"
    AddName wb, "Spr_Cuts", ws, "B" & CUT_FIRST & ":J" & CUT_LAST

    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 26
    ws.Columns("D:W").ColumnWidth = 11
    ws.Columns("F").ColumnWidth = 40
    ws.Columns("R").ColumnWidth = 36
End Sub

Sub MakeSimSetup(ByVal wb As Object)
    Dim ws As Object, g As String, rgs As String, r As Long, th As String, dg As String
    Set ws = GetOrAddSheet(wb, "Sim setup")
    g = "'CAD geometry'!"
    rgs = "'Component register'!"
    th = ChrW(952): dg = ChrW(176)
    SetTitle ws, "B2", "Sim setup"
    ws.Range("B3").Value = "Blue cells are inputs. SetupSim reads this tab."

    ws.Range("B4:D4").Value = Array("GLOBAL", "Value", "Unit / note")
    Head ws.Range("B4:D4")
    SetRow wb, ws, 5, "Study type", "Thermal", "Sim_Type", "Thermal = SolidWorks Simulation. Flow = Flow Simulation (later)."
    ListDV ws.Range("C5"), "Lst_SimType"
    ws.Range("B6").Value = "Run name (SetupSim writes it)"
    ApplyLook ws.Range("C6"), stCalc
    AddName wb, "Sim_RunName", ws, "C6"
    ws.Range("B7").Value = "Ambient"
    Calc ws, "C7", "=" & rgs & "$D$5"
    ws.Range("D7").Value = dg & "C, from the Component register"
    AddName wb, "Sim_Ambient", ws, "C7"
    SetRow wb, ws, 8, "Global mesh size", 3, "Sim_MeshGlobal", "mm"
    SetRow wb, ws, 9, "Fine mesh size (chips and TIMs)", 0.5, "Sim_MeshFine", "mm"
    SetRow wb, ws, 10, "Include parts without heat", "Yes", "Sim_IncludeNonHeat", "No = those parts are left out of the study"
    ListDV ws.Range("C10"), "Lst_YesNo"
    SetRow wb, ws, 11, "k of parts without heat", 0.5, "Sim_kParts", "W/m K"
    SetRow wb, ws, 12, "Two-resistor: outer layer thickness", 0.1, "Sim_LayerFrac", "fraction of the chip height, for each of the top and bottom layers"
    SetRow wb, ws, 13, "Two-resistor: k of the junction core", 100, "Sim_kCore", "W/m K, high so the core is one temperature"

    ws.Range("B14:D14").Value = Array("SURFACE GROUP", "h (W/m2 K)", "Emissivity")
    Head ws.Range("B14:D14")
    ws.Range("B15:B19").Value = xlCol(Array("Spreader top", "Spreader other faces", "PCB top", "PCB bottom", "Parts"))
    InpRange ws, "C15:D19"
    FillIfEmpty ws.Range("C15:D15"), Array(10, 0.85)
    FillIfEmpty ws.Range("C16:D16"), Array(5, 0.85)
    FillIfEmpty ws.Range("C17:D17"), Array(8, 0.9)
    FillIfEmpty ws.Range("C18:D18"), Array(5, 0.9)
    FillIfEmpty ws.Range("C19:D19"), Array(8, 0.9)
    AddName wb, "Sim_Surfaces", ws, "B15:D19"

    ws.Range("B22:D22").Value = Array("INTERFACE", "Contact R (K m2/W)", "Note")
    Head ws.Range("B22:D22")
    ws.Range("B23:B25").Value = xlCol(Array("Part to PCB", "Chip to TIM", "TIM to spreader"))
    InpRange ws, "C23:C25"
    FillIfEmpty ws.Range("C23"), Array(0)
    FillIfEmpty ws.Range("C24"), Array(0)
    FillIfEmpty ws.Range("C25"), Array(0)
    ws.Range("D23").Value = "0 = bonded"
    AddName wb, "Sim_Contacts", ws, "B23:D25"

    ws.Range("B27").Value = "PER CHIP  (type a RefDes from CAD geometry in column B)"
    ws.Range("B27").Font.Bold = True
    ws.Range("B28:L28").Value = Array("RefDes", "Chip model", th & "jc (" & dg & "C/W)", th & "jb (" & dg & "C/W)", _
        "H (mm)", "A (mm" & ChrW(178) & ")", "k z calibrated (W/m K)", "k xy (W/m K, blank = k z)", "Q (W)", "Check", _
        "CAD geometry row (helper)")
    Head ws.Range("B28:L28")
    r = SCHIP_FIRST
    InpRange ws, "B" & r & ":C" & SCHIP_LAST
    InpRange ws, "E" & r & ":E" & SCHIP_LAST
    InpRange ws, "I" & r & ":I" & SCHIP_LAST
    ListDV ws.Range("C" & r & ":C" & SCHIP_LAST), "Lst_ChipModel"
    CalcCol ws, "L", r, SCHIP_LAST, "=IF($B29="""","""",IFERROR(MATCH($B29," & g & "$F$26:$F$95,0),""""))"
    CalcCol ws, "D", r, SCHIP_LAST, "=IF($B29="""","""",IFERROR(IF(N(INDEX(" & rgs & "$R$8:$R$53,MATCH($B29," & rgs & "$B$8:$B$53,0)))>0," & _
        "INDEX(" & rgs & "$R$8:$R$53,MATCH($B29," & rgs & "$B$8:$B$53,0)),""""),""""))"
    CalcCol ws, "F", r, SCHIP_LAST, "=IF($L29="""","""",INDEX(" & g & "$Z$26:$Z$95,$L29))"
    CalcCol ws, "G", r, SCHIP_LAST, "=IF($L29="""","""",IFERROR(INDEX(" & g & "$X$26:$X$95,$L29)*INDEX(" & g & "$Y$26:$Y$95,$L29),""""))"
    CalcCol ws, "H", r, SCHIP_LAST, "=IF(AND(N($D29)>0,N($F29)>0,N($G29)>0),$F29*1000/($D29*$G29),"""")"
    CalcCol ws, "J", r, SCHIP_LAST, "=IF($B29="""","""",SUMIFS(" & rgs & "$V$8:$V$53," & rgs & "$B$8:$B$53,$B29))"
    CalcCol ws, "K", r, SCHIP_LAST, "=IF($B29="""","""",IF($L29="""",""RefDes not on CAD geometry""," & _
        "IF(N($J29)<=0,""no heat in the register""," & _
        "IF($C29=""Two-resistor"",IF(AND(N($D29)>0,N($E29)>0),""OK (two-resistor)"",""needs " & th & "jc and " & th & "jb: block model is used"")," & _
        "IF($H29="""",""no " & th & "jc: k of parts without heat is used"",""OK"")))))"
    ws.Range("F" & r & ":H" & SCHIP_LAST).NumberFormat = "0.00"
    AddName wb, "Sim_ChipHeader", ws, "B28:L28"
    AddName wb, "Sim_Chips", ws, "B" & SCHIP_FIRST & ":L" & SCHIP_LAST

    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 44
    ws.Columns("E:J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 40
End Sub

Sub MakeSimResults(ByVal wb As Object)
    Dim ws As Object, dg As String
    Set ws = GetOrAddSheet(wb, "Sim results")
    dg = ChrW(176)
    SetTitle ws, "B2", "Sim results"
    ws.Range("B3").Value = "One row per run, newest at the bottom. ReadResults writes this tab and adds per-chip columns as needed."
    ws.Range("B7:P7").Value = Array("Run", "Date", "Study", "Spreader file", "Mesh global (mm)", "Mesh fine (mm)", _
        "h spreader top", "h PCB top", "TIM modes", "Spreader max (" & dg & "C)", "Spreader min (" & dg & "C)", _
        "PCB max (" & dg & "C)", "PCB min (" & dg & "C)", "Energy balance (%)", "Plot file")
    Head ws.Range("B7:P7")
    AddName wb, "Res_Header", ws, "B7:P7"
    ws.Columns("B:D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 30
    ws.Columns("F:P").ColumnWidth = 12
End Sub

Sub ExtendGeometry(ByVal wb As Object)
    Dim ws As Object, lg As String, hit As String, r1 As Long
    Set ws = PT_Sheet(wb, "CAD geometry")
    lg = "'Change log'!"
    ws.Range("AF25:AH25").Value = Array("Last change (not checked)", "Key", "Cut #")
    Head ws.Range("AF25:AH25")

    ' AG key: imported rows take it from STEP import column K; manual rows make one from the RefDes
    r1 = GEO_FIRST
    CalcCol ws, "AG", r1, GEO_IMP_LAST, "=IF('STEP import'!$K9="""",IF($E26="""","""",IF($F26<>"""",""m_""&LOWER($F26),""m_row""&ROW())),'STEP import'!$K9)"
    CalcCol ws, "AG", GEO_IMP_LAST + 1, GEO_LAST, "=IF($E" & (GEO_IMP_LAST + 1) & "="""","""",IF($F" & (GEO_IMP_LAST + 1) & "<>"""",""m_""&LOWER($F" & _
        (GEO_IMP_LAST + 1) & "),""m_row""&ROW()))"

    ' AF newest unchecked Change log row for this key
    hit = "MATCH(1,INDEX((" & lg & "$D$7:$D$5000=$AG26)*(" & lg & "$L$7:$L$5000<>""Yes""),0),0)"
    CalcCol ws, "AF", r1, GEO_LAST, "=IF($AG26="""","""",IFERROR(INDEX(" & lg & "$G$7:$G$5000," & hit & ")&"" ""&INDEX(" & lg & "$H$7:$H$5000," & hit & "),""""))"
    With ws.Range("AF" & r1 & ":AF" & GEO_LAST)
        .FormatConditions.Delete
        .FormatConditions.Add 1, 4, "="""""                         ' cell value not equal to ""
        .FormatConditions(1).Interior.Color = RGB(255, 199, 206)
        .FormatConditions(1).Font.Color = RGB(156, 0, 6)
    End With

    ' AH running number of the parts that need a clearance cut in the spreader
    CalcCol ws, "AH", r1, GEO_LAST, "=IF(AND($G26=""Yes"",ISNUMBER($W26),ISNUMBER($Z26)),IF(AND($W26+$Z26>Spr_ZUnder-Spr_ClearZ," & _
        "COUNTIFS(Spreader!$C$20:$C$79,$AG26,Spreader!$I$20:$I$79,""Yes"")=0),MAX(AH$25:AH25)+1,""""),"""")"
    ws.Columns("AF").ColumnWidth = 24
    ws.Columns("AG:AH").Hidden = True
End Sub

' First run only: list every included body with a RefDes on the Spreader and Sim setup tabs,
' with defaults picked from the geometry (works for any board).
' FillChipLists moved to PTCommon (PT_FillChipLists) so ImportStep and BuildBoard can use it.

' ================================================================= helpers
Function GetOrAddSheet(ByVal wb As Object, ByVal sheetName As String) As Object
    Dim i As Long, lastVis As Object
    If PT_HasSheet(wb, sheetName) Then
        Set GetOrAddSheet = PT_Sheet(wb, sheetName)
        Exit Function
    End If
    For i = wb.Worksheets.Count To 1 Step -1
        If wb.Worksheets(i).Visible = -1 Then Set lastVis = wb.Worksheets(i): Exit For
    Next
    Set GetOrAddSheet = wb.Worksheets.Add(, lastVis)
    GetOrAddSheet.Name = sheetName
    addedTabs = addedTabs & ", " & sheetName
End Function

Function PrefixFromModel(ByVal wb As Object) As String
    ' default file prefix from the imported model name, e.g. "Raspberry Pi 4 Model B" -> "RaspberryPi4_"
    Dim s As String, i As Long, ch As String, outS As String
    s = PT_Str(PT_Sheet(wb, "STEP import").Range("D4").Value)
    s = Mid(s, InStrRev(Replace(s, "/", "\"), "\") + 1)
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch Like "[A-Za-z0-9]" Then outS = outS & ch
        If Len(outS) >= 12 Then Exit For
    Next
    If outS = "" Then outS = "PCB"
    PrefixFromModel = outS & "_"
End Function

Sub SetTitle(ByVal ws As Object, ByVal addr As String, ByVal txt As String)
    ws.Range(addr).Value = txt
    ws.Range(addr).Font.Bold = True
    ws.Range(addr).Font.Size = 14
End Sub

Sub Head(ByVal rng As Object)
    ApplyLook rng, stHead
    rng.Font.Bold = True
    rng.WrapText = True
End Sub

Sub ApplyLook(ByVal rng As Object, ByVal src As Object)
    rng.Interior.Color = src.Interior.Color
    rng.Font.Color = src.Font.Color
End Sub

' one labelled input row: B label, C value (blue, default only if empty), D note, named C
Sub SetRow(ByVal wb As Object, ByVal ws As Object, ByVal r As Long, ByVal lab As String, ByVal dflt As Variant, _
           ByVal nameS As String, ByVal note As String)
    ws.Range("B" & r).Value = lab
    InpRange ws, "C" & r
    If CStr(dflt) <> "" Then
        If IsEmpty(ws.Range("C" & r).Value) Then ws.Range("C" & r).Value = dflt
    End If
    ws.Range("D" & r).Value = note
    AddName wb, nameS, ws, "C" & r
End Sub

Sub AsBuiltRow(ByVal wb As Object, ByVal ws As Object, ByVal r As Long, ByVal lab As String, ByVal nameS As String, ByVal unitS As String)
    ws.Range("F" & r).Value = lab
    ApplyLook ws.Range("G" & r), stCalc
    ws.Range("H" & r).Value = unitS
    AddName wb, nameS, ws, "G" & r
End Sub

Sub InpRange(ByVal ws As Object, ByVal addr As String)
    ApplyLook ws.Range(addr), stIn
End Sub

Sub FillIfEmpty(ByVal rng As Object, ByVal vals As Variant)
    Dim i As Long
    For i = 1 To rng.Cells.Count
        If IsEmpty(rng.Cells(i).Value) Then rng.Cells(i).Value = vals(i - 1)
    Next
End Sub

Sub Calc(ByVal ws As Object, ByVal addr As String, ByVal fml As String)
    ws.Range(addr).Formula = fml
    ApplyLook ws.Range(addr), stCalc
End Sub

' formula written for the first row, Excel shifts the relative references for the rest
Sub CalcCol(ByVal ws As Object, ByVal colS As String, ByVal r1 As Long, ByVal r2 As Long, ByVal fml As String)
    Calc ws, colS & r1 & ":" & colS & r2, fml
End Sub

Sub AddName(ByVal wb As Object, ByVal nameS As String, ByVal ws As Object, ByVal addr As String)
    wb.Names.Add nameS, "='" & ws.Name & "'!" & ws.Range(addr).Address
End Sub

Sub ListDV(ByVal rng As Object, ByVal listName As String)
    rng.Validation.Delete
    rng.Validation.Add 3, 1, 1, "=" & listName
End Sub

' 1-D array -> column array for writing down a range
Function xlCol(ByVal a As Variant) As Variant
    Dim outv() As Variant, i As Long
    ReDim outv(1 To UBound(a) + 1, 1 To 1)
    For i = 0 To UBound(a): outv(i + 1, 1) = a(i): Next
    xlCol = outv
End Function
