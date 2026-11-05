Attribute VB_Name = "SendSolution"
Option Explicit
' =====================================================================================
' SendSolution  -  SolidWorks macro, the Send Solution button (Ansys page route, ansys\page.html)
'
' After Build Solution and your own edits (mounting holes, shape, cuts), sends the plates back to the board folder,
' for the page's hand check and for Apply, which imports them into AEDT. Changes no SolidWorks file.
' Pick study.json of the board folder; the assembly is <folder>\solidworks\<board>_solution.sldasm. Writes:
'   solution_top.step, solution_bottom.step   each plate part as it is now (part frame; offsetMm moves it into the
'                                             board STEP's frame)
'   solution_cad.json  sentAt, frame (copied from study.json), plates[] and tims[], problems[]:
'     plate: side, part, file, offsetMm [x, y, z], volumeMm3, areaMm2 (all faces), boxMm [x0, y0, z0, x1, y1, z1],
'            undersideZMm, topZMm, tMm, topAreaMm2, undersideAreaMm2, contacts {RefDes: area of Contact_<RefDes>}
'     TIM:   refDes, side, boxMm [x0, y0, z0, x1, y1, z1]
' All in mm, the board STEP's frame. A suppressed plate or TIM is left out (AEDT then suppresses it on Apply).
' Needs PTCommon and PTJson in the same macro.
' =====================================================================================

Const MACRO_NAME As String = "SendSolution"
Const swSaveAsCopy As Long = 2

Dim probs As Collection
Dim topZ As Double, haveTopZ As Boolean      ' board top face in the STEP's frame: tells a top TIM from a bottom one

Sub main()
    Set swApp = Application.SldWorks
    PT_Trace MACRO_NAME & ": started (VBA macro)."
    Set probs = New Collection
    If Not PJ_PickStudy(MACRO_NAME) Then Exit Sub
    Dim asm As Object, asmPath As String, sc As Object, frame As Object
    asmPath = PJ_SolutionAsm()
    Set asm = PT_OpenAsm(asmPath)
    If asm Is Nothing Then
        PT_Msg MACRO_NAME & " did not run. The solution assembly is not there:" & vbCrLf & asmPath & vbCrLf & vbCrLf & _
               "Press Build Solution first.", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    Set sc = PJ_ReadOwn("solution_cad.json", "pcb-thermal-ansys/solution-cad")
    If sc Is Nothing Then
        PT_Msg MACRO_NAME & " did not run. " & PJ_Dir & "\solution_cad.json could not be read:" & vbCrLf & PT_LastErr, vbExclamation, MACRO_NAME
        Exit Sub
    End If
    Set frame = PJ_Obj(PJ_Project, "solution/forSolidWorks/frame")
    haveTopZ = PJ_IsNum(frame, "topZMm")
    topZ = PJ_Num(frame, "topZMm", 0)

    PT_ActivateDoc asm
    asm.ResolveAllLightWeightComponents False
    Dim comps As Variant, i As Long, c As Object, md As Object, role As String, plates As Collection, tims As Collection
    Dim sides As Object, p As Object
    Set plates = New Collection: Set tims = New Collection
    Set sides = CreateObject("Scripting.Dictionary")
    comps = asm.GetComponents(True)
    If Not IsEmpty(comps) Then
        For i = 0 To UBound(comps)
            Set c = comps(i)
            If c.GetSuppression2 <> PT_swSuppressed Then
                Set md = c.GetModelDoc2
                role = PT_GetProp(md, "Role")
                If role = "Spreader" Then
                    Set p = OnePlate(c, md, sides)
                    If Not p Is Nothing Then plates.Add p
                ElseIf role = "TIM" Then
                    Set p = OneTim(c, md)
                    If Not p Is Nothing Then tims.Add p
                End If
            End If
        Next
    End If
    PT_ActivateDoc asm
    If plates.Count = 0 Then probs.Add "No plate in the assembly (Spreader_Top or Spreader_Bottom, not suppressed)."

    PJ_Put sc, "sentAt", PJ_Now()
    PJ_Put sc, "assembly", asmPath
    PJ_Put sc, "frame", frame
    PJ_Put sc, "plates", plates
    PJ_Put sc, "tims", tims
    PJ_Put sc, "problems", probs
    If Not PJ_SaveOwn(sc, "solution_cad.json", MACRO_NAME) Then
        PT_Msg MACRO_NAME & ": " & PJ_Dir & "\solution_cad.json could not be written:" & vbCrLf & PT_LastErr, vbExclamation, MACRO_NAME
        Exit Sub
    End If
    Dim s As String, it As Variant
    s = "Sent " & plates.Count & " plate" & IIf(plates.Count = 1, "", "s") & " and " & tims.Count & " TIM part" & IIf(tims.Count = 1, "", "s") & _
        " to " & PJ_Dir & vbCrLf & "(solution_cad.json, solution_*.step)." & vbCrLf & vbCrLf & _
        "The page shows them in step 4 and uses them in the hand check. Apply puts them into AEDT."
    If probs.Count > 0 Then
        s = s & vbCrLf & vbCrLf & "Problems:"
        For Each it In probs: s = s & vbCrLf & "  " & it: Next
    End If
    PT_Msg s, IIf(probs.Count > 0, vbExclamation, vbInformation), MACRO_NAME
End Sub

' One plate component: its STEP in the board folder and its measures. Nothing when it cannot be sent (in probs).
Function OnePlate(ByVal c As Object, ByVal md As Object, ByVal sides As Object) As Object
    Dim side As String, xf As Variant, k As Long, ox As Double, oy As Double, oz As Double, fileS As String
    Dim pd As Object, errs As Long, warns As Long, ttl As String, wasOpen As Boolean, d As Object
    side = LCase(PT_GetProp(md, "Side"))
    If side <> "bottom" Then side = "top"
    If sides.Exists(side) Then probs.Add c.Name2 & ": a second " & side & " plate, not sent. Suppress the one you do not want.": Exit Function
    sides.Add side, 1
    xf = c.Transform2.ArrayData
    For k = 0 To 8
        If Abs(xf(k) - IIf(k Mod 4 = 0, 1#, 0#)) > 0.000001 Then
            probs.Add c.Name2 & ": the plate is turned in the assembly, not sent. Only move it, or put it back square."
            Exit Function
        End If
    Next
    ox = xf(9) * 1000#: oy = xf(10) * 1000#: oz = xf(11) * 1000#

    ' its own window, for the export and the measures (a part loaded only as a component may have no bodies to read)
    wasOpen = Not swApp.GetOpenDocumentByName(md.GetPathName) Is Nothing And md.Visible
    Set pd = swApp.OpenDoc6(md.GetPathName, PT_swDocPART, PT_swOpenSilent, "", errs, warns)
    If pd Is Nothing Then probs.Add c.Name2 & ": the part could not be opened (error " & errs & ").": Exit Function
    ttl = pd.GetTitle
    fileS = "solution_" & side & ".step"
    If Not pd.Extension.SaveAs(PJ_Dir & "\" & fileS, 0, PT_swSaveSilent + swSaveAsCopy, Nothing, errs, warns) Then
        probs.Add c.Name2 & ": could not write " & fileS & " (error " & errs & "). Is it open in another program?"
        If Not wasOpen Then swApp.CloseDoc ttl
        Exit Function
    End If

    Set d = PJ_Dict()
    PJ_Put d, "side", side
    PJ_Put d, "part", md.GetPathName
    PJ_Put d, "file", fileS
    PJ_Put d, "offsetMm", Nums(Array(ox, oy, oz))
    MeasurePlate pd, d, ox, oy, oz, side
    If Not wasOpen Then swApp.CloseDoc ttl
    Set OnePlate = d
End Function

' Volume, area, box and the named faces (Top, Underside, Contact_<RefDes>) of a plate part, moved by its offset.
Sub MeasurePlate(ByVal pd As Object, ByVal d As Object, ByVal ox As Double, ByVal oy As Double, ByVal oz As Double, ByVal side As String)
    Dim mp As Object, bodies As Variant, i As Long, bb As Variant, lo(2) As Double, hi(2) As Double, k As Long
    Dim faces As Collection, fc As Object, nm As String, contacts As Object, zTop As Double, zUnder As Double, gotT As Boolean, gotU As Boolean
    On Error Resume Next
    Set mp = pd.Extension.CreateMassProperty
    On Error GoTo 0
    If Not mp Is Nothing Then
        PJ_Put d, "volumeMm3", Round(mp.Volume * 1000000000#, 1)
        PJ_Put d, "areaMm2", Round(mp.SurfaceArea * 1000000#, 1)
    Else
        probs.Add pd.GetTitle & ": no mass properties (volume, area)."
    End If
    bodies = pd.GetBodies2(PT_swSolidBody, True)
    If IsEmpty(bodies) Then probs.Add pd.GetTitle & ": no solid body.": Exit Sub
    If UBound(bodies) > 0 Then probs.Add pd.GetTitle & ": " & (UBound(bodies) + 1) & " solid bodies; AEDT gets them all."
    For k = 0 To 2: lo(k) = 1E+30: hi(k) = -1E+30: Next
    For i = 0 To UBound(bodies)
        bb = bodies(i).GetBodyBox
        For k = 0 To 2
            If bb(k) * 1000# < lo(k) Then lo(k) = bb(k) * 1000#
            If bb(k + 3) * 1000# > hi(k) Then hi(k) = bb(k + 3) * 1000#
        Next
    Next
    PJ_Put d, "boxMm", Nums(Array(lo(0) + ox, lo(1) + oy, lo(2) + oz, hi(0) + ox, hi(1) + oy, hi(2) + oz))
    Set contacts = PJ_Dict()
    Set faces = PT_AllFaces(pd)
    For Each fc In faces
        nm = pd.GetEntityName(fc)
        If nm = "Top" Then
            bb = fc.GetBox: zTop = bb(2) * 1000# + oz: gotT = True
            PJ_Put d, "topAreaMm2", Round(fc.GetArea * 1000000#, 2)
        ElseIf nm = "Underside" Then
            bb = fc.GetBox: zUnder = bb(2) * 1000# + oz: gotU = True
            PJ_Put d, "undersideAreaMm2", Round(fc.GetArea * 1000000#, 2)
        ElseIf Left(nm, 8) = "Contact_" Then
            PJ_Put contacts, Mid(nm, 9), Round(fc.GetArea * 1000000#, 3)
        End If
    Next
    If gotU Then PJ_Put d, "undersideZMm", Round(zUnder, 4) Else probs.Add pd.GetTitle & ": no face named Underside. Run Build Solution again to name it."
    If gotT Then PJ_Put d, "topZMm", Round(zTop, 4)
    If gotT And gotU Then PJ_Put d, "tMm", Round(Abs(zTop - zUnder), 4)
    PJ_Put d, "contacts", contacts
End Sub

' One TIM component: RefDes, side and box in the board STEP's frame.
Function OneTim(ByVal c As Object, ByVal md As Object) As Object
    Dim bx As Variant, d As Object
    bx = PT_CompBox(c)
    If IsEmpty(bx) Then probs.Add c.Name2 & ": no box, not sent.": Exit Function
    Set d = PJ_Dict()
    PJ_Put d, "refDes", PT_GetProp(md, "RefDes")
    If haveTopZ Then PJ_Put d, "side", IIf(bx(2) + bx(5) / 2 >= topZ, "top", "bottom")
    PJ_Put d, "boxMm", Nums(Array(bx(0), bx(1), bx(2), bx(0) + bx(3), bx(1) + bx(4), bx(2) + bx(5)))
    Set OneTim = d
End Function

' Numbers as a JSON array, rounded to 0.1 um.
Function Nums(ByVal a As Variant) As Collection
    Dim i As Long
    Set Nums = New Collection
    For i = 0 To UBound(a): Nums.Add Round(CDbl(a(i)), 4): Next
End Function
