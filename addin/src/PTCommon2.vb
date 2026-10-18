Option Strict Off
Option Explicit On

Imports System
Imports Microsoft.VisualBasic
Imports SW = SolidWorks.Interop.sldworks

' Port of src/PTCommon.bas, part 2: SolidWorks helpers shared by the build macros.
' Same names and behaviour as the VBA. Lengths in the arguments are mm (board frame); conversion to metres happens here.
' Late binding (Option Strict Off) like the VBA, typed interop only where the VBA needed typed objects.
Public Module PTCommon2

    Public PT_ShapeRot As Double      ' rotated box part: degrees, counter-clockwise seen from +Z, about the footprint centre
    Public PT_ShapeL As Double        ' rotated box part: rectangle size along its own X (mm)
    Public PT_ShapeW As Double        ' rotated box part: rectangle size along its own Y (mm)
    Private Const PT_PI As Double = 3.14159265358979
    Public Const PT_swSolidBody As Integer = 0
    Public Const PT_swSuppressed As Integer = 0
    Public Const PT_swFullyResolved As Integer = 2
    Public Const PT_swDeleteAbsorbed As Integer = 2
    Public Const PT_swSaveSilent As Integer = 1
    Public Const PT_swSaveReferenced As Integer = 4
    Public Const PT_swOpenSilent As Integer = 1
    Public Const PT_swDocPART As Integer = 1
    Public Const PT_swDocASSEMBLY As Integer = 2

    ' ---------------------------------------------------------------- typed calls
    ' SolidWorks rejects late-bound ByRef Integer / String / Boolean arguments from .NET (they arrive as VARIANT*, the
    ' VBA macros sent VT_I4 | VT_BYREF): "Type mismatch". Every call with such an argument goes through these.
    Public Function SwTyped() As SW.SldWorks
        Return CType(swApp, SW.SldWorks)
    End Function

    Public Function SwOpenDoc6(ByVal path As String, ByVal docType As Integer, ByVal opts As Integer, ByRef errs As Integer, ByRef warns As Integer) As Object
        Return SwTyped().OpenDoc6(path, docType, opts, "", errs, warns)
    End Function

    Public Function SwSaveAs(ByVal md As Object, ByVal path As String, ByRef errs As Integer, ByRef warns As Integer) As Boolean
        Return CType(md, SW.ModelDoc2).Extension.SaveAs(path, 0, PT_swSaveSilent, Nothing, errs, warns)
    End Function

    Public Function SwSave3(ByVal md As Object, ByVal opts As Integer, ByRef errs As Integer, ByRef warns As Integer) As Boolean
        Return CType(md, SW.ModelDoc2).Save3(opts, errs, warns)
    End Function

    Public Function SwGetOpenFileName(ByVal title As String, ByVal initial As String, ByVal filter As String) As String
        Dim opt As Integer = 0, cfg As String = "", disp As String = ""
        Dim f As String = SwTyped().GetOpenFileName(title, initial, filter, opt, cfg, disp)
        If f Is Nothing Then Return ""
        Return f
    End Function

    Public Function SwLoadFile4(ByVal path As String, ByRef errs As Integer) As Object
        Dim imp As Object = SwTyped().GetImportFileData(path)
        Return SwTyped().LoadFile4(path, "r", imp, errs)
    End Function

    ' ---------------------------------------------------------------- properties
    Public Function PT_GetProp(ByVal md As Object, ByVal fld As String) As String
        Dim valOut As String = "", resOut As String = "", wasRes As Boolean, linkProp As Boolean
        If md Is Nothing Then Return ""
        Try
            CType(md, SW.ModelDoc2).Extension.CustomPropertyManager("").Get6(fld, False, valOut, resOut, wasRes, linkProp)
        Catch
            Return ""
        End Try
        If resOut <> "" Then Return resOut
        If valOut Is Nothing Then Return ""
        Return valOut
    End Function

    Public Function PT_PropNum(ByVal md As Object, ByVal fld As String) As Double
        Dim s As String = PT_GetProp(md, fld)
        If s <> "" AndAlso IsNumeric(s) Then Return CDbl(s)
        Return 0
    End Function

    Public Sub PT_SetProp(ByVal md As Object, ByVal fld As String, ByVal valS As String)
        md.Extension.CustomPropertyManager("").Add3(fld, 30, valS, 2)       ' text, replace value
    End Sub

    ' ---------------------------------------------------------------- documents
    Public Sub PT_ActivateDoc(ByVal md As Object)
        Dim errs As Integer = 0
        SwTyped().ActivateDoc3(CStr(md.GetTitle), False, 0, errs)
    End Sub

    ' Thermal assembly path from Settings (blank = <prefix>Thermal.sldasm in the output folder).
    Public Function PT_AsmPath(ByVal wb As Object) As String
        Dim s As String = PT_NameStr(wb, "Set_AsmName", "")
        If s = "" Then s = PT_SafeFile(PT_Prefix(wb)) & "Thermal.sldasm"
        If InStr(s, "\") = 0 Then s = PT_OutputFolder(wb) & s
        If LCase(Right(s, 7)) <> ".sldasm" Then s = s & ".sldasm"
        Return s
    End Function

    ' The assembly if it is open or on disk, else Nothing.
    Public Function PT_OpenAsm(ByVal asmPath As String) As Object
        Dim errs As Integer = 0, warns As Integer = 0
        Dim md As Object = swApp.GetOpenDocumentByName(asmPath)
        If md Is Nothing Then
            If System.IO.File.Exists(asmPath) Then md = SwOpenDoc6(asmPath, PT_swDocASSEMBLY, PT_swOpenSilent, errs, warns)
        End If
        Return md
    End Function

    ' Template path from Settings, else the SolidWorks default; "" if the file is not there.
    Public Function PT_Template(ByVal wb As Object, ByVal nameS As String, ByVal swPref As Integer) As String
        Dim s As String = PT_NameStr(wb, nameS, "")
        If s = "" Then s = CStr(swApp.GetUserPreferenceStringValue(swPref))
        If s <> "" Then
            If Not System.IO.File.Exists(s) Then s = ""
        End If
        Return s
    End Function

    ' Top-level components whose part carries Role = roleS. Fills have (Key -> component) and wasSupp (Key -> True if it
    ' was suppressed). Suppressed ones are resolved so they can be read; the caller suppresses again what it does not
    ' need. Returns the last one found (for single parts).
    Public Function PT_ReadTagged(ByVal asm As Object, ByVal roleS As String, ByVal have As Dict, ByVal wasSupp As Dict) As Object
        Dim comps As Object, c As Object, md As Object, keyS As String, supp As Boolean, lastC As Object = Nothing
        asm.ResolveAllLightWeightComponents(False)
        comps = asm.GetComponents(True)
        If IsEmpty(comps) Then Return Nothing
        For i As Integer = 0 To UBound(comps)
            c = comps(i)
            supp = (c.GetSuppression2 = PT_swSuppressed)
            If supp Then c.SetSuppression2(PT_swFullyResolved)
            md = c.GetModelDoc2
            If PT_GetProp(md, "Role") = roleS Then
                keyS = PT_GetProp(md, "Key")
                If keyS = "" Then keyS = "(no key) " & CStr(c.Name2)
                If Not have.Exists(keyS) Then
                    have.Add(keyS, c)
                    wasSupp.Add(keyS, supp)
                    lastC = c
                End If
            ElseIf supp Then
                c.SetSuppression2(PT_swSuppressed)
            End If
        Next
        Return lastC
    End Function

    ' Suppresses tagged components whose key is not in want. Logs the ones that were active. Returns that count.
    Public Function PT_SuppressUnwanted(ByVal have As Dict, ByVal wasSupp As Dict, ByVal want As Dict,
                                        ByVal macroName As String, ByVal doLog As Boolean) As Integer
        Dim c As Object, md As Object, cnt As Integer = 0
        For Each it As Object In have.Keys
            If Not want.Exists(it) Then
                c = have(it)
                If Not CBool(wasSupp(it)) Then
                    md = c.GetModelDoc2
                    If doLog Then PT_LogAdd(macroName, CStr(it), PT_GetProp(md, "Name"), PT_GetProp(md, "RefDes"),
                                            "removed", "Suppressed in the assembly", "", "")
                    cnt += 1
                End If
                c.SetSuppression2(PT_swSuppressed)
            End If
        Next
        Return cnt
    End Function

    ' ---------------------------------------------------------------- features and faces
    Public Function PT_FindFeature(ByVal md As Object, ByVal featName As String) As Object
        Dim f As Object = md.FirstFeature
        Do While f IsNot Nothing
            If CStr(f.Name) = featName Then Return f
            f = f.GetNextFeature
        Loop
        Return Nothing
    End Function

    ' nth reference plane of the part (1 = Front, the XY plane in the standard templates)
    Public Function PT_RefPlane(ByVal md As Object, ByVal nth As Integer) As Object
        Dim f As Object = md.FirstFeature, k As Integer = 0
        Do While f IsNot Nothing
            If CStr(f.GetTypeName2) = "RefPlane" Then
                k += 1
                If k = nth Then Return f
            End If
            f = f.GetNextFeature
        Loop
        Return Nothing
    End Function

    Public Function PT_CountFeatures(ByVal md As Object, ByVal typeName2 As String) As Integer
        Dim f As Object = md.FirstFeature, n As Integer = 0
        Do While f IsNot Nothing
            If CStr(f.GetTypeName2) = typeName2 Then n += 1
            f = f.GetNextFeature
        Loop
        Return n
    End Function

    Public Function PT_LastFeatureOfType(ByVal md As Object, ByVal typeName2 As String) As Object
        Dim f As Object = md.FirstFeature, lastF As Object = Nothing
        Do While f IsNot Nothing
            If CStr(f.GetTypeName2) = typeName2 Then lastF = f
            f = f.GetNextFeature
        Loop
        Return lastF
    End Function

    Public Sub PT_DeleteFeature(ByVal md As Object, ByVal feat As Object)
        md.ClearSelection2(True)
        feat.Select2(False, 0)
        md.Extension.DeleteSelection2(PT_swDeleteAbsorbed)          ' also deletes its sketch
        md.ClearSelection2(True)
    End Sub

    ' model point (mm) -> sketch coordinates (m) of the active sketch
    Public Function PT_ToSketch(ByVal xf As Object, ByVal x As Double, ByVal y As Double, ByVal z As Double) As Double()
        Dim d(2) As Double
        d(0) = x / 1000.0 : d(1) = y / 1000.0 : d(2) = z / 1000.0
        Dim mu As SW.MathUtility = CType(swApp.GetMathUtility, SW.MathUtility)
        Dim xfT As SW.MathTransform = CType(xf, SW.MathTransform)
        Dim pt As SW.MathPoint = CType(mu.CreatePoint(d), SW.MathPoint)
        pt = CType(pt.MultiplyTransform(xfT), SW.MathPoint)
        Return CType(pt.ArrayData, Double())
    End Function

    ' Rectangle from (x0,y0) to (x1,y1), board frame mm, on plane pl whose height is zPlane (mm).
    ' Uses the sketch's own transform, so the plane orientation does not matter. Returns the sketch feature.
    ' Drawn as four lines: CreateCornerRectangle returns nothing for a rectangle with a corner on the sketch origin and a
    ' side under about 4 mm (tested 2026-10-18 on SolidWorks 2023), which made small parts fail.
    Public Function PT_SketchRect(ByVal md As Object, ByVal pl As Object, ByVal x0 As Double, ByVal y0 As Double,
                                  ByVal x1 As Double, ByVal y1 As Double, ByVal zPlane As Double) As Object
        Dim sk As Object, xf As Object, a As Double(), b As Double(), skF As Object = Nothing
        Dim nBefore As Integer = PT_CountFeatures(md, "ProfileFeature")
        For tryN As Integer = 0 To 1
            md.ClearSelection2(True)
            pl.Select2(False, 0)
            md.SketchManager.InsertSketch(True)
            sk = md.SketchManager.ActiveSketch
            If sk Is Nothing Then Err.Raise(vbObjectError + 20, , "could not start a sketch on " & CStr(pl.Name))
            xf = sk.ModelToSketchTransform
            a = PT_ToSketch(xf, x0, y0, zPlane)
            b = PT_ToSketch(xf, x1, y1, zPlane)
            md.SketchManager.AddToDB = (tryN = 0)          ' second try: normal sketch mode
            md.SketchManager.CreateLine(a(0), a(1), 0.0, b(0), a(1), 0.0)
            md.SketchManager.CreateLine(b(0), a(1), 0.0, b(0), b(1), 0.0)
            md.SketchManager.CreateLine(b(0), b(1), 0.0, a(0), b(1), 0.0)
            md.SketchManager.CreateLine(a(0), b(1), 0.0, a(0), a(1), 0.0)
            md.SketchManager.AddToDB = False
            md.SketchManager.InsertSketch(True)
            md.ClearSelection2(True)
            If PT_CountFeatures(md, "ProfileFeature") > nBefore Then
                skF = PT_LastFeatureOfType(md, "ProfileFeature")
                Exit For
            End If
        Next
        If skF Is Nothing Then Err.Raise(vbObjectError + 21, , "the sketch was not made (rectangle " &
            Format(Math.Abs(x1 - x0), "0.000") & " x " & Format(Math.Abs(y1 - y0), "0.000") & " mm)")
        Return skF
    End Function

    ' Blind boss from a sketch. down = True extrudes the other way. Returns the feature.
    Public Function PT_Boss(ByVal md As Object, ByVal sk As Object, ByVal depthMm As Double, ByVal down As Boolean) As Object
        md.ClearSelection2(True)
        sk.Select2(False, 0)
        Dim feat As Object = md.FeatureManager.FeatureExtrusion2(True, False, down, 0, 0, depthMm / 1000.0, 0.0, False, False, False, False,
                                                                 0.0, 0.0, False, False, False, False, True, True, True, 0, 0.0, False)
        md.ClearSelection2(True)
        If feat Is Nothing Then Err.Raise(vbObjectError + 22, , "the extrusion failed")
        Return feat
    End Function

    ' Cut through everything, both directions. Returns the feature (Nothing if SolidWorks refused, e.g. no overlap).
    Public Function PT_CutThrough(ByVal md As Object, ByVal sk As Object) As Object
        md.ClearSelection2(True)
        sk.Select2(False, 0)
        Dim feat As Object = md.FeatureManager.FeatureCut4(False, False, False, 1, 1, 0.01, 0.01, False, False, False, False,
                                                           0.0, 0.0, False, False, False, False, False, True, True, False, False, False,
                                                           0, 0.0, False, False)
        md.ClearSelection2(True)
        Return feat
    End Function

    Public Sub PT_FlipExtrude(ByVal md As Object, ByVal feat As Object)
        Dim def As Object = feat.GetDefinition
        def.AccessSelections(md, NullDisp)
        def.ReverseDirection = Not CBool(def.ReverseDirection)
        feat.ModifyDefinition(def, md, NullDisp)
    End Sub

    Public Sub PT_SetExtrudeDepth(ByVal md As Object, ByVal feat As Object, ByVal mm As Double)
        Dim def As Object = feat.GetDefinition
        def.AccessSelections(md, NullDisp)
        def.SetDepth(True, mm / 1000.0)
        feat.ModifyDefinition(def, md, NullDisp)
    End Sub

    Public Function PT_ExtrudeDepth(ByVal feat As Object) As Double
        Dim def As Object = feat.GetDefinition
        Return CDbl(def.GetDepth(True)) * 1000.0
    End Function

    Public Sub PT_SetPlaneOffset(ByVal md As Object, ByVal feat As Object, ByVal mm As Double, ByVal flipIt As Boolean)
        Dim def As Object = feat.GetDefinition
        def.AccessSelections(md, NullDisp)
        def.Distance = mm / 1000.0
        If flipIt Then def.ReverseDirection = Not CBool(def.ReverseDirection)
        feat.ModifyDefinition(def, md, NullDisp)
    End Sub

    Public Function PT_PlaneOffset(ByVal feat As Object) As Double
        Dim def As Object = feat.GetDefinition
        Return CDbl(def.Distance) * 1000.0
    End Function

    ' Z range (mm) of the solid bodies of a part: {zmin, zmax}, or Nothing
    Public Function PT_BodyZRange(ByVal md As Object) As Double()
        Dim bodies As Object = md.GetBodies2(PT_swSolidBody, True)
        If IsEmpty(bodies) Then Return Nothing
        Dim zmin As Double = 1.0E+30, zmax As Double = -1.0E+30, bb As Object
        For i As Integer = 0 To UBound(bodies)
            bb = bodies(i).GetBodyBox
            If bb(2) * 1000.0 < zmin Then zmin = bb(2) * 1000.0
            If bb(5) * 1000.0 > zmax Then zmax = bb(5) * 1000.0
        Next
        Return New Double() {zmin, zmax}
    End Function

    ' Lowest Z (mm) over the faces a feature made, or 1E+30 if it made none
    Public Function PT_FeatureZMin(ByVal feat As Object) As Double
        Dim zmin As Double = 1.0E+30, bb As Object
        Dim faces As Object = feat.GetFaces
        If Not IsEmpty(faces) Then
            For i As Integer = 0 To UBound(faces)
                bb = faces(i).GetBox
                If bb(2) * 1000.0 < zmin Then zmin = bb(2) * 1000.0
            Next
        End If
        Return zmin
    End Function

    Public Function PT_AllFaces(ByVal md As Object) As Collection
        Dim res As New Collection(), faces As Object
        Dim bodies As Object = md.GetBodies2(PT_swSolidBody, True)
        If Not IsEmpty(bodies) Then
            For i As Integer = 0 To UBound(bodies)
                faces = bodies(i).GetFaces
                If Not IsEmpty(faces) Then
                    For j As Integer = 0 To UBound(faces)
                        res.Add(faces(j))
                    Next
                End If
            Next
        End If
        Return res
    End Function

    ' Names the largest flat face facing up (or down) that lies at height zmm and whose centre is inside the rectangle
    ' x0..x1, y0..y1 (mm). Any other face holding that name loses it. True if named.
    Public Function PT_NameFlatFace(ByVal md As Object, ByVal up As Boolean, ByVal zmm As Double, ByVal nameS As String,
                                    ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double) As Boolean
        Dim faces As Collection, nv As Object, bb As Object, best As Object = Nothing, bestA As Double = 0, a As Double
        Dim cx As Double, cy As Double
        faces = PT_AllFaces(md)
        For Each fc As Object In faces
            Try
                If CStr(md.GetEntityName(fc)) = nameS Then md.DeleteEntityName(fc)
            Catch
            End Try
        Next
        For Each fc As Object In faces
            nv = Nothing
            Try
                nv = fc.Normal
            Catch
            End Try
            If Not IsEmpty(nv) Then
                If (up And nv(2) > 0.99) Or ((Not up) And nv(2) < -0.99) Then
                    bb = fc.GetBox
                    cx = (bb(0) + bb(3)) * 500.0 : cy = (bb(1) + bb(4)) * 500.0
                    If Math.Abs(bb(2) * 1000.0 - zmm) < 0.001 And Math.Abs(bb(5) * 1000.0 - zmm) < 0.001 And
                       cx >= x0 And cx <= x1 And cy >= y0 And cy <= y1 Then
                        a = CDbl(fc.GetArea)
                        If a > bestA Then
                            bestA = a
                            best = fc
                        End If
                    End If
                End If
            End If
        Next
        If best Is Nothing Then Return False
        Try
            If CStr(md.GetEntityName(best)) <> "" Then md.DeleteEntityName(best)
            Return CBool(md.SetEntityName(best, nameS))
        Catch
            Return False
        End Try
    End Function

    ' ---------------------------------------------------------------- footprint shape (rotated parts)
    ' A part can be a rotated rectangle instead of an axis-aligned box. BuildBoard sets PT_ShapeRot / PT_ShapeL / PT_ShapeW
    ' before PT_AddBoxPart or PT_CheckBoxPart and clears them after (PT_ShapeClear). The L and W given to those functions
    ' are then the ENVELOPE (axis-aligned size) of the rotated rectangle. The shape is stored in the part's custom
    ' properties (Rot_mdeg, OwnL_um, OwnW_um) and PT_BuildBox / PT_BuildLayeredBox read it back.
    Public Sub PT_ShapeClear()
        PT_ShapeRot = 0 : PT_ShapeL = 0 : PT_ShapeW = 0
    End Sub

    ' Axis-aligned size of a rectangle ownL x ownW turned by rotDeg.
    Public Sub PT_ShapeEnvelope(ByVal rotDeg As Double, ByVal ownL As Double, ByVal ownW As Double,
                                ByRef envL As Double, ByRef envW As Double)
        Dim ca As Double = Math.Abs(Math.Cos(rotDeg * PT_PI / 180.0)), sa As Double = Math.Abs(Math.Sin(rotDeg * PT_PI / 180.0))
        envL = ownL * ca + ownW * sa
        envW = ownL * sa + ownW * ca
    End Sub

    ' True when the shape is a real rotated rectangle whose envelope is the box lx x wy (within 0.02 mm).
    Public Function PT_ShapeIsRotated(ByVal rotDeg As Double, ByVal ownL As Double, ByVal ownW As Double,
                                      ByVal lx As Double, ByVal wy As Double) As Boolean
        Dim e1 As Double, e2 As Double
        If Math.Abs(rotDeg) < 0.05 Or ownL <= 0 Or ownW <= 0 Then Return False
        PT_ShapeEnvelope(rotDeg, ownL, ownW, e1, e2)
        Return (Math.Abs(e1 - lx) <= 0.02 And Math.Abs(e2 - wy) <= 0.02)
    End Function

    Public Sub PT_SetShapeProps(ByVal md As Object)
        PT_SetProp(md, "Rot_mdeg", CStr(CInt(Round(PT_ShapeRot * 1000.0))))
        PT_SetProp(md, "OwnL_um", PT_Um(PT_ShapeL))
        PT_SetProp(md, "OwnW_um", PT_Um(PT_ShapeW))
    End Sub

    ' Shape stored in the part: True when it is a rotated rectangle that fits the box lx x wy.
    Public Function PT_ShapeOfPart(ByVal md As Object, ByVal lx As Double, ByVal wy As Double,
                                   ByRef rotDeg As Double, ByRef ownL As Double, ByRef ownW As Double) As Boolean
        rotDeg = PT_PropNum(md, "Rot_mdeg") / 1000.0
        ownL = PT_PropNum(md, "OwnL_um") / 1000.0
        ownW = PT_PropNum(md, "OwnW_um") / 1000.0
        Return PT_ShapeIsRotated(rotDeg, ownL, ownW, lx, wy)
    End Function

    ' Footprint sketch of a box part, from (0,0) to (lx,wy) on plane pl: a rectangle, or the rotated rectangle stored in
    ' the part when there is one (its corners stay inside 0..lx and 0..wy, so the part origin is still the box corner).
    Public Function PT_SketchFootprint(ByVal md As Object, ByVal pl As Object, ByVal lx As Double, ByVal wy As Double,
                                       ByVal zPlane As Double) As Object
        Dim rotDeg As Double, ownL As Double, ownW As Double, sk As Object, xf As Object, skF As Object = Nothing
        Dim ca As Double, sa As Double, j As Integer, sx(3) As Double, sy(3) As Double
        Dim dx As Double, dy As Double, px As Double, py As Double, a As Double()
        Dim nBefore As Integer = PT_CountFeatures(md, "ProfileFeature")
        If Not PT_ShapeOfPart(md, lx, wy, rotDeg, ownL, ownW) Then
            Return PT_SketchRect(md, pl, 0, 0, lx, wy, zPlane)
        End If
        ca = Math.Cos(rotDeg * PT_PI / 180.0) : sa = Math.Sin(rotDeg * PT_PI / 180.0)
        md.ClearSelection2(True)
        pl.Select2(False, 0)
        md.SketchManager.InsertSketch(True)
        sk = md.SketchManager.ActiveSketch
        If sk Is Nothing Then Err.Raise(vbObjectError + 23, , "could not start a sketch on " & CStr(pl.Name))
        xf = sk.ModelToSketchTransform
        For k As Integer = 0 To 3
            dx = ownL / 2.0 : dy = ownW / 2.0
            If k = 0 Or k = 3 Then dx = -dx
            If k = 0 Or k = 1 Then dy = -dy
            px = lx / 2.0 + dx * ca - dy * sa
            py = wy / 2.0 + dx * sa + dy * ca
            a = PT_ToSketch(xf, px, py, zPlane)
            sx(k) = a(0) : sy(k) = a(1)
        Next
        md.SketchManager.AddToDB = True
        For k As Integer = 0 To 3
            j = (k + 1) Mod 4
            md.SketchManager.CreateLine(sx(k), sy(k), 0.0, sx(j), sy(j), 0.0)
        Next
        md.SketchManager.AddToDB = False
        md.SketchManager.InsertSketch(True)
        md.ClearSelection2(True)
        If PT_CountFeatures(md, "ProfileFeature") > nBefore Then skF = PT_LastFeatureOfType(md, "ProfileFeature")
        If skF Is Nothing Then Err.Raise(vbObjectError + 24, , "the footprint sketch was not made (" & Format(lx, "0.000") & " x " &
                                         Format(wy, "0.000") & " mm, turned " & Format(rotDeg, "0.0") & " deg)")
        Return skF
    End Function

    ' ---------------------------------------------------------------- box parts
    ' Box from (0,0,0) to (L,W,H) in the part. An older box made by this code is deleted first, so this also resizes.
    ' When the part carries a rotated shape (PT_SetShapeProps) the footprint is that rotated rectangle.
    Public Sub PT_BuildBox(ByVal md As Object, ByVal lx As Double, ByVal wy As Double, ByVal hz As Double,
                           ByVal topName As String, ByVal bottomName As String)
        Dim f As Object, pl As Object, sk As Object, feat As Object, zr As Double()
        PT_DropLayers(md)
        f = PT_FindFeature(md, "PT_Box")
        If f IsNot Nothing Then PT_DeleteFeature(md, f)
        pl = PT_RefPlane(md, 1)
        If pl Is Nothing Then Err.Raise(vbObjectError + 10, , "the part template has no reference plane")
        sk = PT_SketchFootprint(md, pl, lx, wy, 0)
        sk.Name = "PT_BoxSketch"
        feat = PT_Boss(md, sk, hz, False)
        feat.Name = "PT_Box"
        zr = PT_BodyZRange(md)
        If zr IsNot Nothing Then
            If zr(0) < -hz / 2 Then PT_FlipExtrude(md, feat)               ' went down: turn it up
        End If
        PT_NameBoxFaces(md, hz, topName, bottomName)
    End Sub

    ' Names the top and bottom face of a box part. Raises an error when either face cannot be found or named,
    ' so the caller reports the part instead of building it without the faces SetupSim needs.
    Private Sub PT_NameBoxFaces(ByVal md As Object, ByVal hz As Double, ByVal topName As String, ByVal bottomName As String)
        If Not PT_NameFlatFace(md, True, hz, topName, -1.0E+30, -1.0E+30, 1.0E+30, 1.0E+30) Then
            Err.Raise(vbObjectError + 13, , "could not name the top face " & topName)
        End If
        If Not PT_NameFlatFace(md, False, 0, bottomName, -1.0E+30, -1.0E+30, 1.0E+30, 1.0E+30) Then
            Err.Raise(vbObjectError + 14, , "could not name the bottom face " & bottomName)
        End If
    End Sub

    ' Removes the junction core and top layer made by PT_BuildLayeredBox (M8) and clears their tag.
    Public Sub PT_DropLayers(ByVal md As Object)
        Dim f As Object
        For Each nm As String In New String() {"PT_LayerTop", "PT_LayerCore"}
            f = PT_FindFeature(md, nm)
            If f IsNot Nothing Then PT_DeleteFeature(md, f)
        Next
        Dim cur As String = PT_GetProp(md, "Layers_um")
        If cur <> "" And cur <> "-" Then PT_SetProp(md, "Layers_um", "-")      ' "-" = no layers
    End Sub

    ' Box made of three stacked bodies for the two-resistor chip model (M8): bottom layer (tBot, toward the board),
    ' junction core, top layer (tTop). Same envelope as PT_BuildBox. Feature names: PT_Box (bottom layer),
    ' PT_LayerCore, PT_LayerTop. Tag Layers_um = "<tBot>|<tTop>" in micrometres. False on failure (PT_LastErr).
    Public Function PT_BuildLayeredBox(ByVal md As Object, ByVal lx As Double, ByVal wy As Double, ByVal hz As Double,
                                       ByVal tBot As Double, ByVal tTop As Double,
                                       ByVal topName As String, ByVal bottomName As String) As Boolean
        Dim f As Object, pl As Object, sk As Object, feat As Object, zr As Double(), down As Boolean = False
        PT_LastErr = ""
        Try
            If tBot <= 0 Or tTop <= 0 Or tBot + tTop >= hz Then Throw New Exception("layer thicknesses do not fit the chip height")
            PT_DropLayers(md)
            f = PT_FindFeature(md, "PT_Box")
            If f IsNot Nothing Then PT_DeleteFeature(md, f)
            pl = PT_RefPlane(md, 1)
            If pl Is Nothing Then Throw New Exception("the part template has no reference plane")
            sk = PT_SketchFootprint(md, pl, lx, wy, 0)
            sk.Name = "PT_BoxSketch"
            feat = PT_Boss(md, sk, tBot, False)
            feat.Name = "PT_Box"
            zr = PT_BodyZRange(md)
            If zr IsNot Nothing Then
                If zr(0) < -tBot / 2 Then
                    PT_FlipExtrude(md, feat)
                    down = True
                End If
            End If
            PT_BossLayer(md, pl, lx, wy, tBot, hz - tBot - tTop, "PT_LayerCore", down)
            PT_BossLayer(md, pl, lx, wy, hz - tTop, tTop, "PT_LayerTop", down)
            PT_NameBoxFaces(md, hz, topName, bottomName)
            PT_SetProp(md, "Layers_um", PT_Um(tBot) & "|" & PT_Um(tTop))
            Return True
        Catch ex As Exception
            PT_LastErr = ex.Message
            Return False
        End Try
    End Function

    ' Separate body (not merged) from height z0 to z0 + depth (mm), made by a boss with a start offset.
    ' Tries the extrude direction that is known to work first, the other one when the result lands in the wrong place.
    Private Sub PT_BossLayer(ByVal md As Object, ByVal pl As Object, ByVal lx As Double, ByVal wy As Double,
                             ByVal z0 As Double, ByVal depth As Double, ByVal featName As String, ByRef down As Boolean)
        Dim sk As Object, feat As Object, d As Boolean, zmn As Double
        For attempt As Integer = 0 To 1
            If attempt = 0 Then d = down Else d = Not down
            sk = PT_SketchFootprint(md, pl, lx, wy, 0)
            md.ClearSelection2(True)
            sk.Select2(False, 0)
            feat = md.FeatureManager.FeatureExtrusion2(True, False, d, 0, 0, depth / 1000.0, 0.0, False, False, False, False,
                                                       0.0, 0.0, False, False, False, False, False, True, True, 3, z0 / 1000.0, False)
            md.ClearSelection2(True)
            If feat IsNot Nothing Then
                zmn = PT_FeatureZMin(feat)
                If Math.Abs(zmn - z0) < 0.001 Then
                    feat.Name = featName
                    sk.Name = featName & "Sketch"
                    down = d
                    Return
                End If
                PT_DeleteFeature(md, feat)
            Else
                PT_DeleteFeature(md, sk)
            End If
        Next
        Throw New Exception("could not make the layer " & featName)
    End Sub

    ' Bodies of a component that carry the heat: the middle one of three stacked layers (M8), else all bodies.
    Public Function PT_CoreBodies(ByVal c As Object) As Object
        Dim b As Object = PT_CompBodies(c)
        If IsEmpty(b) Then Return b
        If UBound(b) <> 2 Then Return b
        Dim zs(2) As Double, bb As Object, mid3 As Integer = 0, j As Integer
        For i As Integer = 0 To 2
            bb = b(i).GetBodyBox
            zs(i) = CDbl(bb(2))
        Next
        For i As Integer = 0 To 2
            j = 0
            If zs(i) > zs(0) Then j += 1
            If zs(i) > zs(1) Then j += 1
            If zs(i) > zs(2) Then j += 1
            If j = 1 Then mid3 = i
        Next
        Return New Object() {b(mid3)}
    End Function

    Public Sub PT_TagBox(ByVal md As Object, ByVal roleS As String, ByVal keyS As String, ByVal nameS As String, ByVal refD As String,
                         ByVal srcRow As Integer, ByVal lx As Double, ByVal wy As Double, ByVal hz As Double, ByVal macroName As String)
        PT_SetProp(md, "Role", roleS)
        PT_SetProp(md, "Key", keyS)
        PT_SetProp(md, "Name", nameS)
        PT_SetProp(md, "RefDes", refD)
        PT_SetProp(md, "SourceRow", CStr(srcRow))
        PT_SetProp(md, "L_um", PT_Um(lx))
        PT_SetProp(md, "W_um", PT_Um(wy))
        PT_SetProp(md, "H_um", PT_Um(hz))
        PT_SetProp(md, "BuiltBy", macroName)
    End Sub

    Public Sub PT_PlaceAndFix(ByVal asm As Object, ByVal comp As Object, ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double)
        Dim c As SW.Component2 = CType(comp, SW.Component2)
        Dim d(15) As Double
        d(0) = 1 : d(4) = 1 : d(8) = 1
        d(9) = x0 / 1000.0 : d(10) = y0 / 1000.0 : d(11) = z0 / 1000.0
        d(12) = 1
        Dim mu As SW.MathUtility = CType(swApp.GetMathUtility, SW.MathUtility)
        Dim xf As SW.MathTransform = CType(mu.CreateTransform(d), SW.MathTransform)
        asm.ClearSelection2(True)
        If c.IsFixed Then
            c.Select4(False, Nothing, False)
            asm.UnfixComponent()
            asm.ClearSelection2(True)
        End If
        c.Transform2 = xf
        asm.EditRebuild3()
        PT_FixComp(asm, c)
    End Sub

    Public Sub PT_FixComp(ByVal asm As Object, ByVal c As Object)
        asm.ClearSelection2(True)
        c.Select4(False, NullDisp, False)
        asm.FixComponent()
        asm.ClearSelection2(True)
    End Sub

    ' New box part file, inserted in the assembly, placed and fixed. False on failure (reason in PT_LastErr).
    Public Function PT_AddBoxPart(ByVal asm As Object, ByVal partTmpl As String, ByVal fpath As String,
                                  ByVal roleS As String, ByVal keyS As String, ByVal nameS As String, ByVal refD As String,
                                  ByVal srcRow As Integer, ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double,
                                  ByVal lx As Double, ByVal wy As Double, ByVal hz As Double,
                                  ByVal topName As String, ByVal bottomName As String, ByVal macroName As String) As Boolean
        Dim md As Object = Nothing, errs As Integer = 0, warns As Integer = 0, c As Object, ttl As String
        PT_LastErr = ""
        Try
            If swApp.GetOpenDocumentByName(fpath) IsNot Nothing Then Throw New Exception("a file with this name is already open: " & fpath)
            md = swApp.NewDocument(partTmpl, 0, 0, 0)
            If md Is Nothing Then Throw New Exception("SolidWorks could not make a new part from " & partTmpl)
            PT_SetShapeProps(md)
            PT_BuildBox(md, lx, wy, hz, topName, bottomName)
            PT_TagBox(md, roleS, keyS, nameS, refD, srcRow, lx, wy, hz, macroName)
            If Not SwSaveAs(md, fpath, errs, warns) Then
                Throw New Exception("could not save " & fpath & " (error " & errs & ")")
            End If
            ttl = CStr(md.GetTitle)
            PT_ActivateDoc(asm)
            c = asm.AddComponent5(fpath, 0, "", False, "", 0.0, 0.0, 0.0)
            If c Is Nothing Then Throw New Exception("could not insert " & fpath & " into the assembly")
            PT_PlaceAndFix(asm, c, x0, y0, z0)
            swApp.CloseDoc(ttl)
            Return True
        Catch ex As Exception
            PT_LastErr = ex.Message
            Try
                If md IsNot Nothing Then swApp.CloseDoc(md.GetTitle)           ' drop the part, saved or not, so a rerun is not blocked
            Catch
            End Try
            Try
                PT_ActivateDoc(asm)
            Catch
            End Try
            Return False
        End Try
    End Function

    ' Box of a component in the assembly (board frame, mm): {xmin, ymin, zmin, L, W, H}, or Nothing
    Public Function PT_CompBox(ByVal c As Object) As Double()
        Dim v As Object = Nothing
        Try
            v = c.GetBox(False, False)
        Catch
        End Try
        If IsEmpty(v) Then Return Nothing
        Dim outv(5) As Double
        For k As Integer = 0 To 2
            outv(k) = v(k) * 1000.0
            outv(3 + k) = (v(3 + k) - v(k)) * 1000.0
        Next
        Return outv
    End Function

    ' Existing box part: compares its real box in the assembly with the wanted box, rebuilds and moves it when needed,
    ' refreshes its tags, logs what changed. Returns 0 no change, +1 resized, +2 moved (so 3 = both), -1 failed (PT_LastErr).
    Public Function PT_CheckBoxPart(ByVal asm As Object, ByVal c As Object, ByVal roleS As String, ByVal keyS As String,
                                    ByVal nameS As String, ByVal refD As String, ByVal srcRow As Integer,
                                    ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double,
                                    ByVal lx As Double, ByVal wy As Double, ByVal hz As Double,
                                    ByVal topName As String, ByVal bottomName As String,
                                    ByVal tolPos As Double, ByVal tolSize As Double, ByVal macroName As String, ByVal doLog As Boolean) As Integer
        Dim md As Object, bx As Double(), oldS(2) As Double, newS(2) As Double, oldP(2) As Double, newP(2) As Double
        Dim sizeChanged As Boolean = False, posChanged As Boolean = False, ad As Object, errs As Integer = 0, warns As Integer = 0
        Dim fields As String() = {"X min", "Y min", "Z min", "L", "W", "H"}, refChanged As Boolean, res As Integer = 0
        Dim shapeChanged As Boolean = False, oldRot As Double, oldOwnL As Double, oldOwnW As Double
        PT_LastErr = ""
        Try
            md = c.GetModelDoc2
            If md Is Nothing Then Throw New Exception("the part file could not be loaded")
            bx = PT_CompBox(c)
            If bx Is Nothing Then Throw New Exception("the part has no geometry")
            For k As Integer = 0 To 2
                oldP(k) = bx(k) : oldS(k) = bx(3 + k)
            Next
            newS(0) = lx : newS(1) = wy : newS(2) = hz
            newP(0) = x0 : newP(1) = y0 : newP(2) = z0
            For k As Integer = 0 To 2
                If Math.Abs(newS(k) - oldS(k)) > tolSize Then sizeChanged = True
                If Math.Abs(newP(k) - oldP(k)) > tolPos Then posChanged = True
            Next
            refChanged = (PT_GetProp(md, "RefDes") <> refD)
            ' rotated footprint: parts made before this feature have no shape properties = not rotated
            oldRot = PT_PropNum(md, "Rot_mdeg") / 1000.0
            oldOwnL = PT_PropNum(md, "OwnL_um") / 1000.0
            oldOwnW = PT_PropNum(md, "OwnW_um") / 1000.0
            If Math.Abs(oldRot - PT_ShapeRot) > 0.05 Then shapeChanged = True
            If PT_ShapeRot <> 0 Or oldRot <> 0 Then
                If Math.Abs(oldOwnL - PT_ShapeL) > tolSize Or Math.Abs(oldOwnW - PT_ShapeW) > tolSize Then shapeChanged = True
            End If

            If sizeChanged Or refChanged Or shapeChanged Then
                ' rebuilt box starts at the part origin, so the component goes exactly to the wanted corner
                PT_SetShapeProps(md)
                PT_BuildBox(md, lx, wy, hz, topName, bottomName)
                PT_TagBox(md, roleS, keyS, nameS, refD, srcRow, lx, wy, hz, macroName)
                SwSave3(md, PT_swSaveSilent, errs, warns)
                PT_PlaceAndFix(asm, c, x0, y0, z0)
            Else
                If PT_GetProp(md, "Name") <> nameS Or PT_GetProp(md, "SourceRow") <> CStr(srcRow) Then
                    PT_TagBox(md, roleS, keyS, nameS, refD, srcRow, lx, wy, hz, macroName)
                    SwSave3(md, PT_swSaveSilent, errs, warns)
                End If
                If posChanged Then
                    ' shift the component by the difference, whatever is inside the part
                    ad = c.Transform2.ArrayData
                    PT_PlaceAndFix(asm, c, ad(9) * 1000.0 + (x0 - oldP(0)), ad(10) * 1000.0 + (y0 - oldP(1)), ad(11) * 1000.0 + (z0 - oldP(2)))
                ElseIf Not CBool(c.IsFixed) Then
                    PT_FixComp(asm, c)
                End If
            End If

            If sizeChanged Or shapeChanged Then
                res += 1
                If doLog Then
                    If shapeChanged Then PT_LogAdd(macroName, keyS, nameS, refD, "resized", "Rotation (deg)", Round(oldRot, 3), Round(PT_ShapeRot, 3))
                    For k As Integer = 0 To 2
                        If Math.Abs(newS(k) - oldS(k)) > tolSize Then PT_LogAdd(macroName, keyS, nameS, refD, "resized", fields(3 + k), Round(oldS(k), 3), newS(k))
                    Next
                End If
            End If
            If posChanged Then
                res += 2
                If doLog Then
                    For k As Integer = 0 To 2
                        If Math.Abs(newP(k) - oldP(k)) > tolPos Then PT_LogAdd(macroName, keyS, nameS, refD, "moved", fields(k), Round(oldP(k), 3), newP(k))
                    Next
                End If
            End If
            Return res
        Catch ex As Exception
            PT_LastErr = ex.Message
            Try
                PT_ActivateDoc(asm)
            Catch
            End Try
            Return -1
        End Try
    End Function

    ' ---------------------------------------------------------------- Simulation helpers
    ' Simulation document of the assembly (loads the Simulation add-in if needed). Nothing if not available.
    Public Function PT_GetSimDoc(ByVal asm As Object) As Object
        Try
            PT_ActivateDoc(asm)
            Dim addin As Object = swApp.GetAddInObject("SldWorks.Simulation")
            If addin Is Nothing Then
                Dim exeDir As String = CStr(swApp.GetExecutablePath)
                If Right(exeDir, 1) <> "\" Then exeDir = exeDir & "\"
                swApp.LoadAddIn(exeDir & "Simulation\cosworks.dll")
                addin = swApp.GetAddInObject("SldWorks.Simulation")
            End If
            If addin Is Nothing Then Return Nothing
            Dim cw As Object = addin.CosmosWorks
            If cw Is Nothing Then Return Nothing
            Return cw.ActiveDoc
        Catch
            Return Nothing
        End Try
    End Function

    Public Function PT_FindStudy(ByVal simDoc As Object, ByVal studyName As String) As Object
        Dim sm As Object = simDoc.StudyManager, st As Object
        For i As Integer = 0 To CInt(sm.StudyCount) - 1
            st = sm.GetStudy(i)
            If st IsNot Nothing Then
                If StrComp(CStr(st.Name), studyName, CompareMethod.Text) = 0 Then Return st
            End If
        Next
        Return Nothing
    End Function

    ' Solid bodies of a component, assembly context (array or Nothing)
    Public Function PT_CompBodies(ByVal c As Object) As Object
        Dim info As Object = Nothing
        Return c.GetBodies3(PT_swSolidBody, info)
    End Function

    ' Named face of a component's part, in assembly context, or Nothing
    Public Function PT_AsmFace(ByVal c As Object, ByVal faceName As String) As Object
        Try
            Dim md As Object = c.GetModelDoc2
            If md Is Nothing Then Return Nothing
            Dim f As Object = md.GetEntityByName(faceName, 2)
            If f Is Nothing Then Return Nothing
            Return c.GetCorrespondingEntity(f)
        Catch
            Return Nothing
        End Try
    End Function

    ' Every active (not suppressed) tagged component: Collection of {role, key, refdes, component}
    Public Function PT_ListTagged(ByVal asm As Object) As Collection
        Dim res As New Collection(), c As Object, md As Object, roleS As String
        asm.ResolveAllLightWeightComponents(False)
        Dim comps As Object = asm.GetComponents(True)
        If Not IsEmpty(comps) Then
            For i As Integer = 0 To UBound(comps)
                c = comps(i)
                If c.GetSuppression2 <> PT_swSuppressed Then
                    md = c.GetModelDoc2
                    roleS = PT_GetProp(md, "Role")
                    If roleS <> "" Then res.Add(Arr(roleS, PT_GetProp(md, "Key"), PT_GetProp(md, "RefDes"), c))
                End If
            Next
        End If
        Return res
    End Function

    ' Collection of objects -> array for the Simulation API (Nothing if none)
    Public Function PT_ToArray(ByVal coll As Collection) As Object
        If coll.Count = 0 Then Return Nothing
        Dim a(coll.Count - 1) As Object
        For i As Integer = 1 To coll.Count
            a(i - 1) = coll(i)
        Next
        Return a
    End Function

    ' ---------------------------------------------------------------- board rules
    ' Clips a part box (xmin, ymin, zmin, L, W, H in mm, 0-based array) at the PCB surface so no part goes into the board
    ' (through-hole pins, edge connectors). Top-side parts start at the PCB top, bottom-side parts end at the PCB bottom.
    ' Returns True if the box was changed.
    Public Function PT_ClipToBoard(ByVal bx() As Double, ByVal pcbT As Double) As Boolean
        Dim z0 As Double = bx(2), z1 As Double = bx(2) + bx(5), changed As Boolean = False
        If (z0 + z1) / 2 >= pcbT / 2 Then
            If z0 < pcbT - 0.0005 And z1 > pcbT + 0.0005 Then
                bx(2) = pcbT : bx(5) = z1 - pcbT : changed = True
            End If
        Else
            If z1 > 0.0005 And z0 < -0.0005 Then
                bx(5) = -z0 : changed = True
            End If
        End If
        Return changed
    End Function

    ' The build and sync macros work on the normal configuration, not on PT_HeatOnly made by SetupSim.
    Public Sub PT_LeaveHeatOnly(ByVal asm As Object)
        Try
            If CStr(asm.ConfigurationManager.ActiveConfiguration.Name) <> "PT_HeatOnly" Then Return
            Dim names As Object = asm.GetConfigurationNames
            For i As Integer = 0 To UBound(names)
                If CStr(names(i)) <> "PT_HeatOnly" Then
                    asm.ShowConfiguration2(names(i))
                    Return
                End If
            Next
        Catch
        End Try
    End Sub

    ' ---------------------------------------------------------------- results tabs (ReadResults, RunFlow)
    ' Column of "<RefDes> peak (C)" on the Sim results header row; adds the "peak" and "avg" columns when missing.
    Public Function PT_RefColumn(ByVal ws As Object, ByVal hdr As Object, ByVal refD As String) As Integer
        Dim c As Integer, want As String = refD & " peak (C)"
        Dim lastC As Integer = CInt(ws.Cells(hdr.Row, ws.Columns.Count).End(-4159).Column)
        For c = CInt(hdr.Column) + CInt(hdr.Columns.Count) To lastC
            If StrComp(PT_Str(ws.Cells(hdr.Row, c).Value), want, CompareMethod.Text) = 0 Then Return c
        Next
        c = lastC + 1
        If c < CInt(hdr.Column) + CInt(hdr.Columns.Count) Then c = CInt(hdr.Column) + CInt(hdr.Columns.Count)
        ws.Cells(hdr.Row, c).Value = want
        ws.Cells(hdr.Row, c + 1).Value = refD & " avg (C)"
        ws.Range(ws.Cells(hdr.Row, c), ws.Cells(hdr.Row, c + 1)).Interior.Color = hdr.Cells(1, 1).Interior.Color
        ws.Range(ws.Cells(hdr.Row, c), ws.Cells(hdr.Row, c + 1)).Font.Bold = True
        Return c
    End Function

    ' "U1 Body, U2 Contact R, ..." for the covered chips on the Spreader tab (Spr_Chips offsets: 0 RefDes, 7 Covered, 9 TIM mode, 16 Check)
    Public Function PT_TimModes(ByVal wb As Object) As String
        Dim v As Object(,) = ToTable(PT_Rng(wb, "Spr_Chips").Value), s As String = ""
        For r As Integer = 1 To UBound(v, 1)
            If PT_Str(v(r, 1)) <> "" And UCase(PT_Str(v(r, 8))) = "YES" And PT_Str(v(r, 17)) = "OK" Then
                s = s & IIf(s = "", "", ", ") & PT_Str(v(r, 1)) & " " & PT_Str(v(r, 10))
            End If
        Next
        Return s
    End Function

    ' Thermal solution column AF (TS_SimTj): the peak temperature of each RefDes in column B, blank where there is no result.
    ' peak = Dict, RefDes (upper case) -> degrees C.
    Public Sub PT_WriteSimTj(ByVal wb As Object, ByVal peak As Dict)
        Dim tj As Object = PT_Rng(wb, "TS_SimTj")
        If tj Is Nothing Then Return
        Dim ws As Object = tj.Worksheet
        Dim n As Integer = CInt(tj.Rows.Count)
        Dim refs As Object(,) = ToTable(ws.Range(ws.Cells(tj.Row, 2), ws.Cells(tj.Row + n - 1, 2)).Value)
        Dim outv As Object(,) = NewTable(n, 1)
        Dim rk As String
        For i As Integer = 1 To n
            rk = UCase(PT_Str(refs(i, 1)))
            If rk <> "" Then
                If peak.Exists(rk) Then outv(i, 1) = Round(CDbl(peak(rk)), 2)
            End If
        Next
        tj.Columns(1).Value = outv
    End Sub

    ' One pass over the top-level components for several roles at once (SyncFromCAD reads four roles).
    ' roles, haves and supps in the same order (see PT_ReadTagged).
    Public Sub PT_ReadRoles(ByVal asm As Object, ByVal roles As String(), ByVal haves As Dict(), ByVal supps As Dict())
        Dim c As Object, md As Object, keyS As String, tagRole As String, supp As Boolean, hit As Integer
        asm.ResolveAllLightWeightComponents(False)
        Dim comps As Object = asm.GetComponents(True)
        If IsEmpty(comps) Then Return
        For i As Integer = 0 To UBound(comps)
            c = comps(i)
            supp = (c.GetSuppression2 = PT_swSuppressed)
            If supp Then c.SetSuppression2(PT_swFullyResolved)
            md = c.GetModelDoc2
            tagRole = PT_GetProp(md, "Role")
            hit = -1
            For k As Integer = 0 To roles.Length - 1
                If roles(k) = tagRole Then
                    hit = k
                    Exit For
                End If
            Next
            If hit >= 0 Then
                keyS = PT_GetProp(md, "Key")
                If keyS = "" Then keyS = "(no key) " & CStr(c.Name2)
                If Not haves(hit).Exists(keyS) Then
                    haves(hit).Add(keyS, c)
                    supps(hit).Add(keyS, supp)
                End If
            ElseIf supp Then
                c.SetSuppression2(PT_swSuppressed)
            End If
        Next
    End Sub

    ' ---------------------------------------------------------------- colours (BuildBoard)
    ' Same palette on screen and in the workbook. The PCB is green. Part colour = palette(row index mod 12).
    Public Function PT_PcbRgb() As Integer
        Return RGB(16, 140, 39)
    End Function

    Public Function PT_PartRgb(ByVal idx As Integer) As Integer
        Dim p As Integer() = {RGB(66, 133, 244), RGB(244, 160, 0), RGB(219, 68, 55), RGB(171, 71, 188), RGB(0, 172, 193),
                              RGB(255, 202, 40), RGB(141, 110, 99), RGB(236, 64, 122), RGB(92, 107, 192), RGB(255, 112, 67),
                              RGB(120, 144, 156), RGB(38, 198, 218)}
        Dim k As Integer = idx Mod 12
        If k < 0 Then k += 12
        Return p(k)
    End Function

    ' Sets the colour of a component in the assembly (this configuration). rgbL is an RGB value. True if accepted.
    Public Function PT_ApplyColor(ByVal c As Object, ByVal rgbL As Integer) As Boolean
        Dim v(8) As Double
        v(0) = (rgbL And 255) / 255.0
        v(1) = ((rgbL \ 256) And 255) / 255.0
        v(2) = ((rgbL \ 65536) And 255) / 255.0
        v(3) = 1 : v(4) = 1 : v(5) = 0.5 : v(6) = 0.4 : v(7) = 0 : v(8) = 0
        Try
            c.MaterialPropertyValues = v
            Dim ok As Boolean = CBool(c.HasMaterialPropertyValues)
            If Not ok Then PT_LastErr = "SolidWorks did not store the colour"
            Return ok
        Catch ex As Exception
            PT_LastErr = ex.Message
            Return False
        End Try
    End Function

End Module
