# Core API examples (VBA) from the online SolidWorks 2023 help

Fetched with api_example.py through headless Chromium (the site blocks plain scripts). Paths in the examples point to
sample files of a normal SolidWorks install. Add more with: python api_example.py <outdir> IFace2.GetTessTriangles ...

## Align_Assembly_Component_to_Assembly_Origin_and_Planes_Example_VB.htm
https://help.solidworks.com/2023/english/api/sldworksapi/Align_Assembly_Component_to_Assembly_Origin_and_Planes_Example_VB.htm
```vb
Align Assembly Component to Assembly Origin and Planes Example (VBA)
This example shows how to align an assembly component to the assembly origin and planes.
'----------------------------------------------
' Preconditions:
' 1. Open an assembly document.
' 2. Select an assembly component.
'
' Postconditions:
' 1. Aligns the selected assembly component
'    to the assembly origin and planes.
' 2. Examine the graphics area.
'-----------------------------------------------
Option Explicit
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swMathUtil As SldWorks.MathUtility
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swComp As SldWorks.Component2
    Dim swXform As SldWorks.MathTransform
    Dim i As Long
    Dim bRet As Boolean
    Set swApp = Application.SldWorks
    Set swMathUtil = swApp.GetMathUtility
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager
    Set swComp = swSelMgr.GetSelectedObjectsComponent3(1, 0)
    ' Create unit transform
    Set swXform = swMathUtil.CreateTransform(Nothing)
    swComp.Transform2 = swXform
    ' Rebuild to see updated transform
    bRet = swModel.EditRebuild3
End Sub
```

## Combine_Assembly_Components_into_Part_Example_VB.htm
https://help.solidworks.com/2023/english/api/sldworksapi/Combine_Assembly_Components_into_Part_Example_VB.htm
```vb
Combine Assembly Components into Part Example (VBA)
This example shows how to combine two assembly components into a part.
'------------------------------------------------------------
' Preconditions:
' 1. Open public_documents\samples\tutorial\api\landing_gear.sldasm.
' 2. In the FeatureManager design tree:
'    a. Click oleostrut<1>.
'    b. Ctrl+click oleopiston<1>.
'    NOTE: Both components are single-body components that
'    only contain solid bodies.
' 3. Verify that the specified part document template exists.
'
' Postconditions:
' 1. Creates a new part, which is a Boolean addition of both
'    selected components.
' 2. Click Zoom to Fit.
'
' NOTES:
' * This example:
'   * does not replace saving an assembly as a part.
'   * only illustrates the use of several geometric and
'     topological methods and properties.
' * Because the assembly is used elsewhere, do not save
'   changes.
'------------------------------------------------------------
Option Explicit
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swComp1 As SldWorks.Component2
    Dim swComp2 As SldWorks.Component2
    Dim swXform1 As SldWorks.MathTransform
    Dim swXform2 As SldWorks.MathTransform
    Dim swBody1 As SldWorks.Body2
    Dim swBody1Copy As SldWorks.Body2
    Dim swBody2 As SldWorks.Body2
    Dim swBody2Copy As SldWorks.Body2
    Dim vBodyResArr As Variant
    Dim vBodyRes As Variant
    Dim swBodyRes As SldWorks.Body2
    Dim swPartRes  As SldWorks.PartDoc
    Dim swFeatRes As SldWorks.Feature
    Dim nRetval As Long
    Dim bRet As Boolean
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager
    Set swComp1 = swSelMgr.GetSelectedObjectsComponent2(1)
    Set swComp2 = swSelMgr.GetSelectedObjectsComponent2(2)
    Set swXform1 = swComp1.Transform2
    Set swXform2 = swComp2.Transform2
    Set swBody1 = swComp1.GetBody
    Set swBody2 = swComp2.GetBody
    Set swBody1Copy = swBody1.Copy
    Set swBody2Copy = swBody2.Copy
    bRet = swBody1Copy.ApplyTransform(swXform1)
    bRet = swBody2Copy.ApplyTransform(swXform2)
    vBodyResArr = swBody1Copy.Operations2(SWBODYADD, swBody2Copy, nRetval)
    Debug.Assert Not IsEmpty(vBodyResArr)
    Set swPartRes = swApp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2014\templates\Part.prtdot", 0, 0, 0)
    For Each vBodyRes In vBodyResArr
        Set swBodyRes = vBodyRes
        Set swFeatRes = swPartRes.CreateFeatureFromBody3(swBodyRes, False, swCreateFeatureBodyCheck + swCreateFeatureBodySimplify): Debug.Assert Not swFeatRes Is Nothing
    Next
End Sub
```

## Create_and_Convert_Non-manifold_Bodies_Example_VB.htm
https://help.solidworks.com/2023/english/api/sldworksapi/Create_and_Convert_Non-manifold_Bodies_Example_VB.htm
```vb
Create and Convert Non-manifold Bodies Example (VBA)
This example shows how to create non-manifold bodies, which by default are not allowed in SOLIDWORKS, and then convert the non-manifold bodies to manifold bodies.
'-----------------------------------------------------
' Preconditions: Verify that the specified part document
' template exists.
'
' Postconditions:
' 1. Click Debug > Step Into in the IDE.
' 2. Step through the rest of the macro by pressing F8
'    while observing the graphics area.
' 3. Creates and tessellates non-manifold bodies
'    and coverts them to manifold bodies.
'----------------------------------------------------
Option Explicit
Dim swApp As SldWorks.SldWorks
Dim swModeler As SldWorks.Modeler
Dim swModel As SldWorks.ModelDoc2
Dim swModelDocExt As SldWorks.ModelDocExtension
Dim swSketchManager As SldWorks.SketchManager
Dim swFeature As SldWorks.Feature
Dim swSelMgr As SldWorks.SelectionMgr
Dim swSketchSegment As SldWorks.SketchSegment
Dim swFeatureManager As SldWorks.FeatureManager
Dim tess As SldWorks.Tessellation
Dim tool As SldWorks.Body2
Dim tgt1 As SldWorks.Body2
Dim tgt0 As SldWorks.Body2
Dim sketchLines As Variant
Dim resVar As Variant
Dim resvar2 As Variant
Dim manifVar As Variant
Dim vFacetId As Variant
Dim vFinId As Variant
Dim vVertexId As Variant
Dim vVertex1 As Variant
Dim vVertex2 As Variant
Dim f As Object
Dim boolstatus As Boolean
Dim i As Long
Dim j As Long
Dim clr(0 To 1) As Long
Sub DisplayBody(ByVal b As Object, col As Long)
    Call b.Display2(swModel, col, swTempBodySelectOptions_e.swTempBodySelectable)
End Sub
Sub HideBody(ByVal b As Object)
    Call b.Hide(swModel)
End Sub
Sub main()
    Set swApp = Application.SldWorks
    Set swModeler = swApp.GetModeler
    'SOLIDWORKS requires this option to be
    'false, so make sure it is set to false
    swModeler.GeneralTopology = False
    'Create part having a tool and target bodies
    Set swModel = swApp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2016\templates\Part.prtdot", 0, 0, 0)
    Set swModelDocExt = swModel.Extension
    boolstatus = swModelDocExt.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstToRectEntity, swUserPreferenceOption_e.swDetailingNoOptionSpecified, False)
    boolstatus = swModelDocExt.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
    Set swSketchManager = swModel.SketchManager
    sketchLines = swSketchManager.CreateCornerRectangle(0, 0, 0, 0.13786334229408, 0.069192775961991, 0)
    boolstatus = swModelDocExt.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
    boolstatus = swModelDocExt.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
    boolstatus = swModelDocExt.SelectByID2("Line4", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
    boolstatus = swModelDocExt.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
    boolstatus = swModelDocExt.SelectByID2("Sketch2", "SKETCH", 0, 0, 0, False, 4, Nothing, 0)
    Set swFeatureManager = swModel.FeatureManager
    Set swFeature = swFeatureManager.FeatureExtrusion2(True, False, False, swEndConditions_e.swEndCondBlind, 0, 0.01524, 0.00254, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, swStartConditions_e.swStartSketchPlane, 0, False)
    Set swSelMgr = swModel.SelectionManager
    swSelMgr.EnableContourSelection = False
    swModel.ClearSelection2 True
    swModelDocExt.SelectByID2("Front?Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
    swSketchManager.InsertSketch (True)
    Set swSketchSegment = swSketchManager.CreateLine(0#, 0.034596, 0#, 0.137863, 0.034596, 0#)
    Set swSketchSegment = swSketchManager.CreateLine(0.068932, 0.069193, 0#, 0.068932, 0#, 0#)
    Set swSketchSegment = swSketchManager.CreateLine(0#, 0#, 0#, 0.137863, 0#, 0#)
    Set swSketchSegment = swSketchManager.CreateLine(0.137863, 0#, 0#, 0.137863, 0.069193, 0#)
    Set swSketchSegment = swSketchManager.CreateLine(0.137863, 0.069193, 0#, 0#, 0.069193, 0#)
    Set swSketchSegment = swSketchManager.CreateLine(0#, 0.069193, 0#, 0#, 0#, 0#)
    boolstatus = swModelDocExt.SelectByID2("Sketch2", "SKETCHREGION", 2.95651111315002E-02, 5.62122082077101E-02, 7.61999999999985E-03, True, 4, Nothing, 0)
    boolstatus = swModelDocExt.SelectByID2("Sketch2", "SKETCHREGION", 8.12426680973543E-02, 1.81340083381333E-02, 7.62000000000011E-03, True, 4, Nothing, 0)
    swModel.ClearSelection2 True
    boolstatus = swModelDocExt.SelectByID2("Line9", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
    boolstatus = swModelDocExt.SelectByID2("Sketch2", "SKETCH", 2.95651111315002E-02, 5.62122082077101E-02, 7.61999999999985E-03, True, 4, Nothing, 0)
    swSelMgr.EnableContourSelection = True
    boolstatus = swModelDocExt.SelectByID2("Sketch2", "SKETCHREGION", 2.95651111315002E-02, 5.62122082077101E-02, 7.61999999999985E-03, True, 4, Nothing, 0)
    boolstatus = swModelDocExt.SelectByID2("Sketch2", "SKETCHREGION", 8.12426680973543E-02, 1.81340083381333E-02, 7.62000000000011E-03, True, 4, Nothing, 0)
    Set swFeature = swFeatureManager.FeatureExtrusion2(True, False, False, swEndConditions_e.swEndCondBlind, 0, 0.01524, 0.01524, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, swStartConditions_e.swStartSketchPlane, 0, False)
    swSelMgr.EnableContourSelection = False
    'Hide the boss-extrude and sketch features
    boolstatus = swModelDocExt.SelectByID2("Boss-Extrude1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    swFeatureManager.HideBodies
    boolstatus = swModelDocExt.SelectByID2("Boss-Extrude2", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    swFeatureManager.HideBodies
    'Make selections of tool and target bodies;
    'Boss-Extrude1 is larger cube, whereas Boss-Extrude2[1]
    'and Boss-Extrude2[2] are 1/4 the size of Boss-Extrude1,
    'so (Boss-Extrude - Boss-Extrude2[1])-Boss-Extrude2[2])
    'results in non-manifold bodies; under normal conditions,
    'i.e., when non-manifold bodies are not allowed,
    'such an operation results in two bodies;
    'when creation of non-manifold bodies is allowed,
    'then one general body is the result
    boolstatus = swModelDocExt.SelectByID2("Boss-Extrude1", "SOLIDBODY", 0, 0, 0, False, 0, Nothing, 0)
    boolstatus = swModelDocExt.SelectByID2("Boss-Extrude2[1]", "SOLIDBODY", 0, 0, 0, True, 0, Nothing, 0)
    boolstatus = swModelDocExt.SelectByID2("Boss-Extrude2[2]", "SOLIDBODY", 0, 0, 0, True, 0, Nothing, 0)
    Set tool = swSelMgr.GetSelectedObject6(1, -1)
    Set tgt0 = swSelMgr.GetSelectedObject6(2, -1)
    Set tgt1 = swSelMgr.GetSelectedObject6(3, -1)
    'Create temporary bodies
    Set tool = tool.Copy
    Set tgt0 = tgt0.Copy
    Set tgt1 = tgt1.Copy
    swModel.ClearSelection2 True
    'First cut operation : Boss-Extrude1 - Boss-Extrude2[1]
    Dim errCode As Long
    resVar = tool.Operations2(swBodyOperationType_e.SWBODYCUT, tgt0, errCode)
    'SOLIDWORKS requires this option
    'to be false; thus, switch it back to false
    'as soon as your intended operations complete
    swModeler.GeneralTopology = True
    'Second cut operation: (Boss-Extrude1 - Boss-Extrude2[1])- Boss-Extrude2[2])
    resvar2 = resVar(0).Operations2(swBodyOperationType_e.SWBODYCUT, tgt1, errCode)
    'Reset the option back to false
    swModeler.GeneralTopology = False
    clr(0) = RGB(0, 0, 255)
    clr(1) = RGB(255, 0, 0)
    For i = LBound(resvar2) To UBound(resvar2)
        Call DisplayBody(resvar2(i), clr(i))
    Next i
    'Hide the displayed bodies
    For i = LBound(resvar2) To UBound(resvar2)
        HideBody (resvar2(i))
    Next i
    'Try tessellation
    'Add sketch for this face
    swModel.Insert3DSketch2 False
    'Add lines directly to sketch to increase performance
    swModel.SetAddToDB True
    Set tess = resvar2(0).GetTessellation(Empty)
    tess.NeedFaceFacetMap = True
    tess.MatchType = swTesselationMatchFacetGeometry
    boolstatus = tess.Tessellate
    Set f = resvar2(0).GetFirstFace
    While Not f Is Nothing
        vFacetId = tess.GetFaceFacets(f)
            For i = 0 To UBound(vFacetId)
                vFinId = tess.GetFacetFins(vFacetId(i))
                For j = 0 To 2
                    'Should always be three fins per facet
                    vVertexId = tess.GetFinVertices(vFinId(j))
                    'Should always be two vertices per fin
                    vVertex1 = tess.GetVertexPoint(vVertexId(0))
                    vVertex2 = tess.GetVertexPoint(vVertexId(1))
                    Call swModel.CreateLine2(vVertex1(0), vVertex1(1), vVertex1(2), vVertex2(0), vVertex2(1), vVertex2(2))
                Next j
            Next i
            Set f = f.GetNextFace
    Wend
    'Convert non-manifold bodies to manifold bodies
    manifVar = swModeler.MakeManifoldBodies(resvar2(0))
    For i = LBound(manifVar) To UBound(manifVar)
        Call DisplayBody(manifVar(i), RGB(0, 255, 0))
    Next i
    swModel.ClearSelection2 True
    For i = LBound(manifVar) To UBound(manifVar)
            HideBody (manifVar(i))
    Next i
End Sub
```

## Get_Bodies_in_Components_Example_VB.htm
https://help.solidworks.com/2023/english/api/sldworksapi/Get_Bodies_in_Components_Example_VB.htm
```vb
Get Bodies in Components Example (VBA)
This example shows how to get the number of normal and user bodies in the components in an assembly.
'-----------------------------------------------------
' Preconditions:
' 1. Verify that the specified assembly document
'    to open exists.
' 2. Open the Immediate window.
'
' Postconditions:
' 1. Gets and prints each component's name, number of
'    solid bodies, body names, and body types
'    to the Immediate window.
' 2. Right-click filterholder<1> in the FeatureManager
'    design tree and click the Open Part button.
'    Notice that there are no screw holes in the part.
' 3. Close the part and examine the filterholder<1>
'    component, which is the orange, flat, circular
'    component located on the front of the assembly
'    in the graphics area. There are screw
'    holes in the component.
' 4. Examine the filterholder<1>'s information in the
'    Immediate window. Because the component was
'    modified in the assembly, its body is identified
'    as a user body.
'
' NOTE: Because this assembly document is used by
' elsewhere, do not save changes.
'-----------------------------------------------------
Option Explicit
Dim swApp As SldWorks.SldWorks
Dim swModel As ModelDoc2
Dim swAssembly As SldWorks.AssemblyDoc
Dim vComponents As Variant
Dim oneComponent As Component2
Dim vBodies As Variant
Dim vBodyInfo As Variant
Dim errors As Long, warnings As Long
Dim i As Long, j As Long
Sub main()
Set swApp = Application.SldWorks
    ' Open the assembly
    Set swModel = swApp.OpenDoc6("C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS 2018\samples\tutorial\toolbox\lens_mount.sldasm", swDocASSEMBLY, swOpenDocOptions_Silent, "", errors, warnings)
    Set swAssembly = swModel
    ' Get the components in the assembly
    vComponents = swAssembly.GetComponents(True)
    For i = 0 To UBound(vComponents)
        Set oneComponent = vComponents(i)
        Debug.Print " "
            Debug.Print "Component name: " & oneComponent.Name2
        ' Get the solid bodies in the component
        vBodies = oneComponent.GetBodies3(SwConst.swBodyType_e.swSolidBody, vBodyInfo)
        Debug.Print "  Number of solid bodies: " & (UBound(vBodies) + 1)
        For j = 0 To UBound(vBodies)
            Debug.Print "  Body number: " & (j + 1)
            Debug.Print "  Body name: " & vBodies(j).Name
            ' Print the type of body
                Select Case vBodyInfo(j)
                Case 0
                  Debug.Print "  Body type: user"
                Case 1
                  Debug.Print "  Body type: normal"
            End Select
        Next
    Next
End Sub
```

## Get_Part_Bounding_Box_Example_VB.htm
https://help.solidworks.com/2023/english/api/sldworksapi/Get_Part_Bounding_Box_Example_VB.htm
```vb
Get Part Bounding Box Example (VBA)
This example shows how to get an accurate bounding box for a open part.
'-----------------------------------------------
'
' Problem:
' ??????IPartDoc::GetPartBox returns an approximate
' ??????bounding box and cannot be relied upon to be used
' ??????in calculations. Typically, the bounding box
' ??????returned is larger and not a minimal bounding box.
' ??????The best use for IPartDoc::GetPartbox is as a first
' ??????approximation to determine whether two parts are intersecting.
'
' ??????If there is a need for a more accurate bounding
' ??????box, then use one based on the display tessellation, which
' ??????gives a more minimal bounding box.
'
' Preconditions: Part is open.
'
' Postconditions: None
'
'-----------------------------------------------
Option Explicit
Function GetMax _
( _
????Val1 As Double, _
????Val2 As Double, _
????Val3 As Double, _
????Val4 As Double _
) As Double
' Finds maximum of four values
????GetMax = Val1
????If Val2 > GetMax Then
????????GetMax = Val2
????End If
????If Val3 > GetMax Then
????????GetMax = Val3
????End If
????If Val4 > GetMax Then
????????GetMax = Val4
????End If
End Function
Function GetMin _
( _
????Val1 As Double, _
????Val2 As Double, _
????Val3 As Double, _
????Val4 As Double _
) As Double
' Finds minimum of four values
????GetMin = Val1
????If Val2 < GetMin Then
????????GetMin = Val2
????End If
????If Val3 < GetMin Then
????????GetMin = Val3
????End If
????If Val4 < GetMin Then
????????GetMin = Val4
????End If
End Function
Sub ProcessTessTriangles _
( _
????vTessTriangles As Variant, _
????X_max As Double, _
????X_min As Double, _
????Y_max As Double, _
????Y_min As Double, _
????Z_max As Double, _
????Z_min As Double _
)
????Dim i ??????????????????As Long
????For i = 0 To UBound(vTessTriangles) / (1 * 9) - 1
' ???????' Debugging output only
' ???????Debug.Print "Pt(" + Str(i) + ") = "
' ???????Debug.Print " (" + _
' ???????????Str(vTessTriangles(9 * i + 0)) + "," + _
' ???????????Str(vTessTriangles(9 * i + 1)) + "," + _
' ???????????Str(vTessTriangles(9 * i + 2)) + ")"
' ???????Debug.Print " (" + _
' ???????????Str(vTessTriangles(9 * i + 3)) + "," + _
' ???????????Str(vTessTriangles(9 * i + 4)) + "," + _
' ???????????Str(vTessTriangles(9 * i + 5)) + ")"
' ???????Debug.Print " (" + _
' ???????????Str(vTessTriangles(9 * i + 6)) + "," + _
' ???????????Str(vTessTriangles(9 * i + 7)) + "," + _
' ???????????Str(vTessTriangles(9 * i + 8)) + ")"
????????X_max = GetMax((vTessTriangles(9 * i + 0)), (vTessTriangles(9 * i + 3)), (vTessTriangles(9 * i + 6)), X_max)
????????X_min = GetMin((vTessTriangles(9 * i + 0)), (vTessTriangles(9 * i + 3)), (vTessTriangles(9 * i + 6)), X_min)
????????Y_max = GetMax((vTessTriangles(9 * i + 1)), (vTessTriangles(9 * i + 4)), (vTessTriangles(9 * i + 7)), Y_max)
????????Y_min = GetMin((vTessTriangles(9 * i + 1)), (vTessTriangles(9 * i + 4)), (vTessTriangles(9 * i + 7)), Y_min)
????????Z_max = GetMax((vTessTriangles(9 * i + 2)), (vTessTriangles(9 * i + 5)), (vTessTriangles(9 * i + 8)), Z_max)
????????Z_min = GetMin((vTessTriangles(9 * i + 2)), (vTessTriangles(9 * i + 5)), (vTessTriangles(9 * i + 8)), Z_min)
????Next i
End Sub
Sub ProcessBodies _
( _
????vBodies As Variant, _
????X_max As Double, _
????X_min As Double, _
????Y_max As Double, _
????Y_min As Double, _
????Z_max As Double, _
????Z_min As Double _
)
????Dim swBody ?????????????As SldWorks.body2
????Dim swFace ?????????????As SldWorks.face2
????Dim vTessTriangles ?????As Variant
????Dim i ??????????????????As Long
????' Probably empty if no reference surfaces
????If IsEmpty(vBodies) Then Exit Sub
????For i = 0 To UBound(vBodies)
????????Set swBody = vBodies(i)
????????Set swFace = swBody.GetFirstFace
????????While Not swFace Is Nothing
????????????vTessTriangles = swFace.GetTessTriangles(True)
????????????ProcessTessTriangles vTessTriangles, X_max, X_min, Y_max, Y_min, Z_max, Z_min
????????????Set swFace = swFace.GetNextFace
????????Wend
????Next i
End Sub
Sub main()
????Const MaxDouble ????????As Double = 1.79769313486231E+308
????Const MinDouble ????????As Double = -1.79769313486231E+308
????Dim swApp ??????????????As SldWorks.SldWorks
????Dim swModel ????????????As SldWorks.ModelDoc2
????Dim swPart ?????????????As SldWorks.PartDoc
????Dim vBodies ????????????As Variant
????Dim vBoundBox ??????????As Variant
????Dim X_max ??????????????As Double
????Dim X_min ??????????????As Double
????Dim Y_max ??????????????As Double
????Dim Y_min ??????????????As Double
????Dim Z_max ??????????????As Double
????Dim Z_min ??????????????As Double
????Dim i ??????????????????As Long
????Set swApp = CreateObject("SldWorks.Application")
????Set swModel = swApp.ActiveDoc
????Set swPart = swModel
????' Initialise to large/small values
????X_max = MinDouble
????X_min = MaxDouble
????Y_max = MinDouble
????Y_min = MaxDouble
????Z_max = MinDouble
????Z_min = MaxDouble
????' Solid body
????vBodies = swPart.GetBodies2(swSolidBody, False)
????ProcessBodies vBodies, X_max, X_min, Y_max, Y_min, Z_max, Z_min
????' Reference surfaces
????vBodies = swPart.GetBodies2(swSheetBody, False)
????ProcessBodies vBodies, X_max, X_min, Y_max, Y_min, Z_max, Z_min
????' Approximate bounding box
????vBoundBox = swPart.GetPartBox(True)
????Debug.Print "Tessellation Quality = " + Str(swModel.GetTessellationQuality)
????Debug.Print ""
????Debug.Print "PartBox = "
????Debug.Print " ?(" + _
????????????????????Str(vBoundBox(0) * 1000#) + "," + _
????????????????????Str(vBoundBox(1) * 1000#) + "," + _
????????????????????Str(vBoundBox(2) * 1000#) + ") mm"
????Debug.Print " ?(" + _
????????????????????Str(vBoundBox(3) * 1000#) + "," + _
????????????????????Str(vBoundBox(4) * 1000#) + "," + _
????????????????????Str(vBoundBox(5) * 1000#) + ") mm"
????Debug.Print ""
????Debug.Print "TessBox = "
????Debug.Print " ?(" + _
????????????????????Str(X_min * 1000#) + "," + _
????????????????????Str(Y_min * 1000#) + "," + _
????????????????????Str(Z_min * 1000#) + ") mm"
????Debug.Print " ?(" + _
????????????????????Str(X_max * 1000#) + "," + _
????????????????????Str(Y_max * 1000#) + "," + _
????????????????????Str(Z_max * 1000#) + ") mm"
????Debug.Print ""
End Sub
'------------------------------------------
```

## Get_Transforms_of_Assembly_Components_Example_VB.htm
https://help.solidworks.com/2023/english/api/sldworksapi/Get_Transforms_of_Assembly_Components_Example_VB.htm
```vb
Get Transforms of Assembly Components Example (VBA)
This example shows how to recursively retrieve the transform for each child component in an assembly.
'-----------------------------------------------
'
' Preconditions:
' ??????(1) Assembly is open.
' ??????(2) Assembly is fully resolved.
'
' Postconditions: None
'
' Notes:
' ??????(1) Root component does not have a name
'
' ??????(2) Root component has a NULL transform, that is, no rotation,
' ??????????translation, or scaling
'
' ??????(3) All child component transforms are relative to their
' ??????????root component
'
'-----------------------------------------------
Option Explicit
Sub OutputCompXform _
( _
????swComp As SldWorks.Component2, _
????nLevel As Long _
)
????Dim vChild ?????????????????????As Variant
????Dim swChildComp ????????????????As SldWorks.Component2
????Dim sPadStr ????????????????????As String
????Dim swCompXform ????????????????As SldWorks.MathTransform
????Dim vXform ?????????????????????As Variant
????Dim i ??????????????????????????As Long
????For i = 0 To nLevel
????????sPadStr = sPadStr & " ?"
????Next i
????' Null for root component
????Set swCompXform = swComp.Transform2
????If Not swCompXform Is Nothing Then
????????vXform = swCompXform.ArrayData
????????' Root component has no name
????????Debug.Print sPadStr & "Component = " & swComp.Name2 & " (" & swComp.ReferencedConfiguration & ")"
????????Debug.Print sPadStr & " ?Suppr ??= " & swComp.IsSuppressed
????????Debug.Print sPadStr & " ?Hidden ?= " & swComp.IsHidden(False)
????????Debug.Print sPadStr & " ?Rot1 ?= (" + _
????????????Str(vXform(0)) + ", " + _
????????????Str(vXform(1)) + ", " + _
????????????Str(vXform(2)) + ")" _
????????Debug.Print sPadStr & " ?Rot2 ?= (" + _
????????????Str(vXform(3)) + ", " + _
????????????Str(vXform(4)) + ", " + _
????????????Str(vXform(5)) + ")" _
????????Debug.Print sPadStr & " ?Rot3 ?= (" + _
????????????Str(vXform(6)) + ", " + _
????????????Str(vXform(7)) + ", " + _
????????????Str(vXform(8)) + ")" _
????????Debug.Print sPadStr & " ?Trans = (" + _
????????????Str(vXform(9)) + ", " + _
????????????Str(vXform(10)) + ", " + _
????????????Str(vXform(11)) + ")" _
????????Debug.Print sPadStr & " ?Scale = " + Str(vXform(12))
????????Debug.Print ""
????End If
????' Recurse into subassembly
????vChild = swComp.GetChildren
????For i = 0 To UBound(vChild)
????????Set swChildComp = vChild(i)
????????OutputCompXform swChildComp, nLevel + 1
????Next i
End Sub
Sub main()
????Dim swApp ??????????????????????As SldWorks.SldWorks
????Dim swModel ????????????????????As SldWorks.ModelDoc2
????Dim swConf ?????????????????????As SldWorks.configuration
????Dim swRootComp ?????????????????As SldWorks.Component2
????Dim bRet ???????????????????????As Boolean
????Set swApp = CreateObject("SldWorks.Application")
????Set swModel = swApp.ActiveDoc
????Set swConf = swModel.GetActiveConfiguration
????Set swRootComp = swConf.GetRootComponent
????Debug.Print "File = " & swModel.GetPathName
????OutputCompXform swRootComp, 0
End Sub
'---------------------------------------
```

## Transform_Point_from_Component_Space_to_Assembly_Space_Example_VB.htm
https://help.solidworks.com/2023/english/api/sldworksapi/Transform_Point_from_Component_Space_to_Assembly_Space_Example_VB.htm
```vb
Transform Point from Component Space to Assembly Space Example (VBA)
This example shows how to transform a point from component space to assembly space.
'--------------------------------------------------------------
' Preconditions:
' 1. Verify that the specified assembly document to open
'    exists.
' 2. Open the Immediate window.
'
' Postconditions:
' 1. Opens the specified assembly document.
' 2. Selects a component.
' 3. Transforms the component's origin to a point in
'    assembly space.
' 4. Examine the Immediate window.
'--------------------------------------------------------------
Option Explicit
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swMathUtil As SldWorks.MathUtility
    Dim swModel As SldWorks.ModelDoc2
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swComp As SldWorks.Component2
    Dim swXform As SldWorks.MathTransform
    Dim nPt(2) As Double
    Dim vPt As Variant
    Dim swPt As SldWorks.MathPoint
    Dim bRet As Boolean
    Dim errors As Long
    Dim warnings As Long
    Dim fileName As String
    Set swApp = Application.SldWorks
    ' Open assembly
    fileName = "C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS 2018\samples\tutorial\smartcomponents\stepped_shaft.sldasm"
    Set swModel = swApp.OpenDoc6(fileName, swDocumentTypes_e.swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, warnings)
    Set swModelDocExt = swModel.Extension
    bRet = swModelDocExt.SelectByID2("stepped_shaft-1@stepped_shaft", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    Set swMathUtil = swApp.GetMathUtility
    Set swSelMgr = swModel.SelectionManager
    Set swComp = swSelMgr.GetSelectedObjectsComponent(1)
    Set swXform = swComp.Transform2
    ' Point at component origin
    nPt(0) = 0#
    nPt(1) = 0#
    nPt(2) = 0#
    vPt = nPt
    Set swPt = swMathUtil.CreatePoint(vPt)
    Set swPt = swPt.MultiplyTransform(swXform)
    Debug.Print "File = " & swModel.GetPathName
    Debug.Print "  Component = " & swComp.Name2 & " [" & swComp.GetPathName & "]"
    Debug.Print "    Point in component = (" & nPt(0) * 1000# & ", " & nPt(1) * 1000# & ", " & nPt(2) * 1000# & ") mm"
    Debug.Print "    Point in assembly = (" & swPt.ArrayData(0) * 1000# & ", " & swPt.ArrayData(1) * 1000# & ", " & swPt.ArrayData(2) * 1000# & ") mm"
End Sub
```
