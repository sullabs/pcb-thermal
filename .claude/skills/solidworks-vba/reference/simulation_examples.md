# Simulation example code (VB6 from cworksapi.chm)

Static-study examples are included where they show the call chain for materials, mesh, run and plots.
The thermal calls (AddHeatPower, AddConvection, AddRadiation) are in simulation.md.

## Apply Thermostat-controlled Heat Power for Transient Thermal Study (VB6)
```vb
Apply Thermostat-controlled Heat Power for Transient Thermal Study Example (VBA)
SOLIDWORKS Simulation
API Help
Apply Thermostat-controlled Heat Power for Transient Thermal Study
Example (VBA)
This example shows how to apply thermostat-controlled heat power and flux in
a transient thermal study.
'---------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in (in SOLIDWORKS, click
' Tools > Add-ins > SOLIDWORKS Simulation > OK ).
' 2. Add the SOLIDWORKS Simulation type library as a reference (in the IDE,
' click Tools > References > SOLIDWORKS
Simulation version type library ).
' 3. Ensure that c:\temp exists.
' 4. Ensure that the specified material library exists.
' 5. Ensure that the specified model document exists.
' 6. Open the Immediate window.
'
' Postconditions:
' 1. Opens the specified model document.
' 2. Creates Thermal_One study.
' 3. Defines default thermal study results plot and plot options.
' 4. Applies user-defined and SOLIDWORKS materials to components.
' 5. Specifies initial temperatures for assembly components.
' 6. Specifies thermostat location and the temperature range.
' 7. Prints heat flux properties to the Immediate window.
' 8. Prints radiation properties to the Immediate window.
' 9. Meshes the model.
' 10. Analyzes Thermal_One .
' 11. Creates 60 thermal plots.
' 12. Saves the Thermal1 results plot as an eDrawings document in c:\temp .
' 13. Exports Thermal_One study to the NASTRAN finite-element analysis program
' in c:\temp .
' 14. Renames Thermal_One study to TransientThermal .
' 15. Inspect the Immediate window, the Simulation study tree,
' and the graphics area.
'
' NOTES:
' * To get persistent reference identifiers (PIDs) for model selections,
' you can use
pidcollector.exe or IModelDocExtension::GetPersistReference3.
' * Solving this study can take some time. Examine the Immediate window
' to monitor the macro's progress. Done! indicates that the
macro
' has finished.
' * Because the assembly is used elsewhere, do not save any changes.
'--------------------------------------------------------------------
Option Explicit
Sub main()
Dim swApp As SldWorks.SldWorks
Dim COSMOSWORKS As CosmosWorksLib.COSMOSWORKS
Dim CWObject As CosmosWorksLib.CwAddincallback
Dim ActDoc As CosmosWorksLib.CWModelDoc
Dim StudyMngr As CosmosWorksLib.CWStudyManager
Dim Study As CosmosWorksLib.CWStudy
Dim SolidMgr As CosmosWorksLib.CWSolidManager
Dim SolidComp As CosmosWorksLib.CWSolidComponent
Dim SolidBody As CosmosWorksLib.CWSolidBody
Dim ThermalOptions As CosmosWorksLib.CWThermalStudyOptions
Dim CWConv As CosmosWorksLib.CWConvection
Dim CwMesh As CosmosWorksLib.CwMesh
Dim CWResult As CosmosWorksLib.CWResults
Dim CWMat As CosmosWorksLib.CWMaterial
Dim CWTemp As CosmosWorksLib.CWTemperature
Dim Part As SldWorks.ModelDoc2
Dim LBCMgr As CosmosWorksLib.CWLoadsAndRestraintsManager
Dim heatFlux As CosmosWorksLib.CWHeatFlux
Dim CWHeatPower As CosmosWorksLib.CWHeatPower
Dim rad As CosmosWorksLib.CWRadiation
Dim oselect1 As Object, oselect2 As Object, oselect3 As
Object
Dim oselect4 As Object, oselect5 As Object, oselect6 As
Object
Dim oselect7 As Object, oselect8 As Object, oselect9 As
Object, oselect10 As Object
Dim therm1 As Object
Dim var1 As Variant, var2 As Variant, var3 As Variant, var4
As Variant
Dim var5 As Variant, var6 As Variant, var7 As Variant, var8
As Variant
Dim var9 As Variant, var10 As Variant, var11 As Variant
Dim Temp As Variant
Dim varArray1 As Variant
Dim varArray2 As Variant
Dim varArray3 As Variant
Dim varArray4 As Variant
Dim varArray5 As Variant
Dim varArray6 As Variant
Dim bApp As Boolean
Dim selection1 As String, selection2 As String, selection3 As
String
Dim selection4 As String, selection5 As String, selection6 As
String
Dim selection7 As String, selection8 As String, selection9 As
String
Dim selection10 As String, selection11 As String
Dim longstatus As Long, longwarnings As Long
Dim errCode As Long
Dim el As Double, tl As Double
Dim nStep As Long
Dim res As Long
If swApp Is Nothing Then Set swApp =
Application.SldWorks
Set Part = swApp. OpenDoc6 ("C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS
2019\samples\Simulation Examples\Thermal\CoffeeJar.SLDASM",
swDocASSEMBLY, swOpenDocOptions_Silent, "", longstatus, longwarnings)
If Part Is Nothing Then ErrorMsg swApp, "Failed to open
document"
' Get the SOLIDWORKS Simulation add-in
object
Set CWObject = swApp. GetAddInObject ("SldWorks.Simulation")
If CWObject Is Nothing Then ErrorMsg swApp, "Failed to get
Simulation add-in"
Set COSMOSWORKS = CWObject. COSMOSWORKS
If COSMOSWORKS Is Nothing Then ErrorMsg swApp, "Failed to get
CosmosWorks object"
' Get active document
Set ActDoc = COSMOSWORKS. ActiveDoc ()
If ActDoc Is Nothing Then ErrorMsg swApp, "Failed to get
active document"
' Add a default thermal study results plot
res =
ActDoc. AddDefaultThermalStudyPlot (swsThermalResultComponentTypes_TEMP, True)
' Set default thermal plot options
' Show maximum value annotation on the plot
bApp =
ActDoc. SetSimulationOptionToggle (swsPlotAnnotationShowMaxValue, True)
Debug.Print "Show maximum value annotation on the plot? " &
ActDoc. GetSimulationOptionToggle (swsPlotAnnotationShowMaxValue)
' Set sub-folder option
bApp =
ActDoc. SetSimulationOptionToggle (swsResultFolderUnderSubFolder, True)
Debug.Print "Set sub-folder option? " &
ActDoc. GetSimulationOptionToggle (swsResultFolderUnderSubFolder)
' Put results in
C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS 2019\samples\Simulation Examples\Thermal\temp
bApp =
ActDoc. SetSimulationOptionStringValue (swsSOLIDWORKSDocumentFolderSubFolderLocation,
"temp")
Debug.Print "Place results in model document sub-folder, " &
ActDoc. GetSimulationOptionStringValue (swsSOLIDWORKSDocumentFolderSubFolderLocation)
' Set boundary settings option to
translucent single color
bApp =
ActDoc. SetSimulationOptionIntegerValue (swsPlotSettingsBoundaryOption,
swsPlotBoundaryTranslucentSingleColor)
Debug.Print "Boundary settings option as defined in
swsPlotSettingsBoundaryOption is " &
ActDoc. GetSimulationOptionIntegerValue (swsPlotSettingsBoundaryOption)
' Set boundary color; Red 1 is 0xFF0000 in
http://cloford.com/resources/colours/500col.htm
' Reverse the hexadecimal (&H0000FF) in the call below
bApp =
ActDoc. SetSimulationOptionIntegerValue (swsPlotBoundaryOptionTranslucentSingleColorSetting,
&HFF)
Debug.Print "Boundary color is " &
ActDoc. GetSimulationOptionIntegerValue (swsPlotBoundaryOptionTranslucentSingleColorSetting)
' Set boundary transparency of translucent
part color to 50%
bApp =
ActDoc. SetSimulationOptionDoubleValue (swsPlotBoundaryTransparency, 50#)
Debug.Print "Transparency of boundary translucent part color
is " & ActDoc. GetSimulationOptionDoubleValue (swsPlotBoundaryTransparency)
' Create new thermal study
Debug.Print "Creating thermal study..."
Set StudyMngr = ActDoc. StudyManager ()
If StudyMngr Is Nothing Then ErrorMsg swApp, "No
CWStudyManager object"
Set Study =
StudyMngr. CreateNewStudy3 ("Thermal_One", swsAnalysisStudyTypeThermal,
swsMeshTypeSolid, errCode)
If Study Is Nothing Then ErrorMsg swApp, "Failed to create
study", True
' Get first solid component
Set SolidMgr = Study. SolidManager
If SolidMgr Is Nothing Then ErrorMsg swApp, "No
CWStudyManager object"
Set SolidComp = SolidMgr. GetComponentAt (0, errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to get
component"
Set SolidBody =
SolidComp. GetSolidBodyAt (0, errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to get
component body"
' Add material
Set CWMat = SolidBody. GetDefaultMaterial
If CWMat Is Nothing Then ErrorMsg swApp, "No CWMaterial
object"
CWMat. MaterialName = "Coffee"
Call CWMat. SetPropertyByName ("DENS", 1000, 0)
Call CWMat. SetPropertyByName ("KX", 40, 0)
Call CWMat. SetPropertyByName ("C", 4200#, 0)
errCode =
SolidBody. SetSolidBodyMaterial (CWMat)
If errCode <> 0 Then ErrorMsg swApp, "Failed to set
material"
' Get second solid component
Set SolidComp = SolidMgr. GetComponentAt (1, errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to get
component"
Set SolidBody =
SolidComp. GetSolidBodyAt (0, errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to get
component body"
' Add material
bApp = SolidBody. SetLibraryMaterial ("C:\Program
Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\sldmaterials\solidworks
materials.sldmat", "Glass")
If bApp = False Then ErrorMsg swApp, "Failed to apply glass
material"
' Get third solid component
Set SolidComp = SolidMgr. GetComponentAt (2, errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to get
component"
Set SolidBody =
SolidComp. GetSolidBodyAt (0, errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to get
component body"
' Add material
bApp = SolidBody. SetLibraryMaterial ("C:\Program
Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\sldmaterials\solidworks
materials.sldmat", "Nylon 6/10")
If bApp = False Then ErrorMsg swApp, "Failed to apply Nylon
6/10 material"
' CoffeePot
selection1 =
"216,14,0,0,5,0,0,0,255,254,255,21,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,11,0,0,0"
selection1 = selection1 & ",Type=1"
StringtoArray selection1, var1
Set oselect1 =
Part. Extension . GetObjectByPersistReference3 ((var1), longstatus)
' Coffee
selection2 =
"216,14,0,0,5,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0"
selection2 = selection2 & ",Type=1"
StringtoArray selection2, var2
Set oselect2 =
Part. Extension . GetObjectByPersistReference3 ((var2), longstatus)
' Top
selection3 =
"216,14,0,0,5,0,0,0,255,254,255,15,84,0,111,0,112,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,13,0,0,0"
selection3 = selection3 & ",Type=1"
StringtoArray selection3, var3
Set oselect3 =
Part. Extension . GetObjectByPersistReference3 ((var3), longstatus)
' Get the four coffee faces that make up
the surface of coffee in the coffee pot
' Face 1
selection4 =
"216,14,0,0,3,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,6,99,0,111,0,102,0,102,0,101,0,101,0,126,253,47,63,11,128,255,254,255,57,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,10"
selection4 = selection4 &
"8,0,92,0,99,0,111,0,102,0,102,0,101,0,101,0,46,0,115,0,108,0,100,0,112,0,114,0,116,0,90,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,4,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,5,0,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,80,114,111,106,73,100,82,101,112,95,99,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,19,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,1,0,0,0,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,11,0,0,0,0,0,0,0,0,0"
selection4 = selection4 & ",Type=1"
StringtoArray selection4, var4
Set oselect4 =
Part. Extension . GetObjectByPersistReference3 ((var4), longstatus)
' Face 2
selection5 =
"216,14,0,0,3,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,6,99,0,111,0,102,0,102,0,101,0,101,0,126,253,47,63,11,128,255,254,255,57,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,"
selection5 = selection5 &
"108,0,92,0,99,0,111,0,102,0,102,0,101,0,101,0,46,0,115,0,108,0,100,0,112,0,114,0,116,0,90,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,4,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,6,0,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,80,114,111,106,73,100,82,101,112,95,99,0,0,5,128,8,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,1,0,0,0,19,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,11,0,0,0,0,0,0,0,0,0"
selection5 = selection5 & ",Type=1"
StringtoArray selection5, var5
Set oselect5 =
Part. Extension . GetObjectByPersistReference3 ((var5), longstatus)
' Face 3
selection6 =
"216,14,0,0,3,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,6,99,0,111,0,102,0,102,0,101,0,101,0,126,253,47,63,11,128,255,254,255,57,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,"
selection6 = selection6 &
"108,0,92,0,99,0,111,0,102,0,102,0,101,0,101,0,46,0,115,0,108,0,100,0,112,0,114,0,116,0,90,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,4,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,2,0,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,80,114,111,106,73,100,82,101,112,95,99,0,0,5,128,8,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,1,0,0,0,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,11,0,0,0,19,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,0,0,0,0,0,0"
selection6 = selection6 & ",Type=1"
StringtoArray selection6, var6
Set oselect6 =
Part. Extension . GetObjectByPersistReference3 ((var6), longstatus)
' Face 4
selection7 =
"216,14,0,0,3,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,6,99,0,111,0,102,0,102,0,101,0,101,0,126,253,47,63,11,128,255,254,255,57,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,"
selection7 = selection7 &
"108,0,92,0,99,0,111,0,102,0,102,0,101,0,101,0,46,0,115,0,108,0,100,0,112,0,114,0,116,0,90,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,4,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,1,0,0,0,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,11,0,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,80,114,111,106,73,100,82,101,112,95,99,0,0,5,128,8,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,1,0,0,0,21,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,0,0,0,0,0,0"
selection7 = selection7 & ",Type=1"
StringtoArray selection7, var7
Set oselect7 =
Part. Extension . GetObjectByPersistReference3 ((var7), longstatus)
' Get the selections for convection
' Top-outer face of coffee pot
selection8 =
"216,14,0,0,3,0,0,0,255,254,255,21,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,11,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,9,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,174,232,47,63,11,128,255,254,255,60,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,"
selection8 = selection8 &
"114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,108,0,92,0,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,46,0,83,0,76,0,68,0,80,0,82,0,84,0,153,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,5,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,39,0,0,0,234,237,47,63,3,0,0,0,3,128,0,0,5,128,8,0,39,0,0,0,234,237,47,63,2,0,0,0,0,0,3,128,0,0,5,128,8,0,39,0,0,0,234,237,47,63,4,0,0,0,0,0,0,0,0,0"
selection8 = selection8 & ",Type=1"
StringtoArray selection8, var8
Set oselect8 =
Part. Extension . GetObjectByPersistReference3 ((var8), longstatus)
' Bottom-outer face of coffee pot
selection9 =
"216,14,0,0,3,0,0,0,255,254,255,21,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,11,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,9,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,174,232,47,63,11,128,255,254,255,60,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,"
selection9 = selection9 &
"0,114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,108,0,92,0,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,46,0,83,0,76,0,68,0,80,0,82,0,84,0,153,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,5,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,39,0,0,0,234,237,47,63,2,0,0,0,3,128,0,0,5,128,8,0,39,0,0,0,234,237,47,63,3,0,0,0,0,0,3,128,0,0,5,128,8,0,39,0,0,0,234,237,47,63,1,0,0,0,0,0,0,0,0,0"
selection9 = selection9 & ",Type=1"
StringtoArray selection9, var9
Set oselect9 =
Part. Extension . GetObjectByPersistReference3 ((var9), longstatus)
' Thermostat sensor located on the top of
the bottom face of the coffee pot
selection10 =
"216,14,0,0,3,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,99,0,111,0,102,0,102,0,101,0,101,0,106,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0,255,255,1,0,13,0,109,111,86,101,114,116,101,120,82,101,102,95,99,255,255,255,255,255,255,255,255,8,0,0,0,0,2,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,80,114,111,106,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,"
selection10 = selection10 &
"1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,6,99,0,111,0,102,0,102,0,101,0,101,0,126,253,47,63,11,128,255,254,255,57,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,114,0,115,0,92,0,84,0,104,"
selection10 = selection10 &
"0,101,0,114,0,109,0,97,0,108,0,92,0,99,0,111,0,102,0,102,0,101,0,101,0,46,0,115,0,108,0,100,0,112,0,114,0,116,0,90,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,4,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,92,0,0,0,241,69,48,63,255,255,1,0,18,0,109,111,80,76,105,110,101,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,90,0,0,0,219,69,48,63,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,"
selection10 = selection10 &
"110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,1,0,0,0,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,3,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,14,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,1,0,0,0,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,"
selection10 = selection10 &
"135,69,48,63,5,0,0,0,1,0,0,0,14,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,14,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,5,0,0,0,14,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,14,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,1,0,0,0,14,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,14,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,"
selection10 = selection10 &
"5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,2,0,0,0,14,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,14,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,6,0,0,0,4,0,0,0,0,0,0,0"
selection10 = selection10 & ",Type=1"
StringtoArray selection10, var10
Set therm1 =
Part. Extension . GetObjectByPersistReference3 ((var10), longstatus)
' Face to which to apply heat power and
flux
selection11 =
"216,14,0,0,3,0,0,0,255,254,255,21,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,11,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,3,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,9,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,174,232,47,63,11,128,255,254,255,60,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,"
selection11 = selection11 &
"114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,108,0,92,0,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,46,0,83,0,76,0,68,0,80,0,82,0,84,0,153,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,5,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,39,0,0,0,234,237,47,63,1,0,0,0,3,128,0,0,5,128,8,0,39,0,0,0,234,237,47,63,2,0,0,0,0,0,0,0,0,0"
selection11 = selection11 & ",Type=1"
StringtoArray selection11, var11
Set oselect10 =
Part. Extension . GetObjectByPersistReference3 ((var11), longstatus)
' Create arrays
varArray1 = Array(oselect1, oselect3)
varArray2 = Array(oselect2)
varArray3 = Array(oselect4, oselect5, oselect6, oselect7)
varArray4 = Array(oselect8, oselect9)
varArray5 = Array(therm1)
varArray6 = Array(oselect10)
' Set transient thermal study options
Debug.Print "Setting transient thermal study options..."
Set ThermalOptions = Study. ThermalStudyOptions
If ThermalOptions Is Nothing Then ErrorMsg swApp, "No
CWThermalStudyOptions object"
ThermalOptions. CheckFlowConvectionCoef2 = 0
ThermalOptions. SolutionType = 0
ThermalOptions. SolverType = 2
ThermalOptions. TotalTime = 3600
ThermalOptions. TimeIncrement = 60
' Apply initial temperature to coffee pot
and top
Debug.Print "Applying initial temperature to coffee pot and
top..."
Set LBCMgr = Study. LoadsAndRestraintsManager
If LBCMgr Is Nothing Then ErrorMsg swApp, "No
CWLoadsAndRestraintsManager object"
Set CWTemp =
LBCMgr. AddTemperature ((varArray1), errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to add
temperature load"
CWTemp. TemperatureBeginEdit
CWTemp. TemperatureType = 0
CWTemp. Unit = 1
CWTemp. TemperatureValue = 72
errCode = CWTemp. TemperatureEndEdit
If errCode <> 0 Then ErrorMsg swApp, "Failed to edit
temperature load"
' Apply initial temperature to coffee
Debug.Print "Applying initial temperature to coffee..."
Set CWTemp = LBCMgr. AddTemperature ((varArray2), errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to add
temperature load"
CWTemp. TemperatureBeginEdit
CWTemp. TemperatureType = 0
CWTemp. Unit = 1
CWTemp. TemperatureValue = 195
errCode = CWTemp. TemperatureEndEdit
If errCode <> 0 Then ErrorMsg swApp, "Failed to edit
temperature load"
' Create convection for coffee component
Debug.Print "Creating convection for coffee component..."
Set CWConv = LBCMgr. AddConvection ((varArray3), errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to add
convection load"
CWConv. ConvectionBeginEdit
CWConv. Unit = 1
CWConv. ConvectionCoefficient = 0.0000085
CWConv. BulkAmbientTemperature = 72
errCode = CWConv. ConvectionEndEdit
If errCode <> 0 Then ErrorMsg swApp, "Failed to edit
convection load"
' Create convection for coffee pot
component
Debug.Print "Creating convection for coffee pot component..."
Set CWConv = LBCMgr. AddConvection ((varArray4), errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to add
convection"
CWConv. ConvectionBeginEdit
CWConv. Unit = 1
CWConv. ConvectionCoefficient = 0.000061
CWConv. BulkAmbientTemperature = 72
errCode = CWConv. ConvectionEndEdit
If errCode <> 0 Then ErrorMsg swApp, "Failed to edit
convection load"
' Add heat power
Debug.Print "Creating heat power for face..."
Set CWHeatPower = LBCMgr. AddHeatPower ((varArray6), errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to add
heat power"
CWHeatPower. HeatPowerBeginEdit
CWHeatPower. Unit = 0
CWHeatPower. HPValue = 2000
CWHeatPower. ReverseDirection2 = 0
Debug.Print " Setting thermostat to a vertex..."
CWHeatPower. IncludeThermostat2 = True
CWHeatPower. SetThermostat therm1
CWHeatPower. ThermostatUnits = 1
CWHeatPower. SetCutOffTemperatures 190, 200
errCode = CWHeatPower. HeatPowerEndEdit
If errCode <> 0 Then ErrorMsg swApp, "Failed to edit
heat power"
' Add heat flux
Debug.Print "Creating heat flux for face..."
Set heatFlux = LBCMgr. AddHeatFlux ((varArray6), errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to add
heat flux"
heatFlux. HeatFluxBeginEdit
heatFlux. HFValue = 100
heatFlux. SetThermostat therm1
errCode = heatFlux. HeatFluxEndEdit
If errCode <> 0 Then ErrorMsg swApp, "Failed to edit
heat flux"
Dim dval1 As Double, dval2 As Double
heatFlux. GetCutOffTemperatures dval1, dval2
Debug.Print " Thermostat lower-bound
temperature: " & dval1
Debug.Print " Thermostat upper-bound temperature: " &
dval2
Debug.Print " Heat flux: " & heatFlux. HFValue
Debug.Print " Reverse direction? (False=no,
True=yes) " & heatFlux. ReverseDirection2
Debug.Print " Include thermostat effects? (False=no,
True=yes) " & heatFlux. IncludeThermostat2
Debug.Print " Thermostat units as defined in
swsTemperatureUnit_e: " & heatFlux. ThermostatUnits
Debug.Print " Use temperature curve? (False=no,
True=yes)
" & heatFlux. UseTemperatureCurve2
Debug.Print " Use time curve? (False=no,
True=yes) " &
heatFlux. UseTimeCurve2
Debug.Print " Units as defined in swsTemperatureUnit_e:
" & heatFlux. Unit
' Add radiation
Debug.Print "Creating radiation for face..."
Set rad = LBCMgr. AddRadiation (swsRadiationTypeSurfaceToAmbient,
(varArray6), errCode)
If errCode <> 0 Then ErrorMsg swApp, "Failed to add
radiation"
rad. RadiationBeginEdit
rad. Emmisivity = 0.5
rad. ViewFactor = 0.5
errCode = rad. RadiationEndEdit
Debug.Print " Ambient temperature: " & rad. AmbientTemperature
Debug.Print " Emissivity: " & rad. Emmisivity
Debug.Print " View factor: " & rad. ViewFactor
Debug.Print " Open system? (False = no, True =
yes) " & rad. OpenSystem2
Debug.Print " Use temperature curve? (False = no,
True = yes)
" & rad. UseTemperatureCurve2
Debug.Print " Use time curve? (False = no, True =
yes) " &
rad. UseTimeCurve2
Debug.Print " Units as defined in swsTemperatureUnit_e:
" & rad. Unit
If errCode <> 0 Then ErrorMsg swApp, "Failed to edit
radiation"
' Mesh model
Debug.Print "Creating mesh..."
Set CwMesh = Study. Mesh
If CwMesh Is Nothing Then ErrorMsg swApp, "No CWMesh object"
CwMesh. Quality = 1
Call CwMesh. GetDefaultElementSizeAndTolerance (0, el, tl)
errCode = Study. CreateMesh (0, el, tl)
If errCode <> 0 Then ErrorMsg swApp, "Failed to create
mesh"
' Analyze study
Debug.Print "Running analysis..."
errCode = Study. RunAnalysis
If errCode <> 0 Then ErrorMsg swApp, "Analysis failed
with error code as defined in swsRunAnalysisError_e: " & errCode
Debug.Print "Getting results..."
Set CWResult = Study. Results
If CWResult Is Nothing Then ErrorMsg swApp, "No CWResults
object"
' Get temperature at vertex
Debug.Print "Getting temperature at vertex..."
nStep = CWResult. GetMaximumAvailableSteps
Temp = CWResult. GetThermalForEntities (0, nStep, Nothing,
(varArray5), 2, errCode)
If errCode <> 0 Then ErrorMsg swApp, "No temperature
result"
' Save Thermal1 plot as eDrawing
CWResult. SavePlotsAseDrawings "c:\temp", "ThermalPlot",
"Thermal1"
' Export this study to the NASTRAN
finite-element analysis program
res = Study. ExportSimulationStudy ("c:\temp", "ThermalStudyExport",
swsNastranExportOption_LongFixed, 0, 0, swsStudyExportOption_Nastran,
swsNastranExportUnit_IPS)
' Rename the study
res = StudyMngr. RenameStudyFromName ("Thermal_One", "TransientThermal")
Debug.Print "Done!"
End Sub
Sub ErrorMsg(swApp As SldWorks.SldWorks, Message As String)
swApp. SendMsgToUser2 Message, 0, 0
swApp. RecordLine "'*** WARNING - General"
swApp. RecordLine "'*** " & Message
swApp. RecordLine ""
End Sub
Sub StringtoArray(inputSTR As String, varEntity As Variant)
Dim PID() As Byte
Dim i As Integer
varEntity = Split(inputSTR, ",")
ReDim PID(UBound(varEntity))
For i = 0 To (UBound(varEntity) - 1)
PID(i) = varEntity(i)
Next i
varEntity = PID
End Sub
```

## Get Thermal Values at Points (VB6)
```vb
Get Thermal Values Example (VBA)
SOLIDWORKS Simulation
API Help
Get Thermal Values Example (VBA)
This example shows how to get thermal values in a thermal study.
'---------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in (in SOLIDWORKS, click
' Tools > Add-ins > SOLIDWORKS Simulation > OK ).
' 2. Add the SOLIDWORKS Simulation type library as a reference (in the IDE,
' click Tools > References > SOLIDWORKS
Simulation version type library ).
' 3. Ensure that the specified model exists.
' 4. Modify the path to solidworks materials.sldmat if needed.
' 5. Uncomment the thermal method whose return values you want to see.
' 6. Open the Immediate window.
'
' Postconditions:
' 1. Opens document.
' 2. Creates Thermal_One study.
' 3. Applies user-defined and SOLIDWORKS materials to components.
' 4. Gets objects by persistent reference and populates input arrays.
' 5. Sets thermal study options.
' 6. Assigns initial temperatures to assembly components.
' 7. Adds heat power and convection to assembly components.
' 8. Assigns thermostat location and the temperature range.
' 9. Meshes the assembly.
' 10. Analyzes the study.
' 11. Prints thermal values at solution step 60 to the Immediate window.
'
' NOTES:
' * To get persistent reference identifiers (PIDs) for model selections,
' use
pidcollector.exe
or IModelDocExtension::GetPersistReference3.
' * Solving this study can take some time. Examine the Immediate window
' to monitor the macro's progress. Done! indicates that the
macro
' has finished.
' * Because this assembly document is used by a SOLIDWORKS Simulation
' online tutorial, do not save any changes.
'--------------------------------------------------------------------
Option Explicit
Sub main()
Dim swApp As SldWorks.SldWorks
Dim COSMOSWORKS As CosmosWorksLib.COSMOSWORKS
Dim CWObject As CosmosWorksLib.CwAddincallback
Dim ActDoc As CosmosWorksLib.CWModelDoc
Dim StudyMngr As CosmosWorksLib.CWStudyManager
Dim Study As CosmosWorksLib.CWStudy
Dim SolidMgr As CosmosWorksLib.CWSolidManager
Dim SolidComp As CosmosWorksLib.CWSolidComponent
Dim SolidBody As CosmosWorksLib.CWSolidBody
Dim ThermalOptions As CosmosWorksLib.CWThermalStudyOptions
Dim CWConv As CosmosWorksLib.CWConvection
Dim CWHeatPower As CosmosWorksLib.CWHeatPower
Dim CwMesh As CosmosWorksLib.CwMesh
Dim CWResult As CosmosWorksLib.CWResults
Dim CWMat As CosmosWorksLib.CWMaterial
Dim CWTemp As CosmosWorksLib.CWTemperature
Dim Part As SldWorks.ModelDoc2
Dim LBCMgr As CosmosWorksLib.CWLoadsAndRestraintsManager
Dim oselect1 As Object, oselect2 As Object, oselect3 As Object
Dim oselect4 As Object, oselect5 As Object, oselect6 As Object
Dim oselect7 As Object, oselect8 As Object, oselect9 As Object,
oselect10 As Object
Dim therm1 As Object
Dim var1 As Variant, var2 As Variant, var3 As Variant, var4 As
Variant
Dim var5 As Variant, var6 As Variant, var7 As Variant, var8 As
Variant
Dim var9 As Variant, var10 As Variant, var11 As Variant
Dim Temp As Variant
Dim varArray1 As Variant
Dim varArray2 As Variant
Dim varArray3 As Variant
Dim varArray4 As Variant
Dim varArray5 As Variant
Dim varArray6 As Variant
Dim varArray7 As Variant
Dim bApp As Boolean
Dim selection1 As String, selection2 As String, selection3 As
String
Dim selection4 As String, selection5 As String, selection6 As
String
Dim selection7 As String, selection8 As String, selection9 As
String
Dim selection10 As String, selection11 As String
Dim longstatus As Long, longwarnings As Long
Dim errCode As Long
Dim el As Double, tl As Double
Dim nStep As Long
' Connect to SOLIDWORKS
If swApp Is Nothing Then Set swApp = Application.SldWorks
' Open document
Debug.Print "Opening document..."
Set Part = swApp. OpenDoc6 ("C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS 2019\samples\Simulation Examples\Thermal\CoffeeJar.SLDASM", swDocASSEMBLY,
swOpenDocOptions_Silent, "", longstatus, longwarnings)
If Part Is Nothing Then ErrorMsg swApp,
"Failed to open document", True
' Get the SOLIDWORKS Simulation add-in
object
Set CWObject = swApp. GetAddInObject ("SldWorks.Simulation")
If CWObject Is Nothing Then ErrorMsg swApp, "No
CwAddincallback object", True
Set COSMOSWORKS = CWObject. COSMOSWORKS
If COSMOSWORKS Is Nothing Then ErrorMsg swApp, "No
CosmosWorks object", True
' Get active document
Set ActDoc = COSMOSWORKS. ActiveDoc ()
If ActDoc Is Nothing Then ErrorMsg swApp, "No active
document", True
' Create new thermal study
Debug.Print "Creating thermal study..."
Set StudyMngr = ActDoc. StudyManager ()
If StudyMngr Is Nothing Then ErrorMsg swApp, "No study
manager object", True
Set Study = StudyMngr. CreateNewStudy3 ("Thermal_One",
swsAnalysisStudyTypeThermal, swsMeshTypeSolid, errCode)
If Study Is Nothing Then ErrorMsg swApp, "No study", True
' Get first solid component
Set SolidMgr = Study. SolidManager
If SolidMgr Is Nothing Then ErrorMsg swApp, "No solid manager
object", True
Set SolidComp = SolidMgr. GetComponentAt (0,
errCode)
If errCode <> 0 Then ErrorMsg swApp, "No solid component",
True
Set SolidBody = SolidComp. GetSolidBodyAt (0,
errCode)
If errCode <> 0 Then ErrorMsg swApp, "No solid body", True
' Add material
Set CWMat = SolidBody. GetDefaultMaterial
If CWMat Is Nothing Then ErrorMsg swApp, "No default
material", True
CWMat. MaterialName = "Coffee"
Call CWMat. SetPropertyByName ("DENS",
1000, 0)
Call CWMat. SetPropertyByName ("KX", 40, 0)
Call CWMat. SetPropertyByName ("C", 4200#, 0)
errCode = SolidBody. SetSolidBodyMaterial (CWMat)
If errCode <> 0 Then ErrorMsg swApp, "Solid body material not
set", True
' Get second solid component
Set SolidComp = SolidMgr. GetComponentAt (1, errCode)
If errCode <> 0 Then ErrorMsg swApp, "No solid component",
True
Set SolidBody = SolidComp. GetSolidBodyAt (0,
errCode)
If errCode <> 0 Then ErrorMsg swApp, "No solid body", True
' Add material
bApp = SolidBody. SetLibraryMaterial ("C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\sldmaterials\solidworks materials.sldmat", "Glass")
If bApp = False Then ErrorMsg swApp, "No glass material
applied", True
' Get third solid component
Set SolidComp = SolidMgr. GetComponentAt (2, errCode)
If errCode <> 0 Then ErrorMsg swApp, "No solid component",
True
Set SolidBody = SolidComp. GetSolidBodyAt (0,
errCode)
If errCode <> 0 Then ErrorMsg swApp, "No solid body", True
' Add material
bApp = SolidBody. SetLibraryMaterial ("C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\sldmaterials\solidworks materials.sldmat", "Nylon
6/10")
If bApp = False Then ErrorMsg swApp, " No Nylon 6/10 material
applied", True
' CoffeePot
selection1 =
"216,14,0,0,5,0,0,0,255,254,255,21,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,11,0,0,0"
selection1 = selection1 & ",Type=1"
StringtoArray selection1, var1
Set oselect1 = Part. Extension . GetObjectByPersistReference3 ((var1),
longstatus)
' Coffee
selection2 =
"216,14,0,0,5,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0"
selection2 = selection2 & ",Type=1"
StringtoArray selection2, var2
Set oselect2 = Part. Extension . GetObjectByPersistReference3 ((var2),
longstatus)
' Top
selection3 =
"216,14,0,0,5,0,0,0,255,254,255,15,84,0,111,0,112,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,13,0,0,0"
selection3 = selection3 & ",Type=1"
StringtoArray selection3, var3
Set oselect3 = Part. Extension . GetObjectByPersistReference3 ((var3),
longstatus)
' Get the four coffee faces that make up the
surface of coffee in the coffee pot
' Face 1
selection4 =
"216,14,0,0,3,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,6,99,0,111,0,102,0,102,0,101,0,101,0,126,253,47,63,11,128,255,254,255,57,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,10"
selection4 = selection4 &
"8,0,92,0,99,0,111,0,102,0,102,0,101,0,101,0,46,0,115,0,108,0,100,0,112,0,114,0,116,0,90,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,4,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,5,0,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,80,114,111,106,73,100,82,101,112,95,99,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,19,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,1,0,0,0,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,11,0,0,0,0,0,0,0,0,0"
selection4 = selection4 & ",Type=1"
StringtoArray selection4, var4
Set oselect4 = Part. Extension . GetObjectByPersistReference3 ((var4),
longstatus)
' Face 2
selection5 =
"216,14,0,0,3,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,6,99,0,111,0,102,0,102,0,101,0,101,0,126,253,47,63,11,128,255,254,255,57,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,"
selection5 = selection5 &
"108,0,92,0,99,0,111,0,102,0,102,0,101,0,101,0,46,0,115,0,108,0,100,0,112,0,114,0,116,0,90,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,4,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,6,0,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,80,114,111,106,73,100,82,101,112,95,99,0,0,5,128,8,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,1,0,0,0,19,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,11,0,0,0,0,0,0,0,0,0"
selection5 = selection5 & ",Type=1"
StringtoArray selection5, var5
Set oselect5 = Part. Extension . GetObjectByPersistReference3 ((var5),
longstatus)
' Face 3
selection6 =
"216,14,0,0,3,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,6,99,0,111,0,102,0,102,0,101,0,101,0,126,253,47,63,11,128,255,254,255,57,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,"
selection6 = selection6 &
"108,0,92,0,99,0,111,0,102,0,102,0,101,0,101,0,46,0,115,0,108,0,100,0,112,0,114,0,116,0,90,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,4,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,2,0,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,80,114,111,106,73,100,82,101,112,95,99,0,0,5,128,8,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,1,0,0,0,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,11,0,0,0,19,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,0,0,0,0,0,0"
selection6 = selection6 & ",Type=1"
StringtoArray selection6, var6
Set oselect6 = Part. Extension . GetObjectByPersistReference3 ((var6),
longstatus)
' Face 4
selection7 =
"216,14,0,0,3,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,6,99,0,111,0,102,0,102,0,101,0,101,0,126,253,47,63,11,128,255,254,255,57,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,"
selection7 = selection7 &
"108,0,92,0,99,0,111,0,102,0,102,0,101,0,101,0,46,0,115,0,108,0,100,0,112,0,114,0,116,0,90,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,4,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,1,0,0,0,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,11,0,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,80,114,111,106,73,100,82,101,112,95,99,0,0,5,128,8,0,92,0,0,0,241,69,48,63,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,1,0,0,0,21,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,16,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,0,0,0,0,0,0"
selection7 = selection7 & ",Type=1"
StringtoArray selection7, var7
Set oselect7 = Part. Extension . GetObjectByPersistReference3 ((var7),
longstatus)
' Get the selections for convection
' Top-outer face of coffee pot
selection8 =
"216,14,0,0,3,0,0,0,255,254,255,21,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,11,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,9,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,174,232,47,63,11,128,255,254,255,60,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,"
selection8 = selection8 &
"114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,108,0,92,0,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,46,0,83,0,76,0,68,0,80,0,82,0,84,0,153,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,5,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,39,0,0,0,234,237,47,63,3,0,0,0,3,128,0,0,5,128,8,0,39,0,0,0,234,237,47,63,2,0,0,0,0,0,3,128,0,0,5,128,8,0,39,0,0,0,234,237,47,63,4,0,0,0,0,0,0,0,0,0"
selection8 = selection8 & ",Type=1"
StringtoArray selection8, var8
Set oselect8 = Part. Extension . GetObjectByPersistReference3 ((var8),
longstatus)
' Lower-outer face of coffee pot
selection9 =
"216,14,0,0,3,0,0,0,255,254,255,21,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,11,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,5,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,9,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,174,232,47,63,11,128,255,254,255,60,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,"
selection9 = selection9 &
"0,114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,108,0,92,0,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,46,0,83,0,76,0,68,0,80,0,82,0,84,0,153,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,5,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,39,0,0,0,234,237,47,63,2,0,0,0,3,128,0,0,5,128,8,0,39,0,0,0,234,237,47,63,3,0,0,0,0,0,3,128,0,0,5,128,8,0,39,0,0,0,234,237,47,63,1,0,0,0,0,0,0,0,0,0"
selection9 = selection9 & ",Type=1"
StringtoArray selection9, var9
Set oselect9 = Part. Extension . GetObjectByPersistReference3 ((var9),
longstatus)
' Get the selections for heat power
' Vertex located on the top of the bottom face of the coffee pot
selection10 =
"216,14,0,0,3,0,0,0,255,254,255,18,67,0,111,0,102,0,102,0,101,0,101,0,45,0,49,0,64,0,99,0,111,0,102,0,102,0,101,0,101,0,106,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,12,0,0,0,255,255,1,0,13,0,109,111,86,101,114,116,101,120,82,101,102,95,99,255,255,255,255,255,255,255,255,8,0,0,0,0,2,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,18,0,109,111,80,76,105,110,101,80,114,111,106,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,"
selection10 = selection10 &
"1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,6,99,0,111,0,102,0,102,0,101,0,101,0,126,253,47,63,11,128,255,254,255,57,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,114,0,115,0,92,0,84,0,104,"
selection10 = selection10 &
"0,101,0,114,0,109,0,97,0,108,0,92,0,99,0,111,0,102,0,102,0,101,0,101,0,46,0,115,0,108,0,100,0,112,0,114,0,116,0,90,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,4,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,92,0,0,0,241,69,48,63,255,255,1,0,18,0,109,111,80,76,105,110,101,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,90,0,0,0,219,69,48,63,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,"
selection10 = selection10 &
"110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,1,0,0,0,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,3,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,14,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,1,0,0,0,3,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,"
selection10 = selection10 &
"135,69,48,63,5,0,0,0,1,0,0,0,14,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,14,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,5,0,0,0,14,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,14,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,1,0,0,0,14,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,14,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,"
selection10 = selection10 &
"5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,2,0,0,0,2,0,0,0,14,128,0,0,5,128,8,0,92,0,0,0,241,69,48,63,14,128,0,0,5,128,8,0,90,0,0,0,219,69,48,63,17,128,0,0,5,128,8,0,88,0,0,0,135,69,48,63,5,0,0,0,1,0,0,0,6,0,0,0,4,0,0,0,0,0,0,0"
selection10 = selection10 & ",Type=1"
StringtoArray selection10, var10
Set therm1 = Part. Extension . GetObjectByPersistReference3 ((var10),
longstatus)
' Vertex where sensor (thermostat) is
located
selection11 =
"216,14,0,0,3,0,0,0,255,254,255,21,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,45,0,49,0,64,0,67,0,111,0,102,0,102,0,101,0,101,0,74,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,11,0,0,0,255,255,1,0,11,0,109,111,70,97,99,101,82,101,102,95,99,1,0,0,0,0,0,0,0,3,0,0,0,0,3,0,0,0,0,0,0,125,195,148,37,173,73,178,84,125,195,148,37,173,73,178,84,0,0,255,255,1,0,23,0,109,111,70,114,111,109,83,107,116,69,110,116,83,117,114,102,73,100,82,101,112,95,99,0,0,255,255,1,0,6,0,109,111,70,82,95,99,255,255,1,0,13,0,109,111,69,120,116,79,98,106,101,99,116,95,99,255,255,1,0,14,0,109,111,79,98,106,70,105,108,101,68,101,102,95,99,255,255,1,0,17,0,109,111,67,83,116,114,105,110,103,72,97,110,100,108,101,95,99,255,254,255,9,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,174,232,47,63,11,128,255,254,255,60,67,0,58,0,92,0,67,0,79,0,83,0,77,0,79,0,83,0,68,0,111,0,99,0,115,0,92,0,69,0,120,0,97,0,109,0,112,0,108,0,101,0,115,0,95,0,102,0,111,0,114,0,67,0,117,0,115,0,116,0,111,0,109,0,101,0,"
selection11 = selection11 &
"114,0,115,0,92,0,84,0,104,0,101,0,114,0,109,0,97,0,108,0,92,0,67,0,111,0,102,0,102,0,101,0,101,0,80,0,111,0,116,0,46,0,83,0,76,0,68,0,80,0,82,0,84,0,153,14,28,65,0,0,0,0,2,0,1,0,0,0,0,0,0,0,5,0,0,0,255,254,255,7,68,0,101,0,102,0,97,0,117,0,108,0,116,0,0,0,0,0,0,0,0,0,0,0,0,0,39,0,0,0,234,237,47,63,1,0,0,0,3,128,0,0,5,128,8,0,39,0,0,0,234,237,47,63,2,0,0,0,0,0,0,0,0,0"
selection11 = selection11 & ",Type=1"
StringtoArray selection11, var11
Set oselect10 = Part. Extension . GetObjectByPersistReference3 ((var11),
longstatus)
' Create arrays
varArray1 = Array(oselect1, oselect3)
varArray2 = Array(oselect2)
varArray3 = Array(oselect4, oselect5, oselect6, oselect7)
varArray4 = Array(oselect8, oselect9)
varArray5 = Array(therm1)
varArray6 = Array(oselect10)
varArray7 = Array(-19.63, 31.23, 173.15)
' Set transient thermal study with solver
type FFEPlus
Debug.Print "Setting transient thermal study with solver type
FFEPlus..."
Set ThermalOptions = Study. ThermalStudyOptions
Study. ThermalStudyOptions . SolutionType = 0
If ThermalOptions Is Nothing Then ErrorMsg swApp, "No thermal
options", False
ThermalOptions. SolverType = 2
ThermalOptions. TotalTime = 3600
ThermalOptions. TimeIncrement = 60
' Apply initial temperature to coffee pot
and top
Debug.Print "Applying initial temperature to coffee pot and
top..."
Set LBCMgr = Study. LoadsAndRestraintsManager
If LBCMgr Is Nothing Then ErrorMsg swApp, "No loads and
restraints manager", False
Set CWTemp = LBCMgr. AddTemperature ((varArray1),
errCode)
If errCode <> 0 Then ErrorMsg swApp, "No temperature
applied", True
CWTemp. TemperatureBeginEdit
CWTemp. TemperatureType = 0
CWTemp. Unit = 1
CWTemp. TemperatureValue = 72
errCode = CWTemp. TemperatureEndEdit
If errCode <> 0 Then ErrorMsg swApp, "No temperature
applied", True
' Apply initial temperature to coffee
Debug.Print "Applying initial temperature to coffee..."
Set CWTemp = LBCMgr. AddTemperature ((varArray2),
errCode)
If errCode <> 0 Then ErrorMsg swApp, "No temperature
applied", True
CCWTemp. TemperatureBeginEdit
CWTemp. TemperatureType = 0
CWTemp. Unit = 1
CWTemp. TemperatureValue = 195
errCode = CWTemp. TemperatureEndEdit
If errCode <> 0 Then ErrorMsg swApp, "No temperature
applied", True
' Create convection for coffee component
Debug.Print "Creating convection for coffee component..."
Set CWConv = LBCMgr. AddConvection ((varArray3),
errCode)
If errCode <> 0 Then ErrorMsg swApp, "No convection added",
True
CWConv. ConvectionBeginEdit
CWConv. Unit = 1
CWConv. ConvectionCoefficient = 0.0000085
CWConv. BulkAmbientTemperature = 72
errCode = CWConv. ConvectionEndEdit
If errCode <> 0 Then ErrorMsg swApp, "Convection end-edit
failed for coffee", True
' Create convection for coffee pot
component
Debug.Print "Creating convection for coffee pot component..."
Set CWConv = LBCMgr. AddConvection ((varArray4), errCode)
If errCode <> 0 Then ErrorMsg swApp, "No convection added",
True
CWConv. ConvectionBeginEdit
CWConv. Unit = 1
CWConv. ConvectionCoefficient = 0.000061
CWConv. BulkAmbientTemperature = 72
errCode = CWConv. ConvectionEndEdit
If errCode <> 0 Then ErrorMsg swApp, "Convection end-edit
failed for coffee pot", True
' Create heat power for face
Debug.Print "Creating heat power for face..."
Set CWHeatPower = LBCMgr. AddHeatPower ((varArray6),
errCode)
If errCode <> 0 Then ErrorMsg swApp, "No heat power created",
True
CWHeatPower. HeatPowerBeginEdit
CWHeatPower. Unit = 0
CWHeatPower. HPValue = 2000
' Set thermostat to a vertex
Debug.Print "Setting thermostat to a vertex..."
CWHeatPower. IncludeThermostat = True
CWHeatPower. SetThermostat therm1
CWHeatPower. ThermostatUnits = 1
CWHeatPower. SetCutOffTemperatures 190, 200
errCode = CWHeatPower. HeatPowerEndEdit
If errCode <> 0 Then ErrorMsg swApp, "Heat power not
updated", True
' Create mesh
Debug.Print "Creating mesh..."
Set CwMesh = Study. Mesh
If CwMesh Is Nothing Then ErrorMsg swApp, "No mesh object",
False
CwMesh. Quality = 1
CwMesh. GetDefaultElementSizeAndTolerance 0, el, tl
errCode = Study. CreateMesh (0, el, tl)
If errCode <> 0 Then ErrorMsg swApp, "Mesh failed", True
' Run analysis
Debug.Print "Running analysis..."
errCode = Study. RunAnalysis
If errCode <> 0 Then ErrorMsg swApp, "Analysis failed", True
' Get thermal results
Debug.Print "Getting thermal results..."
Set CWResult = Study. Results
If CWResult Is Nothing Then ErrorMsg swApp, "No results
object", False
nStep = CWResult. GetMaximumAvailableSteps
' solution step 60
' Get temperature value at point {-19.63,
31.23, 173.15} at solution step 60
Temp = CWResult. GetThermalValuesAtPoints (swsThermalComponentTEMP,
nStep, Nothing, (varArray7), swsTemperatureUnitKelvin, errCode)
' Get temperature values at all nodes at
solution step 60
'Temp = CWResult. GetThermalValues (nStep, Nothing,
swsTemperatureUnitKelvin, errCode)
' Get heat energy values for the specified
vertex at solution step 60
'Temp = CWResult. GetHeatPowerOrEnergy2 (-1, (varArray6),
swsUnitSI, 1, nStep, errCode) ' Heat energy
' Get temperature values for all steps at
node 9236
'Temp = CWResult. GetThermalComponentForAllStepsAtNode (swsThermalComponentTEMP,
9236, Nothing, swsTemperatureUnitKelvin, errCode)
' Get temperature values for the specified
entities at solution step 60
'Temp = CWResult. GetThermalForEntities (swsThermalComponentTEMP,
nStep, Nothing, (varArray4), swsTemperatureUnitKelvin, errCode)
If errCode <> 0 Then ErrorMsg swApp, "No temperature result",
True
' Print out the contents of Temp
' See the Simulation API Help for information about the
contents of the returned array
Dim i As Long
For i = 0 To UBound(Temp)
Debug.Print Temp(i)
Next
Debug.Print "Done!"
End Sub
Sub ErrorMsg(swApp As SldWorks.SldWorks, Message As String,
EndTest As Boolean)
swApp. SendMsgToUser2 Message, 0, 0
swApp. RecordLine "'*** WARNING - General"
swApp. RecordLine "'*** " & Message
swApp. RecordLine ""
End Sub
Function StringtoArray(inputSTR As String, varEntity As
Variant)
Dim PID() As Byte
Dim i As Integer
varEntity = Split(inputSTR, ",")
ReDim PID(UBound(varEntity))
For i = 0 To (UBound(varEntity) - 1)
PID(i) = varEntity(i)
Next i
varEntity = PID
End Function
```

## Apply Material to Bodies (VB6)
```vb
Apply Material to Bodies Example (VBA)
SOLIDWORKS Simulation
API Help
Apply Material to Bodies Example (VBA)
This example shows how to apply user-defined and SOLIDWORKS-defined
material to different bodies.
'----------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in (in SOLIDWORKS, click
' Tools > Add-ins > SOLIDWORKS Simulation > OK ).
' 2. Add the SOLIDWORKS Simulation type library as a reference (in the IDE,
' click Tools > References > SOLIDWORKS
Simulation version type library ).
' 3. Verify that the specified material library exists.
' 4. Verify that the specified assembly exists.
'
' Postconditions:
' 1. Opens the assembly.
' 2. Creates a nonlinear study.
' 3. Gets the number of components.
' 4. Applies materials to all components.
' 5. Expand the Parts folder in the Simulation Study tree
' }}--> }}-->to
verify step 4.
'
' NOTE: Because this assembly is used elsewhere, do not save changes.
'----------------------------------------------------------------------------
Option Explicit
Sub main()
}}--> }}-->Dim
SwApp As SldWorks.SldWorks
}}--> }}-->Dim
COSMOSWORKS As Object
}}--> }}-->Dim
COSMOSObject As CosmosWorksLib.CwAddincallback
}}--> }}-->Dim
ActDoc As CosmosWorksLib.CWModelDoc
}}--> }}-->Dim
StudyMngr As CosmosWorksLib.CWStudyManager
}}--> }}-->Dim
Study As CosmosWorksLib.CWStudy
}}--> }}-->Dim
SolidMgr As CosmosWorksLib.CWSolidManager
}}--> }}-->Dim
SolidComponent As CosmosWorksLib.CWSolidComponent
}}--> }}-->Dim
SolidBody As CosmosWorksLib.CWSolidBody
}}--> }}-->Dim
errorCode As Long
}}--> }}-->Dim
longstatus As Long
}}--> }}-->Dim
longwarnings As Long
}}--> }}-->Dim
errCode As Long
}}--> }}-->Dim
CompCount As Integer
}}--> }}-->Dim
j As Integer
}}--> }}-->Dim
CWMat As CosmosWorksLib.CWMaterial
}}--> }}-->Dim
returnValue As Long
}}--> }}-->Dim
SName As String
}}--> }}-->
}}--> }}-->'
Connect to SOLIDWORKS
}}--> }}-->If
SwApp Is Nothing Then Set SwApp = Application.SldWorks
}}--> }}-->
}}--> }}-->'
Get SOLIDWORKS Simulation object>
}}--> }}-->Set
COSMOSObject = SwApp. GetAddInObject ("SldWorks.Simulation")
}}--> }}-->If
COSMOSObject Is Nothing Then ErrorMsg SwApp, "No CwAddincallback object"
}}--> }}-->Set
COSMOSWORKS = COSMOSObject. COSMOSWORKS
}}--> }}-->If
COSMOSWORKS Is Nothing Then ErrorMsg SwApp, "No COSMOSWORKS object"
}}--> }}-->'
Open and get document
}}--> }}-->SwApp. OpenDoc6 "C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS 2019\samples\Simulation Examples\basketball_hoop.sldasm", swDocASSEMBLY,
swOpenDocOptions_Silent, "", longstatus, longwarnings
}}--> }}-->Set
ActDoc = COSMOSWORKS. ActiveDoc ()
}}--> }}-->If
ActDoc Is Nothing Then ErrorMsg SwApp, "No active document"
}}--> }}-->'
Create new nonlinear study
}}--> }}-->Set
StudyMngr = ActDoc. StudyManager ()
}}--> }}-->If
StudyMngr Is Nothing Then ErrorMsg SwApp, "No CWStudyManager object"
}}--> }}-->Set
Study = StudyMngr. CreateNewStudy ("Nonlinear",
swsAnalysisStudyTypeNonlinear, swsMeshTypeSolid, errCode)
}}--> }}-->If
Study Is Nothing Then ErrorMsg SwApp, "No CWStudy object"
}}--> }}-->'
Get number of solid components
}}--> }}-->Set
SolidMgr = Study. SolidManager
}}--> }}-->If
SolidMgr Is Nothing Then ErrorMsg SwApp, "No CWSolidManager object"
}}--> }}-->CompCount
= SolidMgr. ComponentCount
}}--> }}-->Set
SolidComponent = SolidMgr. GetComponentAt (0,
errorCode)
}}--> }}-->If
SolidComponent Is Nothing Then ErrorMsg SwApp, "No CWSolidComponent
object"
}}--> }}-->
}}--> }}-->'
Get name of solid component
}}--> }}-->SName
= SolidComponent. ComponentName
}}--> }}-->'
Apply user-defined material for the first component in the tree
}}--> }}-->Set
SolidBody = SolidComponent. GetSolidBodyAt (0,
errCode)
}}--> }}-->If
errCode <> 0 Then ErrorMsg SwApp, "No solid body"
}}--> }}-->If
SolidBody Is Nothing Then ErrorMsg SwApp, "No CWSolidBody object"
}}--> }}-->Set
CWMat = SolidBody. GetDefaultMaterial
}}--> }}-->CWMat. MaterialUnits = 0
}}--> }}-->If
CWMat Is Nothing Then ErrorMsg SwApp, "No CWMaterial object"
}}--> }}-->CWMat. MaterialName = "Alloy Steel"
}}--> }}-->CWMat. SetPropertyByName "EX", 210000000000#,
0
}}--> }}-->Debug.Print
"Property EX: " & CWMat. GetPropertyByName (swsUnitSystemSI,
"EX", 0)
}}--> }}-->CWMat. SetPropertyByName "NUXY",
0.28, 0
}}--> }}-->CWMat. SetPropertyByName "GXY", 79000000000#,
0
}}--> }}-->CWMat. SetPropertyByName "DENS",
7700, 0
}}--> }}-->CWMat. SetPropertyByName "SIGXT",
723825600, 0
}}--> }}-->CWMat. SetPropertyByName "SIGYLD",
620422000, 0
}}--> }}-->CWMat. SetPropertyByName "ALPX",
0.000013, 0
}}--> }}-->CWMat. SetPropertyByName "KX", 50,
0
}}--> }}-->CWMat. SetPropertyByName "C", 460,
0
}}--> }}-->errCode
= SolidBody. SetSolidBodyMaterial (CWMat)
}}--> }}-->If
errCode <> 0 Then ErrorMsg SwApp, "Failed to apply material"
}}--> }}-->'
Apply a different SOLIDWORKS Simulation library material to other components
}}--> }}-->Set
SolidBody = Nothing
}}--> }}-->Set
SolidComponent = Nothing
}}--> }}-->For
j = 1 To (CompCount - 1)
}}--> }}-->Set
SolidComponent = SolidMgr. GetComponentAt (j,
errorCode)
}}--> }}-->SName
= SolidComponent. ComponentName
}}--> }}-->Set
SolidBody = SolidComponent. GetSolidBodyAt (0,
errCode)
}}--> }}-->If
errCode <> 0 Then ErrorMsg SwApp, "No CWSolidBody object"
}}--> }}-->returnValue
= SolidBody. SetLibraryMaterial ("c:\Program
Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\sldmaterials\solidworks
materials.sldmat", "2024 Alloy (SN)")
}}--> }}-->If
returnValue = 0 Then ErrorMsg SwApp, "No material applied"
}}--> }}-->Set
SolidBody = Nothing
}}--> }}-->Next
j
End Sub
' Error function
Function ErrorMsg(SwApp As Object, Message As String)
}}--> }}-->SwApp. SendMsgToUser2 Message, 0, 0
}}--> }}-->SwApp. RecordLine "'*** WARNING - General"
}}--> }}-->SwApp. RecordLine "'*** " & Message
}}--> }}-->SwApp. RecordLine ""
}}-->
End Function
```

## Apply Mesh Control (VB6)
```vb
'----------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in (in SOLIDWORKS, click
'    Tools > Add-ins > SOLIDWORKS Simulation > OK).
' 2. Add the SOLIDWORKS Simulation type library as a reference (in the IDE, 
'    click Tools > References > SOLIDWORKS Simulation version type library).
' 3. Ensure that the specified part and material library exist.
' 4. Open the Immediate window.
'
' Postconditions:
' 1. Creates a static study named frame.
' 2. Gets beam information.
' 3. Applies Plain Carbon Steel material to all beams.
' 4. Calculates joints for all beams, gets a neutral axis for each beam,  
'    and gets the pinball tolerance value and unit.
' 5. Sets the mesher type to alternate curvature based.
' 6. Applies a mesh control and gets its values.
' 7. Creates a mixed mesh and gets its type and state.
'    NOTE: This can take several seconds to complete.
' 8. Inspect the Immediate window.
'
' NOTES:
' *  Beam elements are created by default for parts with structural members.
' *  Because the part is used elsewhere, do not save changes.
'----------------------------------------------------------------------------
```

## Add Component Contacts (VB6)
```vb
Add Component Contacts Example (VBA)
SOLIDWORKS Simulation API
Help
Add Component Contacts Example (VBA)
This example shows how to add component contacts to a frequency study.
'----------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in (in SOLIDWORKS, click
' Tools > Add-ins > SOLIDWORKS Simulation > OK ).
' 2. Add the SOLIDWORKS Simulation type library as a reference (in the IDE,
' click Tools > References > SOLIDWORKS
Simulation version type library ).
' 3. Ensure that the specified material library exists.
' 4. Ensure that the specified model document exists.
' 5. Open the Immediate window.
'
' Postconditions:
' 1. Opens the model.
' 2. Creates Frequency 1 study.
' 3. Uses a soft spring to stabilize the model.
' 4. Adds Component Contact-1 and Component Contact-2 to the
' Component Contacts folder in the Simulation
study tree.
' 5. Prints the component contacts' properties to the Immediate window.
' 6. Applies ductile iron to the model.
' 7. Meshes the model.
' 8. Analyzes Frequency 1 .
' 9. Prints the resonant frequencies of each mode to the Immediate window.
' 10. Inspect the Immediate window, the Simulation study tree, and the
' graphics area.
'
' NOTE: Because the model is used elsewhere, do not save changes.
'----------------------------------------------------------------------------
Option Explicit
Sub main()
Dim SwApp As SldWorks.SldWorks
Dim COSMOSWORKS As Object
Dim COSMOSObject As CosmosWorksLib.CwAddincallback
Dim ActDoc As CosmosWorksLib.CWModelDoc
Dim StudyMngr As CosmosWorksLib.CWStudyManager
Dim Study As CosmosWorksLib.CWStudy
Dim SolidMgr As CosmosWorksLib.CWSolidManager
Dim SolidComponent As CosmosWorksLib.CWSolidComponent
Dim SolidBody As CosmosWorksLib.CWSolidBody
Dim CwMesh As CosmosWorksLib.CwMesh
Dim CWResult As CosmosWorksLib.CWResults
Dim Part As SldWorks.ModelDoc2
Dim ContactMgr As CosmosWorksLib.CWContactManager
Dim CWComponentContact As CosmosWorksLib.CWContactComponent
Dim pDisp3 As Object, pDisp4 As Object
Dim var20 As Variant, var21 As Variant
Dim Freq As Variant
Dim bApp As Boolean
Dim str3 As String, str4 As String
Dim longstatus As Long, longwarnings As Long
Dim errCode As Long
Dim el As Double, tl As Double
Dim i As Long
Set SwApp = Application.SldWorks
SwApp. OpenDoc6 "C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS 2019\samples\Simulation Examples\Contact\quartereyebar.sldasm",
swDocASSEMBLY, swOpenDocOptions_Silent, "", longstatus, longwarnings
Set Part = SwApp. ActiveDoc ()
Set COSMOSObject = SwApp. GetAddInObject ("SldWorks.Simulation")
If COSMOSObject Is Nothing Then ErrorMsg SwApp, "Failed to get
Simulation add-in"
Set COSMOSWORKS = COSMOSObject. COSMOSWORKS
If COSMOSWORKS Is Nothing Then ErrorMsg SwApp, "Failed to get
COSMOSWORKS"
Set ActDoc = COSMOSWORKS. ActiveDoc ()
If ActDoc Is Nothing Then ErrorMsg SwApp, "Failed to get active
document"
' Create new frequency study
Set StudyMngr = ActDoc. StudyManager ()
If StudyMngr Is Nothing Then ErrorMsg SwApp, "Failed to get study
manager"
Set Study = StudyMngr. CreateNewStudy3 ("Frequency
1", swsAnalysisStudyTypeFrequency, 0, errCode)
If Study Is Nothing Then ErrorMsg SwApp, "Failed to create
frequency study"
Dim StudyOptions As ICWFrequencyStudyOptions
Set StudyOptions = Study. FrequencyStudyOptions
StudyOptions. UseSoftSpring = 1
str3 =
"64,31,0,0,4,0,0,0,255,254,255,29,81,0,117,0,97,0,114,0,116,0,101,0,114,0,69,0,121,0,101,0,66,0,97,0,114,0,45,0,49,0,64,0,113,0,117,0,97,0,114,0,116,0,101,0,114,0,101,0,121,0,101,0,98,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,9,0,0,0,3,0,0,0,103,0,0,0,102,0,0,0,101,0,0,0"
StringtoArray str3, var20
Set pDisp3 = Part. Extension . GetObjectByPersistReference3 ((var20),
longstatus)
str4 =
"64,31,0,0,5,0,0,0,255,254,255,27,81,0,117,0,97,0,114,0,116,0,101,0,114,0,66,0,111,0,108,0,116,0,45,0,49,0,64,0,113,0,117,0,97,0,114,0,116,0,101,0,114,0,101,0,121,0,101,0,98,0,97,0,114,0,4,0,0,0,16,0,0,0,1,0,0,0,1,0,0,0,11,0,0,0"
StringtoArray str4, var21
Set pDisp4 = Part. Extension . GetObjectByPersistReference3 ((var21),
longstatus)
Set ContactMgr = Study. ContactManager
If errCode <> 0 Then ErrorMsg SwApp, "Failed to get contact
manager"
ContactMgr. DeleteContactComponent
("Global Contact")
' Add bonded component contacts
Set CWComponentContact = ContactMgr. CreateContactComponent (swsContactTypeBonded,
swsMeshCompatibilityCompatible, pDisp3, errCode)
Debug.Print CWComponentContact. ContactName
Debug.Print " Type as defined in swsContactType_e: " &
CWComponentContact. ContactComponentType
Debug.Print " Include friction? " & CWComponentContact. IncludeFriction
If CWComponentContact. IncludeFriction = True Then
Debug.Print " Friction coefficient: " &
CWComponentContact. FrictionValue
End If
Debug.Print " Global contact condition? (1=yes, 0=no) " &
CWComponentContact. GlobalContact
Debug.Print " Include non-touching faces? (1=yes, 0=no) " &
CWComponentContact. IncludeClearance
If CWComponentContact. IncludeClearance = True Then
Debug.Print " Clearance value: " &
CWComponentContact. ClearanceValue
Debug.Print " Clearance units: " &
CWComponentContact. ClearanceUnit
Debug.Print " Automatically create contact
sets? " & CWComponentContact. IncludeShellEdgeSolidOrShellFace
End If
Debug.Print " Mesh compatibility as defined in
swsMeshCompatibility_e: " & CWComponentContact. Option
Debug.Print " Suppression state as defined in
swsSuppressionState_e: " & CWComponentContact. State
' If suppressed, unsuppress
If CWComponentContact. State = 1 Then
errCode = ContactMgr. SuppressUnsuppressComponentContact (CWComponentContact. ContactName ,
0)
End If
Debug.Print " "
Set CWComponentContact = ContactMgr. CreateContactComponent (swsContactTypeBonded,
swsMeshCompatibilityCompatible, pDisp4, errCode)
Debug.Print CWComponentContact. ContactName
Debug.Print " Type as defined in swsContactType_e: " &
CWComponentContact. ContactComponentType
Debug.Print " Include friction? " & CWComponentContact. IncludeFriction
If CWComponentContact. IncludeFriction = True Then
Debug.Print " Friction coefficient: " &
CWComponentContact. FrictionValue
End If
Debug.Print " Global contact condition? (1=yes, 0=no) " &
CWComponentContact. GlobalContact
Debug.Print " Include non-touching faces? (1=yes, 0=no) " &
CWComponentContact. IncludeClearance
If CWComponentContact. IncludeClearance = True Then
Debug.Print " Clearance value: " &
CWComponentContact. ClearanceValue
Debug.Print " Clearance units: " &
CWComponentContact. ClearanceUnit
Debug.Print " Automatically create contact
sets? " & CWComponentContact. IncludeShellEdgeSolidOrShellFace
End If
Debug.Print " Mesh compatibility as defined in
swsMeshCompatibility_e: " & CWComponentContact. Option
Debug.Print " Suppression state as defined in
swsSuppressionState_e: " & CWComponentContact. State
If CWComponentContact. State = 1 Then
errCode = ContactMgr. SuppressUnsuppressComponentContact (CWComponentContact. ContactName ,
0)
End If
Debug.Print " "
' Apply material to model
Set SolidMgr = Study. SolidManager
If SolidMgr Is Nothing Then ErrorMsg SwApp, "SolidMgr object
not there"
Set SolidComponent = SolidMgr. GetComponentAt (0,
errCode)
If errCode <> 0 Then ErrorMsg SwApp, "No solid component"
Set SolidBody = SolidComponent. GetSolidBodyAt (0,
errCode)
If errCode <> 0 Then ErrorMsg SwApp, "No solid body"
bApp = SolidBody. SetLibraryMaterial ("c:\Program
Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\sldmaterials\solidworks
materials.sldmat", "Ductile Iron")
If bApp = False Then ErrorMsg SwApp, "No material applied"
Set SolidComponent = SolidMgr. GetComponentAt (1,
errCode)
If errCode <> 0 Then ErrorMsg SwApp, "No solid component"
Set SolidBody = SolidComponent. GetSolidBodyAt (0,
errCode)
If errCode <> 0 Then ErrorMsg SwApp, "No solid body"
bApp = SolidBody. SetLibraryMaterial ("c:\Program
Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\sldmaterials\solidworks
materials.sldmat", "Ductile Iron")
If bApp = False Then ErrorMsg SwApp, "No material applied"
' Mesh model
Set CwMesh = Study. Mesh
If CwMesh Is Nothing Then ErrorMsg SwApp, "Failed to get mesh
object"
CwMesh. Quality = 1
Call CwMesh. GetDefaultElementSizeAndTolerance (swsLinearUnitMillimeters,
el, tl)
errCode = Study. CreateMesh (swsLinearUnitMillimeters,
el, tl)
If errCode <> 0 Then ErrorMsg SwApp, "Failed to create mesh"
' Run analysis
errCode = Study. RunAnalysis
If errCode <> 0 Then ErrorMsg SwApp, "Analysis failed"
' Get resonant frequencies for each mode
Set CWResult = Study. Results
If CWResult Is Nothing Then ErrorMsg SwApp, "Failed to get results
object"
Freq = CWResult. GetResonantFrequencies (errCode)
If errCode <> 0 Then ErrorMsg SwApp, "Failed to get resonant
frequencies"
Debug.Print "Resonant frequencies (mode number,
radians/second, cycles/second, period in seconds):"
For i = 0 To UBound(Freq)
Debug.Print " " & Freq(i)
Next i
End Sub
Sub ErrorMsg(SwApp As SldWorks.SldWorks, Message As String)
SwApp. SendMsgToUser2 Message, 0, 0
SwApp. RecordLine "'*** WARNING - General"
SwApp. RecordLine "'*** " & Message
SwApp. RecordLine ""
End Sub
Sub StringtoArray(inputSTR As String, varEntity As Variant)
Dim PID() As Byte
Dim i As Integer
varEntity = Split(inputSTR, ",")
ReDim PID(UBound(varEntity))
For i = 0 To (UBound(varEntity) - 1)
PID(i) = varEntity(i)
Next i
varEntity = PID
End Sub
```

## Run Studies in Batch Mode (VB6)
```vb
Run Studies in Batch Mode Example (VBA)
SOLIDWORKS Simulation API
Help
Run Studies in Batch Mode Example (VBA)
This example shows how to run studies in batch mode.
'---------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in (in SOLIDWORKS, click
' Tools > Add-ins > SOLIDWORKS Simulation > OK ).
' 2. Add the SOLIDWORKS Simulation type library as a reference (in the IDE,
' click Tools > References > SOLIDWORKS
Simulation version type library ).
' 3. Ensure that the specified model exists.
' 4. Open the Immediate window.
'
' Postconditions:
' 1. Opens the model document.
' 2. Sets the run and mesh option for the studies in the batch.
' 3. Analyzes the studies in the batch.
' 4. Gets the result codes of the studies.
' 5. Inspect the Immediate window.
'
' NOTE: Because the model is used elsewhere, do not save changes.
'---------------------------------------------------------------------------
Option Explicit
Dim swApp
As SldWorks.SldWorks
Dim Part As
SldWorks.ModelDoc2
Dim fileName As String
Dim COSMOSWORKS As CosmosWorksLib.COSMOSWORKS
Dim CWAddinCallBack As CosmosWorksLib.CWAddinCallBack
Dim ActDoc As
CosmosWorksLib.CWModelDoc
Dim StudyMngr As
CosmosWorksLib.CWStudyManager
Dim RunStudyResults As CosmosWorksLib.CWRunStudiesResults
Dim RunStudyOptions As CosmosWorksLib.CWRunSpecStudiesRunMeshOptions
Dim errCode As Long
Dim studyName As String
Dim result As Long
Dim errors As Long
Dim warnings As Long
Sub main()
Set swApp = Application.SldWorks
fileName = "C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS 2018\samples\tutorial\api\tjoint.sldprt"
Set Part = swApp. OpenDoc6 (fileName, swDocPART, 1, "",
errors, warnings)
Set CWAddinCallBack = swApp. GetAddInObject ("SldWorks.Simulation")
Set COSMOSWORKS = CWAddinCallBack. COSMOSWORKS
Set ActDoc = COSMOSWORKS. ActiveDoc ()
Set StudyMngr = ActDoc. StudyManager ()
Set RunStudyOptions = StudyMngr. RunSpecifiedStudyOptions ()
errCode = RunStudyOptions. AddStudyOption ("Partial",
swsRunStudiesRunMeshOptions_MeshAndRun)
errCode = RunStudyOptions. AddStudyOption ("Ready - 8
plies", swsRunStudiesRunMeshOptions_MeshAndRun)
Set RunStudyResults = StudyMngr. RunSpecifiedStudyByName (errCode)
errCode = RunStudyResults. GetFirstItem (studyName,
result)
Debug.Print "First study in batch: " & studyName
Debug.Print " Result code as defined in
swsRunStudiesResultsErrorCode_e: " & result
errCode = RunStudyResults. GetNextItem (studyName,
result)
Debug.Print "Next study in batch: " & studyName
Debug.Print " Result code as defined in
swsRunStudiesResultsErrorCode_e: " & result
End Sub
```

## Get Mesh Elements and Nodes (VB6)
```vb
Get Mesh Elements and Nodes Example (VBA)
SOLIDWORKS Simulation API
Help
Get Mesh Elements and Nodes Example (VBA)
This example shows how to get the elements and nodes of a solid mesh.
'----------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in
' (in SOLIDWORKS, click Tools > Add-ins > SOLIDWORKS
Simulation ).
' 2. Add the SOLIDWORKS Simulation type library as a reference
' (in the IDE, click Tools > References > SOLIDWORKS
' Simulation version type library ).
' 3. Open public_documents \samples\Simulation Examples\tutor1.sldprt .
' 4. Open the Immediate window.
'
' Postconditions:
' 1. Activates the Ready study.
' 2. Creates a solid mesh.
' 3. Gets the elements and nodes of the solid mesh.
' 4. Gets the elements and nodes at a depth of 0.001m in the solid mesh.
' 5. Runs the study.
' 6. Inspect the Immediate window.
'
' NOTE: Because the model is used elsewhere, do not save changes.
' ---------------------------------------------------------------------------
Option Explicit
Sub main()
Dim SwApp As SldWorks.SldWorks
Dim COSMOSWORKS As
CosmosWorksLib.COSMOSWORKS
Dim CWAddinCallBack As CosmosWorksLib.CWAddinCallBack
Dim ActDoc As CosmosWorksLib.CWModelDoc
Dim StudyMngr As CosmosWorksLib.CWStudyManager
Dim Study As CosmosWorksLib.CWStudy
Dim CWMesh As CosmosWorksLib.CWMesh
Dim errCode As Long
Dim el As Double, tl As Double
Dim num As Long
Dim idList As Variant
Dim normalNum As Variant
Dim normalVec As Variant
Set SwApp = Application.SldWorks
Set CWAddinCallBack =
SwApp. GetAddInObject ("CosmosWorks.CosmosWorks")
If CWAddinCallBack Is Nothing Then ErrorMsg SwApp, "No
CWAddinCallBack object"
Set COSMOSWORKS = CWAddinCallBack. COSMOSWORKS
If COSMOSWORKS Is Nothing Then ErrorMsg SwApp, "No
CosmosWorks object"
'Get active document
Set ActDoc = COSMOSWORKS. ActiveDoc ()
If ActDoc Is Nothing Then ErrorMsg SwApp, "No active
document"
'Get Ready study
Set StudyMngr = ActDoc. StudyManager ()
If StudyMngr Is Nothing Then ErrorMsg SwApp, "No study
manager object"
StudyMngr. ActiveStudy = 0
Set Study = StudyMngr. GetStudy (0)
If Study Is Nothing Then ErrorMsg SwApp, "No study object"
'Mesh
Set CWMesh = Study. Mesh
If CWMesh Is Nothing Then ErrorMsg SwApp, "No mesh object"
CWMesh. Quality = 1
Call CWMesh. GetDefaultElementSizeAndTolerance (0, el, tl)
errCode = Study. CreateMesh (0, el, tl)
If errCode <> 0 Then ErrorMsg SwApp, "Mesh failed"
'Get the elements of this solid
mesh
num = CWMesh. GetElementList (0, idList)
Debug.Print "Number of elements in the mesh: " & num
'Get the nodes of this solid mesh
num = CWMesh. GetNodeList (0, idList)
Debug.Print "Number of nodes in the mesh: " & num
'Get the elements at a depth of
0.001m in this solid mesh
num = CWMesh. GetSolidElementList (0.001, 2, idList)
Debug.Print "Number of elements at a depth of 0.001m: " & num
'Get the nodes at a depth of 0.001m in this solid mesh
num = CWMesh. GetSolidNodeList (0.001, 2, idList)
Debug.Print "Number of nodes at a depth of 0.001m: " & num
'Get the nodes and normal vectors at the
surface of this solid mesh
num = CWMesh. GetSurfaceNodesAndNormals (idList, normalNum,
normalVec)
Debug.Print "Number of surface nodes: " & num
'Run analysis
errCode = Study. RunAnalysis
If errCode <> 0 Then ErrorMsg SwApp, "Analysis failed
with error code as defined in swsRunAnalysisError_e: " & errCode
End Sub
Sub ErrorMsg(SwApp As SldWorks.SldWorks, Message As String)
SwApp. SendMsgToUser2 Message, 0, 0
SwApp. RecordLine "'*** WARNING - General"
SwApp. RecordLine "'*** " & Message
SwApp. RecordLine ""
End Sub
```

## Create Plots for Static Study (VB6)
```vb
Create Plots for a Static Study Example (VBA)
SOLIDWORKS Simulation API
Help
Create Plots for a Static Study Example (VBA)
This example shows how to create plots of displacement, stress, and strain
results in a static study.
'----------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in (in SOLIDWORKS, click
' Tools > Add-ins > SOLIDWORKS Simulation > OK ).
' 2. Add the SOLIDWORKS Simulation type library as a reference (in the
IDE,
' click Tools > References > SOLIDWORKS Simulation
version type library ).
' 3. Ensure that c:\temp exists.
' 4. Open public_documents \samples\Simulation Examples\tutor1.sldprt .
' 5. Open the Immediate window.
'
' Postconditions:
' 1. Deletes all default static study plots from the model document.
' 2. Activates the Ready static study.
' 3. Meshes and runs the study.
' 4. Creates Displacement1 plot.
' 5. Changes the units of the displacement plot.
' 6. Creates and activates Stress2 plot.
' 7. Probes and annotates Stress2 and saves the
probed result plot to
' c:\temp\tutor1-Ready-Image-1.png .
' 8. Creates Strain1 plot.
' 9. Applies a view orientation to each plot.
' 10. Gets the minimum and maximum values for the result plots.
' 11. Click OK to close each message box.
' 12. Detect stress hot spots and singularities.
' 13. Creates the Stress Hot Spot1 plot and gets hot spot nodes and
elements.
' 14. Inspect c:\temp\tutor1-Ready-Image-1.png .
' 15. Inspect the Immediate window.
'
' NOTE: Because the model is used elsewhere, do not save changes.
' ---------------------------------------------------------------------------
Option Explicit
Sub main()
Dim SwApp As SldWorks.SldWorks
Dim COSMOSWORKS As
CosmosWorksLib.COSMOSWORKS
Dim CWAddinCallBack As CosmosWorksLib.CWAddinCallBack
Dim ActDoc As CosmosWorksLib.CWModelDoc
Dim StudyMngr As CosmosWorksLib.CWStudyManager
Dim Study As CosmosWorksLib.CWStudy
Dim CWMesh As CosmosWorksLib.CWMesh
Dim CWResult As CosmosWorksLib.CWResults
Dim CWCFf As CosmosWorksLib.CWPlot
Dim CWProbeResultsManager As
CosmosWorksLib.CWResultsProbeManager
Dim errCode As Long
Dim el As Double, tl As Double
Dim Disp As Variant, Stress As
Variant, Strn As Variant
Dim SelNodeElemVariant As Variant
Dim SelNodeElemWarnings As Variant
Const URESMax As Double = 1004928
'nm
Const URESMin As Double = 0#
'nm
Const VONMax As Double = 284
'MPa
Const VONMin As Double = 0.797
'MPa
Const MaxStrn As Double = 0.00352
Const MinStrn As Double = 0.00000294
If SwApp Is Nothing Then Set SwApp =
Application.SldWorks
Set CWAddinCallBack = SwApp. GetAddInObject ("CosmosWorks.CosmosWorks")
If CWAddinCallBack Is Nothing Then ErrorMsg SwApp, "No
CWAddinCallBack object", True
Set COSMOSWORKS = CWAddinCallBack. COSMOSWORKS
If COSMOSWORKS Is Nothing Then ErrorMsg SwApp, "No
CosmosWorks object", True
'Get active document
Set ActDoc = COSMOSWORKS. ActiveDoc ()
If ActDoc Is Nothing Then ErrorMsg SwApp, "No active
document", True
'Delete all default static study plots
from this model
ActDoc. DeleteAllDefaultStaticStudyPlots
'Get Ready study
Set StudyMngr = ActDoc. StudyManager ()
If StudyMngr Is Nothing Then ErrorMsg SwApp, "No study
manager object", True
StudyMngr. ActiveStudy = 0
Set Study = StudyMngr. GetStudy (0)
If Study Is Nothing Then ErrorMsg SwApp, "No study object",
True
'Set static study options
Dim StudyOptions As CosmosWorksLib.CWStaticStudyOptions
Set StudyOptions = Study. StaticStudyOptions
StudyOptions. AdaptiveMethodType = 2
StudyOptions. PAdaptiveConvergenceCriteria = 0
StudyOptions. PAdaptiveErrorLimit = 1#
StudyOptions. PAdaptiveMaxIterations = 4
StudyOptions. PAdaptiveMaxPOrder = 5
StudyOptions. PAdaptiveStartingPOrder = 2
StudyOptions. PAdaptiveStrainEnergyErrorLimit = 2#
StudyOptions. IncludeRemarkInReport = 1
StudyOptions. RemarkComment = "Remark comment"
StudyOptions. NoPenetration = 1
StudyOptions. IgnoreClearanceForSurfaceContact = 1
StudyOptions. IncompatibleBondingOption =
swsIncompatibleBondingOption_Automatic
StudyOptions. IncludeGlobalFriction = 1
StudyOptions. FrictionCoefficient = 0.5
'Mesh
Set CWMesh = Study. Mesh
If CWMesh Is Nothing Then ErrorMsg SwApp, "No mesh object",
False
CWMesh. Quality = 1
Call CWMesh. GetDefaultElementSizeAndTolerance (0, el,
tl)
errCode = Study. CreateMesh (0, el, tl)
If errCode <> 0 Then ErrorMsg SwApp, "Mesh failed", True
'Run analysis
errCode = Study. RunAnalysis
If errCode <> 0 Then ErrorMsg SwApp, "Analysis failed", True
'Get results
Set CWResult = Study. Results
If CWResult Is Nothing Then ErrorMsg SwApp, "No result
object", False
Debug.Print "Time taken in seconds for
input phase: " & Study. GetTimeTakenForInputPhase
Debug.Print "Time taken in seconds to solve this study: " &
Study. GetTotalSolutionTime
'Create displacement plot
Set CWCFf = CWResult. CreatePlot (swsResultDisplacementOrAmplitude,
swsDisplacementComponentURES, swsLinearUnitCentimeters, False, errCode)
If CWCFf Is Nothing Then ErrorMsg SwApp, "Failed to create
plot", False
'Change the units from cm to in
errCode = CWCFf. SetComponentUnitAndValueByElem (swsDisplacementComponentURES,
swsLinearUnitInches, False)
'Apply top view orientation
errCode = CWCFf. ApplyNameViewOrientation2 (swsNameViewOrientation_Top)
'Activate plot
errCode = CWCFf. ActivatePlot ()
If errCode <> 0 Then ErrorMsg SwApp, "Plot is not activated",
True
'Get min/max resultant displacements from
plot
Disp = CWCFf. GetMinMaxResultValues (errCode)
If errCode <> 0 Then ErrorMsg SwApp, "No displacement result
values", True
If (Disp(1) < 0.95 * URESMin) Or (Disp(1)
> 1.05 * URESMin) Then
ErrorMsg SwApp, "URES minimum % error = " & (((Disp(1) -
URESMin) / URESMin) * 100), False
End If
If (Disp(3) < 0.95 * URESMax) Or (Disp(3)
> 1.05 * URESMax) Then
ErrorMsg SwApp, "URES maximum % error = " & (((Disp(3) -
URESMax) / URESMax) * 100), False
End If
'Create stress plot
Set CWCFf = CWResult. CreatePlot (swsResultStress,
swsStaticNodalStress_VON, swsStrengthUnitPascal, False, errCode)
If errCode <> 0 Then ErrorMsg SwApp, "Failed to create plot",
True
'Apply bottom view orientation
errCode = CWCFf. ApplyNameViewOrientation2 (swsNameViewOrientation_Bottom)
'Activate plot
errCode = CWCFf. ActivatePlot ()
If errCode <> 0 Then ErrorMsg SwApp, "Plot is not activated",
True
'Get min/max von Mises stresses from plot
Stress = CWCFf. GetMinMaxResultValues (errCode)
If errCode <> 0 Then ErrorMsg SwApp, "No stress results",
True
If (Stress(1) < 0.95 * VONMin) Or
(Stress(1) > 1.05 * VONMin) Then
ErrorMsg SwApp, " Minimum von Mises stress % error = " &
(((Stress(1) - VONMin) / VONMin) * 100), False
End If
If (Stress(3) < 0.95 * VONMax) Or
(Stress(3) > 1.05 * VONMax) Then
ErrorMsg SwApp, "Maximum von Mises stress % error = " &
(((Stress(3) - VONMax) / VONMax) * 100), False
End If
' Probe stress plot at two locations
Set CWProbeResultsManager = CWCFf. GetResultsProbeManager (errCode)
errCode = CWProbeResultsManager. BeginProbing ()
CWProbeResultsManager. ClearSelectionAndAnnotations
CWProbeResultsManager. ProbingOption =
swsProbePostResultOption_AtNodeElemNumber
SelNodeElemVariant = Array(1033, 924)
errCode = CWProbeResultsManager. ShowAnnotationOnSelNodeElems ((SelNodeElemVariant),
SelNodeElemWarnings)
If errCode = 14 Then ErrorMsg SwApp, "Not all specified nodes
or elements are annotated", True
errCode = CWProbeResultsManager. CaptureProbedResultPlotAsPNGImage ("c:\temp\",
"tutor1-Ready-Image-1")
errCode = CWProbeResultsManager. EndProbing ()
'Create strain plot
Set CWCFf = CWResult. CreatePlot (swsResultStrain,
swsStaticElementalStrain_ESTRN, swsStrengthUnitPascal, True, errCode)
If errCode <> 0 Then ErrorMsg SwApp, "Failed to create plot",
True
'Apply left view orientation
errCode = CWCFf. ApplyNameViewOrientation2 (swsNameViewOrientation_Left)
'Activate plot
errCode = CWCFf. ActivatePlot ()
If errCode <> 0 Then ErrorMsg SwApp, "Plot is not activated",
True
'Get min/max strains from plot
Strn = CWCFf. GetMinMaxResultValues (errCode)
If errCode <> 0 Then ErrorMsg SwApp, "No strain results",
True
If (Strn(1) < 0.95 * MinStrn) Or (Strn(1)
> 1.05 * MinStrn) Then
ErrorMsg SwApp, "Minimum strain % error = " & (((Strn(1) -
MinStrn) / MinStrn) * 100), False
End If
If (Strn(3) < 0.95 * MaxStrn) Or (Strn(3)
> 1.05 * MaxStrn) Then
ErrorMsg SwApp, "Maximum strain % error = " & (((Strn(3) -
MaxStrn) / MaxStrn) * 100), False
End If
'Detect stress hot spots and singularities
Dim hotSpotsFound As Boolean
Dim singularitiesFound As Boolean
Dim nodeArray As Variant
Dim elemArray As Variant
errCode = CWResult. RunStressHotSpotDiagnosticsAndDetectSingularities (25,
True, 3, 0.6, 1.5, 0, hotSpotsFound, singularitiesFound)
'Create stress hot spot plot and get
detected hot spots
If hotSpotsFound = True Then
errCode = CWResult. GetDetectedHotSpotNodes (nodeArray)
errCode = CWResult. GetDetectedHotSpotElements (elemArray)
Set CWCFf = CWResult. CreateStressHotSpotPlot (True,
True, errCode)
End If
End Sub
Sub ErrorMsg(SwApp As SldWorks.SldWorks, Message As String,
EndTest As Boolean)
SwApp. SendMsgToUser2 Message, 0, 0
SwApp. RecordLine "'*** WARNING - General"
SwApp. RecordLine "'*** " & Message
SwApp. RecordLine ""
End Sub
```

## Create Results Equation Plot (VB6)
```vb
Create Results Equation Plot Example (VBA)
SOLIDWORKS Simulation API
Help
Create Results Equation Plot Example (VBA)
This example shows how to create a results equation plot in a static study.
'----------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in (in SOLIDWORKS, click
' Tools > Add-ins > SOLIDWORKS Simulation > OK ).
' 2. Add the SOLIDWORKS Simulation type library as a reference (in the IDE,
' click Tools > References > SOLIDWORKS
Simulation version type library ).
' 3. Click View > Locals Window .
' 4. Add a breakpoint to the line that begins with "EqnResultMinMax = ".
' 5. Add a breakpoint to the line that begins with "EqnResultEntire = ".
' 6. Open public_documents \samples\Simulation Examples\tutor1.sldprt .
'
' Postconditions:
' 1. Activates Ready static study.
' 2. Meshes and runs the study.
' 3. Creates results equation plot, Equation1 .
' 4. Press F8 to execute the current line.
' 5. Inspect local variable EqnResultMinMax to see the nodal minimum and maximum
' for Equation1 for the entire
model.
' 6. Press F8 to execute the current line.
' 7. Inspect local variable EqnResultEntire to see the nodal results
' for Equation1 for the entire model.
' 8. Press F5 to complete the macro.
' 9. Inspect My Equation Plot in the graphics area.
'
' NOTE: Because the model is used elsewhere, do not save changes.
' ---------------------------------------------------------------------------
Option Explicit
Sub main()
Dim SwApp As SldWorks.SldWorks
Dim COSMOSWORKS As
CosmosWorksLib.COSMOSWORKS
Dim CWAddinCallBack As CosmosWorksLib.CWAddinCallBack
Dim ActDoc As CosmosWorksLib.CWModelDoc
Dim StudyMngr As CosmosWorksLib.CWStudyManager
Dim Study As CosmosWorksLib.CWStudy
Dim CWMesh As CosmosWorksLib.CWMesh
Dim CWResult As CosmosWorksLib.CWResults
Dim PostPlot As CosmosWorksLib.CWPlot
Dim errCode As Long
Dim el As Double, tl As Double
Dim sEqtn As String
Dim sTitle As String
If SwApp Is Nothing Then Set SwApp =
Application.SldWorks
Set CWAddinCallBack = SwApp. GetAddInObject ("CosmosWorks.CosmosWorks")
If CWAddinCallBack Is Nothing Then ErrorMsg SwApp, "No
CWAddinCallBack object", True
Set COSMOSWORKS = CWAddinCallBack. COSMOSWORKS
If COSMOSWORKS Is Nothing Then ErrorMsg SwApp, "No
CosmosWorks object", True
'Get active document
Set ActDoc = COSMOSWORKS. ActiveDoc ()
If ActDoc Is Nothing Then ErrorMsg SwApp, "No active
document", True
'Get Ready study
Set StudyMngr = ActDoc. StudyManager ()
If StudyMngr Is Nothing Then ErrorMsg SwApp, "No study
manager object", True
StudyMngr. ActiveStudy = 0
Set Study = StudyMngr. GetStudy (0)
If Study Is Nothing Then ErrorMsg SwApp, "No study object",
True
'Mesh
Set CWMesh = Study. Mesh
If CWMesh Is Nothing Then ErrorMsg SwApp, "No mesh object",
False
CWMesh. Quality = 0
Call CWMesh. GetDefaultElementSizeAndTolerance (0, el,
tl)
errCode = Study. CreateMesh (0, el, tl)
If errCode <> 0 Then ErrorMsg SwApp, "Mesh failed", True
'Run analysis
errCode = Study. RunAnalysis
If errCode <> 0 Then ErrorMsg SwApp, "Analysis failed", True
'Get results
Set CWResult = Study. Results
If CWResult Is Nothing Then ErrorMsg SwApp, "No results",
False
'Create results equation plot
sEqtn = """P1: 1st Principal Stress"" + ""P2: 2nd Principal
Stress"" + ""P3: 3rd Principal Stress"""
sTitle = "My Equation Plot"
Set PostPlot = CWResult. CreateResultsEquationPlot (sEqtn,
sTitle, False, swsUnitSI, 1, swsShellFace_Top, errCode)
Dim EqnResultMinMax
As Variant
Dim EqnResultEntire As Variant
EqnResultMinMax = CWResult. GetMinMaxResultsEquationValues (sEqtn,
False, 0, 1, 0, errCode)
EqnResultEntire = CWResult. GetResultsEquationValuesForEntities (sEqtn,
False, 0, 1, 0, Nothing, errCode)
End Sub
Sub ErrorMsg(SwApp As SldWorks.SldWorks, Message As String,
EndTest As Boolean)
SwApp. SendMsgToUser2 Message, 0, 0
SwApp. RecordLine "'*** WARNING - General"
SwApp. RecordLine "'*** " & Message
SwApp. RecordLine ""
End Sub
```

## Duplicate Study (VB6)
```vb
Sub ErrorMsg(SwApp As SldWorks.SldWorks, Message As String, EndTest As Boolean)
    SwApp.SendMsgToUser2 Message, 0, 0
    SwApp.RecordLine "'*** WARNING - General"
    SwApp.RecordLine "'*** " & Message
    SwApp.RecordLine ""
End Sub
```

## Copy Mesh and Gen Report (VB6)
```vb
Copy Mesh and Generate Report Example (VBA)
SOLIDWORKS Simulation API
Help
Copy Mesh and Generate Report Example (VBA)
This example shows how to copy the mesh from one study to another and
generate a report.
'----------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in (in SOLIDWORKS, click
' Tools > Add-ins > SOLIDWORKS Simulation > OK ).
' 2. Add the SOLIDWORKS Simulation type library as a reference (in the IDE,
' click Tools > References > SOLIDWORKS
Simulation version type library ).
' 3. Open public_documents \samples\Simulation Examples\Verification\buckling2.sldprt .
'
' Postconditions:
' 1. Meshes the Quality and Draft studies.
' 2. Copies the mesh from the Quality study to the Draft study.
' 3. Click OK to close any dialog boxes.
' 4. Generates a report for the Draft study in c:\temp .
'
' NOTE: Because the model is used elsewhere, do not save changes.
' ---------------------------------------------------------------------------
Option Explicit
Sub main()
Dim SwApp As SldWorks.SldWorks
Dim COSMOSWORKS As
CosmosWorksLib.COSMOSWORKS
Dim CWAddinCallBack As CosmosWorksLib.CWAddinCallBack
Dim ActDoc As CosmosWorksLib.CWModelDoc
Dim StudyMngr As CosmosWorksLib.CWStudyManager
Dim Study As CosmosWorksLib.CWStudy
Dim CWMesh As CosmosWorksLib.CWMesh
Dim errCode As Long, n As Long
Dim el As Double, tl As Double
Const NoOfElements1 As Double = 712
Const NoOfElements2 As Double = 2686
If SwApp Is Nothing Then Set SwApp =
Application.SldWorks
Set CWAddinCallBack = SwApp. GetAddInObject ("CosmosWorks.CosmosWorks")
If CWAddinCallBack Is Nothing Then ErrorMsg SwApp, "No CwAddincallback
object"
Set COSMOSWORKS = CWAddinCallBack. COSMOSWORKS
If COSMOSWORKS Is Nothing Then ErrorMsg SwApp, "No
CosmosWorks object"
'Get active document
Set ActDoc = COSMOSWORKS. ActiveDoc ()
If ActDoc Is Nothing Then ErrorMsg SwApp, "No active
document"
'Get the first study
Set StudyMngr = ActDoc. StudyManager ()
If StudyMngr Is Nothing Then ErrorMsg SwApp, "No
CWStudyManager object"
StudyMngr. ActiveStudy = 0
Set Study = StudyMngr. GetStudy (0)
If Study Is Nothing Then ErrorMsg SwApp, "No CWStudy object"
'Mesh the first study
Set CWMesh = Study. Mesh
If CWMesh Is Nothing Then ErrorMsg SwApp, "No CWMesh object"
CWMesh. Quality = 1
Call CWMesh. GetDefaultElementSizeAndTolerance (0, el,
tl)
errCode = Study. CreateMesh (0, el, tl)
If errCode <> 0 Then ErrorMsg SwApp, "Mesh failed"
n = CWMesh. NodeCount
If (n < 0.95 * NoOfElements2) Or (n > 1.05
* NoOfElements2) Then
ErrorMsg SwApp, " % error = " & (((n - NoOfElements2) /
NoOfElements2) * 100)
End If
'Get the second study
StudyMngr. ActiveStudy = 1
Set Study = StudyMngr. GetStudy (1)
If Study Is Nothing Then ErrorMsg SwApp, "No CWStudy object"
'Mesh the second study
Set CWMesh = Study. Mesh
If CWMesh Is Nothing Then ErrorMsg SwApp, "No CWMesh object"
CWMesh. Quality = 0
Call CWMesh. GetDefaultElementSizeAndTolerance (0, el,
tl)
errCode = Study. CreateMesh (0, el, tl)
If errCode <> 0 Then ErrorMsg SwApp, "Mesh failed"
n = CWMesh. NodeCount
If (n < 0.95 * NoOfElements1) Or (n > 1.05 * NoOfElements1)
Then
ErrorMsg SwApp, " % error = " & (((n - NoOfElements1) /
NoOfElements1) * 100)
End If
'Copy mesh from the first study to the second
study
errCode = Study. CopyMeshFromStudy ("Quality")
If errCode <> 0 Then ErrorMsg SwApp, "Mesh failed to
copy"
n = CWMesh. NodeCount
If (n < 0.95 * NoOfElements2) Or (n > 1.05 * NoOfElements2)
Then
ErrorMsg SwApp, " % error = " & (((n - NoOfElements2) /
NoOfElements2) * 100)
End If
'Generate Report
errCode = Study. GenerateReport ("C:\temp",
"report", False)
If errCode <> 1 Then ErrorMsg SwApp, "Failed to
generate report"
End Sub
Sub ErrorMsg(SwApp As SldWorks.SldWorks, Message As String)
SwApp. SendMsgToUser2 Message, 0, 0
SwApp. RecordLine "'*** WARNING - General"
SwApp. RecordLine "'*** " & Message
SwApp. RecordLine ""
End Sub
```

## Create Curvature-based Mesh (VB6)
```vb
Create Curvature-based Mesh Example (VBA)
SOLIDWORKS Simulation API
Help
Create Curvature-based Mesh Example (VBA)
This example shows how to create a curvature-based mesh for a part.
'----------------------------------------------------------------------------
' Preconditions:
' 1. Add the SOLIDWORKS Simulation as an add-in (in SOLIDWORKS, click
' Tools > Add-ins > SOLIDWORKS Simulation > OK ).
' 2. Add the SOLIDWORKS Simulation type library as a reference (in the IDE,
' click Tools > References > SOLIDWORKS
Simulation version type library ).
' 3. Ensure that the specified part exists.
' 4. Open the Immediate window.
'
' Postconditions:
' 1. Opens the part.
' 2. Creates a curvature-based mesh for study, Ready.
' 3. In SOLIDWORKS, click Ready , right-click Mesh
' in the Simulation study tree, and click Show Mesh .
' 4. Gets mesh connectivity data and node normal coordinates. Because the
' amount of data returned is large, the display code is
commented out.
' Uncomment the display code or add break points to inspect
the returned
' arrays in the Locals window.
' 5. Zoom in on the part and examine the mesh.
' 6. Examine the Immediate window.
'
' NOTE: Because the part document is used elsewhere, do not save changes.
'---------------------------------------------------------------------------
Option Explicit
Sub main()
Dim swApp As SldWorks.SldWorks
Dim COSMOSWORKS As Object
Dim COSMOSObject As CosmosWorksLib.CwAddincallback
Dim ActDoc As CosmosWorksLib.CWModelDoc
Dim StudyMngr As CosmosWorksLib.CWStudyManager
Dim Study As CosmosWorksLib.CWStudy
Dim CwMesh As CosmosWorksLib.CwMesh
Dim errCode As Long
Dim errors As Long
Dim warnings As Long
Const fileName As String = "C:\Users\Public\Documents\SOLIDWORKS\SOLIDWORKS
2019\samples\Simulation Examples\bikeframe.sldprt"
Set swApp = Application.SldWorks
Set COSMOSObject = swApp. GetAddInObject ("SldWorks.Simulation")
If COSMOSObject Is Nothing Then ErrorMsg swApp, "No
Simulation add-in"
Set COSMOSWORKS = COSMOSObject. COSMOSWORKS
If COSMOSWORKS Is Nothing Then ErrorMsg swApp, "No
COSMOSWORKS object"
swApp. OpenDoc6 fileName,
swDocumentTypes_e.swDocPART, swOpenDocOptions_Silent, "", errors, warnings
Set ActDoc = COSMOSWORKS. ActiveDoc ()
If ActDoc Is Nothing Then ErrorMsg swApp, "No active
document"
Set StudyMngr = ActDoc. StudyManager ()
If StudyMngr Is Nothing Then ErrorMsg swApp, "No
CWStudyManager object"
Set Study = StudyMngr. GetStudy (0)
If Study Is Nothing Then ErrorMsg swApp, "No study"
If (Study. Name <> "Ready") Then ErrorMsg swApp,
"Wrong study"
Debug.Print ("Name of study: " & Study. Name )
Set CwMesh = Study. Mesh
If CwMesh Is Nothing Then ErrorMsg swApp, "No mesh"
' Set type of mesh to curvature-based
CwMesh. MesherType = swsMesherTypeAlternate
' Set mesh parameters
CwMesh. GrowthRatio = 1.6
CwMesh. MinElementsInCircle = 8
CwMesh. GetDefaultMaxAndMinElementSize
swsLinearUnitMillimeters, 20, 4
' Create mesh
errCode = Study. CreateMesh (swsLinearUnitMillimeters,
20, 1)
Debug.Print "Error code: " & errCode '0 indicates successful
creation of the mesh
If errCode <> 0 Then ErrorMsg swApp, "Mesh failed;
check error code"
' Get maximum and minimum element sizes
used for boundaries
Debug.Print "Maximum element size used for boundaries with
lowest curvature: " & CwMesh. MaxElementSize
Debug.Print "Minimum element size used for boundaries with
highest curvature: " & CwMesh. MinElementSize
Dim Var As Variant
Dim Var1 As Variant
Dim i As Long
'Debug.Print "Mesh connectivity data"
'Debug.Print "Number of nodes in connectivity unit, node
numbers in connectivity unit"
i = 0
Var = CwMesh. GetConnectivity (errCode)
'Do While i <= UBound(Var)
'Debug.Print Var(i)
'i = i + 1
'Loop
'Debug.Print "Mesh node normals"
'Debug.Print "Node #, x-coord, y-coord, z-coord of node's
normal vector"
i = 0
Var1 = CwMesh. GetNodeOutwardDirections (errCode)
'Do While i <= UBound(Var1)
'If 0 = i Mod 4 Then
'Debug.Print
"Node: " & Var1(i)
'Debug.Print
" Normal vector (x, y, z coordinates):"
'Else
'Debug.Print
" " & Var1(i)
'End If
'i = i + 1
'Loop
End Sub
Sub ErrorMsg(swApp As SldWorks.SldWorks, Message As String)
swApp. SendMsgToUser2 Message, 0, 0
swApp. RecordLine "'*** WARNING - General"
swApp. RecordLine "'*** " & Message
swApp. RecordLine ""
End Sub
```

