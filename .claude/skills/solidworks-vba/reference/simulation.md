# Simulation (CosmosWorks) API reference, SolidWorks 2023

Generated from cworksapi.chm. Thermal-relevant interfaces only. Entry: ICosmosWorks from the add-in
(swApp.GetAddInObject("SldWorks.Simulation")), then .ActiveDoc -> ICWModelDoc -> .StudyManager.

## ICosmosWorks

### ICosmosWorks.ActiveDoc
Gets the model document.
```vb
ReadOnly Property ActiveDoc As CWModelDoc
```
Property Value Model document
VBA: See CosmosWorks::ActiveDoc .
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Materials

### ICosmosWorks.VersionNumber
Gets the SOLIDWORKS Simulation version number.
```vb
ReadOnly Property VersionNumber As System.String
```
Property Value Version number
VBA: See CosmosWorks::VersionNumber .
Examples: Visual Basic for Applications

## ICWModelDoc

### ICWModelDoc.AddDefaultDropTestStudyPlot
Specifies a default drop test study result plot for the active document.
```vb
Function AddDefaultDropTestStudyPlot( ByVal NResultType As System.Integer, ByVal NResultComponent As System.Integer ) As System.Integer
```
Parameters NResultType Type of drop test study result to plot as defined by swsDropTestStudyResultType_e NResultComponent If NResultType is swsDropTestStudyResultType_e... Then NResultComponent is... swsDropTestResultDisplacement Displacement component to plot as defined by swsDisplacementComponent_e swsDropTestResultElementalStrain Strain component to plot as defined by swsStrainComponent_e swsDropTestResultElementalStress or swsDropTestResultNodalStress Stress component to plot as defined by swsStressComponent_e Return Value Error code as defined by swsAddDefaultDropTestStudyPlotResultError_e
Return Value Error code as defined by swsAddDefaultDropTestStudyPlotResultError_e
VBA: See CWModelDoc::AddDefaultDropTestStudyPlot .
Examples: Visual Basic for Applications; Example Create Drop Test Study

### ICWModelDoc.AddDefaultFatigueStudyPlot
Specifies a default fatigue study result plot for the active document.
```vb
Function AddDefaultFatigueStudyPlot( ByVal NResultType As System.Integer ) As System.Integer
```
Parameters NResultType Type of fatigue study result to plot as defined by swsFatigueStudyResultType_e Return Value Error code as defined by swsAddDefaultFatigueStudyPlotResultError_e
Return Value Error code as defined by swsAddDefaultFatigueStudyPlotResultError_e
VBA: See CWModelDoc::AddDefaultFatigueStudyPlot .
Examples: Visual Basic for Applications; Example Create Fatigue Study

### ICWModelDoc.AddDefaultFrequencyOrBucklingStudyPlot
Specifies a default frequency or buckling study result plot for the active document.
```vb
Function AddDefaultFrequencyOrBucklingStudyPlot( ByVal BAllModeShapes As System.Boolean, ByVal NDefMaxModeShapeValue As System.Integer, ByVal BDisplacement As System.Boolean, ByVal NDisplacementValue As System.Integer ) As System.Integer
```
Parameters BAllModeShapes True to create plots for all mode shapes, false to create plots for the first NDefMaxModeShapeValue mode shapes NDefMaxModeShapeValue Maximum number of mode shapes to plot; valid only if BAllModeShapes is false BDisplacement True to plot displacement, false to not NDisplacementValue Type of displacement result to plot as defined by swsFrequencyBucklingResultDisplacementComponentTypes_e Return Value Error code as defined by swsAddDefaultFrequencyOrBucklingStudyPlotResultError_e
Return Value Error code as defined by swsAddDefaultFrequencyOrBucklingStudyPlotResultError_e
VBA: See CWModelDoc::AddDefaultFrequencyOrBucklingStudyPlot .
Examples: Visual Basic for Applications; Example Create Frequency Study with Solid Mesh

### ICWModelDoc.AddDefaultNonLinearStudyPlot
Specifies a default nonlinear study result plot for the active document.
```vb
Function AddDefaultNonLinearStudyPlot( ByVal NResultType As System.Integer, ByVal NResultComponent As System.Integer ) As System.Integer
```
Parameters NResultType Default nonlinear study result plot types as defined in swsNonlinearStudyResultTypes_e NResultComponent If NResultType is ... Then NResultComponent is... 0 = Nodal stress or 1 = Elemental stress Stress component to plot as defined by swsStressComponent_e 2 = Displacement Displacement component to plot as defined by swsDisplacementComponent_e 3 = Nodal strain or 4 = Elemental strain Strain component to plot as defined by swsStrainComponent_e Return Value Error code as defined by swsAddDefaultNonLinearStudyPlotResultError_e
Return Value Error code as defined by swsAddDefaultNonLinearStudyPlotResultError_e
VBA: See CWModelDoc::AddDefaultNonLinearStudyPlot .
Examples: Visual Basic for Applications

### ICWModelDoc.AddDefaultOptimizationStudyPlot
Specifies a default optimization study result plot for the active document.
```vb
Function AddDefaultOptimizationStudyPlot( ByVal NOptimizationResultType As System.Integer ) As System.Integer
```
Parameters NOptimizationResultType Type of optimization study result to plot as defined by swsOptimizationStudyResultType_e Return Value Error code as defined by swsDefaultOptimizationDesignStudyPlotResultError_e
Return Value Error code as defined by swsDefaultOptimizationDesignStudyPlotResultError_e
VBA: See CWModelDoc::AddDefaultOptimizationStudyPlot .
Examples: Visual Basic for Applications

### ICWModelDoc.AddDefaultStaticStudyPlot
Specifies a default static study result plot for the active document.
```vb
Function AddDefaultStaticStudyPlot( ByVal NResultType As System.Integer, ByVal NResultComponent As System.Integer ) As System.Integer
```
Parameters NResultType Type of static study result to plot as defined by swsDefaultStaticResultTypes_e NResultComponent If NResultType is swsDefaultStaticResultTypes_e... Then NResultComponent is... swsStaticResultDisplacement Displacement to plot as defined by swsStaticResultDisplacementComponentTypes_e swsStaticResultElementalStrain Elemental strain to plot as defined by swsStaticResultElementalStrainComponentTypes_e swsStaticResultElementalStress Elemental stress to plot as defined by swsStaticResultElementalStressComponentTypes_e swsStaticResultNodalStrain Nodal strain to plot as defined by swsStaticResultNodalStrainComponentTypes_e swsStaticResultNodalStress Nodal stress to plot as defined by swsStaticResultNodalStressComponentTypes_e swsStaticResultBoltPinCheck Not valid swsStaticResultFactorOfSafety Not valid Return Value Error code as defined by swsAddDefaultStaticStudyPlotResultError_e
Return Value Error code as defined by swsAddDefaultStaticStudyPlotResultError_e
VBA: See CWModelDoc::AddDefaultStaticStudyPlot .
Examples: Visual Basic for Applications; Example Analyze Part

### ICWModelDoc.AddDefaultThermalStudyPlot
Specifies a default thermal study result plot for the active document.
```vb
Function AddDefaultThermalStudyPlot( ByVal NResultComponent As System.Integer, ByVal BTransient As System.Boolean ) As System.Integer
```
Parameters NResultComponent Type of thermal study result to plot as defined by swsThermalResultComponentTypes_e BTransient True if transient, false if steady state (see Remarks ) Return Value Error code as defined by swsAddDefaultThermalStudyPlotResultError_e
Return Value Error code as defined by swsAddDefaultThermalStudyPlotResultError_e
VBA: See CWModelDoc::AddDefaultThermalStudyPlot .
Remarks: If BTransient is true, this plot is for a transient thermal study in which an initial temperature profile is defined for the thermal load.
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWModelDoc.DefaultUnitSystem
Gets and sets the default unit system for this model document.
```vb
Property DefaultUnitSystem As System.Integer
```
Property Value Default unit system as defined in swsUnitSystem_e
VBA: See CWModelDoc::DefaultUnitSystem .
Examples: Visual Basic for Applications

### ICWModelDoc.DeleteAllDefaultDropTestStudyPlots
Deletes all default drop test study plots.
```vb
Function DeleteAllDefaultDropTestStudyPlots() As System.Integer
```
Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
VBA: See CWModelDoc::DeleteAllDefaultDropTestStudyPlots .
Examples: Visual Basic for Applications

### ICWModelDoc.DeleteAllDefaultFatigueStudyPlots
Deletes all default fatigue study plots.
```vb
Function DeleteAllDefaultFatigueStudyPlots() As System.Integer
```
Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
VBA: See CWModelDoc::DeleteAllDefaultFatigueStudyPlots .
Examples: Visual Basic for Applications

### ICWModelDoc.DeleteAllDefaultNonLinearStudyPlots
Deletes all default nonlinear study plots.
```vb
Function DeleteAllDefaultNonLinearStudyPlots() As System.Integer
```
Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
VBA: See CWModelDoc::DeleteAllDefaultNonLinearStudyPlots .
Examples: Visual Basic for Applications

### ICWModelDoc.DeleteAllDefaultStaticStudyPlots
Deletes all default static study plots from this model.
```vb
Function DeleteAllDefaultStaticStudyPlots() As System.Integer
```
Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
VBA: See CWModelDoc::DeleteAllDefaultStaticStudyPlots .
Examples: Visual Basic for Applications; Example Create Plots for a Static Study

### ICWModelDoc.DeleteAllDefaultThermalStudyPlots
Deletes all default thermal study plots.
```vb
Function DeleteAllDefaultThermalStudyPlots() As System.Integer
```
Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
VBA: See CWModelDoc::DeleteAllDefaultThermalStudyPlots .
Examples: Visual Basic for Applications

### ICWModelDoc.DeleteDefaultDropTestStudyPlot
Deletes the specified default drop test study plot.
```vb
Function DeleteDefaultDropTestStudyPlot( ByVal NResultType As System.Integer, ByVal NResultComponent As System.Integer ) As System.Integer
```
Parameters NResultType Type of drop test study result plot to delete as defined by swsDropTestStudyResultType_e NResultComponent If NResultType is swsDropTestStudyResultType_e... Then NResultComponent is... swsDropTestResultDisplacement Displacement component to plot as defined by swsDisplacementComponent_e swsDropTestResultElementalStrain Strain component to plot as defined by swsStrainComponent_e swsDropTestResultElementalStress or swsDropTestResultNodalStress Stress component to plot as defined by swsStressComponent_e Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
VBA: See CWModelDoc::DeleteDefaultDropTestStudyPlot .
Examples: Visual Basic for Applications

### ICWModelDoc.DeleteDefaultFatigueStudyPlot
Deletes the specified default fatigue study plot.
```vb
Function DeleteDefaultFatigueStudyPlot( ByVal NResultType As System.Integer ) As System.Integer
```
Parameters NResultType Type of fatigue study result plot to delete as defined by swsFatigueStudyResultType_e Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
VBA: See CWModelDoc::DeleteDefaultFatigueStudyPlot .
Examples: Visual Basic for Applications

### ICWModelDoc.DeleteDefaultNonLinearStudyPlot
Deletes the specified default nonlinear study plot.
```vb
Function DeleteDefaultNonLinearStudyPlot( ByVal NResultType As System.Integer, ByVal NResultComponent As System.Integer ) As System.Integer
```
Parameters NResultType Type of nonlinear study result plot to delete as defined in swsNonlinearStudyResultTypes_e NResultComponent If NResultType is ... Then NResultComponent is... 0 = Nodal stress or 1 = Elemental stress Stress component to plot as defined by swsStressComponent_e 2 = Displacement Displacement component to plot as defined by swsDisplacementComponent_e 3 = Nodal strain or 4 = Elemental strain Strain component to plot as defined by swsStrainComponent_e Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
VBA: See CWModelDoc::DeleteDefaultNonLinearStudyPlot .
Examples: Visual Basic for Applications

### ICWModelDoc.DeleteDefaultStaticStudyPlot
Deletes the specified default static study plot.
```vb
Function DeleteDefaultStaticStudyPlot( ByVal NResultType As System.Integer, ByVal NResultComponent As System.Integer ) As System.Integer
```
Parameters NResultType Type of static study result plot to delete as defined by swsDefaultStaticResultTypes_e NResultComponent If NResultType is swsDefaultStaticResultTypes_e... Then NResultComponent is... swsStaticResultDisplacement Displacement to plot as defined by swsStaticResultDisplacementComponentTypes_e swsStaticResultElementalStrain Elemental strain to plot as defined by swsStaticResultElementalStrainComponentTypes_e swsStaticResultElementalStress Elemental stress to plot as defined by swsStaticResultElementalStressComponentTypes_e swsStaticResultNodalStrain Nodal strain to plot as defined by swsStaticResultNodalStrainComponentTypes_e swsStaticResultNodalStress Nodal stress to plot as defined by swsStaticResultNodalStressComponentTypes_e swsStaticResultBoltPinCheck Not valid swsStaticResultFactorOfSafety Not valid Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
VBA: See CWModelDoc::DeleteDefaultStaticStudyPlot .
Examples: Visual Basic for Applications

### ICWModelDoc.DeleteDefaultThermalStudyPlot
Deletes the specified default thermal study plot.
```vb
Function DeleteDefaultThermalStudyPlot( ByVal NResultComponent As System.Integer, ByVal BTransient As System.Boolean ) As System.Integer
```
Parameters NResultComponent Type of thermal study result plot to delete as defined by swsThermalResultComponentTypes_e BTransient True if transient, false if steady state (see Remarks ) Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
Return Value Error code as defined in swsResultPlotDelete_ErrorCode_e
VBA: See CWModelDoc::DeleteDefaultThermalStudyPlot .
Remarks: If BTransient is true, this plot is for a transient thermal study in which an initial temperature profile is defined for the thermal load.
Examples: Visual Basic for Applications

### ICWModelDoc.GetSimulationOptionDefaultPlotsBoundaryColorInRGB
Gets the RGB color for the specified default plot boundary option.
```vb
Function GetSimulationOptionDefaultPlotsBoundaryColorInRGB( ByVal NBoundaryOption As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NBoundaryOption Boundary option as defined in swsSimulationOptionDefaultPlotsBoundaryColorInRGBBoundaryOption_e ErrorCode Error code as defined in swsSimulationOptionDefaultPlotsBoundaryColorInRGBError_e Return Value Array of three values; [0-255][0-255][0-255]
Return Value Array of three values; [0-255][0-255][0-255]
VBA: See CWModelDoc::GetSimulationOptionDefaultPlotsBoundaryColorInRGB .
Examples: Visual Basic for Applications; Example Analyze Part

### ICWModelDoc.GetSimulationOptionDoubleValue
Gets the double value for the specified user preference.
```vb
Function GetSimulationOptionDoubleValue( ByVal NUserPreferenceValue As System.Integer ) As System.Double
```
Parameters NUserPreferenceValue User preference as defined by swsUserPreferenceDoubleValue_e Return Value Double value
Return Value Double value
VBA: See CWModelDoc::GetSimulationOptionDoubleValue .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWModelDoc.GetSimulationOptionIntegerValue
Gets the integer value for the specified user preference.
```vb
Function GetSimulationOptionIntegerValue( ByVal UserPreferenceValue As System.Integer ) As System.Integer
```
Parameters UserPreferenceValue User preference as defined by swsUserPreferenceIntegerValue_e Return Value Integer value (see Remarks )
Return Value Integer value (see Remarks )
VBA: See CWModelDoc::GetSimulationOptionIntegerValue .
Remarks: Values for colors are defined in swsWindowsBasicColors_e .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWModelDoc.GetSimulationOptionStringValue
Gets the string value for the specified user preference.
```vb
Function GetSimulationOptionStringValue( ByVal NUserPreferenceValue As System.Integer ) As System.String
```
Parameters NUserPreferenceValue User preference as defined by swsUserPreferenceStringValue_e Return Value String value
Return Value String value
VBA: See CWModelDoc::GetSimulationOptionStringValue .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWModelDoc.GetSimulationOptionToggle
Gets the boolean value for the specified user preference.
```vb
Function GetSimulationOptionToggle( ByVal NUserPreferenceValue As System.Integer ) As System.Boolean
```
Parameters NUserPreferenceValue User preference as defined by swsUserPreferenceToggle_e Return Value Boolean value
Return Value Boolean value
VBA: See CWModelDoc::GetSimulationOptionToggle .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWModelDoc.SetSimulationOptionDefaultPlotsBoundaryColorInRGB
Sets the RGB color for the specified default plot boundary option.
```vb
Function SetSimulationOptionDefaultPlotsBoundaryColorInRGB( ByVal NBoundaryOption As System.Integer, ByVal NRed As System.Integer, ByVal NGreen As System.Integer, ByVal NBlue As System.Integer ) As System.Integer
```
Parameters NBoundaryOption Boundary option as defined in swsSimulationOptionDefaultPlotsBoundaryColorInRGBBoundaryOption_e NRed 0 <= Red component <= 255 NGreen 0 <= Green component <= 255 NBlue 0 <= Blue component <= 255 Return Value Error code as defined in swsSimulationOptionDefaultPlotsBoundaryColorInRGBError_e
Return Value Error code as defined in swsSimulationOptionDefaultPlotsBoundaryColorInRGBError_e
VBA: See CWModelDoc::SetSimulationOptionDefaultPlotsBoundaryColorInRGB .
Examples: Visual Basic for Applications; Example Analyze Part

### ICWModelDoc.SetSimulationOptionDoubleValue
Sets the double value for the specified user preference.
```vb
Function SetSimulationOptionDoubleValue( ByVal NUserPreferenceValue As System.Integer, ByVal DVal As System.Double ) As System.Boolean
```
Parameters NUserPreferenceValue User preference as defined by swsUserPreferenceDoubleValue_e DVal Double value Return Value True if option successfully set, false if not
Return Value True if option successfully set, false if not
VBA: See CWModelDoc::SetSimulationOptionDoubleValue .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWModelDoc.SetSimulationOptionIntegerValue
Sets the integer value for the specified user preference.
```vb
Function SetSimulationOptionIntegerValue( ByVal UserPreferenceValue As System.Integer, ByVal Value As System.Integer ) As System.Boolean
```
Parameters UserPreferenceValue User preference as defined by swsUserPreferenceIntegerValue_e Value Integer value (see Remarks ) Return Value True if option successfully set, false if not
Return Value True if option successfully set, false if not
VBA: See CWModelDoc::SetSimulationOptionIntegerValue .
Remarks: To set color options you can either: Locate the color you want, for example swsAqua, from swsWindowsBasicColors_e . Call this method, specifying Value with the integer value for swsAqua, 16776960, e.g., ICWModelDoc::SetSimulationOptionIntegerValue(swsUserPreferenceIntegerValue_e.swsPlotShowHiddenBodyTranslucentSingleColor, 16776960). -or- Select a color from http://cloford.com/resources/colours/500col.htm (e.g., #FF0000 is the hexadecimal for Red 1). Reverse the hexadecimal value of the selected color (e.g., #0000FF is the reverse hexadecimal for Red 1). Call ICWModelDoc::SetSimulationOptionIntegerValue (swsUserPreferenceIntegerValue_e. option , &h color_reverse_hexadecimal ), e.g., ICWModelDoc::SetSimulationOptionIntegerValue(swsUserPreferenceIntegerValue_e.swsPlotShowHiddenBodyTranslucentSingleColor, &h0000ff).
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWModelDoc.SetSimulationOptionStringValue
Sets the string value for the specified user preference.
```vb
Function SetSimulationOptionStringValue( ByVal NUserPreferenceValue As System.Integer, ByVal StringVal As System.String ) As System.Boolean
```
Parameters NUserPreferenceValue User preference as defined by swsUserPreferenceStringValue_e StringVal String value Return Value True if option successfully set, false if not
Return Value True if option successfully set, false if not
VBA: See CWModelDoc::SetSimulationOptionStringValue .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWModelDoc.SetSimulationOptionToggle
Sets the boolean value for the specified user preference.
```vb
Function SetSimulationOptionToggle( ByVal NUserPreferenceValue As System.Integer, ByVal BValue As System.Boolean ) As System.Boolean
```
Parameters NUserPreferenceValue User preference as defined by swsUserPreferenceToggle_e BValue Boolean value Return Value True if option successfully set, false if not
Return Value True if option successfully set, false if not
VBA: See CWModelDoc::SetSimulationOptionToggle .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWModelDoc.StudyManager
Gets the study manager.
```vb
ReadOnly Property StudyManager As CWStudyManager
```
Property Value Study manager
VBA: See CWModelDoc::StudyManager .
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Materials

## ICWStudyManager

### ICWStudyManager.ActiveStudy
Gets or sets the index of the active study.
```vb
Property ActiveStudy As System.Integer
```
Property Value Index of the active study
VBA: See CWStudyManager::ActiveStudy .
Examples: Visual Basic for Applications; Calculate Edge Weld Results

### ICWStudyManager.ConvertStudy
Converts the specified static study to the specified target study.
```vb
Function ConvertStudy( ByVal SName As System.String, ByVal NSrcStudyType As System.Integer, ByVal SNewName As System.String, ByVal NTargetStudyType As System.Integer, ByVal SConfiguration As System.String, ByVal NTargetStudySubType As System.Integer, ByRef Errors As System.Integer ) As CWStudy
```
Parameters SName Name of static study to convert NSrcStudyType swsAnalysisStudyType_e .swsAnalysisStudyTypeStatic SNewName Name of target study to which to convert NTargetStudyType (see Remarks ) SConfiguration Name of model configuration NTargetStudySubType (see Remarks ) Errors Error as defined in swsStudyError_e Return Value ICWStudy
Return Value ICWStudy
VBA: See CWStudyManager::ConvertStudy .
Remarks: This method is valid only for SOLIDWORKS Simulation Premium. If NTargetStudyType is swsAnalysisStudyType_e ... Then NTargetStudySubType is... swsAnalysisStudyTypeDynamic As defined in swsDynamicAnalysisSubType_e swsAnalysisStudyTypeNonlinear 0 = Static 1 = Dynamic swsAnalysisStudyTypeStatic Ignored The following boundary conditions are transferred to the target study during conversion: Material properties Connections External loads Fixtures Mesh Study properties
Examples: Visual Basic for Applications; Example Convert Study

### ICWStudyManager.CreateNewStudy
Obsolete. Superseded by ICWStudyManager::CreateNewStudy2 .
```vb
Function CreateNewStudy( ByVal SName As System.String, ByVal NAnalysisType As System.Integer, ByVal NMeshType As System.Integer, ByRef Errors As System.Integer ) As CWStudy
```
Parameters SName New study name NAnalysisType Type of study as defined in swsAnalysisStudyType_e NMeshType Mesh type as defined in swsMeshType_e (see Remarks ) Errors Error as defined in swsStudyError_e Return Value Study
Return Value Study
VBA: See CWStudyManager::CreateNewStudy .
Remarks: Valid mesh types for different types of studies: Study Type Solid Mesh Shell Mesh Mixed Mesh (Solid & Shell) Beam Mesh Static Supported Supported Supported Supported Frequency Supported Supported Supported Supported Buckling Supported Supported Supported Supported Thermal Supported Supported Supported Supported Nonlinear Supported Supported Supported Not Supported Linear Dynamic Supported Supported Supported Not Supported Drop Test Supported Not Supported Not Supported Not Supported Fatigue Not Applicable Not Applicable Not Applicable Not Applicable Optimization Not Applicable Not Applicable Not Applicable Not Applicable NOTES: For documents with surface geometry only (no solids), only RefSurfShellElementMesh is supported. BeamElementMesh is supported only for static, buckling, and frequency studies. Drop test studies support solid mesh only. For optimization and fatigue studies do not have mesh on their own. They use mesh of referenced studies. Mesh type is set to SolidElementMesh.
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Materials

### ICWStudyManager.CreateNewStudy2
Obsolete. Superseded by ICWStudyManager::CreateNewStudy3 .
```vb
Function CreateNewStudy2( ByVal SName As System.String, ByVal NAnalysisType As System.Integer, ByRef Errors As System.Integer ) As CWStudy
```
Parameters SName New study name NAnalysisType Type of study as defined in swsAnalysisStudyType_e Errors Error as defined in swsStudyError_e Return Value Study
Return Value Study
VBA: See CWStudyManager::CreateNewStudy2 .
Examples: Visual Basic for Applications; Create Frequency Study with Mixed Mesh

### ICWStudyManager.CreateNewStudy3
Creates a new study.
```vb
Function CreateNewStudy3( ByVal SName As System.String, ByVal NAnalysisType As System.Integer, ByVal NStudySubOptions As System.Integer, ByRef Errors As System.Integer ) As CWStudy
```
Parameters SName New study name NAnalysisType Type of study as defined in swsAnalysisStudyType_e NStudySubOptions If NAnalysisType is ... Then NStudySubOptions is a sub-study as defined in ... swsAnalysisStudyType_e.swsAnalysisStudyTypeDynamic swsDynamicAnalysisSubType_e swsAnalysisStudyType_e.swsAnalysisStudyTypeFatigue swsFatigueStudySubOption_e Errors Error as defined in swsStudyError_e Return Value ICWStudy
Return Value ICWStudy
VBA: See CWStudyManager::CreateNewStudy3 .
Examples: Visual Basic for Applications; Example Create Linear Dynamic Study; Create Fatigue Study; Add Remote Load; Create Drop Test Study

### ICWStudyManager.DeleteStudy
Deletes the specified study.
```vb
Sub DeleteStudy( ByVal StudyName As System.String )
```
Parameters StudyName Study name
VBA: See CWStudyManager::DeleteStudy .
Examples: Visual Basic for Applications; Calculate Edge Weld Results

### ICWStudyManager.DuplicateStudy
Duplicates the specified study.
```vb
Function DuplicateStudy( ByVal SName As System.String, ByVal SNewName As System.String, ByVal SConfiguration As System.String, ByVal NNLDynamic As System.Integer, ByRef Errors As System.Integer ) As CWStudy
```
Parameters SName Name of study to duplicate SNewName Name of new study SConfiguration Name of configuration NNLDynamic 1 to copy to a nonlinear dynamic study, 0 to not Errors Error code as defined in swsDuplicateStudyError_e Return Value ICWStudy
Return Value ICWStudy
VBA: See CWStudyManager::DuplicateStudy .
Examples: Visual Basic for Applications; Example Duplicate Study

### ICWStudyManager.GetStudy
Gets the study at the specified index.
```vb
Function GetStudy( ByVal NIndex As System.Integer ) As CWStudy
```
Parameters NIndex 0-based index of the study to get Return Value Study
Return Value Study
VBA: See CWStudyManager::GetStudy .
Remarks: Before calling this method, call ICWStudyManager::ActiveStudy to set the index of the active study or ICWStudyManager::StudyCount to get a valid value for NIndex.
Examples: Visual Basic for Applications; Change Beam to Solid Body and Back

### ICWStudyManager.ManageDistributedSimulation
Obsolete. Superseded by ICWStudyManager::ManageDistributedSimulation2 .
```vb
Function ManageDistributedSimulation( ByVal BOffLoad As System.Integer, ByVal ComputerNames As System.Object ) As System.Integer
```
Parameters BOffLoad True to offload the coordinating computer's simulation processing to other computers on the network, false to let the coordinating computer participate in network simulation ComputerNames Array of names of computers participating in network simulation Return Value Error code as defined in swsDistributedSimulationError_e
Return Value Error code as defined in swsDistributedSimulationError_e
VBA: See CWStudyManager::ManageDistributedSimulation .
Remarks: This method is valid only for SOLIDWORKS Simulation Professional or Premium. Network simulation requires the Intel Direct Sparse solver. Before running network simulation: Set the solver type to swsSolverType_e .swsSolverTypeINTELCluster using one of: ICWBucklingStudyOptions::SolverType ICWDynamicStudyOptions::SolverType ICWFrequencyStudyOptions::SolverType ICWNonLinearStudyOptions::SolverType ICWStaticStudyOptions::SolverType ICWThermalStudyOptions::SolverType Call ICWStudyManager::SetDistributedSimulation to enable network simulation.
Examples: Visual Basic for Applications

### ICWStudyManager.ManageDistributedSimulation2
Sets up network simulation.
```vb
Function ManageDistributedSimulation2( ByVal BOffLoad As System.Boolean, ByVal ComputerNames As System.Object ) As System.Integer
```
Parameters BOffLoad -1 or true to offload the coordinating computer's simulation processing to other computers on the network, 0 or false to let the coordinating computer participate in network simulation ComputerNames Array of names of computers participating in network simulation Return Value Error code as defined in swsDistributedSimulationError_e
Return Value Error code as defined in swsDistributedSimulationError_e
VBA: See CWStudyManager::ManageDistributedSimulation2 .
Examples: Visual Basic for Applications

### ICWStudyManager.RenameStudyFromID
Renames the study that has the specified ID.
```vb
Function RenameStudyFromID( ByVal NIndex As System.Integer, ByVal SName As System.String ) As System.Integer
```
Parameters NIndex ID of study to rename SName New name of study Return Value Error as defined by swsStudyError_e
Return Value Error as defined by swsStudyError_e
VBA: See CWStudyManager::RenameStudyFromID .
Examples: Visual Basic for Applications

### ICWStudyManager.RenameStudyFromName
Renames the study that has the specified name.
```vb
Function RenameStudyFromName( ByVal SOldName As System.String, ByVal SNewName As System.String ) As System.Integer
```
Parameters SOldName Old name of study SNewName New name of study Return Value Error as defined by swsStudyError_e
Return Value Error as defined by swsStudyError_e
VBA: See CWStudyManager::RenameStudyFromName .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWStudyManager.RunAllStudies
Runs all studies in batch mode.
```vb
Function RunAllStudies( ByRef ErrorCode As System.Integer ) As CWRunStudiesResults
```
Parameters ErrorCode Error code as defined in swsRunStudiesErrorCode_e Return Value ICWRunStudiesResults
Return Value ICWRunStudiesResults
VBA: See CWStudyManager::RunAllStudies .
Remarks: Before calling this method, set the run and mesh options for the studies using ICWStudyManager::RunSpecifiedStudyOptions .
Examples: Visual Basic for Applications

### ICWStudyManager.RunSpecifiedStudyByName
Runs specified studies in batch mode.
```vb
Function RunSpecifiedStudyByName( ByRef ErrorCode As System.Integer ) As CWRunStudiesResults
```
Parameters ErrorCode Error code as defined in swsRunStudiesErrorCode_e Return Value ICWRunStudiesResults
Return Value ICWRunStudiesResults
VBA: See CWStudyManager::RunSpecifiedStudyByName .
Remarks: Before calling this method, set the run and mesh options for the study using ICWStudyManager::RunSpecifiedStudyOptions .
Examples: Visual Basic for Applications; Example Run Studies in Batch Mode

### ICWStudyManager.RunSpecifiedStudyOptions
Gets the run and mesh option for studies run in batch mode.
```vb
ReadOnly Property RunSpecifiedStudyOptions As CWRunSpecStudiesRunMeshOptions
```
Property Value ICWRunSpecStudiesRunMeshOptions
VBA: See CWStudyManager::RunSpecifiedStudyOptions .
Examples: Visual Basic for Applications; Example Run Studies in Batch Mode

### ICWStudyManager.SetDistributedSimulation
Obsolete. Superseded by ICWStudyManager::SetDistributedSimulation2 .
```vb
Function SetDistributedSimulation( ByVal BNetSim As System.Integer ) As System.Integer
```
Parameters BNetSim True to enable network simulation, false to disable it Return Value Error code as defined in swsDistributedSimulationError_e
Return Value Error code as defined in swsDistributedSimulationError_e
VBA: See CWStudyManager::SetDistributedSimulation .
Remarks: This method is valid only for SOLIDWORKS Simulation Professional or Premium.
Examples: Visual Basic for Applications

### ICWStudyManager.SetDistributedSimulation2
Sets whether to enable network simulation.
```vb
Function SetDistributedSimulation2( ByVal BNetSim As System.Boolean ) As System.Integer
```
Parameters BNetSim -1 or true to enable network simulation, 0 or false to disable it Return Value Error code as defined in swsDistributedSimulationError_e
Return Value Error code as defined in swsDistributedSimulationError_e
VBA: See CWStudyManager::SetDistributedSimulation2 .
Examples: Visual Basic for Applications

### ICWStudyManager.StudyCount
Gets the number of studies in the open document.
```vb
ReadOnly Property StudyCount As System.Integer
```
Property Value Number of studies
VBA: See CWStudyManager::StudyCount .
Examples: Visual Basic for Applications

## ICWStudy

### ICWStudy.ActivateConfiguration
Activates the configuration of this study.
```vb
Sub ActivateConfiguration()
```
VBA: See CWStudy::ActivateConfiguration .
Remarks: If the model has studies that use different configurations, you must set the active study with ICWStudyManager::ActiveStudy after calling this method.
Examples: Visual Basic for Applications; Example Create Trend Tracker

### ICWStudy.AddDropTestSetup
Adds a setup to this drop test study.
```vb
Function AddDropTestSetup( ByVal EntityForGravity As System.Object, ByRef ErrorCode As System.Integer ) As CWDropTestSetup
```
Parameters EntityForGravity Entity to drop ErrorCode Status as defined by swsDropTestSetUpError_e Return Value ICWDropTestSetup
Return Value ICWDropTestSetup
VBA: See CWStudy::AddDropTestSetup .
Examples: Visual Basic for Applications; Example Create Drop Test Study

### ICWStudy.AnalysisType
Gets the type of the analysis of the active study.
```vb
ReadOnly Property AnalysisType As System.Integer
```
Property Value Type of study as defined in swsAnalysisStudyType_e
VBA: See CWStudy::AnalysisType .
Remarks: The proper license must be available to create the specified type of analysis. Only static, frequency, buckling, thermal, and nonlinear studies are currently supported. Fatigue, optimization, and pressure vessel studies are not currently supported.
Examples: Visual Basic for Applications

### ICWStudy.BeamManager
Gets the beam manager.
```vb
ReadOnly Property BeamManager As CWBeamManager
```
Property Value Beam manager
VBA: See CWStudy::BeamManager .
Examples: Visual Basic for Applications; Get and Set Beams and Joints

### ICWStudy.BucklingStudyOptions
Gets the options for this buckling study.
```vb
ReadOnly Property BucklingStudyOptions As CWBucklingStudyOptions
```
Property Value Buckling study options
VBA: See CWStudy::BucklingStudyOptions .
Examples: Visual Basic for Applications; Example Create Buckling Study

### ICWStudy.ConfigurationName
Gets the name of the study's configuration.
```vb
ReadOnly Property ConfigurationName As System.String
```
Property Value Configuration name
VBA: See CWStudy::ConfigurationName .
Examples: Visual Basic for Applications; Example Create Linear Dynamic Study

### ICWStudy.ContactManager
Gets the contact manager.
```vb
ReadOnly Property ContactManager As CWContactManager
```
Property Value Contact manager
VBA: See CWStudy::ContactManager .
Examples: Visual Basic for Applications; Create Frequency Study with Mixed Mesh

### ICWStudy.CopyMeshFromStudy
Copies the mesh from the specified study to this study.
```vb
Function CopyMeshFromStudy( ByVal SStudyName As System.String ) As System.Integer
```
Parameters SStudyName Name of study from which to copy a mesh Return Value Error code as defined in swsMeshCopyErrorCode_e
Return Value Error code as defined in swsMeshCopyErrorCode_e
VBA: See CWStudy::CopyMeshFromStudy .
Examples: Visual Basic for Applications; Example Copy Mesh and Generate Report

### ICWStudy.CreateImportedResultsPlot
Creates a plot for this study using the specified external results file.
```vb
Function CreateImportedResultsPlot( ByVal SExternalResultFile As System.String, ByVal SLegendTitle As System.String, ByVal BValueByElem As System.Boolean, ByRef ErrorCode As System.Integer ) As CWPlot
```
Parameters SExternalResultFile Full pathname of the *.csv results file SLegendTitle Legend title of the plot BValueByElem True if SExternalResultFile contains element data, false if it contains nodal data ErrorCode Error code as defined in swsResultPlotErrorCode_e Return Value ICWPlot
Return Value ICWPlot
VBA: See CWStudy::CreateImportedResultsPlot .
Examples: Visual Basic for Applications

### ICWStudy.CreateMesh
Creates the mesh.
```vb
Function CreateMesh( ByVal NUnits As System.Integer, ByVal DElementSize As System.Double, ByVal DTolerance As System.Double ) As System.Integer
```
Parameters NUnits Linear units as defined in swsLinearUnit_e DElementSize Average global element size DTolerance Tolerance Return Value Error as defined in swsStudyMeshError_e
Return Value Error as defined in swsStudyMeshError_e
VBA: See CWStudy::CreateMesh .
Examples: Visual Basic for Applications; Analyze Part; Create Curvature-based Mesh

### ICWStudy.CreateMeshForMeshControl
Creates a mesh for the specified mesh control.
```vb
Function CreateMeshForMeshControl( ByVal SMeshControlName As System.String ) As System.Integer
```
Parameters SMeshControlName Name of the mesh control Return Value Error as defined in swsMeshControlMeshingError_e
Return Value Error as defined in swsMeshControlMeshingError_e
VBA: See CWStudy::CreateMeshForMeshControl .
Remarks: To populate SMeshControlName, call: ICWMesh::GetMeshControlAt ICWMeshControl::Name
Examples: Visual Basic for Applications; Example Apply Mesh Control

### ICWStudy.CreateTrendTracker
Creates a Trend Tracker for this study.
```vb
Function CreateTrendTracker( ByRef ErrorCode As System.Integer ) As CWTrendTracker
```
Parameters ErrorCode Error code as defined in swsTrendTrackerErrorCode_e Return Value ICWTrendTracker
Return Value ICWTrendTracker
VBA: See CWStudy::CreateTrendTracker .
Examples: Visual Basic for Applications; Example Create Trend Tracker

### ICWStudy.DampingOptions
Gets the damping options for this study.
```vb
ReadOnly Property DampingOptions As CWDampingOptions
```
Property Value ICWDampingOptions
VBA: See CWStudy::DampingOptions .
Examples: Visual Basic for Applications; Example Create Linear Dynamic Study

### ICWStudy.DeleteTrendTracker
Deletes the Trend Tracker from this study.
```vb
Function DeleteTrendTracker() As System.Integer
```
Return Value Deletion status as defined in swsTrendTrackerErrorCode_e
VBA: See CWStudy::DeleteTrendTracker .
Examples: Visual Basic for Applications; Example Create Trend Tracker

### ICWStudy.DropTestResultOptions
Gets the result options for this drop test study.
```vb
ReadOnly Property DropTestResultOptions As CWDropTestResultOptions
```
Property Value ICWDropTestResultOptions
VBA: See CWStudy::DropTestResultOptions .
Examples: Visual Basic for Applications; Example Create Drop Test Study

### ICWStudy.DropTestSetup
Gets the setup for this drop test study.
```vb
ReadOnly Property DropTestSetup As CWDropTestSetup
```
Property Value ICWDropTestSetup
VBA: See CWStudy::DropTestSetup .
Examples: Visual Basic for Applications; Create Drop Test Study

### ICWStudy.DropTestStudyOptions
Gets the options for this drop test study.
```vb
ReadOnly Property DropTestStudyOptions As CWDropTestStudyOptions
```
Property Value ICWDropTestStudyOptions
VBA: See CWStudy::DropTestStudyOptions .
Examples: Visual Basic for Applications; Example Create Drop Test Study

### ICWStudy.DynamicAnalysisSubType
Gets the type of this linear dynamic analysis.
```vb
ReadOnly Property DynamicAnalysisSubType As System.Integer
```
Property Value Type of linear dynamic analysis as defined in swsDynamicAnalysisSubType_e
VBA: See CWStudy::DynamicAnalysisSubType .
Examples: Visual Basic for Applications; Example Create Linear Dynamic Study

### ICWStudy.DynamicStudyOptions
Gets the options for this dynamic study.
```vb
ReadOnly Property DynamicStudyOptions As CWDynamicStudyOptions
```
Property Value Dynamic study options
VBA: See CWStudy::DynamicStudyOptions .
Examples: Visual Basic for Applications; Create Linear Dynamic Study

### ICWStudy.ExportSimulationStudy
Exports this study to the specified finite-element analysis program.
```vb
Function ExportSimulationStudy( ByVal SLocationPath As System.String, ByVal SFileName As System.String, ByVal NFormat As System.Integer, ByVal NNodeOffset As System.Integer, ByVal NElementOffset As System.Integer, ByVal NOption As System.Integer, ByVal NUnit As System.Integer ) As System.Integer
```
Parameters SLocationPath Path to which to export this study SFileName Name of file to which to export this study NFormat If nOption is swsStudyExportOption_e... nFormat is... swsStudyExport_Cosmos Export option as defined by swsCosmosExportOption_e swsStudyExport_Nastran Format in which to export this study as defined by swsNastranExportOption_e swsStudyExport_Abacus, swsStudyExport_Ansys, swsStudyExport_Exodus, swsStudyExport_IdeasUniversal, or swsStudyExport_PatranNeutral Not valid NNodeOffset Starting node for exported data; the lowest node label in the generated file is NNodeOffset + 1 NElementOffset Starting element for exported data; the lowest element label in the generated file is NElementOffset + 1 NOption Finite-element analysis program to which to export as defined by swsStudyExportOption_e NUnit If nOption is swsStudyExportOption_e... nUnit is... swsStudyExport_Cosmos Defined by swsLinearUnit_e swsStudyExport_Nastran Defined by swsNastranExportUnit_e swsStudyExport_Abacus, swsStudyExport_Ansys, swsStudyExport_Exodus, swsStudyExport_IdeasUniversal, or swsStudyExport_PatranNeutral Not valid Return Value Error as defined in swsStudyExportError_e
Return Value Error as defined in swsStudyExportError_e
VBA: See CWStudy::ExportSimulationStudy .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWStudy.FatigueStudyOptions
Gets the options for this fatigue study.
```vb
ReadOnly Property FatigueStudyOptions As CWFatigueStudyOptions
```
Property Value ICWFatigueStudyOptions
VBA: See CWStudy::FatigueStudyOptions .
Examples: Visual Basic for Applications; Create Fatigue Study

### ICWStudy.FrequencyStudyOptions
Gets the options for this frequency study.
```vb
ReadOnly Property FrequencyStudyOptions As CWFrequencyStudyOptions
```
Property Value Frequency study options
VBA: See CWStudy::FrequencyStudyOptions .
Examples: Visual Basic for Applications; Example Create Frequency Study with Solid Mesh

### ICWStudy.GenerateReport
Creates a report about all aspects of this study.
```vb
Function GenerateReport( ByVal SReportPath As System.String, ByVal SDocName As System.String, ByVal BShowOnPublish As System.Boolean ) As System.Integer
```
Parameters SReportPath Path in which to create this report SDocName File name of the report BShowOnPublish True to display the report after creation, false to not Return Value Error code as defined in swsGenerateReportError_e
Return Value Error code as defined in swsGenerateReportError_e
VBA: See CWStudy::GenerateReport .
Examples: Visual Basic for Applications; Example Copy Mesh and Generate Report

### ICWStudy.GetClusterComputers
Gets the list of names of the computers participating in this study's distributed simulation.
```vb
Function GetClusterComputers( ByRef ComputerNames As System.Object ) As System.Integer
```
Parameters ComputerNames Array of computer names Return Value Error code as defined in swsDistributedSimulationError_e
Return Value Error code as defined in swsDistributedSimulationError_e
VBA: See CWStudy::GetClusterComputers .
Remarks: This method is valid only if: The license is SOLIDWORKS Simulation Professional or Premium. ICWStudy::IsSimulationDistributable is true. ICWStudy::OffLoad is true.
Examples: Visual Basic for Applications

### ICWStudy.GetMassPropertiesManager
Gets the mass properties manager.
```vb
Function GetMassPropertiesManager() As CWMassPropertiesManager
```
Return Value ICWMassPropertiesManager
VBA: See CWStudy::GetMassPropertiesManager .
Examples: Visual Basic for Applications; Example Get Mass Properties

### ICWStudy.GetTimeTakenForGapIterationsAndContactOperations
Gets the time taken for gap iterations and contact operations in this nonlinear study.
```vb
Function GetTimeTakenForGapIterationsAndContactOperations() As System.Integer
```
Return Value Time taken in seconds for gap iterations and contact operations in this nonlinear study
VBA: See CWStudy::GetTimeTakenforGapIterationsAndContactOperations .
Remarks: This method is valid only for nonlinear studies. model_name - study_name .out is created after the study is analyzed and contains this time taken for gap iterations and contact operations. Check the location of this file in the results folder that is configered in Simulation > Options > Default Options .
Examples: Visual Basic for Applications

### ICWStudy.GetTimeTakenForInputDataTransferFromDatabase
Gets the time taken for input data transfer from the database in this nonlinear study.
```vb
Function GetTimeTakenForInputDataTransferFromDatabase() As System.Integer
```
Return Value Time taken in seconds for input data transfer from the database in this nonlinear study
VBA: See CWStudy::GetTimeTakenForInputDataTransferFromDatabase .
Remarks: This method is valid only for nonlinear studies. model_name - study_name .out is created after the study is analyzed and contains this time taken for input data transfer from the database. Check the location of this file in the results folder that is configered in Simulation > Options > Default Options .
Examples: Visual Basic for Applications

### ICWStudy.GetTimeTakenForInputPhase
Gets the time taken for the input phase of this static study.
```vb
Function GetTimeTakenForInputPhase() As System.Integer
```
Return Value Time taken in seconds for the input phase of this static study
VBA: See CWStudy::GetTimeTakenForInputPhase .
Remarks: This method is valid only for static studies. model_name - study_name .out is created after the study is analyzed and contains this time taken for the input phase. Check the location of this file in the results folder that is configered in Simulation > Options > Default Options .
Examples: Visual Basic for Applications; Example Create Plots for Static Study

### ICWStudy.GetTotalSolutionTime
Gets the total solver time for this study.
```vb
Function GetTotalSolutionTime() As System.Integer
```
Return Value Total solver time in seconds for this study
VBA: See CWStudy::GetTotalSolutionTime .
Remarks: This method is valid for both static and nonlinear studies. model_name - study_name .out is created after the study is analyzed and contains this total solver time. Check the location of this file in the results folder that is configered in Simulation > Options > Default Options .
Examples: Visual Basic for Applications; Example Create Plots for Static Study

### ICWStudy.ImportStudyFeatures
Imports the study features of the specified component's study into this study.
```vb
Function ImportStudyFeatures( ByVal SComponent As System.String, ByVal SStudy As System.String, ByVal SConfig As System.String, ByVal NFilter As System.Integer, ByVal BPropagate As System.Boolean, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters SComponent Name of the assembly component from whose study to import study features SStudy Name of SComponent's study from which to import study features SConfig Name of SComponent's configuration NFilter Study features to import as defined in swsImportStudyFeaturesFilterType_e BPropagate True to propagate imported study features to components in this study that are similar to SComponent, false to not ErrorCode Result code as defined in swsImportStudyFeaturesErrorCode_e Return Value VARIANT_TRUE (-1) if successful, VARIANT_FALSE (0) if not
Return Value VARIANT_TRUE (-1) if successful, VARIANT_FALSE (0) if not
VBA: See CWStudy::ImportStudyFeatures .
Examples: Visual Basic for Applications

### ICWStudy.IsSimulationDistributable
Gets whether this study supports network simulation.
```vb
ReadOnly Property IsSimulationDistributable As System.Boolean
```
Property Value True if this study supports network simulation, false if not
VBA: See CWStudy::IsSimulationDistributable .
Remarks: This method is valid only for SOLIDWORKS Simulation Professional or Premium.
Examples: Visual Basic for Applications

### ICWStudy.LoadCaseManager
Gets the Load Case Manager for this static study.
```vb
ReadOnly Property LoadCaseManager As CWLoadCaseManager
```
Property Value ICWLoadCaseManager
VBA: See CWStudy::LoadCaseManager .
Examples: Visual Basic for Applications; Example Access Load Case Manager

### ICWStudy.LoadsAndRestraintsManager
Gets the loads and restraints manager.
```vb
ReadOnly Property LoadsAndRestraintsManager As CWLoadsAndRestraintsManager
```
Property Value Loads and restraints manager
VBA: See CWStudy::LoadsAndRestraintsManager .
Examples: Visual Basic for Applications; Example Create and Edit Bolt and Pin Connectors

### ICWStudy.Mesh
Gets the mesh of this study.
```vb
ReadOnly Property Mesh As CWMesh
```
Property Value Mesh
VBA: See CWStudy::Mesh .
Examples: Visual Basic for Applications; Analyze Part; Create Curvature-based Mesh

### ICWStudy.MeshAndRun
Meshes the model and analyzes the study.
```vb
Sub MeshAndRun()
```
VBA: See CWStudy::MeshAndRun .
Examples: Visual Basic for Applications; Example Create Trend Tracker

### ICWStudy.MeshType
Gets the mesh type of this study.
```vb
ReadOnly Property MeshType As System.Integer
```
Property Value Mesh type as defined in swsMeshType_e (see Remarks )
VBA: See CWStudy::MeshType .
Remarks: Valid mesh types for different types of studies: Study Type Solid Mesh Shell Mesh Mixed Mesh (Solid & Shell) Beam Mesh Static Supported Supported Supported Supported Frequency Supported Supported Supported Supported Buckling Supported Supported Supported Supported Thermal Supported Supported Supported Supported Nonlinear Supported Supported Supported Not Supported Linear Dynamic Supported Supported Supported Not Supported Drop Test Supported Not Supported Not Supported Not Supported Fatigue Not Applicable Not Applicable Not Applicable Not Applicable Optimization Not Applicable Not Applicable Not Applicable Not Applicable NOTES: For documents with surface geometry only (no solids), only RefSurfShellElementMesh is supported. BeamElementMesh is supported only for static, buckling, and frequency studies. Drop test studies support solid mesh only. Optimization and fatigue studies do not have their own mesh. They use the mesh of referenced studies. Mesh type is set to SolidElementMesh.
Examples: Visual Basic for Applications

### ICWStudy.MultipleComponentContactsEditManager
Gets the Multiple Component Contacts Edit Manager which provides the simultaneous editing of multiple component contacts.
```vb
ReadOnly Property MultipleComponentContactsEditManager As CWMultipleComponentContactsEditManager
```
Property Value ICWMultipleComponentContactsEditManager
VBA: See CWStudy::MultipleComponentContactsEditManager .
Examples: Visual Basic for Applications

### ICWStudy.MultipleContactSetsEditManager
Gets the Multiple Contact Sets Edit Manager which provides the simultaneous editing of multiple contact sets.
```vb
ReadOnly Property MultipleContactSetsEditManager As CWMultipleContactSetsEditManager
```
Property Value ICWMultipleContactSetsEditManager
VBA: See CWStudy::MultipleContactSetsEditManager .
Examples: Visual Basic for Applications

### ICWStudy.Name
Gets the name of this study.
```vb
ReadOnly Property Name As System.String
```
Property Value Name of the active study
VBA: See CWStudy::Name .
Examples: Visual Basic for Applications

### ICWStudy.NonLinearStudyOptions
Gets the options for this nonlinear study.
```vb
ReadOnly Property NonLinearStudyOptions As CWNonLinearStudyOptions
```
Property Value Nonlinear study options
VBA: See CWStudy::NonLinearStudyOptions .
Examples: Visual Basic for Applications

### ICWStudy.NonlinearAnalysisSubType
Gets the subtype of this nonlinear study.
```vb
ReadOnly Property NonlinearAnalysisSubType As System.Integer
```
Property Value Nonlinear subtype as defined in swsNonlinearAnalysisSubType_e
VBA: See CWStudy::NonlinearAnalysisSubType .
Examples: Visual Basic for Applications

### ICWStudy.OffLoad
Gets or sets whether to offload the coordinating computer's simulation processing to other computers on the network.
```vb
Property OffLoad As System.Boolean
```
Property Value True to offload the coordinating computer's simulation processing to other computers on the netowrk, false to let the coordinating computer participate in network simulation
VBA: See CWStudy::OffLoad .
Remarks: This method is valid only if: The license is SOLIDWORKS Simulation Professional or Premium. ICWStudy::IsSimulationDistributable is true.
Examples: Visual Basic for Applications

### ICWStudy.Results
Gets the results of this study.
```vb
ReadOnly Property Results As CWResults
```
Property Value Results
VBA: See CWStudy::Results .
Examples: Visual Basic for Applications; Analzye Part

### ICWStudy.RunAnalysis
Runs this study.
```vb
Function RunAnalysis() As System.Integer
```
Return Value Error as defined in swsRunAnalysisError_e
VBA: See CWStudy::RunAnalysis .
Remarks: See Also ICWStudy Interface ICWStudy Members ICWStudy::AnalysisType Property ICWStudy::MeshAndRun Method ()
Examples: Visual Basic for Applications; Example Analyze Part

### ICWStudy.SetClusterComputers
Sets the list of names of the computers participating in this study's distributed simulation.
```vb
Function SetClusterComputers( ByVal ComputerNames As System.Object ) As System.Integer
```
Parameters ComputerNames Array of computer names Return Value Error code as defined in swsDistributedSimulationError_e
Return Value Error code as defined in swsDistributedSimulationError_e
VBA: See CWStudy::SetClusterComputers .
Remarks: This method is valid only if: The license is SOLIDWORKS Simulation Professional or Premium. ICWStudy::IsSimulationDistributable is true. ICWStudy::OffLoad is true.
Examples: Visual Basic for Applications

### ICWStudy.SetFatigueResultOptions
Sets the result options for this fatigue study.
```vb
Function SetFatigueResultOptions( ByVal NFatigueCalculationsOption As System.Integer, ByVal DispArray As System.Object ) As System.Integer
```
Parameters NFatigueCalculationsOption Fatigue calculation option as defined in swsFatigueCalculationsOption_e DispArray Array of vertices and reference points at which to plot damage matrix charts; only valid in variable amplitude fatigue studies Return Value 0 if successful, 1 if not
Return Value 0 if successful, 1 if not
VBA: See CWStudy::SetFatigueResultOptions .
Examples: Visual Basic for Applications; Example Create Fatigue Study

### ICWStudy.SetNonLinearStudyOptions
Sets the nonlinear study options.
```vb
Function SetNonLinearStudyOptions( ByVal NonLinearOption As System.Object ) As System.Integer
```
Parameters NonLinearOption Nonlinear study options Return Value Error defined in swsNonLinearStudyOptionsError_e
Return Value Error defined in swsNonLinearStudyOptionsError_e
VBA: See CWStudy::SetNonLinearStudyOptions .
Examples: Visual Basic for Applications

### ICWStudy.SetResultCombinationSetup
Sets the result combination setup for this Pressure Vessel Design study.
```vb
Function SetResultCombinationSetup( ByVal NOption As System.Integer, ByVal NItems As System.Integer, ByVal MultiplicationFactors As System.Object, ByVal StudyNames As System.Object ) As System.Integer
```
Parameters NOption Type of result combination as defined in swsPVResultCombinationType_e NItems Number of elements in the MultiplicationFactors and StudyNames arrays MultiplicationFactors Array of multiplication factors; valid only if NOption = swsPVResultCombinationType_e.swsPVResultCombinationType_Linear StudyNames Array of names of the static studies whose results you want to combine; if NOption = swsPVResultCombinationType_e.swsPVResultCombinationType_Linear, the elements of MultiplicationFactors are coefficients of the linear combination of the static study results Return Value Error as defined in swsPVResultCombinationError_e
Return Value Error as defined in swsPVResultCombinationError_e
VBA: See CWStudy::SetResultCombinationSetup .
Examples: Visual Basic for Applications

### ICWStudy.ShellManager
Gets the shell manager.
```vb
ReadOnly Property ShellManager As CWShellManager
```
Property Value Shell manager
VBA: See CWStudy::ShellManager .
Examples: Visual Basic for Applications; Create Frequency Study with Mixed Mesh

### ICWStudy.ShowOrHideFixtures
Shows or hides fixture symbols in the graphics area.
```vb
WriteOnly Property ShowOrHideFixtures As System.Boolean
```
Property Value True to show fixture symbols in the study, false to hide them
VBA: See CWStudy::ShowOrHideFixtures .
Examples: Visual Basic for Applications; Analyze Part

### ICWStudy.ShowOrHideForce
Shows or hides the force symbols in the graphics area.
```vb
WriteOnly Property ShowOrHideForce As System.Boolean
```
Property Value True to show force symbols in the study, false to hide them
VBA: See CWStudy::ShowOrHideForces .
Examples: Visual Basic for Applications; Analyze Part

### ICWStudy.SolidManager
Gets the solid manager.
```vb
ReadOnly Property SolidManager As CWSolidManager
```
Property Value Solid manager
VBA: See CWStudy::SolidManager .
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Materials

### ICWStudy.StaticStudyOptions
Gets the options for this static study.
```vb
ReadOnly Property StaticStudyOptions As CWStaticStudyOptions
```
Property Value Static study options
VBA: See CWStudy::StaticStudyOptions .
Examples: Visual Basic for Applications

### ICWStudy.StudyResultOptions
Gets the result options for this study.
```vb
ReadOnly Property StudyResultOptions As CWStudyResultOptions
```
Property Value Study result options
VBA: See CWStudy::StudyResultOptions .
Examples: Visual Basic for Applications; Create Linear Dynamic Study

### ICWStudy.SwxDRDOneClickWorkflowInfo
*Internal use only*
```vb
Property SwxDRDOneClickWorkflowInfo As System.String
```

### ICWStudy.ThermalStudyOptions
Gets the options for this thermal study.
```vb
ReadOnly Property ThermalStudyOptions As CWThermalStudyOptions
```
Property Value Thermal study options
VBA: See CWStudy::ThermalStudyOptions .
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWStudy.TopologyStudyManager
Gets the manager of this topology study's goals, constraints, and manufacturing controls.
```vb
ReadOnly Property TopologyStudyManager As CWTopologyStudyManager
```
Property Value ICWTopologyStudyManager
VBA: See CWStudy::TopologyStudyManager .
Examples: Visual Basic for Applications; Example Create Topology Study

### ICWStudy.TopologyStudyOptions
Gets this topology study's options.
```vb
ReadOnly Property TopologyStudyOptions As CWTopologyStudyOptions
```
Property Value ICWTopologyStudyOptions
VBA: See CWStudy::TopologyStudyOptions .
Remarks: This property is only valid for topology studies.
Examples: Visual Basic for Applications; Example Create Topology Study

### ICWStudy.UpdateAllComponents
Updates all components in the active study
```vb
Sub UpdateAllComponents()
```
VBA: See CWStudy::UpdateAllComponents .
Remarks: Use this method to update the study if the status of any components changed after creating the study.
Examples: Visual Basic for Applications

## ICWThermalStudyOptions

### ICWThermalStudyOptions.CheckFlowConvectionCoef
Obsolete. Superseded by ICWThermalStudyOptions::CheckFlowConvectionCoef2 .
```vb
Property CheckFlowConvectionCoef As System.Integer
```
Property Value 1 to include fluid convection effects from SOLDWORKS Flow Simulation, 0 to not
VBA: See CWThermalStudyOptions::CheckFlowConvectionCoef .
Remarks: After setting this property to 1, specify ICWThermalStudyOptions::FlowConvectionCoefFile with the result file generated by SOLIDWORKS Flow Simulation ( *.fld ). This property is not valid for 2D simplification studies.
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.CheckFlowConvectionCoef2
Gets or sets whether to include fluid convection effects from SOLIDWORKS Flow Simulation.
```vb
Property CheckFlowConvectionCoef2 As System.Boolean
```
Property Value -1 or true to include fluid convection effects from SOLDWORKS Flow Simulation, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer. After setting this property to -1, specify ICWThermalStudyOptions::FlowConvectionCoefFile with the result file generated by SOLIDWORKS Flow Simulation ( *.fld ). This property is not valid for 2D simplification studies.

### ICWThermalStudyOptions.ConvergenceTolerance
Gets or sets the convergence tolerance.
```vb
Property ConvergenceTolerance As System.Double
```
Property Value Convergence tolerance
VBA: See CWThermalStudyOptions::ConvergenceTolerance .
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.EMail
Obsolete. Superseded by ICWThermalStudyOptions::Email2 .
```vb
Property EMail As System.Integer
```
Property Value 1 to email notifications during simulations, 0 to not
VBA: See CWThermalStudyOptions::EMail .
Remarks: If this property is set to 1: use ICWThermalStudyOptions::EMailTo to specify the recipient of notifications. use ICWThermalStudyOptions::EMailTimebased to send email notifications during simulations.
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.EMail2
Gets or sets whether to email notifications during simulations.
```vb
Property EMail2 As System.Boolean
```
Property Value -1 or true to email notifications during simulations, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer. If this property is set to -1: use ICWThermalStudyOptions::EMailTo to specify the recipient of notifications. use ICWThermalStudyOptions::EMailTimebased2 to send email notifications during simulations.

### ICWThermalStudyOptions.EMailInterval
Gets or sets the time interval for sending email notifications during simulations.
```vb
Property EMailInterval As System.Integer
```
Property Value Time interval
VBA: See CWThermalStudyOptions::EMailInterval .
Remarks: This property is: set in units specified by ICWThermalStudyOptions::EMailIntervalUnit . valid only if ICWThermalStudyOptions::EMailTimebased is set to 1, and ICWThermalStudyOptions::EMail is set to 1.
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.EMailIntervalUnit
Gets or sets the units of time for ICWThermalStudyOptions::EMailInterval .
```vb
Property EMailIntervalUnit As System.Integer
```
Property Value Time units as defined in swsTimeUnits_e : swsSecond swsMinute swsHour
VBA: See CWThermalStudyOptions::EMailIntervalUnit .
Remarks: This property is valid only if: ICWThermalStudyOptions::EMailTimebased is set to 1, - and - ICWThermalStudyOptions::EMail is set to 1.
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.EMailTimebased
Obsolete. Superseded by ICWThermalStudyOptions::EmailTimebased2 .
```vb
Property EMailTimebased As System.Integer
```
Property Value 1 to send email notifications during simulations, 0 to not
VBA: See CWThermalStudyOptions::EMailTimebased .
Remarks: This property is valid only if ICWThermalStudyOptions::EMail is set to 1. If this property is set to 1, use ICWThermalStudyOptions::EMailInterval and ICWThermalStudyOptions::EMailIntervalUnit to specify the time interval for sending email notifications during simulations.
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.EMailTimebased2
Gets or sets whether to send email notifications during simulations.
```vb
Property EMailTimebased2 As System.Boolean
```
Property Value -1 or true to send email notifications during simulations, 0 or false to not
Remarks: This property is valid only if ICWThermalStudyOptions::EMail is set to -1. This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer. If this property is set to -1, use ICWThermalStudyOptions::EMailInterval and ICWThermalStudyOptions::EMailIntervalUnit to specify the time interval for sending email notifications during simulations.

### ICWThermalStudyOptions.EMailTo
Gets or sets the recipient of email notifications.
```vb
Property EMailTo As System.String
```
Property Value Email address
VBA: See CWThermalStudyOptions::EMailTo .
Remarks: This property is valid only if ICWThermalStudyOptions::EMail is set to 1.
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.FixedValue
Gets or sets the fixed value of the relaxation factor.
```vb
Property FixedValue As System.Double
```
Property Value Fixed under-relaxation factor
VBA: See CWThermalStudyOptions::FixedValue .
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.FlowConvectionCoefFile
Gets or sets the result file generated by SOLIDWORKS Flow Simulation.
```vb
Property FlowConvectionCoefFile As System.String
```
Property Value Full path name of the result file generated by SOLIDWORKS Flow Simulation ( *.fld )
VBA: See CWThermalStudyOptions::FlowConvectionCoefFile .
Remarks: This property is valid only if ICWThermalStudyOptions::CheckFlowConvectionCoef is set to 1.
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.RelaxationFactor
Gets or sets the status of the relaxation factor.
```vb
Property RelaxationFactor As System.Integer
```
Property Value Thermal relaxation factor as defined in swsThermalRelaxationFactor_e
VBA: See CWThermalStudyOptions::RelaxationFactor .
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.ResultFolder
Gets or sets the path to the folder for storing the study results.
```vb
Property ResultFolder As System.String
```
Property Value Path for the results folder
VBA: See CWThermalStudyOptions::ResultFolder .
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.SolutionType
Gets or sets the thermal solution type.
```vb
Property SolutionType As System.Integer
```
Property Value Thermal solution as defined in swsThermalSolutionType_e
VBA: See CWThermalStudyOptions::SolutionType .
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.SolverType
Gets or sets the solver type associated with the study.
```vb
Property SolverType As System.Integer
```
Property Value Solver type as defined in swsSolverType_e (see Remarks )
VBA: See CWThermalStudyOptions::SolverType .
Remarks: The automatic option is not supported for thermal studies; it supports static, frequency, and buckling studies only.
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.ThermalStudyNameUsedForInitialTemperature
Gets or sets the thermal study name for initial temperature from the thermal study. Used for thermal transient studies only.
```vb
Property ThermalStudyNameUsedForInitialTemperature As System.String
```
Property Value Thermal study name
VBA: See CWThermalStudyOptions::ThermalStudyNameUsedForInitialTemperature .
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.TimeIncrement
Gets or sets the time increment. Used for thermal transient studies only.
```vb
Property TimeIncrement As System.Double
```
Property Value Time increment
VBA: See CWThermalStudyOptions::TimeIncrement .
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.TimeSteps
Gets or sets the time step at which you want to use the temperature profile as an initial condition. Used for thermal transient studies only.
```vb
Property TimeSteps As System.Integer
```
Property Value Time step
VBA: See CWThermalStudyOptions::TimeSteps .
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.TotalTime
Gets or sets the total time. Used for thermal transient studies only.
```vb
Property TotalTime As System.Double
```
Property Value Total time
VBA: See CWThermalStudyOptions::TotalTime .
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.UseTemperatureFromThermalStudy
Obsolete. Superseded by ICWThermalStudyOptions::UseTemperatureFromThermalStudy2 .
```vb
Property UseTemperatureFromThermalStudy As System.Integer
```
Property Value 1 = Use initial temperature from the thermal study 0 = Do not use initial temperature from the thermal study
VBA: See CWThermalStudyOptions::UseTemperatureFromThermalStudy .
Examples: Visual Basic for Applications

### ICWThermalStudyOptions.UseTemperatureFromThermalStudy2
Gets or sets the status of initial temperature from the thermal study.
```vb
Property UseTemperatureFromThermalStudy2 As System.Boolean
```
Property Value -1 or true = Use initial temperature from the thermal study 0 or false = Do not use initial temperature from the thermal study
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

## ICWLoadsAndRestraintsManager

### ICWLoadsAndRestraintsManager.AddBearingLoad
Adds a bearing load.
```vb
Function AddBearingLoad( ByVal CoordinateSystem As System.Object, ByVal FirstFace As System.Object, ByRef ErrorCode As System.Integer ) As CWBearingLoad
```
Parameters CoordinateSystem Coordinate system (see Remarks ) FirstFace Array of cylindrical faces or circular shell edges (see Remarks ) ErrorCode Error as defined in swsLoadsAndRestraintsManagerBearingLoadError_e Return Value Bearing load
Return Value Bearing load
VBA: See CWLoadsAndRestraintsManager::AddBearingLoad .
Remarks: The Z-axis of CoordinateSystem must coincide with the axes of the faces in the FirstFace array.
Examples: Visual Basic for Applications; Example Add Bearing Load

### ICWLoadsAndRestraintsManager.AddBoltConnector
Adds a bolt connector.
```vb
Function AddBoltConnector( ByVal NBoltType As System.Integer, ByVal DispArrayBoltHead As System.Object, ByVal DispArrayBoltNut As System.Object, ByRef ErrorCode As System.Integer ) As CWBoltConnector
```
Parameters NBoltType Type of bolt as defined in swsBoltType_e DispArrayBoltHead Array of bolt heads DispArrayBoltNut Array of bolt nuts ErrorCode Error as defined in swsBoltConnectorError_e Return Value Bolt connector
Return Value Bolt connector
VBA: See CWLoadsAndRestraintsManager::AddBoltConnector .
Examples: Visual Basic for Applications; Example Create and Edit Bolt and Pin Connectors

### ICWLoadsAndRestraintsManager.AddCentrifugalForce
Creates a centrifugal force.
```vb
Function AddCentrifugalForce( ByVal DispEntity As System.Object, ByRef ErrorCode As System.Integer ) As CWCentriFugalForce
```
Parameters DispEntity Reference geometry (axis, edge, or cylindrical face) to specify direction ErrorCode Error as defined in swsCentrifugalForceError_e Return Value Centrifugal force
Return Value Centrifugal force
VBA: See CWLoadsAndRestraintsManager::AddCentrifugalForce .
Remarks: The density material property must be specified to consider the effect of centrifugal loads.
Examples: Visual Basic for Applications; Example Add Centrifugal Load

### ICWLoadsAndRestraintsManager.AddConvection
Creates a convection load for thermal studies.
```vb
Function AddConvection( ByVal DispArray As System.Object, ByRef ErrorCode As System.Integer ) As CWConvection
```
Parameters DispArray Array of faces or shell edges to apply convection ErrorCode Error as defined in swsConvectionError_e Return Value Convection
Return Value Convection
VBA: See CWLoadsAndRestraintsManager::AddConvection .
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWLoadsAndRestraintsManager.AddCyclicSymmetryRestraint
Adds a cyclic symmetry restraint.
```vb
Function AddCyclicSymmetryRestraint( ByVal FirstFace As System.Object, ByVal SecondFace As System.Object, ByVal Axis As System.Object, ByRef ErrorCode As System.Integer ) As CWRestraint
```
Parameters FirstFace First face (see Remarks ) SecondFace Second face (see Remarks ) Axis Axis of revolution (see Remarks ) ErrorCode Error code as defined in swsCyclicRestraintError_e Return Value ICWRestraint
Return Value ICWRestraint
VBA: See CWLoadsAndRestraintsManager::AddCyclicSymmetryRestraint .
Remarks: The cyclic symmetry fixture is appropriate for turbines, fans, flywheels, and motor rotors. FirstFace and SecondFace are the faces of a cut section of a model that can be repeated in a cyclical pattern about Axis. The angle between FirstFace and SecondFace should be evenly divisible by 360. You cannot apply a cyclic symmetry restraint to an invalid section. See the Creating Valid Sections for Cyclic Symmetry topic in the SOLIDWORKS Simulation Helps for more information. This method is valid only for static and nonlinear studies.
Examples: Visual Basic for Applications; Example Add Cyclic Symmetry Restraint

### ICWLoadsAndRestraintsManager.AddDistributedMass
Creates a distributed mass.
```vb
Function AddDistributedMass( ByVal DispArray As System.Object, ByVal NUnits As System.Integer, ByVal DMass As System.Double, ByRef ErrorCode As System.Integer ) As CWDistributedMass
```
Parameters DispArray Array of faces or shell edges on which to distribute the mass NUnits Units as defined in swsUnitSystem_e DMass Mass to distribute ErrorCode Error code as defined in swsDistributedMassError_e Return Value ICWDistributedMass
Return Value ICWDistributedMass
VBA: See CWLoadsAndRestraintsManager::AddDistributedMass .
Examples: Visual Basic for Applications; Create Linear Dynamic Study

### ICWLoadsAndRestraintsManager.AddEdgeWeldConnector
Creates the specified edge weld connector.
```vb
Function AddEdgeWeldConnector( ByVal FirstFace As System.Object, ByVal SecondFace As System.Object, ByVal Edges As System.Object, ByVal NEdgeWeldStyle As System.Integer, ByRef ErrorCode As System.Integer ) As CWEdgeWeldConnector
```
Parameters FirstFace Array of Face Set1 faces SecondFace Array of Face Set2 faces If NEdgeWeldStyle = swsEdgeWeldConnectorTypes_e... Then SecondFace array contains faces that are... swsEdgeWeldConnectorGrooveDoubleSided - or - swsEdgeWeldConnectorGrooveSingleSided Parallel to faces in FirstFace array swsEdgeWeldConnectorFilletDoubleSided - or - swsEdgeWeldConnectorFilletSingleSided Perpendicular to faces in FirstFace array Edges Array of touching edges between the faces specified in FirstFace and SecondFace arrays NEdgeWeldStyle Edge weld style as defined by swsEdgeWeldConnectorTypes_e ErrorCode Error code as defined by swsEdgeWeldCreationErrorCode_e Return Value ICWEdgeWeldConnector
Return Value ICWEdgeWeldConnector
VBA: See CWLoadsAndRestraintsManager::AddEdgeWeldConnector .
Examples: Visual Basic for Applications; Example Create and Edit Edge Weld Connector

### ICWLoadsAndRestraintsManager.AddElasticConnector
Adds an elastic support fixture.
```vb
Function AddElasticConnector( ByVal DispArray As System.Object, ByRef ErrorCode As System.Integer ) As CWElasticConnector
```
Parameters DispArray Array of faces for the elastic support fixture ErrorCode Error as defined in swsElasticConnectorError_e Return Value Elastic support fixture
Return Value Elastic support fixture
VBA: See CWLoadsAndRestraintsManager::AddElasticConnector .
Examples: Visual Basic for Applications; Example Add Elastic Support Fixture

### ICWLoadsAndRestraintsManager.AddForce
Obsolete. Superseded by ICWLoadsAndRestraintsManager::AddForce2 .
```vb
Function AddForce( ByVal NForceType As System.Integer, ByVal DispArray As System.Object, ByVal RefGeom As System.Object, ByRef ErrorCode As System.Integer ) As CWForce
```
Parameters NForceType Type of force as defined in swsForceType_e DispArray Array of entities (faces, edges, vertices, or points) to which to apply force RefGeom Reference geometry entity to specify direction (see Remarks ) ErrorCode Error as defined in swsForceError_e Return Value Force
Return Value Force
VBA: See CWLoadsAndRestraintsManager::AddForce .
Remarks: Pass null in RefGeom if reference geometry is not used.
Examples: Visual Basic for Applications

### ICWLoadsAndRestraintsManager.AddForce2
Obsolete. Superseded by ICWloadsAndRestraintsManager::AddForce3 .
```vb
Function AddForce2( ByVal NForceType As System.Integer, ByVal NSelectionType As System.Integer, ByVal DispArray As System.Object, ByVal RefGeom As System.Object, ByRef ErrorCode As System.Integer ) As CWForce
```
Parameters NForceType Type of force as defined in swsForceType_e NSelectionType Type of selection as defined in swsSelectionType_e DispArray Array of entities (faces, edges, vertices, or points) to which to apply force RefGeom Reference geometry entity to specify direction (see Remarks ) ErrorCode Error as defined in swsForceError_e Return Value Force
Return Value Force
VBA: See CWLoadsAndRestraintsManager::AddForce2 .
Remarks: Pass null in RefGeom if reference geometry is not used.
Examples: Visual Basic for Applications

### ICWLoadsAndRestraintsManager.AddForce3
Creates a force.
```vb
Function AddForce3( ByVal ForceType As System.Integer, ByVal SelectionType As System.Integer, ByVal RefDirType As System.Integer, ByVal InterpolationMode As System.Integer, ByVal DistPercentageOpt As System.Integer, ByVal NumOfRows As System.Integer, ByVal DistValue As System.Object, ByVal ForceValue As System.Object, ByVal NonUniformLoading As System.Boolean, ByVal NULoadingOnBeam As System.Boolean, ByVal NonUniformLoadDistDef As System.Integer, ByVal NonUniformLoadDistType As System.Integer, ByVal Ucode As System.Integer, ByVal TorqueNFVal As System.Double, ByVal Comps As System.Object, ByVal FlipOrigin As System.Boolean, ByVal IsCurvedBeam As System.Boolean, ByVal DispArray As System.Object, ByVal RefGeom As System.Object, ByVal PerUnitLength As System.Boolean, ByRef ErrorCode As System.Integer ) As CWForce
```
Parameters ForceType Type of force as defined in swsForceType_e (see Remarks ) SelectionType Type of selection as defined in swsSelectionType_e RefDirType Type of RefGeom to specify direction of this force: 0 = edge 1 = axis 2 = face or plane -1 = no direction reference (see Remarks ) InterpolationMode Interpolation mode as defined in swsTableDrivenInterpolationType_e ; only used in calculation if NonUniformLoadDistDef is swsBeamNonUniformLoadDef_e.swsTableDrivenLoad DistPercentageOpt Table-driven distribution option as defined in swsTableDrivenDistOption_e ; only used in calculation if NonUniformLoadDistDef is swsBeamNonUniformLoadDef_e.swsTableDrivenLoad NumOfRows Number of rows for a table-driven nonuniform distribution; only used in calculation if NonUniformLoadDistDef is swsBeamNonUniformLoadDef_e.swsTableDrivenLoad DistValue Array of the percentage or absolute distances of the forces along the beam for a table-driven nonuniform distribution; only used in calculation if NonUniformLoadDistDef is swsBeamNonUniformLoadDef_e.swsTableDrivenLoad; specify an empty array for a non-beam force ForceValue Array of the values of the forces along the beam for a table-driven nonuniform distribution; only used in calculation if NonUniformLoadDistDef is swsBeamNonUniformLoadDef_e.swsTableDrivenLoad; specify an empty array for a non-beam force NonUniformLoading True for nonuniform distribution, false for uniform distribution NULoadingOnBeam True for nonuniform distribution along a beam, false for uniform distribution along a beam NonUniformLoadDistDef Definition of nonuniform distribution along a beam as defined in swsBeamNonUniformLoadDef_e ; valid only if NULoadingOnBeam is true NonUniformLoadDistType Type of nonuniform distribution along a beam as defined in swsBeamNonUniformLoadType_e ; valid only if NULoadingOnBeam is true Ucode Components in Comps to use for a force in a selected direction: 1 = Fx 2 = Fy 4 = Fz 8 = Mx 16 = My 32 = Mz For combinations of forces and moments in various directions, add up the corresponding values; use 0 for a normal force (see Remarks ) TorqueNFVal Torque or normal force value; valid only if: RefDirType is -1 ForceType is not swsForceType_e.swsForceTypeForceOrMoment For beams, valid only for swsBeamNonUniformLoadDef_e.swsTotalLoad or swsBeamNonUniformLoadDef_e.swsCentralLoad Comps Array of force and moment component values; set elements to use in Ucode; for beams, valid only if NonUniformLoadDistDef is swsBeamNonUniformLoadDef_e.swsTotalLoad or swsBeamNonUniformLoadDef_e.swsCentralLoad (see Remarks ) FlipOrigin True to flip the beam origin, false to not; only used if NonUniformLoadDistDef is swsBeamNonUniformLoadDef_e.swsTableDrivenLoad IsCurvedBeam True if the beam is curved, false if the beam is straight DispArray Array of entities to which to apply this force RefGeom Reference direction entity used to indicate the direction of force; valid only if RefDirType is not -1 (see Remarks ) PerUnitLength True to change the unit of the force or moment value to a per unit length, false to not; valid only if SelectionType is swsSelectionType_e.swsSelectionBeams ErrorCode Error as defined in swsForceError_e Return Value Force
Return Value Force
VBA: See CWLoadsAndRestraintsManager::AddForce3 .
Remarks: You must specify Comps with a non-empty array of non-zero elements for all ForceTypes. Use Ucode to specify which elements of the Comps array to use for forces in a selected direction. If ForceType = swsForceType_e ... Then... swsForceTypeNormal or swsForceTypeTorque Specify: Comps with an array of six values: [1.0, 1.0, 1.0, 1.0, 1.0, 1.0] Ucode with 0 swsForceTypeForceOrMoment If RefDirType and RefGeom specify direction reference... And the model is a... Then specify Comps with array... And Ucode with... Edge Shell or beam [ force_along_edge , moment_along_edge ] 4 to use force_along_edge 32 to use moment_along_edge 36 to use both Edge Solid [ force_along_edge ] 4 to use force_along_edge Face or plane Shell or beam [ force_along_plane_dir_1 , force_along_plane_dir_2 , force_along_normal_to_plane , moment_along_plane_dir_1 , moment_along_plane_dir_2 , moment_along_normal_to_plane ] Combinations of values corresponding to the Comps array components to use Face or plane Solid [ force_along_plane_dir_1 , force_along_plane_dir_2 , force_along_normal_to_plane ] Combinations of values corresponding to the Comps array components to use Axis Shell or beam [ force_in_radial_direction, force_in_circumferential_direction, force_along_axis, moment_in_radial_direction, moment_in_circumferential_direction, moment_along_axis ] Combinations of values corresponding to the Comps array components to use Axis Solid [ force_in_radial_direction, force_in_circumferential_direction, force_along_axis ] Combinations of values corresponding to the Comps array components to use
Examples: Visual Basic for Applications; Apply Loads to Beam; Apply Table-driven Load to Multiple Beams; Add Nonuniform Force Distribution

### ICWLoadsAndRestraintsManager.AddGravity
Adds a gravity load.
```vb
Function AddGravity( ByVal DispEntity As System.Object, ByRef ErrorCode As System.Integer ) As CWGravity
```
Parameters DispEntity Reference geometry entity for direction ErrorCode Error as defined in swsGravityError_e Return Value Gravity
Return Value Gravity
VBA: See CWLoadsAndRestraintsManager::AddGravity .
Remarks: Mass density must be defined to include the effect of gravity.
Examples: Visual Basic for Applications; Example Add Gravity Load

### ICWLoadsAndRestraintsManager.AddHeatFlux
Creates a heat flux load for thermal studies.
```vb
Function AddHeatFlux( ByVal DispArray As System.Object, ByRef ErrorCode As System.Integer ) As CWHeatFlux
```
Parameters DispArray Array of entities (faces or shell edges) to which to apply heat flux ErrorCode Error as defined in swsHeatFluxError_e Return Value Heat flux
Return Value Heat flux
VBA: See CWLoadsAndRestraintsManager::AddHeatFlux .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWLoadsAndRestraintsManager.AddHeatPower
Creates a heat power load for thermal studies.
```vb
Function AddHeatPower( ByVal DispArray As System.Object, ByRef ErrorCode As System.Integer ) As CWHeatPower
```
Parameters DispArray Array of entities (bodies, faces, edges, and vertices) to which to apply heat power ErrorCode Error as defined in swsHeatPowerError_e Return Value Heat power
Return Value Heat power
VBA: See CWLoadsAndRestraintsManager::AddHeatPower .
Remarks: The specified value is applied to each entity.
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWLoadsAndRestraintsManager.AddInitialConditionForDynamicStudy
Obsolete. Superseded by ICWLoadsAndRestraintsManager::AddInitialConditionForDynamicStudy2 .
```vb
Function AddInitialConditionForDynamicStudy( ByVal NType As System.Integer, ByVal NSelectionType As System.Integer, ByVal DispArray As System.Object, ByVal RefEntity As System.Object, ByVal NUnits As System.Integer, ByVal BDir1 As System.Integer, ByVal DVal1 As System.Double, ByVal BDir2 As System.Integer, ByVal DVal2 As System.Double, ByVal BDir3 As System.Integer, ByVal DVal3 As System.Double, ByRef ErrorCode As System.Integer ) As CWDynamicInitialCondition
```
Parameters NType Type of dynamic initial condition as defined in swsDynamicInitialConditionType_e NSelectionType Type of selection as defined in swsSelectionType_e DispArray Array of components, bodies, or faces to which to apply the initial condition RefEntity Face, edge, or plane reference geometry NUnits If NType is swsDynamicInitialConditionType_e.... Then units are as defined in ... swsDynamicInitialConditionType_Displacement swsLinearUnit_e swsDynamicInitialConditionType_Velocity swsVelocityUnit_e swsDynamicInitialConditionType_Acceleration swsAccelerationUnit_e BDir1 1 to set a value along plane direction 1 or along a radial direction, 0 to not; valid only if RefEntity is set to a planar face, reference plane, cylindrical face, or null DVal1 Value along plane direction 1 or along a radial direction; valid only if BDir1 = 1 BDir2 1 to set a value along plane direction 2 or along a circumferential direction, 0 to not; valid only if RefEntity is set to a planar face, reference plane, cylindrical face, or null DVal2 Value along plane direction 2 or along a circumferential direction; valid only if BDir2 = 1 BDir3 1 to set a value along the normal to the plane or along an edge, 0 to not; valid only if RefEntity is set to a planar face, reference plane, cylindrical face, edge, or null DVal3 Value along the normal to the plane or along an edge; valid only if BDir3 = 1 ErrorCode Error code as defined in swsDynamicInitialConditionError_e }}--> Return Value ICWDynamicInitialCondition
Return Value ICWDynamicInitialCondition
VBA: See CWLoadsAndRestraintsManager::AddInitialConditionForDynamicStudy .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraintsManager.AddInitialConditionForDynamicStudy2
Creates the initial (time = 0) condition for displacements, velocities, or accelerations in modal time history and nonlinear dynamic studies.
```vb
Function AddInitialConditionForDynamicStudy2( ByVal NType As System.Integer, ByVal NSelectionType As System.Integer, ByVal DispArray As System.Object, ByVal RefEntity As System.Object, ByVal NUnits As System.Integer, ByVal BDir1 As System.Boolean, ByVal DVal1 As System.Double, ByVal BDir2 As System.Boolean, ByVal DVal2 As System.Double, ByVal BDir3 As System.Boolean, ByVal DVal3 As System.Double, ByRef ErrorCode As System.Integer ) As CWDynamicInitialCondition
```
Parameters NType Type of dynamic initial condition as defined in swsDynamicInitialConditionType_e NSelectionType Type of selection as defined in swsSelectionType_e DispArray Array of components, bodies, or faces to which to apply the initial condition RefEntity Face, edge, or plane reference geometry NUnits If NType is swsDynamicInitialConditionType_e.... Then units are as defined in ... swsDynamicInitialConditionType_Displacement swsLinearUnit_e swsDynamicInitialConditionType_Velocity swsVelocityUnit_e swsDynamicInitialConditionType_Acceleration swsAccelerationUnit_e BDir1 -1 or true to set a value along plane direction 1 or along a radial direction, 0 or false to not; valid only if RefEntity is set to a planar face, reference plane, cylindrical face, or null DVal1 Value along plane direction 1 or along a radial direction; valid only if BDir1 = -1 BDir2 -1 or true to set a value along plane direction 2 or along a circumferential direction, 0 or false to not; valid only if RefEntity is set to a planar face, reference plane, cylindrical face, or null DVal2 Value along plane direction 2 or along a circumferential direction; valid only if BDir2 = -1 BDir3 -1 or true to set a value along the normal to the plane or along an edge, 0 or false to not; valid only if RefEntity is set to a planar face, reference plane, cylindrical face, edge, or null DVal3 Value along the normal to the plane or along an edge; valid only if BDir3 = -1 ErrorCode Error code as defined in swsDynamicInitialConditionError_e }}--> Return Value ICWDynamicInitialCondition
Return Value ICWDynamicInitialCondition
Remarks: This method returns booleans or integers in out parameters BDir1, BDir2, and BDir3, depending on their prior declarations. If out parameters BDir1, BDir2, and BDir3 are cast as: Booleans, true or false is returned in each out parameter. Longs or integers, -1 (=true) or 0 (=false) is returned in each out parameter.

### ICWLoadsAndRestraintsManager.AddLinkConnector
Adds a link connector.
```vb
Function AddLinkConnector( ByVal VertexPointforFirstLocation As System.Object, ByVal VertexPointforSecondLocation As System.Object, ByRef ErrorCode As System.Integer ) As CWLinkConnector
```
Parameters VertexPointforFirstLocation Vertex for first location VertexPointforSecondLocation Vertex for second location ErrorCode Error as defined in swsLinkConnectorError_e Return Value Link connector
Return Value Link connector
VBA: See CWLoadsAndRestraintsManager::AddLinkConnector .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraintsManager.AddLinkageRod
Adds the specified linkage rod connector.
```vb
Function AddLinkageRod( ByVal VertexOrCircularFacesForEnd1 As System.Object, ByVal VertexOrCircularFacesForEnd2 As System.Object, ByVal NUnitSys As System.Integer, ByRef ErrorCode As System.Integer ) As CWLinkageRod
```
Parameters VertexOrCircularFacesForEnd1 Array of vertex or concentric circular faces (or edges for shells) for End 1 VertexOrCircularFacesForEnd2 Array of vertex or concentric circular faces (or edges for shells) for End 2 NUnitSys Units as defined by swsUnit_e ErrorCode Error code as defined by swsLinkageRodEndEditErrors_e Return Value ICWLinkageRod
Return Value ICWLinkageRod
VBA: See CWLoadsAndRestraintsManager::AddLinkageRod .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraintsManager.AddPinConnector
Adds a pin connector.
```vb
Function AddPinConnector( ByVal DispArrayComp1 As System.Object, ByVal DispArrayComp2 As System.Object, ByRef ErrorCode As System.Integer ) As CWPinConnector
```
Parameters DispArrayComp1 Array of one full cylindrical face or of multiple cylindrical faces of smaller angles belonging to the same body and coaxial with the same radius DispArrayComp2 Array of one full cylindrical face or of multiple cylindrical faces of smaller angles belonging to the same body, but different than the body in DispArrayComp1, and coaxial with the same radius ErrorCode Error as defined in swsPinConnectorError_e Return Value Pin connector
Return Value Pin connector
VBA: See CWLoadsAndRestraintsManager::AddPinConnector .
Examples: Visual Basic for Applications; Example Add Pin Connector

### ICWLoadsAndRestraintsManager.AddPrescribedDisplacement
Applies the specified prescribed displacement to the specified geometric entities to create a fixture.
```vb
Function AddPrescribedDisplacement( ByVal Displacements As System.Object, ByVal NLengthUnit As System.Integer, ByVal DispArray As System.Object, ByVal RefGeom As System.Object, ByRef ErrorCode As System.Integer ) As CWRestraint
```
Parameters Displacements Array: [ disp_x_coord , disp_y_coord , disp_z_coord , x_bool , y_bool , z_bool ] where: Array Element Value Description disp_x_coord Double Prescribed displacement in the x direction disp_y_coord Double Prescribed displacement in the y direction disp_z_coord Double Prescribed displacement in the z direction x_bool 0 or 1 Whether to use disp_x_coord (1 to use , 0 to not) y_bool 0 or 1 Whether to use disp_y_coord (1 to use , 0 to not) z_bool 0 or 1 Whether to use disp_z_coord (1 to use , 0 to not) (see Remarks ) NLengthUnit Units of length as defined in swsLinearUnit_e DispArray Array of entities to which to apply the prescribed displacement specified in Displacements RefGeom Reference entity for prescribed displacement direction ErrorCode Error code as defined in swsRestraintError_e Return Value ICWRestraint
Return Value ICWRestraint
VBA: See CWLoadsAndRestraintsManager::AddPrescribedDisplacement .
Remarks: The Displacements parameter specifies: length for reference geometry or flat faces - or - angle of rotation for cylindrical or spherical faces
Examples: Visual Basic for Applications; Example Add Prescribed Displacement

### ICWLoadsAndRestraintsManager.AddPressure
Creates pressure.
```vb
Function AddPressure( ByVal NPressureType As System.Integer, ByVal DispArray As System.Object, ByVal RegGeom As System.Object, ByRef ErrorCode As System.Integer ) As CWPressure
```
Parameters NPressureType Pressure type as defined in swsPressureType_e DispArray Array of entities (faces or shell edges) to which to apply pressure RegGeom Reference geometry for direction ErrorCode Error as defined in swsPressureError_e Return Value Pressure
Return Value Pressure
VBA: See CWLoadsAndRestraintsManager::AddPressure .
Remarks: Pass null or nothing if reference geometry is not used. If you pass a Dispatch pointer, it is ignored.
Examples: Visual Basic for Applications; Analyze Part

### ICWLoadsAndRestraintsManager.AddRadiation
Creates a radiation load for thermal studies.
```vb
Function AddRadiation( ByVal NRadType As System.Integer, ByVal DispArray As System.Object, ByRef ErrorCode As System.Integer ) As CWRadiation
```
Parameters NRadType Type of radiation as defined in swsRadiationType_e DispArray Array of faces or shell edges to which to apply radiation ErrorCode Error as defined in swsRadiationError_e Return Value Radiation
Return Value Radiation
VBA: See CWLoadsAndRestraintsManager::AddRadiation .
Remarks: See Also ICWLoadsAndRestraintsManager Interface ICWLoadsAndRestraintsManager Members
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWLoadsAndRestraintsManager.AddRemoteLoad
Creates a remote load of the specified type at the specified location.
```vb
Function AddRemoteLoad( ByVal NLoadType As System.Integer, ByVal DispArray As System.Object, ByVal NLocationUnits As System.Integer, ByVal DXValue As System.Double, ByVal DYValue As System.Double, ByVal DZValue As System.Double, ByRef ErrorCode As System.Integer ) As CWRemoteLoad
```
Parameters NLoadType Type of restraint or load as defined in swsRemoteLoadType_e DispArray If NLoadType is ... Then DispArray is an array of ... swsRemoteLoadType_e .swsRemoteLoadType_RigidLoadOrMass - or - swsRemoteLoadType_e .swsRemoteLoadType_RigidDisplacement Faces, edges, or vertices swsRemoteLoadType_e .swsRemoteLoadType_DirectLoad - or - swsRemoteLoadType_e .swsRemoteLoadType_DirectDisplacement Faces NLocationUnits Units as defined in swsLinearUnit_e DXValue X-coordinate of the point of application of the remote load DYValue Y-coordinate of the point of application of the remote load DZValue Z-coordinate of the point of application of the remote load ErrorCode Error code as defined in swsLoadsAndRestraintsError_e }}--> Return Value ICWRemoteLoad
Return Value ICWRemoteLoad
VBA: See CWLoadsAndRestraintsManager::AddRemoteLoad .
Examples: Visual Basic for Applications; Example Add Remote Load

### ICWLoadsAndRestraintsManager.AddRestraint
Creates a restraint.
```vb
Function AddRestraint( ByVal NRestraintType As System.Integer, ByVal DispArray As System.Object, ByVal RefGeom As System.Object, ByRef ErrorCode As System.Integer ) As CWRestraint
```
Parameters NRestraintType Type of restraint as defined in swsRestraintType_e DispArray Array of entities to which to apply the restraint RefGeom Reference geometry entity for direction; valid only if NRestraintType is set to swsRestraintType_e.swsRestraintTypeReferenceGeometry ErrorCode Error code as defined in swsRestraintError_e Return Value Restraint
Return Value Restraint
VBA: See CWLoadsAndRestraintsManager::AddRestraint .
Examples: Visual Basic for Applications; Analyze Part

### ICWLoadsAndRestraintsManager.AddRigidConnector
Adds a rigid connector.
```vb
Function AddRigidConnector( ByVal PersistIDFaceArray As System.Object, ByVal PersistIDTargetArray As System.Object, ByRef ErrorCode As System.Integer ) As CWRigidConnector
```
Parameters PersistIDFaceArray Array of objects of the faces from a solid body PersistIDTargetArray Array of objects of the faces from the target solid body ErrorCode Error as defined in swsRigidConnectorError_e Return Value Rigid connector
Return Value Rigid connector
VBA: See CWLoadsAndRestraintsManager::AddRigidConnector .
Examples: Visual Basic for Applications; Example Add Rigid Connector

### ICWLoadsAndRestraintsManager.AddSelectedBaseExcitation
Obsolete. Superseded by ICWLoadsAndRestraintsManager::AddSelectedBaseExcitation2 .
```vb
Function AddSelectedBaseExcitation( ByVal NType As System.Integer, ByVal SFixtureName As System.String, ByVal NUnits As System.Integer, ByVal BDir1 As System.Integer, ByVal DVal1 As System.Double, ByVal BDir2 As System.Integer, ByVal DVal2 As System.Double, ByVal BDir3 As System.Integer, ByVal DVal3 As System.Double, ByRef ErrorCode As System.Integer ) As CWBaseExcitation
```
Parameters NType Type of excitation as defined in swsBaseExcitationType_e SFixtureName Model entity to which the restraint is applied NUnits If NType is ... Then units are as defined in ... swsBaseExcitationType_e.swsBaseExcitationType_Displacement swsLinearUnit_e swsBaseExcitationType_e.swsBaseExcitationType_Velocity swsVelocityUnit_e swsBaseExcitationType_e.swsBaseExcitationType_Acceleration swsAccelerationUnit_e BDir1 1 to add excitation along plane direction 1, 0 to not DVal1 Excitation along plane direction 1 (see Remarks ) BDir2 1 to add excitation along plane direction 2, 0 to not DVal2 Excitation along plane direction 2 (see Remarks ) BDir3 1 to add excitation along normal to plane, 0 to not DVal3 Excitation along normal to the plane (see Remarks ) ErrorCode Error code as defined in swsRestraintError_e }}--> Return Value ICWBaseExcitation
Return Value ICWBaseExcitation
VBA: See CWLoadsAndRestraintsManager::AddSelectedBaseExcitation .
Remarks: Set negative values in DVal1, DVal2, and DVal3 to indicate reverse directions. If this is a random vibration dynamic study, then DVal1, DVal2, and DVal3 are the Power Spectral Densities of the specified NType.
Examples: Visual Basic for Applications

### ICWLoadsAndRestraintsManager.AddSelectedBaseExcitation2
Applies a base excitation to the specified restraint along the specified directions for dynamic studies.
```vb
Function AddSelectedBaseExcitation2( ByVal NType As System.Integer, ByVal SFixtureName As System.String, ByVal NUnits As System.Integer, ByVal BDir1 As System.Boolean, ByVal DVal1 As System.Double, ByVal BDir2 As System.Boolean, ByVal DVal2 As System.Double, ByVal BDir3 As System.Boolean, ByVal DVal3 As System.Double, ByRef ErrorCode As System.Integer ) As CWBaseExcitation
```
Parameters NType Type of excitation as defined in swsBaseExcitationType_e SFixtureName Model entity to which the restraint is applied NUnits If NType is ... Then units are as defined in ... swsBaseExcitationType_e.swsBaseExcitationType_Displacement swsLinearUnit_e swsBaseExcitationType_e.swsBaseExcitationType_Velocity swsVelocityUnit_e swsBaseExcitationType_e.swsBaseExcitationType_Acceleration swsAccelerationUnit_e BDir1 -1 or true to add excitation along plane direction 1, 0 or false to not DVal1 Excitation along plane direction 1 (see Remarks ) BDir2 -1 or true to add excitation along plane direction 2, 0 or false to not DVal2 Excitation along plane direction 2 (see Remarks ) BDir3 -1 or true to add excitation along normal to plane, 0 or false to not DVal3 Excitation along normal to the plane (see Remarks ) ErrorCode Error code as defined in swsRestraintError_e }}--> Return Value ICWBaseExcitation
Return Value ICWBaseExcitation
Remarks: This method returns booleans or integers in out parameters BDir1, BDir2, and BDir3, depending on their prior declarations. If out parameters BDir1, BDir2, and BDir3 are cast as: Booleans, true or false is returned in each out parameter. Longs or integers, -1 (=true) or 0 (=false) is returned in each out parameter. Set negative values in DVal1, DVal2, and DVal3 to indicate reverse directions. If this is a random vibration dynamic study, then DVal1, DVal2, and DVal3 are the Power Spectral Densities of the specified NType.

### ICWLoadsAndRestraintsManager.AddSpotWeldConnector
Adds a spot weld connector.
```vb
Function AddSpotWeldConnector( ByVal FirstFace As System.Object, ByVal SecondFace As System.Object, ByVal DispArrayWeldLocations As System.Object, ByRef ErrorCode As System.Integer ) As CWSpotWeldConnector
```
Parameters FirstFace Face of a shell or solid body SecondFace Face of a shell or solid body that belongs to a different body than FirstFace DispArrayWeldLocations Array of vertices or an array of reference points to project on the faces to determine the locations of the spot welds ErrorCode Error as defined in swsSpotWeldConnectorError_e Return Value Spot weld connector
Return Value Spot weld connector
VBA: See CWLoadsAndRestraintsManager::AddSpotWeldConnector .
Examples: Visual Basic for Applications; Example Add Spot Weld Connector

### ICWLoadsAndRestraintsManager.AddSpringConnector
Adds a spring connector.
```vb
Function AddSpringConnector( ByVal NSpringSubType As System.Integer, ByVal DispArrayComp1 As System.Object, ByVal DispArrayComp2 As System.Object, ByRef ErrorCode As System.Integer ) As CWSpringConnector
```
Parameters NSpringSubType Spring connector type as defined in swsSpringConnectorType_e DispArrayComp1 Array of planar faces, cylindrical faces, or vertices from a body depending on NSpringSubType DispArrayComp2 Array of planar faces, cylindrical faces, or vertices, depending on NSpringSubType, from another body ErrorCode Error as defined in swsSpringConnectorError_e Return Value Spring connector
Return Value Spring connector
VBA: See CWLoadsAndRestraintsManager::AddSpringConnector .
Examples: Visual Basic for Applications; Example Add Spring Connector

### ICWLoadsAndRestraintsManager.AddTemperature
Creates a temperature load for thermal studies.
```vb
Function AddTemperature( ByVal DispArray As System.Object, ByRef ErrorCode As System.Integer ) As CWTemperature
```
Parameters DispArray Array of entities (faces, edges, and vertices) to which to apply temperature ErrorCode Error as defined in swsTemperatureError_e Return Value Temperature
Return Value Temperature
VBA: See CWLoadsAndRestraintsManager::AddTemperature .
Remarks: Temperature can be used with static, nonlinear, frequency, buckling, and thermal studies.
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWLoadsAndRestraintsManager.AddUniformBaseExcitation
Obsolete. Superseded by ICWLoadsAndRestraintsManager::AddUniformBaseExcitation2 .
```vb
Function AddUniformBaseExcitation( ByVal NType As System.Integer, ByVal DirectionEntity As System.Object, ByVal NUnits As System.Integer, ByVal BDir1 As System.Integer, ByVal DVal1 As System.Double, ByVal BDir2 As System.Integer, ByVal DVal2 As System.Double, ByVal BDir3 As System.Integer, ByVal DVal3 As System.Double, ByRef ErrorCode As System.Integer ) As CWBaseExcitation
```
Parameters NType Type of excitation as defined in swsBaseExcitationType_e DirectionEntity Face, edge, or plane of direction for the excitation; null for global direction if Response Spectrum study NUnits If NType is ... Then units are as defined in ... swsBaseExcitationType_e.swsBaseExcitationType_Displacement swsLinearUnit_e swsBaseExcitationType_e.swsBaseExcitationType_Velocity swsVelocityUnit_e swsBaseExcitationType_e.swsBaseExcitationType_Acceleration swsAccelerationUnit_e BDir1 1 to add excitation along plane direction 1, 0 to not; valid only if DirectionEntity is a face, plane, or unspecified DVal1 Excitation along plane direction 1; valid only if BDir1 = 1 and DirectionEntity is a face, plane, or unspecified (see Remarks ) BDir2 1 to add excitation along plane direction 2; 0 to not; valid only if DirectionEntity is a face, plane, or unspecified DVal2 Excitation along plane direction 2; valid only if BDir2 = 1 and DirectionEntity is a face, plane, or unspecified (see Remarks ) BDir3 1 to add excitation along normal to the plane or along an edge; 0 to not; valid only if DirectionEntity is a face, edge plane, or unspecified DVal3 Excitation along normal to the plane or along an edge; valid only if BDir3 = 1 and DirectionEntity is a face, edge, plane, or unspecified (see Remarks ) ErrorCode Error code as defined in swsRestraintError_e }}--> Return Value ICWBaseExcitation
Return Value ICWBaseExcitation
VBA: See CWLoadsAndRestraintsManager::AddUniformBaseExcitation .
Remarks: Set negative values in DVal1, DVal2, and DVal3 to indicate reverse directions. If this is a random vibration dynamic study, then DVal1, DVal2, and DVal3 are the Power Spectral Densities of the specified NType.
Examples: Visual Basic for Applications; Example Create Linear Dynamic Study

### ICWLoadsAndRestraintsManager.AddUniformBaseExcitation2
Adds a uniform base excitation to all locations restrained in the specified directions for dynamic studies.
```vb
Function AddUniformBaseExcitation2( ByVal NType As System.Integer, ByVal DirectionEntity As System.Object, ByVal NUnits As System.Integer, ByVal BDir1 As System.Boolean, ByVal DVal1 As System.Double, ByVal BDir2 As System.Boolean, ByVal DVal2 As System.Double, ByVal BDir3 As System.Boolean, ByVal DVal3 As System.Double, ByRef ErrorCode As System.Integer ) As CWBaseExcitation
```
Parameters NType Type of excitation as defined in swsBaseExcitationType_e DirectionEntity Face, edge, or plane of direction for the excitation; null for global direction if Response Spectrum study NUnits If NType is ... Then units are as defined in ... swsBaseExcitationType_e.swsBaseExcitationType_Displacement swsLinearUnit_e swsBaseExcitationType_e.swsBaseExcitationType_Velocity swsVelocityUnit_e swsBaseExcitationType_e.swsBaseExcitationType_Acceleration swsAccelerationUnit_e BDir1 -1 or true to add excitation along plane direction 1, 0 or false to not; valid only if DirectionEntity is a face, plane, or unspecified DVal1 Excitation along plane direction 1; valid only if BDir1 = -1 and DirectionEntity is a face, plane, or unspecified (see Remarks ) BDir2 -1 or true to add excitation along plane direction 2; 0 or false to not; valid only if DirectionEntity is a face, plane, or unspecified DVal2 Excitation along plane direction 2; valid only if BDir2 = -1 and DirectionEntity is a face, plane, or unspecified (see Remarks ) BDir3 -1 or true to add excitation along normal to the plane or along an edge; 0 or false to not; valid only if DirectionEntity is a face, edge plane, or unspecified DVal3 Excitation along normal to the plane or along an edge; valid only if BDir3 = -1 and DirectionEntity is a face, edge, plane, or unspecified (see Remarks ) ErrorCode Error code as defined in swsRestraintError_e Return Value ICWBaseExcitation
Return Value ICWBaseExcitation
Remarks: This method returns booleans or integers in out parameters BDir1, BDir2, and BDir3, depending on their prior declarations. If out parameters BDir1, BDir2, and BDir3 are cast as: Booleans, true or false is returned in each out parameter. Longs or integers, -1 (=true) or 0 (=false) is returned in each out parameter. Set negative values in DVal1, DVal2, and DVal3 to indicate reverse directions. If this is a random vibration dynamic study, then DVal1, DVal2, and DVal3 are the Power Spectral Densities of the specified NType.

### ICWLoadsAndRestraintsManager.CopyLoadsAndRestraintsToStudy
Copies the specified loads and restraints to the specified study.
```vb
Function CopyLoadsAndRestraintsToStudy( ByVal SStudyName As System.String, ByVal LRNames As System.Object ) As System.Integer
```
Parameters SStudyName Name of study to which to copy loads and restraints LRNames Array of names of loads and restraints to copy (see Remarks ) Return Value Error code as defined in swsCopyItemsError_e
Return Value Error code as defined in swsCopyItemsError_e
VBA: See CWLoadsAndRestraintsManager::CopyLoadsAndRestraintsToStudy .
Remarks: To populate the array of LRNames, either: inspect the Simulation study tree. - or - call ICWLoadsAndRestraintsManager::GetLoadsAndRestraints to get each ICWLoadsAndRestraints and then call ICWLoadsAndRestraints::Name to get the name of each load or restraint. If a load or restraint name already exists, Copy[1] is prepended to the name of the copied load or restraint.
Examples: Visual Basic for Applications; Example Copy Items to Another Study

### ICWLoadsAndRestraintsManager.Count
Gets the number of loads and restraints defined in the active study.
```vb
ReadOnly Property Count As System.Integer
```
Property Value Number of loads and restraints
VBA: See CWLoadsAndRestraintsManager::Count .
Remarks: Call this method before calling ICWLoadsAndRestraintsManager::GetLoadsAndRestraints .
Examples: Visual Basic for Applications; Example Copy Items to Another Study

### ICWLoadsAndRestraintsManager.DeleteLoadsAndRestraints
Deletes a load or restraint.
```vb
Sub DeleteLoadsAndRestraints( ByVal SName As System.String )
```
Parameters SName Name of the restraint or load object to delete
VBA: See CWLoadsAndRestraintsManager::DeleteLoadsAndRestraints .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraintsManager.GetEdgeWeldConnector
Gets the edge weld connector at the specified index.
```vb
Function GetEdgeWeldConnector( ByVal NIndex As System.Integer, ByRef ErrorCode As System.Integer ) As CWEdgeWeldConnector
```
Parameters NIndex Index of the edge weld connector to get (see Remarks ) ErrorCode Error code as defined by swsEdgeWeldCreationErrorCode_e Return Value ICWEdgeWeldConnector
Return Value ICWEdgeWeldConnector
VBA: See CWLoadsAndRestraintsManager::GetEdgeWeldConnector .
Remarks: Call ICWLoadsAndRestraintsManager::Count to determine NIndex. The vertical position of a load or restraint in the Simulation Study tree does not correspond to NIndex. You must use ICWLoadsAndRestraintsManager::GetLoadsAndRestraints to determine the mapping between NIndex and the study's loads and restraints.
Examples: Visual Basic for Applications; Example Create and Edit Edge Weld Connector

### ICWLoadsAndRestraintsManager.GetLinkageRod
Gets the specified linkage rod connector.
```vb
Function GetLinkageRod( ByVal SName As System.String, ByRef ErrorCode As System.Integer ) As CWLinkageRod
```
Parameters SName Name of the linkage rod connector as shown Connections > Connectors branch of the simulation study tree ErrorCode Error as defined in swsLinkConnectorError_e Return Value ICWLinkageRod
Return Value ICWLinkageRod
VBA: See CWLoadsAndRestraintsManager::GetLinkageRod .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraintsManager.GetLoadsAndRestraints
Gets the load or restraint at the specified index.
```vb
Function GetLoadsAndRestraints( ByVal NIndex As System.Integer, ByRef ErrorCode As System.Integer ) As CWLoadsAndRestraints
```
Parameters NIndex 0-based index of load or restraint to get ErrorCode Error as defined in swsLoadsAndRestraintsError_e Return Value Load or restraint at specified index
Return Value Load or restraint at specified index
VBA: See CWLoadsAndRestraintsManager::GetLoadsAndRestraints .
Remarks: Before calling this method, call ICWLoadsAndRestraintsManager::Count to get a valid value for NIndex. NIndex does not correspond to the vertical position of the load or restraint in the Simulation Study tree. Cast the return object to the load or restraint class of interest (e.g., CWPinConnector in the example).
Examples: Visual Basic for Applications; Example Create and Edit Bolt and Pin Connectors

## ICWLoadsAndRestraints

### ICWLoadsAndRestraints.EntityCount
Gets the number of entities in this load or restraint.
```vb
ReadOnly Property EntityCount As System.Integer
```
Property Value Number of entities in this load or restraint; does not include reference geometry
VBA: See CWLoadsAndRestraints::EntityCount .
Remarks: Call this method before calling ICWLoadAndRestraint::GetEntityAt .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraints.GetDirection
Gets the direction of this load or restraint.
```vb
Function GetDirection() As System.Object
```
Return Value Array of x, y, z coordinates
VBA: See CWLoadsAndRestraints::GetDirection .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraints.GetEntityAt
Gets the entity at the specified index in this load or restraint.
```vb
Function GetEntityAt( ByVal NIndex As System.Integer, ByRef NSelType As System.Integer ) As System.Object
```
Parameters NIndex 0-based index of entity NSelType Entity type as defined in swSelectType_e Return Value Entity
Return Value Entity
VBA: See CWLoadsAndRestraints::GetEntityAt .
Remarks: Call ICWLoadsAndRestraints::EntityCount before calling this method to get a valid value for NIndex.
Examples: Visual Basic for Applications

### ICWLoadsAndRestraints.GetReferenceGeometry
Gets the entity used for reference geometry in this load or restraint.
```vb
Function GetReferenceGeometry( ByRef NSelType As System.Integer ) As System.Object
```
Parameters NSelType Entity type as defined in swSelectType_e Return Value Entity
Return Value Entity
VBA: See CWLoadsAndRestraints::GetReferenceGeometry .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraints.Name
Gets the name of load or restraint.
```vb
ReadOnly Property Name As System.String
```
Property Value Name of the load or restraint
VBA: See CWLoadsAndRestraints::Name .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraints.State
Gets the suppression state of the load or restraint.
```vb
ReadOnly Property State As System.Integer
```
Property Value Suppression state as defined in swsSuppressionState_e
VBA: See CWLoadsAndRestraints::State .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraints.SuppressUnSuppress
Suppresses or unsuppresses the load or restraint depending on its state .
```vb
Sub SuppressUnSuppress()
```
VBA: See CWLoadsAndRestraints::SuppressUnSuppress .
Examples: Visual Basic for Applications

### ICWLoadsAndRestraints.Type
Gets the type of the load or restraint.
```vb
ReadOnly Property Type As System.Integer
```
Property Value Type of load and restraint as defined in swsLoadsAndRestraintsType_e
VBA: See CWLoadsAndRestraints::Type .
Examples: Visual Basic for Applications

## ICWHeatPower

### ICWHeatPower.GetCutOffTemperatures
Gets the thermostat cut off temperatures (lower and upper bounds) used in heat power for transient thermal studies.
```vb
Sub GetCutOffTemperatures( ByRef DVal1 As System.Double, ByRef DVal2 As System.Double )
```
Parameters DVal1 Lower-bound temperature DVal2 Lower-bound temperature
VBA: See CWHeatPower::GetCutOffTemperatures .
Examples: Visual Basic for Applications

### ICWHeatPower.GetTemperatureCurve
Gets the temperature curve data for heat power.
```vb
Function GetTemperatureCurve() As System.Object
```
Return Value Array of temperature curve data (see Remarks )
VBA: See CWHeatPower::GetTemperatureCurve .
Remarks: Temperature curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi
Examples: Visual Basic for Applications

### ICWHeatPower.GetThermostat
Gets the location of the thermostat for heat power.
```vb
Function GetThermostat() As System.Object
```
Return Value Vertex for the location of heat power
VBA: See CWHeatPower::GetThermostat .
Remarks: Thermostat is used for transient thermal studies only.
Examples: Visual Basic for Applications

### ICWHeatPower.GetTimeCurve
Gets the time curve data for heat power.
```vb
Function GetTimeCurve() As System.Object
```
Return Value Array of time curve data (see Remarks )
VBA: See CWHeatPower::GetTimeCurve .
Remarks: Time curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi
Examples: Visual Basic for Applications

### ICWHeatPower.HPValue
Gets or sets the heat power value for heat power.
```vb
Property HPValue As System.Double
```
Property Value Heat power value
VBA: See CWHeatPower::HPValue .
Examples: Visual Basic for Applications

### ICWHeatPower.HeatPowerBeginEdit
Starts editing heat power.
```vb
Sub HeatPowerBeginEdit()
```
VBA: See CWHeatPower::HeatPowerBeginEdit .
Remarks: To start editing heat power, you must call this method. You must call ICWHeatPower::HeatPowerBeginEdit to end editing heat power. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWHeatPower.HeatPowerEndEdit
Ends editing heat power.
```vb
Function HeatPowerEndEdit() As System.Integer
```
Return Value Error as defined in swsHeatPowerEndEditError_e
VBA: See CWHeatPower::HeatPowerEndEdit .
Remarks: You must call ICWHeatPower::HeatPowerBeginEdit to start editing heat power. To ends editing heat power, you must call this method. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWHeatPower.IncludeThermostat
Obsolete. Superseded by ICWHeatPower::IncludeThermostat2 .
```vb
Property IncludeThermostat As System.Integer
```
Property Value 1 to include thermostat, 0 to not
VBA: See CWHeatPower::IncludeThermostat .
Remarks: Thermostat is used for transient thermal studies only.
Examples: Visual Basic for Applications

### ICWHeatPower.IncludeThermostat2
Gets or sets whether to include thermostat effects in heat power for transient thermal studies.
```vb
Property IncludeThermostat2 As System.Boolean
```
Property Value -1 or true to include thermostat, 0 or false to not
Remarks: Thermostat is used for transient thermal studies only. This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWHeatPower.InsertEntity
Inserts a new entity for heat power.
```vb
Sub InsertEntity( ByVal DispEntity As System.Object )
```
Parameters DispEntity Entity
VBA: See CWHeatPower::InsertEntity .
Examples: Visual Basic for Applications

### ICWHeatPower.RemoveEntity
Removes the entity at the specified }}-->index to which to apply heat power.
```vb
Sub RemoveEntity( ByVal NIndex As System.Integer )
```
Parameters NIndex 0-based index of entity to remove
VBA: See CWHeatPower::RemoveEntity .
Examples: Visual Basic for Applications

### ICWHeatPower.ReverseDirection
Obsolete. Superseded by ICWHeatPower::ReverseDirection2 .
```vb
Property ReverseDirection As System.Integer
```
Property Value 1 to reverse direction, 0 to not
VBA: See CWHeatPower::ReverseDirection .
Examples: Visual Basic for Applications

### ICWHeatPower.ReverseDirection2
Gets or sets whether to reverse the direction of heat power.
```vb
Property ReverseDirection2 As System.Boolean
```
Property Value -1 or true to reverse direction, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWHeatPower.SetCutOffTemperatures
Sets the thermostat cut off temperatures (lower and upper bounds) used in heat power for transient thermal studies.
```vb
Sub SetCutOffTemperatures( ByVal DVal1 As System.Double, ByVal DVal2 As System.Double )
```
Parameters DVal1 Lower-bound temperature DVal2 Upper-bound temperature
VBA: See CWHeatPower::SetCutOffTemperatures .
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWHeatPower.SetTemperatureCurve
Obsolete. Superseded by ICWHeatPower::SetTemperatureCurve2 .
```vb
Function SetTemperatureCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters VarCurveData Array of temperature curve data (see Remarks ) ErrorCode Temperature curve error as defined in swsTemperatureCurveError_e Return Value 1 to use curve, 0 to not
Return Value 1 to use curve, 0 to not
VBA: See CWHeatPower::SetTemperatureCurve .
Remarks: Temperature curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi
Examples: Visual Basic for Applications

### ICWHeatPower.SetTemperatureCurve2
Sets the temperature curve data for heat power.
```vb
Function SetTemperatureCurve2( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters VarCurveData Array of temperature curve data (see Remarks ) ErrorCode Temperature curve error as defined in swsTemperatureCurveError_e Return Value -1 or true to use curve, 0 or false to not
Return Value -1 or true to use curve, 0 or false to not
Remarks: Temperature curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi

### ICWHeatPower.SetThermostat
Sets a vertex for the location of a thermostat for heat power.
```vb
Sub SetThermostat( ByVal DispVertex As System.Object )
```
Parameters DispVertex Vertex for the location of a thermostat
VBA: See CWHeatPower::SetThermostat .
Remarks: Thermostat is used for transient thermal studies only.
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWHeatPower.SetTimeCurve
Obsolete. Superseded by ICWHeatPower::SetTimeCurve2 .
```vb
Function SetTimeCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters VarCurveData Array of time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value 1 to use curve, 0 to not
Return Value 1 to use curve, 0 to not
VBA: See CWHeatPower::SetTimeCurve .
Remarks: Time curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi
Examples: Visual Basic for Applications

### ICWHeatPower.SetTimeCurve2
Sets the time curve data for heat power.
```vb
Function SetTimeCurve2( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters VarCurveData Array of time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value -1 or true to use curve, 0 or false to not
Return Value -1 or true to use curve, 0 or false to not
Remarks: Time curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi

### ICWHeatPower.ThermostatUnits
Gets or sets the thermostat units for heat power.
```vb
Property ThermostatUnits As System.Integer
```
Property Value Temperature unit as defined swsTemperatureUnit_e
VBA: See CWHeatPower::ThermostatUnits .
Remarks: Thermostat is used for transient thermal studies only.
Examples: Visual Basic for Applications

### ICWHeatPower.Unit
Gets or sets the units of heat power.
```vb
Property Unit As System.Integer
```
Property Value Units as defined in swsUnit_e
VBA: See CWHeatPower::Unit .
Examples: Visual Basic for Applications

### ICWHeatPower.UseTemperatureCurve
Obsolete. Superseded by ICWHeatPower::UseTemperatureCurve2 .
```vb
Property UseTemperatureCurve As System.Integer
```
Property Value 1 to use temperature curve, 0 to not
VBA: See CWHeatPower::UseTemperatureCurve .
Examples: Visual Basic for Applications

### ICWHeatPower.UseTemperatureCurve2
Gets or sets whether to use a temperature curve for heat power.
```vb
Property UseTemperatureCurve2 As System.Boolean
```
Property Value -1 or true to use temperature curve, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWHeatPower.UseTimeCurve
Obsolete. Superseded by ICWHeatPower::UseTimeCurve2 .
```vb
Property UseTimeCurve As System.Integer
```
Property Value 1 to use time curve, 0 to not
VBA: See CWHeatPower::UseTimeCurve .
Examples: Visual Basic for Applications

### ICWHeatPower.UseTimeCurve2
Gets or sets whether to use a time curve for heat power.
```vb
Property UseTimeCurve2 As System.Boolean
```
Property Value -1 or true to use time curve, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

## ICWHeatFlux

### ICWHeatFlux.GetCutOffTemperatures
Gets the thermostat lower- and upper-bound temperatures in this heat flux load. }}--> }}-->
```vb
Sub GetCutOffTemperatures( ByRef DVal1 As System.Double, ByRef DVal2 As System.Double )
```
Parameters DVal1 Lower-bound temperature DVal2 Upper-bound temperature
VBA: See CWHeatFlux::GetCutOffTemperatures .
Examples: Visual Basic for Applications

### ICWHeatFlux.GetTemperatureCurve
Gets the temperature curve data for heat flux.
```vb
Function GetTemperatureCurve() As System.Object
```
Return Value Array of temperature curve data (see Remarks )
VBA: See CWHeatFlux::GetTemperatureCurve .
Remarks: Temperature curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi
Examples: Visual Basic for Applications

### ICWHeatFlux.GetThermostat
Gets the location of the thermostat for heat flux.
```vb
Function GetThermostat() As System.Object
```
Return Value Vertex of the location of the thermostat
VBA: See CWHeatFlux::GetThermostat .
Remarks: Thermostat is used for transient thermal studies only.
Examples: Visual Basic for Applications

### ICWHeatFlux.GetTimeCurve
Gets the time curve data for heat flux.
```vb
Function GetTimeCurve() As System.Object
```
Return Value Array of time curve data (see Remarks )
VBA: See CWHeatFlux::GetTimeCurve .
Remarks: Time curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi
Examples: Visual Basic for Applications

### ICWHeatFlux.HFValue
Gets or sets the heat flux value used in thermal studies.
```vb
Property HFValue As System.Double
```
Property Value Heat flux value
VBA: See CWHeatFlux::HFValue .
Remarks: After setting this property, you must select the location of the thermostat in the model and pass it to ICWHeatFlux::SetThermostat .
Examples: Visual Basic for Applications

### ICWHeatFlux.HeatFluxBeginEdit
Starts editing heat flux.
```vb
Sub HeatFluxBeginEdit()
```
VBA: See CWHeatFlux::HeatFluxBeginEdit .
Remarks: To start editing heat flux, you must call this method. You must call ICWHeatFlux::HeatFluxEndEdit to end editing heat flux. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications

### ICWHeatFlux.HeatFluxEndEdit
Ends editing heat flux.
```vb
Function HeatFluxEndEdit() As System.Integer
```
Return Value Error as defined in swsHeatFluxEndEditError_e
VBA: See CWHeatFlux::HeatFluxEndEdit .
Remarks: You must call ICWHeatFlux::HeatFluxBeginEdit to start editing heat flux. To end editing heat flux, you must call this method. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications

### ICWHeatFlux.IncludeThermostat
Obsolete. Superseded by ICWHeatFlux::IncludeThermostat2 .
```vb
Property IncludeThermostat As System.Integer
```
Property Value 1 to include thermostat effects, 0 to not
VBA: See CWHeatFlux::IncludeThermostat .
Remarks: Thermostat is used for transient thermal studies only.
Examples: Visual Basic for Applications

### ICWHeatFlux.IncludeThermostat2
Gets or sets whether to include thermostat effects in this heat flux.
```vb
Property IncludeThermostat2 As System.Boolean
```
Property Value -1 or true to include thermostat effects, 0 or false to not
Remarks: Thermostat is used for transient thermal studies only. This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWHeatFlux.InsertEntity
Inserts an entity on which to apply this heat flux.
```vb
Sub InsertEntity( ByVal DispEntity As System.Object )
```
Parameters DispEntity Entity
VBA: See CWHeatFlux::InsertEntity .
Examples: Visual Basic for Applications

### ICWHeatFlux.PerUnitLength
Obsolete. Superseded by ICWHeatFlux::PerUnitLength2 .
```vb
Property PerUnitLength As System.Integer
```
Property Value True if heat flux uses per unit length while selecting beam bodies, false if not
VBA: See CWHeatFlux::PerUnitLength .
Examples: Visual Basic for Applications

### ICWHeatFlux.PerUnitLength2
Gets or sets whether heat flux uses per unit length while selecting beam bodies.
```vb
Property PerUnitLength2 As System.Boolean
```
Property Value -1 or true if heat flux uses per unit length while selecting beam bodies, 0 or false if not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWHeatFlux.RemoveEntity
Removes an entity from this heat flux.
```vb
Sub RemoveEntity( ByVal NIndex As System.Integer )
```
Parameters NIndex 0-based index of entity to remove
VBA: See CWHeatFlux::RemoveEntity .
Examples: Visual Basic for Applications

### ICWHeatFlux.ReverseDirection
Obsolete. Superseded by ICWHeatFlux::ReverseDirection2 .
```vb
Property ReverseDirection As System.Integer
```
Property Value 1 to reverse direction, 0 to not
VBA: See CWHeatFlux::ReverseDirection .
Examples: Visual Basic for Applications

### ICWHeatFlux.ReverseDirection2
Gets or sets whether to reverse the direction of heat flux.
```vb
Property ReverseDirection2 As System.Boolean
```
Property Value -1 or true to reverse direction, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWHeatFlux.SetCutOffTemperatures
Sets the thermostat lower- and upper-bound temperatures for this heat flux.
```vb
Sub SetCutOffTemperatures( ByVal DVal1 As System.Double, ByVal DVal2 As System.Double )
```
Parameters DVal1 Lower-bound temperature DVal2 Upper-bound temperature
VBA: See CWHeatFlux::SetCutOffTemperatures .
Remarks: Thermostat is used for transient thermal studies only.
Examples: Visual Basic for Applications

### ICWHeatFlux.SetTemperatureCurve
Obsolete. Superseded by ICWHeatFlux::SetTemperatureCurve2 .
```vb
Function SetTemperatureCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters VarCurveData Array of temperature curve data (see Remarks ) ErrorCode Temperature curve error as defined in swsTemperatureCurveError_e Return Value 1 to use temperature curve, 0 to not
Return Value 1 to use temperature curve, 0 to not
VBA: See CWHeatFlux::SetTemperatureCurve .
Remarks: Temperature curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi
Examples: Visual Basic for Applications

### ICWHeatFlux.SetTemperatureCurve2
Sets the temperature curve data for heat flux.
```vb
Function SetTemperatureCurve2( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters VarCurveData Array of temperature curve data (see Remarks ) ErrorCode Temperature curve error as defined in swsTemperatureCurveError_e Return Value -1 or true to use temperature curve, 0 or false to not
Return Value -1 or true to use temperature curve, 0 or false to not
Remarks: Temperature curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi

### ICWHeatFlux.SetThermostat
Sets the location of the thermostat for heat flux used in transient thermal studies.
```vb
Sub SetThermostat( ByVal DispCoord As System.Object )
```
Parameters DispCoord Vertex of the location of the thermostat
VBA: See CWHeatFlux::SetThermostat .
Remarks: Thermostat is used for transient thermal studies only.
Examples: Visual Basic for Applications

### ICWHeatFlux.SetTimeCurve
Obsolete. Superseded by ICWHeatFlux::SetTimeCurve2 .
```vb
Function SetTimeCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters VarCurveData Array of time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value 1 to use curve, 0 to not
Return Value 1 to use curve, 0 to not
VBA: See CWHeatFlux::SetTimeCurve .
Remarks: Time curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi
Examples: Visual Basic for Applications

### ICWHeatFlux.SetTimeCurve2
Sets the time curve data for heat flux.
```vb
Function SetTimeCurve2( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters VarCurveData Array of time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value -1 or true to use curve, 0 or false to not
Return Value -1 or true to use curve, 0 or false to not
Remarks: Time curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi

### ICWHeatFlux.ThermostatUnits
Gets or sets the thermostat units for heat flux.
```vb
Property ThermostatUnits As System.Integer
```
Property Value Temperature unit as defined swsTemperatureUnit_e
VBA: See CWHeatFlux::ThermostatUnits .
Remarks: Thermostat is used for transient thermal studies only.
Examples: Visual Basic for Applications

### ICWHeatFlux.Unit
Gets or sets the temperature unit to use with this heat flux.
```vb
Property Unit As System.Integer
```
Property Value Temperature unit as defined swsTemperatureUnit_e
VBA: See CWHeatFlux::Unit .
Examples: Visual Basic for Applications

### ICWHeatFlux.UseTemperatureCurve
Obsolete. Superseded by ICWHeatFlux::UseTemperatureCurve2 .
```vb
Property UseTemperatureCurve As System.Integer
```
Property Value 1 to use temperature curve, 0 to not
VBA: See CWHeatFlux::UseTemperatureCurve .
Examples: Visual Basic for Applications

### ICWHeatFlux.UseTemperatureCurve2
Gets or sets whether to use a temperature curve for this heat flux.
```vb
Property UseTemperatureCurve2 As System.Boolean
```
Property Value -1 or true to use temperature curve, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWHeatFlux.UseTimeCurve
Obsolete. Superseded by ICWHeatFlux::UseTimeCurve2 .
```vb
Property UseTimeCurve As System.Integer
```
Property Value 1 to use time curve, 0 to not
VBA: See CWHeatFlux::UseTimeCurve .
Examples: Visual Basic for Applications

### ICWHeatFlux.UseTimeCurve2
Gets or sets whether to use a time curve for this heat flux.
```vb
Property UseTimeCurve2 As System.Boolean
```
Property Value -1 or true to use time curve, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

## ICWConvection

### ICWConvection.BulkAmbientTemperature
studies.
```vb
Property BulkAmbientTemperature As System.Double
```
Property Value Bulk ambient temperature
VBA: See CWConvection::BulkAmbientTemperature .
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power with Transient Thermal Study

### ICWConvection.ConvectionBeginEdit
Starts editing convection.
```vb
Sub ConvectionBeginEdit()
```
VBA: See CWConvection::ConvectionBeginEdit .
Remarks: To start editing convection, you must call this method. You must call ICWConvection::ConvectionEndEdit to end editing that convection. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power with Transient Thermal Study

### ICWConvection.ConvectionCoefficient
studies.
```vb
Property ConvectionCoefficient As System.Double
```
Property Value Convection coefficient
VBA: See CWConvection::ConvectionCoefficient .
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power with Transient Thermal Study

### ICWConvection.ConvectionEndEdit
Ends the editing convection.
```vb
Function ConvectionEndEdit() As System.Integer
```
Return Value Error as defined in swsConvectionEndEditError_e
VBA: See CWConvection::ConvectionEndEdit .
Remarks: You must call ICWConvection::ConvectionBeginEdit to start editing convection. To end editing convection, you must call this method. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power with Transient Thermal Study

### ICWConvection.GetBulkTemperatureTimeCurve
Gets the bulk temperature versus time curve data for convection.
```vb
Function GetBulkTemperatureTimeCurve() As System.Object
```
Return Value Array of bulk temperature-time curve (see Remarks )
VBA: See CWConvection::GetBulkTemperatureTimeCurve .
Remarks: Bulk temperature-time curve array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = bulk temperature at time xi
Examples: Visual Basic for Applications

### ICWConvection.GetTemperatureCurve
Gets the temperature curve data for convection.
```vb
Function GetTemperatureCurve() As System.Object
```
Return Value Array of temperature curve data for convection (see Remarks )
VBA: See CWConvection::GetTemperatureCurve .
Remarks: Temperature curve array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi
Examples: Visual Basic for Applications

### ICWConvection.GetTimeCurve
Gets the time curve data for convection.
```vb
Function GetTimeCurve() As System.Object
```
Return Value Array of time curve data for convection (see Remarks )
VBA: See CWConvection::GetTimeCurve .
Remarks: Time curve array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi
Examples: Visual Basic for Applications

### ICWConvection.InsertEntity
Inserts an entity for applying convection.
```vb
Sub InsertEntity( ByVal DispEntity As System.Object )
```
Parameters DispEntity Entity for applying convection
VBA: See CWConvection::InsertEntity .
Examples: Visual Basic for Applications

### ICWConvection.RemoveEntity
Removes the entity at the specified index to which to apply convection.
```vb
Sub RemoveEntity( ByVal NIndex As System.Integer )
```
Parameters NIndex 0-based index of entity to remove
VBA: See CWConvection::RemoveEntity .
Examples: Visual Basic for Applications

### ICWConvection.SetBulkTemperatureTimeCurve
Obsolete. Superseded by ICWConvection::SetBulkTemperatureTimeCurve2 .
```vb
Function SetBulkTemperatureTimeCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters VarCurveData Array of bulk temperature-time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value 1 to use curve, 0 to not
Return Value 1 to use curve, 0 to not
VBA: See CWConvection::SetBulkTemperatureTimeCurve .
Remarks: Bulk temperature-time curve array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = bulk temperature at time xi
Examples: Visual Basic for Applications

### ICWConvection.SetBulkTemperatureTimeCurve2
Sets the bulk temperature-time curve data.
```vb
Function SetBulkTemperatureTimeCurve2( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters VarCurveData Array of bulk temperature-time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value -1 or true to use curve, 0 or false to not
Return Value -1 or true to use curve, 0 or false to not
Remarks: Bulk temperature-time curve array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = bulk temperature at time xi

### ICWConvection.SetTemperatureCurve
Obsolete. Superseded by ICWConvection::SetTemperatureCurve2 .
```vb
Function SetTemperatureCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters VarCurveData Array of temperature curve data (see Remarks ) ErrorCode Temperature curve error as defined in swsTemperatureCurveError_e Return Value 1 to use curve, 0 to not
Return Value 1 to use curve, 0 to not
VBA: See CWConvection::SetTemperatureCurve .
Remarks: Temperature curve array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi
Examples: Visual Basic for Applications

### ICWConvection.SetTemperatureCurve2
Sets the temperature curve data for convection.
```vb
Function SetTemperatureCurve2( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters VarCurveData Array of temperature curve data (see Remarks ) ErrorCode Temperature curve error as defined in swsTemperatureCurveError_e Return Value -1 or true to use curve, 0 or false to not
Return Value -1 or true to use curve, 0 or false to not
Remarks: Temperature curve array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi

### ICWConvection.SetTimeCurve
Obsolete. Superseded by ICWConvection::SetTimeCurve2 .
```vb
Function SetTimeCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters VarCurveData Array of time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value 1 to use curve, 0 to not
Return Value 1 to use curve, 0 to not
VBA: See CWConvection::SetTimeCurve .
Remarks: Time curve array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi
Examples: Visual Basic for Applications

### ICWConvection.SetTimeCurve2
Sets the time curve data for convection.
```vb
Function SetTimeCurve2( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters VarCurveData Array of time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value -1 or true to use curve, 0 or false to not
Return Value -1 or true to use curve, 0 or false to not
Remarks: Time curve array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi

### ICWConvection.Unit
Gets or sets the units for convection.
```vb
Property Unit As System.Integer
```
Property Value Temperature units as defined swsTemperatureUnit_e
VBA: See CWConvection::Unit .
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power with Transient Thermal Study

### ICWConvection.UseBulkTemperatureTimeCurve
Obsolete. Superseded by ICWConvection::UseBulkTemperatureTimeCurve2 .
```vb
Property UseBulkTemperatureTimeCurve As System.Integer
```
Property Value 1 to use bulk temperature curve, 0 to not
VBA: See CWConvection::UseBulkTemperatureTimeCurve .
Examples: Visual Basic for Applications

### ICWConvection.UseBulkTemperatureTimeCurve2
Gets or sets whether to use a bulk temperature curve for convection.
```vb
Property UseBulkTemperatureTimeCurve2 As System.Boolean
```
Property Value -1 or true to use bulk temperature curve, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWConvection.UseTemperatureCurve
Obsolete. Superseded by ICWConvection::UseTemperatureCurve2 .
```vb
Property UseTemperatureCurve As System.Integer
```
Property Value 1 to use temperature curve, 0 to not
VBA: See CWConvection::UseTemperatureCurve .
Examples: Visual Basic for Applications

### ICWConvection.UseTemperatureCurve2
Gets or sets whether to use a temperature curve for convection.
```vb
Property UseTemperatureCurve2 As System.Boolean
```
Property Value -1 or true to use temperature curve, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWConvection.UseTimeCurve
Obsolete. Superseded by ICWConvection::UseTimeCurve2 .
```vb
Property UseTimeCurve As System.Integer
```
Property Value 1 to use a time curve, 0 to not
VBA: See CWConvection::UseTimeCurve .
Examples: Visual Basic for Applications

### ICWConvection.UseTimeCurve2
Gets or sets whether to use a time curve for convection.
```vb
Property UseTimeCurve2 As System.Boolean
```
Property Value -1 or true to use a time curve, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

## ICWRadiation

### ICWRadiation.AmbientTemperature
Gets or sets the value of ambient temperature for this radiation.
```vb
Property AmbientTemperature As System.Double
```
Property Value Ambient temperature
VBA: See CWRadiation::AmbientTemperature .
Examples: Visual Basic for Applications

### ICWRadiation.Emmisivity
Gets or sets the value for emissivity for this radiation.
```vb
Property Emmisivity As System.Double
```
Property Value Emissivity
VBA: See CWRadiation::Emmisivity .
Examples: Visual Basic for Applications

### ICWRadiation.GetTemperatureCurve
Gets the temperature curve data for this radiation.
```vb
Function GetTemperatureCurve() As System.Object
```
Return Value Array of temperature curve data (see Remarks )
VBA: See CWRadiation::GetTemperatureCurve .
Remarks: Temperature curve data array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi
Examples: Visual Basic for Applications

### ICWRadiation.GetTimeCurve
Gets the time curve data for this radiation.
```vb
Function GetTimeCurve() As System.Object
```
Return Value Array of time curve data (see Remarks )
VBA: See CWRadiation::GetTimeCurve .
Remarks: Array of time curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi
Examples: Visual Basic for Applications

### ICWRadiation.InsertEntity
Inserts a new entity for radiation.
```vb
Sub InsertEntity( ByVal DispEntity As System.Object )
```
Parameters DispEntity Entity
VBA: See CWRadiation::InsertEntity .
Examples: Visual Basic for Applications

### ICWRadiation.OpenSystem
Obsolete. Superseded by ICWRadiation::OpenSystem2 .
```vb
Property OpenSystem As System.Integer
```
Property Value System as defined in swsRadiationOpenSystem_e
VBA: See CWRadiation::OpenSystem .
Examples: Visual Basic for Applications

### ICWRadiation.OpenSystem2
Gets or sets whether to consider radiation to the ambient (open system).
```vb
Property OpenSystem2 As System.Boolean
```
Property Value -1 or true to consider radiation to the ambient (open system), 0 or false to ignore radiation to the ambient (closed system)
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWRadiation.RadiationBeginEdit
Starts editing radiation.
```vb
Sub RadiationBeginEdit()
```
VBA: See CWRadiation::RadiationBeginEdit .
Remarks: To start editing radiation, you must call this method. You must call ICWRadiation::RadiationEndEdit to end editing radiation. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications

### ICWRadiation.RadiationEndEdit
Ends editing radiation.
```vb
Function RadiationEndEdit() As System.Integer
```
Return Value Error as defined in swsRadiationEndEditError_e
VBA: See CWRadiation::RadiationEndEdit .
Remarks: You must call ICWRadiation::RadiationBeginEdit to start editing radiation. To end editing radiation, you must call this method. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications

### ICWRadiation.RadiationType
Gets or sets the type of radiation.
```vb
Property RadiationType As System.Integer
```
Property Value Radiation type as defined in swsRadiationType_e
VBA: See CWRadiation::RadiationType .
Examples: Visual Basic for Applications

### ICWRadiation.RemoveEntity
Removes the entity at the specified index from this radiation.
```vb
Sub RemoveEntity( ByVal NIndex As System.Integer )
```
Parameters NIndex 0-based index of entity to remove
VBA: See CWRadiation::RemoveEntity .
Examples: Visual Basic for Applications

### ICWRadiation.SetTemperatureCurve
Obsolete. Superseded by ICWRadiation::SetTemperatureCurve2 .
```vb
Function SetTemperatureCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters VarCurveData Array of temperature curve data (see Remarks ) ErrorCode Temperature curve error as defined in swsTemperatureCurveError_e
VBA: See CWRadiation::SetTemperatureCurve .
Remarks: Temperature curve data array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi
Examples: Visual Basic for Applications

### ICWRadiation.SetTemperatureCurve2
Sets the temperature curve data for this radiation.
```vb
Function SetTemperatureCurve2( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters VarCurveData Array of temperature curve data (see Remarks ) ErrorCode Temperature curve error as defined in swsTemperatureCurveError_e Return Value -1 or true to use curve, 0 or false to not
Return Value -1 or true to use curve, 0 or false to not
Remarks: Temperature curve data array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = temperature value at the ith data point yi = property value associated with temperature xi

### ICWRadiation.SetTimeCurve
Obsolete. Superseded by ICWRadiation::SetTimeCurve2 .
```vb
Function SetTimeCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters VarCurveData Array of time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e
VBA: See CWRadiation::SetTimeCurve .
Remarks: Array of time curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi
Examples: Visual Basic for Applications

### ICWRadiation.SetTimeCurve2
Sets the time curve data for this radiation.
```vb
Function SetTimeCurve2( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters VarCurveData Array of time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value -1 or true to use time curve, 0 or false to not
Return Value -1 or true to use time curve, 0 or false to not
Remarks: Array of time curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi

### ICWRadiation.Unit
Gets or sets the units of radiation.
```vb
Property Unit As System.Integer
```
Property Value Temperature units as defined in swsTemperatureUnit_e
VBA: See CWRadiation::Unit .
Examples: Visual Basic for Applications

### ICWRadiation.UseTemperatureCurve
Obsolete. Superseded by ICWRadiation::UseTemperatureCurve2 .
```vb
Property UseTemperatureCurve As System.Integer
```
Property Value 1 to use temperature curve, 0 to not
VBA: See CWRadiation::UseTemperatureCurve .
Examples: Visual Basic for Applications

### ICWRadiation.UseTemperatureCurve2
Gets or sets whether to use a temperature curve for this radiation.
```vb
Property UseTemperatureCurve2 As System.Boolean
```
Property Value -1 or true to use temperature curve, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWRadiation.UseTimeCurve
Obsolete. Superseded by ICWRadiation::UseTimeCurve2 .
```vb
Property UseTimeCurve As System.Integer
```
Property Value 1 to use time curve, 0 to not
VBA: See CWRadiation::UseTimeCurve .
Examples: Visual Basic for Applications

### ICWRadiation.UseTimeCurve2
Gets or sets whether to use a time curve for this radiation.
```vb
Property UseTimeCurve2 As System.Boolean
```
Property Value -1 or true to use time curve, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWRadiation.ViewFactor
Gets or sets the value of the view factor for this radiation.
```vb
Property ViewFactor As System.Double
```
Property Value Value of view factor
VBA: See CWRadiation::ViewFactor .
Examples: Visual Basic for Applications

## ICWTemperature

### ICWTemperature.GetTimeCurve
Gets the time curve data for the temperature.
```vb
Function GetTimeCurve() As System.Object
```
Return Value Array of time curve data (see Remarks )
VBA: See CWTemperature::GetTimeCurve .
Remarks: Time curve data array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi xi = time value at the ith data point yi = property value associated with time xi
Examples: Visual Basic for Applications

### ICWTemperature.InsertEntity
Inserts a new entity for the temperature.
```vb
Sub InsertEntity( ByVal DispEntity As System.Object )
```
Parameters DispEntity Entity
VBA: See CWTemperature::InsertEntity .
Examples: Visual Basic for Applications

### ICWTemperature.RemoveEntity
Removes the entity at the specified index to which to apply the temperature.
```vb
Sub RemoveEntity( ByVal NIndex As System.Integer )
```
Parameters NIndex 0-based index of the entity to remove
VBA: See CWTemperature::RemoveEntity .
Examples: Visual Basic for Applications

### ICWTemperature.SetTimeCurve
Obsolete. Superseded by ICWTemperature::SetTimeCurve2 .
```vb
Function SetTimeCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters VarCurveData Array of time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value 1 to use the time curve, 0 to not
Return Value 1 to use the time curve, 0 to not
VBA: See CWTemperature::SetTimeCurve .
Remarks: Time curve data array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi
Examples: Visual Basic for Applications

### ICWTemperature.SetTimeCurve2
Sets the time curve data.
```vb
Function SetTimeCurve2( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer ) As System.Boolean
```
Parameters VarCurveData Array of time curve data (see Remarks ) ErrorCode Error as defined in swsTimeCurveError_e Return Value -1 or true to use the time curve, 0 or false to not
Return Value -1 or true to use the time curve, 0 or false to not
Remarks: Time curve data array: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs xi = time value at the ith data point yi = property value associated with time xi

### ICWTemperature.TemperatureBeginEdit
Starts editing the temperature.
```vb
Sub TemperatureBeginEdit()
```
VBA: See CWTemperature::TemperatureBeginEdit .
Remarks: To start editing the temperature, you must call this method. You must call ICWTemperature::TemperatureEndEdit to end editing the temperature. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWTemperature.TemperatureEndEdit
Ends editing the temperature.
```vb
Function TemperatureEndEdit() As System.Integer
```
Return Value Error as defined in swsTemperatureEndEditError_e
VBA: See CWTemperature::TemperatureEndEdit .
Remarks: You must call ICWTemperature::TemperatureBeginEdit to start editing the temperature. To end editing the temperature, you must call this method. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWTemperature.TemperatureType
Gets or sets the temperature type for transient studies.
```vb
Property TemperatureType As System.Integer
```
Property Value Temperature type as defined in swsTemperatureType_e
VBA: See CWTemperature::TemperatureType .
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWTemperature.TemperatureValue
Gets or sets the temperature.
```vb
Property TemperatureValue As System.Double
```
Property Value Temperature
VBA: See CWTemperature::TemperatureValue .
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWTemperature.Unit
Gets or sets the units of temperature.
```vb
Property Unit As System.Integer
```
Property Value Temperature units as defined in swsTemperatureUnit_e
VBA: See CWTemperature::Unit .
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWTemperature.UseTimeCurve
Obsolete. Superseded by ICWTemperature::UseTimeCurve2 .
```vb
Property UseTimeCurve As System.Integer
```
Property Value 1 to use time curve data, 0 to not
VBA: See CWTemperature::UseTimeCurve .
Examples: Visual Basic for Applications

### ICWTemperature.UseTimeCurve2
Gets or sets whether to use time curve data with the temperature.
```vb
Property UseTimeCurve2 As System.Boolean
```
Property Value -1 or true to use time curve data, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

## ICWContactManager

### ICWContactManager.ContactComponentCount
Gets the number of component contacts in the active study.
```vb
ReadOnly Property ContactComponentCount As System.Integer
```
Property Value Number of component contacts
VBA: See CWContactManager::ContactComponentCount .
Examples: Visual Basic for Applications

### ICWContactManager.ContactSetCount
Gets the number of contact sets in the active study.
```vb
ReadOnly Property ContactSetCount As System.Integer
```
Property Value Number of contact sets
VBA: See CWContactManager::ContactSetCount .
Examples: Visual Basic for Applications; Example Copy Items to Another Study

### ICWContactManager.CopyContactsToStudy
Copies the specified contact sets to the specified study.
```vb
Function CopyContactsToStudy( ByVal SStudyName As System.String, ByVal VarContactNames As System.Object ) As System.Integer
```
Parameters SStudyName Name of study to which to copy contact sets VarContactNames Array of contact names (see Remarks ) Return Value Error code as defined in swsCopyItemsError_e
Return Value Error code as defined in swsCopyItemsError_e
VBA: See CWContactManager::CopyContactsToStudy .
Remarks: To populate the array of VarContactNames, either: inspect Connections > Contact Sets in the Simulation study tree. - or - call ICWContactManager::GetContactSetAt to get each ICWContactSet and then call ICWContactSet::ContactName to get the name of each contact set. If a contact set name already exists, Copy[1] is prepended to the name of the copied contact set.
Examples: Visual Basic for Applications; Example Copy Items to Another Study

### ICWContactManager.CreateContactComponent
Creates the specified component contact.
```vb
Function CreateContactComponent( ByVal NContactType As System.Integer, ByVal NOption As System.Integer, ByVal DispEntity As System.Object, ByRef ErrorCode As System.Integer ) As CWContactComponent
```
Parameters NContactType Type of component contact for: static , nonlinear , buckling and thermal studies as defined in swsContactType_e frequency studies as defined in swsContactType_e , excluding swsContactTypeStaticNoPenetration NOption Mesh compatibility as defined in swsMeshCompatibility_e DispEntity Component entity ErrorCode Error code as defined in swsContactComponentEndEditError_e Return Value Component contact
Return Value Component contact
VBA: See CWContactManager::CreateContactComponent .
Remarks: See the SOLIDWORKS Simulation Help for guidelines about studies with contact conditions.
Examples: Visual Basic for Applications; Example Add Component Contacts

### ICWContactManager.CreateContactSet
Obsolete. Superseded by ICWContactManager::CreateContactSet2 .
```vb
Function CreateContactSet( ByVal NContactSetType As System.Integer, ByVal ArraySourceEntities As System.Object, ByVal ArrayTargetEntities As System.Object, ByRef ErrorCode As System.Integer ) As CWContactSet
```
Parameters NContactSetType Type of contact set for static and nonlinear studies as defined in swsContactSetTypeStaticNonLinear_e thermal studies as defined in swsContactSetTypeThermal_e buckling and frequency studies as defined in swsContactType_e , excluding swsContactTypeStaticNoPenetration ArraySourceEntities Array of source entities ArrayTargetEntities Array of target faces ErrorCode Error code as defined in swsContactSetError_e Return Value Contact set
Return Value Contact set
VBA: See CWContactManager::CreateContactSet .
Remarks: Contact set definitions override global and component contact definitions. Component contact definitions override global contact definitions.
Examples: Visual Basic for Applications

### ICWContactManager.CreateContactSet2
Creates a contact set.
```vb
Function CreateContactSet2( ByVal NContactSetType As System.Integer, ByVal NSelectionType As System.Integer, ByVal ArraySourceEntities As System.Object, ByVal ArrayTargetEntities As System.Object, ByRef ErrorCode As System.Integer ) As CWContactSet
```
Parameters NContactSetType Type of contact set for: static and nonlinear studies as defined in swsContactSetTypeStaticNonLinear_e thermal studies as defined in swsContactSetTypeThermal_e buckling and frequency studies as defined in swsContactType_e , excluding swsContactTypeStaticNoPenetration NSelectionType Type of selection as defined in swsSelectionType_e ArraySourceEntities Array of source entities ArrayTargetEntities Array of target faces ErrorCode Error code as defined in swsContactSetError_e Return Value Contact set
Return Value Contact set
VBA: See CWContactManager::CreateContactSet2 .
Remarks: Contact set definitions override both global and component contact definitions. Component contact definitions override global contact definitions. See the SOLIDWORKS Simulation Help for guidelines about studies with contact conditions.
Examples: Visual Basic for Applications; Create Frequency Study with Mixed Mesh

### ICWContactManager.CreateContactSetsFromPairList
Creates contact sets from the specified contact pairs.
```vb
Function CreateContactSetsFromPairList( ByVal NContactSetType As System.Integer, ByVal VarPairs As System.Object, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NContactSetType Type of contact set for: static and nonlinear studies as defined in swsContactSetTypeStaticNonLinear_e thermal studies as defined in swsContactSetTypeThermal_e buckling and frequency studies as defined in swsContactType_e , excluding swsContactTypeStaticNoPenetration VarPairs Array of contact pairs (see Remarks ) ErrorCode Error code as defined in swsContactSetError_e Return Value Array of contact sets
Return Value Array of contact sets
VBA: See CWContactManager::CreateContactSetsFromPairList .
Remarks: To populate varPairs, call one of the following: ICWContactManager::FindNonTouchingPairs ICWContactManager::FindTouchingEdgeFacePairs ICWContactManager::FindTouchingFacePairs See the SOLIDWORKS Simulation Help for guidelines about studies with contact conditions.
Examples: Visual Basic for Applications

### ICWContactManager.DeleteContactComponent
Deletes the specified component contact.
```vb
Sub DeleteContactComponent( ByVal SName As System.String )
```
Parameters SName Name of the component contact to delete
VBA: See CWContactManager::DeleteContactComponent .
Examples: Visual Basic for Applications; Example Add Component Contacts

### ICWContactManager.DeleteContactSet
Deletes the specified contact set.
```vb
Sub DeleteContactSet( ByVal SName As System.String )
```
Parameters SName Name of contact set to delete
VBA: See CWContactManager::DeleteContactSet .
Examples: Visual Basic for Applications

### ICWContactManager.FindNonTouchingPairs
Finds non-touching faces within the specified minimum and maximum distance in the specified bodies or components.
```vb
Function FindNonTouchingPairs( ByVal VarCompBodies As System.Object, ByVal DMin As System.Double, ByVal DMax As System.Double, ByRef NEnts As System.Integer ) As System.Object
```
Parameters VarCompBodies Array of bodies or components DMin Minimum clearance; face pairs closer than this distance do not appear in the returned array DMax Maximum clearance; face pairs further apart than this distance do not appear in the returned array NEnts Number of non-touching faces Return Value Array of non-touching face pairs
Return Value Array of non-touching face pairs
VBA: See CWContactManager::FindNonTouchingPairs .
Remarks: nEnts contains the size of the returned array. The returned array contains consecutive pairs of non-touching faces. To create contact sets using the returned array, call ICWContactManager::CreateContactSetsFromPairList .
Examples: Visual Basic for Applications

### ICWContactManager.FindTouchingEdgeFacePairs
Finds touching shell edges and faces in the specified bodies or components.
```vb
Function FindTouchingEdgeFacePairs( ByVal VarCompBodies As System.Object, ByRef NEdgeFaces As System.Integer ) As System.Object
```
Parameters VarCompBodies Array of bodies or components NEdgeFaces Number of touching shell edges and faces Return Value Array of touching shell edge and face pairs
Return Value Array of touching shell edge and face pairs
VBA: See CWContactManager::FindTouchingEdgeFacePairs .
Remarks: NEdgeFaces contains the size of the returned array. The returned array contains consecutive pairs of touching entities. To create contact sets using the returned array, call ICWContactManager::CreateContactSetsFromPairList .
Examples: Visual Basic for Applications

### ICWContactManager.FindTouchingFacePairs
Finds touching faces in the specified bodies or components.
```vb
Function FindTouchingFacePairs( ByVal VarCompBodies As System.Object, ByRef NFaces As System.Integer ) As System.Object
```
Parameters VarCompBodies Array of bodies or components NFaces Number of touching faces Return Value Array of touching face pairs
Return Value Array of touching face pairs
VBA: See CWContactManager::FindTouchingFacePairs .
Remarks: NFaces contains the size of the returned array. The returned array contains consecutive pairs of touching faces. To create contact sets using the returned array, call ICWContactManager::CreateContactSetsFromPairList .
Examples: Visual Basic for Applications

### ICWContactManager.GetContactComponentAt
Gets the component contact at the specified index.
```vb
Function GetContactComponentAt( ByVal NIndex As System.Integer ) As CWContactComponent
```
Parameters NIndex 0-based index Return Value Component contact
Return Value Component contact
VBA: See CWContactManager::GetContactComponentAt .
Remarks: Before calling this method, call ICWContactManager::ContactComponentCount to determine the value of NIndex. If no contact is present at the specified index, then this method returns null with a non-zero error code.
Examples: Visual Basic for Applications

### ICWContactManager.GetContactSetAt
Gets the contact set at the specified index.
```vb
Function GetContactSetAt( ByVal NIndex As System.Integer ) As CWContactSet
```
Parameters NIndex 0-based index of contact set Return Value Contact set
Return Value Contact set
VBA: See CWContactManager::GetContactSetAt .
Remarks: Before calling this method, call ICWContactManager::ContactSetCount to determine the value of NIndex. If no contact set is present at the specified index, then this method returns Nothing or null with a non-0 error code.
Examples: Visual Basic for Applications; Example Copy Items to Another Study

### ICWContactManager.GetGlobalContact
Gets the contact type and options for global contact.
```vb
Sub GetGlobalContact( ByRef NContactType As System.Integer, ByRef NOption As System.Integer )
```
Parameters NContactType Type of global contact: static and nonlinear studies as defined in swsContactSetTypeStaticNonLinear_e thermal studies as defined in swsContactSetTypeThermal_e buckling and frequency studies as defined swsContactType_e , excluding swsContactTypeStaticNoPenetration NOption Mesh compatibility as defined in swsMeshCompatibility_e
VBA: See CWContactManager::GetGlobalContact .
Remarks: Contact set definitions override both global and component contact definitions. Component contact definitions override global contact definitions.
Examples: Visual Basic for Applications

### ICWContactManager.SetGlobalContact
Sets the contact type and options for global contacts.
```vb
Sub SetGlobalContact( ByVal NContactType As System.Integer, ByVal NOption As System.Integer )
```
Parameters NContactType Type global contact: static and nonlinear studies as defined in swsContactSetTypeStaticNonLinear_e thermal studies as defined in swsContactSetTypeThermal_e buckling and frequency studies as defined swsContactType_e , excluding swsContactTypeStaticNoPenetration NOption Mesh compatibility as defined in swsMeshCompatibility_e
VBA: See CWContactManager::SetGlobalContact .
Remarks: Contact set definitions override both global and component contact definitions. Component contact definitions override global contact definitions.
Examples: Visual Basic for Applications

### ICWContactManager.SuppressUnsuppressComponentContact
Obsolete. Superseded by ICWContactManager::SuppressUnsuppressComponentContact2 .
```vb
Function SuppressUnsuppressComponentContact( ByVal SName As System.String, ByVal BSuppress As System.Integer ) As System.Integer
```
Parameters SName Name of component contact (see Remarks ) BSuppress 1 to suppress, 0 to unsuppress (see Remarks ) Return Value Error code as defined in swsContactSuppressUnsuppressError_e
Return Value Error code as defined in swsContactSuppressUnsuppressError_e
VBA: See CWContactManager::SuppressUnsuppressComponentContact .
Remarks: Before calling this method, use: ICWContactComponent::ContactName to get the name of the component contact. ICWContactComponent:State to get the current suppression state of the component contact.
Examples: Visual Basic for Applications; Example Add Component Contacts

### ICWContactManager.SuppressUnsuppressComponentContact2
Sets the suppression state of the specified component contact.
```vb
Function SuppressUnsuppressComponentContact2( ByVal SName As System.String, ByVal BSuppress As System.Boolean ) As System.Integer
```
Parameters SName Name of component contact (see Remarks ) BSuppress -1 or true to suppress, 0 or false to unsuppress Return Value Error code as defined in swsContactSuppressUnsuppressError_e
Return Value Error code as defined in swsContactSuppressUnsuppressError_e
VBA: See CWContactManager::SuppressUnsuppressComponentContact2 .
Remarks: Before calling this method, use: ICWContactComponent::ContactName to get the name of the component contact. ICWContactComponent:State to get the current suppression state of the component contact.
Examples: Visual Basic for Applications

### ICWContactManager.SuppressUnsuppressContactPair
Obsolete. Superseded by ICWContactManager::SuppressUnsuppressContactPair2 .
```vb
Function SuppressUnsuppressContactPair( ByVal SName As System.String, ByVal BSuppress As System.Integer ) As System.Integer
```
Parameters SName Name of contact set (see Remarks ) BSuppress 1 to suppress, 0 to unsuppress (see Remarks ) Return Value Error code as defined in swsContactSuppressUnsuppressError_e
Return Value Error code as defined in swsContactSuppressUnsuppressError_e
VBA: See CWContactManager::SuppressUnsuppressContactPair .
Remarks: Before calling this method, use: ICWContactSet::ContactName to get the name of the contact set. ICWContactSet::State to get the current suppression state of the contact set.
Examples: Visual Basic for Applications

### ICWContactManager.SuppressUnsuppressContactPair2
Sets the suppression state of the specified contact set.
```vb
Function SuppressUnsuppressContactPair2( ByVal SName As System.String, ByVal BSuppress As System.Boolean ) As System.Integer
```
Parameters SName Name of contact set (see Remarks ) BSuppress -1 or true to suppress, 0 or false to unsuppress Return Value Error code as defined in swsContactSuppressUnsuppressError_e
Return Value Error code as defined in swsContactSuppressUnsuppressError_e
VBA: See CWContactManager::SuppressUnsuppressContactPair2 .
Remarks: Before calling this method, use: ICWContactSet::ContactName to get the name of the contact set. ICWContactSet::State to get the current suppression state of the contact set.
Examples: Visual Basic for Applications

## ICWContactSet

### ICWContactSet.AxialStiffnessValue
Gets or sets the wall axial stiffness value for this flexible virtual wall contact set.
```vb
Property AxialStiffnessValue As System.Double
```
Property Value Wall axial stiffness value
VBA: See CWContactSet::AxialStiffnessValue .
Remarks: This property is valid only if ICWContactSet::ContactSetType is set to swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearVirtualWall and ICWContactSet::WallType is set to swsWallType_e .swsWallTypeFlexible.
Examples: Visual Basic for Applications

### ICWContactSet.ClearanceUnit
Gets or sets the units of clearance for this contact set.
```vb
Property ClearanceUnit As System.Integer
```
Property Value Units of clearance as defined in swsLinearUnit_e
VBA: See CWContactSet::ClearanceUnit .
Remarks: This property is valid only if ICWContactSet::ContactSetType is one of: swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearVirtualWall swsContactSetTypeStaticNonLinear_e.swsContactSetTypeStaticNonLinearNoPenetration
Examples: Visual Basic for Applications

### ICWContactSet.ClearanceValue
Gets or sets the clearance (gap) value for this contact set.
```vb
Property ClearanceValue As System.Double
```
Property Value Clearance value
VBA: See CWContactSet::ClearanceValue .
Remarks: This property is valid only if ICWContactSet::ContactSetType is one of: swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearVirtualWall swsContactSetTypeStaticNonLinear_e.swsContactSetTypeStaticNonLinearNoPenetration
Examples: Visual Basic for Applications

### ICWContactSet.ContactName
Gets the name of this contact set.
```vb
ReadOnly Property ContactName As System.String
```
Property Value Name of this contact set
VBA: See CWContactSet::ContactName .
Examples: Visual Basic for Applications

### ICWContactSet.ContactSetBeginEdit
Starts editing this contact set.
```vb
Sub ContactSetBeginEdit()
```
VBA: See CWContactSet::ContactSetBeginEdit .
Remarks: To start editing a contact set, call this method. You must call ICWContactSet::ContactSetEndEit to end editing that contact set. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications

### ICWContactSet.ContactSetEndEdit
Ends editing this contact set.
```vb
Function ContactSetEndEdit() As System.Integer
```
Return Value Error as defined in swsContactSetEndEditError_e
VBA: See CWContactSet::ContactSetEndEdit .
Remarks: To start editing a contact set, call ICWContactSet::ContactSetBeginEdit . You must call this method to end editing that contact set. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications

### ICWContactSet.ContactSetType
Gets or sets the type of this contact set.
```vb
Property ContactSetType As System.Integer
```
Property Value Type of contact set for: static and nonlinear studies as defined in swsContactSetTypeStaticNonLinear_e thermal studies as defined in swsContactSetTypeThermal_e buckling and frequency studies as defined swsContactType_e , excluding swsContactTypeStaticNoPenetration
VBA: See CWContactSet::ContactSetType .
Remarks: See the SOLIDWORKS Simulation Help for guidelines about studies with contact conditions.
Examples: Visual Basic for Applications

### ICWContactSet.FrictionValue
Gets or sets the friction coefficient for this contact set.
```vb
Property FrictionValue As System.Double
```
Property Value 0.0 <= Coefficient of friction <= 1.0
VBA: See CWContactSet::FrictionValue .
Remarks: This property is valid only if ICWContactSet::ContactSetType is one of: swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearNoPrenetration swsContactSetTypeStaticNonLinear_e.swsContactSetTypeStaticNonLinearShrinkFit
Examples: Visual Basic for Applications

### ICWContactSet.GapType
Gets or sets the gap type associated with this contact set.
```vb
Property GapType As System.Integer
```
Property Value Type of gap as defined in swsGapType_e
VBA: See CWContactSet::GapType .
Remarks: This property is valid only if ICWContactSet::ContactSetType is one of: swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearVirtualWall swsContactSetTypeStaticNonLinear_e.swsContactSetTypeStaticNonLinearNoPenetration
Examples: Visual Basic for Applications

### ICWContactSet.GetSourceEntityAt
Gets the entity type and source entity of this contact set at the specified index.
```vb
Function GetSourceEntityAt( ByVal NIndex As System.Integer, ByRef NSel As System.Integer ) As System.Object
```
Parameters NIndex 0-based index of the entity NSel Type of entity (see Remarks ) Return Value Entity
Return Value Entity
VBA: See CWContactSet::GetSourceEntityAt .
Remarks: Before calling this method, call ICWContactSet::SourceEntityCount to get the value of NIndex. Faces, edges, and vertices are allowed as source entities. See swSelectType_e in the SOLIDWORKS API Enumerations Help for their types.
Examples: Visual Basic for Applications

### ICWContactSet.GetTargetEntityAt
Gets the entity type and target entity of this contact set at the specified index.
```vb
Function GetTargetEntityAt( ByVal NIndex As System.Integer, ByRef NSel As System.Integer ) As System.Object
```
Parameters NIndex 0-based index of the entity NSel Type of entity (see Remarks ) Return Value Entity
Return Value Entity
VBA: See CWContactSet::GetTargetEntityAt .
Remarks: Before calling this method, call ICWContactSet::TargetEntityCount to get the value of NIndex. Faces, edges, and vertices are allowed as target entities. See swSelectType_e in the SOLIDWORKS API Enumerations Help for their types.
Examples: Visual Basic for Applications

### ICWContactSet.IncludeFriction
Obsolete. Superseded by ICWContactSet::IncludeFriction2 .
```vb
Property IncludeFriction As System.Integer
```
Property Value 1 to include friction, 0 to not
VBA: See CWContactSet::IncludeFriction .
Remarks: This property is valid only if ICWContactSet::ContactSetType is one of: swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearNoPrenetration swsContactSetTypeStaticNonLinear_e.swsContactSetTypeStaticNonLinearShrinkFit
Examples: Visual Basic for Applications

### ICWContactSet.IncludeFriction2
Sets whether to include friction in this contact set.
```vb
Property IncludeFriction2 As System.Boolean
```
Property Value -1 or true to include friction, 0 or false to not
VBA: See CWContactSet::IncludeFriction2 .
Remarks: This property is valid only if ICWContactSet::ContactSetType is one of: swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearNoPrenetration swsContactSetTypeStaticNonLinear_e.swsContactSetTypeStaticNonLinearShrinkFit To set this property, you can specify either the boolean or the integer.
Examples: Visual Basic for Applications

### ICWContactSet.InsertSourceEntity
Inserts the specified source entity into this contact set.
```vb
Sub InsertSourceEntity( ByVal DispEntity As System.Object )
```
Parameters DispEntity Entity (see Remarks )
VBA: See CWContactSet::InsertSourceEntity .
Remarks: Faces, edges, and vertices are allowed as source entities. See swSelectType_e in the SOLIDWORKS API Enumerations Help for their types.
Examples: Visual Basic for Applications

### ICWContactSet.InsertTargetEntity
Inserts the specified target entity into this contact set.
```vb
Sub InsertTargetEntity( ByVal DispEntity As System.Object )
```
Parameters DispEntity Entity (see Remarks )
VBA: See CWContactSet::InsertTargetEntity .
Remarks: Faces, edges, and vertices are allowed as target entities. See swSelectType_e in the SOLIDWORKS API Enumerations Help for their types.
Examples: Visual Basic for Applications

### ICWContactSet.Options
Gets or sets the advanced options for this contact set.
```vb
Property Options As System.Integer
```
Property Value Option for no penetration ( static and nonlinear studies only) advanced options as defined in swsNoPenetrationOption_e thermal resistance ( thermal studies only) advanced options as defined in swsNoPenetrationOption_e (except swsNoPenetrationOptionNodeToNode) shrink fit (static and nonlinear studies only) advanced options as defined in swsShrinkFitOption_e
VBA: See CWContactSet::Options .
Examples: Visual Basic for Applications

### ICWContactSet.RemoveSourceEntity
Removes a source entity at the specified index from this contact set.
```vb
Sub RemoveSourceEntity( ByVal NIndex As System.Integer )
```
Parameters NIndex 0-based index of entity
VBA: See CWContactSet::RemoveSourceEntity .
Remarks: Before calling this method, call ICWContactSet::SourceEntityCount to get the value of NIndex.
Examples: Visual Basic for Applications

### ICWContactSet.RemoveTargetEntity
Removes a target entity at the specified index from this contact set.
```vb
Sub RemoveTargetEntity( ByVal NIndex As System.Integer )
```
Parameters NIndex 0-based index of entity
VBA: See CWContactSet::RemoveTargetEntity .
Remarks: Before calling this method, call ICWContactSet::TargetEntityCount to get the value of NIndex.
Examples: Visual Basic for Applications

### ICWContactSet.ResistanceType
Gets or sets the thermal resistance type for this thermal resistance contact set.
```vb
Property ResistanceType As System.Integer
```
Property Value Thermal resistance type as defined in swsResistanceType_e
VBA: See CWContactSet::ResistanceType .
Remarks: This property is valid only for thermal studies and when ICWContactSet::ContactSetType is set to swsContactSetTypeThermal_e .swsContactSetTypeThermalResistance.
Examples: Visual Basic for Applications

### ICWContactSet.ResistanceUnit
Gets or sets the unit system for this thermal resistance contact set.
```vb
Property ResistanceUnit As System.Integer
```
Property Value Unit system as defined in swsUnitSystem_e
VBA: See CWContactSet::ResistanceUnit .
Remarks: This property is valid only for thermal studies and when ICWContactSet::ContactSetType is set to swsContactSetTypeThermal_e .swsContactSetTypeThermalResistance.
Examples: Visual Basic for Applications

### ICWContactSet.ResistanceValue
Gets or sets the thermal resistance value for this thermal resistance contact set.
```vb
Property ResistanceValue As System.Double
```
Property Value Thermal resistance value
VBA: See CWContactSet::ResistanceValue .
Remarks: This property is valid only for thermal studies and when ICWContactSet::ContactSetType is set to swsContactSetTypeThermal_e .swsContactSetTypeThermalResistance.
Examples: Visual Basic for Applications

### ICWContactSet.SourceEntityCount
Gets the number of source entities in this contact set.
```vb
ReadOnly Property SourceEntityCount As System.Integer
```
Property Value Number of source entities
VBA: See CWContactSet::SourceEntityCount . Requirements Call this method before calling ICWContactSet::GetSourceEntityAt and ICWContactSet::RemoveSourceEntity to determine a valid index for each method.
Examples: Visual Basic for Applications

### ICWContactSet.State
Gets the suppression state of this contact set.
```vb
ReadOnly Property State As System.Integer
```
Property Value Suppression state as defined in swsSuppressionState_e
VBA: See CWContactSet::State .
Examples: Visual Basic for Applications

### ICWContactSet.SuppressUnSuppress
Suppresses or unsuppresses this contact set depending on its state .
```vb
Sub SuppressUnSuppress()
```
VBA: See CWContactSet::SuppressUnSuppress .
Examples: Visual Basic for Applications

### ICWContactSet.TangentialStiffnessValue
Gets or sets the wall shear stiffness value for this flexible virtual wall contact set.
```vb
Property TangentialStiffnessValue As System.Double
```
Property Value Wall shear stiffness value
VBA: See CWContactSet::TangentialStiffnessValue .
Remarks: This property is valid only if ICWContactSet::ContactSetType is set to swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearVirtualWall and ICWContactSet::WallType is set to swsWallType_e .swsWallTypeFlexible.
Examples: Visual Basic for Applications

### ICWContactSet.TargetEntityCount
Gets the number of target entities in this contact set.
```vb
ReadOnly Property TargetEntityCount As System.Integer
```
Property Value Number of target entities
VBA: See CWContactSet::TargetEntityCount .
Remarks: Call this method before calling ICWContactSet::GetTargetEntityAt and ICWContactSet::RemoveTargetEntity to determine a valid index for each method.
Examples: Visual Basic for Applications

### ICWContactSet.WallFriction
Gets or sets the coefficient of friction for this contact set with a rigid virtual wall.
```vb
Property WallFriction As System.Double
```
Property Value 0.0 <= Coefficient of friction for rigid wall types <= 1.0
VBA: See CWContactSet::WallFriction .
Remarks: This property is valid only if ICWContactSet::ContactSetType is set to swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearVirtualWall, and ICWContactSet::WallType is set to swsWallType_e .swsWallTypeRigid.
Examples: Visual Basic for Applications

### ICWContactSet.WallStiffnessUnit
Gets or sets the wall stiffness units of this contact set.
```vb
Property WallStiffnessUnit As System.Integer
```
Property Value Units as defined in swsUnitSystem_e
VBA: See CWContactSet::WallStiffnessUnit .
Remarks: This property is valid only if ICWContactSet::ContactSetType is set to swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearVirtualWall and ICWContactSet::WallType is set to swsWallType_e .swsWallTypeFlexible.
Examples: Visual Basic for Applications

### ICWContactSet.WallType
Gets or sets the wall type of this contact set.
```vb
Property WallType As System.Integer
```
Property Value Virtual wall type as defined in swsWallType_e
VBA: See CWContactSet::WallType .
Remarks: This property is valid only if ICWContactSet::ContactSetType is set to swsContactSetTypeStaticNonLinear_e .swsContactSetTypeStaticNonLinearVirtualWall.
Examples: Visual Basic for Applications

## ICWContactComponent

### ICWContactComponent.BondingFormulation
Sets the bonding formulation for this contact.
```vb
WriteOnly Property BondingFormulation As System.Integer
```
Property Value Option for no penetration ( static and nonlinear studies only) advanced options as defined in swsNoPenetrationOption_e (except swsNoPenetrationOptionNodeToNode) thermal resistance ( thermal studies only) advanced options as defined in swsNoPenetrationOption_e (except swsNoPenetrationOptionNodeToNode) shrink fit (static and nonlinear studies only) advanced options as defined in swsShrinkFitOption_e
VBA: See CWContactComponent::BondingFormulation .
Remarks: This property is valid only for selected components that behave during simulation as if they were welded. For more information, see SOLIDWORKS user-interface Help > Simulation > Interaction Options > : Component Interaction PropertyManager - and - Types of Interactions > Bonded Interaction
Examples: Visual Basic for Applications

### ICWContactComponent.ClearanceUnit
Gets or sets the units of clearance between non-touching faces.
```vb
Property ClearanceUnit As System.Integer
```
Property Value Units of clearance as defined in swsPinballUnit_e
VBA: See CWContactComponent::ClearanceUnit .
Remarks: This property is valid only if... Is set to... ICWContactComponent::ContactComponentType swsContactType_e .swsContactTypeBonded ICWContactComponent::IncludeClearance 1
Examples: Visual Basic for Applications

### ICWContactComponent.ClearanceValue
Gets or sets the maximum clearance (gap) between non-touching faces or shell edges.
```vb
Property ClearanceValue As System.Double
```
Property Value Maximum clearance (gap) between non-touching faces or shell edges
VBA: See CWContactComponent::ClearanceValue .
Remarks: This property is valid only if... Is set to... ICWContactComponent::ContactComponentType swsContactType_e .swsContactTypeBonded ICWContactComponent::IncludeClearance 1
Examples: Visual Basic for Applications

### ICWContactComponent.ContactComponentBeginEdit
Start editing of this contact.
```vb
Sub ContactComponentBeginEdit()
```
VBA: See CWContactComponent::ContactComponentBeginEdit .
Remarks: To start editing a contact, you must call this method. You must call ICWContactComponent::ContactComponentEndEdit to end editing that contact. Changes are not applied unless you call both methods. Contact sets override both global and component contacts. Component contacts override global contacts.
Examples: Visual Basic for Applications

### ICWContactComponent.ContactComponentEndEdit
Ends editing of this contact.
```vb
Function ContactComponentEndEdit() As System.Integer
```
Return Value Error as defined in swsContactComponentEndEditError_e
VBA: See CWContactComponent::ContactComponentEndEdit .
Remarks: To start editing a contact, call ICWContactComponent::ContactComponentBeginEdit . You must call this method to end editing that contact. Changes are not applied unless you call both methods. Contact sets override both global and component contacts. Component contacts override global contacts.
Examples: Visual Basic for Applications

### ICWContactComponent.ContactComponentType
Gets or sets the type of contact.
```vb
Property ContactComponentType As System.Integer
```
Property Value Type of contact for: static , nonlinear , buckling and thermal studies as defined in swsContactType_e frequency studies as defined in swsContactType_e , excluding swsContactTypeStaticNoPenetration
VBA: See CWContactComponent::ContactComponentType .
Remarks: See the SOLIDWORKS Simulation Help for guidelines about studies with contact conditions.
Examples: Visual Basic for Applications

### ICWContactComponent.ContactName
Gets the name of the contact.
```vb
ReadOnly Property ContactName As System.String
```
Property Value Name of contact
VBA: See CWContactComponent::ContactName .
Examples: Visual Basic for Applications

### ICWContactComponent.FrictionValue
Gets or sets the friction coefficient for this contact.
```vb
Property FrictionValue As System.Double
```
Property Value 0.0 <= Coefficient of friction for contact <= 1.0
VBA: See CWContactComponent::FrictionValue .
Remarks: This property is valid only in static and nonlinear studies where ICWContactComponent::IncludeFriction is set to 1, and ICWContactComponent::ContactComponentType is set to swsContactType_e .swsContactTypeStaticNoPenetration.
Examples: Visual Basic for Applications

### ICWContactComponent.GlobalContact
Obsolete. Superseded by ICWContactComponent::GlobalContact2 .
```vb
Property GlobalContact As System.Integer
```
Property Value 1 to apply a global contact condition, 0 to not
VBA: See CWContactComponent::GlobalContact .
Remarks: Set this property to 1 to apply a bonded contact to all components in an assembly. If you set this property to 0, call ICWContactComponent::InsertEntity to specify contact components.
Examples: Visual Basic for Applications

### ICWContactComponent.GlobalContact2
Sets whether to apply a global contact condition.
```vb
Property GlobalContact2 As System.Boolean
```
Property Value -1 or true to apply a global contact condition, 0 or false to not
VBA: See CWContactComponent::GlobalContact2 .
Remarks: To set this property, you can specify either the boolean or the integer. Set this property to -1 or true to apply a bonded contact to all components in an assembly. If you set this property to 0 or false, call ICWContactComponent::InsertEntity to specify contact components.
Examples: Visual Basic for Applications

### ICWContactComponent.IncludeClearance
Obsolete. Superseded by ICWContactComponent::IncludeClearance2 .
```vb
Property IncludeClearance As System.Integer
```
Property Value 1 to specify clearance for non-touching faces, 0 to not
VBA: See CWContactComponent::IncludeClearance .
Examples: Visual Basic for Applications

### ICWContactComponent.IncludeClearance2
Sets whether to specify clearance for non-touching faces in this contact.
```vb
Property IncludeClearance2 As System.Boolean
```
Property Value -1 or true to specify clearance for non-touching faces, 0 or false to not
VBA: See CWContactComponent::IncludeClearance2 .
Remarks: To set this property, you can specify either the boolean or the integer.
Examples: Visual Basic for Applications

### ICWContactComponent.IncludeFriction
Obsolete. Superseded by ICWContactComponent::IncludeFriction2 .
```vb
Property IncludeFriction As System.Integer
```
Property Value 1 to specify friction, 0 to not
VBA: See CWContactComponent::IncludeFriction .
Remarks: This property is valid only in static and nonlinear studies where ICWContactComponent::ContactComponentType is set to swsContactType_e .swsContactTypeStaticNoPenetration.
Examples: Visual Basic for Applications

### ICWContactComponent.IncludeFriction2
Sets whether to specify a friction coefficient for this contact.
```vb
Property IncludeFriction2 As System.Boolean
```
Property Value -1 or true to include friction, 0 or false to not
VBA: See CWContactComponent::IncludeFriction2 .
Remarks: This property is valid only in static and nonlinear studies where ICWContactComponent::ContactComponentType is set to swsContactType_e .swsContactTypeStaticNoPenetration. To set this property, you can specify either the boolean or the integer.
Examples: Visual Basic for Applications

### ICWContactComponent.IncludeShellEdgeSolidOrShellFace
Obsolete. Superseded by ICWContactComponent::IncludeShellEdgeSolidOrShellFace2 .
```vb
Property IncludeShellEdgeSolidOrShellFace As System.Integer
```
Property Value 1 to create contact sets, 0 to not
VBA: See CWContactComponent::IncludeShellEdgeSolidOrShellFace .
Remarks: This property sets whether to create bonded contact sets for shell edge pairs, shell edge-to-shell face pairs, and shell edge-to-solid face pairs that are non-touching and within clearance. This property is valid only if... Is set to... ICWContactComponent::ContactComponentType swsContactType_e .swsContactTypeBonded ICWContactComponent::IncludeClearance 1
Examples: Visual Basic for Applications

### ICWContactComponent.IncludeShellEdgeSolidOrShellFace2
Sets whether to create contact sets.
```vb
Property IncludeShellEdgeSolidOrShellFace2 As System.Boolean
```
Property Value -1 or true to create contact sets, 0 or false to not
VBA: See CWContactComponent::IncludeShellEdgeSolidOrShellFace2 .
Remarks: To set this property, you can specify either the boolean or the integer. This property sets whether to create bonded contact sets for shell edge pairs, shell edge-to-shell face pairs, and shell edge-to-solid face pairs that are non-touching and within clearance. This property is valid only if... Is set to... ICWContactComponent::ContactComponentType swsContactType_e .swsContactTypeBonded ICWContactComponent::IncludeClearance2 -1 or true
Examples: Visual Basic for Applications

### ICWContactComponent.InsertEntity
Inserts a component into this contact.
```vb
Sub InsertEntity( ByVal DispEntity As System.Object )
```
Parameters DispEntity Component to insert
VBA: See CWContactComponent::InsertEntity .
Remarks: This method is valid only if ICWContactComponent::GlobalContact is set to 0.
Examples: Visual Basic for Applications

### ICWContactComponent.Option
Gets or sets the mesh compatibility for this contact.
```vb
Property Option As System.Integer
```
Property Value Mesh compatibility as defined in swsMeshCompatibility_e
VBA: See CWContactComponent::Option .
Examples: Visual Basic for Applications

### ICWContactComponent.RemoveEntity
Removes a component from this contact.
```vb
Sub RemoveEntity( ByVal NIndex As System.Integer )
```
Parameters NIndex Index of component to remove
VBA: See CWContactComponent::RemoveEntity .
Examples: Visual Basic for Applications

### ICWContactComponent.ReplaceEntity
Replaces the entity of this contact.
```vb
Sub ReplaceEntity( ByVal DispEntity As System.Object )
```
Parameters DispEntity Replacing component
VBA: See CWContactComponent::ReplaceEntity .
Examples: Visual Basic for Applications

### ICWContactComponent.State
Gets the suppression state of this contact.
```vb
ReadOnly Property State As System.Integer
```
Property Value Suppression state as defined in swsSuppressionState_e
VBA: See CWContactComponent::State .
Examples: Visual Basic for Applications

### ICWContactComponent.SuppressUnSuppress
Suppresses or unsuppresses this contact depending on its state .
```vb
Sub SuppressUnSuppress()
```
VBA: See CWContactComponent::SuppressUnSuppress .
Examples: Visual Basic for Applications

## ICWSolidManager

### ICWSolidManager.ComponentCount
Gets the number of solid components.
```vb
ReadOnly Property ComponentCount As System.Integer
```
Property Value Number of solid components
VBA: See CWSolidManager::ComponentCount .
Examples: Visual Basic for Applications

### ICWSolidManager.GetComponentAt
Gets the solid component at the specified index.
```vb
Function GetComponentAt( ByVal NIndex As System.Integer, ByRef ErrorCode As System.Integer ) As CWSolidComponent
```
Parameters NIndex 0-based index of the solid component to get ErrorCode 0 if successful, 1 if no solid component at NIndex Return Value Solid component or NULL if no solid component at NIndex
Return Value Solid component or NULL if no solid component at NIndex
VBA: See CWSolidManager::GetComponentAt .
Remarks: Before calling this method, call ICWSolidManager::ComponentCount to get NIndex. Index starts at 0.
Examples: Visual Basic for Applications

### ICWSolidManager.GetSolidBodyByName
Gets the specified solid body.
```vb
Function GetSolidBodyByName( ByVal SBodyName As System.String, ByRef ErrorCode As System.Integer ) As CWSolidBody
```
Parameters SBodyName Name of solid body ErrorCode 0 if successful, 1 if not Return Value ICWSolidBody
Return Value ICWSolidBody
VBA: See CWSolidManager::GetSolidBodyByName .
Examples: Visual Basic for Applications

### ICWSolidManager.SetFavLibraryMaterialToAllBodiesByCompName
Applies the specified favorite library material to all bodies in the specified component.
```vb
Function SetFavLibraryMaterialToAllBodiesByCompName( ByVal SCompName As System.String, ByVal SFavLibName As System.String, ByVal SFavMaterialName As System.String ) As System.Integer
```
Parameters SCompName Name of component SFavLibName Name of favorite material library SFavMaterialName Name of material in SFavLibName Return Value 0 if material library and material name are set, 1 if not
Return Value 0 if material library and material name are set, 1 if not
VBA: See CWSolidManager::SetFavLibraryMaterialToAllBodiesByCompName .
Examples: Visual Basic for Applications

### ICWSolidManager.SetFavLibraryMaterialToSelectedEntities
Applies the specified favorite library material to selected entities.
```vb
Function SetFavLibraryMaterialToSelectedEntities( ByVal SFavLibName As System.String, ByVal SFavMaterialName As System.String ) As System.Integer
```
Parameters SFavLibName Name of favorite material library SFavMaterialName Name of material in SFavLibName Return Value 0 if material library and material name are set, 1 if not
Return Value 0 if material library and material name are set, 1 if not
VBA: See CWSolidManager::SetFavLibraryMaterialToSelectedEntities .
Remarks: Before calling this method, select one or more components to which to apply the specified favorite library material.
Examples: Visual Basic for Applications

### ICWSolidManager.SetLibraryMaterialToAllBodiesByCompName
Applies the specified library material to all bodies in the specified component.
```vb
Function SetLibraryMaterialToAllBodiesByCompName( ByVal SCompName As System.String, ByVal SLibName As System.String, ByVal SMaterialName As System.String ) As System.Integer
```
Parameters SCompName Name of component SLibName Name of material library SMaterialName Name of material in SLibName Return Value 0 if material library and material name are set, 1 if not
Return Value 0 if material library and material name are set, 1 if not
VBA: See CWSolidManager::SetLibraryMaterialToAllBodiesByCompName .
Examples: Visual Basic for Applications

### ICWSolidManager.SetLibraryMaterialToSelectedEntities
Applies the specified library material to selected entities.
```vb
Function SetLibraryMaterialToSelectedEntities( ByVal SLibName As System.String, ByVal SMaterialName As System.String ) As System.Integer
```
Parameters SLibName Name of material library SMaterialName Name of material in SLibName Return Value 0 if material library and material name are set, 1 if not
Return Value 0 if material library and material name are set, 1 if not
VBA: See CWSolidManager::SetLibraryMaterialToSelectedEntities .
Remarks: Before calling this method, select one or more components to which to apply the specified library material.
Examples: Visual Basic for Applications; Example Apply Material to Selected Entities

## ICWSolidBody

### ICWSolidBody.ConvertToBeamBody
Treat a solid body as a beam.
```vb
Function ConvertToBeamBody() As System.Integer
```
Return Value 0 to treat a solid body as a beam, 1 to not
VBA: See CWSolidBody::ConvertToBeamBody .
Examples: Visual Basic for Applications; Change Beam to Solid Body and Back

### ICWSolidBody.GetDefaultMaterial
Gets the CAD material of the solid body.
```vb
Function GetDefaultMaterial() As CWMaterial
```
Return Value Material
VBA: See CWSolidBody::GetDefaultMaterial .
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Materials

### ICWSolidBody.GetSolidBodyMaterial
Gets the material applied to the solid body for analysis.
```vb
Function GetSolidBodyMaterial() As CWMaterial
```
Return Value Material applied to the solid body or NULL if no material object was applied
VBA: See CWSolidBody::GetSolidBodyMaterial .
Examples: Visual Basic for Applications

### ICWSolidBody.SetFavMaterial
Obsolete. Superseded by ICWSolidBody::SetFavMaterial2 .
```vb
Function SetFavMaterial( ByVal NFav As System.Integer ) As System.Integer
```
Parameters NFav 1 <= index of material in favorites list <= 100 (see Remarks ) Return Value 0 if material applied successfully, 1 if not
Return Value 0 if material applied successfully, 1 if not
VBA: See CWSolidBody::SetFavMaterial .
Remarks: The list of material favorites appears when you right-click a component in the Simulation study tree and click Apply Favorite Material .
Examples: Visual Basic for Applications

### ICWSolidBody.SetFavMaterial2
Applies the specified material from the material favorites list.
```vb
Function SetFavMaterial2( ByVal NFav As System.Integer ) As System.Boolean
```
Parameters NFav 1 <= index of material in favorites list <= 100 (see Remarks ) Return Value -1 or true if material applied successfully, 0 or false if not
Return Value -1 or true if material applied successfully, 0 or false if not
Remarks: The list of material favorites appears when you right-click a component in the Simulation study tree and click Apply Favorite Material .

### ICWSolidBody.SetLibraryMaterial
Obsolete. Superseded by ICWSolidBody::SetLibraryMaterial2 .
```vb
Function SetLibraryMaterial( ByVal SLibWithPathName As System.String, ByVal SMaterialName As System.String ) As System.Integer
```
Parameters SLibWithPathName Path to the material library folder SMaterialName Material name Return Value 0 if material library and name are set, 1 if not
Return Value 0 if material library and name are set, 1 if not
VBA: See CWSolidBody::SetLibraryMaterial .
Examples: Visual Basic for Applications

### ICWSolidBody.SetLibraryMaterial2
Sets the material library for the solid body.
```vb
Function SetLibraryMaterial2( ByVal SLibWithPathName As System.String, ByVal SMaterialName As System.String ) As System.Boolean
```
Parameters SLibWithPathName Path to the material library folder SMaterialName Material name Return Value -1 or true if material library and name are set, 0 or false if not
Return Value -1 or true if material library and name are set, 0 or false if not

### ICWSolidBody.SetMeshQuality
Sets the mesh quality for the solid body.
```vb
Function SetMeshQuality( ByVal NMeshQuality As System.Integer ) As System.Boolean
```
Parameters NMeshQuality Mesh quality as defined in swsMeshQuality_e Return Value 0 if successful
Return Value 0 if successful
VBA: See CWSolidBody::SetMeshQuality .
Examples: Visual Basic for Applications

### ICWSolidBody.SetSolidBodyMaterial
Sets the material to the solid body for analysis.
```vb
Function SetSolidBodyMaterial( ByVal RetVal As CWMaterial ) As System.Integer
```
Parameters RetVal Material Return Value Error or warning as defined in swsMaterialErrorWarning_e
Return Value Error or warning as defined in swsMaterialErrorWarning_e
VBA: See CWSolidBody::SetSolidBodyMaterial .
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Materials

### ICWSolidBody.SolidBodyName
Gets the name of the solid body.
```vb
ReadOnly Property SolidBodyName As System.String
```
Property Value Name of solid body
VBA: See CWSolidBody::SolidBodyName .
Examples: Visual Basic for Applications; Change Beam to Solid Body and Back

## ICWSolidComponent

### ICWSolidComponent.ComponentName
Gets the name of the solid component.
```vb
ReadOnly Property ComponentName As System.String
```
Property Value Name of solid component
VBA: See CWSolidComponent::ComponentName .
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Materials

### ICWSolidComponent.GetSolidBodyAt
Gets the solid body at the specified index.
```vb
Function GetSolidBodyAt( ByVal NIndex As System.Integer, ByRef ErrorCode As System.Integer ) As CWSolidBody
```
Parameters NIndex 0-based index of solid body to get ErrorCode 0 if successful, 1 if no solid body at NIndex Return Value Solid body or NULL if no solid body at NIndex
Return Value Solid body or NULL if no solid body at NIndex
VBA: See CWSolidComponent::GetSolidBodyAt .
Remarks: Before calling this method, call ICWSolidComponent::SolidBodyCount to get NIndex. Index starts at 0.
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Materials

### ICWSolidComponent.SolidBodyCount
Gets the number of solid bodies in the solid component.
```vb
ReadOnly Property SolidBodyCount As System.Integer
```
Property Value Number of solid bodies in the solid component
VBA: See CWSolidComponent::SolidBodyCount .
Examples: Visual Basic for Applications; Create Frequency Study with Solid Mesh

## ICWMaterial

### ICWMaterial.Category
Gets or sets the category of the material.
```vb
Property Category As System.String
```
Property Value Category
VBA: See CWMaterial::Category .
Examples: Visual Basic for Applications

### ICWMaterial.Count
Gets the number of properties defined for this material.
```vb
ReadOnly Property Count As System.Integer
```
Property Value Number of material properties
VBA: See CWMaterial::Count .
Examples: Visual Basic for Applications

### ICWMaterial.Description
Gets or sets the description for the material.
```vb
Property Description As System.String
```
Property Value Description
VBA: See CWMaterial::Description .
Examples: Visual Basic for Applications

### ICWMaterial.GetAustenticSteelCurve
Gets the austentic steel curve data.
```vb
Function GetAustenticSteelCurve() As System.Object
```
Return Value Array of austentic steel curve data (see Remarks )
VBA: See CWMaterial::GetAustenticSteelCurve .
Remarks: Austentic steel curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs x1 = number of cycles for point 1 y1 = alternating stress for point 1 ... xn = number of cycles for point n yn = alternating stress for point n
Examples: Visual Basic for Applications

### ICWMaterial.GetCarbonSteelCurve
Gets the carbon steel curve data.
```vb
Function GetCarbonSteelCurve() As System.Object
```
Return Value Array of carbon steel curve data (see Remarks )
VBA: See CWMaterial::GetCarbonSteelCurve .
Remarks: Array of carbon steel curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs x1 = number of cycles for point 1 y1 = alternating stress for point 1 ... xn = number of cycles for point n yn = alternating stress for point n
Examples: Visual Basic for Applications

### ICWMaterial.GetFatigueSNCurve
Gets the fatigue S-N curve data for user-defined curve sources.
```vb
Function GetFatigueSNCurve( ByVal NIndex As System.Integer, ByRef DStressRatio As System.Double ) As System.Object
```
Parameters NIndex 0-based S-N curve data index DStressRatio Stress ratio Return Value Array of S-N curve data (see Remarks )
Return Value Array of S-N curve data (see Remarks )
VBA: See CWMaterial::GetFatigueSNCurve .
Remarks: This method is valid only if ICWMaterial::SNCurveSource is set to swsMaterialSNCurveSource_e .swsMaterialSNCurveSourceUserDefined. Array of S-N curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs x1 = number of cycles for point 1 y1 = alternating stress for point 1 ... xn = number of cycles for point n yn = alternating stress for point n See the SOLIDWORKS Help topic, Material Dialog Box - Fatigue SN Curves Tab , for more information about S-N curve equations.
Examples: Visual Basic for Applications

### ICWMaterial.GetMaterialDataCurve
Obsolete. Superseded by ICWMaterial::GetMaterialDataCurve3 .
```vb
Function GetMaterialDataCurve( ByVal NIndex As System.Integer, ByRef BUseCurve As System.Integer ) As System.Object
```
Parameters NIndex 0-based index of material data curve BUseCurve 1 = Use material data curve 0 = Do not use material data curve Return Value Array of material curve data (see Remarks )
Return Value Array of material curve data (see Remarks )
VBA: See CWMaterial::GetMaterialDataCurve .
Remarks: Array of material curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs x1 = X property value for point 1 y1 = Y property value for point 1 ... xn = X property value for point n yn = Y property value for point n
Examples: Visual Basic for Applications

### ICWMaterial.GetMaterialDataCurve3
Gets the material data curve.
```vb
Function GetMaterialDataCurve3( ByVal NIndex As System.Integer, ByRef BUseCurve As System.Boolean ) As System.Object
```
Parameters NIndex 0-based index of material data curve BUseCurve -1 or true = Uses material data curve 0 or false = Does not use material data curve (see Remarks ) Return Value Array of material curve data (see Remarks )
Return Value Array of material curve data (see Remarks )
VBA: See CWMaterial::GetMaterialDataCurve3 .
Remarks: This method returns a boolean or integer in out parameter BUseCurve, depending on its prior declaration. If out parameter BUseCurve is cast as a: Boolean, true or false is returned. Long or integer, -1 (=true) or 0 (=false) is returned. Array of material curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs x1 = X property value for point 1 y1 = Y property value for point 1 ... xn = X property value for point n yn = Y property value for point n
Examples: Visual Basic for Applications

### ICWMaterial.GetPropertyByName
Obsolete. Superseded by er ICWMaterial::GetPropertyByName2 .
```vb
Function GetPropertyByName( ByVal NUnit As System.Integer, ByVal SName As System.String, ByRef BTempDependent As System.Integer ) As System.Double
```
Parameters NUnit Unit system as defined in swsUnitSystem_e SName Property name; for example, 'EX' for Elastic modulus, 'NUXY' for Poisson's ratio, etc. BTempDependent True if temperature dependent, false if not Return Value Value of material property
Return Value Value of material property
VBA: See CWMaterial::GetPropertyByName .
Examples: Visual Basic for Applications; Apply Material to Bodies

### ICWMaterial.GetPropertyByName2
Gets the value of the material property by the property name.
```vb
Function GetPropertyByName2( ByVal NUnit As System.Integer, ByVal SName As System.String, ByRef BTempDependent As System.Boolean ) As System.Double
```
Parameters NUnit Unit system as defined in swsUnitSystem_e SName Property name; for example, 'EX' for Elastic modulus, 'NUXY' for Poisson's ratio, etc. BTempDependent -1 or true if temperature dependent, 0 or false if not (see Remarks ) Return Value Value of material property
Return Value Value of material property
VBA: See CWMaterial::GetPropertyByName2 .
Remarks: This method returns a boolean or integer in out parameter BTempDependent, depending on its prior declaration. If out parameter BTempDependent is cast as a: Boolean, true or false is returned. Long or integer, -1 (=true) or 0 (=false) is returned.
Examples: Visual Basic for Applications

### ICWMaterial.GetReferencePlaneName
Gets the name of the reference plane or reference axis used to specify material properties for orthotropic materials.
```vb
Function GetReferencePlaneName() As System.String
```
Return Value Name of reference plane or reference geometry
VBA: See CWMaterial::GetReferencePlaneName .
Examples: Visual Basic for Applications

### ICWMaterial.GetStressStrainCurve
Gets the stress strain curve data.
```vb
Function GetStressStrainCurve() As System.Object
```
Return Value Array of stress curve data
VBA: See CWMaterial::GetStressStrainCurve .
Examples: Visual Basic for Applications

### ICWMaterial.GetTemperatureCurveForProperty
Gets the temperature curve data for the material property.
```vb
Function GetTemperatureCurveForProperty( ByVal SPropName As System.String ) As System.Object
```
Parameters SPropName Property name Return Value Array of temperature data (see Remarks )
Return Value Array of temperature data (see Remarks )
VBA: See CWMaterial::GetTemperatureCurveForProperty .
Remarks: Array of temperature data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs x1 = }}--> }}-->temperature value for point 1 y1 = }}--> }}-->property value for point 1 ... xn = temperature value for point n yn = property value for point n
Examples: Visual Basic for Applications

### ICWMaterial.IncludeCreep
Obsolete. Superseded by ICWMaterial::IncludeCreep2 .
```vb
Property IncludeCreep As System.Integer
```
Property Value 1 = Include creep effect 0 = Do not include creep effect }}--> }}-->
VBA: See CWMaterial::IncludeCreep .
Remarks: Creep effect is valid for nonlinear studies only.
Examples: Visual Basic for Applications

### ICWMaterial.IncludeCreep2
Gets or sets whether to include creep effect for the material model in nonlinear studies only.
```vb
Property IncludeCreep2 As System.Boolean
```
Property Value -1 or true = Include creep effect 0 or false = Do not include creep effect }}-->
Remarks: Creep effect is valid for nonlinear studies only. This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWMaterial.MaterialName
Gets or sets the name of the material name.
```vb
Property MaterialName As System.String
```
Property Value Material name
VBA: See CWMaterial::MaterialName .
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Matierals

### ICWMaterial.MaterialUnits
Gets or sets the units for the material.
```vb
Property MaterialUnits As System.Integer
```
Property Value Unit system as defined in swsUnitSystem_e
VBA: See CWMaterial::MaterialUnits .
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Materials

### ICWMaterial.ModelType
Gets or sets the material model.
```vb
Property ModelType As System.Integer
```
Property Value Type of material model as defined in swsMaterialModelType_e
VBA: See CWMaterial::ModelType .
Examples: Visual Basic for Applications

### ICWMaterial.MooneyRivlinConstants
model used in nonlinear studies only.
```vb
Property MooneyRivlinConstants As System.Integer
```
Property Value Number of Mooney Rivlin constants
VBA: See CWMaterial::MooneyRivlinConstants .
Remarks: Mooney Rivlin constants are valid in nonlinear studies only.
Examples: Visual Basic for Applications

### ICWMaterial.NoOfBulkModuli
nonlinear studies only.
```vb
Property NoOfBulkModuli As System.Integer
```
Property Value Number of bulk moduli
VBA: See CWMaterial::NoOfBulkModuli .
Remarks: Number of bulk moduli for the viscelastic material model are valid in nonlinear studies only.
Examples: Visual Basic for Applications

### ICWMaterial.NoOfShearModuli
Gets or sets the number of shear moduli for the viscoelastic material model used in nonlinear studies only.
```vb
Property NoOfShearModuli As System.Integer
```
Property Value Number of shear moduli
VBA: See CWMaterial::NoOfShearModuli .
Remarks: Number of shear moduli for the viscelastic material model are valid in nonlinear studies only.
Examples: Visual Basic for Applications

### ICWMaterial.OgdenConstants
nonlinear studies only.
```vb
Property OgdenConstants As System.Integer
```
Property Value Number of Ogden constants
VBA: See CWMaterial::OgdenConstants .
Remarks: Ogden constants are valid in nonlinear studies only.
Examples: Visual Basic for Applications

### ICWMaterial.SNCurveEstimateConstants
Gets or sets whether to specify the Basquin Equation constants, B and slope (m), or let the program calculate them from the S-N curve.
```vb
Property SNCurveEstimateConstants As System.Integer
```
Property Value 0 = Specify Basquin Equation constants, B and m 1 = Let the program calculate the Basquin Equation constants from the S-N curve (see Remarks )
VBA: See CWMaterial::SNCurveEstimateConstants .
Remarks: This property is valid only: for fatigue studies that are based on results from linear dynamic random vibration studies. if ICWMaterial::SNCurveSource is set to swsMaterialSNCurveSource_e .swsMaterialSNCurveSourceEquation. If this property is set to... Then use... 0 ICWMaterial::SNCurveSlopeM , ICWMaterial::SNCurveSpecificConstantB , and ICWMaterial::SNCurveSpecificConstantBUnit to specify the Basquin Equation constants. 1 ICWMaterial::SNCurveEstimateCutoff to define the cut-off point used by the program to fit a straight line to the S-N curve and calculate the Basquin Equation constants. See the SOLIDWORKS Help topic, Material Dialog Box - Fatigue SN Curves Tab , for more information about Basquin Equations.
Examples: Visual Basic for Applications; Example Create Fatigue Study for Dynamic Random Vibration Study

### ICWMaterial.SNCurveEstimateCutoff
Gets or sets the cut-off point for estimating Basquin Eqation constants, B and slope (m).
```vb
Property SNCurveEstimateCutoff As System.Integer
```
Property Value 0 or 1: no cut-off; all of the points of the S-N curve are used for fitting a straight line and calculating the Basquin Equation constants, B and m n >= 2 : cut-off point on material S-N curve; the first n points of the S-N curve are used by the program to fit a straight line and calculate the Basquin Equation constants, B and m
VBA: See CWMaterial::SNCurveEstimateCutoff .
Remarks: This property is valid only: for fatigue studies that are based on results from linear dynamic random vibration studies. if ICWMaterial::SNCurveSource is set to swsMaterialSNCurveSource_e .swsMaterialSNCurveSourceEquation, and ICWMaterial::SNCurveEstimateConstants is set to 1. See the SOLIDWORKS Help topic, Material Dialog Box - Fatigue SN Curves Tab , for more information about Basquin Equations.
Examples: Visual Basic for Applications

### ICWMaterial.SNCurveSlopeM
Gets or sets the slope (m) of the S-N curve.
```vb
Property SNCurveSlopeM As System.Double
```
Property Value 0.0 < slope (m) of the S-N curve
VBA: See CWMaterial::SNCurveSlopeM .
Remarks: This property is valid only: for fatigue studies that are based on results from linear dynamic random vibration studies. if ICWMaterial::SNCurveSource is set to swsMaterialSNCurveSource_e .swsMaterialSNCurveSourceEquation, and ICWMaterial::SNCurveEstimateConstants is set to 0. See the SOLIDWORKS Help topic, Material Dialog Box - Fatigue SN Curves Tab , for more information about Basquin Equations.
Examples: Visual Basic for Applications; Example Create Fatigue Study for Dynamic Random Vibration Study

### ICWMaterial.SNCurveSource
Gets or sets the source for the material S-N curve used in fatigue studies.
```vb
Property SNCurveSource As System.Integer
```
Property Value Source as defined in swsMaterialSNCurveSource_e
VBA: See CWMaterial::SNCurveSource .
Remarks: See the SOLIDWORKS Help topic, Material Dialog Box - Fatigue SN Curves Tab , for more information about S-N curve equations.
Examples: Visual Basic for Applications; Example Create Fatigue Study for Dynamic Random Vibration Study

### ICWMaterial.SNCurveSpecificConstantB
Gets or sets the B constant of the Basquin Equation.
```vb
Property SNCurveSpecificConstantB As System.Double
```
Property Value B constant of Basquin Equation (see Remarks )
VBA: See CWMaterial::SNCurveSpecificConstantB .
Remarks: The B constant of the Basquin Equation is the positive real value of the stress range for one cycle. This property is valid only: for fatigue studies that are based on results from linear dynamic random vibration studies. if ICWMaterial::SNCurveSource is set to swsMaterialSNCurveSource_e .swsMaterialSNCurveSourceEquation, and ICWMaterial::SNCurveEstimateConstants is set to 0. See the SOLIDWORKS Help topic, Material Dialog Box - Fatigue SN Curves Tab , for more information about Basquin Equations.
Examples: Visual Basic for Applications; Example Create Fatigue Study for Dynamic Random Vibration Study

### ICWMaterial.SNCurveSpecificConstantBUnit
Gets or sets the units of the Basquin Equation constants.
```vb
Property SNCurveSpecificConstantBUnit As System.Integer
```
Property Value Units of the Basquin Equation constants as defined in swsPressureUnit_e
VBA: See CWMaterial::SNCurveSpecificConstantBUnit .
Remarks: This property is valid only: for fatigue studies that are based on results from linear dynamic random vibration studies. if ICWMaterial::SNCurveSource is set to swsMaterialSNCurveSource_e .swsMaterialSNCurveSourceEquation, and ICWMaterial::SNCurveEstimateConstants is set to 0. See the SOLIDWORKS Help topic, Material Dialog Box - Fatigue SN Curves Tab , for more information about Basquin Equations.
Examples: Visual Basic for Applications; Example Create Fatigue Study for Dynamic Random Vibration Study

### ICWMaterial.SetFatigueSNCurve
Sets the fatigue S-N curve data for user-defined curve sources.
```vb
Sub SetFatigueSNCurve( ByVal NIndex As System.Integer, ByVal DStressRatio As System.Double, ByVal VarFatigueSNCurveData As System.Object, ByRef ErrorCode As System.Integer )
```
Parameters NIndex 0-based S-N curve data index DStressRatio Stress ratio VarFatigueSNCurveData Array of S-N curve data (see Remarks ) ErrorCode Error as defined in swsMaterialFatigueSNCurveError_e
VBA: See CWMaterial::SetFatigueSNCurve .
Remarks: This method is valid only if ICWMaterial::SNCurveSource is set to swsMaterialSNCurveSource_e .swsMaterialSNCurveSourceUserDefined. VarFatigueSNCurveData: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs x1 = number of cycles for point 1 y1 = alternating stress for point 1 ... xn = number of cycles for point n yn = alternating stress for point n See the SOLIDWORKS Help topic, Material Dialog Box - Fatigue SN Curves Tab , for more information about S-N curve equations.
Examples: Visual Basic for Applications

### ICWMaterial.SetMaterialDataCurve
Obsolete. Superseded by ICWMaterial::SetMaterialDataCurve3 .
```vb
Sub SetMaterialDataCurve( ByVal NIndex As System.Integer, ByVal BUseCurve As System.Integer, ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer )
```
Parameters NIndex 0-based index of material data curve BUseCurve 1 = Use material data curve 0 = Do not use material data curve VarCurveData Array of material curve data (see Remarks ) ErrorCode Error as defined in swsMaterialDataCurveError_e
VBA: See CWMaterial::SetMaterialDataCurve .
Remarks: Array of material curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs x1 = X property value for point 1 y1 = Y property value for point 1 ... xn = X property value for point n yn = Y property value for point n
Examples: Visual Basic for Applications

### ICWMaterial.SetMaterialDataCurve3
Sets the material data curve.
```vb
Sub SetMaterialDataCurve3( ByVal NIndex As System.Integer, ByVal BUseCurve As System.Boolean, ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer )
```
Parameters NIndex 0-based index of material data curve BUseCurve -1 or true = Use material data curve 0 or false = Do not use material data curve VarCurveData Array of material curve data (see Remarks ) ErrorCode Error as defined in swsMaterialDataCurveError_e
Remarks: Array of material curve data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs x1 = X property value for point 1 y1 = Y property value for point 1 ... xn = X property value for point n yn = Y property value for point n

### ICWMaterial.SetPropertyByName
Obsolete. Superseded by ICWMaterial::SetPropertyByName2 .
```vb
Sub SetPropertyByName( ByVal SName As System.String, ByVal DValue As System.Double, ByVal BValue As System.Integer )
```
Parameters SName Property name; for example, 'EX' for Elastic modulus, 'NUXY' for Poisson's ratio, etc. DValue Value of material property BValue True if temperature dependent, false if not
VBA: See CWMaterial::SetPropertyByName .
Examples: Visual Basic for Applications; Create Nonlinear Study and Apply Materials

### ICWMaterial.SetPropertyByName2
Sets the value of the specified material property.
```vb
Sub SetPropertyByName2( ByVal SName As System.String, ByVal DValue As System.Double, ByVal BValue As System.Boolean )
```
Parameters SName Property name; for example, 'EX' for Elastic modulus, 'NUXY' for Poisson's ratio, etc. DValue Value of material property BValue -1 or true if temperature dependent, 0 or false if not (see Remarks )
VBA: See CWMaterial::SetPropertyByName2 .
Remarks: Specify BValue as a boolean or integer: true = -1, false = 0.
Examples: Visual Basic for Applications

### ICWMaterial.SetReferencePlane
Sets the name of the reference plane or reference axis used to specify material properties for orthotropic materials.
```vb
Sub SetReferencePlane( ByVal PlaneDisp As System.Object, ByRef ErrorCode As System.Integer )
```
Parameters PlaneDisp Reference entity ErrorCode Error as defined in swsMaterialReferencePlaneError_e
VBA: See CWMaterial::SetReferencePlane .
Examples: Visual Basic for Applications

### ICWMaterial.SetStressStrainCurve
Sets the stress-strain curve data.
```vb
Sub SetStressStrainCurve( ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer )
```
Parameters VarCurveData Array of stress-strain curve data ErrorCode Error as defined in swsMaterialStressStrainCurveError_e
VBA: See CWMaterial::SetStressStrainCurve .
Examples: Visual Basic for Applications

### ICWMaterial.SetTemperatureCurveForProperty
Sets the temperature curve data for the material property.
```vb
Sub SetTemperatureCurveForProperty( ByVal SPropName As System.String, ByVal VarCurveData As System.Object, ByRef ErrorCode As System.Integer )
```
Parameters SPropName Property name VarCurveData Array of temperature data (see Remarks ) ErrorCode Error as defined in swsMaterialTemperatureCurveForPropertyError_e
VBA: See CWMaterial::SetTemperatureCurveForProperty .
Remarks: Array of temperature data: [ n, x1, y1, x2, y2, x3, y3,...xn, yn ] where: n = number of xi,yi pairs x1 = }}--> }}-->temperature value for point 1 y1 = }}--> }}-->property value for point 1 ... xn = temperature value for point n yn = property value for point n
Examples: Visual Basic for Applications

### ICWMaterial.Source
Returns the source of the material.
```vb
ReadOnly Property Source As System.Integer
```
Property Value Source of material as defined in swsMaterialSource_e
VBA: See CWMaterial::Source .
Examples: Visual Basic for Applications; Analyze Part

## ICWMesh

### ICWMesh.ApplyMeshControl
Creates a mesh control.
```vb
Function ApplyMeshControl( ByVal EntitiesArray As System.Object, ByRef ErrorCode As System.Integer ) As CWMeshControl
```
Parameters EntitiesArray Array of entities for mesh control ErrorCode Error as defined in swsMeshControlError_e Return Value Mesh control
Return Value Mesh control
VBA: See CWMesh::ApplyMeshControl .
Examples: Visual Basic for Applications; Apply Mesh Control

### ICWMesh.AutomaticLooping
Obsolete. Superseded by ICWMesh::AutomaticLooping2 .
```vb
Property AutomaticLooping As System.Integer
```
Property Value 1 if automatic looping is on, 0 if not
VBA: See CWMesh::AutomaticLooping .
Examples: Visual Basic for Applications

### ICWMesh.AutomaticLooping2
Gets or sets whether to automatically retry to mesh the model using a different global element size for solids.
```vb
Property AutomaticLooping2 As System.Boolean
```
Property Value -1 or true if automatic looping is on, 0 or false if not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWMesh.AutomaticTransition
Obsolete. Superseded by ICWMesh::AutomaticTransition2 .
```vb
Property AutomaticTransition As System.Integer
```
Property Value 1 if automatic transition is on, 0 if not
VBA: See CWMesh::AutomaticTransition .
Examples: Visual Basic for Applications

### ICWMesh.AutomaticTransition2
Gets or sets whether mesh controls are automatically applied to small features, holes, fillets, and other fine details in the model.
```vb
Property AutomaticTransition2 As System.Boolean
```
Property Value -1 or true if automatic transition is on, 0 or false if not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWMesh.DeleteMeshControl
Deletes a mesh control.
```vb
Sub DeleteMeshControl( ByVal SName As System.String )
```
Parameters SName Name of the mesh control to delete
VBA: See CWMesh::DeleteMeshControl .
Examples: Visual Basic for Applications

### ICWMesh.ElementCount
Gets the number of elements in the mesh.
```vb
ReadOnly Property ElementCount As System.Integer
```
Property Value Number of elements in the mesh
VBA: See CWMesh::ElementCount .
Examples: Visual Basic for Applications

### ICWMesh.ElementSize
Gets or sets the global element size.
```vb
Property ElementSize As System.Double
```
Property Value Global element size
VBA: See CWMesh::ElementSize .
Examples: Visual Basic for Applications

### ICWMesh.ElementSizeFactorForEachLoop
Obsolete. Not superseded.
```vb
Property ElementSizeFactorForEachLoop As System.Double
```
Property Value Element size factor for looping
VBA: See CWMesh::ElementSizeFactorForEachLoop .
Examples: Visual Basic for Applications

### ICWMesh.FlipShellElements
bottom face and vice versa).
```vb
Function FlipShellElements( ByVal DispFaceArray As System.Object ) As System.Integer
```
Parameters DispFaceArray Array of shells Return Value Error as defined in swsMeshFlipShellError_e
Return Value Error as defined in swsMeshFlipShellError_e
VBA: See CWMesh::FlipShellElements .
Remarks: Align shells correctly for stress results.
Examples: Visual Basic for Applications

### ICWMesh.GetConnectivity
Gets the connectivity of nodes on the surface of the model.
```vb
Function GetConnectivity( ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters ErrorCode Error code as defined in swsMeshQueryErrorCode_e Return Value Array of node connectivity data (see Remarks )
Return Value Array of node connectivity data (see Remarks )
VBA: See CWMesh::GetConnectivity .
Remarks: Array of node connectivity data: [ c 1 , n 11 , n 12 , n 13 , c 2 , n 21 , n 22 , n 23 , ..., c i , n i 1 , n i 2 , n i 3 ] where: c i is number of nodes in the i th connectivity unit; 2 for beams, 3 for solids and shells, 4 for hexagons n i 1 , n i 2 , n i 3 are the nodes in the i th connectivity unit
Examples: Visual Basic for Applications; Example Create Curvature-based Mesh

### ICWMesh.GetDefaultElementSizeAndTolerance
Gets the default element size and tolerance.
```vb
Sub GetDefaultElementSizeAndTolerance( ByVal NUnits As System.Integer, ByRef DElementSize As System.Double, ByRef DTolerance As System.Double )
```
Parameters NUnits Mesh linear units as defined swsLinearUnit_e DElementSize Average global element size DTolerance Tolerance (if the distance between any two nodes is less than this distance, then the nodes are merged)
VBA: See CWMesh::GetDefaultElementSizeAndTolerance .
Examples: Visual Basic for Applications; Analyze Part

### ICWMesh.GetDefaultMaxAndMinElementSize
Gets the default maximum and minimum element sizes used for boundaries in a curvature-based mesh.
```vb
Sub GetDefaultMaxAndMinElementSize( ByVal NUnits As System.Integer, ByRef DMaxElementSize As System.Double, ByRef DMinElementSize As System.Double )
```
Parameters NUnits Mesh linear units as defined swsLinearUnit_e DMaxElementSize Maximum element size used for boundaries with lowest curvature DMinElementSize Minimum element size used for boundaries with highest curvature
VBA: See CWMesh::GetDefaultMaxAndMinElementSize .
Examples: Visual Basic for Applications; Create Curvature-based Mesh

### ICWMesh.GetElementDataFromEntity
Gets the element associated with an entity.
```vb
Function GetElementDataFromEntity( ByVal DispEntity As System.Object, ByRef NCount As System.Integer ) As System.Object
```
Parameters DispEntity Entity NCount Number of elements Return Value Array of element data (see Remarks )
Return Value Array of element data (see Remarks )
VBA: See CWMesh::GetElementDataFromEntity .
Remarks: Array of element data: [ E 1 , N 1 E 1 , N 2 E 1 , N 3 E 1 , N 4 E 1 , N 5 E 1 , N 6 E 1 , N 7 E 1 , N 8 E 1 , N 9 E 1 , N 10 E 1 , X 1 , Y 1 , Z 1 , AR 1 , JR 1 , E 2 , N 1 E 2 , N 2 E 2 , N 3 E 2 , N 4 E 2 , N 5 E 2 , N 6 E 2 , N 7 E 2 , N 8 E 2 , N 9 E 2 , N 10 E 2 , X 2 , Y 2 , Z 2 , AR 2 , JR 2 , ..., E i , N 1 E i , N2E i , N3E i , N4E i , N5E i , N6E i , N7E i , N8E i , N9E i , N10E i , X i , Y i , Z i , AR i , JR i ] where: E i = i th element number N 1 E i to N 10 E i = Node numbers associated with element E i ; returns 10 node numbers for higher order solid elements; returns -1 for shells and lower order meshes. X i , Y i , Z i = x, y and z coordinates of center of E i AR i }}--> }}-->= Aspect ratio for element E i JR i }}--> }}-->= Jacobian ratio for element E i ; returns -1 if not applicable
Examples: Visual Basic for Applications

### ICWMesh.GetElementList
Gets the elements of the specified type.
```vb
Function GetElementList( ByVal NElementType As System.Integer, ByRef VarElementIDs As System.Object ) As System.Integer
```
Parameters NElementType Type of elements in VarElementIDs as defined in swsSimulationElementTypes_e VarElementIDs Array of IDs of elements of type NElementType Return Value Number of elements in VarElementIDs
Return Value Number of elements in VarElementIDs
VBA: See CWMesh::GetElementList .
Remarks: You must set ICWMesh::Quality to swsMeshQuality_e .swsMeshQualityHigh before calling this method.
Examples: Visual Basic for Applications; Example Get Mesh Elements and Nodes

### ICWMesh.GetElementLocation
Gets the coordinates of an element's center.
```vb
Function GetElementLocation( ByVal NNodeNo As System.Integer, ByRef XVal As System.Double, ByRef YVal As System.Double, ByRef ZVal As System.Double ) As System.Integer
```
Parameters NNodeNo Element number XVal X coordinate of the element's center YVal Y coordinate of the element's center ZVal Z coordinate of the element's center Return Value Error as defined in swsMeshElementNodeLocation_e
Return Value Error as defined in swsMeshElementNodeLocation_e
VBA: See CWMesh::GetElementLocation .
Examples: Visual Basic for Applications

### ICWMesh.GetElements
Gets the elements of the mesh.
```vb
Function GetElements() As System.Object
```
Return Value Array of the elements of the mesh (see Remarks )
VBA: See CWMesh::GetElements .
Remarks: Probe functionality. Array of element data: [ E 1 , N 1 E 1 , N 2 E 1 , N 3 E 1 , N 4 E 1 , N 5 E 1 , N 6 E 1 , N 7 E 1 , N 8 E 1 , N 9 E 1 , N 10 E 1 , X 1 , Y 1 , Z 1 , AR 1 , JR 1 , E 2 , N 1 E 2 , N 2 E 2 , N 3 E 2 , N 4 E 2 , N 5 E 2 , N 6 E 2 , N 7 E 2 , N 8 E 2 , N 9 E 2 , N 10 E 2 , X 2 , Y 2 , Z 2 , AR 2 , JR 2 , ..., E i , N 1 E i , N2E i , N3E i , N4E i , N5E i , N6E i , N7E i , N8E i , N9E i , N10E i , X i , Y i , Z i , AR i , JR i ] where: E i = i th element number N 1 E i to N 10 E i = Node numbers associated with element E i ; returns 10 node numbers for higher order solid elements; returns -1 for shells and lower order meshes. X i , Y i , Z i = x, y and z coordinates of center of E i AR i }}--> }}-->= Aspect ratio for element E i JR i }}--> }}-->= Jacobian ratio for element E i ; returns -1 if not applicable
Examples: Visual Basic for Applications

### ICWMesh.GetFailedComponents
Gets the components that failed to mesh.
```vb
Function GetFailedComponents( ByRef NComp As System.Integer ) As System.Object
```
Parameters NComp Number of components that failed to mesh Return Value Array of bodies that failed to mesh
Return Value Array of bodies that failed to mesh
VBA: See CWMesh::GetFailedComponents .
Examples: Visual Basic for Applications

### ICWMesh.GetFailedEdges
Gets the edges that failed to mesh.
```vb
Function GetFailedEdges( ByRef NFailedEdges As System.Integer ) As System.Object
```
Parameters NFailedEdges Number of edges that failed to mesh Return Value Array of edges that failed to mesh
Return Value Array of edges that failed to mesh
VBA: See CWMesh::GetFailedEdges .
Examples: Visual Basic for Applications

### ICWMesh.GetFailedFaces
Gets the faces that failed to mesh.
```vb
Function GetFailedFaces( ByRef NFailedFaces As System.Integer ) As System.Object
```
Parameters NFailedFaces Number of faces that failed to mesh Return Value Array of faces that failed to mesh
Return Value Array of faces that failed to mesh
VBA: See CWMesh::GetFailedFaces .
Examples: Visual Basic for Applications

### ICWMesh.GetMeshControlAt
Gets mesh control at the specified index.
```vb
Function GetMeshControlAt( ByVal NIndex As System.Integer ) As CWMeshControl
```
Parameters NIndex 0-based index of mesh control to get Return Value Mesh control
Return Value Mesh control
VBA: See CWMesh::GetMeshControlAt .
Examples: Visual Basic for Applications

### ICWMesh.GetNoOfFailedComponents
Gets the number of components that failed to mesh.
```vb
Function GetNoOfFailedComponents() As System.Integer
```
Return Value Number of components
VBA: See CWMesh::GetNoOfFailedComponents .
Examples: Visual Basic for Applications; Torques at a Time

### ICWMesh.GetNodeDataFromEntity
Gets the node data associated with an entity.
```vb
Function GetNodeDataFromEntity( ByVal DispEntity As System.Object, ByRef NCount As System.Integer ) As System.Object
```
Parameters DispEntity Entity NCount Number of nodes Return Value Array of nodal data (see Remarks )
Return Value Array of nodal data (see Remarks )
VBA: See CWMesh::GetNodeDataFromEntity .
Remarks: Probe functionality. Array of nodal data: [ N 1 , X 1 , Y 1 , Z 1 , N 2 , X 2 , Y 2 , Z 2 , ..., N i , X i , Y i , Z i ] where: N i = i th node number X i , Y i , Z i = coordinates of node N i
Examples: Visual Basic for Applications

### ICWMesh.GetNodeList
Gets the nodes of elements of the specified type.
```vb
Function GetNodeList( ByVal NElementType As System.Integer, ByRef VarNodeIDs As System.Object ) As System.Integer
```
Parameters NElementType Type of elements as defined in swsSimulationElementTypes_e VarNodeIDs Array of IDs of nodes of elements of type NElementType Return Value Number of nodes in VarNodeIDs
Return Value Number of nodes in VarNodeIDs
VBA: See CWMesh::GetNodeList .
Remarks: You must set ICWMesh::Quality to swsMeshQuality_e .swsMeshQualityHigh before calling this method.
Examples: Visual Basic for Applications; Example Get Mesh Elements and Nodes

### ICWMesh.GetNodeLocation
Gets the coordinates of the specified node.
```vb
Function GetNodeLocation( ByVal NNodeNo As System.Integer, ByRef XVal As System.Double, ByRef YVal As System.Double, ByRef ZVal As System.Double ) As System.Integer
```
Parameters NNodeNo Node number XVal X coordinate YVal Y coordinate ZVal Z coordinate Return Value Error as defined in swsMeshControlError_e
Return Value Error as defined in swsMeshControlError_e
VBA: See CWMesh::GetNodeLocation .
Examples: Visual Basic for Applications

### ICWMesh.GetNodeOutwardDirections
Gets the normal directions at all of the nodes on the surface of this mesh.
```vb
Function GetNodeOutwardDirections( ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters ErrorCode Error code as defined in swsMeshQueryErrorCode_e Return Value Array of x, y, and z coordinates of normal directions at all of the nodes on the surface of this mesh (see Remarks )
Return Value Array of x, y, and z coordinates of normal directions at all of the nodes on the surface of this mesh (see Remarks )
VBA: See CWMesh::GetNodeOutwardDirections .
Remarks: Array: [ x 1 , y 1 , z 1 , x 2 , y 2 , z 2 , ..., x i , y i , z i ] where: x i , y i , z i is the normal direction at the i th node
Examples: Visual Basic for Applications; Example Create Curvature-based Mesh

### ICWMesh.GetNodes
Gets all of the nodes in the mesh.
```vb
Function GetNodes() As System.Object
```
Return Value Array of all of the nodes in the mesh
VBA: See CWMesh::GetNodes .
Examples: Visual Basic for Applications

### ICWMesh.GetPoorMeshQualityElementList
Obsolete. Superseded by ICWMesh::GetPoorMeshQualityElementList2 .
```vb
Function GetPoorMeshQualityElementList( ByVal NMeshQualityKPI As System.Integer, ByVal DLimitVal As System.Double, ByVal NJacobianPts As System.Integer, ByVal BCreateplot As System.Integer, ByRef VarElementIDs As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters NMeshQualityKPI Mesh element quality criterion as defined in swsMeshElementQualityKPI_e (see Remarks ) DLimitVal Upper or lower limit of NMeshQualityKPI, beyond which the mesh becomes poor quality (see Remarks ) NJacobianPts Jacobian points criterion (see Remarks ) BCreateplot 1 to create a plot, 0 to not VarElementIDs Array of element IDs ErrorCode Error code as defined by swsMeshKPIErrCode_e Return Value Number of elements in VarElementIDs
Return Value Number of elements in VarElementIDs
VBA: See CWMesh::GetPoorMeshQualityElementList .
Remarks: This method: Corresponds to the Mesh Quality Diagnostics tool in SOLIDWORKS Simulation. Access the Mesh Quality Diagnostics PropertyManager in a SOLIDWORKS Simulation study by right-clicking Mesh in the study's FeatureManager design tree and selecting Mesh Quality Diagnostics . Retrieves the IDs of elements satisfying the criterion specified by NMeshQualityKPI, DLimitVal, and NJacobianPts if valid. If NMeshQualityKPI is swsMeshElementQualityKPI_e.swsMeshElementQualityKPI_: Volume or Area, then mesh has poor quality if less than DLimitVal; use NJacobianPts to specify the Jacobian points criterion. AspectRatio or JacobianRatio or ElemSkewRatio, then mesh has poor quality if greater than DLimitVal. NJacobianPts is ignored. JacobianRatio, then mesh has poor quality if greater than DLimitVal; use NJacobianPts to specify the Jacobian points criterion. Possible values for NJacobianPts: 4 = Number of points inside the element tetrahedron near its vertices 16 = Number of points inside the element tetrahedron 29 = Number of points inside the element tetrahedron (same as 16 but calculates a more precise Jacobian ratio) 10 = At nodes; Number of points = 4 vertices + 6 midpoints on 6 edges of the element tetrahedron For more information, read the Simulation > Meshing > Mesh Quality Checks > Defining a Mesh Quality Plot topic in the SOLIDWORKS Help.
Examples: Visual Basic for Applications

### ICWMesh.GetPoorMeshQualityElementList2
Gets the Mesh Quality Diagnostics element list.
```vb
Function GetPoorMeshQualityElementList2( ByVal NMeshQualityKPI As System.Integer, ByVal DLimitVal As System.Double, ByVal NJacobianPts As System.Integer, ByVal BCreateplot As System.Boolean, ByRef VarElementIDs As System.Object, ByRef ErrorCode As System.Integer ) As System.Integer
```
Parameters NMeshQualityKPI Mesh element quality criterion as defined in swsMeshElementQualityKPI_e (see Remarks ) DLimitVal Upper or lower limit of NMeshQualityKPI, beyond which the mesh becomes poor quality (see Remarks ) NJacobianPts Jacobian points criterion (see Remarks ) BCreateplot -1 or true to create a plot, 0 or false to not VarElementIDs Array of element IDs ErrorCode Error code as defined by swsMeshKPIErrCode_e Return Value Number of elements in VarElementIDs
Return Value Number of elements in VarElementIDs
Remarks: This method: Corresponds to the Mesh Quality Diagnostics tool in SOLIDWORKS Simulation. Access the Mesh Quality Diagnostics PropertyManager in a SOLIDWORKS Simulation study by right-clicking Mesh in the study's FeatureManager design tree and selecting Mesh Quality Diagnostics . Retrieves the IDs of elements satisfying the criterion specified by NMeshQualityKPI, DLimitVal, and NJacobianPts if valid. If NMeshQualityKPI is swsMeshElementQualityKPI_e.swsMeshElementQualityKPI_: Volume or Area, then mesh has poor quality if less than DLimitVal; use NJacobianPts to specify the Jacobian points criterion. AspectRatio or JacobianRatio or ElemSkewRatio, then mesh has poor quality if greater than DLimitVal. NJacobianPts is ignored. JacobianRatio, then mesh has poor quality if greater than DLimitVal; use NJacobianPts to specify the Jacobian points criterion. Possible values for NJacobianPts: 4 = Number of points inside the element tetrahedron near its vertices 16 = Number of points inside the element tetrahedron 29 = Number of points inside the element tetrahedron (same as 16 but calculates a more precise Jacobian ratio) 10 = At nodes; Number of points = 4 vertices + 6 midpoints on 6 edges of the element tetrahedron For more information, read the Simulation > Meshing > Mesh Quality Checks > Defining a Mesh Quality Plot topic in the SOLIDWORKS Help.

### ICWMesh.GetShellElementNormalAt
Gets shell normal for a face.
```vb
Function GetShellElementNormalAt( ByVal DispFace As System.Object ) As System.Integer
```
Parameters DispFace Face Return Value Shell normal for a face as defined at swsMeshShellNormal_e
Return Value Shell normal for a face as defined at swsMeshShellNormal_e
VBA: See CWMesh::GetShellElementNormalAt .
Remarks: This method determines the positioning of the top and bottom faces of the shell elements associated with the face. The top face is towards the positive normal direction by default. Use this method to align shells. Flipping switches the faces.
Examples: Visual Basic for Applications

### ICWMesh.GetSolidElementList
Gets the elements at the specified depth of this solid mesh.
```vb
Function GetSolidElementList( ByVal DDepth As System.Double, ByVal NUnit As System.Integer, ByRef VarElementIDs As System.Object ) As System.Integer
```
Parameters DDepth Depth at which to find elements NUnit Linear units of DDepth as defined in swsLinearUnit_e VarElementIDs Array of element IDs Return Value Number of elements in VarElementIDs
Return Value Number of elements in VarElementIDs
VBA: See CWMesh::GetSolidElementList .
Remarks: You must set ICWMesh::Quality to swsMeshQuality_e .swsMeshQualityHigh before calling this method.
Examples: Visual Basic for Applications; Example Get Mesh Elements and Nodes

### ICWMesh.GetSolidNodeList
Gets the nodes at the specified depth of this solid mesh.
```vb
Function GetSolidNodeList( ByVal DDepth As System.Double, ByVal NUnit As System.Integer, ByRef VarNodeIDs As System.Object ) As System.Integer
```
Parameters DDepth Depth at which to find nodes NUnit Linear units of DDepth as defined swsLinearUnit_e VarNodeIDs Array of node IDs Return Value Number of nodes in VarNodeIDs
Return Value Number of nodes in VarNodeIDs
VBA: See CWMesh::GetSolidNodeList .
Remarks: You must set ICWMesh::Quality to swsMeshQuality_e .swsMeshQualityHigh before calling this method.
Examples: Visual Basic for Applications; Example Get Mesh Elements and Nodes

### ICWMesh.GetSurfaceNodesAndNormals
Gets the nodes and normal vectors at the surface of this solid mesh.
```vb
Function GetSurfaceNodesAndNormals( ByRef VarNodeIDs As System.Object, ByRef VarNumNormals As System.Object, ByRef VarNormalVecs As System.Object ) As System.Integer
```
Parameters VarNodeIDs Array of node IDs (see Remarks ) VarNumNormals Array of numbers (1, 2 or 3) of normal vectors for the corresponding node ID in VarNodeIDs (see Remarks ) VarNormalVecs Array of normal vectors (see Remarks ) Return Value Number of nodes
Return Value Number of nodes
VBA: See CWMesh::GetSurfaceNodesAndNormals .
Remarks: You must set ICWMesh::Quality to swsMeshQuality_e .swsMeshQualityHigh before calling this method. Each node can have one or more normal vectors. For example, for a meshed cube: Node location Number of normal vectors Middle of side 1 Edge 2 Vertex 3 VarNormalVecs contains an array of all of the normal vectors: [ xa 1 , ya 1 , za 1 , xb 1 , yb 1 , zb 1 , xc 1 , yc 1 , zc 1 , xa 2 , ya 2 , za 2 , ..., xa i , ya i , za i ] where: x , y , and z denote location a , b , and c denote normal vectors xa i , ya i , za i denote the location of the a th normal vector for the i th node Use VarNumNormals to calculate which vectors in VarNormalVecs map to the node IDs in VarNodeIDs.
Examples: Visual Basic for Applications; Example Get Mesh Elements and Nodes

### ICWMesh.GetWorstJacobianRatio
Gets the worst Jacobian ratio for this mesh.
```vb
Function GetWorstJacobianRatio() As System.Double
```
Return Value Worst Jacobian ratio (see Remarks )
VBA: See CWMesh::GetWorstJacobianRatio .
Remarks: The data returned by this method is valid only if ICWMesh::Quality is set to swsMeshQuality_e .swsMeshQualityHigh. Analysis of curved geometry requires a mesh of parabolic tetrahedral elements. With extremely curved edges, placement of the mid-side nodes of these elements on the geometry can distort the elements and the analysis. The Jacobian ratio calculated at a point inside a parabolic mesh element provides a measure of the degree of distortion at that location. The higher the ratio, the greater the distortion of the element. A Jacobian ratio of 40 or less is acceptable. A negative Jacobian ratio causes SOLIDWORKS analysis to stop. Call this method after ICWStudy::CreateMesh to calculate the worst or highest Jacobian ratio among the elements of this mesh. See the SOLIDWORKS Simulation Help for more information about Jacobian ratios and meshes.
Examples: Visual Basic for Applications; Example Create Linear Dynamic Study

### ICWMesh.GrowthRatio
Gets or sets the global element size growth ratio starting from regions of highest curvatures in all directions in a curvature-based mesh.
```vb
Property GrowthRatio As System.Double
```
Property Value Global element size growth ratio starting from regions of highest curvatures in all directions
VBA: See CWMesh::GrowthRatio .
Examples: Visual Basic for Applications; Create Curvature-based Mesh

### ICWMesh.IsComponentFailed
Obsolete. Superseded by ICWMesh::IsComponentFailed2 .
```vb
Function IsComponentFailed( ByVal SName As System.String ) As System.Integer
```
Parameters SName Name of component Return Value 1 if failed, 0 if not
Return Value 1 if failed, 0 if not
VBA: See CWMesh::IsComponentFailed .
Examples: Visual Basic for Applications

### ICWMesh.IsComponentFailed2
Checks the status of the specified component.
```vb
Function IsComponentFailed2( ByVal SName As System.String ) As System.Boolean
```
Parameters SName Name of component Return Value -1 or true if failed, 0 or false if not
Return Value -1 or true if failed, 0 or false if not

### ICWMesh.IsMeshFailed
Obsolete. Superseded by ICWMesh::IsMeshFailed2 .
```vb
ReadOnly Property IsMeshFailed As System.Integer
```
Property Value 1 if mesh failed, 0 if not
VBA: See CWMesh::IsMeshFailed .
Examples: Visual Basic for Applications

### ICWMesh.IsMeshFailed2
Checks the status of the mesh.
```vb
ReadOnly Property IsMeshFailed2 As System.Boolean
```
Property Value -1 or true if mesh failed, 0 or false if not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWMesh.JacobianPoints
Gets or sets the number of points to be used in calculating a Jacobian ratio.
```vb
Property JacobianPoints As System.Integer
```
Property Value Number of Jacobian points (see Remarks )
VBA: See CWMesh::JacobianPoints .
Remarks: This property: is valid only if ICWMesh::Quality is set to swsMeshQuality_e .swsMeshQualityHigh. has 4 possible values: 4 = Number of points inside the element tetrahedron near its vertices 16 = Number of points inside the element tetrahedron 29 = Number of points inside the element tetrahedron (same as 16 but calculates a more precise Jacobian ratio) 10 = At nodes; Number of points = 4 vertices + 6 midpoints on 6 edges of the element tetrahedron For more information, read: The Advanced section of Simulation > Meshing > Mesh PropertyManager topic - and - The Jacobian Points section of the Simulation > Meshing > Mesh Quality Checks topic in the SOLIDWORKS help.
Examples: Visual Basic for Applications

### ICWMesh.MaxAspectRatio
Gets the maximum aspect ratio for all elements in the mesh.
```vb
ReadOnly Property MaxAspectRatio As System.Double
```
Property Value Maximum aspect ratio
VBA: See CWMesh::MaxAspectRatio .
Examples: Visual Basic for Applications

### ICWMesh.MaxElementSize
Gets the maximum element size used for boundaries with lowest curvature in a curvature-based mesh.
```vb
Property MaxElementSize As System.Double
```
Property Value Maximum element size used for boundaries with lowest curvature
VBA: See CWMesh::MaxElementSize .
Examples: Visual Basic for Applications; Create Curvature-based Mesh

### ICWMesh.MeshControlCount
Gets the number of mesh controls defined in the active study.
```vb
ReadOnly Property MeshControlCount As System.Integer
```
Property Value Number of mesh controls in the active study
VBA: See CWMesh::MeshControlCount .
Examples: Visual Basic for Applications; Get and Set Beams and Joints

### ICWMesh.MeshState
Gets the mesh state.
```vb
ReadOnly Property MeshState As System.Integer
```
Property Value Mesh state as defined in swsMeshState_e
VBA: See CWMesh::MeshState .
Examples: Visual Basic for Applications; Get and Set Beams and Joints

### ICWMesh.MeshType
Returns the mesh type.
```vb
ReadOnly Property MeshType As System.Integer
```
Property Value Mesh type as defined in swsMeshType_e (see Remarks )
VBA: See CWMesh::MeshType .
Remarks: Valid mesh types for different types of studies: Study Type Solid Mesh Shell Mesh Mixed Mesh (Solid & Shell) Beam Mesh Static Supported Supported Supported Supported Frequency Supported Supported Supported Supported Buckling Supported Supported Supported Supported Thermal Supported Supported Supported Supported Nonlinear Supported Supported Supported Not Supported Linear Dynamic Supported Supported Supported Not Supported Drop Test Supported Not Supported Not Supported Not Supported Fatigue Not Applicable Not Applicable Not Applicable Not Applicable Optimization Not Applicable Not Applicable Not Applicable Not Applicable NOTES: For documents with surface geometry only (no solids), only RefSurfShellElementMesh is supported. BeamElementMesh is supported only for static, buckling, and frequency studies. Drop test studies support solid mesh only. Optimization and fatigue studies do not have mesh on their own. They use mesh of referenced studies. Mesh type is set to SolidElementMesh.
Examples: Visual Basic for Applications

### ICWMesh.MesherType
Gets or sets the type of mesh.
```vb
Property MesherType As System.Integer
```
Property Value Type of mesh as defined in swsMesherType_e
VBA: See CWMesh::MesherType .
Examples: Visual Basic for Applications; Get and Set Beams and Joints; Create Curvature-based Mesh

### ICWMesh.MinElementSize
Gets the minimum element size used for boundaries with highest curvature in a curvature-based mesh.
```vb
Property MinElementSize As System.Double
```
Property Value Minimum element size used for boundaries with highest curvature
VBA: See CWMesh::MinElementSize .
Examples: Visual Basic for Applications; Example Create Curvature-based Mesh

### ICWMesh.MinElementsInCircle
angle in a curvature-based mesh.
```vb
Property MinElementsInCircle As System.Integer
```
Property Value Number of elements around a circle to determine the maximum angle
VBA: See CWMesh::MinElementsInCircle .
Examples: Visual Basic for Applications; Create Curvature-based Mesh

### ICWMesh.NegativeJacobianRatioCheck
Obsolete. Superseded by ICWMesh::NegativeJacobianRatioCheck2 .
```vb
Property NegativeJacobianRatioCheck As System.Integer
```
Property Value 1 to check for a negative Jacobian ratio, 0 to not
VBA: See CWMesh::NegativeJacobianRatioCheck .
Remarks: For more information, read the Jacobian Points section of the Simulation > Meshing > Mesh Quality Checks topic in the SOLIDWORKS help.
Examples: Visual Basic for Applications

### ICWMesh.NegativeJacobianRatioCheck2
Gets or sets whether to check for a negative Jacobian ratio while meshing.
```vb
Property NegativeJacobianRatioCheck2 As System.Boolean
```
Property Value -1 or true to check for a negative Jacobian ratio, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer. For more information, read the Jacobian Points section of the Simulation > Meshing > Mesh Quality Checks topic in the SOLIDWORKS help.

### ICWMesh.NodeCount
Gets the number of nodes in the active study.
```vb
ReadOnly Property NodeCount As System.Integer
```
Property Value Number of nodes in the active study
VBA: See CWMesh::NodeCount .
Examples: Visual Basic for Applications; Example Copy Mesh and Generate Report

### ICWMesh.NumberOfLoops
Gets or sets the number of loops for automatic mesh looping.
```vb
Property NumberOfLoops As System.Integer
```
Property Value Number of loops
VBA: See CWMesh::NumberOfLoops .
Examples: Visual Basic for Applications

### ICWMesh.Quality
Gets or sets the mesh quality.
```vb
Property Quality As System.Integer
```
Property Value Mesh quality as defined in swsMeshQuality_e
VBA: See CWMesh::Quality .
Examples: Visual Basic for Applications; Analyze Part

### ICWMesh.SaveSettingsWithoutMeshing
Obsolete. Superseded by ICWMesh::SaveSettingsWithoutMeshing2 .
```vb
Property SaveSettingsWithoutMeshing As System.Integer
```
Property Value 1 to not mesh, but save the current mesh settings; 0 to mesh using the current settings
VBA: See CWMesh::SaveSettingsWithoutMeshing .
Examples: Visual Basic for Applications

### ICWMesh.SaveSettingsWithoutMeshing2
Gets or sets whether to save mesh settings without meshing.
```vb
Property SaveSettingsWithoutMeshing2 As System.Boolean
```
Property Value -1 or true to not mesh, but save the current mesh settings; 0 or false to mesh using the current settings
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWMesh.SetMeshAndParamsForDefaultBCB
Sets the mesher to be blended curvature based.
```vb
Sub SetMeshAndParamsForDefaultBCB()
```
VBA: See CWMesh::SetMeshAndParamsForDefaultBCB .
Remarks: As of Simulation 2022, blended curvature-based is the default mesher.
Examples: Visual Basic for Applications

### ICWMesh.SmoothSurface
Obsolete. Not superseded.
```vb
Property SmoothSurface As System.Integer
```
Property Value 1 if on, 0 if off
VBA: See CWMesh::SmoothSurface .
Examples: Visual Basic for Applications

### ICWMesh.SmoothSurface2
Obsolete. Not superseded.
```vb
Property SmoothSurface2 As System.Boolean
```

### ICWMesh.TimeToCompleteMesh
Gets the estimated time to complete the mesh.
```vb
ReadOnly Property TimeToCompleteMesh As System.Integer
```
Property Value Estimated time to complete the mesh formatted as hh:mm:ss; e.g., return value 12 equals 12 seconds
VBA: See CWMesh::TimeToCompleteMesh .
Examples: Visual Basic for Applications; Get and Set Beams and Joints

### ICWMesh.Tolerance
Gets the tolerance value for mesh loops.
```vb
Property Tolerance As System.Double
```
Property Value Tolerance value for mesh loops
VBA: See CWMesh::Tolerance .
Examples: Visual Basic for Applications

### ICWMesh.ToleranceFactorForEachLoop
Obsolete. Not superseded.
```vb
Property ToleranceFactorForEachLoop As System.Double
```
Property Value Tolerance factor for mesh loops
VBA: See CWMesh::ToleranceFactorForEachLoop .
Remarks: The tolerance factor for each is the factor by which the new global element size if multipled to calculate the new global element size.
Examples: Visual Basic for Applications

### ICWMesh.Unit
Gets the units for the mesh.
```vb
Property Unit As System.Integer
```
Property Value Mesh linear units as defined in swsLinearUnit_e
VBA: See CWMesh::Unit .
Examples: Visual Basic for Applications

### ICWMesh.UseJacobianCheckForShells
Obsolete. Superseded by ICWMesh::UseJacobianCheckForShells2 .
```vb
Property UseJacobianCheckForShells As System.Integer
```
Property Value 1 to perform the Jacobian check, 0 to not
VBA: See CWMesh::UseJacobianCheckForShells .
Examples: Visual Basic for Applications

### ICWMesh.UseJacobianCheckForShells2
Gets or sets whether to perform a Jacobian check for shells.
```vb
Property UseJacobianCheckForShells2 As System.Boolean
```
Property Value -1 or true to perform the Jacobian check, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWMesh.UseJacobianCheckForSolids
Gets or sets whether to perform a Jacobian check for solids.
```vb
Property UseJacobianCheckForSolids As System.Integer
```
Property Value 1 to check the Jacobian for solids, 0 to not
VBA: See CWMesh::UseJacobianCheckForSolids .
Examples: Visual Basic for Applications

## ICWMeshControl

### ICWMeshControl.BeamSelected
Obsolete. Superseded by ICWMeshControl::BeamSelected2 .
```vb
Property BeamSelected As System.Integer
```
Property Value 1 to select beams, 0 to not
VBA: See CWMeshControl::BeamSelected .
Remarks: This property is only valid for models with beams.
Examples: Visual Basic for Applications

### ICWMeshControl.BeamSelected2
Gets and sets whether to select beams for this mesh control.
```vb
Property BeamSelected2 As System.Boolean
```
Property Value -1 or true to select beams, 0 or false to not
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWMeshControl.ElementSize
Gets or sets the element size of the mesh.
```vb
Property ElementSize As System.Double
```
Property Value Element size (see Remarks )
VBA: See CWMeshControl::ElementSize .
Remarks: This property gets or sets the global element size for standard meshes and the maximum element size for curvature-based meshes. To get or set the minimum element size for curvature-based meshes, call ICWMeshControl::MinimumElementSizeForBlendedCurveMesher .
Examples: Visual Basic for Applications

### ICWMeshControl.EntityCount
Gets the number of entities in the mesh control.
```vb
ReadOnly Property EntityCount As System.Integer
```
Property Value Number of entities in the mesh control
VBA: See CWMeshControl::EntityCount .
Examples: Visual Basic for Applications

### ICWMeshControl.GetEntityAt
Obsolete. Superseded by ICWMeshControl::GetEntityAt2 .
```vb
Function GetEntityAt( ByVal NIndex As System.Integer, ByRef NSelectionType As System.Integer ) As System.Object
```
Parameters NIndex 0-based index of entity NSelectionType Type of entity as defined in swSelectType_e Return Value Entity
Return Value Entity
VBA: See CWMeshControl::GetEntityAt .
Examples: Visual Basic for Applications

### ICWMeshControl.GetEntityAt2
Obsolete. Superseded by ICWMeshControl::GetEntityAt3 .
```vb
Function GetEntityAt2( ByVal NIndex As System.Integer, ByVal BExcludeVertexAndReference As System.Integer, ByRef NSelectionType As System.Integer ) As System.Object
```
Parameters NIndex 0-based index of entity BExcludeVertexAndReference True to exclude the vertex and reference, false to not NSelectionType Type of entity as defined in swSelectType_e Return Value Entity
Return Value Entity
VBA: See CWMeshControl::GetEntityAt2 .
Remarks: To specify NIndex, use ICWMeshControl::EntityCount to determine the number of entities in this mesh control.
Examples: Visual Basic for Applications

### ICWMeshControl.GetEntityAt3
Gets the entity at the specified index for this mesh control.
```vb
Function GetEntityAt3( ByVal NIndex As System.Integer, ByVal BExcludeVertexAndReference As System.Boolean, ByRef NSelectionType As System.Integer ) As System.Object
```
Parameters NIndex 0-based index of entity BExcludeVertexAndReference -1 or true to exclude the vertex and reference, 0 or false to not NSelectionType Type of entity as defined in swSelectType_e Return Value Entity
Return Value Entity

### ICWMeshControl.InsertEntity
Adds an entity for mesh control.
```vb
Sub InsertEntity( ByVal DispEntity As System.Object )
```
Parameters DispEntity Entity
VBA: See CWMeshControl::InsertEntity .
Examples: Visual Basic for Applications

### ICWMeshControl.Layers
Obsolete. Not superseded.
```vb
Property Layers As System.Integer
```
Property Value Number of layers
VBA: See CWMeshControl::Layers .
Examples: Visual Basic for Applications

### ICWMeshControl.MeshControlBeginEdit
Starts editing the mesh control.
```vb
Sub MeshControlBeginEdit()
```
VBA: See CWMeshControl::MeshControlBeginEdit .
Remarks: To start editing a mesh control, you must call this method. You must call ICWMeshControl::MeshControlEndEdit to end editing a mesh control. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications

### ICWMeshControl.MeshControlEndEdit
Ends editing a mesh control.
```vb
Function MeshControlEndEdit() As System.Integer
```
Return Value Error as defined in swsMeshControlError_e
VBA: See CWMeshControl::MeshControlEndEdit .
Remarks: To end editing a mesh control, you must call this method. You must call ICWMeshControl::MeshControlBeginEdit to start editing a mesh control. Changes are not applied unless you call both methods.
Examples: Visual Basic for Applications

### ICWMeshControl.MinNumOfElementsPerCircleForBlendedCurveMesher
Gets and sets the minimum number of elements in a circle to determine the maximum angle in a curvature-based mesh.
```vb
Property MinNumOfElementsPerCircleForBlendedCurveMesher As System.Integer
```
Property Value Minimum number of elements in a circle
VBA: See CWMeshControl::MinNumOfElementsPerCircleForBlendedCurveMesher .
Remarks: This property is valid only if: ICWMeshControl::BeamSelected = 0 ICWMesh::MesherType = swsMesherType_e .swsMesherTypeAlternateCB - or - swsMesherType_e.swsMesherTypeAlternate
Examples: Visual Basic for Applications

### ICWMeshControl.MinimumElementSizeForBlendedCurveMesher
Gets and sets the minimum element size for a curvature-based mesh.
```vb
Property MinimumElementSizeForBlendedCurveMesher As System.Double
```
Property Value Minimum element size
VBA: See CWMeshControl::MinimumElementSizeForBlendedCurveMesher .
Remarks: This property is valid only if: ICWMeshControl::BeamSelected = 0 ICWMesh::MesherType = swsMesherType_e .swsMesherTypeAlternateCB To get and set the maximum element size for a curvature-based mesh, call ICWMeshControl::ElementSize .
Examples: Visual Basic for Applications

### ICWMeshControl.Name
Gets the name of the mesh control.
```vb
ReadOnly Property Name As System.String
```
Property Value Name of the mesh control
VBA: See CWMeshControl::Name .
Examples: Visual Basic for Applications

### ICWMeshControl.NumofElementsforBeams
Gets or sets the total number of mesh elements for selected beams.
```vb
Property NumofElementsforBeams As System.Integer
```
Property Value Number of mesh elements for selected beams
VBA: See CWMeshControl::NumofElementsforBeams .
Remarks: This property is valid only for beams.
Examples: Visual Basic for Applications

### ICWMeshControl.Ratio
Gets or sets the ratio of the average element size for element layer (i) to that of layer (i-1) for this mesh control.
```vb
Property Ratio As System.Double
```
Property Value Ratio of the average element size for element layer (i) to that of layer (i-1)
VBA: See CWMeshControl::Ratio .
Remarks: This property is valid only if ICWMesh::MesherType = swsMesherType_e .swsMesherTypeAlternate - or - swsMesherType_e.swsMesherTypeAlternateCB
Examples: Visual Basic for Applications

### ICWMeshControl.RemoveEntity
Removes the specified entity from the mesh control.
```vb
Sub RemoveEntity( ByVal NIndex As System.Integer )
```
Parameters NIndex 0-based index of entity to remove
VBA: See CWMeshControl::RemoveEntity .
Examples: Visual Basic for Applications

### ICWMeshControl.State
Gets whether the mesh control is suppressed.
```vb
ReadOnly Property State As System.Integer
```
Property Value Suppression state as defined in swsSuppressionState_e
VBA: See CWMeshControl::State .
Examples: Visual Basic for Applications

### ICWMeshControl.SuppressUnSuppress
Suppresses or unsuppresses the mesh control depending on its state .
```vb
Sub SuppressUnSuppress()
```
VBA: See CWMeshControl::SuppressUnSuppress .
Examples: Visual Basic for Applications

### ICWMeshControl.Units
Gets or sets the units for the mesh control.
```vb
Property Units As System.Integer
```
Property Value Mesh control linear units as defined in swsLinearUnit_e
VBA: See CWMeshControl::Units .
Examples: Visual Basic for Applications

### ICWMeshControl.UseSameElementSize
Obsolete. Superseded by ICWMeshControl::UseSameElementSize2 .
```vb
Property UseSameElementSize As System.Integer
```
Property Value 1 = Use same size for all components 0 = Use different sizes based on component mesh control
VBA: See CWMeshControl::UseSameElementSize .
Examples: Visual Basic for Applications

### ICWMeshControl.UseSameElementSize2
Gets or sets whether to use the the same element for all selected components in this mesh control.
```vb
Property UseSameElementSize2 As System.Boolean
```
Property Value -1 or true = Use same size for all components 0 or false = Use different sizes based on component mesh control
Remarks: This property returns a boolean value which can be cast to an integer. To set this property, you can specify either the boolean or the integer.

### ICWMeshControl.WeightFactor
Gets or sets the component weight factor for the mesh control.
```vb
Property WeightFactor As System.Double
```
Property Value Component weight factor as defined in swsMeshControlWeightFactor_e
VBA: See CWMeshControl::WeightFactor .
Remarks: The weight factor results in using a different element size for each selected component based on its significance.
Examples: Visual Basic for Applications

## ICWResults

### ICWResults.ActivatePlot
Activates the specified plot.
```vb
Function ActivatePlot( ByVal PlotName As System.String ) As System.Boolean
```
Parameters PlotName Name of plot to activate (see Remarks ) Return Value True if the plot is activated, false if not
Return Value True if the plot is activated, false if not
VBA: See CWResults::ActivatePlot .
Remarks: Call ICWResults::GetPlotNames to populate PlotName.
Examples: Visual Basic for Applications

### ICWResults.AddIsoClippingToPlot
Adds iso clipping to the specified result plot.
```vb
Function AddIsoClippingToPlot( ByVal PlotName As System.String, ByVal NIsoPlanes As System.Integer, ByVal IsoValues As System.Object, ByVal BShowClipSurface As System.Boolean, ByVal BShowCutUncut As System.Boolean ) As System.Integer
```
Parameters PlotName Name of result plot to which to add iso clipping (see Remarks ) NIsoPlanes Number of values in IsoValues IsoValues Array of values; each value corresponds to an iso surface to be plotted BShowClipSurface True to plot on iso surface only, false to not BShowCutUncut True to display result contours on the uncut portion of the model, false to display the uncut portion of the model in shaded view mode Return Value Error as defined by swsIsoClippingErrorCode_e
Return Value Error as defined by swsIsoClippingErrorCode_e
VBA: See CWResults::AddIsoClippingToPlot .
Remarks: Call ICWResults::GetPlotNames or ICWResults::IGetPlotNames to populate PlotName. For stress plots, iso clipping is available for nodal values only.
Examples: Visual Basic for Applications

### ICWResults.CreateDeformedBody
Obsolete. Superseded by ICWResults::CreateDeformedBody2 .
```vb
Sub CreateDeformedBody( ByVal NCreateAs As System.Integer, ByVal SName As System.String, ByRef ErrorCode As System.Integer )
```
Parameters NCreateAs Save option as defined in swsCreateDeformedBodyOption_e SName Name of the part or configuration to save ErrorCode Error code as defined in swsCreateDeformedBodyError_e
VBA: See CWResults::CreateDeformedBody .
Remarks: The new document is saved in the location of the original part or assembly. Deformed shapes of assembly documents are saved as multibody parts.
Examples: Visual Basic for Applications

### ICWResults.CreateDeformedBody2
Saves the deformed shape that results from running a static or nonlinear study.
```vb
Function CreateDeformedBody2( ByVal NCreateAs As System.Integer, ByVal NAdvOption As System.Integer, ByVal NStepNumber As System.Integer, ByVal SName As System.String, ByVal SFolderName As System.String ) As System.Integer
```
Parameters NCreateAs Save option as defined in swsCreateDeformedBodyOption_e NAdvOption Advanced save option as defined in swsCreateDeformedBodyAdvancedOption_e NStepNumber 1 for a static study or a valid step number for a nonlinear study (see Remarks ) SName Name of the deformed shape to save SFolderName Path name of folder in which to save SName Return Value 0 if successful, otherwise error code as defined in swsCreateDeformedBodyError_e
Return Value 0 if successful, otherwise error code as defined in swsCreateDeformedBodyError_e
VBA: See CWResults::CreateDeformedBody2 .
Remarks: Deformed shapes of assembly documents are saved as multibody parts. To populate NStepNumber for a nonlinear study, call ICWResults::GetMaximumAvailableSteps to get the maximum number of steps in the nonlinear study.
Examples: Visual Basic for Applications; Example Create Body From Deformed Shape

### ICWResults.CreatePlot
Creates the specified plot.
```vb
Function CreatePlot( ByVal NResultType As System.Integer, ByVal NComponent As System.Integer, ByVal NUnits As System.Integer, ByVal BValueByElem As System.Boolean, ByRef ErrorCode As System.Integer ) As CWPlot
```
Parameters NResultType Type of results plot as defined by swsPlotResultTypes_e (see Remarks ) NComponent Component to plot (see Remarks ) NUnits Units as appropriate to NComponent (see Remarks ) BValueByElem True to plot element values, false to plot node values ErrorCode Error code as defined in swsResultPlotErrorCode_e Return Value ICWPlot
Return Value ICWPlot
VBA: See CWResults::CreatePlot .
Remarks: If NResultType is swsPlotResultTypes_e... Then NComponent contains the component to plot as defined by... And NUnits contains the units as defined by... swsResultAcceleration swsAccelerationComponent_e swsAccelerationUnit_e swsResultBeamDiagram swsBeamForceType_e swsForceUnit_e swsResultBeamStress swsBeamStressType_e swsStrengthUnit_e swsResultDesignInsight N/A N/A swsResultDisplacementOrAmplitude swsDisplacementComponent_e (Displacement) swsFrequencyBucklingResultDisplacementComponentTypes_e (Amplitude for frequency and buckling studies) swsLinearUnit_e (displacement) or swsForceUnit_e (reaction force) swsResultEdgeWeldConnector N/A N/A swsResultFactorOfSafety swsFOS_NonCompositeCriterion_e (Non-composite shells) swsFOS_CompositeCriterion_e (Composite shells) N/A swsResultFatigue swsFatigueComponent_e N/A swsResultPinBoltBearing N/A N/A swsResultStrain swsStrainComponent_e swsStrengthUnit_e swsResultEquivalentStress* N/A swsStrengthUnit_e swsResultStress swsStressComponent_e swsStrengthUnit_e swsResultThermal swsThermalComponent_e swsTemperatureUnit_e swsResultVelocity swsVelocityComponent_e swsVelocityUnit_e * The following conditions are required to create an equivalent stress plot: Mixed mesh model with beams and solids/shells Dynamic study Nodal von Mises stress only (Faster) option is chosen in Results Options > Quantity to calculate stress results NOTE: After calling this method to create a plot of linear dynamic, frequency, or buckling studies, call ICWPlot::SetModeShape to set the mode shape number of the plot.
Examples: Visual Basic for Applications; Example Create Plots for Static Study

### ICWResults.CreateResultsEquationPlot
Creates a plot of the specified results equation.
```vb
Function CreateResultsEquationPlot( ByVal SResultsEquation As System.String, ByVal SLegendTitle As System.String, ByVal BValueByElem As System.Boolean, ByVal NUnits As System.Integer, ByVal NStepNumber As System.Integer, ByVal NShellOptions As System.Integer, ByRef ErrorCode As System.Integer ) As CWPlot
```
Parameters SResultsEquation Results equation SLegendTitle Title of the plot BValueByElem True to plot element values, false to plot node values NUnits Units as defined in swsUnit_e NStepNumber Number of solution step to plot NShellOptions Shell face as defined in swsShellFace_e ; valid only for shell and mixed mesh models ErrorCode Error code as defined in swsResultPlotErrorCode_e Return Value ICWPlot
Return Value ICWPlot
VBA: See CWResults::CreateResultsEquationPlot .
Examples: Visual Basic for Applications; Example Create Results Equation Plot

### ICWResults.CreateStressHotSpotPlot
Creates a stress hot spot plot.
```vb
Function CreateStressHotSpotPlot( ByVal BIsolateHotSpots As System.Boolean, ByVal BValueByNodes As System.Boolean, ByRef ErrorCode As System.Integer ) As CWPlot
```
Parameters BIsolateHotSpots True to isolate stress hot spots, false to not BValueByNodes True to plot node values, false to plot element values ErrorCode Error code as defined in swsStressHotSpotPlotError_e Return Value ICWPlot
Return Value ICWPlot
VBA: See CWResults::CreateStressHotSpotPlot .
Remarks: Before calling this method, you must call ICWResults::RunStressHotSpotDiagnostics .
Examples: Visual Basic for Applications

### ICWResults.DeletePlot
Deletes the specified plot.
```vb
Function DeletePlot( ByVal PlotName As System.String ) As System.Boolean
```
Parameters PlotName Name of the plot to delete (see Remarks ) Return Value True if the plot is deleted, false if not
Return Value True if the plot is deleted, false if not
VBA: See CWResults::DeletePlot .
Remarks: Call ICWResults::GetPlotNames to populate PlotName.
Examples: Visual Basic for Applications

### ICWResults.GetAccelerationComponentForAllStepsAtNode
Gets the acceleration for all solution steps at the specified node for the specified acceleration component.
```vb
Function GetAccelerationComponentForAllStepsAtNode( ByVal NComponent As System.Integer, ByVal NNodeNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Component of acceleration as defined in swsAccelerationComponent_e NNodeNum Node number DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component NUnits Units of acceleration as defined by swsAccelerationUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of acceleration component values
Return Value Array of acceleration component values
VBA: See CWResults::GetAccelerationComponentForAllStepsAtNode .
Examples: Visual Basic for Applications

### ICWResults.GetAccelerationForEntities
Gets the acceleration for the specified entities, solution step, and acceleration component.
```vb
Function GetAccelerationForEntities( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Component of acceleration as defined in swsAccelerationComponent_e NStepNumber Solution step number (use 1 for steady state) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NUnits Units of acceleration as defined by swsAccelerationUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of acceleration component values
Return Value Array of acceleration component values
VBA: See CWResults::GetAccelerationForEntities .
Examples: Visual Basic for Applications

### ICWResults.GetAverageStressesAtMidnodes
Gets whether to average stress at mid-nodes in this study.
```vb
Function GetAverageStressesAtMidnodes() As System.Boolean
```
Return Value True to average stresses at mid-nodes, false to not
VBA: See CWResults::GetAverageStressesAtMidnodes .
Remarks: This method is valid only for static, nonlinear, drop test, pressure vessel, and topology studies.
Examples: Visual Basic for Applications

### ICWResults.GetBeamForcesForEntities
Gets the force values for the specified force for the selected beams.
```vb
Function GetBeamForcesForEntities( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Force or moment or torque load to apply to selected entities as defined by swsBeamForceType_e NStepNumber Number of steps to complete a non-linear analysis; otherwise, value is 1 ArraySelectedEntities Selected beams NUnits Unit as defined by swsUnit_e ErrorCode Error as defined by swsResultsError_e Return Value Array of the values of the specified forces for the elements of the selected beams. Every element has two ends, so the output is as follows: Beam element number End1 force End2 force Beam element number End1 force End2 force etc.
Return Value Array of the values of the specified forces for the elements of the selected beams. Every element has two ends, so the output is as follows: Beam element number End1 force End2 force Beam element number End1 force End2 force etc.
VBA: See CWResults::GetBeamForcesForEntities .
Examples: Visual Basic for Applications; Example Get Forces and Stresses for Selected Beams

### ICWResults.GetBeamMinMaxStress
Gets the minimum and maximum elemental stresses for beams at the specified solution step.
```vb
Function GetBeamMinMaxStress( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Beam stresses as defined in swsBeamStressComponent_e NStepNumber Solution step number (use 1 for static) NUnits Stress units as defined in swsStrengthUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of stress results for the beams
Return Value Array of stress results for the beams
VBA: See CWResults::GetBeamMinMaxStress .
Examples: Visual Basic for Applications

### ICWResults.GetBeamRadius
Gets the radii of beams used to render the specified plot.
```vb
Function GetBeamRadius( ByVal SPlotName As System.String, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters SPlotName Name of plot ErrorCode Error as defined in swsResultsError_e Return Value Array of beam radii
Return Value Array of beam radii
VBA: See CWResults::GetBeamRadius .
Examples: Visual Basic for Applications

### ICWResults.GetBeamStressForEntities
Gets the stress values for the specified type of stress for the selected beams.
```vb
Function GetBeamStressForEntities( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Type of stress as defined in swsBeamStressType_e NStepNumber Number of steps to complete a non-linear analysis; otherwise, value is 1 ArraySelectedEntities Selected beams NUnits Unit as defined by swsUnit_e ErrorCode Error as defined by swsResultsError_e Return Value Array of the stress values for the specified type of stress for the elements of the selected beams. Every element has two ends, so the output is as follows: Beam element number End1 stress End2 stress Beam element number End1 stress End2 stress etc.
Return Value Array of the stress values for the specified type of stress for the elements of the selected beams. Every element has two ends, so the output is as follows: Beam element number End1 stress End2 stress Beam element number End1 stress End2 stress etc.
VBA: See CWResults::GetBeamStressForEntities
Examples: Visual Basic for Applications; Example Get Forces and Stresses for Selected Beams

### ICWResults.GetBucklingLoadFactors
Gets the buckling load factors.
```vb
Function GetBucklingLoadFactors( ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters ErrorCode Error as defined in swsResultsError_e Return Value Array of buckling load factors
Return Value Array of buckling load factors
VBA: See CWResults::GetBucklingLoadFactors .
Examples: Visual Basic for Applications

### ICWResults.GetConnectorForces
Obsolete. Superseded by ICWResults::GetConnectorForces2 .
```vb
Function GetConnectorForces( ByVal NStepNumber As System.Integer, ByVal SName As System.String, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NStepNumber Solution step number (use 1 for static) SName Name of a pin, bolt, or bearing connector NUnits Output units as defined in swsUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of twelve doubles of the connector forces: x-component of the shear force y-component of the shear force z-component of the shear force Resultant shear force x-component of the axial force y-component of the axial force z-component of the axial force Resultant axial force x-component of the torsional moment y-component of the torsional moment z-component of the torsional moment Resultant torsional moment
Return Value Array of twelve doubles of the connector forces: x-component of the shear force y-component of the shear force z-component of the shear force Resultant shear force x-component of the axial force y-component of the axial force z-component of the axial force Resultant axial force x-component of the torsional moment y-component of the torsional moment z-component of the torsional moment Resultant torsional moment
VBA: See CWResults::GetConnectorForces .
Examples: Visual Basic for Applications

### ICWResults.GetConnectorForces2
Gets the x-, y-, and z-component and resultant forces, bending moments, and torques for the specified connector and step.
```vb
Function GetConnectorForces2( ByVal NStepNumber As System.Integer, ByVal SName As System.String, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NStepNumber Solution step number (use 1 for static) SName Name of a pin, bolt, bearing, spring, spot weld, edge weld, linkage rod connector, or link connector (see Remarks ) NUnits Output units as defined in swsUnit_e ErrorCode Error code as defined in swsResultsError_e Return Value Array of 16 doubles: x-component of the shear force y-component of the shear force z-component of the shear force Resultant shear force x-component of the axial force y-component of the axial force z-component of the axial force Resultant axial force x-component of the bending moment y-component of the bending moment z-component of the bending moment Resultant bending moment x-component of the torque y-component of the torque z-component of the torque Resultant torque
Return Value Array of 16 doubles: x-component of the shear force y-component of the shear force z-component of the shear force Resultant shear force x-component of the axial force y-component of the axial force z-component of the axial force Resultant axial force x-component of the bending moment y-component of the bending moment z-component of the bending moment Resultant bending moment x-component of the torque y-component of the torque z-component of the torque Resultant torque
VBA: See CWResults::GetConnectorForces2 .
Remarks: For spot welds, specify SName with the spot weld point (e.g., "Spot Welds-1/Point1").
Examples: Visual Basic for Applications; Torques

### ICWResults.GetConnectorForcesWithTimeValue
Gets the component and resultant forces, bending moments, and torques for the specified connector at the specified time.
```vb
Function GetConnectorForcesWithTimeValue( ByVal DTimeValue As System.Double, ByVal SName As System.String, ByVal NUnits As System.Integer, ByRef DAccurateTimeValue As System.Double, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters DTimeValue Time at which to get connector values SName Name of a pin, bolt, bearing, spring, spot weld, edge weld, or link connector (see Remarks ) NUnits Output units as defined in swsUnit_e DAccurateTimeValue Accurate time ErrorCode Error code as defined in swsResultsError_e Return Value Array of 16 doubles: x-component of the shear force y-component of the shear force z-component of the shear force Resultant shear force x-component of the axial force y-component of the axial force z-component of the axial force Resultant axial force x-component of the bending moment y-component of the bending moment z-component of the bending moment Resultant bending moment x-component of the torque y-component of the torque z-component of the torque Resultant torque
Return Value Array of 16 doubles: x-component of the shear force y-component of the shear force z-component of the shear force Resultant shear force x-component of the axial force y-component of the axial force z-component of the axial force Resultant axial force x-component of the bending moment y-component of the bending moment z-component of the bending moment Resultant bending moment x-component of the torque y-component of the torque z-component of the torque Resultant torque
VBA: See CWResults::GetConnectorForcesWithTimeValue .
Remarks: For spot welds, specify SName with the spot weld point (e.g., "Spot Welds-1/Point1").
Examples: Visual Basic for Applications; Torques at a Time

### ICWResults.GetContactForcesAndFriction
Gets the specified contact or friction forces for the specified entities.
```vb
Function GetContactForcesAndFriction( ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NForceType As System.Integer, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NStepNumber Solution step number at which to calculate the contact forces for non-linear and drop test studies; 1 for static studies DispPlane Plane, axis, or coordinate system relative to which the contact forces are calculated ArraySelectedEntities Array of entities for which to calculate contact forces NForceType Type of contact force as defined in swsNForceType_e NUnits Units as defined in swsUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of four doubles of the contact forces for all selections in ArraySelectedEntities: Summation of x-components of contact forces Summation of y-components of contact forces Summation of z-components of contact forces Resultant contact force
Return Value Array of four doubles of the contact forces for all selections in ArraySelectedEntities: Summation of x-components of contact forces Summation of y-components of contact forces Summation of z-components of contact forces Resultant contact force
VBA: See CWResults::GetContactForcesAndFriction .
Examples: Visual Basic for Applications; Example Get Contact or Friction Forces

### ICWResults.GetDeformedBodyFailedSewOption
Gets the option to use when a deformed body fails to sew into a solid object.
```vb
Function GetDeformedBodyFailedSewOption() As System.Integer
```
Return Value Option to use when a deformed body fails to sew into a solid object as defined in swsCreateDeformedBodyFailedSewOption_e
VBA: See CWResults::GetDeformedBodyFailedSewOption .
Examples: Visual Basic for Applications; Example Create Body From Deformed Shape

### ICWResults.GetDeformedCoord
Gets nodal deform coordinates for the specified plot.
```vb
Function GetDeformedCoord( ByVal SPlotName As System.String, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters SPlotName Name of plot ErrorCode Error as defined in swsResultsError_e Return Value Array of x, y, z coordinates for all deform nodes
Return Value Array of x, y, z coordinates for all deform nodes
VBA: See CWResults::GetDeformedCoord .
Examples: Visual Basic for Applications; Example Create Body From Deformed Shape

### ICWResults.GetDetectedHotSpotElements
Gets the stress hot spot elements.
```vb
Function GetDetectedHotSpotElements( ByRef RetVal As System.Object ) As System.Integer
```
Parameters RetVal Array of stress hot spot elements Return Value 1 if successful, 0 if not
Return Value 1 if successful, 0 if not
VBA: See CWResults::GetDetectedHotSpotElements .
Remarks: Before calling this method, you must call ICWResults::RunStressHotSpotDiagnostics .
Examples: Visual Basic for Applications

### ICWResults.GetDetectedHotSpotNodes
Gets the stress hot spot nodes.
```vb
Function GetDetectedHotSpotNodes( ByRef RetVal As System.Object ) As System.Integer
```
Parameters RetVal Array of stress hot spot nodes Return Value 1 if successful, 0 if not
Return Value 1 if successful, 0 if not
VBA: See CWResults::GetDetectedHotSpotNodes .
Remarks: Before calling this method, you must call ICWResults::RunStressHotSpotDiagnostics .
Examples: Visual Basic for Applications

### ICWResults.GetDisplacementAtPoints
Gets the specified displacement component at the specified points for the specified solution step.
```vb
Function GetDisplacementAtPoints( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArrayPoints As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Displacement component as defined in swsDisplacementComponent_e NStepNumber Solution step number (use 1 for static) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArrayPoints Array of point coordinates NUnits Linear units as defined in swsLinearUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of displacement component values
Return Value Array of displacement component values
VBA: See CWResults::GetDisplacementAtPoints .
Examples: Visual Basic for Applications

### ICWResults.GetDisplacementComponentForAllStepsAtNode
Gets the specified displacement component at the specified node for all solutions steps.
```vb
Function GetDisplacementComponentForAllStepsAtNode( ByVal NComponent As System.Integer, ByVal NNodeNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Displacement component as defined in swsDisplacementComponent_e NNodeNum Node number DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component NUnits Linear units as defined in swsLinearUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of displacement component values
Return Value Array of displacement component values
VBA: See CWResults::GetDisplacementComponentForAllStepsAtNode .
Examples: Visual Basic for Applications

### ICWResults.GetDisplacementForEntities
Gets the specified displacement component for the specified entities at the specified solution step.
```vb
Function GetDisplacementForEntities( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Displacement component as defined in swsDisplacementComponent_e NStepNumber Solution step number (use 1 for static) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NUnits Linear units for displacement translation as defined in swsLinearUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of displacement component values
Return Value Array of displacement component values
VBA: See CWResults::GetDisplacementForEntities .
Examples: Visual Basic for Applications

### ICWResults.GetEdgeWeldResults
Gets the edge weld results for the specified edge weld connector.
```vb
Function GetEdgeWeldResults( ByVal EdgeWeldConnector As System.String, ByVal NUnit As System.Integer, ByRef BPassFail As System.Boolean, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters EdgeWeldConnector Name of the edge weld connector for which to get edge weld results NUnit Unit as defined in swsUnit_e BPassFail True if the edge weld is safe to use, false if not ErrorCode Error code as defined in swsWeldResultErrorCode_e Return Value Array of edge weld results (see Remarks )
Return Value Array of edge weld results (see Remarks )
VBA: See CWResults::GetEdgeWeldResults .
Remarks: The array of edge weld results contains these minimum, maximum, and mean values in this order: Weld size Weld throat size Joint normal force Shear-weld axis force Shear-surface normal force Bending moment
Examples: Visual Basic for Applications; Calculate Edge Weld Results

### ICWResults.GetEnvelopeAccelerationForEntities
Gets the specified envelope acceleration for the specified entities across all solution steps.
```vb
Function GetEnvelopeAccelerationForEntities( ByVal NComponent As System.Integer, ByVal NEnvelopeType As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Component of acceleration as defined in swsAccelerationComponent_e NEnvelopeType Envelope plot type as defined in swsEnvelopePlotType_e DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NUnits Units of acceleration as defined by swsAccelerationUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of acceleration component values
Return Value Array of acceleration component values
VBA: See CWResults::GetEnvelopeAccelerationForEntities .
Examples: Visual Basic for Applications

### ICWResults.GetEnvelopeDisplacementForEntities
Gets the specified envelope displacement for the specified entities across all solution steps.
```vb
Function GetEnvelopeDisplacementForEntities( ByVal BValueByNode As System.Boolean, ByVal NComponent As System.Integer, ByVal NEnvelopeType As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByVal NShellFace As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BValueByNode True to get the displacement value by node, false to get the stress value by element NComponent Displacement component as defined in swsDisplacementComponent_e NEnvelopeType Envelope plot type as defined in swsEnvelopePlotType_e DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NUnits Linear units as defined in swsLinearUnit_e NShellFace Option as defined in swsShellFace_e ErrorCode Error as defined in swsResultsError_e Return Value Array of displacement component values
Return Value Array of displacement component values
VBA: See CWResults::GetEnvelopeDisplacementForEntities .
Examples: Visual Basic for Applications

### ICWResults.GetEnvelopeStrainForEntities
Gets the specified envelope strain for the specified entities across all solution steps.
```vb
Function GetEnvelopeStrainForEntities( ByVal BValueByNode As System.Boolean, ByVal NComponent As System.Integer, ByVal NEnvelopeType As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NShellFace As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BValueByNode True to get the strain value by node, false to get the stress value by element NComponent Strain component as defined in swsStrainComponent_e NEnvelopeType Envelope plot type as defined in swsEnvelopePlotType_e DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NShellFace Option as defined in swsShellFace_e ErrorCode Error as defined in swsResultsError_e Return Value Array of strain component values
Return Value Array of strain component values
VBA: See CWResults::GetEnvelopeStrainForEntities .
Examples: Visual Basic for Applications

### ICWResults.GetEnvelopeStressForEntities
Gets the specified envelope stress for the specified entities across all solution steps.
```vb
Function GetEnvelopeStressForEntities( ByVal BValueByNode As System.Boolean, ByVal NComponent As System.Integer, ByVal NEnvelopeType As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByVal NShellFace As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BValueByNode True to get the stress value by node, false to get the stress value by element NComponent Stress component as defined in swsStressComponent_e NEnvelopeType Envelope plot type as defined in swsEnvelopePlotType_e DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NUnits Unit as defined in swsStrengthUnit_e NShellFace Option as defined in swsShellFace_e ErrorCode Error as defined in swsResultsError_e Return Value Array of stress component values
Return Value Array of stress component values
VBA: See CWResults::GetEnvelopeStressForEntities .
Examples: Visual Basic for Applications

### ICWResults.GetEnvelopeVelocityForEntities
Gets the specified envelope velocity for the specified entities across all solution steps.
```vb
Function GetEnvelopeVelocityForEntities( ByVal NComponent As System.Integer, ByVal NEnvelopeType As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Velocity component as defined in swsVelocityComponent_e NEnvelopeType Envelope plot type as defined in swsEnvelopePlotType_e DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NUnits Units of velocity as defined in swsVelocityUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of velocity component values
Return Value Array of velocity component values
VBA: See CWResults::GetEnvelopeVelocityForEntities .
Examples: Visual Basic for Applications

### ICWResults.GetFactorOfSafetyForComposites
Gets the range of stress or buckling factors of safety for the specified shell type, surface bodies, failure criteria, and other options.
```vb
Function GetFactorOfSafetyForComposites( ByVal BIsComposite As System.Boolean, ByVal ArraySelectedEntities As System.Object, ByVal NonCompositeCriterion As System.Integer, ByVal CompositeCriterion As System.Integer, ByVal PlyNumber As System.Integer, ByVal ShellFaceOption As System.Integer, ByVal NormalShellFaceOption As System.Integer, ByVal FosDistributionOption As System.Integer, ByVal MinFOSVal As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BIsComposite True for a composite shell type, false for a thin or thick shell type ArraySelectedEntities Array of selected surface body objects NonCompositeCriterion Criterion for non-composite shells as defined in swsFOS_NonCompositeCriterion_e CompositeCriterion Criterion for composite shells as defined in swsFOS_CompositeCriterion_e PlyNumber Number of plies; 0 = across all plies (worst case) ShellFaceOption Shell face option as defined in swsFOS_ShellfaceOption_e ; valid only if PlyNumber > 0 NormalShellFaceOption Normal shell face option as defined in swsFOS_NormalShellFaceOption_e FosDistributionOption Plot type as defined in swsFOS_DistributionOpt_e MinFOSVal Minimum factor of safety; default is 1 ErrorCode Result status as defined in swsFOS_ErrorCode_e Return Value Array of two doubles of the minimum and maximum factors of safety
Return Value Array of two doubles of the minimum and maximum factors of safety
VBA: See CWResults::GetFactorOfSafetyForComposites .
Examples: Visual Basic for Applications; Example Get Factor of Safety Values for Composites

### ICWResults.GetFatigueForEntities
Gets the specified fatigue component for the specified entities.
```vb
Function GetFatigueForEntities( ByVal NComponent As System.Integer, ByVal ArraySelectedEntities As System.Object, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Fatigue component as defined in swsFatigueComponent_e ArraySelectedEntities Array of geometric entities ErrorCode Error as defined in swsResultsError_e Return Value Array of fatigue component values
Return Value Array of fatigue component values
VBA: See CWResults::GetFatigueForEntities .
Examples: Visual Basic for Applications

### ICWResults.GetFreeBodyForcesAndMoments
Gets the x-, y-, and z-component and resultant free body forces and moments for the specified entities and the entire model.
```vb
Function GetFreeBodyForcesAndMoments( ByVal DispPlane As System.Object, ByVal SelectedRefPoint As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters DispPlane Plane, axis, or coordinate system relative to which the results are listed SelectedRefPoint Reference point used to list the moments for the forces that are available in ArraySelectedEntities; NULL to not list the moments ArraySelectedEntities Selected faces, edges, vertices, and components NUnits Output units as defined in swsUnit_e ErrorCode Error as defined in swsResultsError_e Return Value If SelectedRefPoint is NULL: Array of eight doubles of the free body forces for all selections in ArraySelectedEntities and for the entire model: Summation of x-components of free body force for all selections Summation of y-components of free body force for all selections Summation of z-components of free body force for all selections Resultant free body force for all selections Summation of x-components of free body moment for the entire model Summation of y-components of free body moment for the entire model Summation of z-components of free body moment for the entire model Resultant free body moment for the entire model If SelectedRefPoint is not NULL: Array of sixteen doubles of the free body forces and moments for all selections in ArraySelectedEntities and for the entire model: Summation of x-components of free body force for all selections Summation of y-components of free body force for all selections Summation of z-components of free body force for all selections Resultant free body force for all selections Summation of x-components of free body moment for all selections Summation of y-components of free body moment for all selections Summation of z-components of free body moment for all selections Resultant free body moment for all selections Summation of x-components of free body force for the entire model Summation of y-components of free body force for the entire model Summation of z-components of free body force for the entire model Resultant free body force for the entire model Summation of x-components of free body moment for the entire model Summation of y-components of free body moment for the entire model Summation of z-components of free body moment for the entire model Resultant free body moment for the entire model
Return Value If SelectedRefPoint is NULL: Array of eight doubles of the free body forces for all selections in ArraySelectedEntities and for the entire model: Summation of x-components of free body force for all selections Summation of y-components of free body force for all selections Summation of z-components of free body force for all selections Resultant free body force for all selections Summation of x-components of free body moment for the entire model Summation of y-components of free body moment for the entire model Summation of z-components of free body moment for the entire model Resultant free body moment for the entire model If SelectedRefPoint is not NULL: Array of sixteen doubles of the free body forces and moments for all selections in ArraySelectedEntities and for the entire model: Summation of x-components of free body force for all selections Summation of y-components of free body force for all selections Summation of z-components of free body force for all selections Resultant free body force for all selections Summation of x-components of free body moment for all selections Summation of y-components of free body moment for all selections Summation of z-components of free body moment for all selections Resultant free body moment for all selections Summation of x-components of free body force for the entire model Summation of y-components of free body force for the entire model Summation of z-components of free body force for the entire model Resultant free body force for the entire model Summation of x-components of free body moment for the entire model Summation of y-components of free body moment for the entire model Summation of z-components of free body moment for the entire model Resultant free body moment for the entire model
VBA: See CWResults::GetFreeBodyForcesAndMoments .
Examples: Visual Basic for Applications; Example Get Free Body Forces and Moments

### ICWResults.GetFreeBodyForcesAndMomentsForAStep
Gets the x-, y-, and z-component and resultant free body forces and moments for the specified entities for the specified step.
```vb
Function GetFreeBodyForcesAndMomentsForAStep( ByVal DispPlane As System.Object, ByVal SelectedRefPoint As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByVal NStepNumber As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters DispPlane Plane, axis, or coordinate system relative to which the results are listed SelectedRefPoint Reference point used to list the moments for the forces that are available in ArraySelectedEntities; NULL to not list the moments ArraySelectedEntities Selected faces, edges, vertices, and components NUnits Output units as defined in swsUnit_e NStepNumber Step number ErrorCode Error as defined in swsResultsError_e Return Value If SelectedRefPoint is NULL: Array of eight doubles of the free body forces for all selections in ArraySelectedEntities and for the entire model: Summation of x-components of free body force for all selections Summation of y-components of free body force for all selections Summation of z-components of free body force for all selections Resultant free body force for all selections Summation of x-components of free body moment for the entire model Summation of y-components of free body moment for the entire model Summation of z-components of free body moment for the entire model Resultant free body moment for the entire model If SelectedRefPoint is not NULL: Array of sixteen doubles of the free body forces and moments for all selections in ArraySelectedEntities and for the entire model: Summation of x-components of free body force for all selections Summation of y-components of free body force for all selections Summation of z-components of free body force for all selections Resultant free body force for all selections Summation of x-components of free body moment for all selections Summation of y-components of free body moment for all selections Summation of z-components of free body moment for all selections Resultant free body moment for all selections Summation of x-components of free body force for the entire model Summation of y-components of free body force for the entire model Summation of z-components of free body force for the entire model Resultant free body force for the entire model Summation of x-components of free body moment for the entire model Summation of y-components of free body moment for the entire model Summation of z-components of free body moment for the entire model Resultant free body moment for the entire model
Return Value If SelectedRefPoint is NULL: Array of eight doubles of the free body forces for all selections in ArraySelectedEntities and for the entire model: Summation of x-components of free body force for all selections Summation of y-components of free body force for all selections Summation of z-components of free body force for all selections Resultant free body force for all selections Summation of x-components of free body moment for the entire model Summation of y-components of free body moment for the entire model Summation of z-components of free body moment for the entire model Resultant free body moment for the entire model If SelectedRefPoint is not NULL: Array of sixteen doubles of the free body forces and moments for all selections in ArraySelectedEntities and for the entire model: Summation of x-components of free body force for all selections Summation of y-components of free body force for all selections Summation of z-components of free body force for all selections Resultant free body force for all selections Summation of x-components of free body moment for all selections Summation of y-components of free body moment for all selections Summation of z-components of free body moment for all selections Resultant free body moment for all selections Summation of x-components of free body force for the entire model Summation of y-components of free body force for the entire model Summation of z-components of free body force for the entire model Resultant free body force for the entire model Summation of x-components of free body moment for the entire model Summation of y-components of free body moment for the entire model Summation of z-components of free body moment for the entire model Resultant free body moment for the entire model
VBA: See CWResults::GetFreeBodyForcesAndMomentsForAStep .
Examples: Visual Basic for Applications

### ICWResults.GetHeatPowerOrEnergy
Obsolete. Superseded by ICWResults::GetHeatPowerOrEnergy2 .
```vb
Function GetHeatPowerOrEnergy( ByVal BHeatEnergy As System.Integer, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByVal NStepNumber As System.Integer, ByVal NStepNumber2 As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BHeatEnergy 0 for heat power, 1 for heat energy ArraySelectedEntities Array of geometric entities NUnits Units as defined in swsUnit_e NStepNumber Solution step number (use 1 for steady state) NStepNumber2 End solution step number (use 1 for steady state); valid only if BHeatEnergy is 1 ErrorCode Error as defined in swsResultsError_e Return Value Array of heat values (see Remarks )
Return Value Array of heat values (see Remarks )
VBA: See CWResults::GetHeatPowerOrEnergy .
Remarks: This method returns the following array : { sel_power_or_energy_in , sel_power_or_energy_out , sel_net_power , model_power_or_energy_in , model_power_or_energy_out , model_net_power_or_energy }, where sel_power_or_energy_in and sel_power_or_energy_out are the sum of individual entity in and out values, and model_power_or_energy_in and model_power_or_energy_out are total in and out values for the entire model. sel_net_power = ( sel_power_or_energy_in - sel_power_or_energy_out) model_net_power_or_energy = ( model_power_or_energy_in - model_power_or_energy_out)
Examples: Visual Basic for Applications

### ICWResults.GetHeatPowerOrEnergy2
Gets the heat power or heat energy for the specified entities.
```vb
Function GetHeatPowerOrEnergy2( ByVal BHeatEnergy As System.Boolean, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByVal NStepNumber As System.Integer, ByVal NStepNumber2 As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BHeatEnergy 0 or false for heat power, -1 or true for heat energy ArraySelectedEntities Array of geometric entities NUnits Units as defined in swsUnit_e NStepNumber Solution step number (use 1 for steady state) NStepNumber2 End solution step number (use 1 for steady state); valid only if BHeatEnergy is 1 ErrorCode Error as defined in swsResultsError_e Return Value Array of heat values (see Remarks )
Return Value Array of heat values (see Remarks )
Remarks: This method returns the following array : { sel_power_or_energy_in , sel_power_or_energy_out , sel_net_power , model_power_or_energy_in , model_power_or_energy_out , model_net_power_or_energy }, where sel_power_or_energy_in and sel_power_or_energy_out are the sum of individual entity in and out values, and model_power_or_energy_in and model_power_or_energy_out are total in and out values for the entire model. sel_net_power = ( sel_power_or_energy_in - sel_power_or_energy_out) model_net_power_or_energy = ( model_power_or_energy_in - model_power_or_energy_out)
Examples: Example Get Thermal Values

### ICWResults.GetLegendContourColors
Gets the contour colors based on user-input colors.
```vb
Function GetLegendContourColors( ByVal SPlotName As System.String, ByRef VarColor As System.Object ) As System.Integer
```
Parameters SPlotName Plot name (see Remarks ) VarColor Array of colors (see Remarks ) Return Value Result code as defined in swsResultPlotColorOption_ErrorCode_e
Return Value Result code as defined in swsResultPlotColorOption_ErrorCode_e
VBA: See CWResults::GetLegendContourColors .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. VarColor contains an array of RGB triplets. Each triplet represents a computed legend color: [ R[0,255] G[0,255] B[0,255] ... ] If there are 14 colors in the legend, then the size of VarColor array is (14 * 3) = 42.
Examples: Visual Basic for Applications; Example Analyze Part

### ICWResults.GetLinearizedStressValues
Obsolete. Superseded by ICWPlot::GetLinearizedStressValuesAlongSCL .
```vb
Function GetLinearizedStressValues( ByVal NComponent As System.Integer, ByVal DPoint1X As System.Double, ByVal DPoint1Y As System.Double, ByVal DPoint1Z As System.Double, ByVal DPoint2X As System.Double, ByVal DPoint2Y As System.Double, ByVal DPoint2Z As System.Double, ByVal NIntermPoints As System.Integer, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Stress component for which to get linearized values (see Remarks ) DPoint1X X coordinate of first point on the section view (see Remarks ) DPoint1Y Y coordinate of first point on the section view (see Remarks ) DPoint1Z Z coordinate of first point on the section view (see Remarks ) DPoint2X X coordinate of second point on the section view (see Remarks ) DPoint2Y Y coordinate of second point on the section view (see Remarks ) DPoint2Z Z coordinate of second point on the section view (see Remarks ) NIntermPoints Number of intermediate points between the points specified by DPoint1X, DPoint1Y, DPoint1Z, DPoint2X, DPoint2Y, and DPoint2Z to define the resolution of the line NUnits Units as defined in swsStrengthUnit_e (see Remarks ) ErrorCode Error as defined in swsResultsError_e Return Value Array of seven linearized membrane and bending stress values for the specified NComponent (see Remarks )
Return Value Array of seven linearized membrane and bending stress values for the specified NComponent (see Remarks )
VBA: See CWResults::GetLinearizedStressValues .
Remarks: Before calling this method: Create a pressure vessel study using the results of two static studies. Analyse the pressure vessel study to generate a stress results plot. Create a planar section view of the stress results plot. Right-click the stress results plot and click Section Clipping to open the Section PropertyManager. In Section 1, click Plane . In Options, click Union and select Show mesh on section plane and Plot on section only . Click OK . Specify DPoint1X, DPoint1Y, DPoint1Z, DPoint2X, DPoint2Y, DPoint2Z with the coordinates of two points on the planar section view you created in step 3. Specify NComponent with one of the following values: 0 = Normal stress in X-direction 1 = Normal stress in Y-direction 2 = Normal stress in Z-direction 3 = Shear stress in the Y direction acting in the YZ plane 4 = Shear stress in the Z direction acting in the YZ plane 5 = Shear stress in the Z direction acting in the XZ plane 6 = Von Mises stress 7 = Stress intensity (P1 - P3) 8 = Normal stress in the first principal direction (P1) 9 = Normal stress in the second principal direction (P2) 10 = Normal stress in the third principal direction (P3) The returned array contains the following seven values in the specified NUnits for the specified NComponent: { membrane_stress , bending_point_1 , ( membrane_stress + bending_point_1 ), bending_point_2 , ( membrane_stress + bending_point_2 ), peak_point_1 , peak_point_2 } NOTE: This method is valid only for solid meshes.
Examples: Visual Basic for Applications

### ICWResults.GetMassParticipation
Gets the mass participation factors in the X, Y, and Z directions for all modes.
```vb
Function GetMassParticipation( ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters ErrorCode Error as defined in swsResultsError_e Return Value Array of mass participation factors (in X, Y, Z directions)
Return Value Array of mass participation factors (in X, Y, Z directions)
VBA: See CWResults::GetMassParticipation .
Examples: Visual Basic for Applications; Example Create Frequency Study with Solid Mesh

### ICWResults.GetMaximumAvailableSteps
Gets the maximum solution step number for which results exist.
```vb
Function GetMaximumAvailableSteps() As System.Integer
```
Return Value Maximum solution step number
VBA: See CWResults::GetMaximumAvailableSteps .
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study; Create Drop Test Study

### ICWResults.GetMinMaxAcceleration
Gets the algebraic minimum and maximum for the specified acceleration component and solution step.
```vb
Function GetMinMaxAcceleration( ByVal NComponent As System.Integer, ByVal NStepNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Component of acceleration as defined in swsAccelerationComponent_e NStepNum Solution step number (use 1 for steady state) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component NUnits Units of acceleration as defined by swsAccelerationUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxAcceleration .
Remarks: This method returns the following array: { node_with_minimum_acceleration , minimum_acceleration , node_with_maximum_acceleration , maximum_acceleration }, where the nodes are integer indexes, and the accelerations are in scientific notation.
Examples: Visual Basic for Applications; Example Create Linear Dynamic Study

### ICWResults.GetMinMaxDisplacement
Gets the algebraic minimum and maximum displacement for the specified component and solution step.
```vb
Function GetMinMaxDisplacement( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Displacement component as defined in swsDisplacementComponent_e NStepNumber Solution step number (use 1 for static) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist NUnits Linear units for displacement translation as defined in swsLinearUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxDisplacement .
Remarks: This method returns the following array: { node_with_minimum_displacement , minimum_displacement , node_with_maximum_displacement , maximum_displacement }, where the nodes are integer indexes, and the displacements are in scientific notation. For random vibration studies, this method returns power spectral density (PSD) values. To obtain root-mean-square (RMS) values in random vibration studies, call ICWResults::GetMinMaxDisplacementRMS .
Examples: Visual Basic for Applications; Create Linear Dynamic Study

### ICWResults.GetMinMaxDisplacementForHarmonic
Gets the algebraic minimum and maximum displacement for the specified component and solution step of this harmonic study.
```vb
Function GetMinMaxDisplacementForHarmonic( ByVal NComponent As System.Integer, ByVal NStepNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Displacement component as defined in swsDisplacementComponent_e NStepNum Solution step number (use 1 for static) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist NUnits Linear units for displacement translation as defined in swsLinearUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxDisplacementForHarmonic .
Remarks: This method returns the following array: { node_with_minimum_displacement , minimum_displacement , node_with_maximum_displacement , maximum_displacement }, where the nodes are integer indexes, and the displacements are in scientific notation.
Examples: Visual Basic for Applications

### ICWResults.GetMinMaxDisplacementRMS
Gets the algebraic minimum and maximum displacement RMS values for the specified component of this random vibration study.
```vb
Function GetMinMaxDisplacementRMS( ByVal NComponent As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Displacement component as defined in swsDisplacementComponent_e DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist NUnits Linear units for displacement translation as defined in swsLinearUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxDisplacementRMS .
Remarks: This method returns the following array: { node_with_minimum_displacement , minimum_displacement , node_with_maximum_displacement , maximum_displacement }, where the nodes are integer indexes, and the displacements are in scientific notation. This method returns root-mean-square (RMS) values. To obtain power spectral density (PSD) values, call ICWResults::GetMinMaxDisplacement .
Examples: Visual Basic for Applications

### ICWResults.GetMinMaxFactorOfSafety
Obsolete. Superseded by ICWResults::GetMinMaxFactorOfSafety2 .
```vb
Function GetMinMaxFactorOfSafety( ByVal BAllBodies As System.Boolean, ByVal ArraySelectedEntities As System.Object, ByVal NComponent As System.Integer, ByVal NShellOptions As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BAllBodies True to select all bodies to plot the factor of safety, false to select specific bodies ArraySelectedEntities Array of bodies for which to plot the factor of safety; valid only if BAllBodies is set to false NComponent Failure criterion as defined by swsFOS_NonCompositeCriterion_e (see Remarks ) NShellOptions Shell face on which to perform the factor of safety as defined by swsFOS_ShellFaceOption_e ErrorCode Error code as defined by swsFosPlotErrorCode_e Return Value Array of two doubles of the minimum and maximum factors of safety
Return Value Array of two doubles of the minimum and maximum factors of safety
VBA: See CWResults::GetMinMaxFactorOfSafety .
Remarks: For pure beam studies, you can only specify NComponent with swsFOS_NonCompositeCriterion_e.swsFOSNonCompositeCriterion_Automatic.
Examples: Visual Basic for Applications

### ICWResults.GetMinMaxFactorOfSafety2
Gets the algebraic minimum and maximum factors of safety (FOS) for non-composite shells.
```vb
Function GetMinMaxFactorOfSafety2( ByVal BAllBodies As System.Boolean, ByVal ArraySelectedEntities As System.Object, ByVal NComponent As System.Integer, ByVal NShellOptions As System.Integer, ByVal NStepNumber As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BAllBodies True to use all bodies to plot the factor of safety, false to use specific bodies ArraySelectedEntities Array of bodies for which to plot the factor of safety; valid only if BAllBodies is set to false NComponent Failure criterion as defined by swsFOS_NonCompositeCriterion_e (see Remarks ) NShellOptions Shell face on which to perform the factor of safety as defined by swsFOS_ShellFaceOption_e NStepNumber Plot step number or 0; 0 gets step 1 for linear static studies and the last available plot step for nonlinear static studies ErrorCode Error code as defined by swsFosPlotErrorCode_e Return Value Array of two doubles of the minimum and maximum factors of safety
Return Value Array of two doubles of the minimum and maximum factors of safety
VBA: See CWResults::GetMinMaxFactorOfSafety2 .
Remarks: If nonlinearities exist in the study, be advised to not use Factor of Safety to estimate the load-bearing capacity of the system. For... Specify NComponent with swsFOS_NonCompositeCriterion_e.... Ductile materials swsFOSNonCompositeCriterion_VonMisesHencky - or - swsFOSNonCompositeCriterion_Tresca Brittle materials swsFOSNonCompositeCriterion_MohrCoulomb - or swsFOSNonCompositeCriterion_Coulomb Pure beam studies swsFOSNonCompositeCriterion_Automatic
Examples: Visual Basic for Applications; Example Get Factor of Safety Values

### ICWResults.GetMinMaxFactorOfSafetyWithDetailSettings
Obsolete. Superseded by ICWResults::GetMinMaxFactorOfSafetyWithDetailSettings2 .
```vb
Function GetMinMaxFactorOfSafetyWithDetailSettings( ByVal BAllBodies As System.Boolean, ByVal ArraySelectedEntities As System.Object, ByVal NComponent As System.Integer, ByVal BUpperLimit As System.Boolean, ByVal DUpperValue As System.Double, ByVal NStressUnit As System.Integer, ByVal NStressLimitOption As System.Integer, ByVal DStressValue As System.Double, ByVal NCompressiveStressLimitOption As System.Integer, ByVal DCompressiveStressValue As System.Double, ByVal DMultiplicationFactor As System.Double, ByVal DCompressiveStressMultiplicationFactor As System.Double, ByVal BCombinedStressOnBeams As System.Boolean, ByVal NShellOptions As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BAllBodies True to select all bodies to plot the factor of safety, false to select specific bodies ArraySelectedEntities Array of bodies for which to plot the factor of safety; valid only if BAllBodies is set to false NComponent Failure criterion as defined by swsFOS_NonCompositeCriterion_e (see Remarks ) BUpperLimit True to set an upper limit, false to not DUpperValue Upper limit; valid only if BUpperLimit is true NStressUnit Units of stress as defined in swsStrengthUnit_e NStressLimitOption Tensile stress limit as defined in swsFactorOfSafetyStressLimitOption_e DStressValue Tensile stress value; valid only if NStressLimitOption is swsFactorOfSafetyStressLimitOption_e.swsFactorOfSafetyStressLimitOption_UserDefined NCompressiveStressLimitOption Compressive stress limit as defined in swsFactorOfSafetyStressLimitOption_e DCompressiveStressValue Compressive stress value; valid only if NCompressiveStressLimitOption is swsFactorOfSafetyStressLimitOption_e.swsFactorOfSafetyStressLimitOption_UserDefined DMultiplicationFactor Tensile stress limit multiplication factor DCompressiveStressMultiplicationFactor Compressive stress limit multiplication factor BCombinedStressOnBeams True to combine stress on beams, false to not NShellOptions Shell face on which to perform the factor of safety as defined by swsFOS_ShellFaceOption_e ErrorCode Error code as defined by swsFosPlotErrorCode_e Return Value Array of two doubles of the minimum and maximum factors of safety
Return Value Array of two doubles of the minimum and maximum factors of safety
VBA: See CWResults::GetMinMaxFactorOfSafetyWithDetailSettings .
Remarks: For pure beam studies, you can only specify NComponent with swsFOS_NonCompositeCriterion_e.swsFOSNonCompositeCriterion_Automatic.
Examples: Visual Basic for Applications

### ICWResults.GetMinMaxFactorOfSafetyWithDetailSettings2
Gets the algebraic minimum and maximum factors of safety (FOS) for non-composite shells and the specified detail settings.
```vb
Function GetMinMaxFactorOfSafetyWithDetailSettings2( ByVal BAllBodies As System.Boolean, ByVal ArraySelectedEntities As System.Object, ByVal NComponent As System.Integer, ByVal BUpperLimit As System.Boolean, ByVal DUpperValue As System.Double, ByVal NStressUnit As System.Integer, ByVal NStressLimitOption As System.Integer, ByVal DStressValue As System.Double, ByVal NCompressiveStressLimitOption As System.Integer, ByVal DCompressiveStressValue As System.Double, ByVal DMultiplicationFactor As System.Double, ByVal DCompressiveStressMultiplicationFactor As System.Double, ByVal BCombinedStressOnBeams As System.Boolean, ByVal NShellOptions As System.Integer, ByVal NStepNumber As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BAllBodies True to use all bodies to plot the factor of safety, false to use specific bodies ArraySelectedEntities Array of bodies for which to plot the factor of safety; valid only if BAllBodies is set to false NComponent Failure criterion as defined by swsFOS_NonCompositeCriterion_e (see Remarks ) BUpperLimit True to set an upper limit, false to not DUpperValue Upper limit; valid only if BUpperLimit is true NStressUnit Units of stress as defined in swsStrengthUnit_e NStressLimitOption Limit stress types as defined in swsFactorOfSafetyStressLimitOption_e (see Remarks ) DStressValue Stress value; valid only if NStressLimitOption is swsFactorOfSafetyStressLimitOption_e.swsFactorOfSafetyStressLimitOption_UserDefined NCompressiveStressLimitOption Compressive stress limit as defined in swsFactorOfSafetyStressLimitOption_e (see Remarks ) DCompressiveStressValue Compressive stress value; valid only if NCompressiveStressLimitOption is swsFactorOfSafetyStressLimitOption_e.swsFactorOfSafetyStressLimitOption_UserDefined DMultiplicationFactor Stress limit multiplication factor DCompressiveStressMultiplicationFactor Compressive stress limit multiplication factor BCombinedStressOnBeams True to combine stress on beams, false to not NShellOptions Shell face on which to perform the factor of safety as defined by swsFOS_ShellFaceOption_e NStepNumber Plot step number or 0; 0 gets step 1 for linear static studies and the last available plot step for nonlinear static studies ErrorCode Error code as defined by swsFosPlotErrorCode_e Return Value Array of two doubles of the minimum and maximum factors of safety
Return Value Array of two doubles of the minimum and maximum factors of safety
VBA: See CWResults::GetMinMaxFactorOfSafetyWithDetailSettings2 .
Remarks: If nonlinearities exist in the study, be advised to not use Factor of Safety to estimate the load-bearing capacity of the system. For... Specify NComponent with swsFOS_NonCompositeCriterion_e.... Ductile materials swsFOSNonCompositeCriterion_VonMisesHencky - or - swsFOSNonCompositeCriterion_Tresca Brittle materials swsFOSNonCompositeCriterion_MohrCoulomb - or swsFOSNonCompositeCriterion_Coulomb Pure beam studies swsFOSNonCompositeCriterion_Automatic How you set stress limits using NStressLimitOption and NCompressiveStressLimitOption depends on the model material . Specify NStressLimitOption and NCompressiveStressLimitOption with swsFactorOfSafetyStressLimitOption_e... For swsMaterialModelType_e ... swsFactorOfSafetyStressLimitOption_YieldStrength swsMaterialModelTypeLinearElasticIsotropic swsMaterialModelTypeLinearElasticOrthtropic swsMaterialModelTypeElastoPlasticvonMisesKinematic swsMaterialModelTypeElastoPlasticTrescaKinematic swsFactorOfSafetyStressLimitOption_UltimateStrength swsMaterialModelTypeLinearElasticIsotropic swsMaterialModelTypeLinearElasticOrthtropic swsMaterialModelTypeNonlinearElastic swsMaterialModelTypeElastoPlasticvonMmisesKinematic swsMaterialModelTypeElastoPlasticTrescaKinematic swsMaterialModelTypeElastoPlasticDruckerPrager swsMaterialModelTypeHyperElasticBlatzko swsMaterialModelTypeHyperElasticMooneyRivlin swsMaterialModelTypeHyperElasticOgden swsMaterialModelTypeViscoElastic swsMaterialModelTypeNitinol
Examples: Visual Basic for Applications; Example Get Factor of Safety Values

### ICWResults.GetMinMaxFatigue
Gets the algebraic minimum and maximum for the specified fatigue component.
```vb
Function GetMinMaxFatigue( ByVal NComponent As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Type of fatigue as defined in swsFatigueComponent_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxFatigue .
Remarks: This method returns the following array: { node_with_minimum_fatigue , minimum_fatigue , node_with_maximum_fatigue , maximum_fatigue }, where the nodes are integers, and the fatigue values are floats.
Examples: Visual Basic for Applications; Create Fatigue Study

### ICWResults.GetMinMaxResultsEquationValues
Gets the algebraic minimum and maximum from the plot of the specified results equation.
```vb
Function GetMinMaxResultsEquationValues( ByVal SResultsEquation As System.String, ByVal BValueByElem As System.Boolean, ByVal NUnits As System.Integer, ByVal NStepNumber As System.Integer, ByVal NShellOptions As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters SResultsEquation Results equation BValueByElem True to plot element values, false to plot node values NUnits Units as defined in swsUnit_e NStepNumber Number of solution step to plot NShellOptions Shell face as defined in swsShellFace_e ; valid only for shell and mixed mesh models ErrorCode Error code as defined in swsResultPlotErrorCode_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxResultsEquationValues .
Remarks: This method returns the following array: { node_or_element_with_minimum_value , minimum_value , node_or_element_with_maximum_value , maximum_value }, where the nodes/elements are integer indexes, and the values are in scientific notation.
Examples: Visual Basic for Applications; Example Create Results Equation Plot

### ICWResults.GetMinMaxStrain
Gets the algebraic minimum and maximum for the specified strain component, element, and solution step.
```vb
Function GetMinMaxStrain( ByVal NComponent As System.Integer, ByVal NElementNumber As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Strain component as defined in swsStrainComponent_e NElementNumber Element number NStepNumber Solution step number (use 1 for static) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxStrain .
Remarks: This method returns the following array: { node_with_minimum_strain , minimum_strain , node_with_maximum_strain , maximum_strain }, where the nodes are integer indexes, and the strains are in scientific notation.
Examples: Visual Basic for Applications

### ICWResults.GetMinMaxStress
Gets the algebraic minimum and maximum for the specified stress component, element, and solution step.
```vb
Function GetMinMaxStress( ByVal NComponent As System.Integer, ByVal NElementNumber As System.Integer, ByVal NStepNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Stress component as defined in swsStressComponent_e NElementNumber Element number NStepNum Solution step number (use 1 for static) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist NUnits Units as defined in swsStrengthUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxStress .
Remarks: This method returns the following array: { node_with_minimum_stress , minimum_stress , node_with_maximum_stress , maximum_stress }, where the nodes are integer indexes, and the stresses are in scientific notation. For random vibration studies, this method returns power spectral density (PSD) values. To obtain root-mean-square (RMS) values in random vibration studies, call ICWResults::GetMinMaxStressRMS .
Examples: Visual Basic for Applications; Analyze Part; Create Drop Test Study

### ICWResults.GetMinMaxStressForHarmonic
Obsolete. Superseded by ICWResults::GetMinMaxStressForHarmonic2 .
```vb
Function GetMinMaxStressForHarmonic( ByVal NComponent As System.Integer, ByVal BByElementNumber As System.Integer, ByVal BAvgOnBoundary As System.Integer, ByVal NStepNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Stress component as defined in swsStressComponent_e BByElementNumber 1 to obtain element stress values, 0 to obtain nodal stress values BAvgOnBoundary 1 to activate averaging of nodal stress results across common part boundaries, 0 to deactivate averaging; valid only if BByElementNumber = 0 NStepNum Solution step number (use 1 for static) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist NUnits Units as defined in swsStrengthUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxStressForHarmonic .
Remarks: This method returns the following array: { node/element_with_minimum_stress , minimum_stress , node/element_with_maximum_stress , maximum_stress }, where the nodes or elements are integer indexes, and the stresses are in scientific notation.
Examples: Visual Basic for Applications

### ICWResults.GetMinMaxStressForHarmonic2
Gets the algebraic minimum and maximum for the specified stress component and solution step of this harmonic study.
```vb
Function GetMinMaxStressForHarmonic2( ByVal NComponent As System.Integer, ByVal BByElementNumber As System.Boolean, ByVal BAvgOnBoundary As System.Boolean, ByVal NStepNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Stress component as defined in swsStressComponent_e BByElementNumber -1 or true to obtain element stress values, 0 or false to obtain nodal stress values BAvgOnBoundary -1 or true to activate averaging of nodal stress results across common part boundaries, 0 or false to deactivate averaging; valid only if BByElementNumber = 0 NStepNum Solution step number (use 1 for static) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist NUnits Units as defined in swsStrengthUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
Remarks: This method returns the following array: { node/element_with_minimum_stress , minimum_stress , node/element_with_maximum_stress , maximum_stress }, where the nodes or elements are integer indexes, and the stresses are in scientific notation.

### ICWResults.GetMinMaxStressRMS
Gets the algebraic minimum and maximum RMS values for the specified stress component and element of this random vibration study.
```vb
Function GetMinMaxStressRMS( ByVal NComponent As System.Integer, ByVal NElementNumber As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Stress component as defined in swsStressComponent_e NElementNumber Element number DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist NUnits Units as defined in swsStrengthUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxStressRMS .
Remarks: This method returns the following array: { node_with_minimum_stress , minimum_stress , node_with_maximum_stress , maximum_stress }, where the nodes are integer indexes, and the stresses are in scientific notation. This method returns root-mean-square (RMS) values. To obtain power spectral density (PSD) values, call ICWResults::GetMinMaxStress .
Examples: Visual Basic for Applications

### ICWResults.GetMinMaxThermal
Gets the algebraic minimum and maximum for the specified thermal component and solution step.
```vb
Function GetMinMaxThermal( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Thermal component as defined in swsThermalComponent_e NStepNumber Solution step number (use 1 for steady state) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist NUnits Temperature unit as defined swsTemperatureUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxThermal .
Remarks: This method returns the following array: { node_with_minimum_temp , minimum_temp , node_with_maximum_temp , maximum_temp }, where the nodes are integer indexes, and the temperatures are in scientific notation.
Examples: Visual Basic for Applications

### ICWResults.GetMinMaxVelocity
Gets the algebraic minimum and maximum for the specified velocity component and solution step.
```vb
Function GetMinMaxVelocity( ByVal NComponent As System.Integer, ByVal NStepNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Velocity component as defined in swsVelocityComponent_e NStepNum Solution step number (use 1 for steady state) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist NUnits Units of velocity as defined in swsVelocityUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetMinMaxVelocity .
Remarks: This method returns the following array: { node_with_minimum_velocity , minimum_velocity , node_with_maximum_velocity , maximum_velocity }, where the nodes are integer indexes, and the velocities are in scientific notation.
Examples: Visual Basic for Applications; Example Create Linear Dynamic Study

### ICWResults.GetPlot
Gets the specified results plot.
```vb
Function GetPlot( ByVal SPlotName As System.String, ByRef ErrorCode As System.Integer ) As CWPlot
```
Parameters SPlotName Name of results plot (see Remarks ) ErrorCode Error code as defined by swsResultPlotErrorCode_e Return Value ICWPlot
Return Value ICWPlot
VBA: See CWResults::GetPlot .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName.
Examples: Visual Basic for Applications; Example Create a Design Insight Plot for a Static Study

### ICWResults.GetPlotColorOptions
Gets the color options for the specified plot.
```vb
Function GetPlotColorOptions( ByVal SPlotName As System.String, ByRef NType As System.Integer, ByRef NBaseColor As System.Integer, ByRef NContour As System.Integer, ByRef BFlip As System.Boolean, ByRef BSpecifyColorLimit As System.Boolean, ByRef VarColor As System.Object ) As System.Integer
```
Parameters SPlotName Plot name (see Remarks ) NType Color map as defined in swsColorChartOptionLegendTypeValue_e (see Remarks ) NBaseColor Number of base colors in color map specified by NType (see Remarks ) NContour Number of gradient colors BFlip True to reverse the color mapping, false to not BSpecifyColorLimit True to specify a color for values above the yield limit, false to not; valid only for von Mises stress plots of single body parts where the yield strength is defined VarColor Array of RGB triplets: RGB color for values above the yield limit; valid only if BSpecifyColorLimit is true One or more RGB triplets of user-defined colors Return Value Error as defined in swsResultsError_e
Return Value Error as defined in swsResultsError_e
VBA: See CWResults::GetPlotColorOptions .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. If NType is swsColorChartOptionLegendTypeValue_e... Then NBaseColor is... swsColorChartOptionLegendDefault 5 swsColorChartOptionLegendRainbow 7 swsColorChartOptionLegendGrayScale 2 swsColorChartOptionLegendUserDefined 2 - 9
Examples: Visual Basic for Applications; Example Analyze Part

### ICWResults.GetPlotCount
Gets the number of plots in this study.
```vb
Function GetPlotCount() As System.Integer
```
Return Value Number of plots
VBA: See CWResults::GetPlotCount .
Examples: Visual Basic for Applications; Example Analyze Part

### ICWResults.GetPlotDefinition
Gets the definition of the specified plot.
```vb
Function GetPlotDefinition( ByVal SPlotName As System.String, ByRef NPlotType As System.Integer, ByRef NComponentName As System.String, ByRef BNodal As System.Boolean, ByRef BDeformed As System.Boolean, ByRef DScaleFactor As System.Double ) As System.Integer
```
Parameters SPlotName Name of plot NPlotType Type of plot as defined in swsPlotResultTypes_e NComponentName Name of plotted component BNodal True if plot is nodal, false if elemental BDeformed True if deformed, false if not DScaleFactor Scale factor for deformation Return Value Error as defined in swsResultsError_e
Return Value Error as defined in swsResultsError_e
VBA: See CWResults::GetPlotDefinition .
Examples: Visual Basic for Applications; Example Create Body From Deformed Shape

### ICWResults.GetPlotDisplayOptions
Gets the display options for the specified plot.
```vb
Function GetPlotDisplayOptions( ByVal SPlotName As System.String, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters SPlotName Plot name (see Remarks ) ErrorCode Error as defined in swsResultsError_e Return Value Array of display options (see Remarks )
Return Value Array of display options (see Remarks )
VBA: See CWResults::GetPlotDisplayOptions .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. The 0-based array contains eight elements: Array element... Contains... 0 Boolean; whether to display minimum value 1 Boolean; whether to display maximum value 2 Boolean; whether to display plot details 3 Boolean; whether to display plot legend 4 Boolean; whether to display minimum/maximum ranges only for parts 5 Boolean; whether to allow user-defined minimum/maximum values 6 Double; user-defined minimum value; valid only if array element 5 is true 7 Double; user-defined maximum value; valid only if array element 5 is true
Examples: Visual Basic for Applications; Example Analyze Part

### ICWResults.GetPlotNames
Gets the names of the plots for these results.
```vb
Function GetPlotNames() As System.Object
```
Return Value Array of the names of the plots for these results
VBA: See CWResults::GetPlotNames .
Examples: Visual Basic for Applications; Example Analyze Part

### ICWResults.GetPlotPositionFormatOptions
Gets the position format options for the specified plot.
```vb
Function GetPlotPositionFormatOptions( ByVal SPlotName As System.String, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters SPlotName Plot name (see Remarks ) ErrorCode Error as defined in swsResultsError_e Return Value Array of plot position/format options (see Remarks )
Return Value Array of plot position/format options (see Remarks )
VBA: See CWResults::GetPlotPositionFormatOptions .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. The 0-based array contains eight elements: Array element... Contains... 0 Double[0,1]; horizontal position from the left of the graphics area; percentage of the window width 1 Double[0,1]; vertical position from the top of the graphics area; percentage of the window height 2 Chart width option as defined in swsColorChartWidthOptionValue_e 3 Chart number format as defined in swsColorChartNumberFormatOptionValue_e 4 0 <= Number of decimal places to display <= 16 5 Boolean; whether to display chart numbers with a 1000 comma separator; valid only if array element 3 is swsColorChartNumberFormatOptionValue_e.swsColorChartNumberFormatFloating or swsColorChartNumberFormatOptionValue_e.swsColorChartNumberFormatGeneral 6 Boolean; whether to use different number format for small numbers (0.001 < |x| < 1000); valid only if array element 3 is swsColorChartNumberFormatOptionValue_e.swsColorChartNumberFormatScientific 7 Number format for small numbers as defined in swsColorNumberFormatUseDiffNumberFormatOptionValue_e ; valid only if array element 6 is true
Examples: Visual Basic for Applications; Example Analyze Part

### ICWResults.GetPlotSettings
Gets the settings for the specified plot.
```vb
Function GetPlotSettings( ByVal SPlotName As System.String, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters SPlotName Plot name (see Remarks ) ErrorCode Error as defined in swsResultsError_e Return Value Array of plot settings (see Remarks )
Return Value Array of plot settings (see Remarks )
VBA: See CWResults::GetPlotSettings .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. The 0-based array contains 11 elements: Array element... Contains... 0 Display option for active fringe plot as defined in swsPlotFringeSettingsOptionValue_e 1 Display option for model boundary as defined in swsPlotBoundarySettingsOptionValue_e 2 R[0-255] value for model/mesh color; valid only if array element 1 is not swsPlotBoundarySettingsOptionValue_e.swsPlotBoundaryNone 3 G[0-255] value for model/mesh color; valid only if array element 1 is not swsPlotBoundarySettingsOptionValue_e.swsPlotBoundaryNone 4 B[0-255] value for model/mesh color; valid only if array element 1 is not swsPlotBoundarySettingsOptionValue_e.swsPlotBoundaryNone 5 Boolean; whether to superimpose the undeformed model on the deformed model 6 Deformed plot translucent color option as defined in swsPlotDeformedShapeOptionSuperImposeValue_e ; valid only if array element 5 is true 7 R[0-255] value for single translucent color; valid only if array element 5 is true, and array element 6 is swsPlotDeformedShapeOptionSuperImposeValue_e.swsPlotDeformedShapeSuperImposeModel_TranslucentSingleColor 8 G[0-255] value for single translucent color; valid only if array element 5 is true, and array element 6 is swsPlotDeformedShapeOptionSuperImposeValue_e.swsPlotDeformedShapeSuperImposeModel_TranslucentSingleColor 9 B[0-255] value for single translucent color; valid only if array element 5 is true, and array element 6is swsPlotDeformedShapeOptionSuperImposeValue_e.swsPlotDeformedShapeSuperImposeModel_TranslucentSingleColor 10 0.0 <= Transparency or intensity of the translucent single color or part colors <= 1.0; valid only if array element 5 is true
Examples: Visual Basic for Applications; Example Analyze Part

### ICWResults.GetPlotSettingsOptionForHiddenAndExcludedBody
Gets hidden and excluded body options for the specified plot.
```vb
Function GetPlotSettingsOptionForHiddenAndExcludedBody( ByVal SPlotName As System.String, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters SPlotName Plot name (see Remarks ) ErrorCode Error code as defined in swsPlotSettingsErrorCode_e Return Value Array of hidden and excluded body options (see Remarks )
Return Value Array of hidden and excluded body options (see Remarks )
VBA: See CWResults::GetPlotSettingsOptionForHiddenAndExcludedBody .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. The 0-based array contains 12 elements: Array element... Contains... 0 Boolean; whether to show hidden bodies 1 Hidden body translucent color option as defined in swsPlotShowHiddenBodiesOptionValue_e ; valid only if array element 0 is true 2 R[0-255] value for single translucent color; valid only if array element 0 is true, and array element 1 is swsPlotShowHiddenBodiesOptionValue_e.swsPlotHiddenBodyTranslucentSingleColor 3 G[0-255] value for single translucent color; valid only if array element 0 is true, and array element 1 is swsPlotShowHiddenBodiesOptionValue_e.swsPlotHiddenBodyTranslucentSingleColor 4 B[0-255] value for single translucent color; valid only if array element 0 is true, and array element 1 is swsPlotShowHiddenBodiesOptionValue_e.swsPlotHiddenBodyTranslucentSingleColor 5 0.0 <= transparency or intensity of the translucent single color or part colors <= 1.0; valid only if array element 0 is true 6 Boolean; whether to show excluded bodies 7 Excluded body translucent color option as defined in swsPlotShowExcludedBodiesOptionValue_e ; valid only if array element 6 is true 8 R[0-255] value for single translucent color; valid only if array element 6 is true, and array element 7 is swsPlotShowExcludedBodiesOptionValue_e.swsPlotExcludedBodyTranslucentSingleColor 9 G[0-255] value for single translucent color; valid only if array element 6 is true, and array element 7 is swsPlotShowExcludedBodiesOptionValue_e.swsPlotExcludedBodyTranslucentSingleColor 10 B[0-255] value for single translucent color; valid only if array element 6 is true, and array element 7 is swsPlotShowExcludedBodiesOptionValue_e.swsPlotExcludedBodyTranslucentSingleColor 11 0.0 <= transparency or intensity of the translucent single color or part colors <= 1.0; valid only if array element 6 is true
Examples: Visual Basic for Applications; Example Analyze Part

### ICWResults.GetReactionForcesAndMoments
Obsolete. Superseded by ICWResults::GetReactionForcesAndMomentsWithSelections .
```vb
Function GetReactionForcesAndMoments( ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NStepNumber Solution step number (use 1 for static) DispPlane Reference geometry NUnits Unit as defined in swsForceUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of reaction forces and moments for the entire model (see Remarks )
Return Value Array of reaction forces and moments for the entire model (see Remarks )
VBA: See CWResults::GetReactionForcesAndMoments .
Remarks: This method is valid only for static, nonlinear, drop test, and dynamic studies. This method returns the reaction forces and moments at each node of the entire model: { node_1 , node_1_xcoord_reaction_force, node_1_ycoord_reaction_force, node_1_zcoord_reaction_force, node_1_resultant_reaction_force, node_1_xcoord_reaction_moment, node_1_ycoord_reaction_moment, node_1_zcoord_reaction_moment, node_1_resultant_reaction_moment, ... node_n, node_n_xcoord_reaction_force, node_n_ycoord_reaction_force, node_n_zcoord_reaction_force, node_n_resultant_reaction_force, node_n_xcoord_reaction_moment, node_n_ycoord_reaction_moment, node_n_zcoord_reaction_moment, node_n_resultant_reaction_moment }, where the nodes are integers, and the reaction forces and moments are in decimal or scientific notation.
Examples: Visual Basic for Applications

### ICWResults.GetReactionForcesAndMomentsWithSelections
Gets the reaction forces and moments for selections and the entire model at the specified solution step.
```vb
Function GetReactionForcesAndMomentsWithSelections( ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByVal SelectedObjects As System.Object, ByRef SelectionAndEntireModelReactionForcesAndMoments As System.Object, ByRef EachSelectedObjectReactionForcesAndMoments As System.Object, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NStepNumber Solution step number (use 1 for static) DispPlane Plane, axis, or coodinate system relative to which to list reaction results NUnits Unit as defined in swsForceUnit_e SelectedObjects Array of entities for which to get reaction results SelectionAndEntireModelReactionForcesAndMoments Array of reaction forces and moments for the selected entities and the entire model (see Remarks ) EachSelectedObjectReactionForcesAndMoments Array of reaction forces and moments for each selected entity (see Remarks ) ErrorCode Error as defined in swsResultsError_e Return Value Array of reaction forces and moments (see Remarks )
Return Value Array of reaction forces and moments (see Remarks )
VBA: See CWResults::GetReactionForcesAndMomentsWithSelections .
Remarks: This method is valid only for static, nonlinear, drop test, and dynamic studies. SelectionAndEntireModelReactionForcesAndMoments array contains 16 elements: { summation_xcoord_forces_all_selections, summation_ycoord_forces_all_selections, summation_zcoord_forces_all_selections, resultant_force_all_selections , summation_xcoord_moments_all_selections, summation_ycoord_moments_all_selections, summation_zcoord_moments_all_selections, resultant_moment_all_selections , summation_xcoord_forces_entire_model, summation_ycoord_forces_entire_model, summation_zcoord_forces_entire_model, resultant_force_entire_model , summation_xcoord_moments_entire_model, summation_ycoord_moments_entire_model, summation_zcoord_moments_entire_model, resultant_moment_entire_model } EachSelectedObjectReactionForcesAndMoments array contains (8 X no_of_selections ) elements: { selection_1_xcoord_reaction_force, selection_1_ycoord_reaction_force, selection_1_zcoord_reaction_force, selection_1_resultant_reaction_force , selection_1_xcoord_reaction_moment, selection_1_ycoord_reaction_moment, selection_1_zcoord_reaction_moment, selection_1_resultant_reaction_moment, ... selection_n_xcoord_reaction_force, selection_n_ycoord_reaction_force, selection_n_zcoord_reaction_force, selection_n_resultant_reaction_force , selection_n_xcoord_reaction_moment, selection_n_ycoord_reaction_moment, selection_n_zcoord_reaction_moment, selection_n_resultant_reaction_moment } This method returns the reaction forces and moments at each node of the entire model: { node_1 , node_1_xcoord_reaction_force, node_1_ycoord_reaction_force, node_1_zcoord_reaction_force, node_1_resultant_reaction_force , node_1_xcoord_reaction_moment, node_1_ycoord_reaction_moment, node_1_zcoord_reaction_moment, node_1_resultant_reaction_moment , ... node_n , node_n_xcoord_reaction_force, node_n_ycoord_reaction_force, node_n_zcoord_reaction_force, node_n_resultant_reaction_force , node_n_xcoord_reaction_moment, node_n_ycoord_reaction_moment, node_n_zcoord_reaction_moment, node_n_resultant_reaction_moment } In all of the arrays, the nodes are integers, and the reaction forces and moments are in decimal or scientific notation.
Examples: Visual Basic for Applications; Example Create Linear Dynamic Study

### ICWResults.GetRemoteForces
Gets the x-, y-, and z-component and resultant forces applied to selected entities as a result of transferring a remote load in static studies.
```vb
Function GetRemoteForces( ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NStepNumber Solution step number (use 1 for static) DispPlane Plane, axis, or coordinate system relative to which the results are listed ArraySelectedEntities Selected faces, edges (shells only), vertices, and components NUnits Output units as defined in swsUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of four doubles of the remote load forces for all selections in ArraySelectedEntities: Summation of x-components of remote load forces for all selections Summation of y-components of remote load forces for all selections Summation of z-components of remote load forces for all selections Resultant remote load force for all selections
Return Value Array of four doubles of the remote load forces for all selections in ArraySelectedEntities: Summation of x-components of remote load forces for all selections Summation of y-components of remote load forces for all selections Summation of z-components of remote load forces for all selections Resultant remote load force for all selections
VBA: See CWResults::GetRemoteForces .
Examples: Visual Basic for Applications; Example Get Remote Load Forces

### ICWResults.GetResonantFrequencies
Gets the resonant frequency values.
```vb
Function GetResonantFrequencies( ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters ErrorCode Error as defined in swsResultsError_e Return Value Array of resonant frequencies containing groups of these results: Mode number Resonant frequency (rad/sec) Resonant frequency (Hertz) Resonant period (seconds)
Return Value Array of resonant frequencies containing groups of these results: Mode number Resonant frequency (rad/sec) Resonant frequency (Hertz) Resonant period (seconds)
VBA: See CWResults::GetResonantFrequencies .
Examples: Visual Basic for Applications; Create Frequency Study with Solid Mesh

### ICWResults.GetResultsEquationValuesForEntities
Gets the result values for the specified equation and selected entities.
```vb
Function GetResultsEquationValuesForEntities( ByVal SResultsEquation As System.String, ByVal BValueByElem As System.Boolean, ByVal NUnits As System.Integer, ByVal NStepNumber As System.Integer, ByVal NShellOptions As System.Integer, ByVal ArraySelectedEntities As System.Object, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters SResultsEquation Results equation BValueByElem True to plot element values, false to plot node values NUnits Units as defined in swsUnit_e NStepNumber Number of solution step to plot NShellOptions Shell face as defined in swsShellFace_e ; valid only for shell and mixed mesh models ArraySelectedEntities Array of selected entities; Nothing or null to return values for the entire model ErrorCode Error code as defined in swsResultPlotErrorCode_e Return Value Array (see Remarks )
Return Value Array (see Remarks )
VBA: See CWResults::GetResultsEquationValuesForEntities .
Remarks: This method returns the following array: { node_or_element_1 , value_1, ... node_or_element_n, value_n }, where the nodes/elements are integer indexes, and the values are in scientific notation.
Examples: Visual Basic for Applications; Example Create Results Equation Plot

### ICWResults.GetRotationalDisplacement
Gets the rotational displacements at the specified solution step.
```vb
Function GetRotationalDisplacement( ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NStepNumber Solution step number (use 1 for static) DispPlane Reference geometry NUnits Unit as defined in swsResultsRotationalDisplacementUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of rotational displacements
Return Value Array of rotational displacements
VBA: See CWResults::GetRotationalDisplacement .
Remarks: Rotational degrees of freedom are available for shells and beams.
Examples: Visual Basic for Applications

### ICWResults.GetStrain
Gets the strain results for all nodes or elements at the specified solution step.
```vb
Function GetStrain( ByVal NElementNumber As System.Integer, ByVal NStepNum As System.Integer, ByVal DispPlane As System.Object, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NElementNumber Element number: 0 = Nodal 1 = Elemental NStepNum Solution step number (use 1 for static) DispPlane Reference geometry ErrorCode Error as defined in swsResultsError_e Return Value Array of strain results
Return Value Array of strain results
VBA: See CWResults::GetStrain .
Remarks: EPSx, EPSy, EPSz, GMxy, GMyz, GMxz, ESTRN, SEDENS, ENERGY, E1, E2, and E3 for all nodes or elements at NStepNum.
Examples: Visual Basic for Applications

### ICWResults.GetStrainComponentForAllStepsAtNode
Gets the specified strain component at the specified node for all solution steps.
```vb
Function GetStrainComponentForAllStepsAtNode( ByVal NComponent As System.Integer, ByVal NNodeNum As System.Integer, ByVal DispPlane As System.Object, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Strain component as defined in swsStrainComponent_e NNodeNum Node number DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if reference geometry does not exist ErrorCode Error as defined in swsResultsError_e Return Value Array of strain results
Return Value Array of strain results
VBA: See CWResults::GetStrainComponentForAllStepsAtNode .
Examples: Visual Basic for Applications

### ICWResults.GetStrainForEntities
Obsolete. Superseded by ICWResults::GetStrainForEntities2 .
```vb
Function GetStrainForEntities( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Strain as defined in swsStrainComponent_e NStepNumber Solution step number (use 1 for static) DispPlane Reference geometry ArraySelectedEntities Array of geometric entities ErrorCode Error as defined in swsResultsError_e Return Value Array of strain results
Return Value Array of strain results
VBA: See CWResults::GetStrainForEntities .
Examples: Visual Basic for Applications

### ICWResults.GetStrainForEntities2
Obsolete. Superseded by ICWResults::GetStrainForEntities3 .
```vb
Function GetStrainForEntities2( ByVal BValueByNode As System.Boolean, ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BValueByNode True to get the strain value by node, false to get the strain value by element NComponent Strain as defined in swsStrainComponent_e NStepNumber Solution step number (use 1 for static) DispPlane Reference geometry ArraySelectedEntities Array of geometric entities ErrorCode Error as defined in swsResultsError_e Return Value Array of strain results (each node or element followed by the strain value)
Return Value Array of strain results (each node or element followed by the strain value)
VBA: See CWResults::GetStrainForEntities2 .
Examples: Visual Basic for Applications

### ICWResults.GetStrainForEntities3
Gets the specified strain component for the specified entities at the specified solution step.
```vb
Function GetStrainForEntities3( ByVal BValueByNode As System.Boolean, ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NShellFace As System.Integer, ByVal NPlyNumber As System.Integer, ByVal BInPlyDirection As System.Boolean, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BValueByNode True to get the strain value by node, false to get the strain value by element NComponent Strain component as defined in swsStrainComponent_e NStepNumber Solution step number (use 1 for static) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NShellFace Option as defined in swsShellFace_e NPlyNumber Ply number to show the strain; available only when the model uses composite shells BInPlyDirection 1 to display the results in ply direction on the surface of the ply, 0 to display the results in the transverse direction to the ply direction on the surface of the ply ErrorCode Error as defined in swsResultsError_e Return Value Array of {node/element #, strain value} pairs
Return Value Array of {node/element #, strain value} pairs
VBA: See CWResults::GetStrainForEntities3 .
Examples: Visual Basic for Applications; Example Analyze Part

### ICWResults.GetStress
Gets the stress results for all nodes or elements at the specified solution step.
```vb
Function GetStress( ByVal NElementNumber As System.Integer, ByVal NStepNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NElementNumber Element number: 0 = Nodal 1 = Elemental NStepNum Solution step number (use 1 for static) DispPlane Reference geometry NUnits Unit as defined in swsStrengthUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of stress results
Return Value Array of stress results
VBA: See CWResults::GetStress .
Remarks: Sx, Sy, Sz, Txy, Tyz, Txz, P1, P2, P3, VON, and INT for all nodes or elements at the specified step number.
Examples: Visual Basic for Applications

### ICWResults.GetStressComponentForAllStepsAtNode
Gets the specified stress component at the specified node for all solution steps.
```vb
Function GetStressComponentForAllStepsAtNode( ByVal NComponent As System.Integer, ByVal NNodeNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Stress component as defined in swsStressComponent_e NNodeNum Node number DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component NUnits Unit as defined in swsStrengthUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of stress results
Return Value Array of stress results
VBA: See CWResults::GetStressComponentForAllStepsAtNode .
Remarks: Sx, Sy, Sz, Txy, Tyz, Txz, P1, P2, P3, VON, and INT for all nodes or elements at the specified step number.
Examples: Visual Basic for Applications

### ICWResults.GetStressForEntities
Obsolete. Superseded by ICWResults::GetStressForEntities2 .
```vb
Function GetStressForEntities( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArrayArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Stress as defined in swsStressComponent_e NStepNumber Solution step number (use 1 for steady state) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArrayArraySelectedEntities Array of geometric entities NUnits Unit as defined in swsStrengthUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of stress results
Return Value Array of stress results
VBA: See CWResults::GetStressForEntities .
Examples: Visual Basic for Applications

### ICWResults.GetStressForEntities2
Obsolete. Superseded by ICWResults::GetStressForEntities3 .
```vb
Function GetStressForEntities2( ByVal BValueByNode As System.Boolean, ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BValueByNode True to get the stress value by node, false to get the stress value by element NComponent Stress as defined in swsStressComponent_e NStepNumber Solution step number (use 1 for steady state) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NUnits Unit as defined in swsStrengthUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of stress results (each node or element followed by the stress value)
Return Value Array of stress results (each node or element followed by the stress value)
VBA: See CWResults::GetStressForEntities2 .
Examples: Visual Basic for Applications

### ICWResults.GetStressForEntities3
Gets the specified stress component for the specified entities at the specified solution step.
```vb
Function GetStressForEntities3( ByVal BValueByNode As System.Boolean, ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByVal NShellFace As System.Integer, ByVal NPlyNumber As System.Integer, ByVal BInPlyDirection As System.Boolean, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters BValueByNode True to get the stress value by node, false to get the stress value by element NComponent Stress component as defined in swsStressComponent_e NStepNumber Solution step number (use 1 for steady state) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NUnits Unit as defined in swsStrengthUnit_e NShellFace Option as defined in swsShellFace_e NPlyNumber Ply number to show the stress; available only when the model uses composite shells BInPlyDirection 1 to display the results in ply direction on the surface of the ply, 0 to display the results in the transverse direction to the ply direction on the surface of the ply ErrorCode Error as defined in swsResultsError_e Return Value Array of {node/element #, stress value} pairs
Return Value Array of {node/element #, stress value} pairs
VBA: See CWResults::GetStressForEntities3 .
Examples: Visual Basic for Applications; Example Analyze Part

### ICWResults.GetStressTensorValuesForAllNodesOfElement
Gets the stress tensor values for all nodes of the specified mesh elements at the specified solution step.
```vb
Function GetStressTensorValuesForAllNodesOfElement( ByVal ElementNumbers As System.Object, ByVal NUnits As System.Integer, ByVal NStepNum As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters ElementNumbers An array of mesh element numbers for which to return nodal results; specify the element number of a single element (see Remarks) NUnits Units as defined in swsStrengthUnit_e NStepNum Solution step number (use 1 for steady state) ErrorCode Error as defined in swsNodalResultsOfElementError_e Return Value Two-dimensional array of tensor stress values (see Remarks )
Return Value Two-dimensional array of tensor stress values (see Remarks )
VBA: See CWResults::GetStressTensorValuesForAllNodesOfElement .
Remarks: This method is valid only for solid and shell meshes. Beam meshes are not supported. Call ICWMesh::ElementCount to specify ElementNumbers. The number of rows in the returned two-dimensional array is: ( n * m * p ) where: n = # elements in the mesh m = # nodes for each element p = 2 for shell meshes (top and bottom faces); 1 for solid meshes Each row contains nine values: element_i, shell_face_i , node_i , tensor_value_1 , tensor_value_2 , tensor_value_3 , tensor_value_4 , tensor_value_5 , tensor_value_6 where: element_i = the zero-based index of the element in the mesh shell_face_i = 0 for the top face of a shell mesh, 1 for the bottom face of a shell mesh, and -1 for a solid mesh node_i = the zero-based index of the node for this row's element Therefore, the array returned for a shell mesh contains 18* n * m values in the following order: [ element_1 , shell_face_0, node_1 , tensor_value_1 , tensor_value_2 , tensor_value_3 , tensor_value_4 , tensor_value_5 , tensor_value_6 element_1 , shell_face_0, node_2 , tensor_value_1 , tensor_value_2 , tensor_value_3 , tensor_value_4 , tensor_value_5 , tensor_value_6 ... element_1 , shell_face_0 , node_m , tensor_value_1 , tensor_value_2 , tensor_value_3 , tensor_value_4 , tensor_value_5 , tensor_value_6 (repeat above for shell_face_1 ) (repeat above for element_2 through element_n ) ]
Examples: Visual Basic for Applications; Example Get Stress Tensor Values For Mesh

### ICWResults.GetThermalComponentForAllStepsAtNode
Gets thermal values at the specified node for all solutions steps.
```vb
Function GetThermalComponentForAllStepsAtNode( ByVal NComponent As System.Integer, ByVal NNodeNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Thermal component as defined in swsThermalComponent_e NNodeNum Node number DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is a scalar component and not a directional component NUnits Temperature unit as defined swsTemperatureUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of thermal results (see Remarks )
Return Value Array of thermal results (see Remarks )
VBA: See CWResults::GetThermalComponentForAllStepsAtNode .
Remarks: This method returns the following array : { solution_step_1 , time_1 , temp_1, solution_step_2 , time_2 , temp_2, solution_step_3 , time_3 , temp_3, ..., solution_step_n , time_n , temp_n }, where the solution steps are integers, time_n = ( solution_step_n * 60) seconds, and the temperatures are doubles.
Examples: Visual Basic for Applications; Example Get Thermal Values

### ICWResults.GetThermalForEntities
Gets thermal values at the specified geometric entities.
```vb
Function GetThermalForEntities( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Thermal component as defined in swsThermalComponent_e NStepNumber Solution step number (use 1 for steady state) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NUnits Temperature unit as defined swsTemperatureUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of thermal results (see Remarks )
Return Value Array of thermal results (see Remarks )
VBA: See CWResults::GetThermalForEntities .
Remarks: This method returns the following array : { node _1 , temp_1, node_2 , temp_2 , node _3 , temp_3 , ..., node _n , temp_n }, where the nodes are integers, and the temperatures are doubles.
Examples: Visual Basic for Applications; Apply Thermostat-controlled Heat Power for Transient Thermal Study; Get Thermal Values

### ICWResults.GetThermalValues
Gets thermal values for all nodes at the specified solution step.
```vb
Function GetThermalValues( ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NStepNumber Solution step number (use 1 for steady state) DispPlane Reference geometry NUnits Temperature unit as defined swsTemperatureUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of thermal results (see Remarks )
Return Value Array of thermal results (see Remarks )
VBA: See CWResults::GetThermalValues .
Remarks: This method returns the following array : { node _1 , temp_1, gradx_1, grady_1, gradz_1, gradn_1, hfluxx_1, hfluxy_1, hfluxz_1, hfluxn_1, node_2 , temp_2 , gradx_2, grady_2, gradz_2, gradn_2, hfluxx_1, hfluxy_1, hfluxz_1, hfluxn_1, node _3 , temp_3 , gradx_3, grady_3, gradz_3, gradn_3, hfluxx_1, hfluxy_1, hfluxz_1, hfluxn_1, ..., node _n , temp_n, gradx_n, grady_n, gradz_n, gradn_n, hfluxx_1, hfluxy_1, hfluxz_1, hfluxn_1 }, where the nodes are integers, and the temperatures, temperature gradients, and heat fluxes are doubles.
Examples: Visual Basic for Applications; Example Get Thermal Values

### ICWResults.GetThermalValuesAtPoints
Gets thermal values at the specified solution step and the specified points.
```vb
Function GetThermalValuesAtPoints( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArrayPoints As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Thermal component as defined in swsThermalComponent_e NStepNumber Solution step number (use 1 for steady state) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArrayPoints Array of point coordinates NUnits Temperature unit as defined swsTemperatureUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of NComponent results for ArrayPoints at NStepNumber
Return Value Array of NComponent results for ArrayPoints at NStepNumber
VBA: See CWResults::GetThermalValuesAtPoints .
Examples: Visual Basic for Applications; Example Get Thermal Values

### ICWResults.GetTimeOrFrequencyAtEachStep
Obsolete. Superseded by ICWResults::GetTimeOrFrequencyAtEachStep2 .
```vb
Function GetTimeOrFrequencyAtEachStep( ByVal ForStressAndStrain As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters ForStressAndStrain 1 for stress or strain results, 0 for other results ErrorCode Error as defined in swsResultsError_e Return Value 0-based array of: Times for time-domain studies - or - Frequencies for frequency-domain studies for all solution steps or mode shapes
Return Value 0-based array of: Times for time-domain studies - or - Frequencies for frequency-domain studies for all solution steps or mode shapes
VBA: See CWResults::GetTimeOrFrequencyAtEachStep .
Examples: Visual Basic for Applications

### ICWResults.GetTimeOrFrequencyAtEachStep2
Gets the times or frequencies for all solution steps or mode shapes in these results.
```vb
Function GetTimeOrFrequencyAtEachStep2( ByVal ForStressAndStrain As System.Boolean, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters ForStressAndStrain -1 or true for stress or strain results, 0 or false for other results ErrorCode Error as defined in swsResultsError_e Return Value 0-based array of: Times for time-domain studies - or - Frequencies for frequency-domain studies for all solution steps or mode shapes
Return Value 0-based array of: Times for time-domain studies - or - Frequencies for frequency-domain studies for all solution steps or mode shapes
Examples: Example Get Frequencies in Mode Shape Plots

### ICWResults.GetTranslationalDisplacement
Gets the translational displacements at the specified solution step.
```vb
Function GetTranslationalDisplacement( ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NStepNumber Solution step number (use 1 for static) DispPlane Reference geometry NUnits Linear units for displacement as defined in swsLinearUnit_e ErrorCode Error as defined in swsResultsError_e Return Value Array of translational displacements
Return Value Array of translational displacements
VBA: See CWResults::GetTranslationalDisplacement .
Examples: Visual Basic for Applications

### ICWResults.GetVelocityComponentForAllStepsAtNode
Gets the specified velocity component for all solution steps at the specified node.
```vb
Function GetVelocityComponentForAllStepsAtNode( ByVal NComponent As System.Integer, ByVal NNodeNum As System.Integer, ByVal DispPlane As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Velocity component as defined in swsVelocityComponent_e NNodeNum Node number DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component NUnits Units of velocity as defined in swsVelocityUnit_e ErrorCode Error as defined in swsResultsError_e
VBA: See CWResults::GetVelocityComponentForAllStepsAtNode .
Examples: Visual Basic for Applications

### ICWResults.GetVelocityForEntities
Gets the specified velocity component for the specified entities and at the specified solution step.
```vb
Function GetVelocityForEntities( ByVal NComponent As System.Integer, ByVal NStepNumber As System.Integer, ByVal DispPlane As System.Object, ByVal ArraySelectedEntities As System.Object, ByVal NUnits As System.Integer, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NComponent Velocity component as defined in swsVelocityComponent_e NStepNumber Solution step number (use 1 for steady state) DispPlane Plane, axis, or coordinate system; valid only if NComponent is set to a directional component; specify nothing if NComponent is not a directional component ArraySelectedEntities Array of geometric entities NUnits Units of velocity as defined in swsVelocityUnit_e ErrorCode Error as defined in swsResultsError_e
VBA: See CWResults::GetVelocityForEntities .
Examples: Visual Basic for Applications

### ICWResults.RunStressHotSpotDiagnostics
Obsolete. Superseded by ICWResults::RunStressHotSpotDiagnosticsAndDetectSingularities .
```vb
Function RunStressHotSpotDiagnostics( ByVal NSensitivityFactor As System.Integer, ByVal BRunForNodes As System.Boolean, ByRef BFoundHotSpots As System.Boolean ) As System.Integer
```
Parameters NSensitivityFactor 5 <= Highest equivalent strain percent <= 100 BRunForNodes True to run for nodes, false to not (see Remarks ) BFoundHotSpots True if stress hot spots are found, false if not (see Remarks ) Return Value Error code as defined in swsRunStressHotSpotDiagnosticsError_e
Return Value Error code as defined in swsRunStressHotSpotDiagnosticsError_e
VBA: See CWResults::RunStressHotSpotDiagnostics .
Remarks: After calling this method, you can call: ICWResults::CreateStressHotSpotPlot and ICWResults::GetDetectedHotSpotElements if BRunForNodes is false. ICWResults::GetDetectedHotSpotNodes if BRunForNodes is true.
Examples: Visual Basic for Applications

### ICWResults.RunStressHotSpotDiagnosticsAndDetectSingularities
Detects stress hot spots and singularities.
```vb
Function RunStressHotSpotDiagnosticsAndDetectSingularities( ByVal NSensitivityFactor As System.Integer, ByVal BRunForNodes As System.Boolean, ByVal NMeshRefineLevels As System.Integer, ByVal DElemSizeReductionFactor As System.Double, ByVal dGrowthRatio As System.Double, ByVal NResultsRestoreOption As System.Integer, ByRef BFoundHotSpots As System.Boolean, ByRef BFoundSingularities As System.Boolean ) As System.Integer
```
Parameters NSensitivityFactor 5 <= Highest equivalent strain percent <= 100 (see Remarks ) BRunForNodes True to run for nodes, false to run for elements (see Remarks ) NMeshRefineLevels 2 <= Mesh refinement level to detect singularities <= 3 DElemSizeReductionFactor 0.01 <= Element size reduction factor across levels to detect singularities <= 0.9 dGrowthRatio 1.0 <= Element size growth ratio to detect singularities <= 3.1 NResultsRestoreOption Results restore option as defined in swsStressHotSpotResultsRestoreOptions_e use after detecting stress hot spots and singularities BFoundHotSpots True if stress hot spots are found, false if not BFoundSingularities True if singularities are found, false if not Return Value Error code as defined in swsRunStressHotSpotDiagnosticsError_e
Return Value Error code as defined in swsRunStressHotSpotDiagnosticsError_e
VBA: See CWResults::RunStressHotSpotDiagnosticsAndDetectSingularities .
Remarks: This method corresponds to the Stress Hot Spot diagnostic tool that identifies adjacent elements with high stress gradiants. High stress gradients occur between elements with significantly different stress values which indicates the presence of stress singularities or stress concentrations. Set NSensitivity to a lower number to restrict the search to elements with the highest strain values. Set NSensitivity to a higher number to expand the search to elements with lower strain values. After calling this method, you can call: ICWResults::CreateStressHotSpotPlot and ICWResults::GetDetectedHotSpotElements if BRunForNodes is false, or ICWResults::GetDetectedHotSpotNodes if BRunForNodes is true.
Examples: Visual Basic for Applications

### ICWResults.SavePlotsAseDrawings
Saves the specified results plot as an eDrawings file with the specified name in the specified location.
```vb
Function SavePlotsAseDrawings( ByVal SFileLocation As System.String, ByVal SFileName As System.String, ByVal SPlotName As System.String ) As System.Integer
```
Parameters SFileLocation Path where to save the eDrawings file SFileName Name of eDrawings file to save SPlotName Plot to save as an eDrawings file (see Remarks ) Return Value Error as defined by swsSaveeDrawingsErrorCode_e
Return Value Error as defined by swsSaveeDrawingsErrorCode_e
VBA: See CWResults::SavePlotsAseDrawings .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName.
Examples: Visual Basic for Applications; Example Apply Thermostat-controlled Heat Power for Transient Thermal Study

### ICWResults.SetAverageStressesAtMidnodes
Sets whether to average stresses at mid-nodes for this study.
```vb
Function SetAverageStressesAtMidnodes( ByVal BValue As System.Boolean ) As System.Boolean
```
Parameters BValue True to average stresses at mid-nodes, false to not Return Value True if option successfully set, false if not
Return Value True if option successfully set, false if not
VBA: See CWResults::SetAverageStressesAtMidnodes .
Remarks: This method is valid only for static, nonlinear, drop test, pressure vessel, and topology studies.
Examples: Visual Basic for Applications

### ICWResults.SetDeformedBodyFailedSewOption
Sets the option to use when a deformed body fails to sew into a solid object.
```vb
Sub SetDeformedBodyFailedSewOption( ByVal NFailedSewOption As System.Integer )
```
Parameters NFailedSewOption Option to use when a deformed body fails to sew into a solid object as defined in swsCreateDeformedBodyFailedSewOption_e
VBA: See CWResults::SetDeformedBodyFailedSewOption .
Examples: Visual Basic for Applications

### ICWResults.SetPlotColorOptions
Sets the color options for the specified plot.
```vb
Function SetPlotColorOptions( ByVal SPlotName As System.String, ByVal NType As System.Integer, ByVal NContour As System.Integer, ByVal NBaseColor As System.Integer, ByVal BFlip As System.Boolean, ByVal BSpecifyColorLimit As System.Boolean, ByVal VarColor As System.Object ) As System.Integer
```
Parameters SPlotName Plot name (see Remarks ) NType Color map as defined in swsColorChartOptionLegendTypeValue_e (see Remarks ) NContour 2 <= Number of gradient colors <= 24 NBaseColor Number of base colors for NType; valid only if NType is swsColorChartOptionLegendTypeValue_e.swsColorChartOptionLegendUserDefined (see Remarks ) BFlip True to reverse the color mapping, false to not BSpecifyColorLimit True to specify a color for values above the yield limit, false to not; valid only for von Mises stress plots of single body parts where the yield strength is defined VarColor Array of RGB triplets: RGB color for values above the yield limit; valid only if BSpecifyColorLimit is true One or more RGB triplets of user-defined colors Return Value Error as defined in swsResultsError_e
Return Value Error as defined in swsResultsError_e
VBA: See CWResults::SetPlotColorOptions .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. If NType is swsColorChartOptionLegendTypeValue_e... Then NBaseColor is... swsColorChartOptionLegendDefault 5 swsColorChartOptionLegendRainbow 7 swsColorChartOptionLegendGrayScale 2 swsColorChartOptionLegendUserDefined 2-9
Examples: Visual Basic for Applications

### ICWResults.SetPlotDisplayOptions
Sets the display options for the specified plot.
```vb
Function SetPlotDisplayOptions( ByVal SPlotName As System.String, ByVal ArrayInput As System.Object ) As System.Integer
```
Parameters SPlotName Plot name (see Remarks ) ArrayInput Array of plot display options (see Remarks ) Return Value Error as defined in swsResultsError_e
Return Value Error as defined in swsResultsError_e
VBA: See CWResults::SetPlotDisplayOptions .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. The 0-based array contains eight elements: Array element... Contains... 0 Boolean; whether to display minimum value 1 Boolean; whether to display maximum value 2 Boolean; whether to display plot details 3 Boolean; whether to display plot legend 4 Boolean; whether to display minimum/maximum ranges only for parts 5 Boolean; whether to allow user-defined minimum/maximum values 6 Double; user-defined minimum value; valid only if array element 5 is true 7 Double; user-defined maximum value; valid only if array element 5 is true
Examples: Visual Basic for Applications

### ICWResults.SetPlotPositionFormatOptions
Sets the position/format options for the specified plot.
```vb
Function SetPlotPositionFormatOptions( ByVal SPlotName As System.String, ByVal ArrayInput As System.Object ) As System.Integer
```
Parameters SPlotName Plot name (see Remarks ) ArrayInput Array of position/format options (see Remarks ) Return Value Error as defined in swsResultsError_e
Return Value Error as defined in swsResultsError_e
VBA: See CWResults::SetPlotPositionFormatOptions .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. The 0-based array contains eight elements: Array element... Contains... 0 Double[0,1]; horizontal position from the left of the graphics area; percentage of the window width 1 Double[0,1]; vertical position from the top of the graphics area; percentage of the window height 2 Chart width option as defined in swsColorChartWidthOptionValue_e 3 Chart number format as defined in swsColorChartNumberFormatOptionValue_e 4 0 <= Number of decimal places to display <= 16 5 Boolean; whether to display chart numbers with a 1000 comma separator; valid only if array element 3 is swsColorChartNumberFormatOptionValue_e.swsColorChartNumberFormatFloating or swsColorChartNumberFormatOptionValue_e.swsColorChartNumberFormatGeneral 6 Boolean; whether to use different number format for small numbers (0.001 < |x| < 1000); valid only if array element 3 is swsColorChartNumberFormatOptionValue_e.swsColorChartNumberFormatScientific 7 Number format for small numbers as defined in swsColorNumberFormatUseDiffNumberFormatOptionValue_e ; valid only if array element 6 is true
Examples: Visual Basic for Applications

### ICWResults.SetPlotSettings
Sets various settings for the specified plot.
```vb
Function SetPlotSettings( ByVal SPlotName As System.String, ByVal ArrayInput As System.Object ) As System.Integer
```
Parameters SPlotName Plot name (see Remarks ) ArrayInput Array of plot settings (see Remarks ) Return Value Error code as defined in swsPlotSettingsErrorCode_e
Return Value Error code as defined in swsPlotSettingsErrorCode_e
VBA: See CWResults::SetPlotSettings .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. The 0-based array contains 11 elements: Array element... Contains... 0 Display option for active fringe plot as defined in swsPlotFringeSettingsOptionValue_e 1 Display option for model boundary as defined in swsPlotBoundarySettingsOptionValue_e 2 R[0-255] value for model/mesh color; valid only if array element 1 is not swsPlotBoundarySettingsOptionValue_e.swsPlotBoundaryNone 3 G[0-255] value for model/mesh color; valid only if array element 1 is not swsPlotBoundarySettingsOptionValue_e.swsPlotBoundaryNone 4 B[0-255] value for model/mesh color; valid only if array element 1 is not swsPlotBoundarySettingsOptionValue_e.swsPlotBoundaryNone 5 Boolean; whether to superimpose the undeformed model on the deformed model 6 Deformed plot translucent color option as defined in swsPlotDeformedShapeOptionSuperImposeValue_e ; valid only if array element 5 is true 7 R[0-255] value for single translucent color; valid only if array elmeent 5 is true, and array element 6 is swsPlotDeformedShapeOptionSuperImposeValue_e.swsPlotDeformedShapeSuperImposeModel_TranslucentSingleColor 8 G[0-255] value for single translucent color; valid only if array element 5 is true, and array element 6 is swsPlotDeformedShapeOptionSuperImposeValue_e.swsPlotDeformedShapeSuperImposeModel_TranslucentSingleColor 9 B[0-255] value for single translucent color; valid only if array element 5 is true, and array element 6 is swsPlotDeformedShapeOptionSuperImposeValue_e.swsPlotDeformedShapeSuperImposeModel_TranslucentSingleColor 10 0.0 <= Transparency or intensity of the translucent single color or part colors <= 1.0; valid only if array element 5 is true
Examples: Visual Basic for Applications

### ICWResults.SetPlotSettingsOptionForHiddenAndExcludedBody
Sets hidden and excluded body options for the specified plot.
```vb
Function SetPlotSettingsOptionForHiddenAndExcludedBody( ByVal SPlotName As System.String, ByVal ArrayInput As System.Object ) As System.Integer
```
Parameters SPlotName Plot name (see Remarks ) ArrayInput Array of hidden and excluded body options (see Remarks ) Return Value Error code as defined in swsPlotSettingsErrorCode_e
Return Value Error code as defined in swsPlotSettingsErrorCode_e
VBA: See CWResults::SetPlotSettingsOptionForHiddenAndExcludedBody .
Remarks: Call ICWResults::GetPlotNames to populate SPlotName. ArrayInput is a 0-based array that contains 12 elements: Array element... Contains... 0 Boolean; whether to show hidden bodies 1 Hidden body translucent color option as defined in swsPlotShowHiddenBodiesOptionValue_e ; valid only if array element 0 is true 2 R[0-255] value for single translucent color; valid only if array element 0 is true, and array element 1 is swsPlotShowHiddenBodiesOptionValue_e.swsPlotHiddenBodyTranslucentSingleColor 3 G[0-255] value for single translucent color; valid only if array element 0 is true, and array element 1 is swsPlotShowHiddenBodiesOptionValue_e.swsPlotHiddenBodyTranslucentSingleColor 4 B[0-255] value for single translucent color; valid only if array element 0 is true, and array element 1 is swsPlotShowHiddenBodiesOptionValue_e.swsPlotHiddenBodyTranslucentSingleColor 5 0.0 <= transparency or intensity of the translucent single color or part colors <= 1.0; valid only if array element 0 is true 6 Boolean; whether to show excluded bodies 7 Excluded body translucent color option as defined in swsPlotShowExcludedBodiesOptionValue_e ; valid only if array element 6 is true 8 R[0-255] value for single translucent color; valid only if array element 6 is true, and array element 7 is swsPlotShowExcludedBodiesOptionValue_e.swsPlotExcludedBodyTranslucentSingleColor 9 G[0-255] value for single translucent color; valid only if array element 6 is true, and array element 7 is swsPlotShowExcludedBodiesOptionValue_e.swsPlotExcludedBodyTranslucentSingleColor 10 B[0-255] value for single translucent color; valid only if array element 6 is true, and array element 7 is swsPlotShowExcludedBodiesOptionValue_e.swsPlotExcludedBodyTranslucentSingleColor 11 0.0 <= transparency or intensity of the translucent single color or part colors <= 1.0; valid only if array element 6 is true
Examples: Visual Basic for Applications; Example Analyze Part

## ICWPlot

### ICWPlot.ActivatePlot
Activates this plot.
```vb
Function ActivatePlot() As System.Integer
```
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ActivatePlot .
Examples: Visual Basic for Applications

### ICWPlot.AddRefGeometry
Specifies the reference geometry for directional component plots.
```vb
Function AddRefGeometry( ByVal NDispType As System.Integer, ByVal DispPlaneAxisCoordSys As System.Object ) As System.Integer
```
Parameters NDispType Type of reference geometry as defined in swsRefDispType_e DispPlaneAxisCoordSys Plane, axis, or coordinate system Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::AddRefGeometry .
Remarks: This method is valid only for the following directional component plots: Displacement Mode Shape/Amplitude Strain Stress Thermal
Examples: Visual Basic for Applications

### ICWPlot.ApplyNameViewOrientation
Obsolete. Superseded by ICWPlot::ApplyNameViewOrientation2 .
```vb
Function ApplyNameViewOrientation( ByVal SNameViewOrientation As System.String ) As System.Integer
```
Parameters SNameViewOrientation Name view orientation Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ApplyNameViewOrientation .
Remarks: This method is valid for the following plots: Acceleration Beam diagram Displacement Mode shape/amplitude Fatigue Strain Strain energy density Stress Thermal Vector
Examples: Visual Basic for Applications

### ICWPlot.ApplyNameViewOrientation2
Associates this plot with the specified view orientation.
```vb
Function ApplyNameViewOrientation2( ByVal NNameViewOrientation As System.Integer ) As System.Integer
```
Parameters NNameViewOrientation View orientation as defined in swsNameViewOrientation_e Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ApplyNameViewOrientation2 .
Remarks: This method is valid for the following plots: Acceleration Beam diagram Displacement Mode shape/amplitude Fatigue Strain Strain energy density Stress Thermal Vector
Examples: Visual Basic for Applications

### ICWPlot.GetLinearizedStressValuesAlongSCL
Calculates the membrane and bending stresses between the specified points on a planar section view of this pressure vessel stress plot.
```vb
Function GetLinearizedStressValuesAlongSCL( ByVal NIntermediatePoints As System.Integer, ByVal DPoint1X As System.Double, ByVal DPoint1Y As System.Double, ByVal DPoint1Z As System.Double, ByVal DPoint2X As System.Double, ByVal DPoint2Y As System.Double, ByVal DPoint2Z As System.Double, ByVal DispReferencePlane As System.Object, ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters NIntermediatePoints 0 < Number of intermediate points between the points specified by DPoint1X, DPoint1Y, DPoint1Z, DPoint2X, DPoint2Y, and DPoint2Z to define the resolution of the SCL axis < 100 DPoint1X X coordinate of Point 1 on the section view (see Remarks ) DPoint1Y Y coordinate of Point 1 on the section view (see Remarks ) DPoint1Z Z coordinate of Point 1 on the section view (see Remarks ) DPoint2X X coordinate of Point 2 on the section view (see Remarks ) DPoint2Y Y coordinate of Point 2 on the section view (see Remarks ) DPoint2Z Z coordinate of Point 2 on the section view (see Remarks ) DispReferencePlane Datum plane reference that creates a section view containing Point 1 and Point 2 (see Remarks ) ErrorCode Result code as defined in swsResultStressLinearizationErrors_e Return Value Array of seven membrane and bending stress values that are linearized along Stress Classification Lines (SCL) (see Remarks )
Return Value Array of seven membrane and bending stress values that are linearized along Stress Classification Lines (SCL) (see Remarks )
VBA: See CWPlot::GetLinearizedStressValuesAlongSCL .
Remarks: This method adheres to current ASME Pressure Vessel design codes (ASME BPVC Section VIII, Annex 5-A.4.1.2) for deriving the membrane, bending, and peak components of a stress distribution based on a local coordinate system that is defined by the orientation of Stress Classification Lines (SCL). In the SCL coordinate system, SCL axis (Tangential) is the x direction, the axis perpendicular to the section plane (Hoop) is the z direction, and the orthogonal axis (Normal) is the y direction. This method calculates: A membrane stress tensor that is comprised of the average of each stress component along the SCL. Bending stresses for the hoop and normal components, but not for the tangential component that is parallel to the SCL. The bending stress tensor is comprised of the linear varying portion of each stress component along the SCL. A peak stress tensor. Three principal stresses (P1, P2, P3) at the ends of the SCL based on components of membrane and membrane + bending stress. Equivalent stresses at the ends of the SCL based on components of membrane and membrane + bending stress. Before calling this method: Create a pressure vessel study by defining a Result Combination Setup of one static study with von Mises stress results. Analyse the pressure vessel study to generate a von Mises stress results plot. Specify DispReferencePlane using a reference plane that sections the stress plot. Specify DPoint1X, DPoint1Y, DPoint1Z, DPoint2X, DPoint2Y, DPoint2Z with the coordinates of two points on a stress line of the section plane in step 3. To locate valid points on the stress line of a section plane: Make sure the document units are set to MKS in the graphics area. Right-click the pressure vessel stress plot and select Section Clipping . Specify Top for Section 1 > Reference Entity . Click OK. Right-click the pressure vessel stress plot and select Linearize . Select two points in the graphics area that fall on the sectioned plot, prefereably spanning the thickness. (The tangential line goes from the first location to the second. The hoop is the outward normal to the cutting plane. The normal completes the right--handed coordinate system.) Specify 5 for No of Intermediate Points . Click Calculate . Click Save and export results as a CSV file. Open the CSV file in Excel. Locate the coordinates of the start and end points of the stress line to be used to specify DPoint1X, DPoint1Y, DPoint1Z, DPoint2X, DPoint2Y, DPoint2Z. Be sure to specify DispReferencePlane using the same section plane used to locate the points. The returned array contains the following seven values: [ membrane_stress , bending_point_1 , ( membrane_stress + bending_point_1 ), bending_point_2 , ( membrane_stress + bending_point_2 ), peak_point_1 , peak_point_2 ] NOTE: This method is valid only for solid meshes.
Examples: Visual Basic for Applications; Example Get Linearized Stress Results

### ICWPlot.GetMinMaxResultValues
Gets the minimum and maximum values in this plot.
```vb
Function GetMinMaxResultValues( ByRef ErrorCode As System.Integer ) As System.Object
```
Parameters ErrorCode Error code as defined in swsResultPlotErrorCode_e Return Value Array of minimum and maximum values (see Remarks )
Return Value Array of minimum and maximum values (see Remarks )
VBA: See CWPlot::GetMinMaxResultValues .
Remarks: This method returns the following array: { node_with_minimum_value , minimum_value , node_with_maximum_value , maximum_value }, where the nodes are integers, and the values are in scientific notation.
Examples: Visual Basic for Applications

### ICWPlot.GetPlotName
Gets the name of this plot.
```vb
Function GetPlotName( ByRef SPlotName As System.String ) As System.Integer
```
Parameters SPlotName Plot name Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::GetPlotName .
Examples: Visual Basic for Applications

### ICWPlot.GetResultsProbeManager
Gets the results probe for this plot.
```vb
Function GetResultsProbeManager( ByRef ErrorCode As System.Integer ) As CWResultsProbeManager
```
Parameters ErrorCode Error code as defined in swsResultPlotErrorCode_e Return Value ICWResultsProbeManager
Return Value ICWResultsProbeManager
VBA: See CWPlot::GetResultsProbeManager .
Examples: Visual Basic for Applications

### ICWPlot.IsoValue
Gets or sets the load level of this Design Insight plot.
```vb
Property IsoValue As System.Integer
```
Property Value 0 <= load level <= 100
VBA: See CWPlot::IsoValue .
Examples: Visual Basic for Applications; Example Create a Design Insight Plot for a Static Study

### ICWPlot.Set2DPlanarRevolveAngle
Sets the revolve angle of the 3D plot of a 2D Simplification study.
```vb
Function Set2DPlanarRevolveAngle( ByVal D2DPlanarRevolveAngle As System.Double ) As System.Integer
```
Parameters D2DPlanarRevolveAngle Angle of revolution about a selected axis to create the 3D plot Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::Set2DPlanarRevolveAngle .
Remarks: This method is valid only: for 2D Simplification axi-symmetric studies where the geometry, loads, and restraints are symmetric (360 degrees) about an axis. if ICWPlot::ShowAs3DPlot is set to true.
Examples: Visual Basic for Applications

### ICWPlot.SetComponentUnitAndValueByElem
Sets the component, units, and values to plot.
```vb
Function SetComponentUnitAndValueByElem( ByVal NComponent As System.Integer, ByVal NUnits As System.Integer, ByVal BValueByElem As System.Boolean ) As System.Integer
```
Parameters NComponent Component to plot (see Remarks ) NUnits Units as appropriate for NComponent (see Remarks ) BValueByElem True to plot element values, false to plot node values Return Value Error code as defined by swsResultPlotErrorCode_e
Return Value Error code as defined by swsResultPlotErrorCode_e
VBA: See CWPlot::SetComponentUnitAndValueByElem .
Remarks: If this plot's result type is swsPlotResultTypes_e... Then NComponent contains the component to plot as defined by... And NUnits contains the units as defined by... swsResultAcceleration swsAccelerationComponent_e swsAccelerationUnit_e swsResultBeamDiagram swsBeamForceType_e swsForceUnit_e swsResultBeamStress swsBeamStressType_e swsStrengthUnit_e swsResultDesignInsight N/A N/A swsResultDisplacementOrAmplitude swsDisplacementComponent_e (Displacement) swsFrequencyBucklingResultDisplacementComponentTypes_e (Amplitude for frequency and buckling) swsLinearUnit_e (displacement) or swsForceUnit_e (reaction force) swsResultEdgeWeldConnector N/A N/A swsResultFactorOfSafety swsFOS_NonCompositeCriterion_e (Non-composite shells) swsFOS_CompositeCriterion_e (Composite shells) N/A swsResultFatigue swsFatigueComponent_e N/A swsResultPinBoltBearing N/A N/A swsResultStrain swsStrainComponent_e swsStrengthUnit_e swsResultStress swsStressComponent_e swsStrengthUnit_e swsResultThermal swsThermalComponent_e swsTemperatureUnit_e swsResultVelocity swsVelocityComponent_e swsVelocityUnit_e
Examples: Visual Basic for Applications

### ICWPlot.SetDisplayOption
Sets the display option for this stress results plot of a mixed mesh model.
```vb
Function SetDisplayOption( ByVal NDisplayOption As System.Integer ) As System.Integer
```
Parameters NDisplayOption Display option as defined in swsDisplayOption_e Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::SetDisplayOption .
Remarks: This method is valid only for stress plots of mixed mesh models having beams and solids/shells. At least one joint group must be present in the study.
Examples: Visual Basic for Applications

### ICWPlot.SetFOSPlot
Sets whether to plot the factor of safety throughout the model.
```vb
Function SetFOSPlot( ByVal BFOSPlot As System.Boolean ) As System.Integer
```
Parameters BFOSPlot True to plot the factor of safety throughout the model, false to not Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::SetFOSPlot .
Examples: Visual Basic for Applications

### ICWPlot.SetModeShape
Sets the mode shape number for this amplitude or displacement plot.
```vb
Function SetModeShape( ByVal NModeShape As System.Integer ) As System.Integer
```
Parameters NModeShape Mode shape number for which amplitudes or displacements are plotted Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::SetModeShape .
Examples: Visual Basic for Applications; Example Create Amplitude Plot and Set Mode Shape

### ICWPlot.SetNormalizeModeShape
Sets whether to normalize the mode shape of this plot.
```vb
Function SetNormalizeModeShape( ByVal BNormModeShape As System.Boolean ) As System.Integer
```
Parameters BNormModeShape True to normalize the mode shape (maximum amplitude is 1), false to not Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::SetNormalizeModeShape .
Remarks: This method is valid only for amplitude plots.
Examples: Visual Basic for Applications

### ICWPlot.SetPlotStepNumber
Sets the number of the solution step to plot.
```vb
Function SetPlotStepNumber( ByVal NStepNumber As System.Integer ) As System.Integer
```
Parameters NStepNumber Number of solution step to plot Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::SetPlotStepNumber .
Remarks: This method is valid only for the following plots: Acceleration Displacement Strain Stress Transient thermal Velocity
Examples: Visual Basic for Applications; Example Create Nonlinear Dynamic Plots

### ICWPlot.SetPlotTitle
Sets a title for this plot.
```vb
Function SetPlotTitle( ByVal SPlotTitle As System.String ) As System.Integer
```
Parameters SPlotTitle Plot title Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::SetPlotTitle .
Remarks: This method is valid for the following plots: Acceleration Beam diagram Displacement Mode shape/amplitude Fatigue Strain Strain energy density Stress Thermal Vector
Examples: Visual Basic for Applications

### ICWPlot.ShowAs3DPlot
Sets whether to render the 2D Simplification results in 3D.
```vb
Function ShowAs3DPlot( ByVal BShowAs3D As System.Boolean ) As System.Integer
```
Parameters BShowAs3D True to render the results in 3D, false to not Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ShowAs3DPlot .
Remarks: This method is valid only for 2D Simplification studies. Also call ICWPlot::Set2DPlanarRevolveAngle to set the angle of revolution to generate the 3D plot.
Examples: Visual Basic for Applications

### ICWPlot.ShowAvgResultsAcrossBoundaryForParts
Sets whether to show average results at the common nodes between two components in an assembly.
```vb
Function ShowAvgResultsAcrossBoundaryForParts( ByVal BAvgResultsAcrossBoundaryForParts As System.Boolean ) As System.Integer
```
Parameters BAvgResultsAcrossBoundaryForParts True to show average results at the common nodes between two components in the assembly, false to not Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ShowAvgResultsAcrossBoundaryForParts .
Remarks: This method is valid only for stress plots of assemblies.
Examples: Visual Basic for Applications

### ICWPlot.ShowBeamProfile
Sets whether to render a beam diagram on the actual beam geometry.
```vb
Function ShowBeamProfile( ByVal BShowBeamProfile As System.Boolean ) As System.Integer
```
Parameters BShowBeamProfile True to render a beam diagram on the beam geometry, false to not Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ShowBeamProfile .
Remarks: This method is valid only for studies with beams and the following plots: Displacement Stress
Examples: Visual Basic for Applications

### ICWPlot.ShowDeformedPlot
Sets whether and how to plot on the deformed shape of the model.
```vb
Function ShowDeformedPlot( ByVal BShowDeformed As System.Boolean, ByVal NDeformOption As System.Integer, ByVal DUserDefValue As System.Double, ByVal BShowColors As System.Boolean ) As System.Integer
```
Parameters BShowDeformed True to plot on the deformed shape, false to not NDeformOption Deformed shape option as defined in swsDeformType_e DUserDefValue User-defined scale factor; valid only if NDeformOption is swsDeformType_e.swsUserDefined BShowColors True to show colors, false to not Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ShowDeformedPlot .
Remarks: This method is valid for the following plots: Acceleration Displacement Mode shape/amplitude Strain Strain energy density Stress Vector
Examples: Visual Basic for Applications

### ICWPlot.ShowNormalizeValuesFrom0To1
Sets whether to normalize values from 0 to 1 in this plot.
```vb
Function ShowNormalizeValuesFrom0To1( ByVal BNormalizeValuesFrom0To1 As System.Boolean ) As System.Integer
```
Parameters BNormalizeValuesFrom0To1 True to normalize values from 0 to 1, false to not Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ShowNormalizeValuesFrom0To1 .
Examples: Visual Basic for Applications

### ICWPlot.ShowPlotOnSelEntities
Sets whether to display this plot only for selected entities.
```vb
Function ShowPlotOnSelEntities( ByVal BSelFaceOnly As System.Boolean, ByVal ArraySelectedEntities As System.Object ) As System.Integer
```
Parameters BSelFaceOnly True to display plot only for ArraySelectedEntities, false to display plot for all entities ArraySelectedEntities Array of selected faces and bodies; valid only if BSelFaceOnly is true Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ShowPlotOnSelEntities .
Remarks: This method is valid for the following plots only if ICWPlot::ShowShellin3D and ICWPlot::ShowTensorOrVector are both false: Acceleration Displacement Mode shape/amplitude Strain Stress Velocity
Examples: Visual Basic for Applications

### ICWPlot.ShowShellin3D
Sets whether to render the results on 3D shell bodies.
```vb
Function ShowShellin3D( ByVal BShowShellin3D As System.Boolean ) As System.Integer
```
Parameters BShowShellin3D True to render results on 3D shell bodies, false to not Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ShowShellin3D .
Remarks: This method is valid for the following plots only if ICWPlot::ShowTensorOrVector and ICWPlot::ShowPlotOnSelEntities are both false: Acceleration Displacement Mode shape/amplitude Strain Stress Velocity
Examples: Visual Basic for Applications

### ICWPlot.ShowTensorOrVector
Sets whether to render a vector or tensor plot.
```vb
Function ShowTensorOrVector( ByVal BShowTensorOrVector As System.Boolean ) As System.Integer
```
Parameters BShowTensorOrVector True to render a tensor or vector plot, false to not Return Value Error code as defined in swsResultPlotErrorCode_e
Return Value Error code as defined in swsResultPlotErrorCode_e
VBA: See CWPlot::ShowTensorOrVector .
Remarks: If the results plot is for component... And BShowTensorOrVector is... Then the plot renders a... von Mises stress True Tensor of the three principal stresses at each node or element center Directional stress, displacement, acceleration, or velocity True Vector of the magnitude and direction of the component at each node Thermal resultant temperature gradient or resultant heat flux True Vector of the magnitude and direction of the component at each node
Examples: Visual Basic for Applications

## ICWResultsProbeManager

### ICWResultsProbeManager.BeginProbing
Starts the results probe.
```vb
Function BeginProbing() As System.Integer
```
Return Value Error code as defined in swsProbePostResultErrorCode_e
VBA: See CWResultsProbeManager::BeginProbing .
Examples: Visual Basic for Applications

### ICWResultsProbeManager.CaptureProbedResultPlotAsPNGImage
Saves the probed results plot as a PNG image in the specified file and folder.
```vb
Function CaptureProbedResultPlotAsPNGImage( ByVal SImageFolder As System.String, ByVal SImageName As System.String ) As System.Integer
```
Parameters SImageFolder Name of folder to which to save the image SImageName Name of file to which to save the image Return Value Error code as defined in swsProbePostResultErrorCode_e
Return Value Error code as defined in swsProbePostResultErrorCode_e
VBA: See CWResultsProbeManager::CaptureProbedResultPlotAsPNGImage .
Examples: Visual Basic for Applications

### ICWResultsProbeManager.ClearSelectionAndAnnotations
Clears selections and annotations for the results probe.
```vb
Sub ClearSelectionAndAnnotations()
```
VBA: See CWResultsProbeManager::ClearSelectionAndAnnotations .
Examples: Visual Basic for Applications

### ICWResultsProbeManager.EndProbing
Finishes the results probe.
```vb
Function EndProbing() As System.Integer
```
Return Value Error code as defined in swsProbePostResultErrorCode_e
VBA: See CWResultsProbeManager::EndProbing .
Examples: Visual Basic for Applications

### ICWResultsProbeManager.ProbingOption
Gets or sets the option for this results probe.
```vb
Property ProbingOption As System.Integer
```
Property Value Option as defined in swsProbePostResultOption_e
VBA: See CWResultsProbeManager::ProbingOption .
Examples: Visual Basic for Applications

### ICWResultsProbeManager.ShowAnnotationOnSelNodeElems
Obsoleted. Superseded by ICWResultsProbeManager::ShowAnnotationsOnNodesOrElems .
```vb
Function ShowAnnotationOnSelNodeElems( ByVal ArraySelectedNodeElems As System.Object, ByRef ArrayAnnotationWarnings As System.Object ) As System.Integer
```
Parameters ArraySelectedNodeElems Array of nodes or elements ArrayAnnotationWarnings Array of warnings for specified nodes or elements as defined in swsProbePostResultNodeElementSelectionWarning_e Return Value Error code as defined in swsProbePostResultErrorCode_e
Return Value Error code as defined in swsProbePostResultErrorCode_e
VBA: See CWResultsProbeManager::ShowAnnotationOnSelNodeElems .
Examples: Visual Basic for Applications

### ICWResultsProbeManager.ShowAnnotationsOnNodesOrElems
Shows results probe annotations for the specified nodes or elements.
```vb
Function ShowAnnotationsOnNodesOrElems( ByVal SSelectedNodesOrElems As System.String, ByRef ArrayAnnotationWarnings As System.Object ) As System.Integer
```
Parameters SSelectedNodesOrElems Comma-delimited string of nodes or elements, e.g., "1,25,30-35" ArrayAnnotationWarnings Array of warnings for specified nodes or elements as defined in swsProbePostResultNodeElementSelectionWarning_e Return Value 0 if successful; otherwise error code as defined in swsProbePostResultErrorCode_e
Return Value 0 if successful; otherwise error code as defined in swsProbePostResultErrorCode_e
VBA: See CWResultsProbeManager::ShowAnnotationsOnNodesOrElems .
Examples: Visual Basic for Applications

## ICWMassPropertiesManager

### ICWMassPropertiesManager.AddBodies
Adds the specified bodies to the calculation of mass properties.
```vb
Function AddBodies( ByVal DispArray As System.Object ) As System.Integer
```
Parameters DispArray Array of bodies Return Value Error code as defined in swsSimMassPropertiesError_e
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::AddBodies .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.AddBoltConnectors
Adds the specified bolt connectors to the calculation of mass properties.
```vb
Function AddBoltConnectors( ByVal NameArray As System.Object ) As System.Integer
```
Parameters NameArray Array of names of bolt connectors Return Value Error code as defined in swsSimMassPropertiesError_e
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::AddBoltConnectors .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.AddDistributedMasses
Adds the specified distributed masses to the calculation of mass properties.
```vb
Function AddDistributedMasses( ByVal NameArray As System.Object ) As System.Integer
```
Parameters NameArray Array of names of distributed masses Return Value Error code as defined in swsSimMassPropertiesError_e
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::AddDistributedMasses .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.AddLinkageRods
Adds the specified linkage rod connectors to the calculation of mass properties.
```vb
Function AddLinkageRods( ByVal NamesArray As System.Object ) As System.Integer
```
Parameters NamesArray Array of the names of linkage rod connectors to add to the mass properties calculation Return Value Error code as defined by swsSimMassPropertiesError_e
Return Value Error code as defined by swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::AddLinkageRods .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.AddPinConnectors
Adds the specified pin connectors to the calculation of mass properties.
```vb
Function AddPinConnectors( ByVal NameArray As System.Object ) As System.Integer
```
Parameters NameArray Array of names of pin connectors Return Value Error code as defined in swsSimMassPropertiesError_e
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::AddPinConnectors .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.AddRemoteLoadMasses
Adds the specified remote load masses to the calculation of mass properties.
```vb
Function AddRemoteLoadMasses( ByVal NameArray As System.Object ) As System.Integer
```
Parameters NameArray Array of names of remote load masses Return Value Error code as defined in swsSimMassPropertiesError_e
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::AddRemoteLoadMasses .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.Area
Gets the total surface area.
```vb
ReadOnly Property Area As System.Double
```
Property Value Surface area in square meters
VBA: See CWMassPropertiesManager::Area .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.ClearAll
Removes all entities from the calculation of mass properties.
```vb
Sub ClearAll()
```
VBA: See CWMassPropertiesManager::ClearAll .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.GetCenterOfMass
Gets the center of mass.
```vb
Function GetCenterOfMass() As System.Object
```
Return Value Array of x, y, and z coordinates of the center of mass
VBA: See CWMassPropertiesManager::GetCenterOfMass .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.GetMomentOfInertia
Gets the moments of inertia with respect to the center of mass and the output coordinate system.
```vb
Function GetMomentOfInertia() As System.Object
```
Return Value Array of nine moment of inertia components
VBA: See CWMassPropertiesManager::GetMomentOfInertia .
Remarks: The moment of inertia is calculated with respect to the output coordinate system. Before calling this method, call ICWMassPropertiesManager::SetCoordinateSystemName or ICWMassPropertiesManager::SetCoordinateSystemToDefault . The moment of inertia components are in units of kilograms * meters squared. This method returns nine doubles as follows: [ lxx , lxy , lxz , l yx , lyy , lyz , lzx , lzy , lzz ]
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.GetMomentOfInertiaAtOutputCoordinateSystem
Gets the moments of inertia with respect to the output coordinate system.
```vb
Function GetMomentOfInertiaAtOutputCoordinateSystem() As System.Object
```
Return Value Array of nine moment of inertia components (see Remarks )
VBA: See CWMassPropertiesManager::GetMomentOfInertiaAtOutputCoordinateSystem .
Remarks: The moment of inertia is calculated with respect to the output coordinate system. Before calling this method, call ICWMassPropertiesManager::SetCoordinateSystemName or ICWMassPropertiesManager::SetCoordinateSystemToDefault . The moment of inertia components are in units of kilograms * meters squared. This method returns nine doubles as follows: [ lxx , lxy , lxz , l yx , lyy , lyz , lzx , lzy , lzz ]
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.GetPrincipalAxesOfInertia
Gets the principal axes of inertia with respect to the center of mass.
```vb
Function GetPrincipalAxesOfInertia() As System.Object
```
Return Value 2D array of vector coefficients of the three principal axes of inertia (see Remarks )
VBA: See CWMassPropertiesManager::GetPrincipalAxesOfInertia .
Remarks: A principal axis is one of three mutually perpendicular axes in a body about which the moment of inertia is at a maximum. This method returns nine doubles as follows: [ axis1_lx , axis1_ly , axis1_lz , axis2_lx , axis2_ly , axis2_lz , axis3_lx , axis3_ly , axis3_lz ]
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.GetPrincipalMomentOfInertia
Gets the moments of inertia with respect to the principal axes and the center of mass.
```vb
Function GetPrincipalMomentOfInertia() As System.Object
```
Return Value Array of three principal moments of inertia (see Remarks )
VBA: See CWMassPropertiesManager::GetPrincipalMomentOfInertia .
Remarks: The moment of inertia is calculated with respect to the output coordinate system. Before calling this method, call ICWMassPropertiesManager::SetCoordinateSystemName or ICWMassPropertiesManager::SetCoordinateSystemToDefault . The moments of inertia are in units of kilograms * meters squared. This method returns three doubles as follows: [ Px , Py , Pz ]
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.Mass
Gets the total mass.
```vb
ReadOnly Property Mass As System.Double
```
Property Value Mass in kilograms
VBA: See CWMassPropertiesManager::Mass .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveAllBodies
Removes all bodies from the calculation of mass properties.
```vb
Function RemoveAllBodies() As System.Integer
```
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveAllBodies .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveAllBoltConnectors
Removes all bolt connectors from the calculation of mass properties.
```vb
Function RemoveAllBoltConnectors() As System.Integer
```
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveAllBoltConnectors .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveAllDistributedMasses
Removes all distributed masses from the calculation of mass properties.
```vb
Function RemoveAllDistributedMasses() As System.Integer
```
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveAllDistributedMasses .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveAllLinkageRods
Removes all linkage rod connectors.
```vb
Function RemoveAllLinkageRods() As System.Integer
```
Return Value Error code as defined by swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveAllLinkageRods .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveAllPinConnectors
Removes all pin connectors from the calculation of mass properties.
```vb
Function RemoveAllPinConnectors() As System.Integer
```
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveAllPinConnectors .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveAllRemoteLoadMasses
Removes all remote load masses from the calculation of mass properties.
```vb
Function RemoveAllRemoteLoadMasses() As System.Integer
```
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveAllRemoteLoadMasses .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveBodies
Removes the specified bodies from the calculation of mass properties.
```vb
Function RemoveBodies( ByVal DispArray As System.Object ) As System.Integer
```
Parameters DispArray Array of bodies to remove Return Value Error code as defined in swsSimMassPropertiesError_e
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveBodies .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveBoltConnectors
Removes the specified bolt connectors from the calculation of mass properties.
```vb
Function RemoveBoltConnectors( ByVal NameArray As System.Object ) As System.Integer
```
Parameters NameArray Array of the names of bolt connectors to remove Return Value Error code as defined in swsSimMassPropertiesError_e
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveBoltConnectors .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveDistributedMasses
Removes the specified distributed masses from the calculation of mass properties.
```vb
Function RemoveDistributedMasses( ByVal NameArray As System.Object ) As System.Integer
```
Parameters NameArray Array of the names of distributed masses to remove Return Value Error code as defined in swsSimMassPropertiesError_e
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveDistributedMasses .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveLinkageRods
Removes the specified linkage rod connectors.
```vb
Function RemoveLinkageRods( ByVal NamesArray As System.Object ) As System.Integer
```
Parameters NamesArray Array of the names of linkage rod connectors to remove Return Value Error code as defined by swsSimMassPropertiesError_e
Return Value Error code as defined by swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveLinkageRods .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemovePinConnectors
Removes the specified pin connectors from the calculation of mass properties.
```vb
Function RemovePinConnectors( ByVal NameArray As System.Object ) As System.Integer
```
Parameters NameArray Array of the names of pin connectors to remove Return Value Error code as defined in swsSimMassPropertiesError_e
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemovePinConnectors .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.RemoveRemoteLoadMasses
Removes the specified remote load masses from the calculation of mass properties.
```vb
Function RemoveRemoteLoadMasses( ByVal NameArray As System.Object ) As System.Integer
```
Parameters NameArray Array of the names of remote load masses to remove Return Value Error code as defined in swsSimMassPropertiesError_e
Return Value Error code as defined in swsSimMassPropertiesError_e
VBA: See CWMassPropertiesManager::RemoveRemoteLoadMasses .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.SetCoordinateSystemName
Applies the specified coordinate system.
```vb
Sub SetCoordinateSystemName( ByVal SCoordinateSystemName As System.String )
```
Parameters SCoordinateSystemName Name of coordinate system
VBA: See CWMassPropertiesManager::SetCoordinateSystemName .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.SetCoordinateSystemToDefault
Applies the default coordinate system.
```vb
Sub SetCoordinateSystemToDefault()
```
VBA: See CWMassPropertiesManager::SetCoordinateSystemToDefault .
Examples: Visual Basic for Applications

### ICWMassPropertiesManager.Volume
Gets the total volume.
```vb
ReadOnly Property Volume As System.Double
```
Property Value Volume in cubic meters
VBA: See CWMassPropertiesManager::Volume .
Examples: Visual Basic for Applications

## ICWRunSpecStudiesRunMeshOptions

### ICWRunSpecStudiesRunMeshOptions.AddStudyOption
Adds the specified run and mesh option to the specified study.
```vb
Function AddStudyOption( ByVal SStudyName As System.String, ByVal NRunMeshOption As System.Integer ) As System.Integer
```
Parameters SStudyName Name of study NRunMeshOption Run and mesh option as defined in swsRunStudiesRunMeshOptions_e Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
VBA: See CWRunSpecStudiesRunMeshOptions::AddStudyOption .
Examples: Visual Basic for Applications

### ICWRunSpecStudiesRunMeshOptions.GetFirstItem
Gets the first item in the run studies batch.
```vb
Function GetFirstItem( ByRef SStudyName As System.String, ByRef NRunMeshOption As System.Integer ) As System.Integer
```
Parameters SStudyName Name of study NRunMeshOption Run and mesh option as defined in swsRunStudiesRunMeshOptions_e Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
VBA: See CWRunSpecStudiesRunMeshOptions::GetFirstItem .
Examples: Visual Basic for Applications

### ICWRunSpecStudiesRunMeshOptions.GetNextItem
Gets the next item in the run studies batch.
```vb
Function GetNextItem( ByRef SStudyName As System.String, ByRef NRunMeshOption As System.Integer ) As System.Integer
```
Parameters SStudyName Name of study NRunMeshOption Run and mesh option as defined in swsRunStudiesRunMeshOptions_e Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
VBA: See CWRunSpecStudiesRunMeshOptions::GetNextItem .
Examples: Visual Basic for Applications

### ICWRunSpecStudiesRunMeshOptions.RemoveAll
Clears the option in the run studies batch.
```vb
Function RemoveAll() As System.Integer
```
Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
VBA: See CWRunSpecStudiesRunMeshOptions::RemoveAll .
Examples: Visual Basic for Applications

### ICWRunSpecStudiesRunMeshOptions.RemoveStudy
Removes the specified study from the run studies batch.
```vb
Function RemoveStudy( ByVal SStudyName As System.String ) As System.Integer
```
Parameters SStudyName Name of study to remove Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
VBA: See CWRunSpecStudiesRunMeshOptions::RemoveStudy .
Examples: Visual Basic for Applications

### ICWRunSpecStudiesRunMeshOptions.ReplaceOption
Replaces the run and mesh option of the specified study with a new option.
```vb
Function ReplaceOption( ByVal SStudyName As System.String, ByVal NRunMeshOption As System.Integer ) As System.Integer
```
Parameters SStudyName Name of study NRunMeshOption New run and mesh option as defined in swsRunStudiesRunMeshOptions_e Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
Return Value Error code as defined in swsRunStudiesRunMeshOptionErrorCode_e
VBA: See CWRunSpecStudiesRunMeshOptions::ReplaceOption .
Examples: Visual Basic for Applications

## ICWRunStudiesResults

### ICWRunStudiesResults.GetFirstItem
Gets the first study and result code in the run studies batch.
```vb
Function GetFirstItem( ByRef SStudyName As System.String, ByRef NRunStatus As System.Integer ) As System.Integer
```
Parameters SStudyName Name of study NRunStatus Error code of SStudyName as defined in swsRunStudiesStatusCode_e Return Value Error code as defined in swsRunStudiesResultsErrorCode_e
Return Value Error code as defined in swsRunStudiesResultsErrorCode_e
VBA: See CWRunStudiesResults::GetFirstItem .
Examples: Visual Basic for Applications

### ICWRunStudiesResults.GetNextItem
Gets the next study and result code in the run studies batch.
```vb
Function GetNextItem( ByRef SStudyName As System.String, ByRef NRunStatus As System.Integer ) As System.Integer
```
Parameters SStudyName Name of study NRunStatus Error code of SStudyName as defined in swsRunStudiesStatusCode_e Return Value Error code as defined in swsRunStudiesResultsErrorCode_e
Return Value Error code as defined in swsRunStudiesResultsErrorCode_e
VBA: See CWRunStudiesResults::GetNextItem .
Examples: Visual Basic for Applications

