Attribute VB_Name = "BuildSolution"
Option Explicit
' =====================================================================================
' BuildSolution  -  SolidWorks macro, the Build Solution button (Ansys page route, ansys\page.html)
'
' Pick study.json of the board folder. Step 4 of the page (Thermal solution) says which parts the plate touches,
' its thickness, Lift and TIMs. This makes <folder>\solidworks\<board>_solution.sldasm from the board STEP (first
' run only), then the plates Spreader_Top / Spreader_Bottom (Outline sketch, plate, pedestals) and one TIM part per
' ticked part. Run it again after a change on the page: the plates move and the pedestals and TIMs follow, while the
' Outline sketch and every feature you added (holes, cuts, fillets) stay. Then press Send Solution.
' The work is done by BuildSpreader.RunAnsys. Needs PTCommon, PTJson and BuildSpreader in the same macro.
' =====================================================================================

Sub main()
    RunAnsys
End Sub
