Option Strict Off
Option Explicit On

Imports System
Imports System.Diagnostics
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

' Counts components, solid bodies, faces and triangles in the active model and times it.
' The triangle loop is the same kind of work ImportStep does, so the time shows in-process API speed.
Public Module ModelCheck

    Public Sub Run(ByVal swApp As SldWorks)
        Dim doc As ModelDoc2 = CType(swApp.ActiveDoc, ModelDoc2)
        If doc Is Nothing Then
            Log.Write("Check model: no open document.")
            Return
        End If
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Dim comps As Integer = 0, bodies As Integer = 0, faces As Integer = 0
        Dim tris As Long = 0
        Dim docType As Integer = doc.GetType()

        If docType = CInt(swDocumentTypes_e.swDocASSEMBLY) Then
            Dim asm As AssemblyDoc = CType(doc, AssemblyDoc)
            asm.ResolveAllLightWeightComponents(False)
            Dim list As Object = asm.GetComponents(False)
            If list IsNot Nothing Then
                For Each o As Object In CType(list, Object())
                    Dim comp As Component2 = CType(o, Component2)
                    comps += 1
                    If comp.IsSuppressed() Then Continue For
                    Dim info As Object = Nothing
                    CountBodies(comp.GetBodies3(CInt(swBodyType_e.swSolidBody), info), bodies, faces, tris)
                Next
            End If
        ElseIf docType = CInt(swDocumentTypes_e.swDocPART) Then
            Dim part As PartDoc = CType(doc, PartDoc)
            CountBodies(part.GetBodies2(CInt(swBodyType_e.swSolidBody), True), bodies, faces, tris)
        Else
            Log.Write("Check model: open a part or an assembly.")
            Return
        End If

        sw.Stop()
        Log.Write("Check model: " & doc.GetTitle())
        Log.Write("  components " & comps.ToString() & ", solid bodies " & bodies.ToString() &
                  ", faces " & faces.ToString() & ", triangles " & tris.ToString())
        Log.Write("  time " & (sw.ElapsedMilliseconds / 1000.0).ToString("0.00") & " s")
    End Sub

    Private Sub CountBodies(ByVal arr As Object, ByRef bodies As Integer, ByRef faces As Integer, ByRef tris As Long)
        If arr Is Nothing Then Return
        For Each o As Object In CType(arr, Object())
            Dim b As Body2 = CType(o, Body2)
            bodies += 1
            Dim f As Face2 = CType(b.GetFirstFace(), Face2)
            Do While f IsNot Nothing
                faces += 1
                Dim t As Object = f.GetTessTriangles(True)
                If t IsNot Nothing Then tris += CType(t, Single()).Length \ 9
                f = CType(f.GetNextFace(), Face2)
            Loop
        Next
    End Sub

End Module
