Option Strict Off
Option Explicit On

Imports System
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

' Excel by late binding (no Office references), same as the VBA macros.
Public Module ExcelLink

    ' Tabs the macros look up, by the words after the tab number.
    Private ReadOnly NeededTabs As String() = {
        "Component register", "STEP import", "CAD geometry", "Settings",
        "Spreader", "Thermal solution", "Sim setup", "Sim results", "Change log"}

    Public Function GetExcel() As Object
        Try
            Return Marshal.GetActiveObject("Excel.Application")
        Catch
            Return Nothing
        End Try
    End Function

    Public Function ActiveWorkbook() As Object
        Dim xl As Object = GetExcel()
        If xl Is Nothing Then Return Nothing
        Try
            Return xl.ActiveWorkbook
        Catch
            Return Nothing
        End Try
    End Function

    ' The workbook the steps will write to (PT_CurrentWorkbook), or Nothing when they will ask for one.
    Public Function StepsWorkbook() As Object
        Return PTCommon.PT_CurrentWorkbook(GetExcel())
    End Function

    ' Text for the task pane and Check when the steps will ask: names the active workbook so it is not mistaken for the used one.
    Public Function WillAskText() As String
        Dim s As String = "(will ask)"
        Dim act As Object = ActiveWorkbook()
        If act IsNot Nothing Then
            Try
                s &= ", active in Excel: " & CStr(act.Name)
            Catch
            End Try
        End If
        Return s
    End Function

    ' Same rule as PT_TabBase in PTCommon.bas: drop the step number ("7", "7b") and tidy spaces.
    Public Function TabBase(ByVal tabName As String) As String
        Dim s As String = tabName.Replace(ChrW(160), " ").Replace(vbTab, " ")
        s = Regex.Replace(s, " {2,}", " ").Trim()
        s = Regex.Replace(s, "^[0-9]+([A-Za-z](?=[ ._)\-]))?", "")
        s = s.TrimStart(" "c, "."c, "_"c, ")"c, "-"c)
        Return s.Trim()
    End Function

    Public Function FindSheet(ByVal wb As Object, ByVal baseName As String) As Object
        For Each ws As Object In wb.Worksheets
            If String.Compare(TabBase(CStr(ws.Name)), baseName, StringComparison.OrdinalIgnoreCase) = 0 Then Return ws
        Next
        Return Nothing
    End Function

    Public Sub CheckWorkbook()
        Dim wb As Object = StepsWorkbook()
        If wb Is Nothing Then
            Log.Write("Check workbook: " & WillAskText() & ". The next step asks which workbook to use.")
            Return
        End If
        Log.Write("Check workbook: " & CStr(wb.Name))
        Dim missing As Integer = 0
        For Each tabName As String In NeededTabs
            Dim ws As Object = FindSheet(wb, tabName)
            If ws Is Nothing Then
                Log.Write("  missing tab: " & tabName)
                missing += 1
            End If
        Next
        Dim nameCount As Integer = 0
        Try
            nameCount = CInt(wb.Names.Count)
        Catch
        End Try
        Dim reg As Object = FindSheet(wb, "Component register")
        Dim rows As String = "?"
        If reg IsNot Nothing Then
            Try
                rows = CStr(reg.UsedRange.Rows.Count)
            Catch
            End Try
        End If
        Log.Write("  tabs found: " & (NeededTabs.Length - missing).ToString() & " of " & NeededTabs.Length.ToString() &
                  ", defined names: " & nameCount.ToString() & ", register used rows: " & rows)
    End Sub

End Module
