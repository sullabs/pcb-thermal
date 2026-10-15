Option Strict Off
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic

' Port of src/PTCommon.bas, part 1: Excel and workbook, values, keys, change log, files.
' Same procedure names and behaviour as the VBA, so each one can be checked against it.
' Excel by late binding (Option Strict Off), like the macros. Cell positions come from workbook names.
' Part 2 (SolidWorks helpers, PTCommon.bas from "properties" on) follows with the macros that need it.
Public Module PTCommon

    Public Const PT_APP As String = "PCBThermal"
    Public Const PT_LOG_SHEET As String = "Change log"
    Public Const PT_LOG_COLS As Integer = 11     ' Date Macro Key Name RefDes Change Field Old New Delta Checked

    ' Set by the add-in on connect (AddIn.ConnectToSW).
    Public swApp As Object
    Public PT_LastErr As String = ""

    Private logRows As New List(Of Object())
    Private ReadOnly reInst As New Regex("(-\d+|<\d+>|\s#\d+)$")
    Private ReadOnly reDup As New Regex("_\d+(?=\.(step|stp)(?![a-z0-9]))")
    Private ReadOnly reExt As New Regex("\.(step|stp|sldprt|sldasm|x_t|igs|iges)$")
    Private ReadOnly reJunk As New Regex("[^a-z0-9]")
    Private ReadOnly reTail As New Regex("(?:-(\d+)|<(\d+)>|\s#(\d+))$")

    ' ---------------------------------------------------------------- Excel and workbook
    Public Function PT_GetExcel() As Object
        Dim xl As Object = Nothing
        Try
            xl = Marshal.GetActiveObject("Excel.Application")
        Catch
        End Try
        If xl Is Nothing Then xl = CreateObject("Excel.Application")
        xl.Visible = True
        Return xl
    End Function

    ' Asks which workbook to use (starts at the last one used), opens it if needed and checks the tabs in
    ' needSheets ("Tab one|Tab two"). Returns Nothing if the user cancels or a check fails (the user has been told why).
    ' Shares the "last workbook" setting with the VBA macros (GetSetting/SaveSetting, same registry key).
    Public Function PT_PickWorkbook(ByVal xl As Object, ByVal macroName As String, ByVal needSheets As String) As Object
        Dim wbk As Object = Nothing
        Dim lastF As String = GetSetting(PT_APP, "Common", "Workbook", "")
        If lastF = "" Then lastF = GetSetting(PT_APP, "ImportStep", "Workbook", "")
        ' the workbook used last time is still open in Excel: use it without asking (close it to pick another one)
        If lastF <> "" Then wbk = FindOpenWorkbook(xl, lastF)
        If wbk Is Nothing Then
            Dim f As String = SwGetOpenFileName("Pick the project workbook (a copy of PCB_thermal_TEMPLATE)", lastF,
                                                "Excel workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm||")
            If String.IsNullOrEmpty(f) Then Return Nothing
            wbk = FindOpenWorkbook(xl, f)
            If wbk Is Nothing Then wbk = xl.Workbooks.Open(f)
        End If
        Dim full As String = CStr(wbk.FullName)
        If CBool(wbk.ReadOnly) Then
            MsgBox("That workbook is open read-only (maybe open on another PC or in another Excel window)." & vbCrLf &
                   "Close it there and run again.", MsgBoxStyle.Exclamation, macroName)
            Return Nothing
        End If
        PT_DailyBackup(wbk)
        Dim missing As String = ""
        For Each s As String In needSheets.Split("|"c)
            If Not PT_HasSheet(wbk, s) Then missing &= vbCrLf & "  " & s
        Next
        If missing <> "" Then
            MsgBox("That workbook is missing these tabs:" & missing & vbCrLf & vbCrLf &
                   "Tabs found in " & CStr(wbk.Name) & ":" & PT_TabList(wbk) & vbCrLf & vbCrLf &
                   "Pick a copy of PCB_thermal_TEMPLATE.xlsx.", MsgBoxStyle.Exclamation, macroName)
            Return Nothing
        End If
        SaveSetting(PT_APP, "Common", "Workbook", full)
        Return wbk
    End Function

    Private Function FindOpenWorkbook(ByVal xl As Object, ByVal fullName As String) As Object
        For Each wbk As Object In xl.Workbooks
            If String.Compare(CStr(wbk.FullName), fullName, StringComparison.OrdinalIgnoreCase) = 0 Then Return wbk
        Next
        Return Nothing
    End Function

    ' One copy of the workbook per day in the folder "backup" next to it (<name>_yyyymmdd.xlsx), made before the first
    ' macro of the day writes to it. Never stops a macro: a failed copy is ignored.
    Public Sub PT_DailyBackup(ByVal wb As Object)
        Try
            Dim p As String = CStr(wb.Path)
            If p = "" Then Return
            Dim dirS As String = Path.Combine(p, "backup")
            PT_MakeDirs(dirS)
            Dim nm As String = CStr(wb.Name)
            Dim dotP As Integer = nm.LastIndexOf("."c)
            Dim stamp As String = DateTime.Now.ToString("yyyyMMdd")
            Dim target As String
            If dotP > 0 Then
                target = Path.Combine(dirS, nm.Substring(0, dotP) & "_" & stamp & nm.Substring(dotP))
            Else
                target = Path.Combine(dirS, nm & "_" & stamp)
            End If
            If Not File.Exists(target) Then wb.SaveCopyAs(target)
        Catch
        End Try
    End Sub

    ' Keys ("key|text") of the Change log rows whose Change column is changeText and whose Checked cell is not Yes.
    ' SyncFromCAD uses it so a check that failed before and is still unchecked is not logged a second time.
    Public Function PT_OpenLogKeys(ByVal wb As Object, ByVal changeText As String) As Dict
        Dim d As New Dict()
        Dim hdr As Object = PT_Rng(wb, "Log_Header")
        If hdr Is Nothing Then Return d
        Dim ws As Object = hdr.Worksheet
        Dim lastR As Integer = CInt(ws.Cells(ws.Rows.Count, hdr.Column + 5).End(-4162).Row)
        If lastR <= CInt(hdr.Row) Then Return d
        Dim v As Object(,) = ToTable(ws.Range(ws.Cells(hdr.Row + 1, hdr.Column), ws.Cells(lastR, hdr.Column + PT_LOG_COLS - 1)).Value)
        For r As Integer = 1 To v.GetUpperBound(0)
            If PT_Str(v(r, 6)) = changeText AndAlso PT_Str(v(r, 11)).ToUpperInvariant() <> "YES" Then
                Dim k As String = PT_Str(v(r, 3)) & "|" & PT_Str(v(r, 7))
                If Not d.Exists(k) Then d.Add(k, 1)
            End If
        Next
        Return d
    End Function

    ' True when every workbook name in needNames ("A|B|C") exists. Otherwise tells the user.
    Public Function PT_NeedNames(ByVal wb As Object, ByVal needNames As String, ByVal macroName As String) As Boolean
        Dim missing As String = ""
        For Each s As String In needNames.Split("|"c)
            If PT_Rng(wb, s) Is Nothing Then missing &= " " & s
        Next
        If missing <> "" Then
            MsgBox(macroName & " did not run. Nothing was changed." & vbCrLf & vbCrLf &
                   "The workbook has no cells named:" & missing & vbCrLf & vbCrLf &
                   "Use a copy of PCB_thermal_TEMPLATE.xlsx. For an older workbook, run SetupWorkbook (Prepare Workbook) once, " &
                   "then run " & macroName & " again.", MsgBoxStyle.Exclamation, macroName)
            Return False
        End If
        Return True
    End Function

    Public Function PT_HasSheet(ByVal wb As Object, ByVal sheetName As String) As Boolean
        Return PT_Sheet(wb, sheetName) IsNot Nothing
    End Function

    ' Tab by its base name. The template puts a step number in front ("7 Spreader"), older copies do not ("Spreader").
    ' Also accepted: "7. Spreader", "07_Spreader", "7-Spreader", "7)Spreader", a non-breaking space, double or trailing
    ' spaces, any letter case. Returns Nothing if there is no such tab.
    Public Function PT_Sheet(ByVal wb As Object, ByVal baseName As String) As Object
        Try
            Dim ws As Object = wb.Worksheets(baseName)
            If ws IsNot Nothing Then Return ws
        Catch
        End Try
        Dim want As String = PT_TabBase(baseName)
        For Each ws As Object In wb.Worksheets
            If String.Compare(PT_TabBase(CStr(ws.Name)), want, StringComparison.OrdinalIgnoreCase) = 0 Then Return ws
        Next
        Return Nothing
    End Function

    ' Tab name without its leading step number and separators, spaces cleaned up.
    Public Function PT_TabBase(ByVal tabName As String) As String
        Dim s As String = tabName.Replace(ChrW(160), " ").Replace(vbTab, " ")
        Do While s.Contains("  ")
            s = s.Replace("  ", " ")
        Loop
        s = s.Trim()
        Dim k As Integer = 0
        Do While k < s.Length AndAlso Char.IsDigit(s(k)) AndAlso s(k) <= "9"c
            k += 1
        Loop
        If k > 0 Then
            ' one letter after the number is part of the step ("7b Spreader bottom")
            If k < s.Length - 1 Then
                If IsAsciiLetter(s(k)) AndAlso " ._)-".IndexOf(s(k + 1)) >= 0 Then k += 1
            End If
            Do While k < s.Length AndAlso " .-_)".IndexOf(s(k)) >= 0
                k += 1
            Loop
            s = s.Substring(k)
        End If
        Return s
    End Function

    Private Function IsAsciiLetter(ByVal c As Char) As Boolean
        Return (c >= "A"c AndAlso c <= "Z"c) OrElse (c >= "a"c AndAlso c <= "z"c)
    End Function

    ' Names of all tabs, one per line, for error messages.
    Public Function PT_TabList(ByVal wb As Object) As String
        Dim s As String = ""
        For Each ws As Object In wb.Worksheets
            s &= vbCrLf & "  [" & CStr(ws.Name) & "]"
        Next
        Return s
    End Function

    ' Range behind a workbook name, or Nothing.
    Public Function PT_Rng(ByVal wb As Object, ByVal nameS As String) As Object
        Try
            Return wb.Names(nameS).RefersToRange
        Catch
            Return Nothing
        End Try
    End Function

    Public Function PT_NameNum(ByVal wb As Object, ByVal nameS As String, ByVal dflt As Double) As Double
        Dim r As Object = PT_Rng(wb, nameS)
        If r Is Nothing Then Return dflt
        Return PT_Num(r.Cells(1, 1).Value, dflt)
    End Function

    Public Function PT_NameStr(ByVal wb As Object, ByVal nameS As String, ByVal dflt As String) As String
        Dim r As Object = PT_Rng(wb, nameS)
        Dim s As String = ""
        If r IsNot Nothing Then s = PT_Str(r.Cells(1, 1).Value)
        If s = "" Then s = dflt
        Return s
    End Function

    ' Fills the chip rows of the Spreader tab (Spr_Chips) and the Sim setup tab (Sim_Chips) from the RefDes column of
    ' CAD geometry, one row per included RefDes. Only a table that is still empty is filled, so typed rows are never
    ' overwritten. Returns the number of rows written.
    Public Function PT_FillChipLists(ByVal wb As Object) As Integer
        Dim rSpr As Object = PT_Rng(wb, "Spr_Chips")
        Dim rSim As Object = PT_Rng(wb, "Sim_Chips")
        If PT_Rng(wb, "Geo_Table") Is Nothing Then Return 0
        Dim doSpr As Boolean = False, doSim As Boolean = False
        If rSpr IsNot Nothing Then doSpr = (CDbl(wb.Application.WorksheetFunction.CountA(rSpr.Columns(1))) = 0)
        If rSim IsNot Nothing Then doSim = (CDbl(wb.Application.WorksheetFunction.CountA(rSim.Columns(1))) = 0)
        If Not (doSpr OrElse doSim) Then Return 0
        Dim geo As Object(,) = ToTable(PT_Rng(wb, "Geo_Table").Value)   ' column 1 = B; 5 RefDes, 6 Include, 13 side, 22 Z min, 25 H
        Dim tPcb As Double = PT_Num(PT_Rng(wb, "Geo_PCB_Used").Cells(3, 1).Value, 0)
        Dim seen As New HashSet(Of String)()
        Dim n As Integer = 0
        For r As Integer = 1 To geo.GetUpperBound(0)
            Dim refD As String = PT_Str(geo(r, 5))
            If refD <> "" AndAlso PT_Str(geo(r, 6)).ToUpperInvariant() = "YES" AndAlso Not seen.Contains(refD.ToUpperInvariant()) Then
                seen.Add(refD.ToUpperInvariant())
                If doSpr AndAlso n >= CInt(rSpr.Rows.Count) Then Exit For
                If doSim AndAlso n >= CInt(rSim.Rows.Count) Then Exit For
                Dim sideS As String = PT_Str(geo(r, 13))
                Dim zmin As Double = PT_Num(geo(r, 22), 0)
                Dim hgt As Double = PT_Num(geo(r, 25), 99)
                Dim onTop As Boolean
                If sideS = "" Then onTop = (zmin >= tPcb - 0.01) Else onTop = (sideS.ToUpperInvariant() = "TOP")
                If doSpr Then
                    rSpr.Cells(n + 1, 1).Value = refD
                    ' low parts on the top side are covered by default; tall ones (connectors, big inductors) are not
                    rSpr.Cells(n + 1, 8).Value = If(onTop AndAlso hgt <= 3.5, "Yes", "No")
                    rSpr.Cells(n + 1, 9).Value = "Pedestal"
                    rSpr.Cells(n + 1, 10).Value = "Body"
                    rSpr.Cells(n + 1, 11).Value = 0.2
                    rSpr.Cells(n + 1, 12).Value = 3
                End If
                If doSim Then
                    rSim.Cells(n + 1, 1).Value = refD
                    rSim.Cells(n + 1, 2).Value = "Block"
                End If
                n += 1
            End If
        Next
        Return n
    End Function

    ' Range.Value as a 1-based 2D array, also for a single cell (Excel then returns the value itself).
    Public Function ToTable(ByVal v As Object) As Object(,)
        Dim t As Object(,) = TryCast(v, Object(,))
        If t IsNot Nothing Then Return t
        Dim one As Object(,) = CType(Array.CreateInstance(GetType(Object), New Integer() {1, 1}, New Integer() {1, 1}), Object(,))
        one(1, 1) = v
        Return one
    End Function

    ' ---------------------------------------------------------------- values
    ' Excel error cells (#N/A, #VALUE! ...) arrive in .NET as Int32 codes; real numbers come as Double.
    Public Function PT_IsError(ByVal v As Object) As Boolean
        Return TypeOf v Is Integer OrElse TypeOf v Is ErrorWrapper
    End Function

    Public Function PT_IsNum(ByVal v As Object) As Boolean
        If v Is Nothing OrElse TypeOf v Is DBNull OrElse PT_IsError(v) Then Return False
        If TypeOf v Is String Then Return False
        Return IsNumeric(v)
    End Function

    Public Function PT_Num(ByVal v As Object, ByVal dflt As Double) As Double
        If v Is Nothing OrElse TypeOf v Is DBNull OrElse PT_IsError(v) Then Return dflt
        If TypeOf v Is String Then
            If CStr(v).Trim() = "" Then Return dflt
        End If
        If IsNumeric(v) Then Return CDbl(v)
        Return dflt
    End Function

    Public Function PT_Str(ByVal v As Object) As String
        If v Is Nothing OrElse TypeOf v Is DBNull OrElse PT_IsError(v) Then Return ""
        Return CStr(v).Trim()
    End Function

    ' ---------------------------------------------------------------- keys
    ' Matching key of a model name: last path segment, lower case, no instance suffix (-1, <1>, #2),
    ' no file extension (.STEP), letters and digits only.
    Public Function NormKey(ByVal s As String) As String
        s = s.Substring(s.LastIndexOf("/"c) + 1)
        s = reDup.Replace(s.Trim().ToLowerInvariant(), "")   ' SolidWorks adds _1 to a STEP name that is already open in the session
        s = reInst.Replace(s, "")
        s = reExt.Replace(s, "")
        Return reJunk.Replace(s, "")
    End Function

    ' Identity key of a body: NormKey plus "_<n>" when the instance number is above 1,
    ' so two copies of the same part (for example two identical connectors) get different keys.
    Public Function PT_Key(ByVal s As String) As String
        Dim seg As String = s.Substring(s.LastIndexOf("/"c) + 1).Trim()
        Dim inst As Long = 0
        Dim m As Match = reTail.Match(seg)
        If m.Success Then
            For k As Integer = 1 To 3
                If m.Groups(k).Success AndAlso m.Groups(k).Value <> "" Then inst = CLng(m.Groups(k).Value)
            Next
        End If
        Dim key As String = NormKey(s)
        If key = "" Then key = "body"
        If inst > 1 Then key &= "_" & inst.ToString()
        Return key
    End Function

    ' Returns k, or k plus "_d2", "_d3" ... if k is already in the dictionary. Adds the result.
    Public Function PT_UniqueKey(ByVal used As Dict, ByVal k As String) As String
        Dim s As String = k
        If used.Exists(s) Then
            Dim i As Integer = 2
            Do While used.Exists(k & "_d" & i.ToString())
                i += 1
            Loop
            s = k & "_d" & i.ToString()
        End If
        used.Add(s, 1)
        Return s
    End Function

    ' ---------------------------------------------------------------- change log
    Public Sub PT_LogReset()
        logRows = New List(Of Object())()
    End Sub

    Public Sub PT_LogAdd(ByVal macroName As String, ByVal key As String, ByVal nameS As String, ByVal refDes As String,
                         ByVal change As String, ByVal field As String, ByVal oldV As Object, ByVal newV As Object)
        Dim dl As Object = ""
        If PT_IsNum(oldV) AndAlso PT_IsNum(newV) Then dl = Math.Round(CDbl(newV) - CDbl(oldV), 4)
        logRows.Add(New Object() {DateTime.Now, macroName, key, nameS, refDes, change, field, oldV, newV, dl, ""})
    End Sub

    Public Function PT_LogCount() As Integer
        Return logRows.Count
    End Function

    ' Writes the collected rows to the top of the Change log tab (newest on top). Returns the count.
    Public Function PT_LogFlush(ByVal wb As Object) As Integer
        Dim n As Integer = logRows.Count
        If n = 0 Then Return 0
        Dim hdr As Object = PT_Rng(wb, "Log_Header")
        Dim ws As Object = hdr.Worksheet
        Dim r0 As Integer = CInt(hdr.Row) + 1, c0 As Integer = CInt(hdr.Column)
        Dim outv(n - 1, PT_LOG_COLS - 1) As Object
        For i As Integer = 0 To n - 1
            For k As Integer = 0 To PT_LOG_COLS - 1
                outv(i, k) = logRows(i)(k)
            Next
        Next
        ws.Rows(r0.ToString() & ":" & (r0 + n - 1).ToString()).Insert(-4121, 1)       ' shift down, format from below
        ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + n - 1, c0 + PT_LOG_COLS - 1)).Value = outv
        ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + n - 1, c0)).NumberFormat = "yyyy-mm-dd hh:mm"
        Dim val As Object = ws.Range(ws.Cells(r0, c0 + PT_LOG_COLS - 1), ws.Cells(r0 + n - 1, c0 + PT_LOG_COLS - 1)).Validation
        val.Delete()
        val.Add(3, 1, 1, "=Lst_YesNo")
        logRows = New List(Of Object())()
        Return n
    End Function

    ' Build macros call this first. True = refuse (the user has been told which rows to check).
    Public Function PT_RefuseIfUnchecked(ByVal wb As Object, ByVal macroName As String) As Boolean
        wb.Application.Calculate()
        Dim n As Double = PT_NameNum(wb, "Log_Unchecked", 0)
        If n <= 0 Then Return False
        Dim hdr As Object = PT_Rng(wb, "Log_Header")
        Dim ws As Object = hdr.Worksheet
        Dim cChange As Integer = CInt(hdr.Column) + 5, cChecked As Integer = CInt(hdr.Column) + 10
        Dim lastR As Integer = CInt(ws.Cells(ws.Rows.Count, cChange).End(-4162).Row)
        Dim s As String = "", shown As Integer = 0
        For r As Integer = CInt(hdr.Row) + 1 To lastR
            If PT_Str(ws.Cells(r, cChange).Value) <> "" AndAlso PT_Str(ws.Cells(r, cChecked).Value).ToUpperInvariant() <> "YES" Then
                s &= "  row " & r.ToString() & ": " & PT_Str(ws.Cells(r, cChange).Value) & " " &
                     PT_Str(ws.Cells(r, cChange + 1).Value) & "  " & PT_Str(ws.Cells(r, hdr.Column + 3).Value) & vbCrLf
                shown += 1
                If shown >= 8 Then
                    s &= "  ..." & vbCrLf
                    Exit For
                End If
            End If
        Next
        MsgBox(macroName & " did not run. Nothing was changed." & vbCrLf & vbCrLf &
               n.ToString() & " rows on the Change log tab are not checked:" & vbCrLf & s & vbCrLf &
               "Look at each one, type Yes in its Checked cell, save, and run again.", MsgBoxStyle.Exclamation, macroName)
        Return True
    End Function

    ' ---------------------------------------------------------------- files
    ' Output folder from Settings (blank = CAD folder next to the workbook), created if missing, ends with "\".
    Public Function PT_OutputFolder(ByVal wb As Object) As String
        Dim s As String = PT_NameStr(wb, "Set_OutputFolder", "")
        If s = "" Then s = CStr(wb.Path) & "\CAD"
        s = s.TrimEnd("\"c)
        PT_MakeDirs(s)
        Return s & "\"
    End Function

    ' File prefix from Settings. Blank = the workbook file name, e.g. "Board.xlsx" -> "Board_".
    Public Function PT_Prefix(ByVal wb As Object) As String
        Dim s As String = PT_NameStr(wb, "Set_Prefix", "")
        If s = "" Then
            s = CStr(wb.Name)
            If s.LastIndexOf("."c) > 0 Then s = s.Substring(0, s.LastIndexOf("."c))
            s = PT_SafeFile(s) & "_"
        End If
        Return s
    End Function

    Public Sub PT_MakeDirs(ByVal p As String)
        If p = "" Then Return
        Directory.CreateDirectory(p)
    End Sub

    ' mm -> whole micrometres as text (no decimal separator, so the Windows language does not matter)
    Public Function PT_Um(ByVal mm As Double) As String
        Return CLng(Math.Round(mm * 1000.0)).ToString()
    End Function

    Public Function PT_SafeFile(ByVal s As String) As String
        Dim sb As New System.Text.StringBuilder()
        For Each ch As Char In s
            If IsAsciiLetter(ch) OrElse (ch >= "0"c AndAlso ch <= "9"c) OrElse ch = "_"c OrElse ch = "-"c Then sb.Append(ch) Else sb.Append("_"c)
        Next
        Dim outS As String = sb.ToString()
        If outS.Length > 60 Then outS = outS.Substring(0, 60)
        Return outS
    End Function

    Public Function PT_BoxText(ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double,
                               ByVal lx As Double, ByVal wy As Double, ByVal hz As Double) As String
        Return "at " & Math.Round(x0, 3).ToString() & ", " & Math.Round(y0, 3).ToString() & ", " & Math.Round(z0, 3).ToString() &
               " size " & Math.Round(lx, 3).ToString() & " x " & Math.Round(wy, 3).ToString() & " x " & Math.Round(hz, 3).ToString()
    End Function

End Module
