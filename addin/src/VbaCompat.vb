Option Strict Off
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic

' Small helpers so the macros port from VBA almost line for line (see src/*.bas).
' MsgBox, IsError and Round here shadow the Microsoft.VisualBasic ones inside this project.
Public Module VbaCompat

    ' ---------------------------------------------------------------- messages
    ' Quiet mode (automated tests): no dialog; answers come from PT_Answers (default Yes / OK).
    ' Every message also goes to the task pane log.
    Public PT_Quiet As Boolean = False
    Public PT_Answers As New Queue(Of MsgBoxResult)()
    Public PT_Messages As New List(Of String)()

    Public Function MsgBox(ByVal prompt As Object, Optional ByVal buttons As MsgBoxStyle = MsgBoxStyle.OkOnly,
                           Optional ByVal title As Object = Nothing) As MsgBoxResult
        Dim text As String = CStr(prompt)
        Dim head As String = ""
        If title IsNot Nothing Then head = CStr(title) & ": "
        PT_Messages.Add(head & text)
        Log.Write(head & text.Replace(vbCrLf, " | "))
        ' inside an error handler of a ported macro: keep where it happened
        Try
            If Information.Err().Number <> 0 AndAlso Information.Err().GetException() IsNot Nothing Then
                Log.Write("  error detail: " & Information.Err().GetException().ToString())
            End If
        Catch
        End Try
        If PT_Quiet Then
            Dim ans As MsgBoxResult
            If PT_Answers.Count > 0 Then
                ans = PT_Answers.Dequeue()
            ElseIf (buttons And MsgBoxStyle.YesNo) = MsgBoxStyle.YesNo OrElse (buttons And MsgBoxStyle.YesNoCancel) = MsgBoxStyle.YesNoCancel Then
                ans = MsgBoxResult.Yes
            Else
                ans = MsgBoxResult.Ok
            End If
            Log.Write("  (quiet answer: " & ans.ToString() & ")")
            Return ans
        End If
        Return Interaction.MsgBox(text, buttons, If(title Is Nothing, AddIn.TabTitle, CStr(title)))
    End Function

    ' ---------------------------------------------------------------- values
    ' VBA Array(...): a 0-based Object array.
    Public Function Arr(ByVal ParamArray items() As Object) As Object()
        Return items
    End Function

    ' VBA IsEmpty: an Empty Variant arrives in .NET as Nothing. SolidWorks also returns VT_NULL (DBNull here) for
    ' "no array" in some calls, where VBA got Empty: treat both as empty.
    Public Function IsEmpty(ByVal v As Object) As Boolean
        Return v Is Nothing OrElse TypeOf v Is DBNull
    End Function

    ' Excel error cells (#N/A ...) arrive as Int32 codes.
    Public Function IsError(ByVal v As Object) As Boolean
        Return PTCommon.PT_IsError(v)
    End Function

    ' VBA Round: banker's rounding, like Math.Round.
    Public Function Round(ByVal v As Double, Optional ByVal digits As Integer = 0) As Double
        Return Math.Round(v, digits)
    End Function

    ' Format(Now, "yyyy-mm-dd hh:nn") of the macros
    Public Function NowText() As String
        Return DateTime.Now.ToString("yyyy-MM-dd HH:mm")
    End Function

    ' Format(Now, "yyyymmdd_hhnn") of the macros
    Public Function NowStamp() As String
        Return DateTime.Now.ToString("yyyyMMdd_HHmm")
    End Function

    ' VBA ReDim a(1 To rows, 1 To cols): a 1-based 2D array, as Excel returns Range.Value.
    Public Function NewTable(ByVal rows As Integer, ByVal cols As Integer) As Object(,)
        Return CType(Array.CreateInstance(GetType(Object), New Integer() {rows, cols}, New Integer() {1, 1}), Object(,))
    End Function

    ' Pass for an optional SolidWorks object argument that VBA gives as Nothing (VT_DISPATCH null, not VT_EMPTY).
    Public ReadOnly Property NullDisp As Object
        Get
            Return New DispatchWrapper(Nothing)
        End Get
    End Property

    ' ModelDoc2.GetType through late binding (obj.GetType on an Object is .NET's Object.GetType).
    Public Function SwDocType(ByVal md As Object) As Integer
        Return CInt(CallByName(md, "GetType", CallType.Method))
    End Function

    ' Text file like Scripting.FileSystemObject.CreateTextFile (overwrite).
    Public Function NewTextFile(ByVal path As String) As System.IO.StreamWriter
        Return New System.IO.StreamWriter(path, False, System.Text.Encoding.Default)
    End Function

End Module

' Scripting.Dictionary look-alike: keys kept in insertion order, binary (case-sensitive) compare.
' Item on a missing key returns Nothing (Scripting.Dictionary would add an empty entry), Item set adds.
Public Class Dict
    Private ReadOnly map As New Dictionary(Of String, Object)(StringComparer.Ordinal)
    Private ReadOnly order As New List(Of String)()

    Public Function Exists(ByVal key As Object) As Boolean
        Return map.ContainsKey(CStr(key))
    End Function

    Public Sub Add(ByVal key As Object, ByVal value As Object)
        Dim k As String = CStr(key)
        If map.ContainsKey(k) Then Throw New ArgumentException("This key is already associated with an element of this collection: " & k)
        map.Add(k, value)
        order.Add(k)
    End Sub

    Default Public Property Item(ByVal key As Object) As Object
        Get
            Dim v As Object = Nothing
            map.TryGetValue(CStr(key), v)
            Return v
        End Get
        Set(ByVal value As Object)
            Dim k As String = CStr(key)
            If Not map.ContainsKey(k) Then order.Add(k)
            map(k) = value
        End Set
    End Property

    Public ReadOnly Property Count As Integer
        Get
            Return order.Count
        End Get
    End Property

    ' Snapshot of the keys, safe to change the dictionary while looping.
    Public Function Keys() As Object()
        Dim a(order.Count - 1) As Object
        For i As Integer = 0 To order.Count - 1
            a(i) = order(i)
        Next
        Return a
    End Function

    Public Sub Remove(ByVal key As Object)
        Dim k As String = CStr(key)
        If map.Remove(k) Then order.Remove(k)
    End Sub

    Public Sub RemoveAll()
        map.Clear()
        order.Clear()
    End Sub
End Class
