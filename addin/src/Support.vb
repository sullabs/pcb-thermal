Option Strict Off
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text

' User settings, kept in %APPDATA%\PCBThermal\settings.txt as key=value lines.
Public Module Settings

    Private Function FilePath() As String
        Dim dir As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PCBThermal")
        Directory.CreateDirectory(dir)
        Return Path.Combine(dir, "settings.txt")
    End Function

    Private Function ReadAll() As Dictionary(Of String, String)
        Dim d As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim f As String = FilePath()
        If File.Exists(f) Then
            For Each line As String In File.ReadAllLines(f, Encoding.UTF8)
                Dim k As Integer = line.IndexOf("="c)
                If k > 0 Then d(line.Substring(0, k).Trim()) = line.Substring(k + 1).Trim()
            Next
        End If
        Return d
    End Function

    Public Function GetValue(ByVal key As String) As String
        Dim d As Dictionary(Of String, String) = ReadAll()
        If d.ContainsKey(key) Then Return d(key)
        Return ""
    End Function

    Public Sub SetValue(ByVal key As String, ByVal value As String)
        Dim d As Dictionary(Of String, String) = ReadAll()
        d(key) = value
        Dim lines As New List(Of String)
        For Each kv As KeyValuePair(Of String, String) In d
            lines.Add(kv.Key & "=" & kv.Value)
        Next
        File.WriteAllLines(FilePath(), lines.ToArray(), Encoding.UTF8)
    End Sub

    ' The .swp that holds the VBA modules. Default: PCBThermal.swp next to the add-in DLL.
    Public Property MacroProject As String
        Get
            Dim p As String = GetValue("MacroProject")
            If p = "" Then
                Dim guess As String = Path.Combine(AddIn.AddInDir(), "PCBThermal.swp")
                If File.Exists(guess) Then p = guess
            End If
            Return p
        End Get
        Set(ByVal value As String)
            SetValue("MacroProject", value)
        End Set
    End Property

End Module

' Messages go to the task pane log and to %APPDATA%\PCBThermal\log.txt.
Public Module Log

    Private target As TaskPane

    Public Sub Attach(ByVal p As TaskPane)
        target = p
    End Sub

    Public Sub Write(ByVal msg As String)
        Dim line As String = DateTime.Now.ToString("HH:mm:ss") & "  " & msg
        Try
            If target IsNot Nothing Then target.AppendLog(line)
        Catch
        End Try
        Try
            Dim dir As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PCBThermal")
            Directory.CreateDirectory(dir)
            File.AppendAllText(Path.Combine(dir, "log.txt"), DateTime.Now.ToString("yyyy-MM-dd ") & line & Environment.NewLine, Encoding.UTF8)
        Catch
        End Try
    End Sub

End Module
