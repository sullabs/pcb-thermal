Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.Win32

' Loads the add-in into SolidWorks without admin rights.
' SolidWorks only reads its add-in list from HKLM, so a per-user install cannot appear there. ISldWorks.LoadAddIn
' with the dll path works for a per-user COM class, so this loader calls it:
'   PCBThermal_Setup.exe /watch   runs at Windows logon (HKCU Run), waits for each SolidWorks start, loads the add-in
'   PCBThermal_Setup.exe /load    loads it once into the running SolidWorks (Start menu shortcut)
' The watcher respects the Start Up tick in Tools > Add-Ins (HKCU\Software\SolidWorks\AddInsStartup\{guid}).
Public Module Loader

    Public Const AddInGuid As String = "{F7AF128D-DEE1-4BEB-A464-4FE0B17337DF}"
    Public Const ProgId As String = "PCBThermal.AddIn"
    Private Const MutexName As String = "Local\PCBThermalLoader"
    Private Const PollMs As Integer = 3000

    Public Function DllPath() As String
        Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PCBThermal", Setup.DllName)
    End Function

    Private Sub Log(ByVal msg As String)
        Try
            Dim dir As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PCBThermal")
            Directory.CreateDirectory(dir)
            File.AppendAllText(Path.Combine(dir, "loader.log"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "  " & msg & Environment.NewLine)
        Catch
        End Try
    End Sub

    ' True unless the user unticked Start Up for the add-in in Tools > Add-Ins.
    Public Function StartupEnabled() As Boolean
        Try
            Using k As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\SolidWorks\AddInsStartup\" & AddInGuid)
                If k Is Nothing Then Return True
                Dim v As Object = k.GetValue("")
                If v Is Nothing Then Return True
                Return Convert.ToInt32(v) <> 0
            End Using
        Catch
            Return True
        End Try
    End Function

    Private Function GetSolidWorks() As Object
        Try
            Return Marshal.GetActiveObject("SldWorks.Application")
        Catch
            Return Nothing
        End Try
    End Function

    Private Sub ReleaseCom(ByVal o As Object)
        If o IsNot Nothing AndAlso Marshal.IsComObject(o) Then
            Try
                Marshal.FinalReleaseComObject(o)
            Catch
            End Try
        End If
    End Sub

    ' Result of one load attempt: "loaded", "already", "busy" (try later) or an error text.
    Private Function TryLoad(ByVal sw As Object) As String
        Try
            If Not CBool(CallByName(sw, "StartupProcessCompleted", CallType.Get)) Then Return "busy"
            Dim existing As Object = CallByName(sw, "GetAddInObject", CallType.Method, ProgId)
            If existing IsNot Nothing Then
                ReleaseCom(existing)
                Return "already"
            End If
            Dim r As Integer = CInt(CallByName(sw, "LoadAddIn", CallType.Method, DllPath()))
            If r = 0 Then Return "loaded"
            Return "LoadAddIn returned " & r.ToString()
        Catch ex As COMException
            ' RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER: SolidWorks is busy
            If ex.ErrorCode = &H80010001 OrElse ex.ErrorCode = &H8001010A Then Return "busy"
            Return ex.Message
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ' /load: once, into the running SolidWorks. Returns a message for the user.
    Public Function LoadOnce() As String
        If Not File.Exists(DllPath()) Then Return Setup.AppName & " is not installed for " & Environment.UserName & "."
        Dim sw As Object = GetSolidWorks()
        If sw Is Nothing Then Return "SolidWorks is not running. Start it, then try again."
        Try
            Dim res As String = "busy"
            For i As Integer = 1 To 20
                res = TryLoad(sw)
                If res <> "busy" Then Exit For
                Thread.Sleep(PollMs)
            Next
            Log("load: " & res)
            Select Case res
                Case "loaded" : Return Setup.AppName & " loaded. Open a part or assembly and look for the " & Setup.AppName & " tab."
                Case "already" : Return Setup.AppName & " is already loaded in SolidWorks."
                Case "busy" : Return "SolidWorks is still busy. Try again in a moment."
                Case Else : Return "Could not load " & Setup.AppName & ": " & res
            End Select
        Finally
            ReleaseCom(sw)
        End Try
    End Function

    ' /watch: stays running for the Windows session. One SolidWorks start = one load attempt cycle.
    Public Sub Watch()
        Dim created As Boolean = False
        Using m As New Mutex(True, MutexName, created)
            If Not created Then Return ' another watcher already runs
            Log("watcher started")
            Dim done As New HashSet(Of Integer)()
            Do
                If Not File.Exists(DllPath()) Then
                    Log("watcher stopped: add-in removed")
                    Exit Do
                End If
                Dim alive As New HashSet(Of Integer)()
                For Each p As Process In Process.GetProcessesByName("SLDWORKS")
                    alive.Add(p.Id)
                    If done.Contains(p.Id) Then Continue For
                    If Not StartupEnabled() Then
                        done.Add(p.Id)
                        Log("SolidWorks " & p.Id.ToString() & ": Start Up is off, not loading")
                        Continue For
                    End If
                    Dim sw As Object = GetSolidWorks()
                    If sw Is Nothing Then Continue For ' not registered yet, try next round
                    Dim res As String = TryLoad(sw)
                    ReleaseCom(sw)
                    If res = "busy" Then Continue For
                    done.Add(p.Id)
                    Log("SolidWorks " & p.Id.ToString() & ": " & res)
                Next
                done.IntersectWith(alive)
                Thread.Sleep(PollMs)
            Loop
        End Using
    End Sub

    ' Stops watchers started from the given exe (before files are replaced or removed).
    Public Sub StopWatchers(ByVal exePath As String)
        Dim me1 As Integer = Process.GetCurrentProcess().Id
        For Each p As Process In Process.GetProcesses()
            If p.Id = me1 Then Continue For
            If Not p.ProcessName.StartsWith("PCBThermal_Setup", StringComparison.OrdinalIgnoreCase) Then Continue For
            Try
                If String.Equals(p.MainModule.FileName, exePath, StringComparison.OrdinalIgnoreCase) Then
                    p.Kill()
                    p.WaitForExit(5000)
                End If
            Catch
            End Try
        Next
    End Sub

End Module
