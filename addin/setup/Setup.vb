Option Strict On
Option Explicit On

Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms
Imports Microsoft.Win32

' One-file installer for the add-in. All files are embedded as resources named "payload/<relative path>".
' Installs for the current Windows user only (no admin): files to %LOCALAPPDATA%\PCBThermal, registry under HKCU.
'   PCBThermal_Setup.exe              window with Install / Uninstall
'   PCBThermal_Setup.exe /silent      install without questions
'   PCBThermal_Setup.exe /uninstall   remove (used by Windows "Installed apps")
'   PCBThermal_Setup.exe /watch       loader: loads the add-in at each SolidWorks start (see Loader.vb)
'   PCBThermal_Setup.exe /load        loader: loads the add-in into the running SolidWorks once
Public Module Setup

    Private Const ClassName As String = "PCBThermal.AddIn"
    Private Const CatSwAddin As String = "{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"
    Private Const UninstallKey As String = "Software\Microsoft\Windows\CurrentVersion\Uninstall\PCBThermal"
    Private Const PayloadPrefix As String = "payload/"
    Private Const RunKey As String = "Software\Microsoft\Windows\CurrentVersion\Run"
    Private Const RunValue As String = "PCBThermalLoader"
    Public Const AppName As String = "NTIS thermal simulation"
    Public Const DllName As String = "NTIS_thermal_simulation.dll"
    Private ReadOnly OldDllNames As String() = {"PCBThermal.dll", "NTIS_thermal_similation.dll"}
    Private Const ShortcutName As String = "Load NTIS thermal simulation into SolidWorks.lnk"
    Private ReadOnly OldShortcutNames As String() = {"Load PCB Thermal into SolidWorks.lnk", "Load NTIS thermal similation into SolidWorks.lnk"}

    Private Function InstallDir() As String
        Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PCBThermal")
    End Function

    Private Function Version() As String
        Return Assembly.GetExecutingAssembly().GetName().Version.ToString(3)
    End Function

    <STAThread>
    Public Function Main(ByVal args As String()) As Integer
        Dim silent As Boolean = False
        Dim remove As Boolean = False
        Dim watch As Boolean = False
        Dim loadNow As Boolean = False
        For Each a As String In args
            Dim s As String = a.Trim().TrimStart("/"c, "-"c).ToLowerInvariant()
            If s = "silent" OrElse s = "s" OrElse s = "quiet" Then silent = True
            If s = "uninstall" OrElse s = "u" Then remove = True
            If s = "watch" Then watch = True
            If s = "load" Then loadNow = True
        Next
        If watch Then
            Loader.Watch()
            Return 0
        End If
        Application.EnableVisualStyles()
        If loadNow Then
            MessageBox.Show(Loader.LoadOnce(), AppName, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return 0
        End If
        If silent Then
            Try
                If remove Then DoUninstall() Else DoInstall()
                Return 0
            Catch ex As Exception
                Console.Error.WriteLine(ex.Message)
                Return 1
            End Try
        End If
        If remove Then
            If MessageBox.Show("Remove " & AppName & " for " & Environment.UserName & "?", AppName,
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return 0
            Return RunWithMessage(False)
        End If
        Application.Run(New SetupForm())
        Return 0
    End Function

    Public Function RunWithMessage(ByVal install As Boolean) As Integer
        Try
            Dim msg As String
            If install Then msg = DoInstall() Else msg = DoUninstall()
            MessageBox.Show(msg, AppName, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return 0
        Catch ex As Exception
            MessageBox.Show(ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return 1
        End Try
    End Function

    Private Sub CheckSolidWorksClosed()
        If Process.GetProcessesByName("SLDWORKS").Length > 0 Then
            Throw New InvalidOperationException("Close SolidWorks first, then run setup again.")
        End If
    End Sub

    Public Function DoInstall() As String
        CheckSolidWorksClosed()
        Dim dest As String = InstallDir()
        Directory.CreateDirectory(dest)

        ' Unpack the embedded files
        Dim asm As Assembly = Assembly.GetExecutingAssembly()
        Dim count As Integer = 0
        For Each resName As String In asm.GetManifestResourceNames()
            If Not resName.StartsWith(PayloadPrefix, StringComparison.Ordinal) Then Continue For
            Dim rel As String = resName.Substring(PayloadPrefix.Length).Replace("/"c, Path.DirectorySeparatorChar)
            Dim target As String = Path.Combine(dest, rel)
            Directory.CreateDirectory(Path.GetDirectoryName(target))
            Using src As Stream = asm.GetManifestResourceStream(resName)
                Using fs As New FileStream(target, FileMode.Create, FileAccess.Write)
                    src.CopyTo(fs)
                End Using
            End Using
            count += 1
        Next
        ' Keep a copy of setup itself for "Installed apps" > Uninstall and as the loader
        Dim selfCopy As String = Path.Combine(dest, "PCBThermal_Setup.exe")
        Loader.StopWatchers(selfCopy)
        If Not String.Equals(asm.Location, selfCopy, StringComparison.OrdinalIgnoreCase) Then
            File.Copy(asm.Location, selfCopy, True)
        End If

        ' Older versions used other names: remove their DLL and shortcut
        Try
            For Each f As String In OldDllNames
                File.Delete(Path.Combine(dest, f))
            Next
            For Each f As String In OldShortcutNames
                File.Delete(Path.Combine(Path.GetDirectoryName(ShortcutPath()), f))
            Next
        Catch
        End Try

        Dim dll As String = Path.Combine(dest, DllName)
        If Not File.Exists(dll) Then Throw New FileNotFoundException("Setup is incomplete: " & DllName & " missing.")
        Dim an As AssemblyName = AssemblyName.GetAssemblyName(dll)
        Dim codeBase As String = New Uri(dll).AbsoluteUri

        ' COM class, per user
        Dim clsid As String = "Software\Classes\CLSID\" & AddInGuid
        Registry.CurrentUser.DeleteSubKeyTree(clsid, False)   ' drop entries of older versions
        SetValues(clsid, "", ClassName)
        For Each sub1 As String In New String() {clsid & "\InprocServer32", clsid & "\InprocServer32\" & an.Version.ToString()}
            Using k As RegistryKey = Registry.CurrentUser.CreateSubKey(sub1)
                If sub1.EndsWith("InprocServer32", StringComparison.Ordinal) Then
                    k.SetValue("", "mscoree.dll")
                    k.SetValue("ThreadingModel", "Both")
                End If
                k.SetValue("Class", ClassName)
                k.SetValue("Assembly", an.FullName)
                k.SetValue("RuntimeVersion", "v4.0.30319")
                k.SetValue("CodeBase", codeBase)
            End Using
        Next
        SetValues(clsid & "\ProgId", "", ProgId)
        Registry.CurrentUser.CreateSubKey(clsid & "\Implemented Categories\" & CatSwAddin).Close()
        SetValues("Software\Classes\" & ProgId, "", ClassName)
        SetValues("Software\Classes\" & ProgId & "\CLSID", "", AddInGuid)

        ' SolidWorks add-in list, loaded at startup
        Using k As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\SolidWorks\AddIns\" & AddInGuid)
            k.SetValue("", 1, RegistryValueKind.DWord)
            k.SetValue("Title", AppName)
            k.SetValue("Description", "PCB thermal workflow: ribbon tab and task pane")
        End Using
        ' Start Up tick in Tools > Add-Ins. Keep the user's choice on upgrade.
        Using k As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\SolidWorks\AddInsStartup\" & AddInGuid)
            If k.GetValue("") Is Nothing Then k.SetValue("", 1, RegistryValueKind.DWord)
        End Using

        ' SolidWorks reads the add-in list above only from HKLM (admin). The loader calls LoadAddIn at each SolidWorks start.
        Using k As RegistryKey = Registry.CurrentUser.CreateSubKey(RunKey)
            k.SetValue(RunValue, """" & selfCopy & """ /watch")
        End Using
        CreateShortcut(selfCopy)
        Dim psi As New ProcessStartInfo(selfCopy, "/watch")
        psi.UseShellExecute = False
        psi.CreateNoWindow = True
        Process.Start(psi)

        ' Windows "Installed apps" entry, per user
        Using k As RegistryKey = Registry.CurrentUser.CreateSubKey(UninstallKey)
            k.SetValue("DisplayName", AppName & " for SolidWorks")
            k.SetValue("DisplayVersion", Version())
            k.SetValue("Publisher", AppName)
            k.SetValue("InstallLocation", dest)
            k.SetValue("DisplayIcon", selfCopy)
            k.SetValue("UninstallString", """" & selfCopy & """ /uninstall")
            k.SetValue("QuietUninstallString", """" & selfCopy & """ /uninstall /silent")
            k.SetValue("NoModify", 1, RegistryValueKind.DWord)
            k.SetValue("NoRepair", 1, RegistryValueKind.DWord)
        End Using

        Return AppName & " " & Version() & " installed for " & Environment.UserName & " (" & count.ToString() & " files)." &
               Environment.NewLine & "Folder: " & dest & Environment.NewLine &
               "Start SolidWorks and open a part or assembly: look for the " & AppName & " tab." & Environment.NewLine &
               "If it is missing: Start menu > " & Path.GetFileNameWithoutExtension(ShortcutName) & "."
    End Function

    Public Function DoUninstall() As String
        CheckSolidWorksClosed()
        Loader.StopWatchers(Path.Combine(InstallDir(), "PCBThermal_Setup.exe"))
        Try
            Using k As RegistryKey = Registry.CurrentUser.OpenSubKey(RunKey, True)
                If k IsNot Nothing Then k.DeleteValue(RunValue, False)
            End Using
        Catch
        End Try
        Try
            File.Delete(ShortcutPath())
            For Each f As String In OldShortcutNames
                File.Delete(Path.Combine(Path.GetDirectoryName(ShortcutPath()), f))
            Next
        Catch
        End Try
        For Each keyPath As String In New String() {
                "Software\Classes\CLSID\" & AddInGuid, "Software\Classes\" & ProgId,
                "Software\SolidWorks\AddIns\" & AddInGuid, "Software\SolidWorks\AddInsStartup\" & AddInGuid, UninstallKey}
            Try
                Registry.CurrentUser.DeleteSubKeyTree(keyPath, False)
            Catch
            End Try
        Next
        Dim dest As String = InstallDir()
        Dim running As String = Assembly.GetExecutingAssembly().Location
        If Directory.Exists(dest) Then
            For Each f As String In Directory.GetFiles(dest, "*", SearchOption.AllDirectories)
                If String.Equals(f, running, StringComparison.OrdinalIgnoreCase) Then Continue For
                Try
                    File.Delete(f)
                Catch
                End Try
            Next
            ' The running setup copy cannot delete itself: let cmd remove the folder after we exit.
            If running.StartsWith(dest, StringComparison.OrdinalIgnoreCase) Then
                Dim psi As New ProcessStartInfo("cmd.exe", "/c ping 127.0.0.1 -n 3 > nul & rmdir /s /q """ & dest & """")
                psi.WindowStyle = ProcessWindowStyle.Hidden
                psi.CreateNoWindow = True
                Process.Start(psi)
            Else
                Try
                    Directory.Delete(dest, True)
                Catch
                End Try
            End If
        End If
        Return AppName & " removed for " & Environment.UserName & "." & Environment.NewLine &
               "Your settings and log in %APPDATA%\PCBThermal were kept."
    End Function

    Private Function ShortcutPath() As String
        Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), ShortcutName)
    End Function

    ' Start menu shortcut that loads the add-in into the running SolidWorks (if the watcher is off or missed it).
    Private Sub CreateShortcut(ByVal exePath As String)
        Try
            Dim sh As Object = Interaction.CreateObject("WScript.Shell")
            Dim lnk As Object = CallByName(sh, "CreateShortcut", CallType.Method, ShortcutPath())
            CallByName(lnk, "TargetPath", CallType.Let, exePath)
            CallByName(lnk, "Arguments", CallType.Let, "/load")
            CallByName(lnk, "Description", CallType.Let, "Load the " & AppName & " add-in into the running SolidWorks")
            CallByName(lnk, "IconLocation", CallType.Let, exePath & ",0")
            CallByName(lnk, "Save", CallType.Method)
        Catch
        End Try
    End Sub

    Private Sub SetValues(ByVal keyPath As String, ByVal valueName As String, ByVal value As String)
        Using k As RegistryKey = Registry.CurrentUser.CreateSubKey(keyPath)
            k.SetValue(valueName, value)
        End Using
    End Sub

    Public Function IsInstalled() As String
        Using k As RegistryKey = Registry.CurrentUser.OpenSubKey(UninstallKey)
            If k Is Nothing Then Return ""
            Return CStr(k.GetValue("DisplayVersion", "?"))
        End Using
    End Function

    Public Function SetupVersion() As String
        Return Version()
    End Function

End Module

' Small window: what is installed, Install / Uninstall.
Public Class SetupForm
    Inherits Form

    Public Sub New()
        Me.Text = Setup.AppName & " setup " & Setup.SetupVersion()
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 9.5F)
        Me.ClientSize = New Size(420, 170)

        Dim installed As String = Setup.IsInstalled()
        Dim info As New Label()
        info.AutoSize = False
        info.SetBounds(16, 14, 388, 90)
        Dim t As String = Setup.AppName & " add-in for SolidWorks, version " & Setup.SetupVersion() & "." & Environment.NewLine &
                          "Installs for " & Environment.UserName & " only. No admin rights needed." & Environment.NewLine & Environment.NewLine
        If installed = "" Then
            t &= "Not installed yet."
        Else
            t &= "Installed now: version " & installed & "."
        End If
        info.Text = t
        Me.Controls.Add(info)

        Dim inst As New Button()
        If installed = "" Then inst.Text = "Install" Else inst.Text = "Install " & Setup.SetupVersion()
        inst.SetBounds(16, 120, 150, 32)
        AddHandler inst.Click, AddressOf OnInstall
        Me.Controls.Add(inst)
        Me.AcceptButton = inst

        Dim rem1 As New Button()
        rem1.Text = "Uninstall"
        rem1.SetBounds(176, 120, 110, 32)
        rem1.Enabled = (installed <> "")
        AddHandler rem1.Click, AddressOf OnUninstall
        Me.Controls.Add(rem1)

        Dim cls As New Button()
        cls.Text = "Close"
        cls.SetBounds(296, 120, 108, 32)
        AddHandler cls.Click, AddressOf OnCloseClick
        Me.Controls.Add(cls)
        Me.CancelButton = cls
    End Sub

    Private Sub OnInstall(ByVal sender As Object, ByVal e As EventArgs)
        If Setup.RunWithMessage(True) = 0 Then Me.Close()
    End Sub

    Private Sub OnUninstall(ByVal sender As Object, ByVal e As EventArgs)
        If Setup.RunWithMessage(False) = 0 Then Me.Close()
    End Sub

    Private Sub OnCloseClick(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

End Class
