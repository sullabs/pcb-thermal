Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.swpublished

' SolidWorks add-in entry point. Builds the "NTIS thermal similation" ribbon tab, the menu and the task pane.
' Registered per user by install.ps1 (no admin rights).
<ComVisible(True), Guid(AddIn.AddInGuid), ProgId("PCBThermal.AddIn")>
Public Class AddIn
    Implements ISwAddin

    Public Const AddInGuid As String = "F7AF128D-DEE1-4BEB-A464-4FE0B17337DF"
    Public Const TabTitle As String = "NTIS thermal similation"
    Private Const OldTabTitle As String = "PCB Thermal"
    Private Const GroupId As Integer = 7301

    Private swApp As SldWorks
    Private addinCookie As Integer
    Private cmdMgr As CommandManager
    Private paneView As TaskpaneView
    Private pane As TaskPane

    ' Ribbon commands in order. Index = icon position in the strips made by build.ps1.
    ' Steps 1 to 8 run the existing VBA macros (module.main) from the macro project chosen in the pane.
    Public Shared ReadOnly StepTitles As String() = {
        "Prepare Workbook", "Import STEP", "Build Board", "Build Spreader",
        "Sync CAD to Excel", "Run Thermal Study", "Run Flow Sim", "Read Results"}
    Public Shared ReadOnly StepModules As String() = {
        "SetupWorkbook", "ImportStep", "BuildBoard", "BuildSpreader",
        "SyncFromCAD", "SetupSim", "RunFlow", "ReadResults"}

    Public ReadOnly Property App As SldWorks
        Get
            Return swApp
        End Get
    End Property

    Public Function ConnectToSW(ByVal ThisSW As Object, ByVal Cookie As Integer) As Boolean Implements ISwAddin.ConnectToSW
        swApp = CType(ThisSW, SldWorks)
        addinCookie = Cookie
        PTCommon.swApp = swApp
        swApp.SetAddinCallbackInfo2(0, Me, addinCookie)
        Try
            AddTaskPane()
        Catch ex As Exception
            Log.Write("Task pane failed: " & ex.Message)
        End Try
        Try
            AddCommands()
        Catch ex As Exception
            Log.Write("Ribbon failed: " & ex.Message)
        End Try
        Log.Write(TabTitle & " add-in loaded.")
        Return True
    End Function

    Public Function DisconnectFromSW() As Boolean Implements ISwAddin.DisconnectFromSW
        Try
            RemoveCommands()
        Catch
        End Try
        Try
            If paneView IsNot Nothing Then paneView.DeleteView()
        Catch
        End Try
        If pane IsNot Nothing Then pane.Dispose()
        pane = Nothing
        paneView = Nothing
        cmdMgr = Nothing
        PTCommon.swApp = Nothing
        swApp = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Return True
    End Function

    ' ---------------- ribbon ----------------

    Private Function IconFiles(ByVal stem As String) As String()
        Dim sizes As Integer() = {20, 32, 40, 64, 96, 128}
        Dim res(sizes.Length - 1) As String
        For i As Integer = 0 To sizes.Length - 1
            res(i) = Path.Combine(AddInDir(), "icons", stem & "_" & sizes(i) & ".png")
        Next
        Return res
    End Function

    Public Shared Function AddInDir() As String
        Return Path.GetDirectoryName(GetType(AddIn).Assembly.Location)
    End Function

    Private Sub AddCommands()
        cmdMgr = swApp.GetCommandManager(addinCookie)
        Dim errs As Integer = 0
        Dim grp As CommandGroup = cmdMgr.CreateCommandGroup2(GroupId, TabTitle, "PCB thermal workflow", "PCB thermal workflow", -1, True, errs)
        grp.IconList = IconFiles("strip")
        grp.MainIconList = IconFiles("main")

        Dim opts As Integer = CInt(swCommandItemType_e.swMenuItem) Or CInt(swCommandItemType_e.swToolbarItem)
        Dim n As Integer = StepTitles.Length + 2
        Dim idx(n - 1) As Integer
        For i As Integer = 0 To StepTitles.Length - 1
            Dim tip As String = (i + 1).ToString() & " " & StepTitles(i)
            idx(i) = grp.AddCommandItem2(tip, -1, "Runs step " & StepModules(i), tip, i, "OnStep" & (i + 1).ToString(), "", i + 1, opts)
        Next
        idx(n - 2) = grp.AddCommandItem2("Check Model", -1, "Counts bodies and triangles in the open model", "Check Model", n - 2, "OnCheckModel", "", 101, opts)
        idx(n - 1) = grp.AddCommandItem2("Check Workbook", -1, "Checks the open Excel workbook", "Check Workbook", n - 1, "OnCheckWorkbook", "", 102, opts)
        grp.HasToolbar = True
        grp.HasMenu = True
        grp.Activate()

        Dim ids(n - 1) As Integer
        Dim styles(n - 1) As Integer
        For i As Integer = 0 To n - 1
            ids(i) = grp.CommandID(idx(i))
            styles(i) = CInt(swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow)
        Next
        For Each docType As Integer In New Integer() {CInt(swDocumentTypes_e.swDocPART), CInt(swDocumentTypes_e.swDocASSEMBLY)}
            Dim oldTab As CommandTab = cmdMgr.GetCommandTab(docType, OldTabTitle)
            If oldTab IsNot Nothing Then cmdMgr.RemoveCommandTab(oldTab)
            Dim cmdTab As CommandTab = cmdMgr.GetCommandTab(docType, TabTitle)
            If cmdTab IsNot Nothing Then cmdMgr.RemoveCommandTab(cmdTab)
            cmdTab = cmdMgr.AddCommandTab(docType, TabTitle)
            Dim box As CommandTabBox = cmdTab.AddCommandTabBox()
            box.AddCommands(ids, styles)
        Next
    End Sub

    Private Sub RemoveCommands()
        If cmdMgr Is Nothing Then Return
        For Each docType As Integer In New Integer() {CInt(swDocumentTypes_e.swDocPART), CInt(swDocumentTypes_e.swDocASSEMBLY)}
            Dim oldTab As CommandTab = cmdMgr.GetCommandTab(docType, OldTabTitle)
            If oldTab IsNot Nothing Then cmdMgr.RemoveCommandTab(oldTab)
            Dim cmdTab As CommandTab = cmdMgr.GetCommandTab(docType, TabTitle)
            If cmdTab IsNot Nothing Then cmdMgr.RemoveCommandTab(cmdTab)
        Next
        cmdMgr.RemoveCommandGroup2(GroupId, True)
    End Sub

    ' ---------------- task pane ----------------

    Private Sub AddTaskPane()
        paneView = swApp.CreateTaskpaneView3(IconFiles("main"), TabTitle)
        pane = New TaskPane(Me)
        paneView.DisplayWindowFromHandlex64(pane.Handle.ToInt64())
        Log.Attach(pane)
    End Sub

    ' ---------------- callbacks (names are used in AddCommandItem2) ----------------

    Public Sub OnStep1()
        RunStep(0)
    End Sub
    Public Sub OnStep2()
        RunStep(1)
    End Sub
    Public Sub OnStep3()
        RunStep(2)
    End Sub
    Public Sub OnStep4()
        RunStep(3)
    End Sub
    Public Sub OnStep5()
        RunStep(4)
    End Sub
    Public Sub OnStep6()
        RunStep(5)
    End Sub
    Public Sub OnStep7()
        RunStep(6)
    End Sub
    Public Sub OnStep8()
        RunStep(7)
    End Sub

    Public Sub OnCheckModel()
        ShowPane()
        ModelCheck.Run(swApp)
    End Sub

    Public Sub OnCheckWorkbook()
        ShowPane()
        ExcelLink.CheckWorkbook()
        If pane IsNot Nothing Then pane.RefreshWorkbook()
    End Sub

    Private Sub ShowPane()
        Try
            If paneView IsNot Nothing Then paneView.ShowView()
        Catch
        End Try
    End Sub

    ' Runs one workflow step: the built-in (ported) code, or the VBA macro when the pane says so.
    Public Sub RunStep(ByVal i As Integer)
        ShowPane()
        If Settings.UseVba Then
            RunVbaStep(i)
            Return
        End If
        RunNative(i)
    End Sub

    ' The ported macro of step i (index into StepModules).
    Private Shared Sub CallStep(ByVal i As Integer)
        Select Case StepModules(i)
            Case "SetupWorkbook" : SetupWorkbook.Main()
            Case "ImportStep" : ImportStep.Main()
            Case "BuildBoard" : BuildBoard.Main()
            Case "BuildSpreader" : BuildSpreader.Main()
            Case "SyncFromCAD" : SyncFromCAD.Main()
            Case "SetupSim" : SetupSim.Main()
            Case "RunFlow" : RunFlow.Main()
            Case "ReadResults" : ReadResults.Main()
        End Select
    End Sub

    Private Function RunNative(ByVal i As Integer) As Boolean
        Dim stepName As String = (i + 1).ToString() & " " & StepTitles(i)
        Log.Write(stepName & ": started.")
        If pane IsNot Nothing Then pane.SetStepStatus(i, "running")
        Dim t0 As DateTime = DateTime.Now
        Dim ok As Boolean = True
        Try
            CallStep(i)
        Catch ex As Exception
            ok = False
            Try
                Dim xl As Object = ExcelLink.GetExcel()
                If xl IsNot Nothing Then xl.ScreenUpdating = True
            Catch
            End Try
            VbaCompat.MsgBox(StepModules(i) & " stopped: " & ex.Message & vbCrLf & "Send this message with the log file %APPDATA%\PCBThermal\log.txt.",
                             MsgBoxStyle.Exclamation, StepModules(i))
            Log.Write(ex.ToString())
        End Try
        Dim secs As String = (DateTime.Now - t0).TotalSeconds.ToString("0.0")
        Log.Write(stepName & ": " & If(ok, "done", "failed") & " in " & secs & " s.")
        If pane IsNot Nothing Then pane.SetStepStatus(i, If(ok, "done ", "failed ") & DateTime.Now.ToString("HH:mm"))
        Return ok
    End Function

    ' For automated tests (VBScript: sw.GetAddInObject("PCBThermal.AddIn").TestRun("ImportStep", "Y")).
    ' Runs a ported step with no dialogs: message boxes are answered from answers ("Y", "N", "C", "O", comma separated,
    ' default Yes / OK). Returns every message the step showed, separated by lines of "-----".
    Public Function TestRun(ByVal moduleName As String, ByVal answers As String) As String
        Dim i As Integer = Array.IndexOf(StepModules, moduleName)
        If i < 0 Then Return "unknown step " & moduleName
        PT_Messages.Clear()
        PT_Answers.Clear()
        For Each a As String In answers.Split(","c)
            Select Case a.Trim().ToUpperInvariant()
                Case "Y" : PT_Answers.Enqueue(MsgBoxResult.Yes)
                Case "N" : PT_Answers.Enqueue(MsgBoxResult.No)
                Case "C" : PT_Answers.Enqueue(MsgBoxResult.Cancel)
                Case "O" : PT_Answers.Enqueue(MsgBoxResult.Ok)
            End Select
        Next
        PT_Quiet = True
        Dim t0 As DateTime = DateTime.Now
        Try
            RunNative(i)
        Finally
            PT_Quiet = False
        End Try
        Return "time " & (DateTime.Now - t0).TotalSeconds.ToString("0.0") & " s" & vbCrLf & "-----" & vbCrLf &
               String.Join(vbCrLf & "-----" & vbCrLf, PT_Messages.ToArray())
    End Function

    ' Runs one VBA macro (module.main) from the macro project and reports the result in the pane.
    Private Sub RunVbaStep(ByVal i As Integer)
        Dim swpFile As String = Settings.MacroProject
        If swpFile = "" OrElse Not File.Exists(swpFile) Then
            Log.Write("Pick the macro project (.swp) in the " & TabTitle & " pane first.")
            swApp.SendMsgToUser2("Pick the macro project (.swp) in the " & TabTitle & " task pane first.",
                CInt(swMessageBoxIcon_e.swMbWarning), CInt(swMessageBoxBtn_e.swMbOk))
            Return
        End If
        Dim stepName As String = (i + 1).ToString() & " " & StepTitles(i)
        Log.Write(stepName & ": started.")
        If pane IsNot Nothing Then pane.SetStepStatus(i, "running")
        Dim t0 As DateTime = DateTime.Now
        Dim errCode As Integer = 0
        Dim ok As Boolean = False
        Try
            ok = swApp.RunMacro2(swpFile, StepModules(i), "main", CInt(swRunMacroOption_e.swRunMacroUnloadAfterRun), errCode)
        Catch ex As Exception
            Log.Write(stepName & ": " & ex.Message)
        End Try
        Dim secs As String = (DateTime.Now - t0).TotalSeconds.ToString("0.0")
        If ok Then
            Log.Write(stepName & ": done in " & secs & " s.")
            If pane IsNot Nothing Then pane.SetStepStatus(i, "done " & DateTime.Now.ToString("HH:mm"))
        Else
            Log.Write(stepName & ": failed, macro error " & errCode.ToString() & " (" & CType(errCode, swRunMacroError_e).ToString() & ").")
            If pane IsNot Nothing Then pane.SetStepStatus(i, "failed")
        End If
    End Sub

End Class
