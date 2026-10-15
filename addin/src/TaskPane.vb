Option Strict Off
Option Explicit On

Imports System
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

' The task pane: workbook, macro project, workflow steps with status, and a log.
' Built in code (no designer files), shown in SolidWorks by DisplayWindowFromHandlex64.
Public Class TaskPane
    Inherits UserControl

    Private ReadOnly owner As AddIn
    Private wbLabel As Label
    Private swpBox As TextBox
    Private vbaBox As CheckBox
    Private stepList As ListView
    Private logBox As TextBox

    Public Sub New(ByVal addInObj As AddIn)
        owner = addInObj
        BuildUi()
        RefreshWorkbook()
    End Sub

    Private Sub BuildUi()
        Me.BackColor = SystemColors.Window
        Me.Font = New Font("Segoe UI", 9.0F)
        Me.Padding = New Padding(6)

        ' Log fills what is left, so add it first (docking order is reversed).
        logBox = New TextBox()
        logBox.Multiline = True
        logBox.ReadOnly = True
        logBox.ScrollBars = ScrollBars.Vertical
        logBox.Dock = DockStyle.Fill
        logBox.BackColor = SystemColors.Window
        Me.Controls.Add(logBox)

        Dim head As New TableLayoutPanel()
        head.Dock = DockStyle.Top
        head.AutoSize = True
        head.ColumnCount = 2
        head.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        head.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        Me.Controls.Add(head)

        Dim title As New Label()
        title.Text = AddIn.TabTitle
        title.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
        title.AutoSize = True
        head.Controls.Add(title, 0, 0)
        head.SetColumnSpan(title, 2)

        ' Workbook
        head.Controls.Add(Caption("Workbook (active in Excel)"), 0, 1)
        head.SetColumnSpan(head.GetControlFromPosition(0, 1), 2)
        wbLabel = New Label()
        wbLabel.AutoSize = True
        wbLabel.Margin = New Padding(3, 3, 3, 6)
        head.Controls.Add(wbLabel, 0, 2)
        head.Controls.Add(MakeButton("Check", AddressOf OnCheckWorkbook), 1, 2)

        ' Macro project (only used when the VBA macros run instead of the built-in steps)
        vbaBox = New CheckBox()
        vbaBox.Text = "Run the VBA macros (.swp) instead"
        vbaBox.AutoSize = True
        vbaBox.Margin = New Padding(3, 8, 3, 0)
        vbaBox.Checked = Settings.UseVba
        AddHandler vbaBox.CheckedChanged, AddressOf OnVbaChanged
        head.Controls.Add(vbaBox, 0, 3)
        head.SetColumnSpan(vbaBox, 2)
        swpBox = New TextBox()
        swpBox.ReadOnly = True
        swpBox.Dock = DockStyle.Fill
        swpBox.Text = Settings.MacroProject
        swpBox.Enabled = Settings.UseVba
        head.Controls.Add(swpBox, 0, 4)
        head.Controls.Add(MakeButton("Browse", AddressOf OnBrowse), 1, 4)

        ' Steps
        head.Controls.Add(Caption("Workflow (double-click to run)"), 0, 5)
        head.SetColumnSpan(head.GetControlFromPosition(0, 5), 2)
        stepList = New ListView()
        stepList.View = View.Details
        stepList.FullRowSelect = True
        stepList.HeaderStyle = ColumnHeaderStyle.None
        stepList.Columns.Add("Step", 150)
        stepList.Columns.Add("Status", 90)
        stepList.Height = 8 * 20 + 8
        stepList.Dock = DockStyle.Fill
        For i As Integer = 0 To AddIn.StepTitles.Length - 1
            Dim it As New ListViewItem((i + 1).ToString() & "  " & AddIn.StepTitles(i))
            it.SubItems.Add("")
            stepList.Items.Add(it)
        Next
        AddHandler stepList.DoubleClick, AddressOf OnStepDoubleClick
        head.Controls.Add(stepList, 0, 6)
        head.SetColumnSpan(stepList, 2)

        ' Checks
        Dim checks As New FlowLayoutPanel()
        checks.AutoSize = True
        checks.Dock = DockStyle.Fill
        checks.Controls.Add(MakeButton("Check model", AddressOf OnCheckModel))
        checks.Controls.Add(MakeButton("Clear log", AddressOf OnClearLog))
        head.Controls.Add(checks, 0, 7)
        head.SetColumnSpan(checks, 2)

        head.Controls.Add(Caption("Log"), 0, 8)
    End Sub

    Private Function Caption(ByVal text As String) As Label
        Dim l As New Label()
        l.Text = text
        l.AutoSize = True
        l.ForeColor = SystemColors.GrayText
        l.Margin = New Padding(3, 8, 3, 0)
        Return l
    End Function

    Private Function MakeButton(ByVal text As String, ByVal handler As EventHandler) As Button
        Dim b As New Button()
        b.Text = text
        b.AutoSize = True
        AddHandler b.Click, handler
        Return b
    End Function

    ' ---------------- called by the add-in ----------------

    Public Sub AppendLog(ByVal line As String)
        If logBox Is Nothing Then Return
        logBox.AppendText(line & Environment.NewLine)
        logBox.Refresh()
    End Sub

    Public Sub SetStepStatus(ByVal i As Integer, ByVal status As String)
        If i < 0 OrElse i >= stepList.Items.Count Then Return
        stepList.Items(i).SubItems(1).Text = status
        stepList.Refresh()
    End Sub

    Public Sub RefreshWorkbook()
        Dim wb As Object = ExcelLink.ActiveWorkbook()
        If wb Is Nothing Then
            wbLabel.Text = "(none open)"
        Else
            Try
                wbLabel.Text = CStr(wb.Name)
            Catch
                wbLabel.Text = "(Excel busy)"
            End Try
        End If
    End Sub

    ' ---------------- events ----------------

    Private Sub OnCheckWorkbook(ByVal sender As Object, ByVal e As EventArgs)
        RefreshWorkbook()
        ExcelLink.CheckWorkbook()
    End Sub

    Private Sub OnCheckModel(ByVal sender As Object, ByVal e As EventArgs)
        ModelCheck.Run(owner.App)
    End Sub

    Private Sub OnClearLog(ByVal sender As Object, ByVal e As EventArgs)
        logBox.Clear()
    End Sub

    Private Sub OnBrowse(ByVal sender As Object, ByVal e As EventArgs)
        Using dlg As New OpenFileDialog()
            dlg.Title = "Pick the macro project that holds the PCB thermal modules"
            dlg.Filter = "SolidWorks macro (*.swp)|*.swp"
            If swpBox.Text <> "" AndAlso File.Exists(swpBox.Text) Then dlg.InitialDirectory = Path.GetDirectoryName(swpBox.Text)
            If dlg.ShowDialog() = DialogResult.OK Then
                Settings.MacroProject = dlg.FileName
                swpBox.Text = dlg.FileName
                Log.Write("Macro project set: " & dlg.FileName)
            End If
        End Using
    End Sub

    Private Sub OnVbaChanged(ByVal sender As Object, ByVal e As EventArgs)
        Settings.UseVba = vbaBox.Checked
        swpBox.Enabled = vbaBox.Checked
        Log.Write(If(vbaBox.Checked, "Steps run the VBA macros from the .swp.", "Steps run the built-in code."))
    End Sub

    Private Sub OnStepDoubleClick(ByVal sender As Object, ByVal e As EventArgs)
        If stepList.SelectedIndices.Count = 0 Then Return
        owner.RunStep(stepList.SelectedIndices(0))
        RefreshWorkbook()
    End Sub

End Class
