<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        MenuStrip1 = New MenuStrip()
        FindControllerToolStripMenuItem = New ToolStripMenuItem()
        GetZonesToolStripMenuItem = New ToolStripMenuItem()
        ZoneStatusToolStripMenuItem = New ToolStripMenuItem()
        ACAbilityToolStripMenuItem = New ToolStripMenuItem()
        ACStatusToolStripMenuItem = New ToolStripMenuItem()
        GetVersionToolStripMenuItem = New ToolStripMenuItem()
        RefreshDataToolStripMenuItem = New ToolStripMenuItem()
        SnapshotToolStripMenuItem = New ToolStripMenuItem()
        MonitorToolStripMenuItem = New ToolStripMenuItem()
        StartMonitorToolStripMenuItem = New ToolStripMenuItem()
        ExportExcelToolStripMenuItem = New ToolStripMenuItem()
        ChartToolStripMenuItem = New ToolStripMenuItem()
        TextBox1 = New TextBox()
        MenuStrip1.SuspendLayout()
        SuspendLayout()
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.Items.AddRange(New ToolStripItem() {FindControllerToolStripMenuItem, RefreshDataToolStripMenuItem, SnapshotToolStripMenuItem, MonitorToolStripMenuItem})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Size = New Size(800, 24)
        MenuStrip1.TabIndex = 0
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' FindControllerToolStripMenuItem
        ' 
        FindControllerToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {GetZonesToolStripMenuItem, ZoneStatusToolStripMenuItem, ACAbilityToolStripMenuItem, ACStatusToolStripMenuItem, GetVersionToolStripMenuItem})
        FindControllerToolStripMenuItem.Name = "FindControllerToolStripMenuItem"
        FindControllerToolStripMenuItem.Size = New Size(98, 20)
        FindControllerToolStripMenuItem.Text = "Find Controller"
        ' 
        ' GetZonesToolStripMenuItem
        ' 
        GetZonesToolStripMenuItem.Name = "GetZonesToolStripMenuItem"
        GetZonesToolStripMenuItem.Size = New Size(136, 22)
        GetZonesToolStripMenuItem.Text = "Get Zones"
        ' 
        ' ZoneStatusToolStripMenuItem
        ' 
        ZoneStatusToolStripMenuItem.Name = "ZoneStatusToolStripMenuItem"
        ZoneStatusToolStripMenuItem.Size = New Size(136, 22)
        ZoneStatusToolStripMenuItem.Text = "Zone Status"
        ' 
        ' ACAbilityToolStripMenuItem
        ' 
        ACAbilityToolStripMenuItem.Name = "ACAbilityToolStripMenuItem"
        ACAbilityToolStripMenuItem.Size = New Size(136, 22)
        ACAbilityToolStripMenuItem.Text = "AC ability"
        ' 
        ' ACStatusToolStripMenuItem
        ' 
        ACStatusToolStripMenuItem.Name = "ACStatusToolStripMenuItem"
        ACStatusToolStripMenuItem.Size = New Size(136, 22)
        ACStatusToolStripMenuItem.Text = "AC status"
        ' 
        ' GetVersionToolStripMenuItem
        ' 
        GetVersionToolStripMenuItem.Name = "GetVersionToolStripMenuItem"
        GetVersionToolStripMenuItem.Size = New Size(136, 22)
        GetVersionToolStripMenuItem.Text = "Get Version"
        ' 
        ' RefreshDataToolStripMenuItem
        ' 
        RefreshDataToolStripMenuItem.Name = "RefreshDataToolStripMenuItem"
        RefreshDataToolStripMenuItem.Size = New Size(85, 20)
        RefreshDataToolStripMenuItem.Text = "Refresh Data"
        ' 
        ' SnapshotToolStripMenuItem
        ' 
        SnapshotToolStripMenuItem.Name = "SnapshotToolStripMenuItem"
        SnapshotToolStripMenuItem.Size = New Size(68, 20)
        SnapshotToolStripMenuItem.Text = "Snapshot"
        ' 
        ' MonitorToolStripMenuItem
        ' 
        MonitorToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {StartMonitorToolStripMenuItem, ExportExcelToolStripMenuItem, ChartToolStripMenuItem})
        MonitorToolStripMenuItem.Name = "MonitorToolStripMenuItem"
        MonitorToolStripMenuItem.Size = New Size(62, 20)
        MonitorToolStripMenuItem.Text = "Monitor"
        ' 
        ' StartMonitorToolStripMenuItem
        ' 
        StartMonitorToolStripMenuItem.Name = "StartMonitorToolStripMenuItem"
        StartMonitorToolStripMenuItem.Size = New Size(180, 22)
        StartMonitorToolStripMenuItem.Text = "Start Monitor"
        ' 
        ' ExportExcelToolStripMenuItem
        ' 
        ExportExcelToolStripMenuItem.Name = "ExportExcelToolStripMenuItem"
        ExportExcelToolStripMenuItem.Size = New Size(180, 22)
        ExportExcelToolStripMenuItem.Text = "Export Excel"
        ' 
        ' ChartToolStripMenuItem
        ' 
        ChartToolStripMenuItem.Name = "ChartToolStripMenuItem"
        ChartToolStripMenuItem.Size = New Size(180, 22)
        ChartToolStripMenuItem.Text = "Make chart"
        ' 
        ' TextBox1
        ' 
        TextBox1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        TextBox1.Location = New Point(11, 34)
        TextBox1.Multiline = True
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(777, 404)
        TextBox1.TabIndex = 1
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(TextBox1)
        Controls.Add(MenuStrip1)
        MainMenuStrip = MenuStrip1
        Name = "Form1"
        Text = "Form1"
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FindControllerToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SnapshotToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents RefreshDataToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GetZonesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ZoneStatusToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ACAbilityToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ACStatusToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GetVersionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MonitorToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents StartMonitorToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExportExcelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ChartToolStripMenuItem As ToolStripMenuItem

End Class
