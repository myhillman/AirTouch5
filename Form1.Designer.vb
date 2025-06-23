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
        GetVersionToolStripMenuItem = New ToolStripMenuItem()
        GetZonesToolStripMenuItem = New ToolStripMenuItem()
        ZoneStatusToolStripMenuItem = New ToolStripMenuItem()
        ACAbilityToolStripMenuItem = New ToolStripMenuItem()
        ACStatusToolStripMenuItem = New ToolStripMenuItem()
        SnapshotToolStripMenuItem = New ToolStripMenuItem()
        MenuStrip1.SuspendLayout()
        SuspendLayout()
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.Items.AddRange(New ToolStripItem() {FindControllerToolStripMenuItem, GetVersionToolStripMenuItem, GetZonesToolStripMenuItem, ZoneStatusToolStripMenuItem, ACAbilityToolStripMenuItem, ACStatusToolStripMenuItem, SnapshotToolStripMenuItem})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Size = New Size(800, 24)
        MenuStrip1.TabIndex = 0
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' FindControllerToolStripMenuItem
        ' 
        FindControllerToolStripMenuItem.Name = "FindControllerToolStripMenuItem"
        FindControllerToolStripMenuItem.Size = New Size(98, 20)
        FindControllerToolStripMenuItem.Text = "Find Controller"
        ' 
        ' GetVersionToolStripMenuItem
        ' 
        GetVersionToolStripMenuItem.Name = "GetVersionToolStripMenuItem"
        GetVersionToolStripMenuItem.Size = New Size(78, 20)
        GetVersionToolStripMenuItem.Text = "Get Version"
        ' 
        ' GetZonesToolStripMenuItem
        ' 
        GetZonesToolStripMenuItem.Name = "GetZonesToolStripMenuItem"
        GetZonesToolStripMenuItem.Size = New Size(72, 20)
        GetZonesToolStripMenuItem.Text = "Get Zones"
        ' 
        ' ZoneStatusToolStripMenuItem
        ' 
        ZoneStatusToolStripMenuItem.Name = "ZoneStatusToolStripMenuItem"
        ZoneStatusToolStripMenuItem.Size = New Size(81, 20)
        ZoneStatusToolStripMenuItem.Text = "Zone Status"
        ' 
        ' ACAbilityToolStripMenuItem
        ' 
        ACAbilityToolStripMenuItem.Name = "ACAbilityToolStripMenuItem"
        ACAbilityToolStripMenuItem.Size = New Size(70, 20)
        ACAbilityToolStripMenuItem.Text = "AC ability"
        ' 
        ' ACStatusToolStripMenuItem
        ' 
        ACStatusToolStripMenuItem.Name = "ACStatusToolStripMenuItem"
        ACStatusToolStripMenuItem.Size = New Size(69, 20)
        ACStatusToolStripMenuItem.Text = "AC status"
        ' 
        ' SnapshotToolStripMenuItem
        ' 
        SnapshotToolStripMenuItem.Name = "SnapshotToolStripMenuItem"
        SnapshotToolStripMenuItem.Size = New Size(68, 20)
        SnapshotToolStripMenuItem.Text = "Snapshot"
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
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
    Friend WithEvents GetVersionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GetZonesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ZoneStatusToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ACAbilityToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ACStatusToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SnapshotToolStripMenuItem As ToolStripMenuItem

End Class
