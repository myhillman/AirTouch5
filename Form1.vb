Imports System.IO
Imports System.Text

Public Class Form1
    Private Async Sub FindControllerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindControllerToolStripMenuItem.Click
        Await DiscoverConsole()
        If AirTouch5Console.Connected Then
            RefreshData
        End If
    End Sub


    Private Sub GetVersionToolStripMenuItem_Click(sender As Object, e As EventArgs)
        GetVersion()
    End Sub
    Private Sub GetZonesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        GetZones()
    End Sub

    Private Sub ZoneStatusToolStripMenuItem_Click(sender As Object, e As EventArgs)
        GetZoneStatus()
    End Sub

    Private Sub ACAbilityToolStripMenuItem_Click(sender As Object, e As EventArgs)
        GetACability()
    End Sub

    Private Sub ACStatusToolStripMenuItem_Click(sender As Object, e As EventArgs)
        GetACStatus()
    End Sub

    Private Sub SnapshotToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SnapshotToolStripMenuItem.Click
        MakeSnapshot()
    End Sub

    Private Sub RefreshDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshDataToolStripMenuItem.Click
        RefreshData
    End Sub
    Sub RefreshData()
        If AirTouch5Console.Connected Then
            GetVersion()
            GetZones()
            GetZoneStatus()
            GetACability()
            GetACStatus()
        Else
            TextBox1.AppendText("Console not connected. Cannot refresh data." & vbCrLf)
        End If
    End Sub
End Class