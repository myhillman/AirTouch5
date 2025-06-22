Imports System.Text

Public Class Form1
    Private Async Sub FindControllerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindControllerToolStripMenuItem.Click
        Await DiscoverConsole()
        If AirTouch5Console.Connected Then
            GetZones()
            GetZoneStatus()
        End If
    End Sub


    Private Sub GetVersionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetVersionToolStripMenuItem.Click
        GetVersion
    End Sub
    Private Sub GetZonesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetZonesToolStripMenuItem.Click
        GetZones
    End Sub

    Private Sub ZoneStatusToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoneStatusToolStripMenuItem.Click
        GetZoneStatus
    End Sub

    Private Sub ACAbilityToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ACAbilityToolStripMenuItem.Click
        GetACability
    End Sub
End Class