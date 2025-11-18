Imports System
Public Class Form1
    Private Async Sub FindControllerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindControllerToolStripMenuItem.Click
        Await DiscoverConsole()
        If AirTouch5Console.Connected Then
            GetVersion()
            RefreshData()
        End If
    End Sub
    Private Sub GetVersionToolStripMenuItem_Click(sender As Object, e As EventArgs)
        GetVersion()
    End Sub
    Private Sub GetZonesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        GetZoneNames()
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
        RefreshData()
    End Sub
    Sub RefreshData()
        If AirTouch5Console.Connected Then
            GetACability()
            GetACStatus()
            GetZoneNames()
            GetZoneStatus()
            AppendText("Refresh complete" & vbCrLf)
        Else
            AppendText("Console not connected. Cannot refresh data." & vbCrLf)
        End If
    End Sub

    Private Async Sub StartMonitorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartMonitorToolStripMenuItem.Click
        If Monitor.zoneMonitorTimer IsNot Nothing Then
            MsgBox("You cannot start monitoring when it is already running", vbCritical + vbOKOnly, "Monitoring Already Running")
            Exit Sub
        End If
        If Not AirTouch5Console.Connected Then
            Await DiscoverConsole()
            If AirTouch5Console.Connected Then
                GetVersion()
                RefreshData()
            End If
        End If
        Monitor.SetMainForm(Me)
        Me.Text = $"{Me.Text} (Monitoring)"
        Monitor.StartZoneMonitoring()
    End Sub

    Private Sub ExportExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportExcelToolStripMenuItem.Click
        Monitor.ExportExcel()
    End Sub

    Private Async Sub ChartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChartToolStripMenuItem.Click
        ' If we have no zone name data, connect and get it.
        If Not AirTouch5Console.Connected Then
            Await DiscoverConsole()
            GetVersion()
        End If

        ' Read the start and end dates from the user
        Dim startDate As Date
        Dim DefaultDate = Strings.Format(Now(), "yyyy-MM-dd")

        ' Input and validate start date
        While True
            Dim startInput = InputBox("Enter start date (YYYY-MM-DD): ", "Start Date", DefaultDate)
            If Date.TryParseExact(startInput, "yyyy-MM-dd",
                                 System.Globalization.CultureInfo.InvariantCulture,
                                 System.Globalization.DateTimeStyles.None, startDate) Then
                Exit While
            End If
        End While

        ' Get the zone names
        If ZoneNames.Count = 0 Then
            GetZoneNames()
        End If
        Dim svg = GenerateZoneSvgChart(startDate)      ' generate the chart
        Dim DateString = Format(startDate, "yyyy-MM-dd")
        Dim svgPath = $"ZoneChart_{DateString}.svg"
        System.IO.File.WriteAllText(svgPath, svg)
        Monitor.ConvertSvgStringToPng(svg, $"ZoneChart_{DateString}.png")         ' save as png
        AppendText($"SVG file written to {svgPath}{vbCrLf}")
        ' Open with default browser (works in .NET Framework and .NET Core)
        Try
            Dim psi As New ProcessStartInfo With {
                .FileName = svgPath,
                .UseShellExecute = True
            }
            Process.Start(psi)
        Catch ex As Exception
            MessageBox.Show("Could not open SVG file: " & ex.Message)
        End Try
    End Sub
    Public Sub AppendText(text As String)
        ' Append text to textbox1 in a thread safe way
        If TextBox1.InvokeRequired Then
            TextBox1.BeginInvoke(Sub() TextBox1.AppendText(text))
        Else
            TextBox1.AppendText(text)
        End If
    End Sub

    Private Sub MakeAllChartsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeAllChartsToolStripMenuItem.Click
        ' (re)make charts for all days
        Dim dt As Date = DateTime.Parse("2025-07-19") ' Start date for the charts
        Dim endd As Date = Now
        While dt <= endd
            Dim svg = GenerateZoneSvgChart(dt)      ' generate the chart
            Dim DateString = Format(dt, "yyyy-MM-dd")
            Dim svgPath = $"ZoneChart_{DateString}.svg"
            System.IO.File.WriteAllText(svgPath, svg)
            Monitor.ConvertSvgStringToPng(svg, $"ZoneChart_{DateString}.png")         ' save as png
            AppendText($"SVG file written to {svgPath}{vbCrLf}")
            dt = dt.AddDays(1)
        End While
    End Sub
End Class