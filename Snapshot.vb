Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
''' <summary>
''' Module Snapshot
''' Purpose: Generates a snapshot of the AirTouch A/C system status and capabilities
''' Features:
''' - Creates an HTML report with A/C ability, status, and zone information
''' - Formats data in a single-line table for clarity
''' - Opens the generated report in the default web browser
''' </summary>
Friend Module Snapshot
    Sub MakeSnapshot()
        ' Define the HTML file path where the snapshot will be saved
        Dim filePath As String = $"ACSummary_{Now:yyyyMMdd_HHmmss}.html"

        ' Create and write HTML content to file using StreamWriter
        Using writer As New StreamWriter(filePath)
            ' Write HTML document header with CSS styling for single-line table borders
            writer.WriteLine("<!DOCTYPE html>" & vbCrLf &
                       "<html>" & vbCrLf &
                       "<head>" & vbCrLf &
                       "<title>A/C Snapshot</title>" & vbCrLf &
                       "<style>" & vbCrLf &
                       "        table.single-line { border-collapse: collapse; }" & vbCrLf &
                       "        table.single-line th, table.single-line td { border: 1px solid black; padding: 2px; text-align: center; }" & vbCrLf &
                       "    html, body {" & vbCrLf &
                        "   Font-family: Arial, Helvetica, sans-serif;" & vbCrLf &
                        "   Font-Size:  10pt;" & vbCrLf &
                        "   line-height:   1.5;" & vbCrLf &
                        "   -webkit-font-smoothing: antialiased;" & vbCrLf &
                        "   Text-rendering: optimizeLegibility; }" & vbCrLf &
                       "</style>" & vbCrLf &
                       "</head>" & vbCrLf &
                       "<body>" & vbCrLf)

            ' Write report header with timestamp and system information
            writer.WriteLine("<h1>AirTouch A/C Summary</h1>" & vbCrLf)
            writer.WriteLine($"Snapshot taken at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}<br>{vbCrLf}")
            writer.WriteLine($"<p>Console: ID={AirTouch5Console.ConsoleID} System ID={AirTouch5Console.AirTouchID} IP={AirTouch5Console.IP} Version={VersionInfo.VersionText} ({If(VersionInfo.UpdateSign = 0, "Latest", "Update")})</p>{vbCrLf}")
            writer.WriteLine($"Installed address: 28 Murndal Drive, Donvale. Vic 3111.  Contact Marc Hillman 0432 686 808{vbCrLf}")

            ' ========== A/C Ability Section ==========
            writer.WriteLine("<h2>A/C Ability</h2>" & vbCrLf)

            ' Create table header with column groupings for modes and fan speeds
            writer.WriteLine("<table class='single-line'>" & vbCrLf &
            "<tr>" &
            "<th></th>" & ' Empty cell for AC number
            "<th></th>" & ' Empty cell for AC name
            "<th></th>" & ' Empty cell for Start Zone
            "<th></th>" & ' Empty cell for Zone Count
            "<th colspan=5>Modes</th>" & ' Group header for operation modes
            "<th colspan=8>Fan Speed</th>" & ' Group header for fan speeds
            "<th></th>" & ' Empty cell for Min Cool
            "<th></th>" & ' Empty cell for Max Cool
            "<th></th>" & ' Empty cell for Min Heat
            "<th></th>" & ' Empty cell for Max Heat
            "</tr>" & vbCrLf)

            ' Write column headers for A/C Ability table
            writer.WriteLine("<tr>" &
                            "<th>AC<br>number</th>" &
                            "<th>AC name</th>" &
                            "<th>Start<br>Zone</th>" &
                            "<th>Zone<br>Count</th>" &
                            "<th>Cool</th>" &
                            "<th>Fan</th>" &
                            "<th>Dry</th>" &
                            "<th>Heat</th>" &
                            "<th>Auto</th>" &
                            "<th>IA</th>" &
                            "<th>Turbo</th>" &
                            "<th>Powerful</th>" &
                            "<th>High</th>" &
                            "<th>Medium</th>" &
                            "<th>Low</th>" &
                            "<th>Quiet</th>" &
                            "<th>Auto</th>" &
                            "<th>Min Cool<br>Set point</th>" &
                            "<th>Max Cool<br>Set point</th>" &
                            "<th>Min Heat<br>Set point</th>" &
                            "<th>Max Heat<br>Set point</th>" &
                            "</tr>" & vbCrLf)

            ' Write A/C Ability data row
            writer.WriteLine("<tr>" &
                            "<td>" & acData.ACnumber & "</td>" &
                            "<td>" & acData.ACName & "</td>" &
                            "<td>" & acData.StartZone & "</td>" &
                            "<td>" & acData.ZoneCount & "</td>" &
                            "<td>" & acData.CoolMode.ToString() & "</td>" &
                            "<td>" & acData.FanMode.ToString() & "</td>" &
                            "<td>" & acData.DryMode.ToString() & "</td>" &
                            "<td>" & acData.HeatMode.ToString() & "</td>" &
                            "<td>" & acData.AutoMode.ToString() & "</td>" &
                            "<td>" & acData.FanSpeedIA.ToString() & "</td>" &
                            "<td>" & acData.FanSpeedTurbo.ToString() & "</td>" &
                            "<td>" & acData.FanSpeedPowerful.ToString() & "</td>" &
                            "<td>" & acData.FanSpeedHigh.ToString() & "</td>" &
                            "<td>" & acData.FanSpeedMedium.ToString() & "</td>" &
                            "<td>" & acData.FanSpeedLow.ToString() & "</td>" &
                            "<td>" & acData.FanSpeedQuiet.ToString() & "</td>" &
                            "<td>" & acData.FanspeedAuto.ToString() & "</td>" &
                            "<td>" & acData.MinCoolSetPoint.ToString() & "</td>" &
                            "<td>" & acData.MaxCoolSetPoint.ToString() & "</td>" &
                            "<td>" & acData.MinHeatSetPoint.ToString() & "</td>" &
                            "<td>" & acData.MaxHeatSetPoint.ToString() & "</td>" &
                            "</tr>" & vbCrLf)
            writer.WriteLine("</table>" & vbCrLf)

            ' ========== A/C Status Section ==========
            writer.WriteLine("<h2>A/C Status</h2>" & vbCrLf)
            writer.WriteLine("<table class='single-line'>" & vbCrLf &
                       "<tr>" &
                       "<th>AC Power</th>" &
                       "<th>Number</th>" &
                       "<th>Mode</th>" &
                       "<th>Fan Speed</th>" &
                       "<th>Set Point</th>" &
                       "<th>Temperature</th>" &
                       "<th>Turbo</th>" &
                       "<th>Bypass</th>" &
                       "<th>Spill</th>" &
                       "<th>Timer</th>" &
                       "<th>Defrost</th>" &
                       "</tr>" & vbCrLf)

            ' Write A/C Status data row
            writer.WriteLine("<tr>" & vbCrLf &
                       "<td>" & acStatusMsg.ACPower.ToString() & "</td>" &
                       "<td>" & acStatusMsg.Number.ToString() & "</td>" &
                       "<td>" & acStatusMsg.Mode.ToString() & "</td>" &
                       "<td>" & acStatusMsg.FanSpeed.ToString() & "</td>" &
                       "<td>" & acStatusMsg.SetPoint.ToString() & "</td>" &
                       "<td>" & acStatusMsg.Temperature.ToString() & "</td>" &
                       "<td>" & acStatusMsg.Turbo.ToString() & "</td>" &
                       "<td>" & acStatusMsg.Bypass.ToString() & "</td>" &
                       "<td>" & acStatusMsg.Spill.ToString() & "</td>" &
                       "<td>" & acStatusMsg.TimerSet.ToString() & "</td>" &
                       "<td>" & acStatusMsg.Defrost.ToString() & "</td>" &
                       "</tr>" & vbCrLf)
            writer.WriteLine("</table>" & vbCrLf)
            writer.WriteLine("<p>Note: Turbo & Timer seem to be reporting incorrectly</p>")

            ' ========== Zones Section ==========
            writer.WriteLine("<h2>Zones</h2>" & vbCrLf)
            writer.WriteLine("<table class='single-line'>" & vbCrLf &
                       "<tr>" & vbCrLf &
                       "<th>Zone<br>Number</th>" &
                       "<th>Zone Name</th>" &
                       "<th>Power</th>" &
                       "<th>Control<br>Method</th>" &
                       "<th>Damper<br>Open %</th>" &
                       "<th>Set<br>Point</th>" &
                       "<th>Temperature</th>" &
                       "<th>Sensor</th>" &
                       "<th>Spill</th>" &
                       "<th>Battery</th>" &
                       "</tr>" & vbCrLf)

            ' Write data rows for each zone
            For Each item In ZoneStatuses
                Dim zoneStatus = item.Value
                writer.WriteLine("<tr>" & vbCrLf &
                           "<td>" & zoneStatus.ZoneNumber.ToString() & "</td>" &
                           "<td>" & ZoneNames(zoneStatus.ZoneNumber) & "</td>" &
                           "<td>" & zoneStatus.ZoneState.ToString() & "</td>" &
                           "<td>" & zoneStatus.ControlMethod.ToString() & "</td>" &
                           "<td>" & zoneStatus.DamperOpen.ToString() & "</td>" &
                           "<td>" & zoneStatus.SetPoint.ToString() & "</td>" &
                           "<td>" & zoneStatus.Temperature.ToString() & "</td>" &
                           "<td>" & zoneStatus.Sensor.ToString() & "</td>" &
                           "<td>" & zoneStatus.Spill.ToString() & "</td>" &
                           "<td>" & zoneStatus.Battery.ToString() & "</td>" &
                           "</tr>" & vbCrLf)
            Next
            writer.WriteLine("</table>" & vbCrLf)

            ' Close HTML document
            writer.WriteLine("</body></html>")
        End Using

        ' Open the generated HTML file in Chrome browser
        Dim psi As New ProcessStartInfo() With {
    .FileName = filePath,
    .UseShellExecute = True ' Required for opening with default application
}
        Try
            Process.Start(psi)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub
End Module
