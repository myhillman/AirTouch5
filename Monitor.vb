Imports System.Drawing.Imaging
Imports System.Security.Policy
Imports System.Text
Imports Microsoft.Data.Sqlite
Imports Microsoft.Office.Interop.Excel
Imports OfficeOpenXml
Imports Svg

Module Monitor
    Private mainForm As Form
    Public zoneMonitorTimer As System.Timers.Timer = Nothing
    Const MonitorInterval = 10        ' monitor interval time in minutes
    Private Const connectionString As String = "Data Source=ZoneStatus.db;"
    Public Sub SetMainForm(f As Form)
        mainForm = f
    End Sub

    Public Sub StartZoneMonitoring()
        ZoneMonitorTimer_Elapsed()        ' take first sample straight away
        If zoneMonitorTimer Is Nothing Then
            zoneMonitorTimer = New System.Timers.Timer(MonitorInterval * 60 * 1000)  ' 10 minutes
            AddHandler zoneMonitorTimer.Elapsed, AddressOf ZoneMonitorTimer_Elapsed
            zoneMonitorTimer.AutoReset = True
            zoneMonitorTimer.Enabled = True
            Form1.AppendText("Zone monitoring started." & vbCrLf)
        End If
    End Sub

    Private Sub ZoneMonitorTimer_Elapsed()
        If mainForm.IsHandleCreated AndAlso Not mainForm.IsDisposed Then
            mainForm.Invoke(Sub()
                                ACStatus.GetACStatus()      ' get the AC status
                                Zones.GetZoneStatus()       ' get the zone statuses
                                SaveStatusesToDatabase()
                            End Sub)
        Else
            MsgBox("Timer elapsed error", vbCritical + vbOKOnly)
        End If
    End Sub

    Private Sub SaveStatusesToDatabase()
        Dim timestamp As DateTime = DateTime.Now, cmd As SqliteCommand
        Using conn As New SqliteConnection(connectionString)
            conn.Open()
            ' Create table of AC Status
            Dim createTableCmd As New SqliteCommand(
                "CREATE TABLE IF NOT EXISTS ACStatus (
	            id	INTEGER NOT NULL,
	            Date	TEXT,
	            ACPower	INTEGER,
	            SetPoint	INTEGER,
	            Temperature	NUMERIC(3,1),
	            PRIMARY KEY(id AUTOINCREMENT)
                )", conn)
            createTableCmd.ExecuteNonQuery()
            ' Add AC Status to the table
            cmd = New SqliteCommand("INSERT INTO ACStatus (date, ACPower, SetPoint, Temperature) VALUES (@date, @ACPower, @SetPoint, @Temperature)", conn)
            With cmd.Parameters
                .AddWithValue("@date", timestamp.ToString("yyyy-MM-dd HH:mm"))
                .AddWithValue("@ACPower", If(ACStatus.acStatusMsg.ACPower = ACPowerEnum.On Or ACStatus.acStatusMsg.ACPower = ACPowerEnum.OnAlways, 1, 0))
                .AddWithValue("@SetPoint", ACStatus.acStatusMsg.SetPoint)
                .AddWithValue("@Temperature", Math.Round(ACStatus.acStatusMsg.Temperature, 1))
            End With
            cmd.ExecuteNonQuery()
            ' Get the last inserted id
            Dim lastIdCmd As New SqliteCommand("SELECT last_insert_rowid()", conn)
            Dim lastId As Long = CLng(lastIdCmd.ExecuteScalar())

            ' create table of zone data if doesn't exist
            createTableCmd = New SqliteCommand("CREATE TABLE IF NOT EXISTS ZoneData (id INTEGER, ZoneNumber INTEGER, ZoneState BOOLEAN, DamperOpen INTEGER, SetPoint INTEGER, Temperature NUMERIC(3,1), Spill BOOLEAN)", conn)
            createTableCmd.ExecuteNonQuery()

            cmd = New SqliteCommand("INSERT INTO ZoneData (id, ZoneNumber, ZoneState, DamperOpen, SetPoint, Temperature, Spill) VALUES (@id, @ZoneNumber, @ZoneState, @DamperOpen, @SetPoint, @Temperature, @Spill)", conn)
            For Each zone In Zones.ZoneStatuses
                cmd.Parameters.Clear()
                cmd.Parameters.AddWithValue("@id", lastId)
                cmd.Parameters.AddWithValue("@ZoneNumber", zone.Key)
                cmd.Parameters.AddWithValue("@ZoneState", zone.Value.ZoneState)
                cmd.Parameters.AddWithValue("@DamperOpen", zone.Value.DamperOpen)
                cmd.Parameters.AddWithValue("@SetPoint", zone.Value.SetPoint)
                cmd.Parameters.AddWithValue("@Temperature", Math.Round(zone.Value.Temperature, 1))
                cmd.Parameters.AddWithValue("@Spill", If(zone.Value.Spill = SpillStatusEnum.Inactive, 0, 1))
                cmd.ExecuteNonQuery()
            Next
        End Using
    End Sub

    Public Sub ExportExcel()
        Dim sqlQuery As String = "select * from ZoneData join ACStatus using (id) order by id,ZoneNumber"
        Dim excelFilePath As String = "ZoneData.xlsx"

        ExcelPackage.License.SetNonCommercialPersonal("Marc Hillman")

        Using conn As New SqliteConnection(connectionString)
            conn.Open()
            Using cmd As New SqliteCommand(sqlQuery, conn)
                Using reader = cmd.ExecuteReader()
                    Using package As New ExcelPackage()
                        Dim worksheet = package.Workbook.Worksheets.Add("ZoneData")

                        ' Write column headers
                        For i As Integer = 0 To reader.FieldCount - 1
                            worksheet.Cells(1, i + 1).Value = reader.GetName(i)
                        Next

                        ' Write data rows
                        Dim row As Integer = 2
                        While reader.Read()
                            For col As Integer = 0 To reader.FieldCount - 1
                                worksheet.Cells(row, col + 1).Value = reader(col).ToString()
                            Next
                            row += 1
                        End While

                        ' Save to file
                        package.SaveAs(New IO.FileInfo(excelFilePath))
                    End Using
                End Using
            End Using
        End Using
    End Sub
    ''' <summary>
    ''' Generates an SVG chart visualizing zone monitoring data including temperature, setpoints, and damper status
    ''' </summary>
    ''' <param name="StartDate">The starting date/time for the data range</param>
    ''' <param name="EndDate">The ending date/time for the data range</param>
    ''' <returns>SVG markup as a string containing the generated chart</returns>
    ''' <remarks>
    ''' Features:
    ''' - Displays data for up to 8 zones plus AC unit
    ''' - Shows three metrics per zone: DamperOpen (blue), SetPoint (green dashed), Temperature (orange)
    ''' - Visualizes active/inactive states with opacity changes (100% for active, 50% for inactive)
    ''' - Implements seamless transitions between active/inactive segments
    ''' - Detects and renders gaps for data intervals >12 minutes
    ''' - Includes time axis, zone labels, and comprehensive legend
    ''' 
    ''' Data Requirements:
    ''' - Zone data must include: ZoneNumber, date, ZoneState, DamperOpen, SetPoint, Temperature
    ''' - AC unit data must include: date, ACPower, SetPoint, Temperature
    ''' 
    ''' Chart Layout:
    ''' - Width: 1290px (1200 + margins)
    ''' - Height: Dynamic based on zone count (100px per zone + margins)
    ''' - Each zone band: 100px tall with 10px separator
    ''' - Margins: Left 90px, Bottom 140px
    ''' 
    ''' Color Scheme:
    ''' - Damper/AC Status: Blue
    ''' - SetPoints: #32CD32 (LimeGreen) with dashed line
    ''' - Temperature: Red
    ''' 
    ''' Normalization:
    ''' - Temperatures normalized to 17-24°C range
    ''' - Damper values normalized to 0-100% range
    ''' </remarks>
    ''' <example>
    ''' Dim svgChart As String = GenerateZoneSvgChart(DateTime.Now.AddDays(-1), DateTime.Now)
    ''' File.WriteAllText("chart.svg", svgChart)
    ''' </example>

    Public Function GenerateZoneSvgChart(StartDate As Date) As String
        Const dashed = "10 10"       ' pattern for dashed lines
        Const MinTemp = 17          ' minimum temperature on chart
        Const MaxTemp = 24          ' maximum temperature on chart
        Const width = 1200          ' width of series plots
        Const heightPerZone = 100
        Const marginLeft = 110
        Const marginBottom = 140 ' Increased for legend and axis labels
        Const ZoneGap = 10        ' gap between zones
        Const TempScaleMargin = 40  ' area for scale margin
        Const MaxAllowedGapMinutes As Integer = MonitorInterval + 2 ' New constant for gap detection
        Const MaxAllowedGapSeconds As Double = MaxAllowedGapMinutes * 60

        ' Fixed SVG colors for each field
        Const damperColor As String = "Blue"      ' Blue for DamperOpen
        Const setPointColor As String = "LimeGreen"    ' Green for SetPoint
        Const tempColor As String = "Red"        ' Orange/Red for Temperature

        ' Color adjustment for inactive zones
        Const inactiveOpacity As String = "0.3"
        Const activeOpacity As String = "1.0"

        Dim x As Integer, y As Integer
        Dim ZoneTop As Integer, ZoneBottom As Integer, ZoneCenter As Integer     ' y values for top, middle and bottom of data series
        Dim ZoneHeight As Integer           ' height of data series

        Dim data As Data.DataTable = GetZoneDataTable(StartDate)
        ' Get min/max date for scaling
        Dim minDate = data.AsEnumerable().Min(Function(r) DateTime.Parse(r("Date").ToString()))
        Dim maxDate = data.AsEnumerable().Max(Function(r) DateTime.Parse(r("Date").ToString()))

        ' round the max/min dates to hour
        Dim minHour = New DateTime(minDate.Year, minDate.Month, minDate.Day, minDate.Hour, 0, 0)
        Dim maxHour = New DateTime(maxDate.Year, maxDate.Month, maxDate.Day, maxDate.Hour, 0, 0)
        maxHour = maxHour.AddHours(1)             ' round up to next hour
        Dim dateRange = (maxHour - minHour).TotalSeconds

        ' Get zone dimensions
        Dim minZone As Integer = data.AsEnumerable().Min(Function(r) r("ZoneNumber"))
        Dim maxZone As Integer = data.AsEnumerable().Max(Function(r) r("ZoneNumber"))
        Dim ZoneCount As Integer = maxZone - minZone + 1
        Dim height As Integer = heightPerZone * (ZoneCount + 1)   ' +1 for AC unit

        Dim sb As New StringBuilder()

        sb.AppendLine($"<svg width=""{width + marginLeft + TempScaleMargin}"" height=""{height + marginBottom}"" xmlns=""http://www.w3.org/2000/svg"">")
        ' Add copyright details
        sb.AppendLine("<!-- Copyright © " & Date.Now.Year & " Marc Hillman -->")

        sb.AppendLine("<style>.label {font-size:12px; font-family:sans-serif;}</style>")
        sb.AppendLine("<style>.comment {font-size:18px; font-family:sans-serif;}</style>")
        sb.AppendLine("<style>text {font-stretch: normal!important; font-style:normal!important; font-variant:normal!important;}</style>")
        sb.AppendLine("<style>.tempscale {font-size:8px; font-family:sans-serif; dominant-baseline=middle;text-anchor=start</style>")

        ' Add style for tooltips
        sb.AppendLine("
    <style>
      .tooltip {
        visibility: hidden;
        position: absolute;
        font: 12px sans-serif;
        background: white;
        color: #333;                     /* Dark gray text */
        border: 1px solid #1E90FF;       /* DodgerBlue border */
        border-radius: 3px;
        padding: 5px 8px;
        pointer-events: none;
        white-space: nowrap;
        z-index: 9999;
      }
      .tooltip.active {
        opacity: 1;
      }
      .tooltip-text {
        fill: #333;                      /* Text color */
        font-weight: bold;
      }
      .data-point:hover ~ .tooltip { visibility: visible; }
      .data-point{fill:transparent;stroke:none;stroke-width:1}
    </style>")

        ' Shade the area between every zone with a light grey rectangle
        For zone = minZone To maxZone + 1
            Dim yTop = (zone - minZone) * heightPerZone
            sb.AppendLine($"<rect x=""{marginLeft}"" y=""{yTop}"" width=""{width}"" height=""{ZoneGap}"" fill=""#cccccc"" />")
        Next

        ' Draw x&y axes
        sb.AppendLine($"<line x1=""{marginLeft}"" y1=""0"" x2=""{marginLeft}"" y2=""{height}"" stroke=""black""/>")         ' y axis
        sb.AppendLine($"<line x1=""{marginLeft}"" y1=""{height}"" x2=""{width + marginLeft}"" y2=""{height}"" stroke=""black""/>")  ' x axis

        ' Draw y axis labels
        For zone = minZone To maxZone + 1
            ZoneTop = (zone - minZone) * heightPerZone + ZoneGap        ' top of zone series area
            ZoneBottom = (zone - minZone + 1) * heightPerZone           ' bottom of zone series area
            ZoneCenter = (ZoneTop + ZoneBottom) / 2                     ' center of zone series area
            Dim ZoneLabel As String = If(zone <= maxZone, $"[{zone + 1}] {Zones.ZoneNames(zone)}", "AC unit")
            sb.AppendLine($"<text x=""5"" y=""{ZoneCenter}"" class=""label"">{ZoneLabel}</text>")

            ' Add temperature scale
            Dim scaleX = width + marginLeft + 5 ' Position scale 5px right of main chart
            ' Add vertical scale line
            sb.AppendLine($"<line x1=""{scaleX}"" y1=""{ZoneTop}"" x2=""{scaleX}"" y2=""{ZoneBottom}"" stroke=""#888"" stroke-width=""1""/>")
            ' Add temperature ticks and labels
            For temp = MinTemp To MaxTemp
                ' Calculate Y position for this temperature
                Dim normValue = (temp - MinTemp) / (MaxTemp - MinTemp)
                Dim yy = ZoneTop + (1.0 - normValue) * (ZoneBottom - ZoneTop)
                sb.AppendLine($"<line x1=""{scaleX}"" y1=""{yy}"" x2=""{scaleX + 5}"" y2=""{yy}"" stroke=""#888"" stroke-width=""1""/>")

                ' Add temperature label (right-aligned to the left of the tick)
                sb.AppendLine($"<text x=""{scaleX + 10}"" y=""{yy}"" class=""tempscale"" text-anchor=""start"" dominant-baseline=""middle"">{temp}°C</text>")
            Next
        Next

        ' Draw data series for each zone with gap detection
        For zone = minZone To maxZone
            Dim z = zone        ' don't use loop variable in a lambda
            Dim zoneData = data.AsEnumerable().Where(Function(r) CInt(r("ZoneNumber")) = z).OrderBy(Function(r) DateTime.Parse(r("date").ToString())).ToList()
            If zoneData.Count = 0 Then Continue For
            ZoneTop = (zone - minZone) * heightPerZone + ZoneGap + 1        ' top of zone series area
            ZoneBottom = (zone - minZone + 1) * heightPerZone - 1          ' bottom of zone series area
            ZoneHeight = ZoneBottom - ZoneTop                           ' height of zone series area

            For Each field In {"DamperOpen", "SetPoint", "Temperature"}
                Dim prevX As Double = -1
                Dim prevY As Double = -1
                Dim prevState As Boolean = False
                Dim prevTime As DateTime = DateTime.MinValue
                Dim segmentPoints As New List(Of String)
                Dim segmentStarted As Boolean = False

                For Each row In zoneData
                    Dim dt = DateTime.Parse(row("date").ToString())
                    x = marginLeft + ((dt - minHour).TotalSeconds / dateRange) * width
                    Dim normValue As Double = If(field = "DamperOpen",
                                       CDbl(row(field)) / 100.0,
                                       (CDbl(row(field)) - MinTemp) / (MaxTemp - MinTemp))
                    normValue = Math.Clamp(normValue, -0.1, 1.1)
                    y = ZoneTop + (1.0 - normValue) * ZoneHeight
                    Dim zoneState = CBool(row("ZoneState"))

                    ' Add hover circle (4px radius is a good hover target size)
                    sb.AppendLine($"<circle class=""data-point"" cx=""{x}"" cy=""{y}"" r=""4"">")
                    sb.AppendLine($"  <title>{field}: {row(field)} at {DateTime.Parse(row("date").ToString()):HH:mm}</title>")
                    sb.AppendLine("</circle>")

                    ' Add Spill indicator for DamperOpen series only
                    ' Spill was added later, so early records have a DBNull for Spill
                    If field = "DamperOpen" AndAlso Not IsDBNull(row("Spill")) AndAlso CInt(row("Spill")) = 1 Then
                        sb.AppendLine($"<circle cx=""{x}"" cy=""{y}"" r=""6"" fill=""none"" stroke=""red"" stroke-width=""1""/>")
                        sb.AppendLine($"<text x=""{x}"" y=""{y}"" font-size=""10"" text-anchor=""middle"" dominant-baseline=""central"" fill=""red"">S</text>")
                    End If

                    ' Check for time gap
                    Dim hasGap = prevTime <> DateTime.MinValue AndAlso (dt - prevTime).TotalSeconds > MaxAllowedGapSeconds

                    If Not segmentStarted Then
                        segmentPoints.Add($"{x:0.0},{y:0.0}")
                        segmentStarted = True
                        prevState = zoneState
                    ElseIf zoneState = prevState AndAlso Not hasGap Then
                        segmentPoints.Add($"{x:0.0},{y:0.0}")
                    Else
                        ' Draw previous segment including transition point (no gap)
                        If segmentPoints.Count >= 1 Then
                            Dim opacity = If(prevState, activeOpacity, inactiveOpacity)
                            Dim color = If(field = "DamperOpen", damperColor,
                                 If(field = "SetPoint", setPointColor, tempColor))
                            Dim strokeDash = If(field = "SetPoint", dashed, "")

                            ' Add transition point to close segment
                            If Not hasGap Then
                                segmentPoints.Add($"{x:0.0},{y:0.0}")
                            End If

                            sb.AppendLine($"<polyline points=""{String.Join(" ", segmentPoints)}"" fill=""none"" stroke=""{color}"" stroke-width=""2"" opacity=""{opacity}"" stroke-dasharray=""{strokeDash}""/>")

                            ' Start new segment (with current point if no gap)
                            segmentPoints = If(hasGap,
                                     New List(Of String),
                                     New List(Of String) From {$"{x:0.0},{y:0.0}"})
                        End If
                        prevState = zoneState
                    End If

                    prevX = x
                    prevY = y
                    prevTime = dt
                Next

                ' Draw last segment
                If segmentPoints.Count > 1 Then
                    Dim opacity = If(prevState, activeOpacity, inactiveOpacity)
                    Dim color = If(field = "DamperOpen", damperColor,
                         If(field = "SetPoint", setPointColor, tempColor))
                    Dim strokeDash = If(field = "SetPoint", dashed, "")
                    sb.AppendLine($"<polyline points=""{String.Join(" ", segmentPoints)}"" fill=""none"" stroke=""{color}"" stroke-width=""2"" opacity=""{opacity}"" stroke-dasharray=""{strokeDash}""/>")
                End If
            Next
        Next

        ' Draw data series for AC Unit
        Dim ACdata = GetACStatusTable(StartDate)
        ZoneTop = ZoneCount * heightPerZone + ZoneGap + 1        ' top of zone series area
        ZoneBottom = ZoneTop + heightPerZone - ZoneGap - 1     ' bottom of zone series area
        ZoneHeight = ZoneBottom - ZoneTop                               ' height of zone series area
        For Each field In {"ACPower", "SetPoint", "Temperature"}
            Dim prevX As Double = -1
            Dim prevY As Double = -1
            Dim prevState As Boolean = False
            Dim prevTime As DateTime = DateTime.MinValue
            Dim segmentPoints As New List(Of String)
            Dim segmentStarted As Boolean = False

            For Each row In ACdata.Rows
                Dim dt = DateTime.Parse(row("Date").ToString())
                x = marginLeft + ((dt - minHour).TotalSeconds / dateRange) * width
                Dim normValue As Double = If(field = "ACPower",
                                   CDbl(row("ACPower")),
                                   (CDbl(row(field)) - MinTemp) / (MaxTemp - MinTemp))
                normValue = Math.Clamp(normValue, -0.2, 1.2)        ' clamp, but allow 20% over/under shoot
                y = ZoneTop + (1.0 - normValue) * ZoneHeight
                Dim ACState = CBool(row("ACPower"))
                ' Add hover circle (4px radius is a good hover target size)
                sb.AppendLine($"<circle class=""data-point"" cx=""{x}"" cy=""{y}"" r=""4"">")
                sb.AppendLine($"  <title>{field}: {row(field)} at {DateTime.Parse(row("date").ToString()):HH:mm}</title>")
                sb.AppendLine("</circle>")
                ' Check for time gap
                Dim hasGap = prevTime <> DateTime.MinValue AndAlso (dt - prevTime).TotalSeconds > MaxAllowedGapSeconds

                If Not segmentStarted Then
                    segmentPoints.Add($"{x:0.0},{y:0.0}")
                    segmentStarted = True
                    prevState = ACState
                ElseIf ACState = prevState AndAlso Not hasGap Then
                    segmentPoints.Add($"{x:0.0},{y:0.0}")
                Else
                    ' Draw previous segment including transition point (no gap)
                    If segmentPoints.Count >= 1 Then
                        Dim opacity = If(prevState, activeOpacity, inactiveOpacity)
                        Dim color = If(field = "ACPower", damperColor,
                             If(field = "SetPoint", setPointColor, tempColor))
                        Dim strokeDash = If(field = "SetPoint", dashed, "")

                        ' Add transition point to close segment
                        If Not hasGap Then
                            segmentPoints.Add($"{x:0.0},{y:0.0}")
                        End If

                        sb.AppendLine($"<polyline points=""{String.Join(" ", segmentPoints)}"" fill=""none"" stroke=""{color}"" stroke-width=""2"" opacity=""{opacity}"" stroke-dasharray=""{strokeDash}""/>")

                        ' Start new segment (with current point if no gap)
                        segmentPoints = If(hasGap,
                                 New List(Of String),
                                 New List(Of String) From {$"{x:0.0},{y:0.0}"})
                    End If
                    prevState = ACState
                End If

                prevX = x
                prevY = y
                prevTime = dt
            Next

            ' Draw last segment
            If segmentPoints.Count > 1 Then
                Dim opacity = If(prevState, activeOpacity, inactiveOpacity)
                Dim color = If(field = "ACPower", damperColor,
                     If(field = "SetPoint", setPointColor, tempColor))
                Dim strokeDash = If(field = "SetPoint", dashed, "")
                sb.AppendLine($"<polyline points=""{String.Join(" ", segmentPoints)}"" fill=""none"" stroke=""{color}"" stroke-width=""2"" opacity=""{opacity}"" stroke-dasharray=""{strokeDash}""/>")
            End If
        Next

        ' Rest of the function remains unchanged (time axis labels, legend, etc.)
        Dim totalHours As Integer = CInt((maxHour - minHour).TotalHours)
        Dim interval = Math.Ceiling(totalHours / 12)
        Dim AxisRange = (maxHour - minHour).TotalSeconds
        For i = 0 To totalHours Step interval
            Dim dt = minHour.AddHours(i)
            Dim frac = (dt - minHour).TotalSeconds / AxisRange
            x = marginLeft + frac * width
            sb.AppendLine($"<text x=""{x}"" y=""{height + 20}"" class=""label"" text-anchor=""middle"">{dt:dd-MM HH:mm}</text>")        ' draw label
            sb.AppendLine($"<line x1=""{x}"" y1=""{height}"" x2=""{x}"" y2=""{height + 5}"" stroke=""black""/>")                        ' draw tick
            ' Draw vertical dotted line
            If i > 0 And i < totalHours Then
                sb.AppendLine($"<line x1=""{x}"" y1=""{0}"" x2=""{x}"" y2=""{heightPerZone * (ZoneCount + 1)}"" stroke=""black"" opacity=""0.5"" stroke-dasharray=""10 20""/>")
            End If
        Next

        Dim legendX = marginLeft + 20
        Dim legendY = height + 60
        Dim legendSpacing = 30
        ' legend box
        sb.AppendLine($"<rect x=""{legendX - 10}"" y=""{legendY - 15}"" width=""400"" height=""90"" fill=""#fff"" stroke=""#ccc""/>")
        ' Damper Open
        sb.AppendLine($"<polyline points=""{legendX},{legendY} {legendX + 30},{legendY}"" fill=""none"" stroke=""{damperColor}"" stroke-width=""3""/>")
        sb.AppendLine($"<text x=""{legendX + 40}"" y=""{legendY + 5}"" class=""label"">Damper Open/AC On (0-100%, blue, solid)</text>")
        sb.AppendLine($"<polyline points=""{legendX},{legendY + legendSpacing} {legendX + 30},{legendY + legendSpacing}"" fill=""none"" stroke=""{setPointColor}"" stroke-width=""3"" stroke-dasharray=""{dashed}""/>")
        ' Spill Active indicator
        sb.AppendLine($"<circle cx=""{legendX + 289}"" cy=""{legendY + 2}"" r=""6"" fill=""none"" stroke=""red"" stroke-width=""1""/>")
        sb.AppendLine($"<text x=""{legendX + 285}"" y=""{legendY + 5}"" class=""label"" stroke=""red"">S</text>")
        sb.AppendLine($"<text x=""{legendX + 300}"" y=""{legendY + 5}"" class=""label"">Spill Active</text>")

        ' Set Point
        sb.AppendLine($"<text x=""{legendX + 40}"" y=""{legendY + legendSpacing + 5}"" class=""label"">Set Point ({MinTemp}-{MaxTemp}°C, green, dotted)</text>")
        ' Temperature
        sb.AppendLine($"<polyline points=""{legendX},{legendY + legendSpacing * 2} {legendX + 30},{legendY + legendSpacing * 2}"" fill=""none"" stroke=""{tempColor}"" stroke-width=""3""/>")
        sb.AppendLine($"<text x=""{legendX + 40}"" y=""{legendY + legendSpacing * 2 + 5}"" class=""label"">Temperature ({MinTemp}-{MaxTemp}°C, red, solid)</text>")
        ' Add some notes
        sb.AppendLine($"<text x=""{legendX + 400}"" y=""{legendY}"" class=""comment"">Chart of parameters for AirTouch console {AirTouch5Console.AirTouchID} V{VersionInfo.AnnotatedVersion} on {Now():yyyy-MM-dd HH:mm}</text>")
        sb.AppendLine($"<text x=""{legendX + 400}"" y=""{legendY + 20}"" class=""comment"">AirTouch Diagnostic tool ©2025 Marc Hillman</text>")
        Dim OnTime = CalculateACOnTimePercentage(ACdata)
        sb.AppendLine($"<text x=""{legendX + 400}"" y=""{legendY + 40}"" class=""comment"">AC On time {OnTime:0.0}%</text>")
        ' Add tooltip container (place near end before closing </svg>)
        sb.AppendLine("<rect id=""tooltip-box"" class=""tooltip"" width=""120"" height=""50"" rx=""5"" ry=""5""></rect>")
        sb.AppendLine("<text id=""tooltip-text"" class=""tooltip"" font-size=""11""></text>")
        sb.AppendLine("</svg>")
        Return sb.ToString()
    End Function

    Public Function CalculateACOnTimePercentage(ACdata As Data.DataTable) As Double
        ' Get AC data
        If ACdata Is Nothing OrElse ACdata.Rows.Count = 0 Then Throw New Exception("AC data is missing or empty")
        Dim StartDate = ACdata.AsEnumerable().Min(Function(r) DateTime.Parse(r("date").ToString()))
        Dim EndDate = ACdata.AsEnumerable().Max(Function(r) DateTime.Parse(r("date").ToString()))

        Dim totalTimeSpan As TimeSpan = EndDate - StartDate
        Dim totalSeconds As Double = totalTimeSpan.TotalSeconds
        If totalSeconds <= 0.0 Then Throw New Exception($"Total Seconds is {totalSeconds}")
        Dim onTimeSeconds As Double = 0
        Dim prevTime As DateTime = StartDate
        Dim prevPower As Integer = 0 ' Assume AC starts off

        ' Calculate on-time
        For Each row As DataRow In ACdata.Rows
            Dim currentTime As DateTime = DateTime.Parse(row("date").ToString())
            Dim currentPower As Integer = CInt(row("ACPower"))

            ' If AC was on during this period, add to onTimeSeconds
            If prevPower = 1 Then
                onTimeSeconds += (currentTime - prevTime).TotalSeconds
            End If

            prevTime = currentTime
            prevPower = currentPower
        Next

        ' Handle the last segment (from last reading to EndDate)
        If prevPower = 1 Then
            onTimeSeconds += (EndDate - prevTime).TotalSeconds
        End If

        ' Calculate percentage
        Dim percentageOnTime As Double = (onTimeSeconds / totalSeconds) * 100
        Return percentageOnTime
    End Function
    Public Function ConvertSvgStringToPng(svgString As String, outputPath As String) As Boolean
        Try
            ' Load the SVG from string
            Dim svgDoc = SvgDocument.FromSvg(Of SvgDocument)(svgString)

            ' Remove tooltip-related elements as they are useless on a PNG
            RemoveTooltips(svgDoc)

            ' Create bitmap with white background
            Using bitmap As New Bitmap(CInt(svgDoc.Width), CInt(svgDoc.Height))
                ' Fill with white background
                Using g As Graphics = Graphics.FromImage(bitmap)
                    g.Clear(Color.White)

                    ' Draw SVG onto the white background
                    svgDoc.Draw(g)
                End Using

                ' Save as PNG
                bitmap.Save(outputPath, ImageFormat.Png)
            End Using
            Form1.AppendText($"PNG file written to {outputPath}{vbCrLf}")
            Return True
        Catch ex As Exception
            ' Handle conversion errors
            Debug.WriteLine($"Error converting SVG to PNG: {ex.Message}")
            Return False
        End Try
    End Function
    Private Sub RemoveTooltips(svgDoc As SvgDocument)
        ' Remove tooltip container if it exists
        Dim tooltipBox = svgDoc.Children.FirstOrDefault(Function(c)
                                                            Return TypeOf c Is SvgRectangle AndAlso
               c.CustomAttributes.Any(Function(a) a.Key = "class" AndAlso a.Value = "tooltip")
                                                        End Function)

        If tooltipBox IsNot Nothing Then
            svgDoc.Children.Remove(tooltipBox)
        End If

        ' Remove tooltip text if it exists
        Dim tooltipText = svgDoc.Children.FirstOrDefault(Function(c)
                                                             Return TypeOf c Is SvgText AndAlso
               c.CustomAttributes.Any(Function(a) a.Key = "class" AndAlso a.Value = "tooltip")
                                                         End Function)

        If tooltipText IsNot Nothing Then
            svgDoc.Children.Remove(tooltipText)
        End If

        ' Remove all elements with class "data-point" (hover targets)
        Dim dataPoints = svgDoc.Children.Where(Function(c)
                                                   Return c.CustomAttributes.Any(Function(a) a.Key = "class" AndAlso a.Value = "data-point")
                                               End Function).ToList()

        For Each point In dataPoints
            svgDoc.Children.Remove(point)
        Next

        ' Remove tooltip titles
        Dim titles = svgDoc.Descendants().OfType(Of SvgTitle)().ToList()
        For Each title In titles
            title.Parent?.Children.Remove(title)
        Next
    End Sub
    Public Function GetZoneDataTable(StartDate As Date) As Data.DataTable
        Dim dt As New Data.DataTable()

        Dim sqlQuery As String = $"SELECT date,ZoneNumber,ZoneState,DamperOpen,ZoneData.SetPoint,ZoneData.Temperature, Spill FROM ZoneData JOIN ACStatus USING (`id`) WHERE DATE(`date`) = '{StartDate:yyyy-MM-dd}' order by `date`,`ZoneNumber`"

        Using conn As New SqliteConnection(connectionString)
            conn.Open()
            Using cmd As New SqliteCommand(sqlQuery, conn)
                Using reader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using

        Return dt
    End Function
    Public Function GetACStatusTable(StartDate As Date) As Data.DataTable
        Dim dt As New Data.DataTable()

        Dim sqlQuery As String = $"SELECT * FROM ACStatus WHERE DATE(`date`) = '{StartDate:yyyy-MM-dd}' order by id"

        Using conn As New SqliteConnection(connectionString)
            conn.Open()
            Using cmd As New SqliteCommand(sqlQuery, conn)
                Using reader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using

        Return dt
    End Function
End Module