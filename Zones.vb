﻿Imports System.Text
Imports AirTouch5.Form1

' ==============================================================
' Module: Zones
' Purpose: Manages zone information and status for AirTouch5 system
' Features:
' - Maintains dictionary of zone names and statuses
' - Provides enums for various zone states and statuses
' - Handles communication with AirTouch console to get zone data
' ==============================================================
Friend Module Zones
    ' Dictionary to store zone names with zone number as key
    Public ZoneNames As New Dictionary(Of Integer, String)

    ' Dictionary to store zone statuses with zone number as key
    Public ZoneStatuses As New Dictionary(Of Integer, ZoneStatusMessage)

    ' Enum: ZoneStateEnum
    ' Purpose: Defines possible power states for a zone
    Enum ZoneStateEnum
        Off = 0     ' Zone is turned off
        [On] = 1    ' Zone is powered on (On is a reserved word, hence brackets)
        Turbo = 3   ' Zone is in turbo mode
    End Enum

    ' Enum: ControlMethodEnum
    ' Purpose: Defines control methods for zones
    Enum ControlMethodEnum
        Percentage = 0   ' Zone controlled by percentage open
        Temperature = 1  ' Zone controlled by temperature
    End Enum

    ' Enum: SensorStatusEnum
    ' Purpose: Defines status of temperature sensor
    Enum SensorStatusEnum
        Absent = 0  ' No temperature sensor present
        Present = 1  ' Temperature sensor is present
    End Enum

    ' Enum: SpillStatusEnum
    ' Purpose: Defines spill status for zone
    Enum SpillStatusEnum
        Inactive = 0 ' Spill is inactive
        Active = 1   ' Spill is active
    End Enum

    ' Enum: BatteryStatusEnum
    ' Purpose: Defines battery status for zone
    Enum BatteryStatusEnum
        Normal = 0 ' Battery level is normal
        Low = 1    ' Battery level is low
    End Enum

    ' Structure: ZoneStatusMessage
    ' Purpose: Represents the complete status of a zone
    Public Structure ZoneStatusMessage
        Public ZoneState As ZoneStateEnum ' Current power state
        Public ZoneNumber As Byte                   ' Zone number (0-15)
        Public ControlMethod As ControlMethodEnum   ' Current control method
        Public DamperOpen As Byte                   ' Percentage open (0-100%)
        Public SetPoint As Short                    ' Temperature set point
        Public Sensor As SensorStatusEnum           ' Sensor presence status
        Public Temperature As Single                ' Current temperature
        Public Spill As SpillStatusEnum             ' Spill status
        Public Battery As BatteryStatusEnum         ' Battery status

        ' Method: Parse
        ' Purpose: Converts raw byte data into ZoneStatusMessage structure
        ' Parameters:
        '   data - Byte array containing zone status information
        Public Shared Function Parse(data As Byte(), index As Integer) As ZoneStatusMessage
            Return New ZoneStatusMessage With {
                .ZoneState = CType((data(index) >> 6) And &H3, ZoneStateEnum),
                .ZoneNumber = data(index) And &H3F,
                .ControlMethod = CType((data(index + 1) >> 7) And &H1, ControlMethodEnum),
                .DamperOpen = data(index + 1) And &H7F,
                .SetPoint = If(data(index + 2) = &HFF, Nothing, (data(index + 2) + 100) / 10),
                .Sensor = CType((data(index + 3) >> 7) And &H1, SensorStatusEnum),
                .Temperature = (((CInt(data(index + 4) And &H7) << 8) Or data(index + 5)) - 500) / 10.0,
                .Spill = CType((data(index + 6) >> 1) And &H1, SpillStatusEnum),
                .Battery = CType(data(index + 6) And &H1, BatteryStatusEnum)
            }
        End Function
    End Structure

    ' Method: GetZonesNames
    ' Purpose: Retrieves zone names from AirTouch console
    ' Flow:
    ' 1. Sends request for zone information
    ' 2. Processes response to extract zone names
    ' 3. Stores zone names in ZoneNames dictionary
    Public Sub GetZoneNames()
        ' Create request message (Extended type with command bytes FF 13)
        Dim requestData() As Byte = CreateMessage(MessageType.Extended, {&HFF, &H13})
        If Not AirTouch5Console.Connected Then Throw New System.Exception("Console not connected. Cannot retrieve zones.")
        Try
            ' Send request and get response
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(
                AirTouch5Console.IP, AirTouchPort, requestData)

            ' Parse response messages
            Dim msg As Message = Message.Parse(response)
            ZoneNames.Clear() ' Clear existing zone names

            Dim index = 2 ' Start of zone data in message

            ' Process each zone entry in the message
            While index < msg.data.Length - 1
                Dim ZoneNumber = msg.data(index)
                Dim NameLength = msg.data(index + 1)

                ' Extract zone name (UTF8 encoded)
                Dim ZoneName = Encoding.UTF8.GetString(msg.data, index + 2, NameLength)
                Debug.WriteLine($"Zone: {ZoneNumber} {ZoneName}")

                ' Add to dictionary
                ZoneNames.Add(ZoneNumber, ZoneName)

                ' Move to next zone entry
                index += NameLength + 2
            End While

        Catch ex As TimeoutException
            Debug.WriteLine("Request timed out")
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    ' Method: GetZoneStatus
    ' Purpose: Retrieves current status for all zones
    ' Flow:
    ' 1. Sends control request for zone status
    ' 2. Processes response to extract status for each zone
    ' 3. Stores status in ZoneStatuses dictionary
    Public Sub GetZoneStatus()
        ' Create control request message (command byte 21)
        Dim requestData() As Byte = CreateMessage(MessageType.Control, {CommandMessages.ZoneStatus, 0, 0, 0, 0, 0, 0, 0})
        Dim zoneStatus As ZoneStatusMessage

        Try
            ' Send request and get response
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(
                AirTouch5Console.IP, AirTouchPort, requestData)

            ' Parse response messages
            Dim msg As Message = Message.Parse(response)

            ZoneStatuses.Clear()

            Dim index = 8 ' Start of zone status data in message

            ' Process each zone status entry
            While index < msg.data.Length - 1

                ' Parse zone status
                zoneStatus = ZoneStatusMessage.Parse(msg.data, index)

                ' Store in dictionary
                ZoneStatuses(zoneStatus.ZoneNumber) = zoneStatus

                ' Output debug information
                With zoneStatus
                    Debug.WriteLine(
                        $"Zone: {ZoneNames(.ZoneNumber)} " &
                        $"Power State: { .ZoneState} " &
                        $"Control Method: { .ControlMethod} " &
                        $"Damper Open: { .DamperOpen}% " &
                        $"Set Point: { .SetPoint} " &
                        $"Temperature: { .Temperature} " &
                        $"Spill: { .Spill} " &
                        $"Sensor: { .Sensor} " &
                        $"Battery: { .Battery}")
                End With
                ' Move to next zone status entry
                index += 8
            End While
        Catch ex As TimeoutException
            Debug.WriteLine("Request timed out")
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
        Form1.TextBox1.AppendText("Zone Statuses retrieved successfully." & vbCrLf)
    End Sub
End Module