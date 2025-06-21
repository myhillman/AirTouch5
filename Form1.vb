Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1
    Dim udpPort As Integer = 49005
    Dim broadcastMessage As String = "::REQUEST-POLYAIRE-AIRTOUCH-DEVICE-INFO:;"
    Dim AirTouchPort = 9005 ' Port for AirTouch console HTTP requests
    Dim timeoutMs As Integer = 5000
    Dim udpClient As UdpClient ' Make client global to properly dispose
    Private responseMessage As String
    Dim ZoneNames As New Dictionary(Of Integer, String)
    Structure AirTouchConsole
        ' Capture details of discovered AirTouch console
        Dim IP As String    ' IP address of the console
        Dim ConsoleID As String ' Unique identifier for the console
        Dim AirTouch5 As String
        Dim AirTouchID As String ' Unique identifier for the AirTouch device
        Dim DeviceName As String ' Name of the device
    End Structure
    Dim AirTouch5Console As AirTouchConsole
    ' Enum for message types
    Enum MessageType
        ' Two message types
        Control = &HC0
        Extended = &H1F
    End Enum

    Enum CommandMessages
        ' Define command messages here
        ZoneControl = &H20
        ZoneStatus = &H21
        ACcontrol = &H22
        ACStatus = &H23
    End Enum
    Enum ExtendedMessages
        ' Define extended messages here
        ACAbility = &HFF11
        ACError = &HFF10
        ZoneName = &HFF13
        ConsoleVersion = &HFF30
    End Enum
    Public Structure Message
        ' This structure represents the header of the message
        Public h1 As Byte
        Public h2 As Byte
        Public h3 As Byte
        Public h4 As Byte
        Public address As UShort
        Public id As Byte
        Public messageType As MessageType
        Public dataLength As UShort
        Public data As Byte() ' Data field to hold variable-length data
        Public checksum As UShort

        Public Shared Function Parse(data As Byte()) As Message
            ' This function parses the extended header from the byte array
            Dim datalength = (data(8) << 8) Or data(9) ' Combine two bytes for data length
            Dim dat = New Byte(datalength - 1) {} ' Initialize data array with correct length
            Array.Copy(data, 10, dat, 0, datalength) ' Copy data into the array
            Return New Message With {
            .h1 = data(0),
            .h2 = data(1),
            .h3 = data(2),
            .h4 = data(3),
            .address = (data(4) << 8) Or data(5), ' Combine two bytes for address
            .id = data(6),
            .messageType = CType(data(7), MessageType), ' Convert byte to MessageType enum
            .dataLength = datalength,
            .data = dat,
            .checksum = data(data.Length - 2) << 8 Or data(data.Length - 1) ' Last two bytes are checksum
        }
        End Function
    End Structure

    Enum SupportedEnum
        No = 0
        Yes = 1
    End Enum

    Public Structure ACability
        Dim ACnumber As Byte
        Dim ACName As String
        Dim StartZone As Byte
        Dim ZoneCount As Byte
        Dim CoolMode As SupportedEnum
        Dim FanMode As SupportedEnum
        Dim DryMode As SupportedEnum
        Dim HeatMode As SupportedEnum
        Dim AutoMode As SupportedEnum
        Dim FanSpeedIA As SupportedEnum
        Dim FanSpeedTurbo As SupportedEnum
        Dim FanSpeedPowerful As SupportedEnum
        Dim FanSpeedHigh As SupportedEnum
        Dim FanSpeedMedium As SupportedEnum
        Dim FanSpeedLow As SupportedEnum
        Dim FanSpeedQuiet As SupportedEnum
        Dim FanspeedAuto As SupportedEnum
        Dim MinCoolSetPoint As Byte
        Dim MaxCoolSetPoint As Byte
        Dim MinHeatSetPoint As Byte
        Dim MaxHeatSetPoint As Byte
        Public Shared Function Parse(data As Byte()) As ACability
            ' This function parses the extended header from the byte array
            Dim name(15) As Byte
            Array.Copy(data, 2, name, 0, 16)    ' extract 16 byte name
            Dim ACname As String = GetNullTerminatedString(Name, 16)
            Dim SupportedMasks = CInt(data(20)) << 8 Or data(21) ' 13 bits for supported modes
            Dim SupportedList As New List(Of SupportedEnum)
            For i = 1 To 13
                Dim support = CType(SupportedMasks And 1, SupportedEnum) ' Check if the bit is set
                SupportedList.Add(support) ' Add to the list
                SupportedMasks = SupportedMasks >> 1 ' Shift right to check next bit
            Next
            SupportedList.Reverse()
            Return New ACability With {
                .ACnumber = data(0),
                .ACName = ACname,
                .StartZone = data(18),
                .ZoneCount = data(19),
                .CoolMode = SupportedList(0),
                .FanMode = SupportedList(1),
                .DryMode = SupportedList(2),
                .HeatMode = SupportedList(3),
                .AutoMode = SupportedList(4),
                .FanSpeedIA = SupportedList(5),
                .FanSpeedTurbo = SupportedList(6),
                .FanSpeedPowerful = SupportedList(7),
                .FanSpeedHigh = SupportedList(8),
                .FanSpeedMedium = SupportedList(9),
                .FanSpeedLow = SupportedList(10),
                .FanSpeedQuiet = SupportedList(11),
                .FanspeedAuto = SupportedList(12),
                .MinCoolSetPoint = data(22), ' Min cool set point
                .MaxCoolSetPoint = data(23), ' Max cool set point
                .MinHeatSetPoint = data(24),
                .MaxHeatSetPoint = data(25)
        }
        End Function
    End Structure
    Enum ZonePowerStateEnum
        ' Define zone power states
        Off = 0
        [On] = 1
        Turbo = 3
    End Enum
    Enum ControlMethodEnum
        ' Define zone power states
        Percentage = 0
        Temperature = 1
    End Enum
    Enum SensorStatusEnum
        ' Define sensor status
        Absent = 0
        Present = 1
    End Enum
    Enum SpillStatusEnum
        ' Define spill status
        Inactive = 0
        Active = 1
    End Enum
    Enum BatteryStatusEnum
        ' Define battery status
        Normal = 0
        Low = 1
    End Enum
    Public Structure ZoneStatusMessage
        ' This structure represents the Zone Status message
        Public ZonePowerState As ZonePowerStateEnum
        Public ZoneNumber As Byte
        Public ControlMethod As ControlMethodEnum
        Public OpenPercentage As Byte
        Public SetPoint As Short
        Public Sensor As SensorStatusEnum
        Public Temperature As Short
        Public Spill As SpillStatusEnum
        Public Battery As BatteryStatusEnum

        Public Shared Function Parse(data As Byte()) As ZoneStatusMessage
            ' This function parses the extended header from the byte array
            Return New ZoneStatusMessage With {
                        .ZonePowerState = CType((data(0) >> 6) And &H3, ZonePowerStateEnum), ' Convert byte to ZonePowerStateEnum
                        .ZoneNumber = data(0) And &HF,
                        .ControlMethod = CType((data(1) >> 7) And &H1, ControlMethodEnum), ' Extract control method from bits 6-7
                        .OpenPercentage = data(1) And &H7F,
                        .SetPoint = (data(2) + 100) / 10, ' Set point temperature
                        .Sensor = CType(data(3) >> 7 And &H1, SensorStatusEnum), ' Sensor status from bit 0
                        .Temperature = (((CInt(data(4)) << 8) Or data(5)) - 500) / 10,
                        .Spill = CType((data(6) >> 1) And 1, SpillStatusEnum), ' Spill status from bit 2
                        .Battery = CType((data(6) And &H1), BatteryStatusEnum) ' Low battery status from bit 1
        }
        End Function
    End Structure
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If udpClient IsNot Nothing Then
            udpClient.Close()
            udpClient.Dispose()
        End If
    End Sub

    Sub SendUdpBroadcast()
        Try
            ' Create and configure client once
            If udpClient Is Nothing Then
                udpClient = New UdpClient With {
                    .EnableBroadcast = True
                }
                udpClient.Client.SetSocketOption(SocketOptionLevel.Socket,
                                              SocketOptionName.ReuseAddress,
                                              True)
                udpClient.Client.Bind(New IPEndPoint(IPAddress.Any, udpPort))
            End If

            Dim bytes As Byte() = Encoding.ASCII.GetBytes(broadcastMessage)
            Dim broadcastIp As New IPEndPoint(IPAddress.Broadcast, udpPort)
            udpClient.Send(bytes, bytes.Length, broadcastIp)
            Debug.WriteLine($"Broadcast sent on port {udpPort}")
        Catch ex As Exception
            Debug.WriteLine($"Send error: {ex.Message}")
            ' Clean up if error occurs
            If udpClient IsNot Nothing Then
                udpClient.Close()
                udpClient = Nothing
            End If
        End Try
    End Sub

    Async Function ReceiveUdpResponsesAsync() As Task(Of Boolean)
        Try
            udpClient.Client.ReceiveTimeout = timeoutMs

            While True
                Dim receiveResult = Await udpClient.ReceiveAsync()
                responseMessage = Encoding.ASCII.GetString(receiveResult.Buffer)

                Me.Invoke(Sub()
                              ' Update UI here
                              Debug.WriteLine($"Received: {responseMessage}")
                          End Sub)
                ' Check if the response contains the expected data    
                If Not responseMessage = broadcastMessage Then
                    Return True ' we got a response
                End If
            End While
        Catch ex As SocketException When ex.ErrorCode = 10060
            Debug.WriteLine("Receive timeout reached")
            Return False
        Catch ex As Exception
            Debug.WriteLine($"Receive error: {ex.Message}")
            Return False
        End Try
    End Function

    Private Async Sub FindControllerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindControllerToolStripMenuItem.Click
        Debug.WriteLine("Starting discovery...")

        ' Initialize client if needed
        If udpClient Is Nothing Then
            udpClient = New UdpClient With {
                .EnableBroadcast = True
            }
            udpClient.Client.SetSocketOption(SocketOptionLevel.Socket,
                                           SocketOptionName.ReuseAddress,
                                           True)
            udpClient.Client.Bind(New IPEndPoint(IPAddress.Any, udpPort))
        End If

        ' Start receiving
        Dim receiveTask = ReceiveUdpResponsesAsync()

        ' Small delay to ensure receiver is ready
        Await Task.Delay(100)

        ' Send broadcast
        SendUdpBroadcast()

        ' Wait for completion
        Await receiveTask
        Debug.WriteLine("Discovery completed", responseMessage)
        ' Process the response
        Dim decode = responseMessage.Split(",")     ' split into pieces
        With AirTouch5Console
            .IP = decode(0)
            .ConsoleID = decode(1)
            .AirTouch5 = decode(2)
            .AirTouchID = decode(3)
            .DeviceName = decode(4)
        End With
        Debug.WriteLine($"Found AirTouch Console: {AirTouch5Console.DeviceName} at IP {AirTouch5Console.IP}")
    End Sub

    Public Function CreateMessage(typ As MessageType, data As Byte()) As Byte()
        ' This function sends a message to the AirTouch console
        Dim ms As New IO.MemoryStream

        ' Add header
        ms.WriteByte(&H55)
        ms.WriteByte(&H55)
        ms.WriteByte(&H55)
        ms.WriteByte(&HAA)
        ' Address
        If typ = MessageType.Control Then
            AddByteToMessage(ms, &H80)
        Else
            AddByteToMessage(ms, &H90)
        End If
        AddByteToMessage(ms, &HB0)
        ' id
        AddByteToMessage(ms, 1)
        ' Message type
        AddByteToMessage(ms, typ)
        ' data length
        AddByteToMessage(ms, data.Length >> 8)
        AddByteToMessage(ms, data.Length And &HFF)
        ' data
        For i = 0 To data.Length - 1
            AddByteToMessage(ms, data(i))
        Next
        ' check bytes
        Dim checksum = CalculateModbusCRC16WithSpecialRules(ms.ToArray)
        ms.WriteByte(checksum(0))
        ms.WriteByte(checksum(1))

        Dim byteArray As Byte() = ms.ToArray
        Return byteArray ' Return the byte array for sending

    End Function

    Public Shared Sub AddByteToMessage(ms As IO.MemoryStream, b As Byte)
        ' This function adds a byte to the message stream and adds 0x00 byte if there are 3 consecutive 0x55
        ms.WriteByte(b)
        ' Check if stream has at least 3 bytes
        If ms.Length >= 3 Then
            ' Move to the last 3 bytes
            ms.Seek(-3, IO.SeekOrigin.End)

            ' Read the last 3 bytes
            Dim buffer(2) As Byte ' Array to hold 3 bytes
            Dim bytesRead As Integer = ms.Read(buffer, 0, 3)

            ' Check if we read 3 bytes and they are all 0x55
            If bytesRead = 3 AndAlso
               buffer(0) = &H55 AndAlso
               buffer(1) = &H55 AndAlso
               buffer(2) = &H55 Then
                ms.WriteByte(0)
            End If
        End If
    End Sub
    Public Function CalculateModbusCRC16WithSpecialRules(ByVal data As Byte()) As Byte()
        ' This function calculates the Modbus CRC16 with special rules for the AirTouch console
        If data Is Nothing OrElse data.Length < 4 Then
            Return New Byte() {&HFF, &HFF} ' Return invalid CRC for invalid input
        End If

        Dim crc As UShort = &HFFFF
        Dim i As Integer = 4 ' Start at index 4 to skip first 4 bytes (header)

        While i < data.Length
            ' Check for 0x55 0x55 0x55 0x00 sequence
            If i + 3 < data.Length AndAlso
               data(i) = &H55 AndAlso
               data(i + 1) = &H55 AndAlso
               data(i + 2) = &H55 AndAlso
               data(i + 3) = &H0 Then
                ' Skip the 0x00 byte in this special sequence
                For j As Integer = 0 To 2
                    ProcessByteForCRC(data(i + j), crc)
                Next
                i += 4 ' Skip all 4 bytes (we've processed 3, skip the 4th)
            Else
                ProcessByteForCRC(data(i), crc)
                i += 1
            End If
        End While

        ' Return CRC in big-endian format (MSB first)
        Return New Byte() {CByte((crc >> 8) And &HFF), CByte(crc And &HFF)}
    End Function

    Private Sub ProcessByteForCRC(ByVal b As Byte, ByRef crc As UShort)
        ' This function processes a byte for CRC calculation
        crc = crc Xor CUShort(b)

        For bit As Integer = 0 To 7
            If (crc And &H1) <> 0 Then
                crc = CUShort(crc >> 1) Xor &HA001
            Else
                crc = CUShort(crc >> 1)
            End If
        Next
    End Sub

    Private Sub GetVersionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetVersionToolStripMenuItem.Click
        Dim requestData() As Byte = CreateMessage(MessageType.Extended, {&HFF, &H30})
        Try
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(AirTouch5Console.IP, AirTouchPort, requestData)
            Debug.WriteLine("Response: " & Encoding.UTF8.GetString(response))
            Dim messages = ParseMessages(response)
            For Each msg In messages
                Dim m = Message.Parse(msg)
                Dim version = Encoding.UTF8.GetString(m.data, 4, m.data(3))
                Debug.WriteLine("Console Version: " & version)
            Next
        Catch ex As TimeoutException
            Debug.WriteLine("Request timed out")
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    Shared Function ParseMessages(data As Byte()) As List(Of Byte())
        Dim messages As New List(Of Byte())
        Dim index As Integer = 0

        While index < data.Length - 3 ' Ensure we have at least 4 bytes to check
            ' Check for 0x55 0x55 0x55 pattern
            If data(index) = &H55 AndAlso
           data(index + 1) = &H55 AndAlso
           data(index + 2) = &H55 AndAlso
           data(index + 3) = &HAA Then

                ' Make sure we have enough bytes for the header (minimum 4 bytes)
                If index + 3 >= data.Length Then
                    Exit While
                End If

                ' We need at least 10 bytes for a complete header (0x55,0x55,0x55,0xAA + 6 more)
                If index + 9 >= data.Length Then
                    Exit While
                End If

                ' Read the message length (assuming bytes 8-9 contain length)
                Dim length As Integer = (data(index + 8) << 8) Or data(index + 9)
                Dim messageLength As Integer = length + 10 ' Header + length bytes

                ' Check if we have enough data for the complete message
                If index + messageLength <= data.Length Then
                    Dim message As Byte() = New Byte(messageLength - 1) {}
                    Array.Copy(data, index, message, 0, messageLength)
                    messages.Add(message)
                    index += messageLength ' Move to next potential message
                    Continue While ' Skip the index increment at the end
                Else
                    Exit While ' Not enough data for complete message
                End If
            End If

            ' Default increment if no pattern matched
            index += 1
        End While

        Return messages
    End Function

    Private Sub GetZonesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetZonesToolStripMenuItem.Click
        Dim requestData() As Byte = CreateMessage(MessageType.Extended, {&HFF, &H13})
        Try
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(AirTouch5Console.IP, AirTouchPort, requestData)
            'Debug.WriteLine("Response: " & Encoding.UTF8.GetString(response))
            Dim messages = ParseMessages(response)
            For Each msg In messages
                Dim m = Message.Parse(msg)
                Dim index = 2        ' start of zone data
                ZoneNames.Clear()
                While index < m.data.Length - 1
                    Dim ZoneNumber = m.data(index)
                    Dim NameLength = m.data(index + 1)
                    Dim ZoneName = Encoding.UTF8.GetString(m.data, index + 2, NameLength)
                    Debug.WriteLine($"Zone: {ZoneNumber} {ZoneName}")
                    ZoneNames.Add(ZoneNumber, ZoneName)
                    index += NameLength + 2 ' next zone
                End While
            Next
        Catch ex As TimeoutException
            Debug.WriteLine("Request timed out")
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ZoneStatusToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoneStatusToolStripMenuItem.Click
        Dim requestData() As Byte = CreateMessage(MessageType.Control, {&H21, 0, 0, 0, 0, 0, 0, 0})
        Dim zoneData(7) As Byte     ' 8 bytes of raw zone data
        Dim zoneStatus As ZoneStatusMessage     ' parse zone status
        Try
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(AirTouch5Console.IP, AirTouchPort, requestData)
            'Debug.WriteLine("Response: " & Encoding.UTF8.GetString(response))
            Dim messages = ParseMessages(response)
            For Each msg In messages
                Dim m = Message.Parse(msg)
                Dim index = 8        ' start of zone data
                While index < m.data.Length - 1
                    Array.Copy(m.data, index, zoneData, 0, 8)    ' copy 8 bytes of zone data
                    zoneStatus = ZoneStatusMessage.Parse(zoneData) ' parse zone status
                    Debug.WriteLine($"Zone: {ZoneNames(zoneStatus.ZoneNumber)} Power State: {zoneStatus.ZonePowerState} Control Method: {zoneStatus.ControlMethod} Open: {zoneStatus.OpenPercentage}% Set Point: {zoneStatus.SetPoint} Temperature: {zoneStatus.Temperature} Spill: {zoneStatus.Spill} Sensor: {zoneStatus.Sensor} Battery: {zoneStatus.Battery}")
                    index += 8 ' next zone
                End While
            Next
        Catch ex As TimeoutException
            Debug.WriteLine("Request timed out")
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ACAbilityToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ACAbilityToolStripMenuItem.Click
        Dim requestData() As Byte = CreateMessage(MessageType.Extended, {&HFF, &H11})
        Dim acData As ACability     '  raw AC data
        Dim acRawData(28) As Byte     ' 29 bytes of raw AC data

        Try
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(AirTouch5Console.IP, AirTouchPort, requestData)
            Dim messages = ParseMessages(response)
            For Each msg In messages
                Dim m = Message.Parse(msg)
                Array.Copy(m.data, 2, acRawData, 0, m.data.Length - 2) ' copy raw AC data
                acData = ACability.Parse(acRawData) ' parse AC data
                Debug.WriteLine($"AC #: {acData.ACnumber} AC name: {acData.ACName} Start Zone: {acData.StartZone} Zone count: {acData.ZoneCount} Min Cool: {acData.MinCoolSetPoint} Max Cool: {acData.MaxCoolSetPoint} Min Heat: {acData.MinHeatSetPoint} Max Heat: {acData.MaxHeatSetPoint}")
                Debug.WriteLine($"Modes: Cool: {acData.CoolMode} Fan: {acData.FanMode} Dry: {acData.DryMode} Heat: {acData.HeatMode} Auto: {acData.AutoMode}")
                Debug.WriteLine($"Fan Speeds: IA: {acData.FanSpeedIA} Turbo: {acData.FanSpeedTurbo} Powerful: {acData.FanSpeedPowerful} High: {acData.FanSpeedHigh} Medium: {acData.FanSpeedMedium} Low: {acData.FanSpeedLow} Quiet: {acData.FanSpeedQuiet} Auto: {acData.FanspeedAuto}")
            Next
        Catch ex As TimeoutException
            Debug.WriteLine("Request timed out")
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
    End Sub
    Public Shared Function GetNullTerminatedString(bytes As Byte(), maxLength As Integer) As String
        ' Validate input
        If bytes Is Nothing OrElse bytes.Length = 0 Then Return String.Empty

        ' Ensure maxLength doesn't exceed array bounds
        Dim searchLength As Integer = Math.Min(bytes.Length, maxLength)

        ' Find null terminator within the search range
        Dim nullIndex As Integer = -1
        For i As Integer = 0 To searchLength - 1
            If bytes(i) = 0 Then
                nullIndex = i
                Exit For
            End If
        Next

        ' Determine the actual length to extract
        Dim stringLength As Integer = If(nullIndex >= 0, nullIndex, searchLength)

        ' Handle empty string case
        If stringLength = 0 Then Return String.Empty

        ' Extract the string
        Return Encoding.ASCII.GetString(bytes, 0, stringLength)
    End Function
End Class