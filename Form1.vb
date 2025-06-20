Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Text

Public Class Form1
    Dim udpPort As Integer = 49005
    Dim message As String = "::REQUEST-POLYAIRE-AIRTOUCH-DEVICE-INFO:;"
    Dim timeoutMs As Integer = 5000
    Dim udpClient As UdpClient ' Make client global to properly dispose
    Dim responseMessage As String
    Structure AirTouchConsole
        ' Capture details of discovered AirTouch console
        Dim IP As String
        Dim ConsoleID As String
        Dim AirTouch5 As String
        Dim AirTouchID As String
        Dim DeviceName As String
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
                udpClient = New UdpClient()
                udpClient.EnableBroadcast = True
                udpClient.Client.SetSocketOption(SocketOptionLevel.Socket,
                                              SocketOptionName.ReuseAddress,
                                              True)
                udpClient.Client.Bind(New IPEndPoint(IPAddress.Any, udpPort))
            End If

            Dim bytes As Byte() = Encoding.ASCII.GetBytes(message)
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
                If Not responseMessage = message Then
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
            udpClient = New UdpClient()
            udpClient.EnableBroadcast = True
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


    Sub SendMessage(address As Integer, id As Byte, typ As MessageType, data As Byte())
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
            AddByteToMessage(ms, &H90)
        Else
            AddByteToMessage(ms, &H90)
            AddByteToMessage(ms, &HB0)
        End If
        ' id
        addByteToMessage(ms, id)
        ' Message type
        addByteToMessage(ms, typ)
        ' data length
        addByteToMessage(ms, data.Length >> 8)
        addByteToMessage(ms, data.Length And &HFF)
        ' data
        For i = 0 To data.Length - 1
            addByteToMessage(ms, data(i))
        Next
        ' check bytes
        Dim checksum = CalculateModbusCRC16WithSpecialRules(ms.ToArray)
        addByteToMessage(ms, checksum(0))
        addByteToMessage(ms, checksum(1))

        Dim byteArray As Byte() = ms.ToArray

    End Sub
    Sub addByteToMessage(ms As IO.MemoryStream, b As Byte)
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
        Dim i As Integer = 4 ' Start at index 4 to skip first 4 bytes

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

        ' Return CRC in little-endian format (LSB first)
        Return New Byte() {CByte(crc And &HFF), CByte((crc >> 8) And &HFF)}
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
End Class