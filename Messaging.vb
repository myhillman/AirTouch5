
''' <summary>
''' Module: Messaging
''' 
''' Handles all message formatting, parsing, and validation for communication 
''' with AirTouch5 console. Implements the complete message protocol including:
''' - Message structure definitions
''' - Header validation
''' - Special byte sequence handling
''' - Custom CRC calculation
''' </summary>
''' <remarks>
''' Protocol Features:
''' - Messages start with 0x55 0x55 0x55 0xAA header
''' - Special handling for 0x55 0x55 0x55 0x00 sequences
''' - Big-endian length and checksum fields
''' - Two message types (Control and Extended)
''' - Address-based routing
''' </remarks>
Friend Module Messaging

    ' Enum: MessageType
    ' Purpose: Defines the two main message types
    Enum MessageType
        Control = &HC0    ' Control messages (0xC0)
        Extended = &H1F   ' Extended messages (0x1F)
    End Enum

    ' Enum: CommandMessages
    ' Purpose: Defines specific command message types
    Enum CommandMessages
        ZoneControl = &H20   ' Zone control command
        ZoneStatus = &H21    ' Zone status request
        ACcontrol = &H22     ' AC control command
        ACStatus = &H23      ' AC status request
    End Enum

    ' Enum: ExtendedMessages
    ' Purpose: Defines specific extended message types
    Enum ExtendedMessages
        ACAbility = &HFF11      ' AC capability request
        ACError = &HFF10        ' AC error information
        ZoneName = &HFF13       ' Zone name request
        ConsoleVersion = &HFF30 ' Console version request
    End Enum

    ' Structure: Message
    ' Purpose: Represents the complete message structure
    Public Structure Message
        ' Header fields
        Public addrFrom As Byte    ' Device address
        Public addrTo As Byte
        Public id As Byte           ' Message ID
        Public messageType As MessageType ' Message type
        Public dataLength As UShort ' Length of data payload
        Public data As Byte()       ' Variable-length data payload
        Public checksum As UShort   ' Message checksum

        ' Method: Parse
        ' Purpose: Converts raw byte array into Message structure
        ' Parameters:
        '   data - Byte array containing complete message
        ' Returns:
        '   Message structure with parsed fields
        Public Shared Function Parse(data As Byte()) As Message
            ' First validate the 0x55 0x55 0x55 0xAA header sequence exists anywhere in the message
            Dim i As Integer
            Dim headerFound As Boolean = False
            Dim headerIndex As Integer = -1

            ' Search for header sequence (0x55 0x55 0x55 0xAA)
            For i = 0 To data.Length - 4
                If data(i) = &H55 AndAlso
           data(i + 1) = &H55 AndAlso
           data(i + 2) = &H55 AndAlso
           data(i + 3) = &HAA Then
                    headerFound = True
                    headerIndex = i
                    Exit For
                End If
            Next

            If Not headerFound Then
                Form1.TextBox1.AppendText("Error: Header sequence 0x55 0x55 0x55 0xAA not found" & vbCrLf)
                Return Nothing
            End If

            ' Verify we have enough data after the header (minimum 6 bytes: 2 reserved + 2 length + 2 checksum)
            If data.Length < headerIndex + 10 Then
                Form1.TextBox1.AppendText("Error: Insufficient data after header" & vbCrLf)
                Return Nothing
            End If

            Dim datalength As Integer = (data(headerIndex + 8) << 8) Or data(headerIndex + 9) ' Get packet length from offsets 8-9 (big-endian) relative to header
            ' Check if we have complete packet (header + reserved + length + data + checksum)
            If data.Length < headerIndex + 10 + datalength + 2 Then
                Form1.TextBox1.AppendText("Error: Packet length exceeds available data" & vbCrLf)
                Return Nothing
            End If

            Dim csmoffset = headerIndex + datalength + 8
            Dim CheckSum = (data(csmoffset) << 8) Or data(csmoffset + 1) ' Get checksum from offsets after data

            ' Now all we have to do is copy the data payload, but handling redundant zero's
            Dim payload As New IO.MemoryStream()    ' Build payload in memory stream and then convert to byte array
            Dim DataStart As Integer = headerIndex + 10 ' Start of data after header and reserved bytes
            i = DataStart
            While i < DataStart + datalength  ' Process data bytes
                ' Check for redundant zero pattern (0x55 0x55 0x55 0x00)
                If payload.Length >= 3 Then ' check for redundant byte
                    ' Check preceding 3 bytes are 0x55
                    For b As Long = -2 To 0 ' Check last 3 bytes before current position
                        If payload.Seek(b, IO.SeekOrigin.End) <> &H55 Then Exit For
                        If data(i) = 0 Then
                            ' Found redundant 0x00 after 3 0x55 bytes
                            i += 1 ' Skip this byte
                            Form1.TextBox1.AppendText("Removing redundant 0x00 from sequence at index " & i & vbCrLf)
                            Continue While
                        End If
                    Next
                End If
                payload.Position = payload.Length   ' move to end of stream
                payload.WriteByte(data(i))  ' output byte
                i += 1      ' process next byte
            End While

            ' Debug output
            Dim packet(data.Length - headerIndex) As Byte
            Array.Copy(data, headerIndex, packet, 0, data.Length - headerIndex)
            Form1.TextBox1.AppendText($"Parsed message: {BitConverter.ToString(packet.ToArray())}" & vbCrLf)

            Return New Message With {
                .addrFrom = data(headerIndex + 4),
                .addrTo = data(headerIndex + 5),
                .id = data(headerIndex + 6),
                .messageType = CType(data(headerIndex + 7), MessageType),
                .dataLength = payload.Length,
                .data = payload.ToArray,
                .checksum = CheckSum
            }
        End Function
    End Structure

    ' Method: CreateMessage
    ' Purpose: Creates properly formatted message for AirTouch console
    ' Parameters:
    '   typ - Message type (Control or Extended)
    '   data - Payload data to include
    ' Returns:
    '   Byte array containing complete formatted message
    Public Function CreateMessage(typ As MessageType, data As Byte()) As Byte()
        Dim ms As New IO.MemoryStream

        ' Add message header (0x55 0x55 0x55 0xAA)
        ms.WriteByte(&H55)
        ms.WriteByte(&H55)
        ms.WriteByte(&H55)
        ms.WriteByte(&HAA)

        ' Add address bytes (depends on message type)
        If typ = MessageType.Control Then
            AddByteToMessage(ms, &H80)  ' Control message address
        Else
            AddByteToMessage(ms, &H90)  ' Extended message address
        End If
        AddByteToMessage(ms, &HB0)     ' Fixed address byte

        ' Add message ID and type
        AddByteToMessage(ms, 1)         ' Default ID
        AddByteToMessage(ms, typ)       ' Message type

        ' Add data length (big-endian)
        AddByteToMessage(ms, data.Length >> 8)   ' High byte
        AddByteToMessage(ms, data.Length And &HFF) ' Low byte

        ' Add payload data
        For i = 0 To data.Length - 1
            AddByteToMessage(ms, data(i))
        Next

        ' Calculate and append CRC
        Dim checksum = CalculateModbusCRC16WithSpecialRules(ms.ToArray)

        ms.WriteByte((checksum >> 8) And &HFF)  ' CRC high byte
        ms.WriteByte(checksum And &HFF)  ' CRC low byte

        Return ms.ToArray()
    End Function

    ' Method: AddByteToMessage
    ' Purpose: Adds byte to message stream with special 0x55 sequence handling
    ' Parameters:
    '   ms - MemoryStream containing message
    '   b - Byte to add
    Private Sub AddByteToMessage(ms As IO.MemoryStream, b As Byte)
        ms.WriteByte(b)

        ' Check for three consecutive 0x55 bytes
        If ms.Length >= 3 Then
            ms.Seek(-3, IO.SeekOrigin.End)

            Dim buffer(2) As Byte
            Dim bytesRead As Integer = ms.Read(buffer, 0, 3)

            ' If three 0x55 found, insert 0x00 delimiter
            If bytesRead = 3 AndAlso
               buffer(0) = &H55 AndAlso
               buffer(1) = &H55 AndAlso
               buffer(2) = &H55 Then
                ms.WriteByte(0)
            End If
        End If
    End Sub

    ' Method: CalculateModbusCRC16WithSpecialRules
    ' Purpose: Calculates CRC with special handling for 0x55 sequences
    ' Parameters:
    '   data - Byte array to calculate CRC for
    ' Returns:
    '   2-byte CRC
    Public Function CalculateModbusCRC16WithSpecialRules(ByVal data As Byte()) As UShort
        ' Validate input
        If data Is Nothing Then
            Return &HFFFF ' Invalid CRC
        End If

        Dim crc As UShort = &HFFFF
        Dim i As Integer = 4        ' skip header bytes (0x55 0x55 0x55 0xAA)
        While i < data.Length
            ProcessByteForCRC(data(i), crc)
            i += 1
        End While

        ' Return CRC in big-endian format
        Return crc
    End Function

    ' Method: ProcessByteForCRC
    ' Purpose: Processes single byte in CRC calculation
    ' Parameters:
    '   b - Byte to process
    '   crc - Current CRC value (passed by reference)
    Private Sub ProcessByteForCRC(ByVal b As Byte, ByRef crc As UShort)
        crc = crc Xor CUShort(b)

        ' Process each bit
        For bit As Integer = 0 To 7
            If (crc And &H1) <> 0 Then
                crc = CUShort(crc >> 1) Xor &HA001 ' Polynomial
            Else
                crc = CUShort(crc >> 1)
            End If
        Next
    End Sub
End Module