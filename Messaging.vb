Imports AirTouch5.Form1

' ==============================================================
' Module: Messaging
' Purpose: Handles all message formatting, parsing and validation
'          for communication with AirTouch5 console
' Features:
' - Defines message types and commands
' - Provides message creation and parsing functions
' - Implements custom CRC calculation with special rules
' ==============================================================
Module Messaging

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
        Public h1 As Byte           ' Header byte 1
        Public h2 As Byte           ' Header byte 2
        Public h3 As Byte           ' Header byte 3
        Public h4 As Byte           ' Header byte 4
        Public address As UShort    ' Device address
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
            ' Combine two bytes for data length (big-endian)
            Dim datalength = (data(8) << 8) Or data(9)

            ' Initialize and copy data payload
            Dim dat = New Byte(datalength - 1) {}
            Array.Copy(data, 10, dat, 0, datalength)

            Return New Message With {
                .h1 = data(0),
                .h2 = data(1),
                .h3 = data(2),
                .h4 = data(3),
                .address = (data(4) << 8) Or data(5),  ' Combine address bytes
                .id = data(6),
                .messageType = CType(data(7), MessageType),
                .dataLength = datalength,
                .data = dat,
                .checksum = (data(data.Length - 2) << 8) Or data(data.Length - 1)
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
        ms.WriteByte(checksum(0))  ' CRC low byte
        ms.WriteByte(checksum(1))  ' CRC high byte

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
    '   2-byte array containing CRC (big-endian)
    Public Function CalculateModbusCRC16WithSpecialRules(ByVal data As Byte()) As Byte()
        ' Validate input
        If data Is Nothing OrElse data.Length < 4 Then
            Return New Byte() {&HFF, &HFF} ' Invalid CRC
        End If

        Dim crc As UShort = &HFFFF
        Dim i As Integer = 4 ' Skip first 4 bytes (header)

        While i < data.Length
            ' Check for special sequence: 0x55 0x55 0x55 0x00
            If i + 3 < data.Length AndAlso
               data(i) = &H55 AndAlso
               data(i + 1) = &H55 AndAlso
               data(i + 2) = &H55 AndAlso
               data(i + 3) = &H0 Then
                ' Process first three 0x55 bytes
                For j As Integer = 0 To 2
                    ProcessByteForCRC(data(i + j), crc)
                Next
                i += 4 ' Skip all four bytes
            Else
                ProcessByteForCRC(data(i), crc)
                i += 1
            End If
        End While

        ' Return CRC in big-endian format
        Return New Byte() {CByte((crc >> 8) And &HFF), CByte(crc And &HFF)}
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

    ' Method: ParseMessages
    ' Purpose: Extracts complete messages from raw byte stream
    ' Parameters:
    '   data - Byte array containing one or more messages
    ' Returns:
    '   List of byte arrays, each containing one complete message
    Function ParseMessages(data As Byte()) As List(Of Byte())
        Dim messages As New List(Of Byte())
        Dim index As Integer = 0

        While index < data.Length - 3 ' Need at least 4 bytes for header
            ' Look for message start pattern: 0x55 0x55 0x55 0xAA
            If data(index) = &H55 AndAlso
               data(index + 1) = &H55 AndAlso
               data(index + 2) = &H55 AndAlso
               data(index + 3) = &HAA Then

                ' Verify we have enough bytes for complete header
                If index + 9 >= data.Length Then
                    Exit While ' Not enough data
                End If

                ' Extract message length (bytes 8-9, big-endian)
                Dim length As Integer = (data(index + 8) << 8) Or data(index + 9)
                Dim messageLength As Integer = length + 10 ' Header + payload

                ' Verify complete message is available
                If index + messageLength <= data.Length Then
                    ' Extract complete message
                    Dim message As Byte() = New Byte(messageLength - 1) {}
                    Array.Copy(data, index, message, 0, messageLength)
                    messages.Add(message)

                    ' Move to next potential message
                    index += messageLength
                    Continue While
                Else
                    Exit While ' Incomplete message
                End If
            End If

            ' No header found, move to next byte
            index += 1
        End While

        Return messages
    End Function
End Module