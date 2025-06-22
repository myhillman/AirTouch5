Imports System.Text

' ==============================================================
' Module: ACability
' Purpose: Handles Air Conditioner capability information for AirTouch5 system
' Features:
' - Defines structure for AC capability information
' - Parses raw AC capability data from AirTouch console
' - Provides methods to retrieve and display AC capabilities
' ==============================================================
Module ACability

    ' Enum: SupportedEnum
    ' Purpose: Simple boolean-like enum for supported features
    Enum SupportedEnum
        No = 0     ' Feature not supported
        Yes = 1    ' Feature supported
    End Enum

    ' Structure: ACability
    ' Purpose: Contains all capability information for an AC unit
    Private Structure ACability
        ' Basic identification
        Dim ACnumber As Byte         ' AC unit number
        Dim ACName As String         ' Name of AC unit (max 16 chars)

        ' Zone information
        Dim StartZone As Byte        ' First zone controlled by this AC
        Dim ZoneCount As Byte        ' Number of zones controlled

        ' Operation modes
        Dim CoolMode As SupportedEnum    ' Cooling supported
        Dim FanMode As SupportedEnum     ' Fan only supported
        Dim DryMode As SupportedEnum     ' Dry mode supported
        Dim HeatMode As SupportedEnum    ' Heating supported
        Dim AutoMode As SupportedEnum    ' Auto mode supported

        ' Fan speed capabilities
        Dim FanSpeedIA As SupportedEnum      ' IA fan speed
        Dim FanSpeedTurbo As SupportedEnum   ' Turbo speed
        Dim FanSpeedPowerful As SupportedEnum ' Powerful speed
        Dim FanSpeedHigh As SupportedEnum    ' High speed
        Dim FanSpeedMedium As SupportedEnum  ' Medium speed
        Dim FanSpeedLow As SupportedEnum     ' Low speed
        Dim FanSpeedQuiet As SupportedEnum   ' Quiet speed
        Dim FanspeedAuto As SupportedEnum    ' Auto speed

        ' Temperature set points
        Dim MinCoolSetPoint As Byte   ' Minimum cooling temperature
        Dim MaxCoolSetPoint As Byte   ' Maximum cooling temperature
        Dim MinHeatSetPoint As Byte   ' Minimum heating temperature
        Dim MaxHeatSetPoint As Byte   ' Maximum heating temperature

        ' Method: Parse
        ' Purpose: Converts raw byte data into ACability structure
        ' Parameters:
        '   data - Byte array containing AC capability information
        ' Returns:
        '   ACability structure populated with parsed data
        Public Shared Function Parse(data As Byte()) As ACability
            ' Extract AC name (16 bytes null-terminated)
            Dim name(15) As Byte
            Array.Copy(data, 2, name, 0, 16)
            Dim ACname As String = GetNullTerminatedString(name, 16)

            ' Parse supported features bitmask (bits 20-21)
            Dim SupportedMasks = CInt(data(20)) << 8 Or data(21)

            ' Convert bitmask to list of SupportedEnum values
            Dim SupportedList As New List(Of SupportedEnum)
            For i = 1 To 13
                ' Check each bit (LSB first)
                Dim support = CType(SupportedMasks And 1, SupportedEnum)
                SupportedList.Add(support)
                SupportedMasks >>= 1 ' Move to next bit
            Next
            SupportedList.Reverse() ' Reverse to match original bit order

            ' Create and return populated structure
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
                .MinCoolSetPoint = data(22),
                .MaxCoolSetPoint = data(23),
                .MinHeatSetPoint = data(24),
                .MaxHeatSetPoint = data(25)
            }
        End Function
    End Structure

    ' Method: GetACability
    ' Purpose: Retrieves AC capability information from AirTouch console
    ' Flow:
    ' 1. Sends extended request message (FF 11)
    ' 2. Receives and parses response
    ' 3. Outputs capability information to debug
    Public Sub GetACability()
        ' Create extended message requesting AC capabilities
        Dim requestData() As Byte = CreateMessage(MessageType.Extended, {&HFF, &H11})
        Dim acData As ACability      ' Parsed AC capability data
        Dim acRawData(28) As Byte    ' Buffer for raw AC data (29 bytes)
        If Not AirTouch5Console.Connected Then Throw New System.Exception("Console not connected. Cannot retrieve AC ability.")
        Try
            ' Send request and receive response
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(
                AirTouch5Console.IP, AirTouchPort, requestData)

            ' Parse response messages
            Dim messages = ParseMessages(response)

            For Each msg In messages
                ' Parse message and extract AC data
                Dim m = Message.Parse(msg)
                Array.Copy(m.data, 2, acRawData, 0, m.data.Length - 2)

                ' Parse raw data into structured format
                acData = ACability.Parse(acRawData)

                ' Output basic AC information
                Debug.WriteLine(
                    $"AC #: {acData.ACnumber} " &
                    $"AC name: {acData.ACName} " &
                    $"Start Zone: {acData.StartZone} " &
                    $"Zone count: {acData.ZoneCount} " &
                    $"Min Cool: {acData.MinCoolSetPoint} " &
                    $"Max Cool: {acData.MaxCoolSetPoint} " &
                    $"Min Heat: {acData.MinHeatSetPoint} " &
                    $"Max Heat: {acData.MaxHeatSetPoint}")

                ' Output supported modes
                Debug.WriteLine(
                    $"Modes: Cool: {acData.CoolMode} " &
                    $"Fan: {acData.FanMode} " &
                    $"Dry: {acData.DryMode} " &
                    $"Heat: {acData.HeatMode} " &
                    $"Auto: {acData.AutoMode}")

                ' Output supported fan speeds
                Debug.WriteLine(
                    $"Fan Speeds: IA: {acData.FanSpeedIA} " &
                    $"Turbo: {acData.FanSpeedTurbo} " &
                    $"Powerful: {acData.FanSpeedPowerful} " &
                    $"High: {acData.FanSpeedHigh} " &
                    $"Medium: {acData.FanSpeedMedium} " &
                    $"Low: {acData.FanSpeedLow} " &
                    $"Quiet: {acData.FanSpeedQuiet} " &
                    $"Auto: {acData.FanspeedAuto}")
            Next
        Catch ex As TimeoutException
            Debug.WriteLine("Request timed out")
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    ' Method: GetNullTerminatedString
    ' Purpose: Extracts null-terminated string from byte array
    ' Parameters:
    '   bytes - Source byte array
    '   maxLength - Maximum characters to read
    ' Returns:
    '   String containing characters up to null terminator or maxLength
    Private Function GetNullTerminatedString(bytes As Byte(), maxLength As Integer) As String
        ' Validate input
        If bytes Is Nothing OrElse bytes.Length = 0 Then Return String.Empty

        ' Ensure we don't exceed array bounds
        Dim searchLength As Integer = Math.Min(bytes.Length, maxLength)

        ' Find null terminator
        Dim nullIndex As Integer = -1
        For i As Integer = 0 To searchLength - 1
            If bytes(i) = 0 Then
                nullIndex = i
                Exit For
            End If
        Next

        ' Determine length to extract
        Dim stringLength As Integer = If(nullIndex >= 0, nullIndex, searchLength)

        ' Return empty string if no data
        If stringLength = 0 Then Return String.Empty

        ' Convert bytes to ASCII string
        Return Encoding.ASCII.GetString(bytes, 0, stringLength)
    End Function
End Module