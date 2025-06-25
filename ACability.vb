Imports System.Text

' ==============================================================
' Module: ACability
' Purpose: Handles Air Conditioner capability information for AirTouch5 system
' Features:
' - Defines structure for AC capability information
' - Parses raw AC capability data from AirTouch console
' - Provides methods to retrieve and display AC capabilities
' ==============================================================
Friend Module ACability

    ' Enum: SupportedEnum
    ' Purpose: Simple boolean-like enum for supported features
    Enum SupportedEnum
        No = 0     ' Feature not supported
        Yes = 1    ' Feature supported
    End Enum

    ' Structure: ACability
    ' Purpose: Contains all capability information for an AC unit
    Public Structure ACability
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
        Public Shared Function Parse(data As Byte(), index As Integer) As ACability
            ' Extract AC name (16 bytes null-terminated)
            Dim name(15) As Byte
            Array.Copy(data, index + 2, name, 0, 16)
            Dim ACname As String = GetNullTerminatedString(name, 16)

            ' Parse supported features bitmask (bits 20-21)
            Dim twoBytes(1) As Byte
            Array.Copy(data, index + 20, twoBytes, 0, 2)    ' copy 2 bytes of masks
            Array.Reverse(twoBytes)        ' reverse them to match big-endian format
            Dim bitArray As New BitArray(twoBytes)          ' create a bitarray with them

            ' Create and return populated structure
            Return New ACability With {
                .ACnumber = data(index),
                .ACName = ACname,
                .StartZone = data(index + 18),
                .ZoneCount = data(index + 19),
                .CoolMode = If(bitArray(12), SupportedEnum.Yes, SupportedEnum.No),
                .FanMode = If(bitArray(11), SupportedEnum.Yes, SupportedEnum.No),
                .DryMode = If(bitArray(10), SupportedEnum.Yes, SupportedEnum.No),
                .HeatMode = If(bitArray(9), SupportedEnum.Yes, SupportedEnum.No),
                .AutoMode = If(bitArray(8), SupportedEnum.Yes, SupportedEnum.No),
                .FanSpeedIA = If(bitArray(7), SupportedEnum.Yes, SupportedEnum.No),
                .FanSpeedTurbo = If(bitArray(6), SupportedEnum.Yes, SupportedEnum.No),
                .FanSpeedPowerful = If(bitArray(5), SupportedEnum.Yes, SupportedEnum.No),
                .FanSpeedHigh = If(bitArray(4), SupportedEnum.Yes, SupportedEnum.No),
                .FanSpeedMedium = If(bitArray(3), SupportedEnum.Yes, SupportedEnum.No),
                .FanSpeedLow = If(bitArray(2), SupportedEnum.Yes, SupportedEnum.No),
                .FanSpeedQuiet = If(bitArray(1), SupportedEnum.Yes, SupportedEnum.No),
                .FanspeedAuto = If(bitArray(0), SupportedEnum.Yes, SupportedEnum.No),
                .MinCoolSetPoint = data(index + 22),
                .MaxCoolSetPoint = data(index + 23),
                .MinHeatSetPoint = data(index + 24),
                .MaxHeatSetPoint = data(index + 25)
            }
        End Function
    End Structure
    Public acData As ACability      ' Parsed AC capability data

    ' Method: GetACability
    ' Purpose: Retrieves AC capability information from AirTouch console
    ' Flow:
    ' 1. Sends extended request message (FF 11)
    ' 2. Receives and parses response
    ' 3. Outputs capability information to debug
    Public Sub GetACability()
        ' Create extended message requesting AC capabilities
        Dim requestData() As Byte = CreateMessage(MessageType.Extended, {&HFF, &H11})

        If Not AirTouch5Console.Connected Then Throw New System.Exception("Console not connected. Cannot retrieve AC ability.")
        Try
            ' Send request and receive response
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(
                AirTouch5Console.IP, AirTouchPort, requestData)

            ' Parse response messages
            Dim msg As Message = Message.Parse(response)

            ' Parse raw data into structured format
            acData = ACability.Parse(msg.data, 2)

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
        Catch ex As TimeoutException
            Debug.WriteLine("Request timed out")
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
        Form1.TextBox1.AppendText($"Retrieved AC ability information.{vbCrLf}")
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