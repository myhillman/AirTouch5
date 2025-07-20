Friend Module ACStatus
    ''' <summary>
    ''' Module: ACStatus
    ''' Purpose: Contains structures and methods to retrieve and parse AC status information from the AirTouch 5 console.
    ''' </summary>
    Enum ACPowerEnum
        Off = 0
        [On] = 1
        OffAlways = 2
        OnAlways = 3
        Sleep = 5
    End Enum
    Enum ModeEnum
        Auto = 0
        Heat = 1
        Dry = 2
        Fan = 3
        Cool = 4
        AutoHeat = 8
        AuroCool = 9
    End Enum
    Enum FanSpeedEnum
        Auto = 0
        Quiet = 1
        Low = 2
        Medium = 3
        High = 4
        Powerful = 5
        Turbo = 6
        IA1 = 9
        IA2 = 10
        IA3 = 11
        IA4 = 12
        IA5 = 13
        IA6 = 14
    End Enum
    Enum ActiveEnum
        InActive = 0
        Active = 1
    End Enum

    ' Structure: ACStatusMessage
    ' Purpose: Contains all status information for an AC unit
    Public Structure ACStatusMessage
        ' Basic identification
        Dim ACPower As ACPowerEnum
        Dim Number As Byte
        Dim Mode As ModeEnum
        Dim FanSpeed As FanSpeedEnum
        Dim SetPoint As Byte
        Dim Turbo As ActiveEnum
        Dim Bypass As ActiveEnum
        Dim Spill As ActiveEnum
        Dim TimerSet As ActiveEnum
        Dim Defrost As ActiveEnum
        Dim Temperature As Single
        Dim ErrorCode As Short

        ' Method: Parse
        Public Shared Function Parse(data As Byte(), index As Integer) As ACStatusMessage
            ' Read Byte4 and 5 as 16 bit array
            Dim twoBytes(1) As Byte
            Array.Copy(data, index + 3, twoBytes, 0, 2)    ' copy 2 bytes of masks
            Array.Reverse(twoBytes)        ' reverse them to match big-endian format
            Dim bitArray As New BitArray(twoBytes)          ' create a bitarray with them

            ' Create and return populated structure
            Return New ACStatusMessage With {
                .ACPower = CType((data(index) >> 4) And &HF, ACPowerEnum),
                .Number = data(index) And &HF,
                .Mode = CType((data(index + 1) >> 4) And &HF, ModeEnum),
                .FanSpeed = CType(data(index + 1) And &HF, FanSpeedEnum),
                .SetPoint = (data(index + 2) + 100) / 10,
                .Turbo = If(bitArray(11), ActiveEnum.Active, ActiveEnum.InActive),
                .Bypass = If(bitArray(10), ActiveEnum.Active, ActiveEnum.InActive),
                .Spill = If(bitArray(9), ActiveEnum.Active, ActiveEnum.InActive),
                .TimerSet = If(bitArray(8), ActiveEnum.Active, ActiveEnum.InActive),
                .Defrost = If(bitArray(4), ActiveEnum.Active, ActiveEnum.InActive),
                .Temperature = (((CInt(data(index + 4) And &H7) << 8) Or data(index + 5)) - 500) / 10.0,
                .ErrorCode = (CInt(data(index + 6)) << 8) Or data(index + 7)
            }
        End Function
    End Structure
    Public acStatusMsg As ACStatusMessage
    Public Sub GetACStatus()
        ' Create extended message requesting AC capabilities
        Dim requestData() As Byte = CreateMessage(MessageType.Control, {CommandMessages.ACStatus, 0, 0, 0, 0, 0, 0, 0})

        If Not AirTouch5Console.Connected Then Throw New System.Exception("Console not connected. Cannot retrieve AC status.")
        Try
            ' Send request and receive response
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(
                AirTouch5Console.IP, AirTouchPort, requestData)

            ' Parse message and extract AC data
            Dim msg As Message = Message.Parse(response)
            ' Parse raw data into structured format
            acStatusMsg = ACStatusMessage.Parse(msg.data, 8)

            ' Output AC status information
            Debug.WriteLine(
                    $"Power: {acStatusMsg.ACPower} " &
                    $"AC #: {acStatusMsg.Number} " &
                    $"Mode: {acStatusMsg.Mode} " &
                    $"Fan Speed: {acStatusMsg.FanSpeed} " &
                    $"Setpoint: {acStatusMsg.SetPoint} " &
                    $"Turbo: {acStatusMsg.Turbo} " &
                    $"Bypass: {acStatusMsg.Bypass} " &
                    $"Spill: {acStatusMsg.Spill} " &
                    $"Temperature: {acStatusMsg.Temperature} " &
                    $"Defrost: {acStatusMsg.Defrost} " &
                    $"Error Code: {acStatusMsg.ErrorCode} ")
        Catch ex As TimeoutException
            Debug.WriteLine("Request timed out")
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
        Form1.AppendText($"Retrieved AC status information.{vbCrLf}")
    End Sub
End Module
