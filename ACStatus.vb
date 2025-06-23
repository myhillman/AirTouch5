Module ACStatus
    Enum PowerStateEnum
        Off = 0
        [On] = 1
        OffAlways = 2
        OnAlways = 3
        Sleep = 5
    End Enum
    Enum ModeEnum
        Auto = 0
        Cool = 4
        Fan = 3
        Dry = 2
        Heat = 1
        AutoHeat = 8
        AuroCool = 9
    End Enum
    Enum FanSpeedEnum
        Auto = 0
        Turbo = 6
        Powerful = 5
        High = 4
        Medium = 3
        Low = 2
        Quiet = 1
        IA1 = 9
        IA2 = 10
        IA3 = 11
        IA4 = 12
        IA5 = 13
    End Enum
    Enum ActiveEnum
        InActive = 0
        Active = 1
    End Enum

    ' Structure: ACability
    ' Purpose: Contains all capability information for an AC unit
    Public Structure ACStatusMessage
        ' Basic identification
        Dim PowerState As PowerStateEnum
        Dim Number As Byte
        Dim Mode As ModeEnum
        Dim FanSpeed As FanSpeedEnum
        Dim SetPoint As Byte
        Dim TurboActive As ActiveEnum
        Dim BypassActive As ActiveEnum
        Dim SpillActive As ActiveEnum
        Dim TimerSet As ActiveEnum
        Dim Defrost As ActiveEnum
        Dim Temperature As Short
        Dim ErrorCode As Short

        ' Method: Parse
        Public Shared Function Parse(data As Byte(), index As Integer) As ACStatusMessage
            ' Create and return populated structure
            Return New ACStatusMessage With {
                .PowerState = CType((data(index) >> 4) And &HF, PowerStateEnum),
                .Number = data(index) And &HF,
                .Mode = CType((data(index + 1) >> 4) And &HF, ModeEnum),
                .FanSpeed = CType(data(index + 1) And &HF, FanSpeedEnum),
                .SetPoint = (data(index + 2) + 100) / 10,
                .TurboActive = CType((data(index + 3) >> 3) And &H1, ActiveEnum),
                .BypassActive = CType((data(index + 3) >> 2) And &H1, ActiveEnum),
                .SpillActive = CType((data(index + 3) >> 1) And &H1, ActiveEnum),
                .TimerSet = CType(data(index + 3) And &H1, ActiveEnum),
                .Defrost = CType((data(index + 4) >> 5) And &H3, ActiveEnum),
                .Temperature = (((CInt(data(index + 4) And &H7) << 8) Or data(index + 5)) - 500) / 10,
                .ErrorCode = (CInt(data(index + 6)) << 8) Or data(index + 7)
            }
        End Function
    End Structure
    Public acStatusMsg As ACStatusMessage
    Public Sub GetACStatus()
        ' Create extended message requesting AC capabilities
        Dim requestData() As Byte = CreateMessage(MessageType.Control, {&H23, 0, 0, 0, 0, 0, 0, 0})

        If Not AirTouch5Console.Connected Then Throw New System.Exception("Console not connected. Cannot retrieve AC status.")
        Try
            ' Send request and receive response
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(
                AirTouch5Console.IP, AirTouchPort, requestData)

            ' Parse response messages
            Dim messages = ParseMessages(response)

            For Each msg In messages
                ' Parse message and extract AC data
                Dim m = Message.Parse(msg)
                ' Parse raw data into structured format
                acStatusMsg = ACStatusMessage.Parse(m.data, 8)

                ' Output AC status information
                Debug.WriteLine(
                    $"Power: {acStatusMsg.PowerState} " &
                    $"AC #: {acStatusMsg.Number} " &
                    $"Mode: {acStatusMsg.Mode} " &
                    $"Fan Speed: {acStatusMsg.FanSpeed} " &
                    $"Setpoint: {acStatusMsg.SetPoint} " &
                    $"Turbo: {acStatusMsg.TurboActive} " &
                    $"Bypass: {acStatusMsg.BypassActive} " &
                    $"Spill: {acStatusMsg.SpillActive} " &
                    $"Temperature: {acStatusMsg.Temperature} " &
                    $"Defrost: {acStatusMsg.Defrost} " &
                    $"Error Code: {acStatusMsg.ErrorCode} ")
            Next
        Catch ex As TimeoutException
            Debug.WriteLine("Request timed out")
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
    End Sub
End Module
