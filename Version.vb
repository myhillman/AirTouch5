Imports System.Text

' ==============================================================
' Module: Version
' Purpose: Handles retrieval and storage of AirTouch console version information
' Features:
' - Requests version information from AirTouch console
' - Parses and stores version string
' - Provides debug output of version information
' ==============================================================
Friend Module Version
    ' Method: GetVersion
    ' Purpose: Retrieves the console version from the AirTouch system
    ' Flow:
    ' 1. Creates an extended message requesting version info (FF 30)
    ' 2. Sends request and waits for response
    ' 3. Parses response to extract version string
    ' 4. Stores version in public variable and outputs to debug
    ' Error Handling:
    ' - Catches and reports timeout exceptions
    ' - Catches and reports general exceptions
    Public Structure VersionStruct
        Dim UpdateSign As Byte
        Dim VersionText As String
        Public Shared Function Parse(data As Byte()) As VersionStruct
            ' Create and return populated structure
            Return New VersionStruct With {
                .UpdateSign = data(2),
                .VersionText = Encoding.UTF8.GetString(data, 4, data(3))
            }
        End Function
    End Structure
    Public VersionInfo As New VersionStruct
    Public Sub GetVersion()
        ' Create extended message with version request command (FF 30)
        Dim requestData() As Byte = CreateMessage(MessageType.Extended, {&HFF, &H30})

        If Not AirTouch5Console.Connected Then Throw New System.Exception("Console not connected. Cannot retrieve version.")
        Try
            ' Send request and receive response
            Dim response As Byte() = TcpClientWithTimeout.SendAndReceiveBytes(
                AirTouch5Console.IP,
                AirTouchPort,
                requestData)

            ' Debug output of raw response
            Debug.WriteLine("Response: " & Encoding.UTF8.GetString(response))

            ' Parse response into individual messages
            Dim msg As Message = Message.Parse(response)

            ' Parse message structure
            VersionInfo = VersionStruct.Parse(msg.data)

            ' Output version to debug
            Debug.WriteLine("Console Version: " & VersionInfo.VersionText)

        Catch ex As TimeoutException
            ' Handle communication timeout
            Debug.WriteLine("Request timed out")

        Catch ex As Exception
            ' Handle all other exceptions
            Debug.WriteLine("Error: " & ex.Message)
        End Try
        Form1.TextBox1.AppendText($"Retrieved version information.{vbCrLf}")
    End Sub
End Module