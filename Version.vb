Imports System.Text

' ==============================================================
' Module: Version
' Purpose: Handles retrieval and storage of AirTouch console version information
' Features:
' - Requests version information from AirTouch console
' - Parses and stores version string
' - Provides debug output of version information
' ==============================================================
Module Version
    ' Public variable to store the retrieved version string
    Public Version As String = String.Empty

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
            Dim messages = ParseMessages(response)

            ' Process each message in response
            For Each msg In messages
                ' Parse message structure
                Dim m = Message.Parse(msg)

                ' Extract version string:
                ' - m.data(3) contains version string length
                ' - Version starts at m.data(4)
                Version = Encoding.UTF8.GetString(m.data, 4, m.data(3))

                ' Output version to debug
                Debug.WriteLine("Console Version: " & Version)
            Next

        Catch ex As TimeoutException
            ' Handle communication timeout
            Debug.WriteLine("Request timed out")

        Catch ex As Exception
            ' Handle all other exceptions
            Debug.WriteLine("Error: " & ex.Message)
        End Try
    End Sub
End Module