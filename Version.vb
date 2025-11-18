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
    Public VersionInfo As New VersionStruct
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
        Dim Latest As Boolean           ' True if software is latest
        Dim VersionText As String       ' Version text
        ' Read-only calculated property
        Public ReadOnly Property AnnotatedVersion As String
            Get
                Return $"{VersionText}{If(Latest, "", "*")}"  ' add * if update available
            End Get
        End Property
        Public Shared Function Parse(data As Byte()) As VersionStruct
            ' Create and return populated structure
            ' Validate data length
            If data Is Nothing OrElse data.Length < 4 Then
                Throw New ArgumentException("Invalid version data")
            End If
            Dim Latest As Boolean = (data(2) = 0)       ' True if latest version
            Dim versionLength As Integer = data(3)
            ' Check if we have enough bytes for the version string
            If data.Length < 4 + versionLength Then
                Throw New ArgumentException("Version data is truncated")
            End If
            Dim Version As String = System.Text.Encoding.ASCII.GetString(data, 4, versionLength)
            Return New VersionStruct With {
                .Latest = Latest,
                .VersionText = Version
            }
        End Function
    End Structure

    Public Sub GetVersion()
        ' Create extended message with version request command (FF 30)
        Dim requestData() As Byte = CreateMessage(MessageType.Extended, {&HFF, &H30})

        If Not AirTouch5Console.Connected Then Throw New System.Exception("Console Not connected. Cannot retrieve version.")
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
            Debug.WriteLine("Console Version: " & VersionInfo.AnnotatedVersion)

        Catch ex As TimeoutException
            ' Handle communication timeout
            Debug.WriteLine("Request timed out")

        Catch ex As Exception
            ' Handle all other exceptions
            Debug.WriteLine("Error: " & ex.Message)
        End Try
        Form1.AppendText($"Console version {Versioninfo.annotatedversion}{vbCrLf}")
    End Sub
End Module