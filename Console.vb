Imports System.Net
Imports System.Net.Sockets
Imports System.Text

Friend Module Console
    ' ==============================================================
    ' Module: Console
    ' Purpose: Handles discovery and communication with AirTouch consoles
    '          using UDP broadcast and response protocols
    '
    ' Features:
    ' - Discovers AirTouch consoles on the local network
    ' - Manages UDP communication with proper error handling
    ' - Maintains console state information
    ' - Provides clean initialization and disposal of resources
    '
    ' Usage:
    ' 1. Call DiscoverConsole() to initiate discovery
    ' 2. Access discovered console information via AirTouch5Console
    '
    ' Protocol Details:
    ' - Uses UDP port 49005 for broadcast and responses
    ' - Broadcast message format: "::REQUEST-POLYAIRE-AIRTOUCH-DEVICE-INFO:;"
    ' - Response format: "IP,ConsoleID,AirTouch5,AirTouchID,DeviceName"
    '
    ' Dependencies:
    ' - Requires Form1 with timeoutMs property for receive timeout
    ' ==============================================================

    Public AirTouchPort = 9005 ' Port for AirTouch console HTTP requests
    Public timeoutMs As Integer = 5000
    ' Network communication constants
    Private ReadOnly udpPort As Integer = 49005 ' Standard port for AirTouch communication
    Private ReadOnly broadcastMessage As String = "::REQUEST-POLYAIRE-AIRTOUCH-DEVICE-INFO:;" ' Discovery message

    ' Communication objects
    Private udpClient As UdpClient ' Make client global to properly dispose
    Private responseMessage As String = String.Empty ' Stores the last received message

    ' ==============================================================
    ' Structure: AirTouchConsole
    ' Purpose: Stores information about a discovered AirTouch console
    ' Fields:
    ' - IP: Console's network address
    ' - ConsoleID: Unique hardware identifier
    ' - AirTouch5: Version identifier
    ' - AirTouchID: Device-specific identifier
    ' - DeviceName: User-assigned name
    ' - Connected: Status flag
    ' ==============================================================
    Structure AirTouchConsole
        ' Capture details of discovered AirTouch console
        Dim IP As String
        Dim ConsoleID As String
        Dim AirTouch5 As String
        Dim AirTouchID As String
        Dim DeviceName As String
        Dim Connected As Boolean

        ' Constructor to initialize with connection status
        Public Sub New(connect As Boolean)
            IP = String.Empty
            ConsoleID = String.Empty
            AirTouch5 = String.Empty
            AirTouchID = String.Empty
            DeviceName = String.Empty
            Connected = connect
        End Sub
    End Structure

    ' Global instance to store discovered console information
    Public AirTouch5Console As New AirTouchConsole(False) ' Initialize with not connected

    ' ==============================================================
    ' Sub: DiscoverConsole
    ' Purpose: Main discovery routine that coordinates broadcast and response handling
    ' Flow:
    ' 1. Initializes UDP client if needed
    ' 2. Starts background receive task
    ' 3. Sends broadcast message
    ' 4. Processes response
    ' 5. Cleans up resources
    ' ==============================================================

    Public Async Function DiscoverConsole() As Task(Of Boolean)
        Form1.TextBox1.AppendText("Starting discovery...")

        ' Initialize client if needed
        If udpClient Is Nothing Then
            udpClient = New UdpClient With {
                .EnableBroadcast = True ' Allow broadcast messages
            }
            ' Configure socket options
            udpClient.Client.SetSocketOption(SocketOptionLevel.Socket,
                                           SocketOptionName.ReuseAddress,
                                           True)
            ' Bind to any available IP on the specified port
            udpClient.Client.Bind(New IPEndPoint(IPAddress.Any, udpPort))
        End If

        ' Start receiving responses in background
        Dim receiveTask = ReceiveUdpResponsesAsync()

        ' Small delay to ensure receiver is ready before sending
        Await Task.Delay(100)

        ' Send broadcast message to discover consoles
        SendUdpBroadcast()

        ' Wait for receive operation to complete
        Await receiveTask
        Debug.WriteLine("Discovery completed", responseMessage)

        ' Process the response (comma-separated values)
        Dim decode = responseMessage.Split(",")     ' split into pieces
        With AirTouch5Console
            .IP = decode(0)
            .ConsoleID = decode(1)
            .AirTouch5 = decode(2)
            .AirTouchID = decode(3)
            .DeviceName = decode(4)
            .Connected = True
        End With
        Form1.TextBox1.AppendText($"Found AirTouch Console: {AirTouch5Console.DeviceName} at IP {AirTouch5Console.IP}{vbCrLf}")

        ' Clean up client after use
        If udpClient IsNot Nothing Then
            udpClient.Close()
            udpClient = Nothing
        End If
        Return (True) ' Indicate discovery was successful
    End Function

    ' ==============================================================
    ' Sub: SendUdpBroadcast
    ' Purpose: Sends the discovery broadcast message
    ' Features:
    ' - Handles UDP client initialization
    ' - Includes error handling and resource cleanup
    ' ==============================================================
    Sub SendUdpBroadcast()
        Try
            ' Create and configure client once
            If udpClient Is Nothing Then
                udpClient = New UdpClient With {
                    .EnableBroadcast = True
                }
                udpClient.Client.SetSocketOption(SocketOptionLevel.Socket,
                                              SocketOptionName.ReuseAddress,
                                              True)
                udpClient.Client.Bind(New IPEndPoint(IPAddress.Any, udpPort))
            End If

            ' Convert message to bytes and send to broadcast address
            Dim bytes As Byte() = Encoding.ASCII.GetBytes(broadcastMessage)
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

    ' ==============================================================
    ' Function: ReceiveUdpResponsesAsync
    ' Purpose: Listens for responses to the discovery broadcast
    ' Returns: Boolean indicating if valid response was received
    ' Features:
    ' - Implements timeout handling
    ' - Updates UI thread safely
    ' - Filters out our own broadcast messages
    ' ==============================================================
    Async Function ReceiveUdpResponsesAsync() As Task(Of Boolean)
        Try
            ' Set receive timeout from main form
            udpClient.Client.ReceiveTimeout = timeoutMs

            While True
                ' Wait for incoming message
                Dim receiveResult = Await udpClient.ReceiveAsync()
                responseMessage = Encoding.ASCII.GetString(receiveResult.Buffer)

                ' Update UI thread safely
                Form1.Invoke(Sub()
                                 ' Update UI here
                                 Debug.WriteLine($"Received: {responseMessage}")
                             End Sub)

                ' Check if the response contains the expected data (ignore our own broadcasts)    
                If Not responseMessage = broadcastMessage Then
                    Return True ' we got a valid response
                End If
            End While
        Catch ex As SocketException When ex.ErrorCode = 10060
            Debug.WriteLine("Receive timeout reached")
            Return False
        Catch ex As Exception
            Debug.WriteLine($"Receive error: {ex.Message}")
            Return False
        End Try
        Return True
    End Function
End Module