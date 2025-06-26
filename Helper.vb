Imports System.IO
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Public Class TcpClientWithTimeout
    Public Shared Function SendAndReceive(
        host As String,
        port As Integer,
        message As Byte(),
        Optional timeoutMs As Integer = 10000,
        Optional encoding As Encoding = Nothing) As Byte()

        If encoding Is Nothing Then encoding = Encoding.ASCII
        Using client As New TcpClient()
            ' Configure client
            client.SendTimeout = timeoutMs
            client.ReceiveTimeout = timeoutMs

            ' Connect with timeout
            Dim connectTask = client.ConnectAsync(host, port)
            If Not connectTask.Wait(timeoutMs) Then
                Throw New TimeoutException("Connection timed out")
            End If

            ' Get network stream
            Using stream = client.GetStream()
                ' Send message
                stream.Write(message, 0, message.Length)

                ' Receive response
                Dim receiveBuffer(4096) As Byte
                Using ms As New MemoryStream()
                    Dim bytesRead As Integer
                    Dim cts As New CancellationTokenSource(timeoutMs)

                    Do
                        Dim readTask = stream.ReadAsync(receiveBuffer, 0, receiveBuffer.Length, cts.Token)
                        bytesRead = If(readTask.Wait(timeoutMs), readTask.Result, 0)

                        If bytesRead > 0 Then
                            ms.Write(receiveBuffer, 0, bytesRead)
                        End If
                    Loop While stream.DataAvailable AndAlso bytesRead > 0 AndAlso Not cts.Token.IsCancellationRequested

                    Return ms.ToArray()
                End Using
            End Using
        End Using
    End Function

    Public Shared Function SendAndReceiveBytes(
        host As String,
        port As Integer,
        message As Byte(),
        Optional timeoutMs As Integer = 10000) As Byte()

        Debug.WriteLine($"Sending: {BitConverter.ToString(message)}")
        Using client As New TcpClient()
            ' Configure client
            client.SendTimeout = timeoutMs
            client.ReceiveTimeout = timeoutMs

            ' Connect with timeout
            Dim connectTask = client.ConnectAsync(host, port)
            If Not connectTask.Wait(timeoutMs) Then
                Throw New TimeoutException("Connection timed out")
            End If

            ' Get network stream
            Using stream = client.GetStream()
                ' Send message
                stream.Write(message, 0, message.Length)

                ' Receive response
                Dim receiveBuffer(4096) As Byte
                Using ms As New MemoryStream()
                    Dim bytesRead As Integer
                    Dim cts As New CancellationTokenSource(timeoutMs)

                    Do
                        Dim readTask = stream.ReadAsync(receiveBuffer, 0, receiveBuffer.Length, cts.Token)
                        bytesRead = If(readTask.Wait(timeoutMs), readTask.Result, 0)

                        If bytesRead > 0 Then
                            ms.Write(receiveBuffer, 0, bytesRead)
                        End If
                    Loop While stream.DataAvailable AndAlso bytesRead > 0 AndAlso Not cts.Token.IsCancellationRequested
                    Debug.WriteLine($"Received: {BitConverter.ToString(ms.ToArray)}")
                    Return ms.ToArray()
                End Using
            End Using
        End Using
    End Function
End Class