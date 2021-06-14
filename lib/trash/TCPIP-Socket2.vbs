bsFunction SendMessageToClient()
Dim address, port, receiveTimeout, socket, connectionsQueueLength, connectedSocket, broadcast,endpoint, byteType, binaryData, maxLength, receivedLength, byteStr, str
address = "127.0.0.1"
 port = 65303
 receiveTimeout = 6000
 connectionsQueueLength = 0

Set Socket = dotNET.System_Net_Sockets.Socket.zctor(dotNET.System_Net_Sockets.AddressFamily.InterNetwork,dotNET.System_Net_Sockets.SocketType.Stream,dotNET.System_Net_Sockets.ProtocolType.Tcp)
Set broadcast = dotNET.System_Net.IPAddress.Parse(address)
Set endpoint = dotNET.System_Net.IPEndPoint.zctor_2(broadcast, port)
'socket.Bind(endpoint)
'socket.Listen(connectionsQueueLength)
'On Error Resume Next

Socket.Connect(endpoint)

'On Error GOTO 0

' Call connectedSocket.SetSocketOption_3(dotNET.System_Net_Sockets.SocketOptionLevel.Socket,dotNET.System_Net_Sockets.SocketOptionName.ReceiveTimeout,receiveTimeout)
' maxLength = 256
' byteType = dotNET.System.Type.GetType("System.Byte")
' binaryData = dotNET.System.Array.CreateInstance(byteType, maxLength)
binaryData = dotNET.System_Text.Encoding.ASCII.GetBytes_2("Tet")
call Socket.SendTo(binaryData,endpoint)

binaryData1 = dotNET.System_Text.Encoding.ASCII.GetBytes_2("Test")
call Socket.SendTo(binaryData1,endpoint)

'connectedSocket = Socket.SendTo(binaryData,endpoint)
''' If Socket.ReceiveFrom(binaryData, endpoint) Then
'''   receivedLength = Socket.ReceiveFrom(binaryData, endpoint)
'''  Else
'''''     An exception occurs if no data is received
'''    receivedLength = 0
'''  End IF


 str = dotNET.System_Text.ASCIIEncoding.get_ASCII().GetString(binaryData)
 str1 = dotNET.System_Text.ASCIIEncoding.get_ASCII().GetString(binaryData)

 'str.SetLength(receivedLength)


 Log.Message ("The received response is " & str)


Socket.Close  
End Function



Public Function server()
Set oSocket = CreateObject("Socket.TCP")
Dim address, port, receiveTimeout, socket, connectionsQueueLength, connectedSocket, broadcast,endpoint, byteType, binaryData, maxLength, receivedLength, byteStr, str
address = "127.0.0.1"
port = 445
receiveTimeout = 6000
connectionsQueueLength = 10
Set Tsocket = dotNET.System_Net_Sockets.Socket.zctor(dotNET.System_Net_Sockets.AddressFamily.InterNetwork, dotNET.System_Net_Sockets.SocketType.Stream,dotNET.System_Net_Sockets.ProtocolType.Tcp)
Set broadcast = dotNET.System_Net.IPAddress.Parse(address)
Set endpoint = dotNET.System_Net.IPEndPoint.zctor_2(broadcast, port)
Tsocket.Bind(endpoint)
Tsocket.Listen(connectionsQueueLength)
Set connectedSocket = Tsocket.Accept()
Call connectedSocket.SetSocketOption_3(dotNET.System_Net_Sockets.SocketOptionLevel.Socket,dotNET.System_Net_Sockets.SocketOptionName.ReceiveTimeout,receiveTimeout)
maxLength = 256
set byteType = dotNET.System.Type.GetType("System.Byte")
set binaryData = dotNET.System.Array.CreateInstance(byteType, maxLength)
'*********************
If connectedSocket.ReceiveFrom(binaryData, endpoint)Then
 receivedLength = connectedSocket.ReceiveFrom(binaryData, endpoint)
Else
'// An exception occurs if no data is received
 receivedLength = 0
End IF

 byteStr = ByteArrayToHexString(binaryData, receivedLength, "0x")

 str = dotNET.System_Text.ASCIIEncoding.get_ASCII().GetString(binaryData)

 str.SetLength(receivedLength)

 Log.Message (receivedLength + " bytes were received: " + byteStr +"String Value: " + str)

 if (str = "TestMessage") Then

    Log.Message "The received data is valid and a positive response will be sent"

   binaryData = dotNET.System_Text.Encoding.ASCII.GetBytes_2("OK")

 else

   Log.Message "The received data is invalid,a negative response will be sent"

   binaryData = dotNET.System_Text.Encoding.ASCII.GetBytes_2("OK")

 End IF
connectedSocket.Close
socket.Close

End Function