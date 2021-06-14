Function Test()
	Dim address, port, receiveTimeout, socket, connectionsQueueLength, connectedSocket, broadcast, endpoint, byteType, binaryData, maxLength, receivedLength, byteStr, str
	address = "122.188.113.56"
	port = 3000
	receiveTimeout = xxxx
	connectionsQueueLength = 0
	Set Socket = dotNET.System_Net_Sockets.Socket.zctor(dotNET.System_Net_Sockets.AddressFamily.InterNetwork,dotNET.System_Net_Sockets.SocketType.Stream,dotNET.System_Net_Sockets.ProtocolType.Tcp)
	Set broadcast = dotNET.System_Net.IPAddress.Parse(address)
	Set endpoint = dotNET.System_Net.IPEndPoint.zctor_2(broadcast, port)
	socket.Bind(endpoint)
	socket.Listen(connectionsQueueLength)
	connectedSocket = socket.Accept()
	Call connectedSocket.SetSocketOption_3(dotNET.System_Net_Sockets.SocketOptionLevel.Socket,dotNET.System_Net_Sockets.SocketOptionName.ReceiveTimeout,receiveTimeout)
	maxLength = 256
	byteType = dotNET.System.Type.GetType("System.Byte")
	binaryData = dotNET.System.Array.CreateInstance(byteType, maxLength)
	If connectedSocket.ReceiveFrom(binaryData, endpoint)Then
		receivedLength = connectedSocket.ReceiveFrom(binaryData, endpoint)
	Else
		'// An exception occurs if no data is received
		receivedLength = 0
	End If
	byteStr = ByteArrayToHexString(binaryData, receivedLength, "0x")
	str = dotNET.System_Text.ASCIIEncoding.get_ASCII().GetString(binaryData)
	str.SetLength(receivedLength)
	Log.Message (receivedLength + " bytes were received: " + byteStr +"String Value: " + str)
	If (str = "TestMessage") Then
		Log.Message "The received data is valid and a positive response will be sent"
		binaryData = dotNET.System_Text.Encoding.ASCII.GetBytes_2("OK")
	Else 
		Log.Message "The received data is invalid,a negative response will be sent"
		binaryData = dotNET.System_Text.Encoding.ASCII.GetBytes_2("OK")
	End If
	connectedSocket.Close
	socket.Close   
End Function
Function ByteArrayToHexString(byteArray, byteCount, prefix)
	Dim i, string
	string = ""
	i = 1 
	Do While i = (i < byteArray.Length) & (i < byteCount)
		string = dotNET.System.String.Format(prefix + "{0:X} ", byteArray.Get(i))
		return string  
	Loop
End Function
Test