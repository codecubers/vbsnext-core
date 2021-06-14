Dim fso, outFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("output.txt", True)

' This example requires the Chilkat API to have been previously unlocked.
' See Global Unlock Sample for sample code.

set socket = CreateObject("Chilkat_9_5_0.Socket")

' Connect to port 5555 of localhost.
' The string "localhost" is for testing on a single computer.
' It would typically be replaced with an IP hostname, such
' as "www.chilkatsoft.com".
ssl = 0
maxWaitMillisec = 20000
success = socket.Connect("localhost",5555,ssl,maxWaitMillisec)
If (success <> 1) Then
    outFile.WriteLine(socket.LastErrorText)
    WScript.Quit
End If

' Set maximum timeouts for reading an writing (in millisec)
socket.MaxReadIdleMs = 10000
socket.MaxSendIdleMs = 10000

' Pretend, for the sake of the example, that the
' ficticious server is going to send a "Hello World!" 
' after accepting the connection.  
' Note: Technically, the ReceiveString may not receive the
' complete string, although it's highly probable given the short
' length of the "Hello World!" message. 
' See this Chilkat blog post for more information:
' http://www.cknotes.com/?p=302
receivedMsg = socket.ReceiveString()
If (socket.LastMethodSuccess <> 1) Then
    outFile.WriteLine(socket.LastErrorText)
    WScript.Quit
End If

' Close the connection with the server
' Wait a max of 20 seconds (20000 millsec)
success = socket.Close(20000)

outFile.WriteLine(receivedMsg)

outFile.Close
