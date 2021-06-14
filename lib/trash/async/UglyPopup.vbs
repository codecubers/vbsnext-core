uglyPopup "Message 1"
uglyPopup "Message 2"
MsgBox "Message 3"

sub uglyPopup(msg)
Set ie = CreateObject("InternetExplorer.Application")
ie.height = 1: ie.width = 1
ie.visible = true 'false => alert not shown
ie.Navigate ("about:blank")
Do Until ie.readyState=4: WScript.Sleep(500): Loop
ie.Document.parentWindow.opener = "me" 'allows self.close
'next line should be revised to escape quotes in msg
ie.Document.parentWindow.setTimeout _
"alert('" & msg & "');self.close();", 10
End Sub