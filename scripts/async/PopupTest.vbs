
'OK, and now ... what you've been waiting for (put both in the same
'directory and run PopupTest.vbs):

' ------------ start PopupTest.vbs ------------
AsyncPopup "Message 1", "Popup testing"
Set ie = CreateObject("InternetExplorer.Application")
ie.visible = true

ie.Navigate ("yahoo.com")
Do Until ie.readyState=4: WScript.Sleep(500): Loop
AsyncPopup "Message 2", "Popup testing"

'There must be at least one navigation before next line
ie.Navigate2 ("javascript:""<html><head><title>Demo</title></head>" & _
"<body><button type=button accesskey='c' " & _
"onClick='self.close()'><u>C</u>lose</button>" & _
"</body></html>""")
Do Until ie.readyState=4: WScript.Sleep(500): Loop
ie.Document.ParentWindow.opener="me" 'allows self.close()
AsyncPopup "Message 3", "Popup testing"


Sub AsyncPopup (msg, title)
const myType = &H0 'multiple popups
' const myType = &H20000 'always on top
CreateObject("WScript.Shell").Run _
"CScript.exe //NOLOGO " & _
popupPath & " """ & msg & _
""" 0 """ & title & """ " & (myType + 49), 0, false
End Sub

Function popupPath()
Dim i
popupPath = WScript.ScriptFullName
For i=Len(popupPath) to 1 step -1
If Mid(popupPath,i,1)="\" Then _
popupPath = Mid(popupPath,1,i) & _
"AsyncPopup.vbs": Exit Function
Next
End Function
' ------------- end PopupTest.vbs -------------