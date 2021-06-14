'command line arguments are the message, timeout (seconds), title, and
' (style) bitFlags, respectively
Dim idx, ctr, vitak, oPop
Set vitak = Wscript.Arguments.Unnamed
Set oPop = CreateObject("WScript.Shell")
For Each idx in vitak: ctr = ctr+1: Next
Select Case ctr
Case 4: oPop.popup vitak(0), vitak(1), vitak(2), vitak(3)
Case 3: oPop.popup vitak(0), vitak(1), vitak(2)
Case 2: oPop.popup vitak(0), vitak(1)
Case 1: oPop.popup vitak(0)
Case Else: oPop.popup WScript.ScriptFullName & _
" must be called with 1 to 4 args" & _
vbcrlf & _
"theMessage, seconds, title, bitFlags"
End Select