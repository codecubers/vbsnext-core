Dim text,okButton,Finished

Timer 20,"This is the text for the timer box"
MsgBox "done"

Sub Timer(sec,boxtext)
' Sets up the IE window
    Const READYSTATE_COMPLETE = 4
    Set ie = CreateObject ("InternetExplorer.Application")
        ie.navigate "about:blank"
        ie.toolbar = 0
        ie.MenuBar = 0
        ie.statusbar = 0
        ie.width = 1250
        ie.height = 1250
        ie.left = 100
        ie.top = 100
    Do While ie.ReadyState <> READYSTATE_COMPLETE
        WScript.Sleep 50
    Loop
    ie.visible = True
'Displays the text in the IE window
    Set ieDoc = ie.Document
    iedoc.writeln (boxtext & "<br><br>")
'Creates the area for the countdown timer to be displayed
    Set text = ie.document.createElement("textarea")
    loopcnt = sec - 1
    text.Value = sec
'Displays the timer in the textarea
    ie.document.body.AppendChild text
'Creates the OK button
    Set okButton = ie.document.createElement("input")
        okButton.type = "button"
        okButton.value = "OK"
    ie.document.body.AppendChild okButton
    Set img = ie.document.createElement("img")
    img.src = "https://cdn.shortpixel.ai/client/to_avif,q_glossy,ret_img,w_863/https://covid19fighters.page/wp-content/uploads/2021/05/image-3.png"
    img.width = "1000"
    ie.document.body.AppendChild img
    Finished = False
    Set okButton.onclick = GetRef("OK_Clicked")
' Start the loop for the countdown
    For I = 0 To loopcnt
        WScript.sleep 1000
        sec = sec - 1
        text.Value = sec
        If finished = True Then i = loopcnt
    Next
    finished = True
    
    Do While Not Finished
        WScript.Sleep 50
    Loop
    ie.Quit
End Sub

Sub OK_Clicked()
    Finished = True
End Sub
