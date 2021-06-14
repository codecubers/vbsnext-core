Dim bPasswordBoxWait
msgbox PasswordBox("Enter your password.")

Function PasswordBox(sTitle) 
  set oIE = CreateObject("InternetExplorer.Application") 

  With oIE 
    .FullScreen = False
    .ToolBar   = False : .RegisterAsDropTarget = False 
    .StatusBar = False : .Navigate("about:blank") 
    .Width = 400
	.Height = 200
	.Left = 300
	.Top = 200
	.Visible = 1

    While .Busy : WScript.Sleep 100 : Wend 
    With .document 
      With .ParentWindow 
        .resizeto 100,50 
        .moveto .screen.width/2-200, .screen.height/2-50 
      End With 
      .WriteLn("<html><body bgColor=Silver><center>") 
      .WriteLn("[b]" & sTitle & "[b]") 
      .WriteLn("Password <input type=password id=pass>  " & _ 
               "<button id=but0>Submit</button>") 
      .WriteLn("</center></body></html>") 
      With .ParentWindow.document.body 
        .scroll="no" 
        .style.borderStyle = "outset" 
        .style.borderWidth = "3px" 
      End With 
      .all.but0.onclick = getref("PasswordBox_Submit") 
      .all.pass.focus 
      oIE.Visible = True 
      bPasswordBoxOkay = False : bPasswordBoxWait = True 
      On Error Resume Next 
      While bPasswordBoxWait 
        WScript.Sleep 100 
        if oIE.Visible Then bPasswordBoxWait = bPasswordBoxWait 
        if Err Then bPasswordBoxWait = False 
      Wend 
      PasswordBox = .all.pass.value 
    End With ' document 
    .Visible = False 
  End With   ' IE 
End Function 


Sub PasswordBox_Submit() 
  bPasswordBoxWait = False 
End Sub