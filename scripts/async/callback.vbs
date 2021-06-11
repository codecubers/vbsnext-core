' Set objSession = CreateObject("Microsoft.Update.Session")
' Set objSearcher = objSession.CreateUpdateSearcher
' WScript.ConnectObject objSearcher, "searcherCallBack_"
' objSearcher.BeginSearch "abc"

' sub searcherCallBack_Invoke()
'     ' handle the callback
'     msgbox "callback"
' end sub

Class EchoFileName
    public Default sub echo(str)
        WScript.Echo str
    End Sub
End Class
    
    
Function DirWalker(fs, strRootDir, funcEventConsumer)
    Const Directory = 16
    Dim f, f1, s, sf
    Set f = fs.GetFolder(strRootDir)

    Set sf = f.SubFolders
    For Each f1 in sf
        If (f1.Attributes And Directory) Then
            DirWalker fs, f1, funcEventConsumer
            ' ==== SENDING EVENT TO CLASS ===
            funcEventConsumer f1.path
        End If
    Next

    Set files = f.Files
    For Each f1 In files
        funcEventConsumer f1.path
    Next
End Function


Set func = New EchoFileName
Set fs = CreateObject("Scripting.FileSystemObject")
DirWalker fs, "..\", func
Set func = Nothing