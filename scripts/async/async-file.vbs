'https://groups.google.com/g/microsoft.public.scripting.vbscript/c/KMy8CJ3BJAQ
set objFileSys = CreateObject("Scripting.FileSystemObject")
set objWShell = CreateObject("Wscript.shell")
strMsgBoxFile = ".\TempMsgBox.vbs"
' Set objMsgBoxFile = objFileSys.CreateTextFile( strMsgBoxFile, Overwrite )
' objMsgBoxFile.WriteLine( "MsgBox " & Chr(34) & "First line" & Chr(34) )
' objMsgBoxFile.Close
' Set objMsgBoxFile = Nothing
objWShell.Run "cscript.exe " & strMsgBoxFile, 0
Set objMsgBoxFile = objFileSys.GetFile( strMsgBoxFile )
On Error Resume Next
objMsgBoxFile.Delete Force
On Error Goto 0
Set objMsgBoxFile = Nothing
MsgBox "Second line"
