'https://groups.google.com/g/microsoft.public.scripting.vbscript/c/KMy8CJ3BJAQ
set objWShell = CreateObject("Wscript.shell")
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    objWShell.Run "cscript.exe " & WScript.ScriptFullName & " ASYNC_ROUTINE", Hidden
    MsgBox "Second line"
ElseIf objArgs.Item(0) = "ASYNC_ROUTINE" Then
    MsgBox "First line"
    WScript.Quit
End If