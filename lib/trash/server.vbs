Dim i
i = 0
' Currentdir=Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
do
    req = CFS.ReadFile("C:\Users\nanda\git\xps.local.npm\vbspm\test\request.txt")
    CFS.DeleteFile("C:\Users\nanda\git\xps.local.npm\vbspm\test\request.txt")
    i = i + 1
    WScript.Echo i & " " & req
    Wscript.sleep (1000)
loop while true