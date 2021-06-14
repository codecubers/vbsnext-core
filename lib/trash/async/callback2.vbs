' Set objSession = CreateObject("Microsoft.Update.Session")
' Set objSearcher = objSession.CreateUpdateSearcher
' WScript.ConnectObject objSearcher, "searcherCallBack_"
' objSearcher.BeginSearch "abc"

' sub searcherCallBack_Invoke()
'     ' handle the callback
'     msgbox "callback"
' end sub

Class callbackEvents
    public sub started(str)
        WScript.Echo str
        for i = 0 to 10
            Wscript.Echo "ok " & i
            Wscript.sleep 100
        next
    End Sub
    public sub inProgress(str)
        WScript.Echo str
    End Sub
    public sub done(str)
        WScript.Echo str
    End Sub
End Class
    
    
Function TaskScheduler(callback)
    callback.started "I'm started"
    for i = 0 to 10
        callback.inProgress "I'm in progress " & i
        Wscript.sleep 100
    next
    callback.done "I'm done"
End Function


Set cb = New callbackEvents
TaskScheduler cb
Set cb = Nothing