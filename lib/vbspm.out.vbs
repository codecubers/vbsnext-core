


' ================================== Job: job1 ================================== 

' ================= inline ================= 

Dim script
'Read and assess the parameter supplied
if WScript.Arguments.named.exists("script") Then
WScript.Echo "Argument received: " + WScript.Arguments.named("script")
script = WScript.Arguments.named("script")
Else
Wscript.Echo "/script:  Enter the Script to exeucte"

WScript.Quit
End If

' ================= src : scripts\FS.vbs ================= 
Class FS
    Private objFSO
    
    Private Sub Class_Initialize
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    End Sub

    Public Function GetFSO
        set GetFSO = objFSO
    End Function

    Public Function GetFileDir(ByVal file)
        Wscript.Echo "GetFileDir(" + file + ")"
        Set objFile = objFSO.GetFile(file)
        GetFileDir = objFSO.GetParentFolderName(objFile) 
    End Function
End Class

' set oFs = new FS
' Wscript.Echo oFs.GetFileDir(WScript.ScriptFullName)


' ================================== Job: job2 ================================== 

' ================= inline ================= 

Wscript.Echo "We are in Job2"
