' Set objShell = CreateObject("Wscript.Shell")
' Set objFSO = CreateObject("Scripting.FileSystemObject")
' strPath = Wscript.ScriptFullName
' Set objFile = objFSO.GetFile(strPath)
' strFolder = objFSO.GetParentFolderName(objFile) 
' objShell.CurrentDirectory = strFolder
' strPath = ".\signtool-x64.exe sign /f .\ata-authenticode-signer.pfx /p pwd /t http://timestamp.digicert.com " + Wscript.Arguments(0)
' objShell.Run strPath, 0, true

Class Signtool

    private cWShell

    private Sub Class_Initialize
        set cWShell = new WShell
        if cWShell is nothing then
            Wscript.Echo "Signer Class: Unable to initialize WShell class."
            Wscript.Quit
        end if
    End Sub

    public Function Sign(file, pwd)
        Wscript.Echo "Signing file: " & file

        Dim signtool: signtool = "C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\signtool\signtool-x64.exe"
        Dim cert: cert = "C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\signtool\ata-authenticode-signer.pfx"
        Dim timestamp: timestamp = "http://timestamp.digicert.com"
        Dim strPath: strPath = signtool & " sign /f " & cert & " /p " & pwd & " /t " & timestamp & " " + file
        cWShell.Exec(strPath)
    End Function

End Class