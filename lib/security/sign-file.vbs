Set objShell = CreateObject("Wscript.Shell")
strPath = Wscript.ScriptFullName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strPath)
strFolder = objFSO.GetParentFolderName(objFile) 
objShell.CurrentDirectory = strFolder
strPath = ".\signtool-x64.exe sign /f .\ata-authenticode-signer.pfx /p pwd /t http://timestamp.digicert.com " + Wscript.Arguments(0)
objShell.Run strPath, 0, true