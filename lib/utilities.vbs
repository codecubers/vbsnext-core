
Function ReadFile(file)
  Dim objFSO, objFile
  set fs = new FS
  Set objFSO = fs.GetFSO 'CreateObject("Scripting.FileSystemObject")
  Set objFile = objFSO.OpenTextFile(file)
  ReadFile = objFile.ReadAll()
  objFile.Close
End Function

sub CreateFolder(fol)
  set fs = new FS
  Set fso = fs.GetFSO 'CreateObject("Scripting.FileSystemObject")
  If Not fso.FolderExists(fol) Then
    fso.CreateFolder(fol)
  End If
  set fso = Nothing
End Sub

Sub WriteLog(strFileName, strMessage, overwrite)
  Const ForReading = 1
  Const ForWriting = 2
  Const ForAppending = 8
  Dim mode
  mode = ForWriting
  If Not overwrite Then
    mode = ForAppending
  End If

  'WScript.Echo strFileName
  set fs = new FS
  Set objFSO = fs.GetFSO 'CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFileName) Then
		Set oFile = objFSO.OpenTextFile(strFileName, mode)
	Else
		Set oFile = objFSO.CreateTextFile(strFileName)
	End If
	oFile.WriteLine strMessage
    
	'CLose file
	oFile.Close

	'Clean up
  set oFile = Nothing
	Set objFSO = Nothing
End Sub 

Sub Include(file)
  Wscript.Echo "Include(" + file + ")"
  
  set fs = new FS
  Set objFSO = fs.GetFSO 'CreateObject("Scripting.FileSystemObject")
  if Not objFSO.FileExists(file) Then
    Wscript.Echo "Module " + file + " not found."
    Wscript.Quit
  End If
  ExecuteGlobal ReadFile(file)
End Sub

Sub Import(pkg)
  Wscript.Echo "Import(" + Pkg + ")"
  Include "./node_modules/" + pkg + "/index.vbs"
End Sub

Function jobSrc(file)
  jobSrc = "<script language=""VBScript"" src=""" + file + """/>"
End Function