' ==============================================================================================
' Implementation of several use cases of FileSystemObject into this class
' Author: Praveen Nandagiri (pravynandas@gmail.com)
' ==============================================================================================

Class FSO
	Private dir
	Private objFSO
	
	Private Sub Class_Initialize
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		dir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
	End Sub
	
	' Update the current directory of the instance if needed
	Public Sub setDir(s)
		dir = s
	End Sub
	
	Public Function getDir
		getDir = dir
	End Function
	
	Public Function GetFSO
		Set GetFSO = objFSO
	End Function
	
	Public Function FolderExists(fol)
		FolderExists = objFSO.FolderExists(fol)
	End Function
	
	' ===================== Sub Routines =====================
	
	Public Function CreateFolder(fol)
		CreateFolder = False
		If FolderExists(fol) Then
			CreateFolder = True
		Else
			objFSO.CreateFolder(fol)
			CreateFolder = FolderExists(fol)
		End If
	End Function
	
	Public Sub WriteFile(strFileName, strMessage, overwrite)
		Const ForReading = 1
		Const ForWriting = 2
		Const ForAppending = 8
		Dim mode
		Dim oFile
		
		mode = ForWriting
		If Not overwrite Then
			mode = ForAppending
		End If
		
		If objFSO.FileExists(strFileName) Then
			Set oFile = objFSO.OpenTextFile(strFileName, mode)
		Else
			Set oFile = objFSO.CreateTextFile(strFileName)
		End If
		oFile.WriteLine strMessage
		
		oFile.Close
		
		Set oFile = Nothing
	End Sub 
	
	' ===================== Function Routines =====================
	
	Public Function GetFileDir(ByVal file)
		EchoDX "GetFileDir( %x )", Array(file)
		Dim objFile
		Set objFile = objFSO.GetFile(file)
		GetFileDir = objFSO.GetParentFolderName(objFile) 
	End Function
	
	Public Function GetFilePath(ByVal file)
		EchoDX "GetFilePath( %x )", Array(file)
		Dim objFile
		On Error Resume Next
		Set objFile = objFSO.GetFile(file)
		On Error GoTo 0
		If IsObject(objFile) Then
			GetFilePath = objFile.Path 
		Else
			EchoDX "File %x not found; searching in directory %x", Array(file,dir)
			On Error Resume Next
			Set objFile = objFile.GetFile(objFSO.BuildPath(dir, file))
			On Error GoTo 0
			If IsObject(objFile) Then
				GetFilePath = objFile.Path 
			Else
				GetFilePath = "File [" & file & "] Not found"
			End If
		End If
	End Function
	
	''' <summary>Returns a specified number of characters from a string.</summary>
	''' <param name="file">File Name</param>
	Public Function GetFileName(ByVal file)
		GetFileName = objFSO.GetFile(file).Name
	End Function
	
	Public Function GetFileExtn(file)
		GetFileExtn = ""
		On Error Resume Next
		GetFileExtn = LCASE(objFSO.GetExtensionName(file))
		On Error GoTo 0
	End Function
	
	Public Function GetBaseName(ByVal file)
		GetBaseName = Replace(GetFileName(file), "." & GetFileExtn(file), "")
	End Function
	
	Public Function ReadFile(file)
		file = putil.Resolve(file)
		EchoDX "---> File resolved to: %x", Array(file)
		If Not FileExists(file) Then 
			Wscript.Echo "---> File " & file & " does not exists."
			ReadFile = ""
			Exit Function
		End If
		Dim objFile: Set objFile = objFSO.OpenTextFile(file)
		ReadFile = objFile.ReadAll()
		objFile.Close
	End Function
	
	Public Function FileExists(file)
		FileExists = objFSO.FileExists(file)
	End Function
	
	Public Sub DeleteFile(file)
		On Error Resume Next
		objFSO.DeleteFile(file)
		On Error GoTo 0
	End Sub

	Public Sub MoveFile(src, dest)
		On Error Resume Next
		objFSO.MoveFile src, dest
		On Error GoTo 0
	End Sub
	
End Class