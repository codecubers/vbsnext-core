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
	public Sub setDir(s)
		dir = s
	End Sub
	
	Public Function GetFSO
		Set GetFSO = objFSO
	End Function

    ' ===================== Sub Routines =====================

	Public Sub CreateFolder(fol)
		If Not objFSO.FolderExists(fol) Then
			objFSO.CreateFolder(fol)
		End If
	End Sub
	
	Public Sub WriteLog(strFileName, strMessage, overwrite)
		Const ForReading = 1
		Const ForWriting = 2
		Const ForAppending = 8
		Dim mode
		
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
		Set objFile = objFSO.GetFile(file)
		GetFileDir = objFSO.GetParentFolderName(objFile) 
	End Function
	
	Public Function ReadFile(file)
		If Not FileExists(file) Then 
			Wscript.Echo "File " & file & " does not exists."
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
		on Error resume next
		objFSO.DeleteFile(file)
		On Error Goto 0
	End Sub
End Class
