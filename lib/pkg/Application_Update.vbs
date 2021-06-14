'TODO: Error handling for each sheet and validation
Dim Application
Dim ActiveWorkbook
Dim wkbTarget
Dim objFSO 
Dim objFile 
Dim szTargetWorkbook 
Dim szImportPath 
Dim szFileName 
Dim cmpComponents 
Const vbext_ct_Document = 100
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = 2, TristateTrue= 1, TristateFalse= 0

Set objFSO = CreateObject("scripting.filesystemobject")
'Get the path to the folder with modules
sPath = FolderWithVBAProjectFiles
If sPath = "Error" Then
	MsgBox "Import Folder not exist"
	Wscript.Quit
End If

''' NOTE: Path where the code modules are located.
szImportPath = sPath & "\"

If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
   MsgBox "There are no files to import"
   Wscript.Quit
End If

If objFSO.FileExists(szImportPath & "Version.txt") Then
	set oVersionFile = objFSO.OpenTextFile(szImportPath & "Version.txt", ForReading, TristateFalse)
	sNewVersion = Trim(oVersionFile.ReadLine)
	oVersionFile.close
End if



set Application = createobject("Excel.Application")
Application.DisplayAlerts = False
if Application is nothing then
	msgbox "Unable to create Excel application object."
	WScript.Quit
end if

set ActiveWorkbook = Application.workbooks.Open(ObjFSO.GetFolder(".") & "\Scheduler.xlsm")
if ActiveWorkbook is nothing then
	msgbox "Unable to create Excel Workbook object."
	Application.Quit
	WScript.Quit
end if
sOldVersion = replace(Trim(ActiveWorkbook.worksheets("App").Range("N1").value), vbcr, "")
'Msgbox "["& sNewVersion & "],[" & sOldVersion & "]"
If sNewVersion <= sOldVersion Then
	Msgbox "Current version of the application (" & sOldVersion & ") is higher than (or equal to) the update version (" & sNewVersion & "). Please recheck. No Update is taken place.", _
	        vbExclamation, "This update seems to be out-dated."
	Activeworkbook.close
	Application.Quit	
	Wscript.Quit
End If

''' NOTE: This workbook must be open in Excel.
szTargetWorkbook = ActiveWorkbook.Name
Set wkbTarget = Application.Workbooks(szTargetWorkbook)

If wkbTarget.VBProject.Protection = 1 Then
	MsgBox "The VBA in this workbook is protected," & _
			"not possible to Import the code"
	Activeworkbook.close
	Application.Quit			
	Wscript.Quit
End If


'Delete all modules/Userforms from the ActiveWorkbook
Call DeleteVBAModulesAndUserForms

Set cmpComponents = wkbTarget.VBProject.VBComponents

''' Import all the code modules in the specified path
''' to the ActiveWorkbook.
For Each objFile In objFSO.GetFolder(szImportPath).Files

	If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
		(objFSO.GetExtensionName(objFile.Name) = "frm") Or _
		(objFSO.GetExtensionName(objFile.Name) = "bas") Then
		cmpComponents.Import objFile.Path
	End If
	
Next 

ActiveWorkbook.worksheets("App").Range("N1").value = sNewVersion
ActiveWorkbook.Save

MsgBox "Application Schedule.xlsm is updated successfully." & vbcrlf & _
       "Old Version: " & sOldVersion & vbcrlf & _
	   "New Version: " & sNewVersion


ActiveWorkbook.close
set ActiveWorkbook = nothing
Application.DisplayAlerts = True
Application.quit
set Application = nothing

Function FolderWithVBAProjectFiles()
    Dim WshShell 
    Dim SpecialPath 

    'Set WshShell = CreateObject("WScript.Shell")

    'SpecialPath = WshShell.SpecialFolders("MyDocuments")
    SpecialPath = ObjFSO.GetFolder(".\Code")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
'    If FSO.FolderExists(SpecialPath & "Ticket Planning Tool") = False Then
'        On Error Resume Next
'        MkDir SpecialPath & "Ticket Planning Tool"
'        On Error GoTo 0
'    End If
    
    'If FSO.FolderExists(SpecialPath & "Ticket Planning Tool") = True Then
    If objFSO.FolderExists(SpecialPath) = True Then
        FolderWithVBAProjectFiles = SpecialPath
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
	
End Function

Function DeleteVBAModulesAndUserForms()
    Dim VBProj 
    Dim VBComp 
    
    Set VBProj = ActiveWorkbook.VBProject
    
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            'Thisworkbook or worksheet module
            'We do nothing
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next 
End Function
