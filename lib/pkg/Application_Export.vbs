'TODO: Error handling for each sheet and validation
Dim Application
Dim ActiveWorkbook
Const vbext_ct_Document = 100, vbext_ct_ClassModule = 2, vbext_ct_MSForm = 3, vbext_ct_StdModule = 1, vbext_ct_ActiveXDesigner = 11
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = 2, TristateTrue= 1, TristateFalse= 0

Dim bExport 
Dim wkbSource 
Dim szSourceWorkbook 
Dim szExportPath 
Dim szFileName 
Dim cmpComponent 

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
	Set objFSO = CreateObject("scripting.filesystemobject")
    sPath = FolderWithVBAProjectFiles
    If sPath = "Error" Then
        MsgBox "Export Folder not exist"
        WScript.Quit
    End If
    
    On Error Resume Next
        Kill sPath & "\*.cls"
		Kill sPath & "\*.frm"
		Kill sPath & "\*.bas"
		Kill sPath & "\*.frx"
    On Error GoTo 0

	
	
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
	sOldVersion = Replace(Trim(ActiveWorkbook.worksheets("App").Range("N1").value),vbCr,"")
	
	
    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
		Activeworkbook.close
		Application.Quit
    WScript.Quit
    End If
    
	Do 
		sNewVersion = InputBox("Current Version is: " & sOldVersion & ". Please input new version (Higher).","Please input new version number",sOldVersion)
	Loop while sNewVersion <= sOldVersion 
	
    szExportPath = sPath & "\"
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
    Next 

	If objFSO.FileExists(szExportPath & "Version.txt") Then
		set oVersionFile = objFSO.OpenTextFile(szExportPath & "Version.txt", ForWriting, True)
		oVersionFile.WriteLine sNewVersion
		oVersionFile.close
	End if
	
ActiveWorkbook.worksheets("App").Range("N1").value = sNewVersion
ActiveWorkbook.Save

MsgBox "Code exported successfully." & vbcrlf & _
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
