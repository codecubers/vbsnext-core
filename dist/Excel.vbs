Class Excel

    Private Property Get vbext_ct_Document
    vbext_ct_Document = 100
    End Property
    Private Property Get vbext_ct_ClassModule
    vbext_ct_ClassModule = 2
    End Property
    Private Property Get vbext_ct_MSForm
    vbext_ct_MSForm = 3
    End Property
    Private Property Get vbext_ct_StdModule
    vbext_ct_StdModule = 1
    End Property
    Private Property Get vbext_ct_ActiveXDesigner
    vbext_ct_ActiveXDesigner = 11
    End Property
    Private Property Get excel_workbook_protected_level_protected
    excel_workbook_protected_level_protected = 1
    End Property
    Private Property Get ForReading
    ForReading = 1
    End Property
    Private Property Get ForWriting
    ForWriting = 2
    End Property
    Private Property Get ForAppending
    ForAppending = 3
    End Property
    Private Property Get TristateUseDefault
    TristateUseDefault = 2
    End Property
    Private Property Get TristateTrue
    TristateTrue = 1
    End Property
    Private Property Get TristateFalse
    TristateFalse = 0
    End Property

    Public Property Get GetActiveWorkbook
        Set GetActiveWorkbook = ActiveWorkbook
    End Property

    Private Application
    Private ActiveWorkbook
    Private wkbSource
    Private objFSO

    Private Sub Class_Initialize()
        Set objFSO = CreateObject("scripting.filesystemobject")
        set Application = createobject("Excel.Application").Application
        if Application is nothing then
            Echo "Unable to create Excel Application object."
            Err.Clear
            Err.Raise 50001, "Error in Excel Class", "Unable to create Excel application object."
            Class_Terminate
        end if
        SetVisibility False
        ShowAlerts False
    End Sub

    Public Sub OpenWorkBook(path)
        On Error Resume Next
        path = objFSO.GetFile(path).path
        EchoX "Opening Excel Workbook at path: %x", path
        set ActiveWorkbook = Application.workbooks.Open(path)
        On Error Goto 0
        if Not IsObject(ActiveWorkbook) then
            EchoX "Unable to Open Excel Workbook at path %x.", path
            Err.Clear
            Err.Raise 50002, "Error in Excel Class", "Unable to open Excel Workbook at path " & path
        end if
        ''' NOTE: This workbook must be open in Excel.
        Set wkbSource = Application.Workbooks(ActiveWorkbook.Name)
        EchoX "Workbook %x opened successfully.", wkbSource.Name
    End Sub
    Public SUb CloseWorkBook
        On Error Resume Next
        ActiveWorkbook.Close
        On Error Goto 0
    End Sub

    Public Function isProtected
        On Error Resume Next
        isProtected = False
        isProtected = (wkbSource.VBProject.Protection = excel_workbook_protected_level_protected)
        On Error Goto 0
    End Function

    Public Sub SetVisibility(flag)
        Application.Visible = (flag or LCase(flag) = "true")
    End Sub
    Public Sub ShowAlerts(flag)
        Application.DisplayAlerts = (flag Or Lcase(flag) = "true")
    End Sub

    Public Sub ExportVBComponents(destination)
        Dim cmpComponent, bExport, szFileName

        If isProtected Then
            Echo "The workbook is protected. Cannot export VB Components."
            Exit Sub
        End If

        If Not objFSO.FolderExists(destination) Then
            EchoX "Destination folder %x does not exists. Please create it and retry.", destination
            Exit Sub
        End If
        destination = ObjFSO.GetFolder(".\Code")

        'TODO: Move objFSO code to it's own class
        On Error Resume Next
        EchoX "Deleting previously exported VBA Modules in direcotry %x", destination
        objFSO.DeleteFile objFSO.BuildPath(destination, "*.cls"), True
        objFSO.DeleteFile  objFSO.BuildPath(destination, "*.frm"), True
        objFSO.DeleteFile  objFSO.BuildPath(destination, "*.bas"), True
        objFSO.DeleteFile  objFSO.BuildPath(destination, "*.frx"), True
        On Error GoTo 0
        
        EchoX "Exporting VBComponents to folder: %x", destination
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
                EchoX "Exporting Module %x to %x", Array(szFileName, objFSO.BuildPath(destination, szFileName))
                On Error Resume Next
                cmpComponent.Export objFSO.BuildPath(destination, szFileName)
                Echo Err.Description
                On Error Goto 0
                
            ''' remove it from the project if you want
            '''wkbSource.VBProject.VBComponents.Remove cmpComponent
            
            End If
        Next 
    End Sub

    Public Sub ImportVBAComponents(source)
        Dim cmpComponents, objFile
        
        If Not objFSO.FolderExists(source) Then
            EchoX "Unable to get source directory at: %x", source
            Exit Sub
        End If

        If isProtected Then
            Echo "The workbook is protected. Cannot export VB Components."
            Exit Sub
        End If

        Set cmpComponents = wkbSource.VBProject.VBComponents

        ''' Import all the code modules in the specified path
        ''' to the wkbSource.
        DeleteVBAComponents False

        For Each objFile In objFSO.GetFolder(source).Files
            If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
                (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
                (objFSO.GetExtensionName(objFile.Name) = "bas") Then
                cmpComponents.Import objFile.Path
            End If
        Next 

        wkbSource.save
    End Sub

    Public Sub DeleteVBAComponents(save)
        Dim VBComponents, VBComp 

        If isProtected Then
            Echo "The workbook is protected. Cannot delete VB Components."
            Exit Sub
        End If

        Echo "About to delete the VBA components of the workbook"
        Set VBComponents = wkbSource.VBProject.VBComponents
        For Each VBComp In VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBComponents.Remove VBComp
            End If
        Next 
        if(save) Then wkbSource.save
    End Sub

    Private Sub Class_Terminate()
        Echo "Excel Class being terminated."
        On Error Resume Next
        ShowAlerts
        ActiveWorkbook.close
        set ActiveWorkbook = nothing
        Application.quit
        set Application = nothing
        On Error Goto 0
    End Sub

End Class ' Excel