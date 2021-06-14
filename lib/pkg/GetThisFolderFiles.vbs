'Sub GetFileList()
'Dim fso As Scripting.FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")
'Set fso = New Scripting.FileSystemObject
'Dim MainFol As Scripting.Folder

Set MainFol = fso.GetFolder(fso.GetAbsolutePathName("."))

'Dim SubFol As Scripting.Folder
'Dim TxtFile As Scripting.File
'For Each SubFol In MainFol.SubFolders
    'Worksheets(SubFol.Name).Activate
    'ActiveSheet.Name = SubFol.Name
    Index = 2
    For Each TxtFile In MainFol.Files
        If TxtFile.Type = "Text Document" And LCase(Left(TxtFile.Name, 4)) = "mliw" Then
            sTemp = sTemp & Split(TxtFile.Name, "_")(0) & vbCrLf
            Index = Index + 1
        End If
    Next
'Next SubFol
'Dim OutFile As Scripting.TextStream
'msgbox sTemp
Set OutFile = fso.CreateTextFile("ThisFolderFileList.Txt", True, False)
OutFile.WriteLine (BubbleSort(sTemp))
OutFile.Close
Set OutFile = Nothing
Set fso = Nothing
'End Sub


Function BubbleSort(sTmp)
  'cheapo bubble sort
  Dim aTmp, i, j, temp
  aTmp = Split(sTmp, vbCrLf)
  
  For i = UBound(aTmp) - 1 To 0 Step -1
    For j = 0 To i - 1
      If LCase(aTmp(j)) > LCase(aTmp(j + 1)) Then
        temp = aTmp(j + 1)
        aTmp(j + 1) = aTmp(j)
        aTmp(j) = temp
      End If
    Next
  Next
  
  BubbleSort = Join(aTmp, vbCrLf)
End Function
