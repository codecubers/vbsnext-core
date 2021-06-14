Option Explicit

'Sub MAIN()
'On Error Resume Next
'Local variables
Dim sStatusMessage
Dim StrComputer
Dim sCompName
Dim sBodyMessage
Dim iTotalTagCnt: iTotalTagCnt = 0
Dim iPassCnt: iPassCnt = 0

'Create FileSystemObject 
Dim PrioFSO
Dim oFeedFile
Set PrioFSO = CreateObject("Scripting.FileSystemObject")

'Set Variables for Input file
Dim sFolderPath
'sFolderPath = "U:\01.Projects\000-000 Resource Upgradation\Win7\"

sFolderPath = PrioFSO.GetAbsolutePathName(".") & "\"
'msgbox("checkpoint0" & sFolderPath)
Dim sFeedFileName: sFeedFileName = "Machine_IDs"
Dim sTextFileExtn: sTextFileExtn = ".txt"
Dim sMachineIdsFilePath: sMachineIdsFilePath = sFolderPath & sFeedFileName & sTextFileExtn
Dim sOutputFileName
'msgbox("checkpoint1")
'Get feed file for reading
Set oFeedFile = PrioFSO.OpenTextFile(sMachineIdsFilePath, 1, True)
'msgbox("checkpoint2")
'Run through each Line in the input feed file
Do While oFeedFile.AtEndOfStream = False

'Clean Previous entries
sBodyMessage = ""
'sStatusMessage = ""

    'Read the machine tag name
    StrComputer = oFeedFile.ReadLine
    'Check it's not blank

    If Len(StrComputer) > 0 Then
'msgbox("checkpoint3")
        'Get the actual system tag Ex: MLIW000XXXYYYY
        sCompName = GetProbedID(StrComputer)
        
        If sCompName = False Then
'msgbox("checkpoint4")
            'report the failure (Note: Always use feedfile Tag Name (StrComputer, not sCompName) to report)
            sStatusMessage = sStatusMessage & StrComputer & vbTab & "Fail-Unable to connect" & vbCrLf
			iPassCnt = iPassCnt + 1
        Else
			sStatusMessage = sStatusMessage & StrComputer & vbTab & "Pass-Able to connect" & vbCrLf
		end if
'msgbox("checkpoint5")
        
        iTotalTagCnt = iTotalTagCnt + 1
    End If
Loop

'msgbox("checkpoint8")

'Create Results file
sStatusMessage = "Pass:" & iPassCnt & ";" & _
                 "Fail:" & iTotalTagCnt - iPassCnt & _
                 vbCrLf & _
                 sStatusMessage

'msgbox("checkpoint9:" & sStatusMessage)                 
dim sLogFileFullPath: sLogFileFullPath = sFolderPath & sFeedFileName & "_Results_" & GetDTFileName() & sTextFileExtn
Call WriteFile(sStatusMessage, sFolderPath & sFeedFileName & "_Results_" & GetDTFileName() & sTextFileExtn)

Set oFeedFile = Nothing
Set PrioFSO = Nothing

'msgbox("checkpoint10")
if iTotalTagCnt = 0 then
    msgbox ("None of the Tag(s) responded.")
elseIf iTotalTagCnt = iPassCnt Then
    MsgBox ("All Tags responded successfully.")
Else
    MsgBox (iTotalTagCnt - iPassCnt & " of " &  iTotalTagCnt " Tag(s) didn't respond. Check the Log file in the below path" & vbcrlf & vbcrlf & sLogFileFullPath)
End If

'End Sub





'----------------------- Connect to Machine and Get Data ---------------------------

'-------------------------------------- GET PROBED ID ------------------------------
Function GetProbedID(sComp)
On Error Resume Next

  Dim objWMIService, colItems, objItem
  
  Set objWMIService = GetObject("winmgmts:\\" & sComp & "\root\cimv2")
  
  Set colItems = objWMIService.ExecQuery("Select SystemName from " & _
                                         "Win32_NetworkAdapter", , 48)
  If Err.Number <> 0 Then
    GetProbedID = False
    Exit Function
  End If
  
  'Get the last Item's SystemName
  For Each objItem In colItems
    GetProbedID = objItem.SystemName
  Next
  
End Function
'---------------------------------- GET Additional System Info ---------------------
'-------------------------------------- Write to File ------------------------------
Function WriteFile(sData, sFileName)
  Dim fso, OutFile, bWrite

On Error Resume Next
  bWrite = True
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set OutFile = fso.OpenTextFile(sFileName, 2, True)
  
  'Possibly need a prompt to close the file and one recursion attempt.
  If Err = 70 Then
  
    WScript.Echo "Could not write to file " & sFileName & ", results " & _
                 "not saved." & vbCrLf & vbCrLf & "This is probably " & _
                 "because the file is already open."
    bWrite = False
  ElseIf Err Then
  
    'WScript.Echo Err & vbCrLf & Err.Description
    bWrite = False
  End If
  
'  On Error GoTo 0
  
  If bWrite Then
    OutFile.WriteLine (sData)
    OutFile.Close
  End If
  
 If Err Then bWrite = False

  Set fso = Nothing
  Set OutFile = Nothing
  
  WriteFile = bWrite
End Function


'-------------------------------------- Additional Functions -----------------------
'--------------------------- GET Date&Time FileName ------------
Function GetDTFileName()
  Dim sNow, sMth, sDay, sYr, sHr, sMin, sSec
  sNow = Now
  sMth = Right("0" & Month(sNow), 2)
  sDay = Right("0" & Day(sNow), 2)
  sYr = Right("00" & Year(sNow), 4)
  sHr = Right("0" & Hour(sNow), 2)
  sMin = Right("0" & Minute(sNow), 2)
  sSec = Right("0" & Second(sNow), 2)
  GetDTFileName = sMth & sDay & sYr & "_" & sHr & sMin & sSec
End Function