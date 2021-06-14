Option Explicit

Dim sFeedFileName: sFeedFileName = "tags"
Dim sTextFileExtn: sTextFileExtn = ".txt"
'dim sDocFileExtn: sDocFileExtn = ".docx"
Dim sMachineIdsFilePath: sMachineIdsFilePath = sFolderPath & sFeedFileName & sTextFileExtn
Dim sOutputFileName
Dim sOutputDocFileName

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

'Set Variables for Input file
Dim sFolderPath

'Create File system object
Set PrioFSO = CreateObject("Scripting.FileSystemObject")

'Get active/current directory as folder path
sFolderPath = PrioFSO.GetAbsolutePathName(".") & "\"

'Get feed file for reading
Set oFeedFile = PrioFSO.OpenTextFile(sMachineIdsFilePath, 1, True)

	'Run through each Line in the input feed file
	Do While oFeedFile.AtEndOfStream = False

		'Clean Previous entries
		sBodyMessage = ""

		'Read the machine tag name
		StrComputer = oFeedFile.ReadLine
			
		'Check it's not blank
		If Len(StrComputer) > 0 Then
		
			'Get the actual system tag Ex: MLIW000XXXYYYY
			sCompName = GetProbedID(StrComputer)
			
			If sCompName = False Then
				'report the failure (Note: Always use feedfile Tag Name (StrComputer, not sCompName) to report)
				sStatusMessage = sStatusMessage & StrComputer & vbTab & "Fail-Unable to connect" & vbCrLf
			Else

				'Get Logged In UserName
				'++++++++ additional entry to get username logged into the system ++++++++
				sBodyMessage = GetUserID(sCompName) & vbcrlf
				if sBodyMessage = False then sBodyMessage = "<Tag>UnableToGetSystemInfo" & vbcrlf
				'++++++++ additional entry to get username logged into the system ++++++++


				'Get Installed Applications data
				sBodyMessage = sBodyMessage & GetAddRemove(sCompName)
					
				If sBodyMessage = False Then

					'report the failure
					sStatusMessage = sStatusMessage & StrComputer & vbTab & "Fail-Unable to get Data" & vbCrLf
				Else

					'Create the Output file name
					sOutputFileName = sFolderPath & sCompName & "_" & GetDTFileName() & sTextFileExtn
						  
					'Write the body message to the specified file
					If WriteFile(sBodyMessage, sOutputFileName) Then
						iPassCnt = iPassCnt + 1
						sStatusMessage = sStatusMessage & StrComputer & vbTab & "Pass" & vbCrLf
					Else
						'Report the failure
						sOutputDocFileName = sFolderPath & sCompName & "_" & GetDTFileName() & sDocFileExtn

						If WriteDoc(sBodyMessage, sOutputDocFileName) Then
							 iPassCnt = iPassCnt + 1
							 sStatusMessage = sStatusMessage & StrComputer & vbTab & "Pass" & vbCrLf
						Else
							 'Report the failure
							 sStatusMessage = sStatusMessage & StrComputer & vbTab & "Fail-Unable to Create File" & vbCrLf
						End If
						
					End If
				End If
			End If
			
			iTotalTagCnt = iTotalTagCnt + 1
		End If
	Loop

'Create Results file
sStatusMessage = "Pass:" & iPassCnt & ";" & _
                 "Fail:" & iTotalTagCnt - iPassCnt & _
                 vbCrLf & _
                 sStatusMessage

            
dim sLogFileFullPath: sLogFileFullPath = sFolderPath & sFeedFileName & "_Results_" & GetDTFileName() & sTextFileExtn
Call WriteFile(sStatusMessage, sFolderPath & sFeedFileName & "_Results_" & GetDTFileName() & sTextFileExtn)

Set oFeedFile = Nothing
Set PrioFSO = Nothing


if iTotalTagCnt = 0 then
    msgbox ("No file got created.")
elseIf iTotalTagCnt = iPassCnt Then
    MsgBox ("All files created successfully.")
Else
    MsgBox (iTotalTagCnt - iPassCnt & " file(s) failed to create. Check the Log file in the below path" & vbcrlf & vbcrlf & sLogFileFullPath)
End If



	'----------------------- Connect to Machine and Get Data ---------------------------

	Function GetAddRemove(sComp)
	'*********************************
	'Function credit to Torgeir Bakken
	'*********************************
	On Error Resume Next

		Dim cnt, oReg, sBaseKey, iRC, aSubKeys
		Dim sKey, sValue, sTmp, sVersion, sDateValue, sYr, sMth, sDay
		Const HKLM = &H80000002  'HKEY_LOCAL_MACHINE

		'Refer following portal for additional settings
		'http://msdn.microsoft.com/en-us/library/windows/desktop/aa389292(v=vs.85).aspx

		Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & _
						   sComp & _
						   "/root/default:StdRegProv")

		sBaseKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

		iRC = oReg.EnumKey(HKLM, sBaseKey, aSubKeys)

		If Err.Number <> 0 Then
			GetAddRemove = False
			Exit Function
		End If

		For Each sKey In aSubKeys

			iRC = oReg.GetStringValue(HKLM, sBaseKey & sKey, "DisplayName", sValue)

			If iRC <> 0 Then
				oReg.GetStringValue HKLM, sBaseKey & sKey, "QuietDisplayName", sValue
			End If

			If sValue <> "" Then
				'******** Read Application One after the other ******
				sTmp = sTmp & sValue & vbCrLf
				cnt = cnt + 1
			End If

		Next
	  
		'Sort the records for ease of reading
		sTmp = BubbleSort(sTmp)

		'Append Header to the list
		GetAddRemove = "<Tag>AppCount:" & cnt & ";" & _
					 "TimeStamp:" & Now() & _
					 vbCrLf & _
					 sTmp
	End Function


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
	Function GetUserID(sComp)
	On Error Resume Next

		Dim objWMIService, colItems, objItem

		Set objWMIService = GetObject("winmgmts:\\" & sComp & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

		If Err.Number <> 0 Then
			GetUserID = False
			Exit Function
		End If

		'Get the last Item's SystemName
		For Each objItem In colItems
			GetUserID = "<Tag>" & objItem.name & "." & objItem.Domain & vbcrlf & "<Tag>"

			if len(objItem.UserName) > 2 then  'Blank reply size is 2 bytes
				GetUserID = GetUserID & objItem.UserName
			else
				GetUserID = GetUserID & "UnknownLoggedInUser"
			end if
		Next

	End Function
		
	'-------------------------------------- Write to File ------------------------------
	Function WriteFile(sData, sFileName)
	On Error Resume Next
	  
		Dim fso, OutFile, bWrite
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

		'On Error GoTo 0

		If bWrite Then
			OutFile.WriteLine (sData)
			OutFile.Close
		End If

		If Err Then bWrite = False

		Set fso = Nothing
		Set OutFile = Nothing

		WriteFile = bWrite
	End Function

	'---------------------- Write the text to the document --------------
	Function WriteDoc(sData, sFileName)
	on error resume next
		
		Dim msw,doc

		Set msw = CreateObject("Word.Application")
		msw.Visible = True

		Set doc = msw.documents.Add

		with doc
			.Range.Text = sData
			.saveas(sFileName)
			.close
		end with

		Set doc = Nothing
		Set msw = Nothing
	End Function

'============================== Additional Functions =====================

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

	'--------------------------- Bubble Sort -----------------------
	Function BubbleSort(sTmp)
	on error resume next
	'******************
	'cheapo bubble sort
	'******************

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
