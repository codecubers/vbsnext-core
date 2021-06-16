Option Explicit

Dim lCtr, lOpt, lItems
Dim sMsg, sOutput
Dim oApp, oDoc, oTable, oRow
Dim oWorkBk

On Error Resume Next

sMsg = "Enter 0 to send output to a dialog, " & _
       "1 to send output to Word. " 

lOpt = InputBox(sMsg, "Display Error Codes", 1)

Select Case lOpt           ' Output to dialog
   Case 0 
      For lCtr = 1 To 50000
         Err.Raise lCtr
         If Err.Description <> "Unknown runtime error" Then 
				sOutput = sOutput & Err.Number & ": " & Err.Description & vbCrLf
				lItems = lItems + 1
				If lItems Mod 20 = 0 Then
				   MsgBox "Code : Description " & vbCrlf & vbCrLf & sOutput
					lItems = 0
					sOutput = ""
				End If
         End If
         Err.Clear
      Next
   Case 1                  ' Output to Microsoft Word
      Set oApp = WScript.CreateObject("Word.Application")
		Set oDoc = oApp.Documents.Add
      Set oTable = oDoc.Tables.Add(oDoc.Range, 1, 2)
		oTable.Cell(1,1).Range.Text = "Error Code"
		oTable.Cell(1,2).Range.Text = "Description"
      For lCtr = 1 To 50000
         Err.Raise lCtr
         If Err.Description <> "Unknown runtime error" Then 
			   Set oRow = oTable.Rows.Add
				oRow.Cells(1).Range.Text = Err.Number
				oRow.Cells(2).Range.Text = Err.Description
         End If
         Err.Clear
      Next
	 	oApp.Visible = True

	Case Else
      MsgBox "You have entered an invalid value."
End Select
