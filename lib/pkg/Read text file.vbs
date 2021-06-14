Sub ImportTextFileContent()
    Dim LineofText As String
    Dim rw As Long:  rw = 0
   
    Open “C:\TEXTFILE.txt” For Input As #1
    Do While Not EOF(1)
        Line Input #1, LineofText
        rw = rw + 1
        Cells(rw, 1).Value = LineofText
    Loop
   
    Close #1
End Sub