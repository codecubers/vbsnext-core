Dim sURL:sURL = "https://webmail.syntel.in/exchange"

dim appIE,ElementCol

set appIE = createobject("InternetExplorer.Application")

With appIE
    .Navigate sURL
    .Visible = TRUE
End With

Do 

Loop While appIE.Busy = True

Set ElementCol = appIE.Document.getElementsByTagName("Input")

        For Each Link In ElementCol
            If Link.Name = "username" Then
               Link.value = "PN28811"
            End If
        Next 	

        For Each Link In ElementCol
            If Link.Name = "password" Then
               Link.value = "syntel234%"
            End If
        Next 	

        For Each Link In ElementCol
            If Link.Type = "submit" Then
               Link.Click
            End If
        Next 

Set ElementCol = Nothing
set appIe = nothing