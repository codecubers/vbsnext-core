Const ForReading = 1

Set args = WScript.Arguments
'directory = args.Item(0) 'thought that workDir would come in here


'Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objTextFile = objFSO.OpenTextFile(&directory&"\filename.txt", ForReading)
'fileName = objTextFile.ReadLine

Dim ToAddress
Dim FromAddress
Dim MessageSubject
Dim MyTime
Dim MessageBody
'Dim MessageAttachment
Dim ol, ns, newMail
MyTime = Now

ToAddress = "pknan@nets.eu"
MessageSubject = "SUCCESS"
MessageBody = "Received Parameter(s): "  & args.Item(0) & ", " & args.Item(1) & ", " & args.Item(2)
'MessageAttachment = &directory&"\"&fileName&"_Log.txt"
Set ol = WScript.CreateObject("Outlook.Application")
Set ns = ol.getNamespace("MAPI")
Set newMail = ol.CreateItem(olMailItem)
newMail.Subject = MessageSubject
newMail.Body = MessageBody & vbCrLf & MyTime
newMail.RecipIents.Add(ToAddress)
'newMail.Attachments.Add(MessageAttachment)
newMail.Send

'objTextFile.Close