Option Explicit

Dim goFS  : Set goFS = CreateObject("Scripting.FileSystemObject")
Dim goWS  : Set goWS = CreateObject("WScript.Shell")
Dim csDir : csDir = "C:\Users\nandapr\Desktop\ExcelVBA2"

WScript.Quit demoSF()

Function demoSF()
  demoSF = 0
  Dim aDSOrd : aDSOrd = getDSOrd(csDir, "%comspec% /c dir /A:-D /B /O:D /T:C """ & csDir & """")
  Dim oFile
  For Each oFile In aDSOrd
      WScript.Echo oFile.DateCreated, oFile.Name
  Next
End Function ' demoSF

Function getDSOrd(sDir, sCmd)
  Dim dicTmp : Set dicTmp = CreateObject("Scripting.Dictionary")
  Dim oExec  : Set oExec  = goWS.Exec(sCmd)
  Do Until oExec.Stdout.AtEndOfStream
     dicTmp(goFS.GetFile(goFS.BuildPath(sDir, oExec.Stdout.ReadLine()))) = Empty
  Loop
  If Not oExec.Stderr.AtEndOfStream Then
     WScript.Echo "Error:", oExec.Stderr.ReadAll()
  End If
  getDSOrd = dicTmp.Keys()
End Function