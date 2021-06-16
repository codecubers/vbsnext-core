Dim baseDir
Dim cFS
set cFS = new FSO


public Sub Echo(msg)
  Wscript.Echo msg
End Sub
Function log(msg)
  cFS.WriteFile "build.log", msg, false
End Function
log "================================= Call ================================="

Sub Include(file)
  log "Include(" + file + ")"
  
  Dim content: content = cFS.ReadFile(file)
  if content <> "" Then 
    ' Dim pkg
    ' pkg = Replace(file, "\node_modules\", "")
    ' pkg = Replace(pkg, "\index.vbs", "")
    ' cFS.WriteFile "build\imported\" & pkg & ".vbs", content, true
    ExecuteGlobal content
  End If
End Sub

Public Sub Import(pkg)
  log "Import(" + Pkg + ")"
  Include baseDir & "\node_modules\" + pkg + "\index.vbs"
End Sub


With CreateObject("WScript.Shell")
  baseDir=.CurrentDirectory
  'Wscript.Echo  "Base path: " & baseDir
End With
log "Base path: " & baseDir
cFS.setDir(baseDir)