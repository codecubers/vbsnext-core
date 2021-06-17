Dim baseDir
Dim cFS
Redim IncludedScripts(-1)
Dim arrUtil
set cFS = new FSO
set arrUtil = new ArrayUtil


public Sub Echo(msg)
  Wscript.Echo msg
End Sub

Function log(msg)
  cFS.WriteFile "build.log", msg, false
End Function

With CreateObject("WScript.Shell")
  baseDir=.CurrentDirectory
End With
log "Base path: " & baseDir
cFS.setDir(baseDir)

log "================================= Call ================================="

Public Sub Import(pkg)
  log "Import(" + Pkg + ")"
  Include baseDir & "\node_modules\" + pkg + "\index.vbs"
End Sub