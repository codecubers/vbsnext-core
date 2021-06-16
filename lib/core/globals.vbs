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

log "================================= Call ================================="

Sub Include(file)
  log "Include(" + file + ")"
  if Lcase(cFS.GetExtn(file)) = "" Then
    log "File extension missing. Adding .vbs"
    file = file + ".vbs"
  end if
  Dim path: path = cFS.GetFilePath(file)
  log "File full path: " & path
  
  ' Dim pkg
  ' pkg = Replace(file, "\node_modules\", "")
  ' pkg = Replace(pkg, "\index.vbs", "")
  ' cFS.WriteFile "build\imported\" & pkg & ".vbs", content, true
  If Not arrUtil.contains(IncludedScripts, path) Then
    Redim Preserve IncludedScripts(UBound(IncludedScripts)+1)
    IncludedScripts(UBound(IncludedScripts)) = path
    Dim content: content = cFS.ReadFile(file)
    if content <> "" Then 
      ExecuteGlobal content
    Else
      log "File content is empty. Not loaded."
    End If
  Else
    log "File: " & path & " already loaded."
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