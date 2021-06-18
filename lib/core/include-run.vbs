
Public Sub Include(file)
  log "Include(" + file + ")"
  if cFS.GetFileExtn(file) = "" Then
    log "File extension missing. Adding .vbs"
    file = file + ".vbs"
  end if
  Dim path: path = cFS.GetFilePath(file)
  log "File full path: " & path
  cFS.setDir(cFS.GetFileDir(file))
  
  If Not arrUtil.contains(IncludedScripts, path) Then
    Redim Preserve IncludedScripts(UBound(IncludedScripts)+1)
    IncludedScripts(UBound(IncludedScripts)) = path
    Dim content: content = cFS.ReadFile(file)
    if content <> "" Then 
      'cFS.WriteFile "build\bundle.vbs", content, false
      ExecuteGlobal content
    Else
      log "File content is empty. Not loaded."
    End If
  Else
    log "File: " & path & " already loaded."
  End If
End Sub