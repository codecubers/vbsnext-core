
Public Sub Include(file)
  log "Include(" + file + ")"
  if Lcase(cFS.GetExtn(file)) = "" Then
    log "File extension missing. Adding .vbs"
    file = file + ".vbs"
  end if
  Dim path: path = cFS.GetFilePath(file)
  log "File full path: " & path
End Sub