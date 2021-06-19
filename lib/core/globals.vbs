log "================================= Call ================================="

log "Base path: " & baseDir

Public Sub Import(pkg)
  log "Import(" + Pkg + ")"
  Include baseDir & "\node_modules\" + pkg + "\index.vbs"
End Sub