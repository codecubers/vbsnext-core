log "Execution Started for file"

Dim file
file = WScript.Arguments.Named("file")
If file = "" Then
    log "Script file not provided as a named argument [/file:]"
    if Wscript.Arguments.count > 0 then
      file = Wscript.Arguments(0) 
      if file = "" Then
        log "No file argument provided."
        Wscript.Quit
      End If
    else 
      file = "index.vbs"
    end if
End If
' TODO: Assess all possible combinations a user can send in command line
file = baseDir & "\" & file

if cFS.GetFileExtn(file) = "" Then
  log "File extension missing. Adding .vbs"
  file = file + ".vbs"
end if

log "Main Script: " & file
buildBundleFile = buildDir & "\" & cFS.GetBaseName(file) &  "-bundle.vbs"
log "buildBundleFile: " & buildBundleFile

