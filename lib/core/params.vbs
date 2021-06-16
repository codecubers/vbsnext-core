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

if Lcase(cFS.GetExtn(file)) = "vbs" Then
  log "File extension is: .vbs"
Else
  log "File extension missing. Adding .vbs"
  file = file + ".vbs"
end if

log "File: " & file


Dim script
script = cFS.ReadFile(file)
if script = "" Then
  log "No file supplied or is empty."
  Wscript.Quit
End if



'=========================== 
Execute script
'=========================== 