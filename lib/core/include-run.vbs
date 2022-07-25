' Dim iThread: iThread = 1
' Public Function Thread(i)
'     EchoX "Thread %x", i
'     i = i + 1
'     Thread = i
' End Function
Dim sThreadBase: sThreadBase = baseDir
Public Function Include(file)
  log "Include(" + file + ")"
  if cFS.GetFileExtn(file) = "" Then
    log "File extension missing. Adding .vbs"
    file = file + ".vbs"
  end if
  Dim path
  'path = cFS.GetFilePath(file)
  putil.TempBasePath = sThreadBase
  path = putil.Resolve(file)
  log "File full path: " & path
  'cFS.setDir(cFS.GetFileDir(path))
  sThreadBase = cFS.GetFileDir(path)
  
  If Not arrUtil.contains(IncludedScripts, path) Then
    Redim Preserve IncludedScripts(UBound(IncludedScripts)+1)
    IncludedScripts(UBound(IncludedScripts)) = path
    Dim content: content = cFS.ReadFile(path)
    if content <> "" Then 
      'cFS.WriteFile "build\bundle.vbs", content, false
      'EchoX "File: %x", file
      'EchoX "Thread ---> %x", iThread
      'content = "iThread = Thread(iThread)" & VBCRLF & content
      'EchoX "Content: %x", content
      dim lines
      lines = split(content, vbCrLf)
      Dim includeS
      for i = 0 to ubound(lines)
        WScript.Echo "Searching in line:" & lines(i)
        if InStr(lines(i), "Include(") > 0 Or InStr(lines(i), "Include """) > 0 Or InStr(lines(i), "Import(") > 0 or InStr(lines(i), "Import """) > 0 Then
          includeS = includeS & lines(i) & vbCrLf
        end if
      next
      WScript.Echo "Lines to execute:" & includeS
      if includeS <> "" Then
          ExecuteGlobal includeS
      End If
    Else
      log "File content is empty. Not loaded."
    End If
  Else
    log "File: " & path & " already loaded."
  End If
  Include = Include
End Function