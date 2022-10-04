
Public Sub Include(file)
  ' DO NOT REMOVE THIS Sub Routine
End Sub
Public Sub Import(file)
  ' DO NOT REMOVE THIS Sub Routine
End Sub

Dim sThreadBase: sThreadBase = baseDir
Public Function Import(file)
  EchoD "Importing... (" + file + ")"
  if cFS.GetFileExtn(file) = "" Then
    WScript.Echo "File extension missing. Skipping"
    'file = file + ".vbs"
  end if
  Dim path

  putil.TempBasePath = sThreadBase
  path = putil.Resolve(file)
  EchoD "Resolved to: " & path

  sThreadBase = cFS.GetFileDir(path)
  EchoD "Current base path is: " & sThreadBase

  If arrUtil.contains(ImportedScripts, Lcase(path)) Then
    WScript.Echo "Skipping as already imported!!"
  Else
    Redim Preserve ImportedScripts(UBound(ImportedScripts)+1)
    ImportedScripts(UBound(ImportedScripts)) = Lcase(path)
    Dim content: content = cFS.ReadFile(path)
    if content <> "" Then

      dim lines
      Dim sThisLine
      Dim i
      'lines = split(join(split(content, ":"), vbCrLf), vbCrLf)
      lines = split(content, vbCrLf)
      Dim includeS
      for i = 0 to ubound(lines)
        sThisLine = Trim(lines(i))
        if Left(Lcase(sThisLine), 8) = "'import(" Then
          sThisLine = right(sThisLine, len(sThisLine)-1)
          EchoD "--------------> Found:" & sThisLine
          includeS = includeS & sThisLine & vbCrLf
        end if
      next
      
      if includeS <> "" Then
          EchoD "Scanning:" & vbcrlf & includeS
          ExecuteGlobal includeS
      End If
    Else
      WScript.Echo "File content is empty. Not loaded."
    End If
  End If
End Function




Function ResolveImports(entryPoint)
    pUtil.AddBasepath "src\modules"
    pUtil.AddBasepath "src\classes"
    pUtil.AddBasepath "src\processors"

    Redim ImportedScripts(-1)
    Import(entryPoint)
    Wscript.Echo arrUtil.toString(ImportedScripts)
    ResolveImports = ImportedScripts
End Function
