Sub Include(file)
  Wscript.Echo "Include(" + file + ")"
  
  Dim cFS: cFS = new FS
  Dim content: content = cFS.ReadFile(file)
  if content <> "" Then ExecuteGlobal content
End Sub


Public Sub Import(pkg)
  Wscript.Echo "Import(" + Pkg + ")"
  Include "./node_modules/" + pkg + "/index.vbs"
End Sub


public Sub Echo(msg)
    Wscript.Echo msg
End Sub


Public Function jobSrc(file)
  jobSrc = "<script language=""VBScript"" src=""" + file + """/>"
End Function