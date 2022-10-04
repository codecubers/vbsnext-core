WScript.Echo "================================= Call ================================="

WScript.Echo "Base path: " & baseDir

Public Sub Import(pkg)
  WScript.Echo "Import(" + Pkg + ")"
  Include baseDir & "\node_modules\" + pkg + "\index.vbs"
End Sub