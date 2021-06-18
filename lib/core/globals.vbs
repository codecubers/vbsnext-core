Dim vbspmDir
Dim baseDir
Dim cFS
Redim IncludedScripts(-1)
Dim arrUtil
Dim buildDir
Dim createBundle: createBundle = false
Dim buildBundleFile: buildBundleFile = ""
Dim oConsole 

set cFS = new FSO
set arrUtil = new ArrayUtil
public Sub Echo(msg)
  Wscript.Echo msg
End Sub

Function log(msg)
  cFS.WriteFile "build.log", msg, false
End Function

vbspmDir = cFS.GetFileDir(WScript.ScriptFullName)
log "VBSPM Directory: " & vbspmDir
With CreateObject("WScript.Shell")
  baseDir=.CurrentDirectory
End With
log "Base path: " & baseDir
cFS.setDir(baseDir)
buildDir = baseDir & "\build"
If cFS.CreateFolder(buildDir) Then
  createBundle = true
Else
  WScript.Echo "Unable to create build directory at [" & buildDir & "]. Script will not be bundled. Please try again."
End If

log "================================= Call ================================="

Public Sub Import(pkg)
  log "Import(" + Pkg + ")"
  Include baseDir & "\node_modules\" + pkg + "\index.vbs"
End Sub


set oConsole = new Console
PUblic Sub printf(str, args)
    'TODO: If use use %s, %d, %f format the values according to it.
    str = Replace(str, "%s", "%x")
    str = Replace(str, "%i", "%x")
    str = Replace(str, "%f", "%x")
    str = Replace(str, "%d", "%x")
    WScript.Echo oConsole.fmt(str, args)
End Sub

Public Sub debugf(str, args)
    if (debug) Then printf str, args
End Sub
