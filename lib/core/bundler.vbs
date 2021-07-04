Sub BundleScript(file, overwrite)
    Dim isOverwrite: isOverwrite = (overwrite = true)
    Dim content: content = cFS.ReadFile(file)
    if createBundle Then
        cFS.WriteFile buildBundleFile, content, isOverwrite
    End If
End Sub

Sub BundleScriptStr(content, overwrite)
    Dim isOverwrite: isOverwrite = (overwrite = true)
    if createBundle Then
        cFS.WriteFile buildBundleFile, content, isOverwrite
    End If
End Sub


' Just before start writing Include/Import file contents to the builder,
' Write the vbspm.vbs file contents
BundleScript vbsnextDir & "\vbsnext-build.vbs", true

'===========================
On Error Resume Next
Include file
On Error Goto 0
'===========================

' Wscript.Echo arrUtil.toString(IncludedScripts)
Dim i, core
for i = UBound(IncludedScripts) to 0 step -1
    core = cFS.ReadFile(IncludedScripts(i))
    core = Replace(core, "Option Explicit", "")
    core = vbCrLf & vbCrLf & "'================= File: " & IncludedScripts(i) & " =================" & vbCrLf & core
    BundleScriptStr core, false
next