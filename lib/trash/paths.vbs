
''''Way 1

Currentdir=Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
Wscript.Echo "Calling script:" & WScript.ScriptName & " at: " & currentdir

''''Way 2

With CreateObject("WScript.Shell")
CurrentPath=.CurrentDirectory
Wscript.Echo  "Calling from: " & CurrentPath
End With


' ''''Way 3

' With WSH
' CD=Replace(.ScriptFullName,.ScriptName,"")
' Wscript.Echo CD
' End With

' Dim fso: set fso = CreateObject("Scripting.FileSystemObject")
' Dim CurrentDirectory
' CurrentDirectory = fso.GetAbsolutePathName(".")
' Directory = CurrentDirectory & "\attribute.exe"
' Wscript.Echo Directory
' Directory = fso.BuildPath(CurrentDirectory, "attribute.exe")
' Wscript.Echo Directory

' scriptdir = replace(WScript.ScriptFullName,WScript.ScriptName,"")
' Wscript.Echo scriptdir
' Wscript.Echo WSH = Wscript




strScriptHost = LCase(Wscript.FullName)

If Right(strScriptHost, 11) = "wscript.exe" Then

    Wscript.Echo "This script is running under WScript."

Else

    Wscript.Echo "This script is running under CScript."

End If

Wscript.Echo "Build number: " & _
    ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion

	Wscript.Echo "Script name: " & Wscript.ScriptName

Wscript.Echo "Script path: " & Wscript.ScriptFullName


Set objShell = CreateObject( "WScript.Shell" )
resourceLocation=objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")

currentdir=Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))