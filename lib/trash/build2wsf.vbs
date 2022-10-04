
Function jobSrc(file)
  jobSrc = "<script language=""VBScript"" src=""" + file + """/>"
End Function

Include ".\node_modules\VbsJson\index.vbs"
Set json = New VbsJson

dim config, o, script, job, main, pkg
config = ReadFile("package.json")

If config = "" Then
    WScript.Echo "Error!! package.json file is missing"
    WScript.Quit
End if

Set o = json.Decode(config)
If o("name") = "" Then
    WScript.Echo "Error!! Package name is missing in package.json file"
    WScript.Quit
End if
pkg = o("name")

If WScript.Arguments.Count = 0 Then
  WScript.Echo "Warning!! 1st Argument (Job id) missing for build script. Assumed 'index'"
  'WScript.Quit
  main = "index.vbs"
  job = "index"
Else
  main = Trim(WScript.Arguments(0))
  job = replace(replace(Lcase(main), ".vbs", ""), ".", "-")
End If
WScript.Echo "Building job [" + job + "] for script [" + main + "] package [" + pkg + "]"

script = "<package id=""" + pkg + """>" + VBCRLF
script = script + "<job id=""" + job + """ >" + VBCRLF
' script = script + jobSrc("..\node_modules\vbspm\index.vbs") + VBCRLF
script = script + jobSrc("..\node_modules\vbspm\utilities.vbs") + VBCRLF
script = script + jobSrc("..\" + main) + VBCRLF
script = script + "</job>"
script = script + "</package>"

CreateFolder ".\build"
WriteWScript.Echo ".\build\index.wsf", script, true
WScript.Echo "Package " + o("name") + " built successfully."
