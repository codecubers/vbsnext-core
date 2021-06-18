


' ================================== Job: vbspm-build ================================== 

' ================= src : lib/core/init.vbs ================= 
Option Explicit

Dim debug: debug = (WScript.Arguments.Named("debug") = "true")
if (debug) Then WScript.Echo "Debug is enabled"
' ================= src : lib/core/include-build.vbs ================= 

Public Sub Include(file)
  ' DO NOT REMOVE THIS Sub Routine
End Sub
Public Sub Import(file)
  ' DO NOT REMOVE THIS Sub Routine
End Sub


'================= File: C:\Users\nanda\git\xps.local.npm\vbspm\bin\test-cls.vbs =================
Class BUILDTEST
    Public default Property Get Status
            Status = "Successfully.."
    End Property
End Class


'================= File: C:\Users\nanda\git\xps.local.npm\vbspm\bin\test.vbs =================
Include "bin\test-cls.vbs"
set test = new BUILDTEST
Wscript.Echo "Build completed " & test & "."
