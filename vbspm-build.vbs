


' ================================== Job: vbspm-build ================================== 

' ================= src : lib/core/init.vbs ================= 
Option Explicit

' Judging by the declaration and description of the startsWith Java function, 
' the "most straight forward way" to implement it in VBA would either be with Left:
' Author: Blackhawk
' Source: https://stackoverflow.com/a/20805609/1751166

Public Function startsWith(str, prefix)
    startsWith = Left(str, Len(prefix)) = prefix
End Function

Public Function endsWith(str, suffix)
    endsWith = Right(str, Len(suffix)) = suffix
End Function

' ================= src : lib/core/Console/Console.vbs ================= 
Class Console

    ' Author: Uwe Keim
    ' License: The Code Project Open License (CPOL)
    ' https://www.codeproject.com/Articles/250/printf-like-Format-Function-in-VBScript
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' works like the printf-function in C.
    ' takes a string with format characters and an array
    ' to expand.
    '
    ' the format characters are always "%x", independ of the
    ' type.
    '
    ' usage example:
    '	dim str
    '	str = fmt( "hello, Mr. %x, today's date is %x.", Array("Miller",Date) )
    '	response.Write str
    public function fmt( str, args )
        dim res		' the result string.
        res = ""

        dim pos		' the current position in the args array.
        pos = 0

        dim i
        for i = 1 to Len(str)
            ' found a fmt char.
            if Mid(str,i,1)="%" then
                if i<Len(str) then
                    ' normal percent.
                    if Mid(str,i+1,1)="%" then
                        res = res & "%"
                        i = i + 1

                    ' expand from array.
                    elseif Mid(str,i+1,1)="x" then
                        res = res & CStr(args(pos))
                        pos = pos+1
                        i = i + 1
                    end if
                end if

            ' found a normal char.
            else
                res = res & Mid(str,i,1)
            end if
        next

        fmt = res
    end function

End Class
' ================= inline ================= 

Dim oConsole
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

Public Sub EchoX(str, args)
If Not IsNull(args) Then
If IsArray(args) Then
'WScript.Echo str & " with args " & join(args, ",")
WScript.Echo oConsole.fmt(str, args)
Else
'WScript.Echo str & " with arg " & args
WScript.Echo oConsole.fmt(str, Array(args))
End if
Else
WScript.Echo str
End If
End Sub

Public Sub Echo(str)
EchoX str, NULL
End Sub

Public Sub EchoDX(str, args)
if (debug) Then EchoX str, args
End Sub

Public Sub EchoD(str)
EchoDX str, NULL
End Sub

' ================= src : lib/core/include-build.vbs ================= 

Public Sub Include(file)
  ' DO NOT REMOVE THIS Sub Routine
End Sub
Public Sub Import(file)
  ' DO NOT REMOVE THIS Sub Routine
End Sub