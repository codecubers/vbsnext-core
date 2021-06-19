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
