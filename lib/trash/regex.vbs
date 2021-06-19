' Test if string begins with a string?
' 
' The best methods are already given but why not look at a couple of other methods for fun? 
' Warning: these are more expensive methods but do serve in other circumstances.
' The expensive regex method and the css attribute selector with starts with ^ operator

' Author: QHarr
' Source: https://stackoverflow.com/a/65548353/1751166

Option Explicit

Public Function StartWithSubString(ByVal substring As String, ByVal testString As String) As Boolean
    'required reference Microsoft VBScript Regular Expressions
    Dim re As VBScript_RegExp_55.RegExp
    Set re = New VBScript_RegExp_55.RegExp

    re.Pattern = "^" & substring

    StartWithSubString = re.test(testString)

End Function

' Css attribute selector with starts with operator
' Author: QHarr
' Source: https://stackoverflow.com/a/65548353/1751166

' Public Function StartWithSubString(ByVal substring As String, ByVal testString As String) As Boolean
'     'required reference Microsoft HTML Object Library
'     Dim html As MSHTML.HTMLDocument
'     Set html = New MSHTML.HTMLDocument

'     html.body.innerHTML = "<div test=""" & testString & """></div>"

'     StartWithSubString = html.querySelectorAll("[test^=" & substring & "]").Length > 0

' End Function


Debug.Print StartWithSubString("ab", "abc,d")

' ' Judging by the declaration and description of the startsWith Java function, 
' ' the "most straight forward way" to implement it in VBA would either be with Left:
' ' Author: Blackhawk
' ' Source: https://stackoverflow.com/a/20805609/1751166

' Public Function startsWith(str As String, prefix As String) As Boolean
'     startsWith = Left(str, Len(prefix)) = prefix
' End Function

' Or, if you want to have the offset parameter available, with Mid:
' Author: Blackhawk
' Source: https://stackoverflow.com/a/20805609/1751166

Public Function startsWith(str As String, prefix As String, Optional toffset As Integer = 0) As Boolean
    startsWith = Mid(str, toffset + 1, Len(prefix)) = prefix
End Function




' ==================================================================================================
' Author: armstrhb
' Source: https://stackoverflow.com/a/20802871/1751166

' You can use the InStr build-in function to test if a String contains a substring. 
' InStr will either return the index of the first match, or 0. 
' So you can test if a String begins with a substring by doing the following:
' If InStr returns 1, then the String ("Hello World"), begins with the substring ("Hello W").
If InStr(1, "Hello World", "Hello W") = 1 Then
    MsgBox "Yep, this string begins with Hello W!"
End If

' You can also use the like comparison operator along with some basic pattern matching:
' In this, we use an asterisk (*) to test if the String begins with our substring.
If "Hello World" Like "Hello W*" Then
    MsgBox "Yep, this string begins with Hello W!"
End If
' ==================================================================================================