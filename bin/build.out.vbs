


' ================================== Job: vbspm-build ================================== 

' ================= src : lib/core/FSO/FSO.vbs ================= 
Class FSO
	Private objFSO
	
	Private Sub Class_Initialize
		Set objFSO = CreateObject("Scripting.FileSystemObject")
	End Sub
	
	Public Function GetFSO
		Set GetFSO = objFSO
	End Function

    ' ===================== Sub Routines =====================

	Public Sub CreateFolder(fol)
		If Not objFSO.FolderExists(fol) Then
			objFSO.CreateFolder(fol)
		End If
	End Sub
	
	Public Sub WriteLog(strFileName, strMessage, overwrite)
		Const ForReading = 1
		Const ForWriting = 2
		Const ForAppending = 8
		Dim mode
		
        mode = ForWriting
		If Not overwrite Then
			mode = ForAppending
		End If
		
		If objFSO.FileExists(strFileName) Then
			Set oFile = objFSO.OpenTextFile(strFileName, mode)
		Else
			Set oFile = objFSO.CreateTextFile(strFileName)
		End If
		oFile.WriteLine strMessage
		
		oFile.Close
		
		Set oFile = Nothing
	End Sub 

	' ===================== Function Routines =====================

	Public Function GetFileDir(ByVal file)
		Set objFile = objFSO.GetFile(file)
		GetFileDir = objFSO.GetParentFolderName(objFile) 
	End Function
	
	Public Function ReadFile(file)
		If Not FileExists(file) Then 
			Wscript.Echo "File " & file & " does not exists."
			ReadFile = ""
			Exit Function
		End If
		Dim objFile: Set objFile = objFSO.OpenTextFile(file)
		ReadFile = objFile.ReadAll()
		objFile.Close
	End Function

	Public Function FileExists(file)
		FileExists = objFSO.FileExists(file)
	End Function

	Public Sub DeleteFile(file)
		on Error resume next
		objFSO.DeleteFile(file)
		On Error Goto 0
	End Sub
End Class

' ================= inline ================= 


Dim cFS
set cFS = new FSO
Function log(msg)
cFS.WriteLog "bin\\build.log", msg, false
End Function
log "Execution Started for file"


' ================= src : lib/core/globals.vbs ================= 
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
' ================= src : lib/core/Wshell.vbs ================= 

' ================= src : lib/core/VbsJson/VbsJson.vbs ================= 
Class VbsJson
    'Author: Demon
    'Date: 2012/5/3
    'Website: http://demon.tw
    Private Whitespace, NumberRegex, StringChunk
    Private b, f, r, n, t

    Private Sub Class_Initialize
        Whitespace = " " & vbTab & vbCr & vbLf
        b = ChrW(8)
        f = vbFormFeed
        r = vbCr
        n = vbLf
        t = vbTab

        Set NumberRegex = New RegExp
        NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
        NumberRegex.Global = False
        NumberRegex.MultiLine = True
        NumberRegex.IgnoreCase = True

        Set StringChunk = New RegExp
        StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
        StringChunk.Global = False
        StringChunk.MultiLine = True
        StringChunk.IgnoreCase = True
    End Sub
    
    'Return a JSON string representation of a VBScript data structure
    'Supports the following objects and types
    '+-------------------+---------------+
    '| VBScript          | JSON          |
    '+===================+===============+
    '| Dictionary        | object        |
    '+-------------------+---------------+
    '| Array             | array         |
    '+-------------------+---------------+
    '| String            | string        |
    '+-------------------+---------------+
    '| Number            | number        |
    '+-------------------+---------------+
    '| True              | true          |
    '+-------------------+---------------+
    '| False             | false         |
    '+-------------------+---------------+
    '| Null              | null          |
    '+-------------------+---------------+
    Public Function Encode(ByRef obj)
        Dim buf, i, c, g
        Set buf = CreateObject("Scripting.Dictionary")
        Select Case VarType(obj)
            Case vbNull
                buf.Add buf.Count, "null"
            Case vbBoolean
                If obj Then
                    buf.Add buf.Count, "true"
                Else
                    buf.Add buf.Count, "false"
                End If
            Case vbInteger, vbLong, vbSingle, vbDouble
                buf.Add buf.Count, obj
            Case vbString
                buf.Add buf.Count, """"
                For i = 1 To Len(obj)
                    c = Mid(obj, i, 1)
                    Select Case c
                        Case """" buf.Add buf.Count, "\"""
                        Case "\"  buf.Add buf.Count, "\\"
                        Case "/"  buf.Add buf.Count, "/"
                        Case b    buf.Add buf.Count, "\b"
                        Case f    buf.Add buf.Count, "\f"
                        Case r    buf.Add buf.Count, "\r"
                        Case n    buf.Add buf.Count, "\n"
                        Case t    buf.Add buf.Count, "\t"
                        Case Else
                            If AscW(c) >= 0 And AscW(c) <= 31 Then
                                c = Right("0" & Hex(AscW(c)), 2)
                                buf.Add buf.Count, "\u00" & c
                            Else
                                buf.Add buf.Count, c
                            End If
                    End Select
                Next
                buf.Add buf.Count, """"
            Case vbArray + vbVariant
                g = True
                buf.Add buf.Count, "["
                For Each i In obj
                    If g Then g = False Else buf.Add buf.Count, ","
                    buf.Add buf.Count, Encode(i)
                Next
                buf.Add buf.Count, "]"
            Case vbObject
                If TypeName(obj) = "Dictionary" Then
                    g = True
                    buf.Add buf.Count, "{"
                    For Each i In obj
                        If g Then g = False Else buf.Add buf.Count, ","
                        buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
                    Next
                    buf.Add buf.Count, "}"
                Else
                    Err.Raise 8732,,"None dictionary object"
                End If
            Case Else
                buf.Add buf.Count, """" & CStr(obj) & """"
        End Select
        Encode = Join(buf.Items, "")
    End Function

    'Return the VBScript representation of ``str(``
    'Performs the following translations in decoding
    '+---------------+-------------------+
    '| JSON          | VBScript          |
    '+===============+===================+
    '| object        | Dictionary        |
    '+---------------+-------------------+
    '| array         | Array             |
    '+---------------+-------------------+
    '| string        | String            |
    '+---------------+-------------------+
    '| number        | Double            |
    '+---------------+-------------------+
    '| true          | True              |
    '+---------------+-------------------+
    '| false         | False             |
    '+---------------+-------------------+
    '| null          | Null              |
    '+---------------+-------------------+
    Public Function Decode(ByRef str)
        Dim idx
        idx = SkipWhitespace(str, 1)

        If Mid(str, idx, 1) = "{" Then
            Set Decode = ScanOnce(str, 1)
        Else
            Decode = ScanOnce(str, 1)
        End If
    End Function
    
    Private Function ScanOnce(ByRef str, ByRef idx)
        Dim c, ms

        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "{" Then
            idx = idx + 1
            Set ScanOnce = ParseObject(str, idx)
            Exit Function
        ElseIf c = "[" Then
            idx = idx + 1
            ScanOnce = ParseArray(str, idx)
            Exit Function
        ElseIf c = """" Then
            idx = idx + 1
            ScanOnce = ParseString(str, idx)
            Exit Function
        ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = Null
            Exit Function
        ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = True
            Exit Function
        ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
            idx = idx + 5
            ScanOnce = False
            Exit Function
        End If
        
        Set ms = NumberRegex.Execute(Mid(str, idx))
        If ms.Count = 1 Then
            idx = idx + ms(0).Length
            ScanOnce = CDbl(ms(0))
            Exit Function
        End If
        
        Err.Raise 8732,,"No JSON object could be ScanOnced"
    End Function

    Private Function ParseObject(ByRef str, ByRef idx)
        Dim c, key, value
        Set ParseObject = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)
        
        If c = "}" Then
            Exit Function
        ElseIf c <> """" Then
            Err.Raise 8732,,"Expecting property name"
        End If

        idx = idx + 1
        
        Do
            key = ParseString(str, idx)

            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) <> ":" Then
                Err.Raise 8732,,"Expecting : delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            ParseObject.Add key, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "}" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            c = Mid(str, idx, 1)
            If c <> """" Then
                Err.Raise 8732,,"Expecting property name"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
    End Function
    
    Private Function ParseArray(ByRef str, ByRef idx)
        Dim c, values, value
        Set values = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "]" Then
            ParseArray = values.Items
            Exit Function
        End If

        Do
            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            values.Add values.Count, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "]" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
        ParseArray = values.Items
    End Function
    
    Private Function ParseString(ByRef str, ByRef idx)
        Dim chunks, content, terminator, ms, esc, char
        Set chunks = CreateObject("Scripting.Dictionary")

        Do
            Set ms = StringChunk.Execute(Mid(str, idx))
            If ms.Count = 0 Then
                Err.Raise 8732,,"Unterminated string starting"
            End If
            
            content = ms(0).Submatches(0)
            terminator = ms(0).Submatches(1)
            If Len(content) > 0 Then
                chunks.Add chunks.Count, content
            End If
            
            idx = idx + ms(0).Length
            
            If terminator = """" Then
                Exit Do
            ElseIf terminator <> "\" Then
                Err.Raise 8732,,"Invalid control character"
            End If
            
            esc = Mid(str, idx, 1)

            If esc <> "u" Then
                Select Case esc
                    Case """" char = """"
                    Case "\"  char = "\"
                    Case "/"  char = "/"
                    Case "b"  char = b
                    Case "f"  char = f
                    Case "n"  char = n
                    Case "r"  char = r
                    Case "t"  char = t
                    Case Else Err.Raise 8732,,"Invalid escape"
                End Select
                idx = idx + 1
            Else
                char = ChrW("&H" & Mid(str, idx + 1, 4))
                idx = idx + 5
            End If

            chunks.Add chunks.Count, char
        Loop

        ParseString = Join(chunks.Items, "")
    End Function

    Private Function SkipWhitespace(ByRef str, ByVal idx)
        Do While idx <= Len(str) And _
            InStr(Whitespace, Mid(str, idx, 1)) > 0
            idx = idx + 1
        Loop
        SkipWhitespace = idx
    End Function

End Class
' ================= src : lib/core/JSONToXML/JSONToXML.vbs ================= 
' ==============================================================================================
' Adaptation of JSONToXML() function for enhancements and bugfixes.
' Author: Praveen Nandagiri (pravynandas@gmail.com)
' Enhancement#1: Arrays are now rendered as Text Nodes
' Enhancement#2: Handled Escape characters (incl. Hex). Refer: http://www.json.org/
'
' Credits:
' Visit: https://stackoverflow.com/a/12171836/1751166
' Author: https://stackoverflow.com/users/881441/stephen-quan
' ==============================================================================================

Class JSONToXML

  Private stateRoot
  Private stateNameQuoted
  Private stateNameFinished
  Private stateValue
  Private stateValueQuoted
  Private stateValueQuotedEscaped
  Private stateValueQuotedEscapedHex
  Private stateValueUnquoted
  Private stateValueUnquotedEscaped

  Private Sub Class_Initialize
    stateRoot = 0
    stateNameQuoted = 1
    stateNameFinished = 2
    stateValue = 3
    stateValueQuoted = 4
    stateValueQuotedEscaped = 5
    stateValueQuotedEscapedHex = 6
    stateValueUnquoted = 7
    stateValueUnquotedEscaped = 8
  End Sub

  Public Function toXml(json)
    Dim dom, xmlElem, i, ch, state, name, value, sHex
    Set dom = CreateObject("Microsoft.XMLDOM")
    state = stateRoot
    For i = 1 to Len(json)
      ch = Mid(json, i, 1)
      Select Case state
      Case stateRoot
        Select Case ch
        Case "["
          If dom.documentElement is Nothing Then
            Set xmlElem = dom.CreateElement("ARRAY")
            Set dom.documentElement = xmlElem
          Else
            Set xmlElem = XMLCreateChild(xmlElem, "ARRAY")
          End If
        Case "{"
          If dom.documentElement is Nothing Then
            Set xmlElem = dom.CreateElement("ROOT")
            Set dom.documentElement = xmlElem
          Else
            Set xmlElem = XMLCreateChild(xmlElem, "OBJECT")
          End If
        Case """"
          state = stateNameQuoted 
          name = ""
        Case "}"
          Set xmlElem = xmlElem.parentNode
        Case "]"
          Set xmlElem = xmlElem.parentNode
        End Select
      Case stateNameQuoted 
        Select Case ch
        Case """"
          state = stateNameFinished
        Case Else
          name = name + ch
        End Select
      Case stateNameFinished
        Select Case ch
        Case ":"
          value = ""
          State = stateValue
        Case Else						'@@Enhancement#1: Handling Array values
          Set xmlitem = dom.createTextNode(name)
      xmlElem.appendChild(xmlitem)
          State = stateRoot					
        End Select
      Case stateValue
        Select Case ch
        Case """"
          State = stateValueQuoted
        Case "{"
          Set xmlElem = XMLCreateChild(xmlElem, name)
          State = stateRoot
        Case "["
          Set xmlElem = XMLCreateChild(xmlElem, name)
          State = stateRoot
        Case " "
        Case Chr(9)
        Case vbCr
        Case vbLF
        Case Else
          value = ch
          State = stateValueUnquoted
        End Select
      Case stateValueQuoted
        Select Case ch
        Case """"
          xmlElem.setAttribute name, value
          state = stateRoot
        Case "\"
          state = stateValueQuotedEscaped
        Case Else
          value = value + ch
        End Select
      Case stateValueQuotedEscaped ' @@Enhancement#2: Handle escape sequences
      If ch = "u" Then	'Four digit hex. Ex: o = 00f8
        sHex = ""
        state = stateValueQuotedEscapedHex
      Else
        Select Case ch
        Case """"
          value = value + """"
        Case "\"
          value = value + "\"
        Case "/"
          value = value + "/"
        Case "b"	'Backspace
          value = value + chr(08)
        Case "f"	'Form-Feed
          value = value + chr(12)
        Case "n"	'New-line (LineFeed(10))
          value = value + vbLF
        Case "r"	'New-line (CarriageReturn/CRLF(13))
          value = value + vbCR
        Case "t"	'Horizontal-Tab (09)
          value = value + vbTab
        Case Else
          'do not accept any other escape sequence
        End Select
        state = stateValueQuoted
      End If
    Case stateValueQuotedEscapedHex
      sHex = sHex + ch
      If len(sHex) = 4 Then
        on error resume next
        value = value + Chr("&H" & sHex)	'Hex to String conversion
        on error goto 0
        state = stateValueQuoted
      End If
      Case stateValueUnquoted
        Select Case ch
        Case "}"
          xmlElem.setAttribute name, value
          Set xmlElem = xmlElem.parentNode
          state = stateRoot
        Case "]"
          xmlElem.setAttribute name, value
          Set xmlElem = xmlElem.parentNode
          state = stateRoot
        Case ","
          xmlElem.setAttribute name, value
          state = stateRoot
        Case "\"
          state = stateValueUnquotedEscaped
        Case Else
          value = value + ch
        End Select
      Case stateValueUnquotedEscaped ' @@TODO: Handle escape sequences
        value = value + ch
        state = stateValueUnquoted
      End Select
    Next
    set toXml = dom
  End Function

  Private Function XMLCreateChild(xmlParent, tagName)
    Dim xmlChild
    If xmlParent is Nothing Then
      Set XMLCreateChild = Nothing
      Exit Function
    End If
    If xmlParent.ownerDocument is Nothing Then
      Set XMLCreateChild = Nothing
      Exit Function
    End If
    Set xmlChild = xmlParent.ownerDocument.createElement(tagName)
    xmlParent.appendChild xmlChild
    Set XMLCreateChild = xmlChild
  End Function
End Class
'' SIG '' Begin signature block
'' SIG '' MIIR0QYJKoZIhvcNAQcCoIIRwjCCEb4CAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFIAfNENXK9H/
'' SIG '' 6aGZEv90jzPAbBvYoIINQTCCAwYwggHuoAMCAQICEBoe
'' SIG '' /smcD+mIQhSFc8JanE4wDQYJKoZIhvcNAQELBQAwGzEZ
'' SIG '' MBcGA1UEAwwQQVRBIEF1dGhlbnRpY29kZTAeFw0yMTA2
'' SIG '' MDgyMjI0MjBaFw0yMjA2MDgyMjQ0MjBaMBsxGTAXBgNV
'' SIG '' BAMMEEFUQSBBdXRoZW50aWNvZGUwggEiMA0GCSqGSIb3
'' SIG '' DQEBAQUAA4IBDwAwggEKAoIBAQC1qnPumMP+1YKsFrRK
'' SIG '' re5j4Mzk8B59EfJVntNeiuxSCDSzYbgvHLkofXRRpG1m
'' SIG '' DbFhjbtX+lH+qmCF6Zf+NSbE1R2laYTfANShBi5RE70f
'' SIG '' IQ0NGvUGtNPt33CDqkOUUNibpRQO6tfxs82o94v4GekL
'' SIG '' FDAJjWHScqr4zsW3dgD/DixEjoGAWO1UR5FyJ+Z+lJoQ
'' SIG '' hbKX7YhoJsatrAxICRo3XnV5X62LGvLBl5nUa/XPpEZY
'' SIG '' RtTUBcENPK8X8DIRA8meN8NgPidpcriozwFIboaTIjzi
'' SIG '' obf3m+NTjxjGd9sUb148LSAbfHC94D8YQvh4eziytghx
'' SIG '' G47yeQze8ttdBBlRAgMBAAGjRjBEMA4GA1UdDwEB/wQE
'' SIG '' AwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4E
'' SIG '' FgQUSebZ9aqW0d9qm2IaY48oeuR/ssowDQYJKoZIhvcN
'' SIG '' AQELBQADggEBAHinUqMyuCaqZqwGcfdmUY6PEQ+HTMnt
'' SIG '' Yv+2c9niUEeZhUuhs5zVFZ8c1Kvr6n8/An5TgIJHwJB5
'' SIG '' 978W5sCeiTRmySl96ZZT0E+h0t7qupJC7/8HbEPXdYEb
'' SIG '' uxedsfdTDXfRmDk9plQJXG2DRbd+3xB26hblPOHxatOE
'' SIG '' MKaLPppWSnFzc1rwLRNqRARtdYP2IxpnW3u+zqKuK3ZF
'' SIG '' 9Thrj+kouJRIGW0OefvZ8fSinP8q1JrHeAmwgTFqBYzf
'' SIG '' iYtk4n6KlJxA3qW0au0ZlivK/p+nq1oBDfH7sFymv6eE
'' SIG '' 2RJFXaQDYeDchfIkJspdR3c9bsm6DBN0tCzhtAc/Ccg2
'' SIG '' LxkwggT+MIID5qADAgECAhANQkrgvjqI/2BAIc4UAPDd
'' SIG '' MA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUw
'' SIG '' EwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
'' SIG '' dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0
'' SIG '' IFNIQTIgQXNzdXJlZCBJRCBUaW1lc3RhbXBpbmcgQ0Ew
'' SIG '' HhcNMjEwMTAxMDAwMDAwWhcNMzEwMTA2MDAwMDAwWjBI
'' SIG '' MQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
'' SIG '' IEluYy4xIDAeBgNVBAMTF0RpZ2lDZXJ0IFRpbWVzdGFt
'' SIG '' cCAyMDIxMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
'' SIG '' CgKCAQEAwuZhhGfFivUNCKRFymNrUdc6EUK9CnV1TZS0
'' SIG '' DFC1JhD+HchvkWsMlucaXEjvROW/m2HNFZFiWrj/Zwuc
'' SIG '' Y/02aoH6KfjdK3CF3gIY83htvH35x20JPb5qdofpir34
'' SIG '' hF0edsnkxnZ2OlPR0dNaNo/Go+EvGzq3YdZz7E5tM4p8
'' SIG '' XUUtS7FQ5kE6N1aG3JMjjfdQJehk5t3Tjy9XtYcg6w6O
'' SIG '' LNUj2vRNeEbjA4MxKUpcDDGKSoyIxfcwWvkUrxVfbENJ
'' SIG '' Cf0mI1P2jWPoGqtbsR0wwptpgrTb/FZUvB+hh6u+elsK
'' SIG '' IC9LCcmVp42y+tZji06lchzun3oBc/gZ1v4NSYS9AQID
'' SIG '' AQABo4IBuDCCAbQwDgYDVR0PAQH/BAQDAgeAMAwGA1Ud
'' SIG '' EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgw
'' SIG '' QQYDVR0gBDowODA2BglghkgBhv1sBwEwKTAnBggrBgEF
'' SIG '' BQcCARYbaHR0cDovL3d3dy5kaWdpY2VydC5jb20vQ1BT
'' SIG '' MB8GA1UdIwQYMBaAFPS24SAd/imu0uRhpbKiJbLIFzVu
'' SIG '' MB0GA1UdDgQWBBQ2RIaOpLqwZr68KC0dRDbd42p6vDBx
'' SIG '' BgNVHR8EajBoMDKgMKAuhixodHRwOi8vY3JsMy5kaWdp
'' SIG '' Y2VydC5jb20vc2hhMi1hc3N1cmVkLXRzLmNybDAyoDCg
'' SIG '' LoYsaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTIt
'' SIG '' YXNzdXJlZC10cy5jcmwwgYUGCCsGAQUFBwEBBHkwdzAk
'' SIG '' BggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQu
'' SIG '' Y29tME8GCCsGAQUFBzAChkNodHRwOi8vY2FjZXJ0cy5k
'' SIG '' aWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElE
'' SIG '' VGltZXN0YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEBCwUA
'' SIG '' A4IBAQBIHNy16ZojvOca5yAOjmdG/UJyUXQKI0ejq5LS
'' SIG '' JcRwWb4UoOUngaVNFBUZB3nw0QTDhtk7vf5EAmZN7Wmk
'' SIG '' D/a4cM9i6PVRSnh5Nnont/PnUp+Tp+1DnnvntN1BIon7
'' SIG '' h6JGA0789P63ZHdjXyNSaYOC+hpT7ZDMjaEXcw3082U5
'' SIG '' cEvznNZ6e9oMvD0y0BvL9WH8dQgAdryBDvjA4VzPxBFy
'' SIG '' 5xtkSdgimnUVQvUtMjiB2vRgorq0Uvtc4GEkJU+y38kp
'' SIG '' qHNDUdq9Y9YfW5v3LhtPEx33Sg1xfpe39D+E68Hjo0mh
'' SIG '' +s6nv1bPull2YYlffqe0jmd4+TaY4cso2luHpoovMIIF
'' SIG '' MTCCBBmgAwIBAgIQCqEl1tYyG35B5AXaNpfCFTANBgkq
'' SIG '' hkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UE
'' SIG '' ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
'' SIG '' aWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1
'' SIG '' cmVkIElEIFJvb3QgQ0EwHhcNMTYwMTA3MTIwMDAwWhcN
'' SIG '' MzEwMTA3MTIwMDAwWjByMQswCQYDVQQGEwJVUzEVMBMG
'' SIG '' A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
'' SIG '' ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBT
'' SIG '' SEEyIEFzc3VyZWQgSUQgVGltZXN0YW1waW5nIENBMIIB
'' SIG '' IjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAvdAy
'' SIG '' 7kvNj3/dqbqCmcU5VChXtiNKxA4HRTNREH3Q+X1NaH7n
'' SIG '' tqD0jbOI5Je/YyGQmL8TvFfTw+F+CNZqFAA49y4eO+7M
'' SIG '' pvYyWf5fZT/gm+vjRkcGGlV+Cyd+wKL1oODeIj8O/36V
'' SIG '' +/OjuiI+GKwR5PCZA207hXwJ0+5dyJoLVOOoCXFr4M8i
'' SIG '' EA91z3FyTgqt30A6XLdR4aF5FMZNJCMwXbzsPGBqrC8H
'' SIG '' zP3w6kfZiFBe/WZuVmEnKYmEUeaC50ZQ/ZQqLKfkdT66
'' SIG '' mA+Ef58xFNat1fJky3seBdCEGXIX8RcG7z3N1k3vBkL9
'' SIG '' olMqT4UdxB08r8/arBD13ays6Vb/kwIDAQABo4IBzjCC
'' SIG '' AcowHQYDVR0OBBYEFPS24SAd/imu0uRhpbKiJbLIFzVu
'' SIG '' MB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgP
'' SIG '' MBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQD
'' SIG '' AgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHkGCCsGAQUF
'' SIG '' BwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3Au
'' SIG '' ZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8v
'' SIG '' Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
'' SIG '' cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2
'' SIG '' hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNl
'' SIG '' cnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRw
'' SIG '' Oi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
'' SIG '' cmVkSURSb290Q0EuY3JsMFAGA1UdIARJMEcwOAYKYIZI
'' SIG '' AYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3
'' SIG '' dy5kaWdpY2VydC5jb20vQ1BTMAsGCWCGSAGG/WwHATAN
'' SIG '' BgkqhkiG9w0BAQsFAAOCAQEAcZUS6VGHVmnN793afKpj
'' SIG '' erN4zwY3QITvS4S/ys8DAv3Fp8MOIEIsr3fzKx8MIVoq
'' SIG '' twU0HWqumfgnoma/Capg33akOpMP+LLR2HwZYuhegiUe
'' SIG '' xLoceywh4tZbLBQ1QwRostt1AuByx5jWPGTlH0gQGF+J
'' SIG '' OGFNYkYkh2OMkVIsrymJ5Xgf1gsUpYDXEkdws3XVk4WT
'' SIG '' fraSZ/tTYYmo9WuWwPRYaQ18yAGxuSh1t5ljhSKMYcp5
'' SIG '' lH5Z/IwP42+1ASa2bKXuh1Eh5Fhgm7oMLSttosR+u8Ql
'' SIG '' K0cCCHxJrhO24XxCQijGGFbPQTS2Zl22dHv1VjMiLyI2
'' SIG '' skuiSpXY9aaOUjGCA/wwggP4AgEBMC8wGzEZMBcGA1UE
'' SIG '' AwwQQVRBIEF1dGhlbnRpY29kZQIQGh7+yZwP6YhCFIVz
'' SIG '' wlqcTjAJBgUrDgMCGgUAoHAwEAYKKwYBBAGCNwIBDDEC
'' SIG '' MAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
'' SIG '' KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZI
'' SIG '' hvcNAQkEMRYEFGuT6PAJHz2opt77a3X7Cqsl/2sfMA0G
'' SIG '' CSqGSIb3DQEBAQUABIIBAH21XsZK28FcsCcR6h4ZoqrI
'' SIG '' 4dpMVcO4PGvsQ9VkzPOGP1mn4sOkdk6isWxBZInmGVAq
'' SIG '' 3bdHDeGH/8FmvP6L+aBq1lditO0ti079UKvc5qN6v6Ps
'' SIG '' ZBAEFwKjhsybjyI0WSVkAt3oQNAYLBNVbWOvmZ6WYZ9H
'' SIG '' BWjhYxJgCXSGi+EpNV3KjIJ5f0nsIt1bdmuKZhQpI9Hu
'' SIG '' Cg814AATC4LmdnLLM9h714WHanKYEjVtMjK1FqCvleUU
'' SIG '' jq3Ql2oJpXmReNM1csG2IjGxnxRT/ZcTt3Ebcrt/Mb0g
'' SIG '' faFu/8M463eM+Q4VOR8DGGSYN0zuzxcMUeUuLPRGtQr+
'' SIG '' rjH880IsQyehggIwMIICLAYJKoZIhvcNAQkGMYICHTCC
'' SIG '' AhkCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoT
'' SIG '' DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
'' SIG '' ZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBB
'' SIG '' c3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQQIQDUJK4L46
'' SIG '' iP9gQCHOFADw3TANBglghkgBZQMEAgEFAKBpMBgGCSqG
'' SIG '' SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkF
'' SIG '' MQ8XDTIxMDYwODIyNTIwOFowLwYJKoZIhvcNAQkEMSIE
'' SIG '' IJ2OpJTxfwrh4uFQszDj9N5h+ZshupcIt38YVSF+BQB7
'' SIG '' MA0GCSqGSIb3DQEBAQUABIIBAHKZj0LMaeTITCbLaFLC
'' SIG '' SUHp8bOIekHSJPaQa20vv+nHasc2uZvRmOrZrkGs4Amc
'' SIG '' RSqRnmX8FmdueqyJdqZVy8PEiWhaRDaI7fLBB7PbVtI1
'' SIG '' JxQz7MqRDVbOy106vwKeZ/DbMYM97MsV0HN1CQ+fbOOa
'' SIG '' XONeThal0Vu4FDEfpBGzfzxDJfbzD2AvWqkK7BJdnUqb
'' SIG '' K4n2igcDk020Flaq+pwSUrOolnykX2xDH7l5mInp86QW
'' SIG '' ZqL4o194k4+7e7RawEIjw4xlktewndOJp5mHx6TS+u3U
'' SIG '' vDGxOBQXny61trAJ22SP+4HiPZnUvVr4jT8k5BfJ5CxY
'' SIG '' jOggmD8v3m8Nd9w=
'' SIG '' End signature block

' ================= src : lib/core/Signtool/Signtool.vbs ================= 
' Set objShell = CreateObject("Wscript.Shell")
' Set objFSO = CreateObject("Scripting.FileSystemObject")
' strPath = Wscript.ScriptFullName
' Set objFile = objFSO.GetFile(strPath)
' strFolder = objFSO.GetParentFolderName(objFile) 
' objShell.CurrentDirectory = strFolder
' strPath = ".\signtool-x64.exe sign /f .\ata-authenticode-signer.pfx /p pwd /t http://timestamp.digicert.com " + Wscript.Arguments(0)
' objShell.Run strPath, 0, true

Class Signtool

    private cWShell

    private Sub Class_Initialize
        set cWShell = new WShell
        if cWShell is nothing then
            Wscript.Echo "Signer Class: Unable to initialize WShell class."
            Wscript.Quit
        end if
    End Sub

    public Function Sign(file, pwd)
        Wscript.Echo "Signing file: " & file

        Dim signtool: signtool = "C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\signtool\signtool-x64.exe"
        Dim cert: cert = "C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\signtool\ata-authenticode-signer.pfx"
        Dim timestamp: timestamp = "http://timestamp.digicert.com"
        Dim strPath: strPath = signtool & " sign /f " & cert & " /p " & pwd & " /t " & timestamp & " " + file
        cWShell.Exec(strPath)
    End Function

End Class
' ================= src : CPOL/vb_format_function.vbs ================= 

' ================= inline ================= 


Dim file: file = WScript.Arguments.Named("file")
log "File: " & file
Dim script: script = cFS.ReadFile(file)
if script = "" Then
Wscript.Echo "Script file called is empty."
Wscript.Quit
End if
Execute script

