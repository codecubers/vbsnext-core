Option Explicit


' ================= src : lib/core/globals.vbs ================= 


' ================= src : lib/core/FSO/FSO.vbs ================= 
' ==============================================================================================
' Implementation of several use cases of FileSystemObject into this class
' Author: Praveen Nandagiri (pravynandas@gmail.com)
' ==============================================================================================

Class FSO
	Private dir
	Private objFSO
	
	Private Sub Class_Initialize
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		dir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
	End Sub

	' Update the current directory of the instance if needed
	public Sub setDir(s)
		dir = s
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
	
	Public Sub WriteFile(strFileName, strMessage, overwrite)
		Const ForReading = 1
		Const ForWriting = 2
		Const ForAppending = 8
		Dim mode
    Dim oFile
		
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

  Public Function GetExtn(file)
    GetExtn = ""
    on Error Resume Next
    GetExtn = objFSO.GetExtensionName(file)
    On Error goto 0
  End Function

End Class


' ================================== Job: vbspm-build ================================== 

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
        dir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
    End Sub
    
    ' Update the current directory of the instance if needed
    public Sub setDir(s)
        dir = s
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



Dim baseDir
Dim cFS
set cFS = new FSO


public Sub Echo(msg)
  Wscript.Echo msg
End Sub
Function log(msg)
  cFS.WriteFile "build.log", msg, false
End Function
log "================================= Call ================================="

Sub Include(pkg, file)
  log "Include(" + file + ")"
  
  Dim content: content = cFS.ReadFile(file)
  if content <> "" Then 
    cFS.WriteFile "build\imported\" & pkg & ".vbs", content, true
    ExecuteGlobal content
  End If
End Sub

Public Sub Import(pkg)
  log "Import(" + Pkg + ")"
  Include pkg, baseDir & "\node_modules\" + pkg + "\index.vbs"
End Sub


With CreateObject("WScript.Shell")
  baseDir=.CurrentDirectory
  'Wscript.Echo  "Base path: " & baseDir
End With
log "Base path: " & baseDir
cFS.setDir(baseDir)


' ================= inline ================= 


log "Execution Started for file"

Dim file
file = WScript.Arguments.Named("file")
If file = "" Then
    log "Script file not provided as a named argument [/file:]"
    if Wscript.Arguments.count > 0 then
      file = Wscript.Arguments(0) 
      if file = "" Then
        log "No file argument provided."
        Wscript.Quit
      End If
    else 
      file = "index.vbs"
    end if
End If
' TODO: Assess all possible combinations a user can send in command line
file = baseDir & "\" & file

if Lcase(cFS.GetExtn(file)) = "vbs" Then
  log "File extension is: .vbs"
Else
  log "File extension missing. Adding .vbs"
  file = file + ".vbs"
end if

log "File: " & file


Dim script
script = cFS.ReadFile(file)
if script = "" Then
  log "No file supplied or is empty."
  Wscript.Quit
End if



'=========================== 
Execute script
'=========================== 