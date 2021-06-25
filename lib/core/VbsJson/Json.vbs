Option Explicit
' Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'                                 (pDst, pSrc, ByVal ByteLen)
Class cStringBuilder
    ' ======================================================================================
    ' Name:     vbAccelerator cStringBuilder
    ' Author:   Steve McMahon (steve@vbaccelerator.com)
    ' Date:     1 January 2002
    '
    ' Copyright © 2002 Steve McMahon for vbAccelerator
    ' --------------------------------------------------------------------------------------
    ' Visit vbAccelerator - advanced free source code for VB programmers
    ' http://vbaccelerator.com
    ' --------------------------------------------------------------------------------------
    '
    ' VB can be slow to append strings together because of the continual
    ' reallocation of string size.  This class pre-allocates a string in
    ' blocks and hence removes the performance restriction.
    '
    ' Quicker insert and remove is also possible since string space does
    ' not have to be reallocated.
    '
    ' Example:
    ' Adding "http://vbaccelerator.com/" 10,000 times to a string:
    ' Standard VB:   34s
    ' This Class:    0.35s
    '
    ' ======================================================================================



    Private m_sString
    Private m_iChunkSize
    Private m_iPos
    Private m_iLen

    Public Property Get Length()
        Length = m_iPos \ 2
    End Property

    Public Property Get Capacity()
        Capacity = m_iLen \ 2
    End Property

    Public Property Get ChunkSize()
    ' Return the unicode character chunk size:
        ChunkSize = m_iChunkSize \ 2
    End Property

    Public Property Let ChunkSize(ByVal iChunkSize)
    ' Set the chunksize.  We multiply by 2 because internally
    ' we are considering bytes:
        m_iChunkSize = iChunkSize * 2
    End Property

    Public Property Get toString()
    ' The internal string:
        If m_iPos > 0 Then
            toString = Left(m_sString, m_iPos \ 2)
        End If
    End Property

    Public Property Let TheString(ByRef sThis)
        Dim lLen

        ' Setting the string:
        lLen = LenB(sThis)
        If lLen = 0 Then
            'Clear
            m_sString = ""
            m_iPos = 0
            m_iLen = 0
        Else
            If m_iLen < lLen Then
                ' Need to expand string to accommodate:
                Do
                    m_sString = m_sString & Space(m_iChunkSize \ 2)
                    m_iLen = m_iLen + m_iChunkSize
                Loop While m_iLen < lLen
            End If
            ' CopyMemory ByVal StrPtr(m_sString), ByVal StrPtr(sThis), lLen
            m_iPos = lLen
        End If

    End Property

    Public Sub Clear()
        m_sString = ""
        m_iPos = 0
        m_iLen = 0
    End Sub

    Public Sub AppendNL(ByRef sThis)
        Append sThis
        Append vbCrLf
    End Sub

    Public Sub Append(ByRef sThis)
        Dim lLen
        Dim lLenPlusPos

        ' Append an item to the string:
        lLen = LenB(sThis)
        lLenPlusPos = lLen + m_iPos
        If lLenPlusPos > m_iLen Then
            Dim lTemp

            lTemp = m_iLen
            Do While lTemp < lLenPlusPos
                lTemp = lTemp + m_iChunkSize
            Loop

            m_sString = m_sString & Space((lTemp - m_iLen) \ 2)
            m_iLen = lTemp
        End If

        ' CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
        m_iPos = m_iPos + lLen
    End Sub

    Public Sub AppendByVal(ByVal sThis)
        Append sThis
    End Sub

    Public Sub Insert(ByVal iIndex, ByRef sThis)
        Dim lLen
        Dim lPos
        Dim lSize

        ' is iIndex within bounds?
        If (iIndex * 2 > m_iPos) Then
            Err.Raise 9
        Else

            lLen = LenB(sThis)
            If (m_iPos + lLen) > m_iLen Then
                m_sString = m_sString & Space(m_iChunkSize \ 2)
                m_iLen = m_iLen + m_iChunkSize
            End If

            ' Move existing characters from current position
            lPos = UnsignedAdd(StrPtr(m_sString), iIndex * 2)
            lSize = m_iPos - iIndex * 2

            ' moving from iIndex to iIndex + lLen
            ' CopyMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize

            ' Insert new characters:
            ' CopyMemory ByVal lPos, ByVal StrPtr(sThis), lLen

            m_iPos = m_iPos + lLen
        End If
    End Sub

    Public Sub InsertByVal(ByVal iIndex, ByVal sThis)
        Insert iIndex, sThis
    End Sub

    Public Sub Remove(ByVal iIndex, ByVal lLen)
        Dim lSrc
        Dim lDst
        Dim lSize

        ' is iIndex within bounds?
        If (iIndex * 2 > m_iPos) Then
            Err.Raise 9
        Else
            ' is there sufficient length?
            If ((iIndex + lLen) * 2 > m_iPos) Then
                Err.Raise 9
            Else
                ' Need to copy characters from iIndex*2 to m_iPos back by lLen chars:
                lSrc = UnsignedAdd(StrPtr(m_sString), (iIndex + lLen) * 2)
                lDst = UnsignedAdd(StrPtr(m_sString), iIndex * 2)
                lSize = (m_iPos - (iIndex + lLen) * 2)
                ' CopyMemory ByVal lDst, ByVal lSrc, lSize
                m_iPos = m_iPos - lLen * 2
            End If
        End If
    End Sub

    Public Function Find(ByVal sToFind, _
                        ByVal lStartIndex, _
                        ByVal compare _
                        )

        Dim lInstr
        If (lStartIndex > 0) Then
            lInstr = InStr(lStartIndex, m_sString, sToFind, compare)
        Else
            lInstr = InStr(m_sString, sToFind, compare)
        End If
        If (lInstr < m_iPos \ 2) Then
            Find = lInstr
        End If
    End Function

    Public Sub HeapMinimize()
        Dim iLen

        ' Reduce the string size so only the minimal chunks
        ' are allocated:
        If (m_iLen - m_iPos) > m_iChunkSize Then
            iLen = m_iLen
            Do While (iLen - m_iPos) > m_iChunkSize
                iLen = iLen - m_iChunkSize
            Loop
            m_sString = Left(m_sString, iLen \ 2)
            m_iLen = iLen
        End If

    End Sub
    Private Function UnsignedAdd(Start, Incr)
    ' This function is useful when doing pointer arithmetic,
    ' but note it only works for positive values of Incr

        If Start And &H80000000 Then    'Start < 0
            UnsignedAdd = Start + Incr
        ElseIf (Start Or &H80000000) < -Incr Then
            UnsignedAdd = Start + Incr
        Else
            UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
        End If

    End Function
    Private Sub Class_Initialize()
    ' The default allocation: 8192 characters.
        m_iChunkSize = 16384
    End Sub

End Class

Class cJSONScript

    Dim dictVars
    Dim plNestCount


    Public Sub Class_Initialize
        set dictVars = CreateObject("Scripting.Dictionary")
    End Sub

    Public Function Eval(sJSON)
        Dim SB: set SB  = New cStringBuilder
        Dim o
        Dim c
        Dim i

        Set o = JSON.parse(sJSON)
        If (JSON.GetParserErrors = "") And Not (o Is Nothing) Then
            For i = 1 To o.Count
                Select Case VarType(o.Item(i))
                Case vbNull
                    SB.Append "null"
                Case vbDate
                    SB.Append CStr(o.Item(i))
                Case vbString
                    SB.Append CStr(o.Item(i))
                Case Else
                    Set c = o.Item(i)
                    SB.Append ExecCommand(c)
                End Select
            Next
        Else
            MsgBox JSON.GetParserErrors, vbExclamation, "Parser Error"
        End If
        Eval = SB.toString
    End Function

    Public Function ExecCommand(ByRef obj)
        Dim SB: set SB = New cStringBuilder

        If plNestCount > 40 Then
            ExecCommand = "ERROR: Nesting level exceeded."
        Else
            plNestCount = plNestCount + 1

            Select Case VarType(obj)
            Case vbNull
                SB.Append "null"
            Case vbDate
                SB.Append CStr(obj)
            Case vbString
                SB.Append CStr(obj)
            Case vbObject

                Dim i
                Dim j
                Dim this
                Dim key
                Dim paramKeys

                If TypeName(obj) = "Dictionary" Then
                    Dim sOut
                    Dim sRet

                    Dim keys
                    keys = obj.keys
                    For i = 0 To obj.Count - 1
                        sRet = ""

                        key = keys(i)
                        If VarType(obj.Item(key)) = vbString Then
                            sRet = obj.Item(key)
                        Else
                            Set this = obj.Item(key)
                        End If

                        ' command implementation
                        Select Case LCase(key)
                        Case "alert":
                            MsgBox ExecCommand(this.Item("message")), vbInformation, ExecCommand(this.Item("title"))

                        Case "input":
                            SB.Append InputBox(ExecCommand(this.Item("prompt")), ExecCommand(this.Item("title")), ExecCommand(this.Item("default")))

                        Case "switch"
                            sOut = ExecCommand(this.Item("default"))
                            sRet = LCase(ExecCommand(this.Item("case")))
                            For j = 0 To this.Item("items").Count - 1
                                If LCase(this.Item("items").Item(j + 1).Item("case")) = sRet Then
                                    sOut = ExecCommand(this.Item("items").Item(j + 1).Item("return"))
                                    Exit For
                                End If
                            Next
                            SB.Append sOut

                        Case "set":
                            If dictVars.Exists(this.Item("name")) Then
                                dictVars.Item(this.Item("name")) = ExecCommand(this.Item("value"))
                            Else
                                dictVars.Add this.Item("name"), ExecCommand(this.Item("value"))
                            End If

                        Case "get":
                            sRet = ExecCommand(dictVars(CStr(this.Item("name"))))
                            If sRet = "" Then
                                sRet = ExecCommand(this.Item("default"))
                            End If

                            SB.Append sRet

                        Case "if"
                            Dim val1
                            Dim val2
                            Dim bRes
                            val1 = ExecCommand(this.Item("value1"))
                            val2 = ExecCommand(this.Item("value2"))

                            bRes = False
                            Select Case LCase(this.Item("type"))
                            Case "eq"    ' =
                                If LCase(val1) = LCase(val2) Then
                                    bRes = True
                                End If

                            Case "gt"    ' >
                                If val1 > val2 Then
                                    bRes = True
                                End If

                            Case "lt"    ' <
                                If val1 < val2 Then
                                    bRes = True
                                End If

                            Case "gte"    ' >=
                                If val1 >= val2 Then
                                    bRes = True
                                End If

                            Case "lte"    ' <=
                                If val1 <= val2 Then
                                    bRes = True
                                End If

                            End Select

                            If bRes Then
                                SB.Append ExecCommand(this.Item("true"))
                            Else
                                SB.Append ExecCommand(this.Item("false"))
                            End If

                        Case "return"
                            SB.Append obj.Item(key)


                        Case Else
                            If TypeName(this) = "Dictionary" Then
                                paramKeys = this.keys
                                For j = 0 To this.Count - 1
                                    If j > 0 Then
                                        sRet = sRet & ","
                                    End If
                                    sRet = sRet & CStr(this.Item(paramKeys(j)))
                                Next
                            End If


                            SB.Append "<%" & UCase(key) & "(" & sRet & ")%>"

                        End Select
                    Next

                ElseIf TypeName(obj) = "Collection" Then

                    Dim Value
                    For Each Value In obj
                        SB.Append ExecCommand(Value)
                    Next

                End If
                Set this = Nothing

            Case vbBoolean
                If obj Then SB.Append "true" Else SB.Append "false"

            Case vbVariant, vbArray, vbArray + vbVariant

            Case Else
                SB.Append Replace(obj, ",", ".")
            End Select
            plNestCount = plNestCount - 1
        End If

        ExecCommand = SB.toString
        Set SB = Nothing

    End Function

End Class

Const INVALID_JSON = 1
Const INVALID_OBJECT = 2
Const INVALID_ARRAY = 3
Const INVALID_BOOLEAN = 4
Const INVALID_NULL = 5
Const INVALID_KEY = 6
Const INVALID_RPC_CALL = 7

Class JSON
    ' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
    ' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
    ' BSD Licensed


    

    Private psErrors

    Public Function GetParserErrors()
        GetParserErrors = psErrors
    End Function

    Public Function ClearParserErrors()
        psErrors = ""
    End Function


    '
    '   parse string and create JSON object
    '
    Public Function parse(ByVal str)
        Dim index
        index = 1
        psErrors = ""
        On Error Resume Next
        Call skipChar(str, index)
        Select Case Mid(str, index, 1)
        Case "{"
            Set parse = parseObject(str, index)
        Case "["
            Set parse = parseArray(str, index)
        Case Else
            psErrors = "Invalid JSON"
        End Select


    End Function

    '
    '   parse collection of key/value
    '
    Private Function parseObject(ByRef str, ByRef index)

        Set parseObject = New Dictionary
        Dim sKey

        ' "{"
        Call skipChar(str, index)
        If Mid(str, index, 1) <> "{" Then
            psErrors = psErrors & "Invalid Object at position " & index & " : " & Mid(str, index) & vbCrLf
            Exit Function
        End If

        index = index + 1

        Do
            Call skipChar(str, index)
            If "}" = Mid(str, index, 1) Then
                index = index + 1
                Exit Do
            ElseIf "," = Mid(str, index, 1) Then
                index = index + 1
                Call skipChar(str, index)
            ElseIf index > Len(str) Then
                psErrors = psErrors & "Missing '}': " & Right(str, 20) & vbCrLf
                Exit Do
            End If


            ' add key/value pair
            sKey = parseKey(str, index)
            On Error Resume Next

            parseObject.Add sKey, parseValue(str, index)
            If Err.Number <> 0 Then
                psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
                Exit Do
            End If
        Loop
    eh:

    End Function

    '
    '   parse list
    '
    Private Function parseArray(ByRef str, ByRef index)

        Set parseArray = New Collection

        ' "["
        Call skipChar(str, index)
        If Mid(str, index, 1) <> "[" Then
            psErrors = psErrors & "Invalid Array at position " & index & " : " + Mid(str, index, 20) & vbCrLf
            Exit Function
        End If

        index = index + 1

        Do

            Call skipChar(str, index)
            If "]" = Mid(str, index, 1) Then
                index = index + 1
                Exit Do
            ElseIf "," = Mid(str, index, 1) Then
                index = index + 1
                Call skipChar(str, index)
            ElseIf index > Len(str) Then
                psErrors = psErrors & "Missing ']': " & Right(str, 20) & vbCrLf
                Exit Do
            End If

            ' add value
            On Error Resume Next
            parseArray.Add parseValue(str, index)
            If Err.Number <> 0 Then
                psErrors = psErrors & Err.Description & ": " & Mid(str, index, 20) & vbCrLf
                Exit Do
            End If
        Loop

    End Function

    '
    '   parse string / number / object / array / true / false / null
    '
    Private Function parseValue(ByRef str, ByRef index)

        Call skipChar(str, index)

        Select Case Mid(str, index, 1)
        Case "{"
            Set parseValue = parseObject(str, index)
        Case "["
            Set parseValue = parseArray(str, index)
        Case """", "'"
            parseValue = parseString(str, index)
        Case "t", "f"
            parseValue = parseBoolean(str, index)
        Case "n"
            parseValue = parseNull(str, index)
        Case Else
            parseValue = parseNumber(str, index)
        End Select

    End Function

    '
    '   parse string
    '
    Private Function parseString(ByRef str, ByRef index)

        Dim quote
        Dim Char
        Dim Code

        Dim SB: set SB = New cStringBuilder

        Call skipChar(str, index)
        quote = Mid(str, index, 1)
        index = index + 1

        Do While index > 0 And index <= Len(str)
            Char = Mid(str, index, 1)
            Select Case (Char)
            Case "\"
                index = index + 1
                Char = Mid(str, index, 1)
                Select Case (Char)
                Case """", "\", "/", "'"
                    SB.Append Char
                    index = index + 1
                Case "b"
                    SB.Append vbBack
                    index = index + 1
                Case "f"
                    SB.Append vbFormFeed
                    index = index + 1
                Case "n"
                    SB.Append vbLf
                    index = index + 1
                Case "r"
                    SB.Append vbCr
                    index = index + 1
                Case "t"
                    SB.Append vbTab
                    index = index + 1
                Case "u"
                    index = index + 1
                    Code = Mid(str, index, 4)
                    SB.Append ChrW(Val("&h" + Code))
                    index = index + 4
                End Select
            Case quote
                index = index + 1

                parseString = SB.toString
                Set SB = Nothing

                Exit Function

            Case Else
                SB.Append Char
                index = index + 1
            End Select
        Loop

        parseString = SB.toString
        Set SB = Nothing

    End Function

    '
    '   parse number
    '
    Private Function parseNumber(ByRef str, ByRef index)

        Dim Value
        Dim Char

        Call skipChar(str, index)
        Do While index > 0 And index <= Len(str)
            Char = Mid(str, index, 1)
            If InStr("+-0123456789.eE", Char) Then
                Value = Value & Char
                index = index + 1
            Else
                parseNumber = CDec(Value)
                Exit Function
            End If
        Loop
    End Function

    '
    '   parse true / false
    '
    Private Function parseBoolean(ByRef str, ByRef index)

        Call skipChar(str, index)
        If Mid(str, index, 4) = "true" Then
            parseBoolean = True
            index = index + 4
        ElseIf Mid(str, index, 5) = "false" Then
            parseBoolean = False
            index = index + 5
        Else
            psErrors = psErrors & "Invalid Boolean at position " & index & " : " & Mid(str, index) & vbCrLf
        End If

    End Function

    '
    '   parse null
    '
    Private Function parseNull(ByRef str, ByRef index)

        Call skipChar(str, index)
        If Mid(str, index, 4) = "null" Then
            parseNull = Null
            index = index + 4
        Else
            psErrors = psErrors & "Invalid null value at position " & index & " : " & Mid(str, index) & vbCrLf
        End If

    End Function

    Private Function parseKey(ByRef str, ByRef index)

        Dim dquote
        Dim squote
        Dim Char

        Call skipChar(str, index)
        Do While index > 0 And index <= Len(str)
            Char = Mid(str, index, 1)
            Select Case (Char)
            Case """"
                dquote = Not dquote
                index = index + 1
                If Not dquote Then
                    Call skipChar(str, index)
                    If Mid(str, index, 1) <> ":" Then
                        psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                        Exit Do
                    End If
                End If
            Case "'"
                squote = Not squote
                index = index + 1
                If Not squote Then
                    Call skipChar(str, index)
                    If Mid(str, index, 1) <> ":" Then
                        psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                        Exit Do
                    End If
                End If
            Case ":"
                index = index + 1
                If Not dquote And Not squote Then
                    Exit Do
                Else
                    parseKey = parseKey & Char
                End If
            Case Else
                If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
                Else
                    parseKey = parseKey & Char
                End If
                index = index + 1
            End Select
        Loop

    End Function

    '
    '   skip special character
    '
    Private Sub skipChar(ByRef str, ByRef index)
        Dim bComment
        Dim bStartComment
        Dim bLongComment
        Do While index > 0 And index <= Len(str)
            Select Case Mid(str, index, 1)
            Case vbCr, vbLf
                If Not bLongComment Then
                    bStartComment = False
                    bComment = False
                End If

            Case vbTab, " ", "(", ")"

            Case "/"
                If Not bLongComment Then
                    If bStartComment Then
                        bStartComment = False
                        bComment = True
                    Else
                        bStartComment = True
                        bComment = False
                        bLongComment = False
                    End If
                Else
                    If bStartComment Then
                        bLongComment = False
                        bStartComment = False
                        bComment = False
                    End If
                End If

            Case "*"
                If bStartComment Then
                    bStartComment = False
                    bComment = True
                    bLongComment = True
                Else
                    bStartComment = True
                End If

            Case Else
                If Not bComment Then
                    Exit Do
                End If
            End Select

            index = index + 1
        Loop

    End Sub

    Public Function toString(ByRef obj)
        Dim SB: set SB = New cStringBuilder
        Select Case VarType(obj)
        Case vbNull
            SB.Append "null"
        Case vbDate
            SB.Append """" & CStr(obj) & """"
        Case vbString
            SB.Append """" & Encode(obj) & """"
        Case vbObject

            Dim bFI
            Dim i

            bFI = True
            If TypeName(obj) = "Dictionary" Then

                SB.Append "{"
                Dim keys
                keys = obj.keys
                For i = 0 To obj.Count - 1
                    If bFI Then bFI = False Else SB.Append ","
                    Dim key
                    key = keys(i)
                    SB.Append """" & key & """:" & toString(obj.Item(key))
                Next 
                SB.Append "}"

            ElseIf TypeName(obj) = "Collection" Then

                SB.Append "["
                Dim Value
                For Each Value In obj
                    If bFI Then bFI = False Else SB.Append ","
                    SB.Append toString(Value)
                Next 
                SB.Append "]"

            End If
        Case vbBoolean
            If obj Then SB.Append "true" Else SB.Append "false"
        Case vbVariant, vbArray, vbArray + vbVariant
            Dim sEB
            SB.Append multiArray(obj, 1, "", sEB)
        Case Else
            SB.Append Replace(obj, ",", ".")
        End Select

        toString = SB.toString
        Set SB = Nothing

    End Function

    Private Function Encode(str)

        Dim SB: set SB = New cStringBuilder
        Dim i
        Dim j
        Dim aL1
        Dim aL2
        Dim c
        Dim p

        aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
        aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
        For i = 1 To Len(str)
            p = True
            c = Mid(str, i, 1)
            For j = 0 To 7
                If c = Chr(aL1(j)) Then
                    SB.Append "\" & Chr(aL2(j))
                    p = False
                    Exit For
                End If
            Next

            If p Then
                Dim a
                a = AscW(c)
                If a > 31 And a < 127 Then
                    SB.Append c
                ElseIf a > -1 Or a < 65535 Then
                    SB.Append "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
                End If
            End If
        Next

        Encode = SB.toString
        Set SB = Nothing

    End Function

    Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition

        Dim iDU
        Dim iDL
        Dim i

        On Error Resume Next
        iDL = LBound(aBD, iBC)
        iDU = UBound(aBD, iBC)

        Dim SB : set SB = New cStringBuilder

        Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
        If Err.Number = 9 Then
            sPB1 = sPT & sPS
            For i = 1 To Len(sPB1)
                If i <> 1 Then sPB2 = sPB2 & ","
                sPB2 = sPB2 & Mid(sPB1, i, 1)
            Next
            '        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
            SB.Append toString(aBD(sPB2))
        Else
            sPT = sPT & sPS
            SB.Append "["
            For i = iDL To iDU
                SB.Append multiArray(aBD, iBC + 1, i, sPT)
                If i < iDU Then SB.Append ","
            Next
            SB.Append "]"
            sPT = Left(sPT, iBC - 2)
        End If
        Err.Clear
        multiArray = SB.toString

        Set SB = Nothing
    End Function

    ' Miscellaneous JSON functions

    Public Function StringToJSON(st)

        Const FIELD_SEP = "~"
        Const RECORD_SEP = "|"

        Dim sFlds
        Dim sRecs: set sRecs = New cStringBuilder
        Dim lRecCnt
        Dim lFld
        Dim fld
        Dim rows

        lRecCnt = 0
        If st = "" Then
            StringToJSON = "null"
        Else
            rows = Split(st, RECORD_SEP)
            For lRecCnt = LBound(rows) To UBound(rows)
                sFlds = ""
                fld = Split(rows(lRecCnt), FIELD_SEP)
                For lFld = LBound(fld) To UBound(fld) Step 2
                    sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
                Next    'fld
                sRecs.Append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
            Next    'rec
            StringToJSON = ("( {""Records"": [" & vbCrLf & sRecs.toString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
        End If
    End Function


    Public Function RStoJSON(rs)
        ' On Error GoTo errHandler
        Dim sFlds
        Dim sRecs: set sRecs = New cStringBuilder
        Dim lRecCnt
        Dim fld

        lRecCnt = 0
        If rs.State = adStateClosed Then
            RStoJSON = "null"
        Else
            If rs.EOF Or rs.BOF Then
                RStoJSON = "null"
            Else
                Do While Not rs.EOF And Not rs.BOF
                    lRecCnt = lRecCnt + 1
                    sFlds = ""
                    For Each fld In rs.Fields
                        sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.Name & """:""" & toUnicode(fld.Value & "") & """")
                    Next    'fld
                    sRecs.Append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
                    rs.MoveNext
                Loop
                RStoJSON = ("( {""Records"": [" & vbCrLf & sRecs.toString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
            End If
        End If

        Exit Function
    ' errHandler:

    End Function

    'Public Function JsonRpcCall(url, methName, args(), user, pwd)
    '    Dim r
    '    Dim cli
    '    Dim pText
    '    Static reqId 
    '
    '    reqId = reqId + 1
    '
    '    Set r = CreateObject("Scripting.Dictionary")
    '    r("jsonrpc") = "2.0"
    '    r("method") = methName
    '    r("params") = args
    '    r("id") = reqId
    '
    '    pText = toString(r)
    '
    '    Set cli = CreateObject("MSXML2.XMLHTTP.6.0")
    '   ' Set cli = New MSXML2.XMLHTTP60
    '    If Len(user) > 0 Then   ' If Not IsMissing(user) Then
    '        cli.Open "POST", url, False, user, pwd
    '    Else
    '        cli.Open "POST", url, False
    '    End If
    '    cli.setRequestHeader "Content-Type", "application/json"
    '    cli.Send pText
    '
    '    If cli.Status <> 200 Then
    '        Err.Raise vbObjectError + INVALID_RPC_CALL + cli.Status, , cli.statusText
    '    End If
    '
    '    Set r = parse(cli.responseText)
    '    Set cli = Nothing
    '
    '    If r("id") <> reqId Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response id"
    '
    '    If r.Exists("error") Or Not r.Exists("result") Then
    '        Err.Raise vbObjectError + INVALID_RPC_CALL, , "Json-Rpc Response error: " & r("error")("message")
    '    End If
    '
    '    If Not r.Exists("result") Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response, missing result"
    '
    '    Set JsonRpcCall = r("result")
    'End Function




    Public Function toUnicode(str)

        Dim x
        Dim uStr: set uStr = New cStringBuilder
        Dim uChrCode 

        For x = 1 To Len(str)
            uChrCode = Asc(Mid(str, x, 1))
            Select Case uChrCode
            Case 8:    ' backspace
                uStr.Append "\b"
            Case 9:    ' tab
                uStr.Append "\t"
            Case 10:    ' line feed
                uStr.Append "\n"
            Case 12:    ' formfeed
                uStr.Append "\f"
            Case 13:    ' carriage return
                uStr.Append "\r"
            Case 34:    ' quote
                uStr.Append "\"""
            Case 39:    ' apostrophe
                uStr.Append "\'"
            Case 92:    ' backslash
                uStr.Append "\\"
            Case 123, 125:    ' "{" and "}"
                uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
            Case Else
              if uChrCode < 32 Or uChrCode > 127 Then    ' non-ascii characters
                uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
              Else
                uStr.Append Chr(uChrCode)
              End If
            End Select
        Next
        toUnicode = uStr.toString
        Exit Function

    End Function

    Private Sub Class_Initialize()
        psErrors = ""
    End Sub


End Class

Const URl = "http://www.nseindia.com/live_market/dynaContent/live_watch/get_quote/GetQuote.jsp?symbol=ICICIBANK"
    ' Const URl = "https://maps.googleapis.com/maps/api/geocode/json?address=sircilla,Telangana,India&key=AIzaSyDsS8jGpuQLruUrDiPSeGYZ_YqxlJoJ4YI"
    
Class Module1
    Public Sub xmlHttp()

        ' Dim xmlHttp
        ' Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        ' xmlHttp.Open "GET", URl & "&rnd=", False ' & WorksheetFunction.RandBetween(1, 99), False
        ' xmlHttp.setRequestHeader "Content-Type", "text/plain"
        ' xmlHttp.send

        ' ' Dim html
        ' ' Set html = New MSHTML.HTMLDocument
        ' MsgBox xmlHttp.ResponseText
        ' '    html.body.innerHTML = xmlHttp.ResponseText

        ' 'Dim divData
        ' 'Set divData = html.getElementById("responseDiv")
        ' '?divData.innerHTML
        ' ' Here you will get a string which is a JSON data

        ' Dim strDiv, startVal, endVal
        ' '    strDiv = divData.innerHTML
        ' '    startVal = InStr(1, strDiv, "data", vbTextCompare)
        ' '    endVal = InStr(startVal, strDiv, "]", vbTextCompare)
        ' '    strDiv = "{" & Mid(strDiv, startVal - 1, (endVal - startVal) + 2) & "}"
        
        ' strDiv = xmlHttp.ResponseText

        Dim js: set js = New JSON

        Dim p
        set p = js.parse("{\""key\"":\""value\""}")
        
        dim i: i = 1
        For Each Item In p("results")(1)
        Cells(i, 1) = Item
        Cells(i, 2) = p("results")(1)("formatted_address")
            i = i + 1
        Next
    
    End Sub


End Class

Dim mod1: set mod1 = new Module1
call mod1.xmlHttp