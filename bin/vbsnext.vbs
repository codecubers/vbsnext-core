

Option Explicit

Dim debug: debug = (WScript.Arguments.Named("debug") = "true")
if (debug) Then WScript.Echo "Debug is enabled"
Dim VBSNEXT_TEST_INDEX: VBSNEXT_TEST_INDEX = 1
Dim vbsnextDir: vbsnextDir=Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
Dim baseDir
With CreateObject("WScript.Shell")
    baseDir=.CurrentDirectory
End With

Dim SCRIPT_PATH
Dim CURRENT_DIRECTOR

SCRIPT_PATH = vbsnextDir
CURRENT_DIRECTOR = baseDir

Public Function startsWith(str, prefix)
    startsWith = Left(str, Len(prefix)) = prefix
End Function

Public Function endsWith(str, suffix)
    endsWith = Right(str, Len(suffix)) = suffix
End Function

Public Function contains(str, char)
    contains = (Instr(1, str, char) > 0)
End Function

Public Function argsArray()
    Dim i
    ReDim arr(WScript.Arguments.Count-1)
    For i = 0 To WScript.Arguments.Count-1
        arr(i) = """"+WScript.Arguments(i)+""""
    Next
    argsArray = arr
End Function

Public Function argsDict()
    Dim i, param, dict
    set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    ReDim arr(WScript.Arguments.Count-1)
    For i = 1 To WScript.Arguments.Count-1
        param = WScript.Arguments(i)
        If startsWith(param, "/") And contains(param, ":") Then
            param = mid(param, 2)
            WScript.Echo "param to be split: " & param
            dict.Add Lcase(split(param, ":")(0)), split(param, ":")(1)
        Else
            dict.Add i, param
        End If
    Next
    set argsDict = dict
End Function

Redim IncludedScripts(-1)
Redim ImportedScripts(-1)
Dim buildDir
Dim createBundle: createBundle = false
Dim buildBundleFile: buildBundleFile = ""	

Class Console

	Public Function fmt( str, args )
		Dim res
		res = ""

		Dim pos
		pos = 0

		Dim i
		For i = 1 To Len(str)

			If Mid(str,i,1)="%" Then
				If i<Len(str) Then

					If Mid(str,i+1,1)="%" Then
						res = res & "%"
						i = i + 1

					ElseIf Mid(str,i+1,1)="x" Then
						res = res & CStr(args(pos))
						pos = pos+1
						i = i + 1
					End If
				End If

			Else
				res = res & Mid(str,i,1)
			End If
		Next

		fmt = res
	End Function

End Class



Dim oConsole                         
set oConsole = new Console
PUblic Sub printf(str, args)

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

            WScript.Echo oConsole.fmt(str, args)
        Else

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

Class DictUtil

    Function SortDictionary(objDict, intSort)

        Const dictKey  = 1
        Const dictItem = 2

        Dim strDict()
        Dim objKey
        Dim strKey,strItem
        Dim X,Y,Z

        Z = objDict.Count

        If Z > 1 Then

            ReDim strDict(Z,2)
            X = 0

            For Each objKey In objDict
                strDict(X,dictKey)  = CStr(objKey)
                strDict(X,dictItem) = CStr(objDict(objKey))
                X = X + 1
            Next

            For X = 0 To (Z - 2)
            For Y = X To (Z - 1)
                If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
                    strKey  = strDict(X,dictKey)
                    strItem = strDict(X,dictItem)
                    strDict(X,dictKey)  = strDict(Y,dictKey)
                    strDict(X,dictItem) = strDict(Y,dictItem)
                    strDict(Y,dictKey)  = strKey
                    strDict(Y,dictItem) = strItem
                End If
            Next
            Next

            objDict.RemoveAll

            For X = 0 To (Z - 1)
            objDict.Add strDict(X,dictKey), strDict(X,dictItem)
            Next

        End If
    End Function
End Class



	Class ArrayUtil

	Public Function toString(arr)
		If Not isArray(arr) Then
			toString = "Supplied parameter is not an array."
			Exit Function
		End If

		Dim s, i
		s = "Array{" & UBound(arr) & "} [" & vbCrLf
		For i = 0  To UBound(arr)
			s = s & vbTab & "[" & i & "] => [" & arr(i) & "]"
			If i < UBound(arr) Then s = s & ", "
			s = s &  vbCrLf
		Next
		s = s & "]"
		toString = s

	End Function

	Public Function contains(arr, s) 
		If Not isArray(arr) Then
			contains = "Supplied parameter is not an array."
			Exit Function
		End If

		Dim i, bFlag
		bFlag = False
		For i = 0  To UBound(arr)
			If arr(i) = s Then
				bFlag = True
				Exit For
			End If
		Next
		contains = bFlag
	End Function

End Class



Dim arrUtil
set arrUtil = new ArrayUtil	

Class PathUtil

	Private Property Get DOT
	DOT = "."
	End Property
	Private Property Get DOTDOT
	DOTDOT = ".."
	End Property

	Private oFSO
	Private m_base
	Private m_script
	Private m_temp

	Private Sub Class_Initialize()
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		m_script = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\")-1)
		m_base = m_script
		m_temp = Array()
		ReDim Preserve m_temp(0)
		m_temp(0) = m_script
	End Sub

	Public Property Get ScriptPath
	    ScriptPath = m_script
	End Property

	Public Property Get BasePath
	    BasePath = m_base
	End Property

	Public Property Let BasePath(path)
        Do While endsWith(path, "\")
            path = Left(Path, Len(path)-1)
        Loop
        m_base = Resolve(path)
        EchoDX "New Base Path: %x", m_base
	End Property

	Public Property Get TempBasePath
	    TempBasePath = m_temp(UBound(m_temp))
	End Property

	Public Property Let TempBasePath(path)
        Do While endsWith(path, "\")
            path = Left(Path, Len(path)-1)
        Loop
        If arrUtil.contains(m_temp, path) Then
            EchoDX "Temp Path %x already exists; skipped", path
        Else
            ReDim Preserve m_temp(Ubound(m_temp)+1)
            m_temp(Ubound(m_temp)) = Resolve(path)
            EchoDX "New Temp Base Path: %x", m_temp(Ubound(m_temp))
        End If
	End Property

    Public Sub AddBasepath(path)
        TempBasePath = path
    End Sub

	Function Resolve(path)
		Dim pathBase, lPath, final
		EchoDX "path: %x", path
		If path = DOT Or path = DOTDOT Then
			path = path & "\"
		End If
		EchoDX "path: %x", path

		If oFSO.FolderExists(path) Then
			EchoD "FolderExists"
			Resolve = oFSO.GetFolder(path).path
			Exit Function
		End If

		If oFSO.FileExists(path) Then
			EchoD "FileExists"
			Resolve = oFSO.GetFile(path).path
			Exit Function
		End If

		pathBase = oFSO.BuildPath(m_base, path)
		EchoDX "Adding base %x to path %x. New Path: %x", Array(m_base, path, pathBase)

		If endsWith(pathBase, "\") Then
			If isObject(oFSO.GetFolder(pathBase)) Then
				EchoD "EndsWith '\' -> FolderExists"
				Resolve = oFSO.GetFolder(pathBase).Path
				Exit Function
			End If
		Else

			If oFSO.FolderExists(pathBase) Then
				EchoD "FolderExists"
				Resolve = oFSO.GetFolder(pathBase).path
				Exit Function
			End If

			If oFSO.FileExists(pathBase) Then
				EchoD "FileExists"
				Resolve = oFSO.GetFile(pathBase).path
				Exit Function
			End If

			Dim i
			i = Ubound(m_temp)
			Do
				lPath = oFSO.BuildPath(m_temp(i), path)
				EchoDX "Adding Temp Base path (%x) %x to path %x. New Path: %x", Array(i, m_temp(i), path, lPath)
				If oFSO.FileExists(lPath) Then
					final = oFSO.GetFile(lPath).path
					EchoDX "File Resolved with Temp Base %x", final
					Resolve = final
					Exit Function
				End If
				If oFSO.FolderExists(lPath) Then
					final = oFSO.GetFolder(lPath)
					EchoDX "Folder Resolved with Temp Base %x", final
					Resolve = final
					Exit Function
				End If
				i = i - 1
			Loop While i >= 0

			lPath = oFSO.BuildPath(m_script, path)
			EchoDX "Adding script path %x to path %x. New Path: %x", Array(m_script, path, lPath)
			If oFSO.FileExists(lPath) Then
				final = oFSO.GetFile(lPath).path
				EchoDX "File Resolved with Temp Base %x", final
				Resolve = final
				Exit Function
			End If
			If oFSO.FolderExists(lPath) Then
				final = oFSO.GetFolder(lPath)
				EchoDX "Folder Resolved with Temp Base %x", final
				Resolve = final
				Exit Function
			End If
		End If

		EchoD "Unable to Resolve"
		Resolve = path
	End Function

	Private Sub Class_Terminate()
		Set oFSO = Nothing
	End Sub

End Class



Dim putil
set putil = new PathUtil
putil.BasePath = baseDir
EchoX "Project location: %x", putil.BasePath	

Class FSO
	Private dir
	Private objFSO

	Private Sub Class_Initialize
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		dir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
	End Sub

	Public Sub setDir(s)
		dir = s
	End Sub

	Public Function getDir
		getDir = dir
	End Function

	Public Function GetFSO
		Set GetFSO = objFSO
	End Function

	Public Function FolderExists(fol)
		FolderExists = objFSO.FolderExists(fol)
	End Function

	Public Function CreateFolder(fol)
		CreateFolder = False
		If FolderExists(fol) Then
			CreateFolder = True
		Else
			objFSO.CreateFolder(fol)
			CreateFolder = FolderExists(fol)
		End If
	End Function

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

	Public Function GetFileDir(ByVal file)
		EchoDX "GetFileDir( %x )", Array(file)
		Dim objFile
		Set objFile = objFSO.GetFile(file)
		GetFileDir = objFSO.GetParentFolderName(objFile) 
	End Function

	Public Function GetFilePath(ByVal file)
		EchoDX "GetFilePath( %x )", Array(file)
		Dim objFile
		On Error Resume Next
		Set objFile = objFSO.GetFile(file)
		On Error GoTo 0
		If IsObject(objFile) Then
			GetFilePath = objFile.Path 
		Else
			EchoDX "File %x not found; searching in directory %x", Array(file,dir)
			On Error Resume Next
			Set objFile = objFile.GetFile(objFSO.BuildPath(dir, file))
			On Error GoTo 0
			If IsObject(objFile) Then
				GetFilePath = objFile.Path 
			Else
				GetFilePath = "File [" & file & "] Not found"
			End If
		End If
	End Function

	Public Function GetFileName(ByVal file)
		GetFileName = objFSO.GetFile(file).Name
	End Function

	Public Function GetFileExtn(file)
		GetFileExtn = ""
		On Error Resume Next
		GetFileExtn = LCASE(objFSO.GetExtensionName(file))
		On Error GoTo 0
	End Function

	Public Function GetBaseName(ByVal file)
		GetBaseName = Replace(GetFileName(file), "." & GetFileExtn(file), "")
	End Function

	Public Function ReadFile(file)
		file = putil.Resolve(file)
		EchoDX "---> File resolved to: %x", Array(file)
		If Not FileExists(file) Then 
			Wscript.Echo "---> File " & file & " does not exists."
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
		On Error Resume Next
		objFSO.DeleteFile(file)
		On Error GoTo 0
	End Sub

	Public Sub MoveFile(src, dest)
		On Error Resume Next
		objFSO.MoveFile src, dest
		On Error GoTo 0
	End Sub

End Class



Dim cFS
set cFS = new FSO

cFS.setDir(baseDir)

buildDir = baseDir & "\build"
If cFS.CreateFolder(buildDir) Then
createBundle = true
Else
EchoX "Unable to create build directory at [%x]. Script will not be bundled. Please try again.", buildDir
End If

Public Function log(msg)
cFS.WriteFile "build.log", msg, false
End Function

WScript.Echo "VBSNext Directory: " & vbsnextDir	

Class ClassA
    public default sub CallMe
        WScript.Echo "Class-extending resolved successfully."
    End Sub
End Class



	Class ClassB

    Private m_CLASSA

    Private Sub Class_Initialize
        set m_CLASSA = new CLASSA
    End Sub

    public default sub CallMe
        call m_CLASSA.CallMe
    End Sub
End Class



Dim ccb 
set ccb = new ClassB
ccb.CallMe

WScript.Echo "================================= Call ================================="

WScript.Echo "Base path: " & baseDir

Public Sub Import(pkg)
  WScript.Echo "Import(" + Pkg + ")"
  Include baseDir & "\node_modules\" + pkg + "\index.vbs"
End Sub

Dim sThreadBase: sThreadBase = baseDir
Public Function Include(file)
  WScript.Echo "Include(" + file + ")"
  if cFS.GetFileExtn(file) = "" Then
    WScript.Echo "File extension missing. Adding .vbs"
    file = file + ".vbs"
  end if
  Dim path

  putil.TempBasePath = sThreadBase
  path = putil.Resolve(file)
  WScript.Echo "File full path: " & path

  sThreadBase = cFS.GetFileDir(path)

  If Not arrUtil.contains(IncludedScripts, path) Then
    Redim Preserve IncludedScripts(UBound(IncludedScripts)+1)
    IncludedScripts(UBound(IncludedScripts)) = path
    Dim content: content = cFS.ReadFile(path)
    if content <> "" Then

      dim lines
      lines = split(content, vbCrLf)
      Dim includeS
      for i = 0 to ubound(lines)

        if InStr(lines(i), "Include(") > 0 Or InStr(lines(i), "Include """) > 0 Or InStr(lines(i), "Import(") > 0 or InStr(lines(i), "Import """) > 0 Then
          includeS = includeS & lines(i) & vbCrLf
        end if
      next

      if includeS <> "" Then
          ExecuteGlobal includeS
      End If
    Else
      WScript.Echo "File content is empty. Not loaded."
    End If
  Else
    WScript.Echo "File: " & path & " already loaded."
  End If
  Include = Include
End Function

WScript.Echo "Execution Started for file"

Dim file
file = WScript.Arguments.Named("file")
If file = "" Then
    WScript.Echo "Script file not provided as a named argument [/file:]"
    if Wscript.Arguments.count > 0 then
      file = Wscript.Arguments(0) 
      if file = "" Then
        WScript.Echo "No file argument provided."
        Wscript.Quit
      End If
    else 
      file = "index.vbs"
    end if
End If

file = baseDir & "\" & file

if cFS.GetFileExtn(file) = "" Then
  WScript.Echo "File extension missing. Adding .vbs"
  file = file + ".vbs"
end if

WScript.Echo "Main Script: " & file
buildBundleFile = buildDir & "\" & cFS.GetBaseName(file) &  "-bundle-unresolved.vbs"
WScript.Echo "buildBundleFile: " & buildBundleFile

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

BundleScript vbsnextDir & "\vbsnext-build.vbs", true

Include file

Dim i, core
for i = UBound(IncludedScripts) to 0 step -1
    core = cFS.ReadFile(IncludedScripts(i))
    core = Replace(core, "Option Explicit", "")
    core = vbCrLf & vbCrLf & "'================= File: " & IncludedScripts(i) & " =================" & vbCrLf & core
    BundleScriptStr core, false
next