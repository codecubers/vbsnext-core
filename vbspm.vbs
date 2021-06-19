


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

' ================= inline ================= 

Dim debug: debug = (WScript.Arguments.Named("debug") = "true")
if (debug) Then WScript.Echo "Debug is enabled"
Dim vbspmDir
Dim baseDir
Dim cFS
Redim IncludedScripts(-1)
Dim arrUtil
Dim buildDir
Dim createBundle: createBundle = false
Dim buildBundleFile: buildBundleFile = ""
Dim oConsole
Dim putil

With CreateObject("WScript.Shell")
baseDir=.CurrentDirectory
End With

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

' ================= src : lib/core/ArrayUtil/ArrayUtil.vbs ================= 
Class ArrayUtil

    Public Function toString(arr)
        If Not isArray(arr) Then
            toString = "Supplied parameter is not an array."
            Exit Function
        End If

        Dim s, i
        s = "Array{" & UBound(arr) & "} [" & vbCrLf
        For i = 0  to UBound(arr)
            s = s & vbTab & "[" & i & "] => [" & arr(i) & "]"
            if i < UBound(arr) Then s = s & ", "
            s = s &  vbCrLf
        Next
        s = s & "]"
        toString = s

    End Function

    Public Function contains(arr, s) 
        If Not isArray(arr) Then
            toString = "Supplied parameter is not an array."
            Exit Function
        End If

        Dim i, bFlag
        bFlag = false
        For i = 0  to UBound(arr)
            If arr(i) = s Then
                bFlag = true
                Exit For
            End If
        Next
        contains = bFlag
    End Function

End Class
' ================= inline ================= 

set arrUtil = new ArrayUtil

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

	Public Function getDir
		getDir = dir
	End Function
	
	Public Function GetFSO
		Set GetFSO = objFSO
	End Function

  Public Function FolderExists(fol)
    FolderExists = objFSO.FolderExists(fol)
  End Function
    ' ===================== Sub Routines =====================


	Public Function CreateFolder(fol)
    CreateFolder = false
		If FolderExists(fol) Then
      CreateFolder = true
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

	' ===================== Function Routines =====================

	Public Function GetFileDir(ByVal file)
    debugf "GetFileDir( %s )", Array(file)
		Dim objFile
		Set objFile = objFSO.GetFile(file)
		GetFileDir = objFSO.GetParentFolderName(objFile) 
	End Function
	
	Public Function GetFilePath(ByVal file)
    debugf "GetFilePath( %s )", Array(file)
    Dim objFile
    On Error Resume Next
    set objFile = objFSO.GetFile(file)
    On Error Goto 0
    If IsObject(objFile) Then
		  GetFilePath = objFile.Path 
    Else
      debugf "File %s not found; searching in directory %s", Array(file,dir)
      On Error Resume Next
      set objFile = objFile.GetFile(objFSO.BuildPath(dir, file))
      On Error Goto 0
      If IsObject(objFile) Then
		    GetFilePath = objFile.Path 
      Else
        GetFilePath = "File [" & file & "] Not found"
      End If
    End If
	End Function

  ''' <summary>Returns a specified number of characters from a string.</summary>
  ''' <param name="file">File Name</param>
	Public Function GetFileName(ByVal file)
		GetFileName = objFSO.GetFile(file).Name
	End Function

	Public Function GetFileExtn(file)
		GetFileExtn = ""
		on Error Resume Next
		GetFileExtn = LCASE(objFSO.GetExtensionName(file))
		On Error goto 0
	End Function

  Public Function GetBaseName(ByVal file)
    GetBaseName = Replace(GetFileName(file), "." & GetFileExtn(file), "")
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

vbspmDir = cFS.GetFileDir(WScript.ScriptFullName)
log "VBSPM Directory: " & vbspmDir


' ================= src : lib/core/PathUtil/PathUtil.vbs ================= 
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
        set oFSO = CreateObject("Scripting.FileSystemObject")
        m_script = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\")-1)
        m_base = m_script
        m_temp = m_script
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
        TempBasePath = m_temp
    End Property
    Public Property Let TempBasePath(path)
        Do While endsWith(path, "\")
            path = Left(Path, Len(path)-1)
        Loop
        m_temp = Resolve(path)
        EchoDX "New Temp Base Path: %x", m_temp
    End Property

    Function Resolve(path)
        Dim pathBase, lPath
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

            lPath = oFSO.BuildPath(m_temp, path)
            EchoDX "Adding Temp Base path %x to path %x. New Path: %x", Array(m_temp, path, lPath)
            If oFSO.FileExists(lPath) Then
                EchoD "Resolved with Temp Base"
                Resolve = oFSO.GetFile(lPath).path
                Exit Function
            End If

            lPath = oFSO.BuildPath(m_script, path)
            EchoDX "Adding script path %x to path %x. New Path: %x", Array(m_script, path, lPath)
            If oFSO.FileExists(lPath) Then
                EchoD "Resolved with script base"
                Resolve = oFSO.GetFile(lPath).path
                Exit Function
            End If
        End If
        
        EchoD "Unable to Resolve"
        Resolve = path
    End Function ' Resolve


    Private Sub Class_Terminate()
        set oFSO = nothing
    End Sub

End Class ' PathUtil
' ================= inline ================= 

set putil = new PathUtil
putil.BasePath = baseDir
EchoX "Project location: %x", putil.BasePath

' ================= src : lib/core/globals.vbs ================= 
log "================================= Call ================================="

log "Base path: " & baseDir

Public Sub Import(pkg)
  log "Import(" + Pkg + ")"
  Include baseDir & "\node_modules\" + pkg + "\index.vbs"
End Sub
' ================= src : lib/core/include-run.vbs ================= 
' Dim iThread: iThread = 1
' Public Function Thread(i)
'     EchoX "Thread %x", i
'     i = i + 1
'     Thread = i
' End Function
Dim sThreadBase: sThreadBase = baseDir
Public Function Include(file)
  log "Include(" + file + ")"
  if cFS.GetFileExtn(file) = "" Then
    log "File extension missing. Adding .vbs"
    file = file + ".vbs"
  end if
  Dim path
  'path = cFS.GetFilePath(file)
  putil.TempBasePath = sThreadBase
  path = putil.Resolve(file)
  log "File full path: " & path
  'cFS.setDir(cFS.GetFileDir(path))
  sThreadBase = cFS.GetFileDir(path)
  
  If Not arrUtil.contains(IncludedScripts, path) Then
    Redim Preserve IncludedScripts(UBound(IncludedScripts)+1)
    IncludedScripts(UBound(IncludedScripts)) = path
    Dim content: content = cFS.ReadFile(path)
    if content <> "" Then 
      'cFS.WriteFile "build\bundle.vbs", content, false
      'EchoX "File: %x", file
      'EchoX "Thread ---> %x", iThread
      'content = "iThread = Thread(iThread)" & VBCRLF & content
      'EchoX "Content: %x", content
      ExecuteGlobal content
    Else
      log "File content is empty. Not loaded."
    End If
  Else
    log "File: " & path & " already loaded."
  End If
  Include = Include
End Function
' ================= src : lib/core/params.vbs ================= 
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

if cFS.GetFileExtn(file) = "" Then
  log "File extension missing. Adding .vbs"
  file = file + ".vbs"
end if

log "Main Script: " & file
buildBundleFile = buildDir & "\" & cFS.GetBaseName(file) &  "-bundle.vbs"
log "buildBundleFile: " & buildBundleFile


' ================= src : lib/core/bundler.vbs ================= 
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


' Just before start writing Include/Import file contents to the builder,
' Write the vbspm.vbs file contents
BundleScript vbspmDir & "\vbspm-build.vbs", true

'===========================
Include file
'===========================

' Wscript.Echo arrUtil.toString(IncludedScripts)
Dim i, core
for i = UBound(IncludedScripts) to 0 step -1
    core = cFS.ReadFile(IncludedScripts(i))
    core = Replace(core, "Option Explicit", "")
    core = vbCrLf & vbCrLf & "'================= File: " & IncludedScripts(i) & " =================" & vbCrLf & core
    BundleScriptStr core, false
next