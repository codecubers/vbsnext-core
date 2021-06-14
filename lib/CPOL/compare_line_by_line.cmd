REM Comparing Two Files Line by Line
REM This is a small Windows batch utility written in VBscript to read and compare two -
REM text files line by line. It can read multiple files if folders are provided. 
REM It also supports sorting (.NET required) prior to comparison.
REM 
REM Author: Marijan Nikic
REM License: The Code Project Open License (CPOL)
REM https://www.codeproject.com/Tips/5286930/Comparing-Two-Files-Line-by-Line

set "caption=Compare line by line"
title=%caption%
pushd "%~dp0"
@echo off
cls
if EXIST "C:\Windows\SysWOW64\" (set "cscript=C:\Windows\SysWOW64\cscript.exe") ELSE (set cscript=cscript.exe)
echo:

set "psort=Y"

%cscript% //nologo "%~f0?.wsf" %psort% //job:VBS

pause

exit

<package><job id="VBS"><script language="VBScript">

Dim sort: sort = False
If WScript.Arguments(0) = "Y" Then sort = True

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh: Set sh = CreateObject("Wscript.Shell")
Dim file1
Dim file2
Dim folder1
Dim folder2
Dim cnt
Dim errcnt

Dim path1
Dim path2

WScript.stdout.Write "Path 1: "
path1 = WScript.stdin.ReadLine
wscript.echo
WScript.stdout.Write "Path 2: "
path2 = WScript.stdin.ReadLine

' is relative path?
If Replace(path1,":","") = path1 Then
	If Left(path1,1) <> "\" Then path1 = "\" & path1
	path1 = sh.CurrentDirectory & path1
End If
If Replace(path2,":","") = path2 Then
	If Left(path2,1) <> "\" Then path2 = "\" & path2
	path2 = sh.CurrentDirectory & path2
End If

If fso.FolderExists(path1) And fso.FolderExists(path2) Then ' folders
	Set folder1 = fso.GetFolder(path1)
	Set folder2 = fso.GetFolder(path2)
	For Each fil In folder1.Files
		CompareFiles fil.Path, Replace(fil.Path, path1, path2), sort
	Next
ElseIf fso.FileExists(path1) And fso.FileExists(path2) Then ' files
	CompareFiles path1, path2, sort
Else
	WScript.Echo "Invalid path(s)!"
End If

WScript.Quit

Sub CompareFiles(fil1, fil2, sort)
	WScript.echo
	WScript.echo "Comparing files:"
	WScript.echo fil1
	WScript.echo fil2
	WScript.echo "==================================================="
	WScript.stdout.Write "Press [ENTER] to continue..."
	WScript.stdin.readline
	cnt = 1
	errcnt = 0
	If Not fso.FileExists(fil1) Then
		WSCript.echo "File " & fil1 & " does not exist!"
	ElseIf Not fso.FileExists(fil2) Then
		WScript.echo "File " & fil2 & " does not exist!"
	Else
		Set file1=fso.OpenTextFile(fil1)
		Set file2=fso.OpenTextFile(fil2)
		If sort Then
			Set arrlist1 = CreateObject("System.Collections.ArrayList")
			Do Until file1.AtEndOfStream
				arrlist1.Add file1.ReadLine
			Loop
			Set arrlist2 = CreateObject("System.Collections.ArrayList")
			Do Until file2.AtEndOfStream
				arrlist2.Add file2.ReadLine
			Loop
			arrlist1.Sort
			arrlist2.Sort
			Do: For i = 0 To Min(arrlist1.Count - 1, arrlist2.Count - 1)
				If i = arrlist1.Count - 1 And i <> arrlist2.Count - 1 Then
					WScript.echo fil1 & " at end of stream: " & CStr(i) & " lines. Other file has more lines"
					errcnt = errcnt + 1
					Exit Do
				ElseIf i <> arrlist1.Count - 1 And i = arrlist2.Count - 1 Then
					WScript.echo fil2 & " at end of stream: " & CStr(i) & " lines. Other file has more lines"
					errcnt = errcnt + 1
					Exit Do
				End If
				If arrlist1.Item(i) <> arrlist2.Item(i) Then
					WScript.echo "Line " & CStr(i+1) & " mismatch"
					errcnt = errcnt + 1
				End If
			Next: Loop While 0=1
		Else
			Do Until file1.AtEndOfStream And file2.AtEndOfStream
				If file1.AtEndOfStream And Not file2.AtEndOfStream Then
					WScript.echo fil1 & " at end of stream: " & CStr(cnt-1) & " lines. Other file has more lines"
					errcnt = errcnt + 1
					Exit Do
				ElseIf file2.AtEndOfStream And Not file1.AtEndOfStream Then
					WScript.echo fil2 & " at end of stream: " & CStr(cnt-1) & " lines. Other file has more lines"
					errcnt = errcnt + 1
					Exit Do
				End If
				If file1.readline <> file2.readline Then
					WScript.echo "Line " & CStr(cnt) & " mismatch"
					errcnt = errcnt + 1
				End If
				cnt = cnt + 1
			Loop
		End If
		If errcnt = 0 Then WScript.echo "Files match!"
		file1.close
		file2.close
	End If
End Sub

function Min(a,b)
    Min = a
    If b < a then Min = b
end function

</script></job></package>