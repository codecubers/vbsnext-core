' Dependencies:
' Class(es): FS
Class WShell
		
	' 0 Hides the window and activates another window.
	Public Property Get WShell_WINDOW_MODE_HIDE
	WShell_WINDOW_MODE_HIDE = 0
	End Property
	' 1 Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
	Public Property Get WShell_WINDOW_MODE_ORIGINAL
	WShell_WINDOW_MODE_ORIGINAL = 1
	End Property
	' 2 Activates the window and displays it as a minimized window.
	Public Property Get WShell_WINDOW_MODE_MINIMIZED_ACTIVE
	WShell_WINDOW_MODE_MINIMIZED_ACTIVE = 2
	End Property
	' 3 Activates the window and displays it as a maximized window.
	Public Property Get WShell_WINDOW_MODE_MAXIMIZED_ACTIVE
	WShell_WINDOW_MODE_MAXIMIZED_ACTIVE = 3
	End Property
	' 4 Displays a window in its most recent size and position. The active window remains active.
	Public Property Get WShell_WINDOW_MODE_RECENT
	WShell_WINDOW_MODE_RECENT = 4
	End Property
	' 5 Activates the window and displays it in its current size and position.
	Public Property Get WShell_WINDOW_MODE_CURRENT
	WShell_WINDOW_MODE_CURRENT = 5
	End Property
	' 6 Minimizes the specified window and activates the next top-level window in the Z order.
	Public Property Get WShell_WINDOW_MODE_MINIMIZED_NEXT
	WShell_WINDOW_MODE_MINIMIZED_NEXT = 6
	End Property
	' 7 Displays the window as a minimized window. The active window remains active.
	Public Property Get WShell_WINDOW_MODE_MINIMIZED_INACTIVE
	WShell_WINDOW_MODE_MINIMIZED_INACTIVE = 7
	End Property
	' 8 Displays the window in its current state. The active window remains active.
	Public Property Get WShell_WINDOW_MODE_CURRENT_INACTIVE
	WShell_WINDOW_MODE_CURRENT_INACTIVE = 8
	End Property
	' 9 Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
	Public Property Get WShell_WINDOW_MODE_NINE
	WShell_WINDOW_MODE_NINE = 9
	End Property
	' 10 Sets the show-state based on the state of the program that started the application.
	Public Property Get WShell_WINDOW_MODE_SHOW_STATE
	WShell_WINDOW_MODE_SHOW_STATE = 10
	End Property
	
	
	' Command output print options
	Public Property Get PRINT_STDOUT
	PRINT_STDOUT = False
	End Property
	Public Property Get PRINT_ECHO
	PRINT_ECHO = True
	End Property
	Public Property Get PRINT_MSGBOX
	PRINT_MSGBOX = False
	End Property
	
	
	' Private dir
	Private oThis
	
	Private Sub Class_Initialize
		Set oThis = WScript.CreateObject("WScript.Shell")
		
		' Set execution directory
		dir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
	End Sub
	
	' Update the current directory of the instance if needed
	Public Sub setDir(s)
		dir = s
	End Sub
	
	Public Property Get GetObj
	Set GetObj = oThis
	End Property
	
	' ================== Sub Routines ==================
	
	Public Sub Run(ByVal path) 
		oThis.Run path, WShell_WINDOW_MODE_ORIGINAL, True
	End Sub
	
	Public Sub Exec(ByVal cmd) 
		Wscript.Echo cmd
		Set result = oThis.Exec(strPath)
		print result
	End Sub
	
	Private Sub print(execCmdOut)
		Select Case execCmdOut.Status
			Case WshFinished
			strOutput = execCmdOut.StdOut.ReadAll
			Case WshFailed
			strOutput = execCmdOut.StdErr.ReadAll
		End Select
		
		If PRINT_STDOUT Then WScript.StdOut.Write strOutput  'write results to the command line
		If PRINT_ECHO Then Wscript.Echo strOutput          'write results to default output
		If PRINT_MSGBOX Then MsgBox strOutput                'write results in a message box
	End Sub
	
	Public Sub Ping(ip)
		Const WshFinished = 1
		Const WshFailed = 2
		strCommand = "ping.exe " & ip
		
		Set WshShellExec = oThis.Exec(strCommand)
		print WshShellExec
	End Sub
	
	Public Sub PingMe
		Ping "127.0.0.1"
	End Sub
	
	' ================== Function Routines ==================
	
	Public Function OpenTextFile(ByVal path)
		oThis.Run "%windir%\notepad " & path, WShell_WINDOW_MODE_ORIGINAL, True
	End Function
	
End Class
