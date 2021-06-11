Public Const WINDOW_MODE_HIDE = 0
Class WShell

    Private WshShell
    
    Private Sub Class_Initialize
        
        Set WshShell = WScript.CreateObject("WScript.Shell")
        'Hides the window and activates another window.
        'TODO: Add all constants of WShell

' 1

' Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.

' 2

' Activates the window and displays it as a minimized window.

' 3

' Activates the window and displays it as a maximized window.

' 4

' Displays a window in its most recent size and position. The active window remains active.

' 5

' Activates the window and displays it in its current size and position.

' 6

' Minimizes the specified window and activates the next top-level window in the Z order.

' 7

' Displays the window as a minimized window. The active window remains active.

' 8

' Displays the window in its current state. The active window remains active.

' 9

' Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.

' 10

' Sets the show-state based on the state of the program that started the application.
    End Sub

    public Function OpenTextFile(ByVal path)
        WshShell.Run "%windir%\notepad " & path
    End Function
End Class