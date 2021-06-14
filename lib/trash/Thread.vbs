' https://stackoverflow.com/questions/4610686/vbscript-threading

Class Thread
' Usage:
'    Dim x: Set x = New Thread
'    Call x.init(10)
'    Call x.queue("script.bat", Array("Arg1", "output_file.txt"))
'    Call x.queue("cscript.exe prog.vbs", Array("Arg1", "Arg2", "Arg3"))
'    Call x.setMax(20)
''''''''''''''''''''''''''''''''''''''''''''''
    Private p_threads
    Private p_max

    Private Function spawn(action, args)
        Dim wsh: Set wsh = WScript.CreateObject("WScript.Shell")
        Dim command: command = action
        Dim element
        For Each element In args
            command = command & " " & element
        Next
        spawn = wsh.Run(command, 0, False)
        Set wsh = Nothing
    End Function

    Public Sub queue(action, args)
        If Ubound(p_threads,1) < p_max Then
            ' create new thread
            ReDim Preserve p_threads(Ubound(p_threads, 1)+1)
            p_threads(Ubound(p_threads, 1)) = spawn(action, args)
        Else
            ' recycle old thread
            Do
            Dim i
            For i = 1 To Ubound(p_threads, 1)
                ' find a thread which has finished
                If p_threads(i) = 1 Then
                    p_threads(i) = spawn(action, args)
                    Exit Sub
                End If
            Next
            ' wait for a thread to finish
            WScript.Sleep 300
            Loop Until False
        End If
    End Sub

    Public Sub init(n)
        p_threads = Array(1)
        p_max = n
    End Sub

    Public Property Let setMax(n)
        p_max = n
    End Property
End Class
