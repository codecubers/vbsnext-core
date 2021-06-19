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
        m_temp = Array()
        Redim Preserve m_temp(0)
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
            Redim Preserve m_temp(Ubound(m_temp)+1)
            m_temp(Ubound(m_temp)) = Resolve(path)
            EchoDX "New Temp Base Path: %x", m_temp(Ubound(m_temp))
        End If
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

            Dim i
            i = Ubound(m_temp)
            do
                lPath = oFSO.BuildPath(m_temp(i), path)
                EchoDX "Adding Temp Base path (%x) %x to path %x. New Path: %x", Array(i, m_temp(i), path, lPath)
                If oFSO.FileExists(lPath) Then
                    EchoD "Resolved with Temp Base"
                    Resolve = oFSO.GetFile(lPath).path
                    Exit Function
                End If
                i = i - 1
            Loop While i >= 0

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