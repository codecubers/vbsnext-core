

Class Collection

    Private dict
    Private oThis
    Private m_Name

    Private Sub Class_Initialize()
        set dict = CreateObject("Scripting.Dictionary")
        set oThis = Me
        m_Name = "Undefined"
    End Sub

    Public Default Property Get Obj
        set Obj = dict
    End Property 

    Public Property Get Name
        Name = m_Name
    End Property
    Public Property Let Name(Value)
        m_Name = Value
    End Property

    Public Sub Add(Key, Value)
        dict.Add key, value
    End Sub ' Add

    Public Sub Remove(Key)
        If KeyExists(Key) Then
            dict.Remove(Key)
        Else
            RaiseErr "Key [" & Key & "] does not exists in collection."
        End If
    End Sub ' Remove

    Public Sub RemoveAll()
        dict.RemoveAll()
    End Sub

    Public Property Get Count
        Count = dict.Count
    End Property

    Public Function GetItem(Key)
        If KeyExists(Key) Then
            GetItem = dict.Item(Key)
        Else
            'TODO: Should we raise an error?
            RaiseErr "Key [" & Key & "] does not exists in collection."
        End If
    End Function ' GetItem

    Public Function GetItemAtIndex(Index)
        'TODO: How to ensure Index is an integer?
        GetItemAtIndex = dict.Item(Index)
    End Function ' GetItemAtIndex


    Public Function IndexOf(Key)
        IndexOf = dict.IndexOf(Key, 0)
    End Function

    Public Function KeyExists(Key)
        KeyExists = dict.Exists(Key)
    End Function

    Public Function toCSV
        toCSV = join(toArray(), ", ")
    End Function

    Public Function toArray
        toArray = dict.Items
    End Function

    Public Function isEmpty
        isEmpty = (dict.Count = 0)        
    End Function ' isEmpty
    
    ' Public Sub ReverseKeys

    '     Dim i, j, last, half, temp, arr
    '     arr = dict.Keys
    '     WScript.Echo join(arr, ", ")
    '     last = UBound(arr)
    '     half = Int(last/2)

    '     For i = 0 To half
    '         temp = arr(i)
    '         arr(i) = arr(last-i)
    '         arr(last-i) = temp
    '     Next
        
    '     WScript.Echo join(arr, ", ")

    '     dim dict1
    '     set dict1 = New Collection
    '     for i = 0 to UBound(arr)
    '         dict1.Add arr(i), arr(i) & "1"
    '     next
    '     'WScript.Echo dict1.toCSV
    '     ' RemoveAll
    '     set dict = dict1

    ' End Sub

    ' list.Sort
    ' list.Reverse
        ' for each k in dict.Keys
        '     if k = Key Then
        '         KeyExists = true
        '         Exit Function
        '     End If
        ' next


    Private Sub RaiseErr(desc)
        Err.Clear
        Err.Raise 1000, "Collection Class Error", desc
    End Sub

    Private Sub Class_Terminate()
        set dict = Nothing
        set oThis = Nothing
    End Sub

End Class ' Collection