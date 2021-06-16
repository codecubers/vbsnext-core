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