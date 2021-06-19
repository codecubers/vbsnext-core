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
			toString = "Supplied parameter is not an array."
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
	
	'TODO: Add functionality to manage Array (redim, get last, add new etc.,)
	'TODO: With ability to sort, reverse, avoid duplicates etc.,
End Class