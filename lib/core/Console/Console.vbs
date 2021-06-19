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
	Public Function fmt( str, args )
		Dim res		' the result string.
		res = ""
		
		Dim pos		' the current position in the args array.
		pos = 0
		
		Dim i
		For i = 1 To Len(str)
			' found a fmt char.
			If Mid(str,i,1)="%" Then
				If i<Len(str) Then
					' normal percent.
					If Mid(str,i+1,1)="%" Then
						res = res & "%"
						i = i + 1
						
						' expand from array.
					ElseIf Mid(str,i+1,1)="x" Then
						res = res & CStr(args(pos))
						pos = pos+1
						i = i + 1
					End If
				End If
				
				' found a normal char.
			Else
				res = res & Mid(str,i,1)
			End If
		Next
		
		fmt = res
	End Function
	
End Class