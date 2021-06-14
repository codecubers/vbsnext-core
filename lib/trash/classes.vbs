Class SomeObject
	Private Sub Class_Initialize(  )
		'This event is called when an instance of the class is instantiated
		'Initialize properties here and perform other start-up tasks
	End Sub
	
	Private Sub Class_Terminate(  )
		'This event is called when a class instance is destroyed
		'either explicitly (Set objClassInstance = Nothing) or
		'implicitly (it goes out of scope)
	End Sub
End Class


Class Information
	'Create a private property to hold the phone number
	Private strPhoneNumber
	
	Public Property Let PhoneNumber(strPhone)
	'Ensures that strPhone is in the format ###-###-####
	'If it is not, raise an error
	If IsObject(strPhone) then
		Err.Raise vbObjectError + 1000, "Information Class", _
		"Invalid format for PhoneNumber.  Must be in ###-###-#### format."
		Exit property
	End If
	
	Dim objRegExp
	Set objRegExp = New regexp
	
	objRegExp.Pattern = "^\d{3}-\d{3}-\d{4}$"
	
	'Make sure the pattern fits
	If objRegExp.Test(strPhone) then
		strPhoneNumber = strPhone
	Else
		Err.Raise vbObjectError + 1000, "Information Class", _
		"Invalid format for PhoneNumber.  Must be in ###-###-#### format."
	End If    
	
	Set objRegExp = Nothing
	End Property
	
	Public Property Get PhoneNumber(  )
	PhoneNumber = strPhoneNumber
	End Property
End Class


Class MyConnectionClass
	'Create a private property to hold our connection object
	Private objConn
	
	Public Property Get Connection(  )
	Set Connection = objConn
	End Property
	Public Property Set Connection(objConnection)
	'Assign the private property objConn to objConnection
	Set objConn = objConnection
	End Property
End Class


Class Test
	Private m_s
	Private m_i
	
	Public Default Function Init(parameters)
		Select Case UBound(parameters)
			Case 0
			Set Init = InitOneParam(parameters(0))
			Case 1
			Set Init = InitTwoParam(parameters(0), parameters(1))
			Case Else
			Set Init = Me
		End Select
    End Function

    Private Function InitOneParam(parameter1)
        If TypeName(parameter1) = "String" Then
            m_s = parameter1
        Else
            m_i = parameter1
        End If
        Set InitOneParam = Me
    End Function

    Private Function InitTwoParam(parameter1, parameter2)
        m_s = parameter1
        m_i = parameter2
        Set InitTwoParam = Me
    End Function
End Class

' Dim o : Set o = (New Test)(Array())
' Dim o : Set o = (New Test)(Array("Hello World"))
' Dim o : Set o = (New Test)(Array(1024))
' Dim o : Set o = (New Test)(Array("Hello World", 1024))