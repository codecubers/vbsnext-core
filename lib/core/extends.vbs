'TODO: Bug to be fixed; extends in string literals causing stirng to break
Class ClassA
    public default sub CallMe
        WScript.Echo "Class-extending resolved successfully."
    End Sub
End Class

Class ClassB extends ClassA
    
End Class

Dim ccb 
set ccb = new ClassB
ccb.CallMe