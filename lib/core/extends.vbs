Class ClassA
    public default sub CallMe
        WScript.Echo "I'm in ClassA"
    End Sub
End Class

Class ClassB extends ClassA
    
End Class

Dim ccb 
set ccb = new ClassB
ccb.CallMe