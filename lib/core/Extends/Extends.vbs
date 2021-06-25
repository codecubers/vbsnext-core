DIm i:Class CAR 'extends
    Private m_tires
    Public Property Let      Tires(count)
        m_tires = count
        i = "'" & _
        "a"
        j = """:""":k=1
        x="" + "Msgbox ""s:""":i=1 : m=1
    End Property
    Public Property Get Tires
        Tires = m_tires
    End Property 
End Class
'' cpmments
     Class HONDA 'extends CAR

    Private m_brand
    Public Property Let Brand(b)
        m_brand = b
    End Property
    Public Property Get Brand
        Brand = m_brand
    End Property 
End Class

Class MARUTHI
End Class

Dim v
set v = new CAR
v.Tires = 4
WScript.Echo v.Tires
