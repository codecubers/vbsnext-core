Dim conn    ''// As ADODB.Connection
Dim rs      ''// As ADODB.RecordSet
Dim connStr ''// As String
Dim dataDir ''// As String

dataDir = "C:\Users\nanda\git\xps.local.npm\vbspm-test\lib\csv\"                         '"
connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dataDir & ";Extended Properties=""text"""

Set conn = CreateObject("ADODB.Connection")
conn.Open(connStr)
Set rs = conn.Execute("SELECT * FROM [data.csv]")

''// do something with the recordset
WScript.Echo rs.Fields.Count & " columns found."
WScript.Echo "---"

WScript.Echo rs.Fields("Col1Name").Value
If Not rs.EOF Then
  rs.MoveNext
  WScript.Echo rs.Fields("Col3Name").Value
End If

''// explicitly closing stuff is somewhat optional
''// in this script, but consider it a good habit
rs.Close
conn.Close

Set rs = Nothing
Set conn = Nothing