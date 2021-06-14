Dim oFS : Set oFS = CreateObject( "Scripting.FileSystemObject" )
Dim sPFSpec : sPFSpec = ".\abc.properties"
Dim dicProps : Set dicProps = CreateObject( "Scripting.Dictionary" )
Dim oTS : Set oTS = oFS.OpenTextFile( sPFSpec )
Dim sSect : sSect = ""
Do Until oTS.AtEndOfStream
Dim sLine : sLine = Trim( oTS.ReadLine )
If "" <> sLine Then
If "#" = Left( sLine, 1 ) Then
sSect = sLine
WScript.Echo oTS.Line, "starting section", sSect
Else
If "" = sSect Then
WScript.Echo oTS.Line, "no section", sLine
Else
Dim aParts : aParts = Split( sLine, "=" )
If 1 <> UBound( aParts ) Then
WScript.Echo oTS.Line, "bad property line", sLine
Else
dicProps( sSect & "." & Trim( aParts( 0 ) ) ) = Trim( aParts( 1 ) )
WScript.Echo oTS.Line, "good property line", sSect, sLine
End If
End If
End If
End If
Loop
oTS.Close

Dim sKey
For Each sKey In dicProps
WScript.Echo sKey, "=>", dicProps( sKey )
Next

sKey = "#local.local.db.database"
WScript.Echo dicProps( sKey )