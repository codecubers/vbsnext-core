if WScript.Arguments.Count < 3 then
    WScript.Echo "Missing Parameters. Execution halted"
	WScript.Quit
end if

on error Resume Next

Dim objrec, objcnn, sOdbcDSN, strMacStatus, strProfileId, strDatafile, fso, oFile,iStartRow
const adStateClosed  = 0
const adStateOpen  = 1
dim sDQ: sDQ = chr(34)

dtmStartTime = Timer

sUsername = WScript.Arguments(0)
sPassword = WScript.Arguments(1)
sQuery = WScript.Arguments(2)

'sUsername = "db2admin"
'sPassword = "db2admin"
'sQuery = "select * from db2uvd2.tuv8001 with ur;"
'sQuery = "select * from ACRT.FUND with ur;"

	Set objrec = createobject("adodb.recordset")
	Set objcnn = createobject("adodb.connection")

'Wscript.Echo "Objects are initiated" & vbcrlf
	
	'sOdbcDSN = "DSN=xx_d5csdb;Uid=" & sUsername & ";Pwd=" & sPassword & ";"
	
	sOdbcDSN = "Driver={IBM DB2 ODBC DRIVER};Database=" & "acrwd1d1" & _
	";HOSTNAME=" & "dbtoolbox" & _
	";PROTOCOL=TCPIP;PORT=" & "50004" & _
	";uid=" & sUsername & _
	";pwd=" & sPassword & ";CurrentSchema=" & "ACRT" & ";"
	
	'sOdbcDSN = "Driver={IBM DB2 ODBC Driver};Database=" & "TOM00D0" & _
	'";HOSTNAME=" & "mlisdb2sysgw1.manulife.com" & _
	'";PROTOCOL=TCPIP;PORT=" & "3700" & _
	'";uid=" & sUsername & _
	'";pwd=" & sPassword
	
	objcnn.open sOdbcDSN

If Err.Number <> 0 Then	
	Wscript.Echo "There was an internal error (" & Err.Number & ") occured:" & vbcrlf & Err.Description
end if

'Wscript.Echo "Connection opened for :" & sOdbcDSN & vbcrlf
'Wscript.Echo "Connection status: " & objcnn.State
	
	objrec.open sQuery, objcnn 

'Wscript.Echo "Recordset opened" & vbcrlf
	
	sResponse = "{" & sDQ & "Query" &  sDQ  & ":" & sDQ & sQuery & sDQ & "," & _
					  sDQ &	"Records" & sDQ & ":["
	while objrec.EOF <> true and objrec.BOF <> true
		sThisRow = sThisRow & "{"
		sThisItem = ""
		
		for col = 0 to objrec.Fields.Count - 1
			if isNull(objrec.Fields(col)) or len(trim(objrec.Fields(col))) = 0 or objrec.Fields(col) = vbTab then 
				sValue = "--"
			else
				sValue = objrec.Fields(col)
			end if
			
			'Wscript.Echo objrec.Fields(col).Name & "_" & Asc(sValue) & vbcrlf
			sThisItem = sThisItem & sDQ & objrec.Fields(col).Name & sDQ & ":" & sDQ & sValue & sDQ & ","
		next
		
		sThisItem = left(sThisItem, len(sThisItem) - 1)
		sThisRow = sThisRow & sThisItem & "},"
		
		objrec.MoveNext
	wend
	sThisRow = left(sThisRow, len(sThisRow) - 1)
	sResponse = sResponse & sThisRow & "]," & sDQ & "Time" & sDQ & ":" & sDQ & Round(Timer - dtmStartTime, 3) & sDQ & "}"

'Wscript.Echo "Data displayed" & vbcrlf
	
	objcnn.close

'Wscript.Echo "Object is now closed" & vbcrlf

	Set objrec = Nothing
	Set objcnn = Nothing

'Wscript.Echo "objects being killed" & vbcrlf

If Err.Number <> 0 Then	
	Wscript.Echo "There was an internal error (" & Err.Number & ") occured:" & vbcrlf & Err.Description
	
	if Not objrec is nothing then set objrec = nothing
	if Not objcnn is nothing then
		if objcnn.State <> adStateClosed then 
			Wscript.Echo "Closing connection... "
			objcnn.Close
			if objcnn.State = adStateClosed then Wscript.Echo "Closed"
		end if
		set objcnn = nothing
	end if
End if

Wscript.Echo sResponse