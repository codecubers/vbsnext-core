	Server_Details = "Database=d3csdb;hostname=jhrpdbs1;port=60008"
    Login_Creds = "uid=ezcs_ro;pwd=ezcs_ro11"
	Rng_SQL = "Select contract_id, " & _
				"Case MANULIFE_COMPANY_ID  " & _
				"When '019' then 'US'  " & _
				"When '094' then 'NY'  " & _
				"End as Company " & _
				"From psw100.contract_cs " & _
				"Where CONTRACT_ID in ($$Contract_IDs$$) ;"
	'Rng_SQL = "Select * from stp100.submission_log_event;"
	'===========================

'Server	Port	Database	UserName	Password
'jhrpsdbs3	60008	d2csdb	ezcs_ro	ezcs_ro11
'jhrpdbs1	57100	macsdb	ezcs_ro	ezcs_ro11
'dbtoolbox	50004	acrwd1d1	db2admin	db2admin
'dbtoolbox	50004	acrmisd1	db2admin	db2admin
'dbtoolbox	50004	mrldm	db2admin	db2admin
'jhrpsdbs3	60073	q7csdb	ezcs_ro	ezcs_ro11
				
	Rng_SQL  = Replace(Rng_SQL, "$$Contract_IDs$$","70300")
    'DBCONSRT = "Driver={IBM DB2 ODBC DRIVER};" & Server_Details & ";DB2COMM=TCPIP;" & Login_Creds
    DBCONSRT = "Provider=IBMDADB2;" & Server_Details & ";Protocol=TCPIP;" & Login_Creds
    'DBCONSRT = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=d2csdb"
    'CHANGE THE BELOW QUERY STRING ACCORDING TO YOUR NEED
    'QRYSTR = "select * from PSW100.Contractfunds fetch first 10 rows only"
    Set DBCON = CreateObject("ADODB.Connection")
    DBCON.ConnectionString = DBCONSRT
'On Error GoTo ConnClose
    DBCON.Open
    
	'BELOW CODE USED TO GET THE DATABASE CONECTION AND EXECUTE THE QUERY CHANGE ACCORDIGN TO YOUR NEED
    Set DBRS = CreateObject("ADODB.Recordset")
    With DBRS
        .Source = Rng_SQL
        Set .ActiveConnection = DBCON
        
        .Open
            'With Range("Rng_Result")
            '    .ClearContents
            '    .ClearFormats
            '    .CopyFromRecordset DBRS
               
               For i = 1 To DBRS.fields.Count - 1
                             msgbox DBRS.fields(i)
							 exit for
               Next 

            'End With

	End With