<!-- #include virtual="/admin/Adovbs.inc" -->
<%
Dim dbConn, strConnString
Server.ScriptTimeOut = 1500
Sub Connect_RMDatabase()
	Dim strServerName

	set dbConn = Server.CreateObject("ADODB.Connection")
	strServerName = UCase(Request("SERVER_NAME"))

  if instr(ucase(strServerName),"DEV") then
    strConnString = "DRIVER={SQL Server}; SERVER=EVTIBG03; UID=flukewebuser; DATABASE=ReliableMeters; pwd=57OrcaSQL"
  else  
    strConnString = "DRIVER={SQL Server}; SERVER=FLKPRD03; UID=flukewebuser; DATABASE=ReliableMeters; pwd=57OrcaSQL"
	end if
  'strConnString = "DRIVER={SQL Server}; SERVER=FLKPRD03; UID=flukewebuser; DATABASE=ReliableMeters; pwd=57OrcaSQL"
	'strConnString = "DRIVER={SQL Server}; SERVER=EVTIBG03; UID=flukewebuser; DATABASE=ReliableMeters; pwd=57OrcaSQL"

	dbConn.ConnectionTimeOut = 120
	dbConn.CommandTimeout = 120
	dbConn.Open strConnString
End Sub

Sub Disconnect_RMDatabase()
	if not IsObject(dbConn) then
		dbConn.Close
		set dbConn = nothing
	end if
End Sub
%>


