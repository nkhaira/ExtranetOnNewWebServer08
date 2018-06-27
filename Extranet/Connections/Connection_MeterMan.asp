<!-- #include file="servers.asp" -->
<%
Dim dbConnMeterMan
Dim strConnectionString_MeterMan
Dim strDatabaseName_Products
Dim strServerName_Products
Dim strLogin_Products

Server.ScriptTimeOut = 1500
strServerName_MeterMan = UCase(Request("SERVER_NAME"))
strDatabaseName_MeterMan = "METERMAN"
strLogin_MeterMan = "WEBUSER"

strConnectionString_MeterMan = GetConnectionString(strDatabaseName_MeterMan, strServerName_MeterMan, strLogin_MeterMan)

Sub Connect_MeterMandatabase()
	set dbConnMeterMan = Server.CreateObject("ADODB.Connection")

	dbConnMeterMan.ConnectionTimeOut = 120
	dbConnMeterMan.CommandTimeout = 120
	dbConnMeterMan.Open strConnectionString_MeterMan
End Sub

Sub Disconnect_MeterMandatabase()
	if not IsObject(dbConnMeterMan) then
		dbConnMeterMan.Close
		set dbConnMeterMan = nothing
	end if
End Sub
%>

