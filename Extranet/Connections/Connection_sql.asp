<!-- #include file="servers.asp" -->
<%

Function DB_connect(strDbName,strLogin,strServerName)
	Dim strConnectionString
	Dim localConnect
	
	strConnectionString = GetConnectionString(strDbName, strServerName, strLogin)
	
	set localConnect = Server.CreateObject("ADODB.Connection")
	localConnect.ConnectionTimeout = 90
	localConnect.Open strConnectionString
	
	set DB_connect = localConnect	
End Function

Sub DB_disconnect(dbconn)
	dbconn.close
	set dbconn = nothing
End sub

%>