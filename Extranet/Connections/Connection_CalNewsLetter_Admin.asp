<!-- #include file="servers.asp" -->
<!-- #include file="Adovbs.inc" -->

<%
Dim dbConnCalNews

Sub ConnectCalNews()
	Dim strConnectionString_CalNews
	Dim strDatabaseName
	Dim strServerName
	Dim strLogin
	
	strServerName = UCase(Request("SERVER_NAME"))
	'strDatabaseName = "FLUKE_CALNEWSLETTER"
	strDatabaseName = "FLUKE_CALNEWSLETTER"
	strLogin = "ADMIN"
	'strLogin = "WEBUSER"
	
	strConnectionString_CalNews = GetConnectionString(strDatabaseName, strServerName, strLogin)
	 
	set dbConnCalNews = Server.CreateObject("ADODB.Connection")
	dbConnCalNews.ConnectionTimeOut = 120
	dbConnCalNews.CommandTimeout = 120
	
	dbConnCalNews.Open strConnectionString_CalNews
End Sub

Sub DisconnectCalNews()
	if IsObject(dbConnCalNews) then
		dbConnCalNews.Close
		set dbConnCalNews = nothing
	end if
End Sub
%>