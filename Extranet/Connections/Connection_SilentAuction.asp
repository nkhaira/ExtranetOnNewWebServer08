<!-- #include file="servers.asp" -->
<%
Dim dbConnSilentAuction
Dim strConnectionString_SilentAuction
Dim strDatabaseName_SilentAuction
Dim strServerName_SilentAuction
Dim strLogin_SilentAuction

Server.ScriptTimeOut = 1500
strServerName_SilentAuction = UCase(Request("SERVER_NAME"))
strDatabaseName_SilentAuction = "FLUKE_PROMO"
strLogin_SilentAuction = "WEBUSER"

strConnectionString_SilentAuction = GetConnectionString(strDatabaseName_SilentAuction, strServerName_SilentAuction, strLogin_SilentAuction)

Function GetConnSilentAuction()
	Dim localConnect

	set localConnect = Server.CreateObject("ADODB.Connection")
	localConnect.ConnectionTimeout = 90
	localConnect.Open strConnectionString_SilentAuction

	set GetConnSilentAuction = localConnect
End Function

Sub ConnectSilentAuction()
	set dbConnSilentAuction = GetConnSilentAuction()
End Sub

Sub DisconnectSilentAuction()
	dbConnSilentAuction.Close
	set dbConnSilentAuction = nothing
End Sub
%>