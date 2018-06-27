<!-- #include file="servers.asp" -->
<%
' Desctiption: Connection strings for Fluke_WhereToBuy database
' Use: US where-to-buy pages
'	Fluke_WWW\AdminTools\USWTBAdmin
'	Fluke_WWW\WTB\wtb.asp

Dim dbConnWTB
Dim strConnectionString_WTB
Dim strDatabaseName_WTB
Dim strServerName_WTB
Dim strLogin_WTB

Server.ScriptTimeOut = 1500
strServerName_WTB = UCase(Request("SERVER_NAME"))
strDatabaseName_WTB = "FLUKE_WHERETOBUY"
strLogin_WTB = "WEBUSER"

strConnectionString_WTB = GetConnectionString(strDatabaseName_WTB, strServerName_WTB, strLogin_WTB)

Function GetConnWTB()
	Dim localConnect

	set localConnect = Server.CreateObject("ADODB.Connection")
	localConnect.ConnectionTimeout = 90
	localConnect.Open strConnectionString_WTB

	set GetConnWTB = localConnect
End Function

Sub ConnectWTB()
	set dbConnWTB = GetConnWTB()
End Sub

'====================================================================================================================

Sub DisconnectWTB()
	dbConnWTB.Close
	set dbConnWTB = nothing
End Sub
%>
