<!-- #include file="servers.asp" -->
<%
' Desctiption: Pim's poor's mans product database
'	All country sites

Dim dbConnWTB2000
Dim strConnectionString_WTB2000
Dim strDatabaseName_Products
Dim strServerName_Products
Dim strLogin_Products

Server.ScriptTimeOut = 1500
strServerName_WTB2000 = UCase(Request("SERVER_NAME"))
strDatabaseName_WTB2000 = "WTB"
strLogin_WTB2000 = "WEBUSER"

strConnectionString_WTB2000 = GetConnectionString(strDatabaseName_WTB2000, strServerName_WTB2000, strLogin_WTB2000)

Function GetConnWTB2000()
	Dim localConnect

	set localConnect = Server.CreateObject("ADODB.Connection")
	localConnect.ConnectionTimeout = 90
	localConnect.Open strConnectionstring_WTB2000

	set GetConnWTB2000 = localConnect
End Function

Sub ConnectWTB2000()
	set dbConnWTB2000 = GetConnWTB2000()
End Sub

'====================================================================================================================

Sub DisconnectWTB2000()
	dbConnWTB2000.Close
	set dbConnWTB2000 = nothing
End Sub
%>
