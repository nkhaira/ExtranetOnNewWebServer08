<!-- #include file="servers.asp" -->
<%
Dim conn
Dim strConnectionString_Promo
Dim strDatabaseName_Promo
Dim strServerName_Promo
Dim strLogin_Promo

Server.ScriptTimeOut = 1500
strServerName_Promo = UCase(Request("SERVER_NAME"))
strDatabaseName_Promo = "FLUKE_PROMO"
strLogin_Promo = "WEBUSER"

strConnectionString_Promo = GetConnectionString(strDatabaseName_Promo, strServerName_Promo, strLogin_Promo)

set conn=server.createobject("Adodb.connection")

conn.Open strConnectionString_Promo

%>