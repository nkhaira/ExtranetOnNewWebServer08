<!-- #include file="servers.asp" -->
<%
' Desctiption: Connection strings for Fluke_Buying database
' Use: International sites where-to-buy and some warranty pages
'	Fluke_WWW\AdminTools\WTBAdmin\*.*
'	Fluke_WWW\Ex_Home\hs~select.asp
'	Fluke_WWW\Ex_Home\Reps.asp
'	Fluke_WWW\Ex_Home\Select.asp
'	Fluke_WWW\WTB\reps.asp
'	Fluke_WWW\WTB\select.asp
'	All international sites including:
'		reps.asp, sales.asp, sales_table.asp, distributors.asp, purchase.asp
'
'	Approx 274 files total

Dim conn
Dim strConnectionString_Buying
Dim strDatabaseName_Buying
Dim strServerName_Buying
Dim strLogin_Buying

Server.ScriptTimeOut = 1500
strServerName_Buying = UCase(Request("SERVER_NAME"))
strDatabaseName_Buying = "FLUKE_BUYING"
strLogin_Buying = "WEBUSER"

strConnectionString_Buying = GetConnectionString(strDatabaseName_Buying, strServerName_Buying, strLogin_Buying)
Set conn = Server.CreateObject("ADODB.Connection")

conn.Open strConnectionString_Buying

%>
