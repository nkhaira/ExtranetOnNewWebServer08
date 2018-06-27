<!-- #include file="servers.asp" -->
<%
Dim connUse
Dim strConnectionString_VirDir
Dim strDatabaseName_VirDir
Dim strServerName_VirDir
Dim strLogin_VirDir

Server.ScriptTimeOut = 1500
strServerName_VirDir = UCase(Request("SERVER_NAME"))
strDatabaseName_VirDir = "FLUKE_VIRTUALDIRECTORIES"
strLogin_VirDir = "WEBUSER"

strConnectionString_VirDir = GetConnectionString(strDatabaseName_VirDir, strServerName_VirDir, strLogin_VirDir)

Set connUse = Server.CreateObject("ADODB.Connection")

connUse.Open strConnectionString_VirDir

%>
