<!-- #include file="servers.asp" -->
<%
Dim strConnectionString_Forums
Dim strDatabaseName_Forums
Dim strServerName_Forums
Dim strLogin_Forums

Server.ScriptTimeOut = 1500
strServerName_Forums = UCase(Request("SERVER_NAME"))
strDatabaseName_Forums = "FLUKE_SITEWIDE"
strLogin_Forums = "WEBUSER"

strConnectionString_Forums = GetConnectionString(strDatabaseName_Forums, strServerName_Forums, strLogin_Forums)

<%
'response.write "config.ADMINSETTING_DatabaseDSN = ""DRIVER={SQL Server}; SERVER=216.244.76.70; UID=sitewide_usr;DATABASE=fluke_sitewide;pwd=tuggy_boy"";" & vbCrLf
response.write "config.ADMINSETTING_DatabaseDSN = """ & strConnectionString_Forums & """;" & vbCrLf
%>

