<!-- #include file="servers.asp" -->
<%
dim dbconn
dim strConnectionString_Parts
Dim strDatabaseName_Parts
Dim strServerName_Parts
Dim strLogin_Parts

Server.ScriptTimeOut = 1500
strServerName_Parts = UCase(Request("SERVER_NAME"))
strDatabaseName_Parts = "ESTORE"
strLogin_Parts = "WEBUSER"

strConnectionString_Parts = GetConnectionString(strDatabaseName_Parts, strServerName_Parts, strLogin_Parts)

'response.write("string; " & strconnectionstring_parts & "<BR>")

sub connect_parts()
	Set DBConn = Server.CreateObject("ADODB.Connection")
	dbconn.open strConnectionString_Parts
end sub

sub disconnect_parts()
	dbconn.close
	set dbconn = nothing
end sub

%>
