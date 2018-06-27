<!-- #include file="servers.asp" -->
<%
Dim dbConnBrazilWeb
Dim strConnectionString_brazilweb
Dim strDatabaseName
Dim strServerName
Dim strLogin

Server.ScriptTimeOut = 1500
strServerName = UCase(Request("SERVER_NAME"))
strDatabaseName = "BRAZILWEB"
strLogin = "WEBUSER"

strConnectionString_brazilweb = GetConnectionString(strDatabaseName, strServerName, strLogin)

'Response.Write("test: " & strConnectionString_BrazilWeb & "<BR>")
'Response.end


	Dim Conexao
	
	set Conexao = Server.CreateObject("ADODB.Connection")
	Conexao.ConnectionTimeOut = 120
	Conexao.CommandTimeout = 120
	Conexao.Open strConnectionString_brazilweb



%>

