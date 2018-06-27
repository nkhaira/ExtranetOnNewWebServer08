<!-- #include file="servers.asp" -->
<%
Dim dbConnFormData
Dim strConnectionString_FormData
Dim strDatabaseName_FormData
Dim strServerName_FormData
Dim strLogin_FormData

Server.ScriptTimeOut = 1500
strServerName_FormData = UCase(Request("SERVER_NAME"))
strDatabaseName_FormData = "FLUKE_FORMDATA"
strLogin_FormData = "WEBUSER"

strConnectionString_FormData = GetConnectionString(strDatabaseName_FormData, strServerName_FormData, strLogin_FormData)

'Response.Write("test: " & strConnectionString_FormData & "<BR>")
'Response.end

Sub Connect_FormDatabase()
	Dim strServerName
	
	set dbConnFormData = Server.CreateObject("ADODB.Connection")
	dbConnFormData.ConnectionTimeOut = 120
	dbConnFormData.CommandTimeout = 120
	dbConnFormData.Open strConnectionString_FormData
End Sub

Sub Disconnect_FormDatabase()
	if not IsObject(dbConnFormData) then
		dbConnFormData.Close
		set dbConnFormData = nothing
	end if
End Sub
%>

