<!-- #include file="servers.asp" -->
<%
Dim dbConnGlobalFormData
Dim strConnectionString_FormData
Dim strDatabaseName_FormData
Dim strServerName_FormData
Dim strLogin_FormData

Server.ScriptTimeOut = 1500
strServerName_FormData = UCase(Request("SERVER_NAME"))
strDatabaseName_FormData = "GLOBAL_FORMDATA"
strLogin_FormData = "WEBUSER"

strConnectionString_FormData = GetConnectionString(strDatabaseName_FormData, strServerName_FormData, strLogin_FormData)

'Response.Write("test: " & strConnectionString_FormData & "<BR>")
'Response.end

Sub Connect_GlobalFormData()
	Dim strServerName
	
	set dbConnGlobalFormData = Server.CreateObject("ADODB.Connection")
	dbConnGlobalFormData.ConnectionTimeOut = 120
	dbConnGlobalFormData.CommandTimeout = 120
	dbConnGlobalFormData.Open strConnectionString_FormData
End Sub

Sub Disconnect_GlobalFormData()
	if not IsObject(dbConnGlobalFormData) then
		dbConnGlobalFormData.Close
		set dbConnGlobalFormData = nothing
	end if
End Sub
%>

