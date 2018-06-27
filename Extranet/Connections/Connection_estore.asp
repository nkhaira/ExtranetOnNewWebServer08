<!-- #include file="servers.asp" -->
<%
Dim eConn
Dim strConnectionString_eStore
Dim strDatabaseName_eStore
Dim strServerName_eStore
Dim strLogin_eStore

Server.ScriptTimeOut = 1500
strServerName_eStore = UCase(Request("SERVER_NAME"))
strDatabaseName_eStore = "ESTORE"
strLogin_eStore = "WEBUSER"

strConnectionString_eStore = GetConnectionString(strDatabaseName_eStore, strServerName_eStore, strLogin_eStore)

Sub Connect_eStoreDatabase()
	set eConn = Server.CreateObject("ADODB.Connection")

	eConn.open strConnectionString_eStore
End Sub

Sub Disconnect_eStoreDatabase()
	if not IsObject(eConn) then
		eConn.Close
		set eConn = nothing
	end if
End Sub
%>

