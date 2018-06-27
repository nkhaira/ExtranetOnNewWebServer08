<!-- #include file="servers.asp" -->
<%
Dim dbConnProducts
Dim strConnectionString_Products

Sub ConnectProducts()
	Dim strDatabaseName
	Dim strServerName
	Dim strLogin
	
	strServerName = UCase(Request("SERVER_NAME"))
	strDatabaseName = "FLUKE_PRODUCTS"
	strLogin = "WEBUSER"
	
	strConnectionString_Products = GetConnectionString(strDatabaseName, strServerName, strLogin)
	
	set dbConnProducts = Server.CreateObject("ADODB.Connection")
	dbConnProducts.ConnectionTimeout = 30
	
	dbConnProducts.Open strConnectionString_Products
End Sub

Sub DisconnectProducts()
	if IsObject(dbConnProducts) then
		dbConnProducts.Close
		set dbConnProducts = nothing
	end if
End Sub
%>