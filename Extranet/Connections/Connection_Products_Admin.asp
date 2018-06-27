<!-- #include file="servers.asp" -->
<%
Dim dbConnProducts
dim strConnectionString_Products

Sub ConnectProducts()
	Dim strConnectionString
	Dim strDatabaseName
	Dim strServerName
	Dim strLogin
	
	strServerName = UCase(Request.ServerVariables("SERVER_NAME"))
	strDatabaseName = "FLUKE_PRODUCTS"
	strLogin = "ADMIN"
	
	'strConnectionString = GetConnectionString(strDatabaseName, strServerName, strLogin)
	
	'*** the line below is for testing purposes
	strConnectionString_Products = GetConnectionString(strDatabaseName, strServerName, strLogin)
	'*********************************************
	
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

Function GetAdminConnectionString()
	Dim strConnectionString
	Dim strDatabaseName
	Dim strServerName
	Dim strLogin
	
	strServerName = UCase(Request.ServerVariables("SERVER_NAME"))
	strDatabaseName = "FLUKE_PRODUCTS"
	strLogin = "ADMIN"
	
	'strConnectionString = GetConnectionString(strDatabaseName, strServerName, strLogin)
	
	'*** the line below is for testing purposes
	strConnectionString_Products = GetConnectionString(strDatabaseName, strServerName, strLogin)
	'*********************************************
	
	GetAdminConnectionString = strConnectionString_Products
End Function
%>