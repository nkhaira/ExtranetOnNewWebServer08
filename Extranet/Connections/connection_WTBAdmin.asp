<%
'Connection string to buying database which generates the content for the /wtb/select.asp page
'Dim strServerName
strServerName = UCase(Request("SERVER_NAME"))
set mdbconn=server.createobject("Adodb.connection")
mdbconn.connectiontimeout = 45	

if InStr(1, strServerName, "MILHOUSE") or InStr(1, strServerName, "216.244.65.27") or InStr(1, strServerName, "216.244.65.42") then
	if request.querystring("db") = "salesex" then
		mdbconn.open "DefaultDir=c:\inetpub\FlukeSite\AdminTools\WTBAdmin\BulkOps;DBQ=salesex.mdb;CursorTypeEnum=2;DRIVER=Microsoft Access Driver (*.mdb);UID=admin;UserCommitSync=Yes;Threads=3;PageTimeout=5;MaxScanRows=8;MaxBufferSize=512;ImplicitCommitSync=Yes;FIL=MS Access;DriverId=25"
	elseif request.querystring("db") = "installers" then
		mdbconn.open "DefaultDir=E:\wwwroot\Fluke_WWW\AdminTools\WTBAdmin\BulkOps;DBQ=installers.mdb;CursorTypeEnum=2;DRIVER=Microsoft Access Driver (*.mdb);UID=admin;UserCommitSync=Yes;Threads=3;PageTimeout=5;MaxScanRows=8;MaxBufferSize=512;ImplicitCommitSync=Yes;FIL=MS Access;DriverId=25"
	elseif request.querystring("db") = "distributors" then
		mdbconn.open "DefaultDir=E:\wwwroot\Fluke_WWW\AdminTools\WTBAdmin\BulkOps;DBQ=distributors.mdb;CursorTypeEnum=2;DRIVER=Microsoft Access Driver (*.mdb);UID=admin;UserCommitSync=Yes;Threads=3;PageTimeout=5;MaxScanRows=8;MaxBufferSize=512;ImplicitCommitSync=Yes;FIL=MS Access;DriverId=25"
	elseif request.querystring("db") = "productgroup" then
		mdbconn.open "DefaultDir=E:\wwwroot\Fluke_WWW\AdminTools\WTBAdmin\BulkOps;DBQ=productgroup.mdb;CursorTypeEnum=2;DRIVER=Microsoft Access Driver (*.mdb);UID=admin;UserCommitSync=Yes;Threads=3;PageTimeout=5;MaxScanRows=8;MaxBufferSize=512;ImplicitCommitSync=Yes;FIL=MS Access;DriverId=25"
	end if
'    mdbconn.Open "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=/inetpub/flukesite/AdminTools/WTBAdmin/BulkOps/salesex.mdb;"
'    mdbconn.Open "DRIVER={Microsoft Access Driver(*.mdb)};DBQ=/inetpub/flukesite/AdminTools/WTBAdmin/BulkOps/salesex.mdb", "", ""
else
	if request.querystring("db") = "salesex" then
		mdbconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\wwwroot\Fluke_WWW\AdminTools\WTBAdmin\BulkOps\salesex.mdb","Admin", ""  
	elseif request.querystring("db") = "installers" then
		mdbconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\wwwroot\Fluke_WWW\AdminTools\WTBAdmin\BulkOps\installers.mdb","Admin", ""  
	elseif request.querystring("db") = "distributors" then
		mdbconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\wwwroot\Fluke_WWW\AdminTools\WTBAdmin\BulkOps\distributors.mdb","Admin", ""  
	elseif request.querystring("db") = "productgroup" then
		mdbconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\wwwroot\Fluke_WWW\AdminTools\WTBAdmin\BulkOps\productgroup.mdb","Admin", ""  
	end if
'	conn.open "DRIVER={SQL Server};SERVER=216.244.76.70;UID=flukewebuser;DATABASE=fluke_buying;pwd=57OrcaSQL"
end if

%>