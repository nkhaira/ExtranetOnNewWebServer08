<%@ Language="VBScript" CODEPAGE="65001" %>
<html>
<head>
<title>Success!</title>
</head>
<body>
<%
Server.ScriptTimeout = 5 * 60   ' Set to 5 Minutes
Set conn = Server.CreateObject("ADODB.connection")  
'conn.CommandTimeout =180
conn.Open "Provider=MSDASQL.1;Extended Properties=DRIVER=SQL Server;SERVER=EVTIBG18.TC.FLUKE.COM;UID=SITEWIDE_WEB;PWD=tuggy_boy;APP=Internet Information Services;WSID=DTMEVTVSDV15;DATABASE=FLUKE_SITEWIDE"
Set rsUser = Server.CreateObject("ADODB.Recordset")  
'Set rsUser = conn.execute( "exec getmetricsdata 3,2009"  )  
Set rsUser = conn.execute( "select * from activity"  )  
if not rsUser.EOF then
		Response.Write rsUser.recordcount
else
			Response.Write "No Rows"
end if
%>
</body>
</html> 



