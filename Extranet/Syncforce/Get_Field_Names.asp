<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%

Call Connect_SiteWide

SQL = "SELECT * FROM UserData WHERE ID=2674"
Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.Open SQL, conn, 3, 3

if not rsUser.EOF then
  For field_num = 0 To rsUser.Fields.Count - 1
    response.write rsUser.Fields(field_num).Name & "<BR>"
  Next
end if

rsUser.close
set rsUser = nothing

Call Disconnect_SiteWide

%>

</body>
</html>
