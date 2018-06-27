<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' The purpose of this script is to map the language Clone to the original ENG version in SiteWide.Calendar
' This script can be run anytime syncing is required.
'
' Author: Kelly Whitlock

Session.timeout      = 240 ' Set to 4 Hours
Server.ScriptTimeout = 10 * 60

Call Connect_SiteWide

SQL = "SELECT Name AS ID, Value1 AS thumbnail " &_
      "FROM dbo.Calendar_Temp " &_
      "WHERE (Value1 IS NOT NULL) " &_
      "ORDER BY ID"
      
Set rsID = Server.CreateObject("ADODB.Recordset")
rsID.Open SQL, conn, 3, 3

do while not rsID.EOF

  SQLU = "UPDATE Calendar SET Thumbnail='" & rsID("Thumbnail") & "' WHERE ID=" & rsID("ID") & "Site_ID=82"
  response.write SQLU & "<P>"
  'conn.execute SQLU
  
  rsID.MoveNext

loop

rsID.close
set rsID = nothing

Call Disconnect_SiteWide
%>