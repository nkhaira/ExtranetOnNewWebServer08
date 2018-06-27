<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Session.timeout = 240 ' Set to 4 Hours
Server.ScriptTimeout = 2 * 60

Call Connect_SiteWide

SQL = "SELECT DISTINCT dbo.Activity.Account_ID " &_
      "FROM         dbo.Activity LEFT OUTER JOIN " &_
      "             dbo.UserData ON dbo.Activity.Account_ID = dbo.UserData.ID " &_
      "WHERE     (dbo.UserData.Business_Country IS NULL) AND (dbo.UserData.Business_Country <> 'US') " &_
      "ORDER BY Account_ID"

Set rsID = Server.CreateObject("ADODB.Recordset")
rsID.Open SQL, conn, 3, 3

do while not rsID.EOF

  SQL = "SELECT Business_Country, Region FROM UserData WHERE ID=" & rsID("Account_ID")
  Set rsID2 = Server.CreateObject("ADODB.Recordset")
  rsID2.Open SQL, conn, 3, 3
  
  if not rsID2.EOF then
    SQL = "UPDATE Activity SET Region=" & rsID2("Region") & ", Country='" & rsID2("Business_Country") & "' WHERE Account_ID=" & rsID("Account_ID")
    conn.execute SQL
  end if
  
  rsID2.close
  set rsID2 = nothing
    
  rsID.MoveNext

loop

rsID.close
set rsID = nothing

Call Disconnect_SiteWide
%>