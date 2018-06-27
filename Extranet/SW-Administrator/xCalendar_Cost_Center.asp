<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Session.timeout = 240 ' Set to 4 Hours
Server.ScriptTimeout = 2 * 60

Call Connect_SiteWide

SQL = "SELECT DISTINCT dbo.Calendar.Item_Number, dbo.Calendar.Revision_Code, dbo.Literature_Items_US.ITEM, dbo.Literature_Items_US.COST_CENTER " &_
      "FROM         dbo.Calendar LEFT OUTER JOIN " &_
      "             dbo.Literature_Items_US ON dbo.Calendar.Item_Number = dbo.Literature_Items_US.ITEM " &_
      "WHERE     (LEN(dbo.Calendar.Item_Number) = 7) AND (LEN(dbo.Literature_Items_US.ITEM) = 7) AND len(dbo.Literature_Items_US.Cost_Center)=4 " &_
      "ORDER BY dbo.Calendar.Revision_Code"

Set rsID = Server.CreateObject("ADODB.Recordset")
rsID.Open SQL, conn, 3, 3

do while not rsID.EOF

  SQL = "UPDATE Calendar SET Cost_Center=" & rsID("Cost_Center") & " WHERE Item_Number='" & rsID("Item_Number") & "'"
  conn.execute SQL

  rsID.MoveNext

loop

rsID.close
set rsID = nothing

Call Disconnect_SiteWide
%>