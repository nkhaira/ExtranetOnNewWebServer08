<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' The purpose of this script is to update the Cost Center and owner of the SiteWide.Calendar
' asset for those assets that have corresponding entries in Literature_Item_US table.
'
' This script can be run anytime syncing is required.
'
' Author: Kelly Whitlock

Session.timeout      = 240      ' 4-Hours
Server.ScriptTimeout = 60 * 10  ' 10-Minutes

Call Connect_SiteWide

SQL = "SELECT DISTINCT " &_
      " dbo.Calendar.Item_Number, dbo.Calendar.Site_ID, dbo.Calendar.Revision_Code, dbo.Literature_Items_US.ITEM, " &_
      " dbo.Literature_Items_US.COST_CENTER, dbo.Literature_Items_US.MARCOM_MANAGER " &_
      "FROM dbo.Literature_Items_US RIGHT OUTER JOIN " &_
      "     dbo.Calendar ON dbo.Literature_Items_US.REVISION = dbo.Calendar.Revision_Code AND " &_
      "     dbo.Literature_Items_US.ITEM = dbo.Calendar.Item_Number " &_
      "WHERE (LEN(dbo.Calendar.Item_Number) = 7) AND (LEN(dbo.Literature_Items_US.ITEM) = 7) AND (LEN(dbo.Literature_Items_US.COST_CENTER) = 4) AND " &_
      "      (dbo.Literature_Items_US.ACTIVE_FLAG = - 1) " &_
      "ORDER BY dbo.Calendar.Item_Number, dbo.Calendar.Revision_Code DESC"

Set rsID = Server.CreateObject("ADODB.Recordset")
rsID.Open SQL, conn, 3, 3

Dim MarComMgr

do while not rsID.EOF

  if instr(1,rsID("Marcom_Manager"),",") > 0 then
    MarComMgr = split(rsID("Marcom_Manager"),",")
  else  
    MarComMgr = ""
  end if
  
  SQLM = "SELECT ID FROM dbo.UserData " &_
         "WHERE (SubGroups LIKE '%administrator%' OR " &_
         "SubGroups LIKE '%content%') AND (LastName = '" & MarComMgr(0) & "') AND (Site_ID = " & rsID("Site_ID") & ")"

  Set rsM = Server.CreateObject("ADODB.Recordset")
  rsM.Open SQLM, conn, 3, 3
  
  if not rsM.EOF then
    SQLU = "UPDATE Calendar SET Submitted_By=" & rsM("ID") & ", Approved_By=" & rsM("ID") & ", Cost_Center=" & rsID("Cost_Center") & " WHERE Item_Number='" & rsID("Item_Number") & "'"
  else
    SQLU = "UPDATE Calendar SET Cost_Center=" & rsID("Cost_Center") & " WHERE Item_Number='" & rsID("Item_Number") & "'"
  end if
  
  rsM.Close
  set rsM = nothing

  response.write SQLU & "<P>"
  response.flush
  
  conn.execute SQLU

  rsID.MoveNext

loop

rsID.close
set rsID = nothing

Call Disconnect_SiteWide
%>