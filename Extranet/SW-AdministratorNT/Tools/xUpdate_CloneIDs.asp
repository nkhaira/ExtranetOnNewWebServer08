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

SQL = "SELECT DISTINCT Item_Number_2 " &_
      "FROM         dbo.Calendar " &_
      "WHERE (LEN(Item_Number_2) = 13) AND ([Language] <> 'eng') AND (Clone = 0 or Clone IS NULL) and site_id=82" &_
      "ORDER BY Item_Number_2"

Set rsID = Server.CreateObject("ADODB.Recordset")
rsID.Open SQL, conn, 3, 3

do while not rsID.EOF

  if not isnull(rsID("Item_Number_2")) and len(rsID("Item_Number_2")) = 13 then
  
    if instr(1,rsID("Item_Number_2")," ") > 0 then

      ItemDetail = split(rsID("Item_Number_2")," ")
      
      if UBound(ItemDetail) = 2 then
      
        if (len(ItemDetail(0)) = 7 and isnumeric(ItemDetail(0))) and _
           len(ItemDetail(1)) = 1 and _
           len(ItemDetail(2)) = 3 and _
           instr(1,rsID("Item_Number_2"),".") = 0 and _
           instr(1,rsID("Item_Number_2"),"-") = 0 then
          
          SQL = "SELECT ID, Revision_Code FROM Calendar WHERE Item_Number='" & ItemDetail(0) & "' AND Language='" & ItemDetail(2) & "'"
          
          Set rsIDM = Server.CreateObject("ADODB.Recordset")
          rsIDM.Open SQL, conn, 3, 3
          
          if not rsIDM.EOF then
          
            SQL = "UPDATE Calendar SET CLONE=" & rsIDM("ID") & " WHERE Item_Number_2='" & rsID("Item_Number_2") & "' AND Site_ID=82 AND Language <> 'eng'"
            response.write SQL & "<P>"            
            conn.execute SQL  
          end if
          
          rsIDM.close
          set rsIDM = nothing
        
        else
          response.write rsID("Item_Number_2") & "<P>"
        end if
          
      end if
        
    end if    

  end if         
  
  rsID.MoveNext

loop

rsID.close
set rsID = nothing

Call Disconnect_SiteWide
%>