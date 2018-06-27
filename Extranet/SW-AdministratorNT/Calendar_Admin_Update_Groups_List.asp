<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%

Dim RegionColor(4)
RegionColor(0) = "#0000CC"
RegionColor(1) = "#99FFCC"
RegionColor(2) = "#66CCFF"
RegionColor(3) = "#FFCCFF"
RegionColor(4) = "#FFCC99"

Call Connect_SiteWide

Site_ID = 82
SubGroups = "uss"
Category_Code  = 1

response.write "<HTTP><BODY>"

SQL = "SELECT * FROM Subgroups WHERE Enabled=-1 AND Site_ID=" & Site_ID & " ORDER BY Order_Num"
Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
rsSubGroups.Open SQL, conn, 3, 3

SQL = "SELECT ID, Title, Item_Number, Language, SubGroups FROM Calendar WHERE Site_ID=" & Site_ID & " AND Code=" & Category_Code & " ORDER BY Status, Title"
Set rsCalendar = Server.CreateObject("ADODB.Recordset")
rsCalendar.Open SQL, conn, 3, 3

with response

  .write "<TABLE CELLPADDING=0 CELLSPACING=0>"
  
  do while not rsCalendar.EOF

    SubGroups = rsCalendar("SubGroups")  

    .write "<TR>"
    .write "<TD>" & rsCalendar("Title") & "</TD>"
    .write "<TD>" & rsCalendar("Item_Number") & "<TD>"
    .write "<TD>" & rsCalendar("Language") & "<TD>"

  
    rsSubGroups.MoveFirst
    do while not rsSubGroups.EOF
  
      .write "<TD BGCOLOR=""" & RegionColor(rsSubGroups("Region")) & """>"
      .write "<INPUT TYPE=""CHECKBOX"" NAME="""" VALUE="""" TITLE=""" & rsSubGroups("X_Description") & """"
      if Instr(1,LCase(SubGroups),LCase(rsSubGroups("Code"))) > 0 then .write " CHECKED"
      .write ">" & vbCrLf
      .write "</TD>"
    
      rsSubGroups.MoveNext
     loop
     
     .write "</TR>"
     .flush
     
     rsCalendar.MoveNext
     
   loop
 
   .write "</TABLE>"

 end with
 
 rsCalendar.close
 set rsCalendar = nothing

 rsSubGroups.close
 set rsSubGroups = nothing
    

response.write "</BODY></HTTP>"

Call Disconnect_SiteWide
%>

