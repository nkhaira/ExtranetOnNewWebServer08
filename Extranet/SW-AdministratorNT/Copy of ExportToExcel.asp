<%@ Language="VBScript" CODEPAGE="65001" %>
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
' --------------------------------------------------------------------------------------
' Author:     P.Deshpande.
' Date:       06/1/2007
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------
on error resume next
Response.Buffer = false
Session.timeout = 60            ' Set to 1 Hour
Server.ScriptTimeout = 5 * 60   ' Set to 5 Minutes

Call Connect_SiteWide

dim Rs
set Rs=server.createobject("ADODB.recordset")

SQL =       " SELECT distinct dbo.Calendar.ID,dbo.Calendar_Category.Title AS Category,dbo.Calendar.Product,dbo.calendar.title as Title, " &_
            " dbo.Calendar.Item_Number, dbo.Literature_Items_US.Revision AS Lit_Revision, " &_
            " dbo.Calendar.Cost_Center, " & _
            " Calendar.Status, Calendar.Sub_Category, Calendar.Product, " & _
            " dbo.Literature_Items_US.ACTIVE_FLAG,Calendar_Category.Sort,Calendar.BDate,Calendar.Revision_Code,CASE dbo.Calendar.Clone WHEN 0 THEN dbo.Calendar.ID ELSE dbo.Calendar.Clone END AS PC_Order " & _
            " , (select top 1 language.sort from language where code=calendar.language) as languagesort " &_
            " FROM   dbo.Calendar " &_
            " LEFT OUTER JOIN " &_
            " dbo.Literature_Items_US ON dbo.Calendar.Revision_Code = dbo.Literature_Items_US.REVISION AND  " &_
            " dbo.Calendar.Item_Number = dbo.Literature_Items_US.ITEM LEFT OUTER JOIN " &_
            " dbo.UserData ON dbo.Calendar.Submitted_By = dbo.UserData.ID LEFT OUTER JOIN " &_
            " dbo.Calendar_Category ON dbo.Calendar.Category_ID = dbo.Calendar_Category.ID " &_
            " WHERE  dbo.Calendar.Site_ID=" & Request.QueryString("Site_ID") & " and ACTIVE_FLAG = -1 " 
            
if Request.QueryString("language") <> "" then
    SQl = SQL & " AND Calendar.Language='" & Request.QueryString("language") & "'"
end if
            
if Request.QueryString("categoryid") > 0 then
    SQl = SQL & " AND Calendar.Category_ID=" & Request.QueryString("categoryid") 
end if

if Request.QueryString("GroupId") > 0 then
    SQl = SQL & " AND Calendar.Subgroups=" & Request.QueryString("groupid") 
end if

if Request.QueryString("Country") <> "" then
    SQL = SQL & " AND (Calendar.Country = 'none' OR Calendar.Country NOT LIKE '0%' AND Calendar.Country LIKE '%" & Country & "%')" & " "   
end if
 
if Request.QueryString("Submitted_By") > 0 then
      SQL = SQL & "AND Calendar.Submitted_By=" & Submitted_By & " "
end if

if Request.QueryString("Campaign") > 0 then
      SQL = SQL & "ORDER BY Calendar_Category.Sort, Calendar.Status, Calendar_Category.Title, Calendar.Sub_Category, Calendar.Product, Calendar.Revision_Code Desc, Calendar.ID, Lit_ACTIVE_FLAG "
else
  select case Request.QueryString("Sort_By")
        case 1  ' Asset ID
          SQL = SQL & " ORDER BY Calendar.Status, Calendar.ID, Calendar.Revision_Code Desc, dbo.Literature_Items_US.ACTIVE_FLAG "
        case 2  ' Item Number + Revision
          SQL = SQL & " ORDER BY Calendar.Item_Number, dbo.Literature_Items_US.ACTIVE_FLAG, Calendar.Revision_Code Desc "
        case 3  ' Begin Date
          SQL = SQL & " ORDER BY Calendar.BDate, dbo.Literature_Items_US.ACTIVE_FLAG, Calendar.Revision_Code Desc "        
        case 4  ' Category, Sub Category, Begin Date, ID
           SQL = SQL & " ORDER BY Calendar_Category.Sort,Calendar_Category.Title,Calendar.Product,Calendar.BDate, Calendar.ID,dbo.Literature_Items_US.ACTIVE_FLAG, Calendar.Revision_Code Desc "        
        case 5  ' Parent / Clone, Language
          ' Modified by zensar on 09-03-2007 as adding sort to select statement is returning duplicate rows.
          'SQL = SQL & "ORDER BY PC_Order, dbo.[Language].Sort"
          '>>>>>>>>>>>>
          SQL = SQL & " ORDER BY PC_Order,languagesort"
        case 6  ' Title
          SQL = SQL & " ORDER BY Calendar.Title"
        case 7  ' Product Title
          SQL = SQL & " ORDER BY Calendar.Product, Calendar.Title"
        case 8  ' Category, Sub Category
          SQL = SQL & " ORDER BY Calendar_Category.sort, Calendar_Category.Title, Calendar.Sub_Category, Calendar.Product,Calendar.BDate, Calendar.ID, dbo.Literature_Items_US.ACTIVE_FLAG, Calendar.Revision_Code Desc "        
        case 9          
          SQL = SQL & " ORDER BY Calendar_Category.sort, Calendar_Category.Title, Calendar.Product, Calendar.Sub_Category, Calendar.BDate, Calendar.ID, dbo.Literature_Items_US.ACTIVE_FLAG, Calendar.Revision_Code Desc "        
        case else
            SQL = SQL & " ORDER BY Calendar.Status, Calendar.Sub_Category, Calendar.Product, Calendar.Item_Number, dbo.Literature_Items_US.ACTIVE_FLAG, Calendar.Revision_Code Desc "
  end select    
 end if
  
Response.AddHeader "Content-Disposition", "attachment; filename=Asset_List.xls" 
Response.ContentType = "application/vnd.ms-excel"
Response.Charset     = "utf-16"
Set Rs = conn.Execute(SQL)
if Rs.eof <> true then
response.write "<table border=1>"
for iCol=0 to 6
        response.write "<td ><b>" & Rs.fields(iCol).Name & "</b></td>"
next
while not Rs.eof
    Response.Write "<tr>"
    for iCol=0 to 6
        response.write "<td>" & Server.HTMLEncode(Rs.fields(iCol).Value) & "</td>"
    next
    Response.Write "</tr>"
    Rs.movenext
wend
response.write "</table>"
end if
set rs=nothing

'if err.number <> 0 then
'    Response.Write err.Description
'end if  

%>
