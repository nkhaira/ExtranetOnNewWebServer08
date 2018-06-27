<HTML>
<HEAD>
<TITLE>Category / Sub Category Matrix</TITLE>
<LINK REL=STYLESHEET HREF="/SW-Common/SW-Style.css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso8859-1">
</HEAD>

<BODY BGCOLOR="White">

<!--#include virtual="/include/functions_String.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%
Call Connect_SiteWide

Site_ID = request("Site_ID")

if not isblank(Site_ID) then

  SQL       = "SELECT ID, Site_Description FROM Site WHERE ID=" & Site_ID
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3

  Site_Description = rsSite("Site_Description")  
  rsSite.close
  set rsSite = nothing

  SQL       = "SELECT Calendar.Site_ID, Calendar.Sub_Category, Calendar.Code, Calendar_Category.Title, Calendar.Language " & vbCrLf
  SQL = SQL & "FROM Calendar INNER JOIN Calendar_Category ON Calendar.Code=Calendar_Category.Code WHERE Calendar_Category.Site_ID=" & Site_ID & " " & vbCrLf
  SQL = SQL & "GROUP BY Calendar.Site_ID, Calendar.Sub_Category, Calendar.Code, Calendar.Language, Calendar_Category.Title " & vbCrLf
  SQL = SQL & "HAVING Calendar.Site_ID=" & Site_ID & " " & vbCrLf
  SQL = SQL & "AND Calendar.Language='eng' " & vbCrLf
  SQL = SQL & "ORDER BY Calendar_Category.Title, Calendar.Sub_Category" & vbCrLf
  
  Set rsSubCategory = Server.CreateObject("ADODB.Recordset")
  rsSubCategory.Open SQL, conn, 3, 3
  
  response.write "<SPAN CLASS=HEADING4>" & Site_Description & "</SPAN><BR>"
  response.write "<SPAN CLASS=SmallBOLD>Category / Sub Category - Matrix Listing</SPAN><P>"
  response.write "<SPAN CLASS=Small>The Category / SubCategory Matrix listing is an aid to ensure that you are<BR>adding content or event items into the correct sub-categories<BR>that have been pre-determined by your Site Administrator.<BR>If you need to add a new sub-category, see your Site Administrator.<P>"
  
  Call Table_Begin
  
  response.write "<TABLE BGCOLOR=Gray BORDER=0 CELLPADDING=4>"
  
  response.write "<TR>"
  response.write "<TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">Category</FONT></TD>"
  response.write "<TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">Sub Category</FONT></TD>"
  response.write "<TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">Preset</FONT></TD>"
  response.write "</TR>"
  
  Highlight = True
  
  Old_Category = ""
    
  do while not rsSubCategory.EOF
  
    New_Category = rsSubCategory("Title")  
    if rsSubCategory("Sub_Category") = ""  or isnull(rsSubCategory("Sub_Category")) then
      rsSubCategory.MoveNext
      if not rsSubCategory.EOF then
        if New_Category = rsSubCategory("Title") and (rsSubCategory("Sub_Category") <> ""  or not isnull(rsSubCategory("Sub_Category"))) then
        else
          rsSubCategory.MovePrevious
        end if
      else
        rsSubCategory.MovePrevious
      end if
    end if
  
    response.write "<TR>"
    
    if Old_Category <> rsSubCategory("Title") then
      if Highlight = True then
        Highlight = False
      else
        Highlight = True
      end if
    end if  
  
    ' Category Title
    response.write "<TD CLASS=SmallBold BGCOLOR="
    if Highlight = True then
      response.write """#E4E4E4"""
    else  
      response.write """White"""
    end if  
    response.write ">"
    if Old_Category = rsSubCategory("Title") then
      response.write "&nbsp;"
    else
      response.write rsSubCategory("Title")
    end if
    Old_Category = rsSubCategory("Title")      
    response.write "</TD>"
    
    ' Sub Category Title
    response.write "<TD CLASS=Small BGCOLOR="
    if Highlight = True then
      response.write """#E4E4E4"""
    else
      response.write """White"""
    end if  
    response.write ">"
    if rsSubCategory("Sub_Category") = "" or isnull(rsSubCategory("Sub_Category")) then
      response.write "&nbsp;"
    else  
      response.write rsSubCategory("Sub_Category")
    end if  
    response.write "</TD>"

    SQL = "SELECT Site_ID, Sub_Category, Code From Content_Sub_Category WHERE Site_ID=" & Site_ID & " AND Code=" & rsSubCategory("Code") & " AND Sub_Category='" & rsSubCategory("Sub_Category") & "'"
    Set rsSub = Server.CreateObject("ADODB.Recordset")
    rsSub.Open SQL, conn, 3, 3

    ' Type
    response.write "<TD ALIGN=CENTER CLASS=SmallBold BGCOLOR="
    if Highlight = True then
      response.write """#E4E4E4"""
    else
      response.write """White"""
    end if  
    response.write ">"

    if not rsSub.EOF then
      response.write "X"
    else
      response.write "&nbsp;"
    end if    
    response.write "</TD>"
    
    rsSub.close
    set rsSub = nothing
    
    response.write "</TR>"
    
    rsSubCategory.MoveNext
    
  loop
  
  response.write "</TABLE>"
  Call Table_End
  
  response.write "</BODY>"
  response.write "</HTML>"

else

  response.redirect 

end if
  
Call Disconnect_SiteWide

%> 