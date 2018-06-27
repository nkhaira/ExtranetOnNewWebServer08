<HTML>
<HEAD>
  <TITLE>POD File Duplication</TITLE>
  <LINK REL=STYLESHEET HREF="/sw-common/SW-Style.css">
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
  <META NAME="LANG" CONTENT="ENGLISH">
  <META AUTHOR="Kelly Whitlock - Kelly.Whitlock@fluke.com">
</HEAD>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<BODY BGCOLOR="White" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" LINK="#000000" VLINK="#000000" ALINK="#000000">

<%
call Connect_Sitewide

Dim Asset_ID
Dim Site_ID

Asset_ID       = request("Asset_ID")
Item_Number    = request("Item_Number")
Site_ID        = request("Site_ID")
Login_Language = request("Language")

SQL = "SELECT Calendar.*, Calendar_Category.Title AS Category " &_
      "FROM Calendar " &_
      "LEFT OUTER JOIN Calendar_Category ON Calendar.Category_ID = Calendar_Category.ID AND Calendar.Site_ID = Calendar_Category.Site_ID " &_
      "WHERE Calendar.Item_Number='" & Item_Number & "' AND Calendar.Site_ID=" & Site_ID & " AND Calendar.File_Name_POD is not null " &_
      "ORDER BY Calendar.Content_Group DESC, Calendar.Campaign"

Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open SQL, conn, 3, 3

Found = False

do while not rsItem.EOF

  if rsItem("ID") = request("Asset_ID") then
  elseif not isblank(rsItem("File_Name_POD")) then
    Found = true
    exit do
  end if
  
  rsPod.MoveNext
loop

if Found = True then

  with response

  .write vbCrLf & "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf
  .write "self.focus();" & vbCrLf
  .write "</SCRIPT>" & vbCrLf
  
  .write "<TABLE ALIGN=CENTER WIDTH=""95%"" BORDER=0>" & vbCrLf
  .write "<TR><TD WIDTH=""100%"" CLASS=Medium>"
  .write "&nbsp;<P><SPAN CLASS=HEADING5>" & Translate("Alert",Login_Language,conn) & " - " & Translate("POD File Duplication",Login_Language,conn) & "</SPAN><P>"
  .write Translate("The POD File that you have just selected for upload is already being used in another container.",Login_Language,conn) & "&nbsp;&nbsp;"
  .write Translate("This message is to alert you to prevent duplication of the POD File asset file because an existing container can be updated.",Login_Language,conn) & "&nbsp;&nbsp;"
  .write Translate("You can take the following actions:",Login_Language,conn)
  .write "<UL>"
  .write "<LI>" & Translate("Click on the [EDIT] button below to load an existing container.  You can then update that container with the newer version of the POD File asset.",Login_Language,conn) & "</LI><P>"
  .write "<LI>" & Translate("Click on the [Close Window] button below then continue to Add/Update your original container, however you <U>must delete or [Unattach] the POD File Name</U> you found when you clicked on the [Browse] button, prior to [Updating].",Login_Language,conn) & "</LI><P>"
  .write "</UL>"

  .write "<DIV CLASS=Small ALIGN=CENTER>" & vbCrLf
  .write "<A HREF=""JavaScript=void(0);"" LANGUAGE=""JavaScript"" ONCLICK=""window.opener.focus(); self.close();""><SPAN Class=NavLeftHighlight1>&nbsp;" & Translate("Close Window",Login_Language,conn) & "&nbsp;</SPAN></A><P>" & vbCrLf
  
  Call Nav_Border_Begin
  
  .write "<TABLE BORDER=0 COLSPACING=0, COLPADDING=2 BGCOLOR=#666666>" & vbCrLf
  .write "<TR>"
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Action",Login_Language,conn) & "</TD>" & vbCrLf
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Asset ID",Login_Language,conn) & "</TD>" & vbCrLf
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Item Number",Login_Language,conn) & "</TD>" & vbCrLf
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Rev",Login_Language,conn) & "</TD>" & vbCrLf
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("PIC",Login_Language,conn) & "</TD>" & vbCrLf
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("POD Name",Login_Language,conn) & "</TD>" & vbCrLf
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("EFF",Login_Language,conn) & "</TD>" & vbCrLf    
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Category",Login_Language,conn) & "</TD>" & vbCrLf    
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Title",Login_Language,conn) & "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf
  
  rsItem.MoveFirst
    
  do while not rsItem.EOF
  
    if rsItem("ID") <> request("Asset_ID") then  
  
      Edit_Asset = "/SW-Administrator/Calendar_Edit.asp?ID=" & rsItem("ID") & "&Site_ID=" & Site_ID
      
      .write "<TR>" & vbCrLf
      .write "<TD BGCOLOR="
      select case rsItem("Status")
        case 1        
          .write """#00CC00"""
        case 2
          .write """#AAAAFF"""
        case else
          .write """Yellow"""
      end select
      .write " CLASS=Small ALIGN=CENTER>&nbsp;<A HREF=""JavaScript=void(0);"" ONCLICK=""window.opener.location.href='" & Edit_Asset & "'; window.opener.focus(); self.close(0);""><SPAN CLASS=NavLeftHighlight1>&nbsp;" & Translate("Edit",Login_Language,conn) & "&nbsp;</SPAN></A>&nbsp;</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=RIGHT>" & rsItem("ID") & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=WHITE Class=Small ALIGN=RIGHT>" & rsItem("Item_Number") & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=WHITE Class=Small ALIGN=CENTER>" & rsItem("Revision_Code") & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#EAEAEA"" ALIGN=CENTER CLASS=Small NOWRAP>"
      select case CInt(rsItem("Content_Group"))
        case 0
          .write "I"
        case 1
          .write "P+I"
        case 2
          .write "P"
        case 3
          .write "C+I"
        case 4
          .write "C"
        case else
          .write "&nbsp;"  
      end select
      .write "</TD>"
      .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=LEFT>" & rsItem("File_Name_POD") & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=CENTER>"
      if instr(1,LCase(rsItem("Subgroups")),"view") > 0 then
        .write "Y"
      else
        .write "&nbsp;"  
      end if
      .write "</TD>" & vbCrLf                          
      .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=LEFT>" & rsItem("Category") & "</TD>" & vbCrLf    
      .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=LEFT>" & rsItem("Title") & "</TD>" & vbCrLf    
      .write "</TR>" & vbCrLf
    end if  
    rsItem.MoveNext
  loop

  .write "</TABLE>" & vbCrLf
  Call Nav_Border_End
  
  .write "</TD>" & vbCrLf  
  .write "</TR>" & vbCrLf  
  .write "</TABLE>" & vbCrLf
    
  .write "</DIV>" & vbCrLf
  end with
end if

rsItem.close
set rsItem = nothing
call Disconnect_Sitewide

if Found = false then
  response.write "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf
  response.write "window.opener.focus(); self.close();" & vbCrLf
  response.write "</SCRIPT>" & vbCrLf
end if  

%>
</BODY>
</HTML>
