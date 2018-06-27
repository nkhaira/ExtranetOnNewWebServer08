<HTML>
<HEAD>
  <TITLE>Item Number Duplication</TITLE>
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
Dim FormName

Asset_ID       = request("Asset_ID")
Site_ID        = request("Site_ID")
Login_Language = request("Language")
FormName       = request("FormName")

SQL = "SELECT Calendar.*, Calendar_Category.Title AS Category, UserData.FirstName AS FirstName, UserData.LastName AS LastName " &_
      "FROM Calendar " &_
      "LEFT OUTER JOIN UserData ON Calendar.Submitted_By = UserData.ID " &_
      "LEFT OUTER JOIN Calendar_Category ON Calendar.Category_ID = Calendar_Category.ID AND Calendar.Site_ID = Calendar_Category.Site_ID " &_
      "WHERE Calendar.Item_Number='" & Asset_ID & "' AND Calendar.Site_ID=" & Site_ID & " " &_
      "ORDER BY Calendar.Content_Group DESC, Calendar.Campaign"
      
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open SQL, conn, 3, 3

Found = False

if not rsItem.EOF then
  Found = True
end if

if Found = True then

  with response

  .write vbCrLf & "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf
  .write "self.focus();" & vbCrLf
  .write "</SCRIPT>" & vbCrLf
  
  .write "<TABLE ALIGN=CENTER WIDTH=""95%"" BORDER=0>" & vbCrLf
  .write "<TR><TD WIDTH=""100%"" CLASS=Medium>"
  .write "&nbsp;<P><SPAN CLASS=HEADING5>" & Translate("Alert",Login_Language,conn) & " - " & Translate("Item Number Duplication",Login_Language,conn) & "</SPAN><P>"
  .write Translate("The Item Number that you have just entered is already being used at this site in another container.",Login_Language,conn) & "&nbsp;&nbsp;"
  .write Translate("This alert is <U>not an error</U> just status information to prevent duplication of the same asset if an existing asset can be modified.",Login_Language,conn) & "&nbsp;&nbsp;"
  .write Translate("You can take the following actions:",Login_Language,conn)
  .write "<UL>"
  .write "<LI>" & Translate("Click on the [Close Window] button below then continue to add/update this container, however note that all other containers using the same Item Number and Revision, will use this version of the ""Low Resolution - Asset File"", ""POD Asset File"" and ""Thumbnail"", if the file names are the same and you are re-uploading the files to the site.",Login_Language,conn) & "</LI><P>"
  .write "<LI>" & Translate("Click on the [EDIT] button below to load an existing container.",Login_Language,conn) & " " & Translate("You can update or use the [CLONE] button to create a copy of the asset that you can modify.  If you use a [CLONE] asset, set the Content Grouping not to include ""Individual"" or ""+ Individual"" otherwise the asset will appear more than once on the site if the ""Groups allowed to view this information"" and ""Country Permissions/Restrictions"" are the same as the primary asset.",Login_Language,conn) & "</LI><P>"
  .write Translate("Clicking on the [EDIT] button will cancel this Add New/Update Content/Event action and reload the exisiting asset.",Login_Language,conn) & "</LI>"
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
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("PIC ID",Login_Language,conn) & "</TD>" & vbCrLf
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("POD",Login_Language,conn) & "</TD>" & vbCrLf  
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Category",Login_Language,conn) & "</TD>" & vbCrLf    
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Title",Login_Language,conn) & "</TD>" & vbCrLf
  .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Owner",Login_Language,conn) & "</TD>" & vbCrLf  
  .write "</TR>" & vbCrLf
  
    
  do while not rsItem.EOF
  
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
    
    .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=RIGHT>"
    if rsItem("Campaign") <> 0 then
      Edit_Asset = "/SW-Administrator/Calendar_Edit.asp?ID=" & rsItem("Campaign") & "&Site_ID=" & Site_ID
      .write "<A HREF=""JavaScript=void(0);"" TITLE=""Edit PIC Container"" ONCLICK=""window.opener.location.href='" & Edit_Asset & "'; window.opener.focus(); self.close(0);"">" & rsItem("Campaign") & "</A>&nbsp;" & vbCrLf      
'      .write rsItem("Campaign")
    else
      .write "&nbsp;"
    end if
    .write "</TD>" & vbCrLf    
    if not isblank(rsItem("File_Name_Pod")) then
      .write "<TD BGCOLOR=""Yellow"" Class=Small ALIGN=CENTER>Y"    
    else
      .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=CENTER>&nbsp;"
    end if  
    .write "</TD>" & vbCrLf    
    .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=RIGHT>" & rsItem("Category") & "</TD>" & vbCrLf    
    .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=LEFT>" & rsItem("Title") & "</TD>" & vbCrLf    
    .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=LEFT>"    
    if not isblank(rsItem("LastName")) then
      .write "<SPAN CLASS=Smallest>"
      if not isblank(rsItem("FirstName")) then
        .write Mid(rsItem("FirstName"),1,1) & ". "
      end if
      .write rsItem("LastName") & "</SPAN>"
    else
      .write "&nbsp;"  
    end if    
    .write "</TD>" & vbCrLf    
    
    .write "</TR>" & vbCrLf
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
  
  if LCase(FormName) = "addcontent" then
  
    Call Connect_Sitewide
    SQL = "SELECT * FROM Literature_Items_US WHERE Item='" & Asset_ID & "'"
    Set rsItem = Server.CreateObject("ADODB.Recordset")
    rsItem.Open SQL, conn, 3, 3
    
    if not rsItem.EOF then
      response.write "window.opener.document." & FormName & ".Title.value = '" & ProperCase(rsItem("Efulfillment")) & "';" & vbCrLf
      response.write "window.opener.document." & FormName & ".Revision_Code.value = '" & UCase(rsItem("Revision")) & "';" & vbCrLf
      response.write "window.opener.document." & FormName & ".Item_Number_Show.checked = true;" & vbCrLf
      
      if CInt(rsItem("PDF")) = CInt(True) then
        response.write "window.opener.document." & FormName & ".SubGroups[0].checked = true;" & vbCrLf
        response.write "window.opener.document." & FormName & ".SubGroups[1].checked = true;" & vbCrLf        
      end if  

      SQL = "SELECT * FROM Language WHERE Oracle_Enable=-1 OR Enable=-1 ORDER BY Language.Sort"
      Set rsLanguage = Server.CreateObject("ADODB.Recordset")
      rsLanguage.Open SQL, conn, 3, 3
      
      Counter = -1
      do while not rsLanguage.EOF
        Counter = Counter + 1
        if LCase(rsItem("Language")) = LCase(rsLanguage("Code")) then
          exit do
        end if
        rsLanguage.MoveNext
      loop
      
      if Counter > -1 then
        response.write "window.opener.document." & FormName & ".Content_Language.options[" & Counter & "].selected = true;" & vbCrLf
      end if
      
      rsLanguage.Close
      set rsLanguage = nothing             
          
    end if
    
    rsItem.Close
    set rsItem = nothing
    Call Disconnect_Sitewide
    
  end if  
  
  response.write "window.opener.focus(); self.close();" & vbCrLf
  response.write "</SCRIPT>" & vbCrLf
  
end if  

%>
</BODY>
</HTML>
