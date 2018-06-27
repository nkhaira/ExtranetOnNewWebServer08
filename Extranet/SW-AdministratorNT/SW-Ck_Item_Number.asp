<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author: K. Whitlock
' Date:   1/13/2005
' --------------------------------------------------------------------------------------
%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%
Dim Asset_ID, AID, Item_Number
Dim Site_ID
Dim FormName
Dim SearchFlag, Search
Dim Found, Continue, errormsg
Dim Debug_Flag

Debug_Flag = false

Asset_ID       = request("Asset_ID")
AID            = request("AID")
Site_ID        = request("Site_ID")
Login_Language = request("Language")
FormName       = request("FormName")
Search         = request("Search")
Continue       = false

if Debug_Flag then

  with response
    .write "Asset_ID: " & Asset_ID & "<BR>"
    .write "AID: " & AID & "<BR>"
    .write "Site_ID: " & Site_ID & "<BR>"
    .write "Language: " & Login_Language & "<BR>"
    .write "Form: " & FormName & "<BR>"
    .write "Search: " & Search & "<BR>"
  end with
  
  response.flush
  response.end
  
end if  

if not isblank(Search) then SearchFlag = true else SearchFlag = false

with response
  .write "<HTML>" & vbCrLf
  .write "<HEAD>" & vbCrLf
  if SearchFlag then
    .write "<TITLE>Item Number Duplication</TITLE>" & vbCrLf
  else
    .write "<TITLE>Asset Container Search</TITLE>" & vbCrLf
  end if  
  .write "<LINK REL=STYLESHEET HREF=""/sw-common/SW-Style.css"">" & vbCrLf
  .write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=UTF-8"">" & vbCrLf
  .write "<META NAME=""LANG"" CONTENT=""ENGLISH"">" & vbCrLf
  .write "<META AUTHOR=""Kelly Whitlock - Kelly.Whitlock@fluke.com"">" & vbCrLf
  .write "</HEAD>" & vbCrLf
  .write "<BODY BGCOLOR=""White"" TOPMARGIN=""0"" LEFTMARGIN=""0"" MARGINWIDTH=""0"" MARGINHEIGHT=""0"" LINK=""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf
end with

call Connect_Sitewide

if not SearchFlag and not isblank(Asset_ID) then

  SQL = "SELECT Calendar.*, Calendar_Category.Title AS Category, UserData.FirstName AS FirstName, UserData.LastName AS LastName " &_
        "FROM Calendar " &_
        "LEFT OUTER JOIN UserData ON Calendar.Submitted_By = UserData.ID " &_
        "LEFT OUTER JOIN Calendar_Category ON Calendar.Category_ID = Calendar_Category.ID AND Calendar.Site_ID = Calendar_Category.Site_ID " &_
        "WHERE Calendar.Item_Number='" & Asset_ID & "' AND Calendar.Site_ID=" & Site_ID & " " &_
        "ORDER BY Calendar.Content_Group DESC, Calendar.Campaign"
        
   continue = true

else

  if SearchFlag and not isblank(Asset_ID) then
  
    if isnumeric(Asset_ID) and len(Asset_ID) = 7 then
  
      SQL = "SELECT Calendar.*, Calendar_Category.Title AS Category, UserData.FirstName AS FirstName, UserData.LastName AS LastName " &_
            "FROM Calendar " &_
            "LEFT OUTER JOIN UserData ON Calendar.Submitted_By = UserData.ID " &_
            "LEFT OUTER JOIN Calendar_Category ON Calendar.Category_ID = Calendar_Category.ID AND Calendar.Site_ID = Calendar_Category.Site_ID " &_
            "WHERE Calendar.Item_Number='" & Asset_ID & "' AND Calendar.Site_ID=" & Site_ID & " " &_
            "ORDER BY Calendar.Content_Group DESC, Calendar.Campaign, Calendar.ID"

      continue = true
      
    elseif isnumeric(Asset_ID) and len(Asset_ID) < 7 then

      SQL = "SELECT Calendar.*, Calendar_Category.Title AS Category, UserData.FirstName AS FirstName, UserData.LastName AS LastName " &_
            "FROM Calendar " &_
            "LEFT OUTER JOIN UserData ON Calendar.Submitted_By = UserData.ID " &_
            "LEFT OUTER JOIN Calendar_Category ON Calendar.Category_ID = Calendar_Category.ID AND Calendar.Site_ID = Calendar_Category.Site_ID " &_
            "WHERE (Calendar.ID=" & Asset_ID & " OR Calendar.Clone=" & Asset_ID & ") AND Calendar.Site_ID=" & Site_ID & " " &_
            "ORDER BY Calendar.Content_Group DESC, Calendar.Campaign, Calendar.ID"
  
      continue = true
      
    else

      SQL = "SELECT Calendar.*, Calendar_Category.Title AS Category, UserData.FirstName AS FirstName, UserData.LastName AS LastName " &_
            "FROM Calendar " &_
            "LEFT OUTER JOIN UserData ON Calendar.Submitted_By = UserData.ID " &_
            "LEFT OUTER JOIN Calendar_Category ON Calendar.Category_ID = Calendar_Category.ID AND Calendar.Site_ID = Calendar_Category.Site_ID " &_
            "WHERE (Calendar.Title like '%" & Replace(Asset_ID,"'","''") & "%') AND Calendar.Site_ID=" & Site_ID & " " &_
            "ORDER BY Calendar.Content_Group DESC, Calendar.Campaign, Calendar.ID"
      continue = true
      
    end if
    
  elseif SearchFlag and isblank(Asset_ID) then

    errormsg = Translate("You have supplied a blank or invalid Search parameter.",Login_Language,conn)
    
  else
  
    with response
    .write "<P>&nbsp;"
    .write "<FORM METHOD=""POST"" ACTION=""" & Server.Variables("SCRIPT_NAME") & """>" & vbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" Value=""" & Site_ID & """>" & vbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Login_Language"" Value=""" & Login_Language & """>" & vbCrLf    
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Search"" Value=""" & SearchFlag & """>" & vbCrLf    
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Form_Name"" Value=""" & Form_Name & """>" & vbCrLf    
    .write "<DIV ALIGN=CENTER>" & vbCrLf
    Call Nav_Border_Begin
    .write "<TABLE ALIGN=CENTER BORDER=0 BGCOLOR=White>" & vbCrLf
    .write "<TR>"
    .write "<TD COLSPAN=2 CLASS=MediumBold NOWRAP ALIGN=CENTER>" & Translate("Search for Item Number or Asset Container", Login_Language,conn) & "</TD>"
    .write "</TR>" & vbCrLf
    .write "<TR>"
    .write "<TD WIDTH=""50%"" CLASS=Medium NOWRAP>" & Translate("Oracle Item Number", Login_Language,conn) & ": </TD>"
    .write "<TD WIDTH=""50%"" CLASS=Medium NOWRAP><INPUT TYPE=TEXT NAME=""Asset_ID"" CLASS=Medium MAXLENGTH=""7"" SIZE=""10""></TD>" & vbCrLf
    .write "</TR>" & vbCrLf
    .write "<TR>"
    .write "<TD WIDTH=""50%"" CLASS=Medium NOWRAP>" & Translate("or", Login_Language,conn) & ": </TD>"
    .write "<TD WIDTH=""50%"" CLASS=Medium NOWRAP>&nbsp;</TD>" & vbCrLf
    .write "</TR>" & vbCrLf
    .write "<TR>"
    .write "<TD WIDTH=""40%"" CLASS=Medium>" & Translate("Asset ID Number", Login_Language,conn) & ": </TD>"
    .write "<TD WIDTH=""60%"" CLASS=Medium><INPUT TYPE=TEXT NAME=""AID"" CLASS=Medium MAXLENGTH=""7"" SIZE=""10""></TD>" & vbCrLf
    .write "</TR>" & vbCrLf
    .write "<TR>"
    .write "<TD WIDTH=""40%"" CLASS=Medium>&nbsp;</TD>"
    .write "<TD WIDTH=""60%"" CLASS=Medium><INPUT TYPE=SUBMIT NAME=""SUBMIT"" CLASS=NavLeftHighlight1 VALUE=""&nbsp;" & Translate("Go", Login_Language,conn) & "&nbsp;""></TD>" & vbCrLf
    .write "</TR>" & vbCrLf
    .write "</TABLE>"
    Call Nav_Border_End
    .write "</FORM>" & vbCrLf
    .write "</DIV>" & vbCrLf
    end with
  
  end if

end if

if continue then

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
    
    if not SearchFlag then
      .write "&nbsp;<P><SPAN CLASS=HEADING5>" & Translate("Alert",Login_Language,conn) & " - " & Translate("Item Number Duplication",Login_Language,conn) & "</SPAN><P>"
      .write Translate("The Item Number that you have just entered is already being used at this site in another container.",Login_Language,conn) & "&nbsp;&nbsp;"
      .write Translate("This alert <U>is not an error just status information</U> to prevent duplication of the same asset if an existing asset can be modified.",Login_Language,conn) & "&nbsp;&nbsp;"
      .write Translate("You can take the following actions:",Login_Language,conn)
      .write "<UL>"
      .write "<LI>" & Translate("Click on the [Close Window] button below then continue to add/update this container, however note that all other containers using the same Item Number and Revision, will use this version of the ""Low Resolution - Asset File"", ""POD Asset File"" and ""Thumbnail"", if the file names are the same and you are re-uploading the files to the site.",Login_Language,conn) & "</LI><P>"
      .write "<LI>" & Translate("Click on the [EDIT] button below to load an existing asset container.",Login_Language,conn) & " " & Translate("You can update or use the [CLONE] button to create a copy of the asset that you can modify.  If you use a [CLONE] asset, set the Content Grouping not to include ""Individual"" or ""+ Individual"" otherwise the asset will appear more than once on the site if the ""Groups allowed to view this information"" and ""Country Permissions/Restrictions"" are the same as the primary asset.",Login_Language,conn) & "</LI><P>"
      .write Translate("Clicking on the [EDIT] button will cancel this Add New/Update Content/Event action and reload the exisiting asset.",Login_Language,conn) & "</LI>"
      .write "</UL>"
    else
      .write "&nbsp;<P><SPAN CLASS=HEADING5>" & Translate("Search Results",Login_Language,conn) & "</SPAN><P>"
      .write Translate("You can take the following actions:",Login_Language,conn)
      .write "<UL>"
      .write "<LI>" & Translate("Click on the [EDIT] button below to load an existing asset container.",Login_Language,conn) & " " & Translate("You can update or use the [CLONE] button to create a copy of the asset that you can modify.  If you use a [CLONE] asset, set the Content Grouping not to include ""Individual"" or ""+ Individual"" otherwise the asset will appear more than once on the site if the ""Groups allowed to view this information"" and ""Country Permissions/Restrictions"" are the same as the primary asset.",Login_Language,conn) & "</LI><P>"
      .write "<LI>" & Translate("Click on the [Close Window] button below and perform no action.",Login_Language,conn) & "</LI><P>"
      .write "</UL>"
    end if
  
    .write "<DIV CLASS=Small ALIGN=CENTER>" & vbCrLf
    .write "<A HREF=""JavaScript=void(0);"" LANGUAGE=""JavaScript"" ONCLICK=""window.opener.focus(); self.close();""><SPAN Class=NavLeftHighlight1>&nbsp;" & Translate("Close Window",Login_Language,conn) & "&nbsp;</SPAN></A><P>" & vbCrLf
    
    Call Nav_Border_Begin
    
    .write "<TABLE BORDER=0 COLSPACING=0, COLPADDING=2 BGCOLOR=#666666>" & vbCrLf
    .write "<TR>"
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Action",Login_Language,conn) & "</TD>" & vbCrLf
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Asset ID",Login_Language,conn) & "</TD>" & vbCrLf
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Clone ID",Login_Language,conn) & "</TD>" & vbCrLf    
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Item Number",Login_Language,conn) & "</TD>" & vbCrLf
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Rev",Login_Language,conn) & "</TD>" & vbCrLf
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Lan",Login_Language,conn) & "</TD>" & vbCrLf    
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("CC",Login_Language,conn) & "</TD>" & vbCrLf    
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("MAC",Login_Language,conn) & "</TD>" & vbCrLf
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("MAC ID",Login_Language,conn) & "</TD>" & vbCrLf
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("POD",Login_Language,conn) & "</TD>" & vbCrLf  
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Category",Login_Language,conn) & "</TD>" & vbCrLf    
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Title",Login_Language,conn) & "</TD>" & vbCrLf
    .write "<TD BGCOLOR=BLACK CLASS=SmallBoldGold>" & Translate("Country",Login_Language,conn) & "</TD>" & vbCrLf    
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
      .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=RIGHT>"
      if rsItem("Clone") <> "0" then
        .write rsItem("Clone")
      else
        .write "&nbsp;"
      end if
      .write "</TD>" & vbCrLf      
      
      ' Item Number
      .write "<TD BGCOLOR=WHITE Class=Small ALIGN=RIGHT>" & rsItem("Item_Number") & "</TD>" & vbCrLf
      
      'Revision
      .write "<TD BGCOLOR=WHITE Class=Small ALIGN=CENTER>" & rsItem("Revision_Code") & "</TD>" & vbCrLf

      ' Language
      if UCase(rsItem("Language")) = "ENG" then
        .write "<TD BGCOLOR=WHITE Class=Small ALIGN=CENTER>"
      else
        .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=CENTER>"
      end if
      .write UCase(rsItem("Language")) & "</TD>" & vbCrLf
      
      ' Cost Center      
      .write "<TD BGCOLOR=WHITE Class=Small ALIGN=CENTER>"
      if rsItem("Cost_Center") > 0 then
        .write rsItem("Cost_Center")
      else
        .write "&nbsp;"
      end if
      .write "</TD>" & vbCrLf      
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
      .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=LEFT>"
      if SearchFlag then
        if not isnumeric(Asset_ID) then
          .write Highlight_Keyword(rsItem("Title"),Asset_ID, "SmallRed")
        else
          .write rsItem("Title")
        end if
      else  
        .write rsItem("Title")
      end if  
      .write "</TD>" & vbCrLf

      ' Country
      .write "<TD BGCOLOR=""#EAEAEA"" Class=Small ALIGN=LEFT>"
      if not isblank(rsItem("Country")) and LCase(rsItem("Country")) <> "none" then
        if instr(1,rsItem("Country"),"0,") > 0 then
          .write "<Font Color=""Red"">" & Translate("Limit to",Login_Language,conn) & ": " & Trim(Mid(rsItem("Country"),3)) & "</FONT>"
        else  
          .write "<Font Color=""Red"">" & Translate("Exclude",Login_Language,conn) & ": " & rsItem("Country") & "</FONT>"
        end if
      else
        .write Translate("No Exclusion",Login_Language,conn)
      end if
      .write "</TD>" & vbCrLf    
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

  elseif SearchFlag then
  
    errormsg = Translate("An Asset Container cannot be found based on the search criteria that you have supplied.",Login_Language,conn)
     
  end if
  
  rsItem.close
  set rsItem = nothing
  
  if SearchFlag = false then  

    response.write "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf

    if LCase(FormName) = "addcontent" then

      Call Connect_Sitewide
      SQL = "SELECT * FROM Literature_Items_US WHERE Item='" & Asset_ID & "' ORDER BY Active_Flag, Revision DESC"
      Set rsItem = Server.CreateObject("ADODB.Recordset")
      rsItem.Open SQL, conn, 3, 3

      if not rsItem.EOF then

        response.write "window.opener.document." & FormName & ".Title.value = '" & Replace(ProperCase(rsItem("Efulfillment")),"'","\'") & "';" & vbCrLf
        response.write "window.opener.document." & FormName & ".Revision_Code.value = '" & UCase(rsItem("Revision")) & "';" & vbCrLf
        response.write "window.opener.document." & FormName & ".Cost_Center.value = '" & rsItem("Cost_Center") & "';" & vbCrLf        
        response.write "window.opener.document." & FormName & ".Item_Number_Show.checked = true;" & vbCrLf
        
        if CInt(rsItem("PDF")) = CInt(True) or CInt(rsItem("POD")) = CInt(True) then
          response.write "window.opener.document." & FormName & ".SubGroups[0].checked = true;" & vbCrLf
          response.write "window.opener.document." & FormName & ".SubGroups[1].checked = true;" & vbCrLf        
        elseif CInt(rsItem("PDF")) = CInt(False) and CInt(rsItem("POD")) = CInt(True) then         
          response.write "window.opener.document." & FormName & ".SubGroups[0].checked = true;" & vbCrLf
          response.write "window.opener.document." & FormName & ".SubGroups[1].checked = false;" & vbCrLf        
        elseif CInt(rsItem("PDF")) = CInt(True) and CInt(rsItem("POD")) = CInt(False) then         
          response.write "window.opener.document." & FormName & ".SubGroups[0].checked = false;" & vbCrLf
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
          response.write "window.opener.document." & FormName & ".Content_Language.options[" & Counter + 1 & "].selected = true;" & vbCrLf
        end if
        
        rsLanguage.Close
        set rsLanguage = nothing
            
      end if
      
      rsItem.Close
      set rsItem = nothing
      
    end if

    if Found = false then    
      response.write "window.opener.focus(); self.close();" & vbCrLf
    end if
      
    response.write "</SCRIPT>" & vbCrLf
    
  end if

  Call ErrorStatus

else

  Call ErrorStatus

end if

sub ErrorStatus

  if Found = false then
  
    with response
      .write vbCrLf & "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf
      .write "self.focus();" & vbCrLf
      .write "</SCRIPT>" & vbCrLf
      
      .write "<TABLE ALIGN=CENTER WIDTH=""95%"" BORDER=0>" & vbCrLf
      .write "<TR><TD WIDTH=""100%"" CLASS=Medium>"
      .write "&nbsp;<P><SPAN CLASS=HEADING5>" & Translate("Search Results",Login_Language,conn) & "</SPAN><P>"
      .write errormsg & "<P>"
      .write "<A HREF=""JavaScript=void(0);"" LANGUAGE=""JavaScript"" ONCLICK=""window.opener.focus(); self.close();""><SPAN Class=NavLeftHighlight1>&nbsp;" & Translate("Close Window",Login_Language,conn) & "&nbsp;</SPAN></A><P>" & vbCrLf
      .write "</TD></TR></TABLE>"
    end with
  end if
end sub

Call Disconnect_Sitewide

%>
</BODY>
</HTML>
