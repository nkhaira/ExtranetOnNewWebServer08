<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     K. Whitlock
' Date:       2/1/2000
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Session.timeout = 240 ' Set to 4 Hours
Server.ScriptTimeout = 600

Dim Freeze
Freeze = False          ' Disables Admin Tools = False for Software Update

if Request("Language") = "XON" then
  Session("ShowTranslation") = True
elseif Request("Language")="XOF" then
  Session("ShowTranslation") = False
end if

Session("BackURL_Calendar") = ""

Dim Site_ID
Dim Admin_Access
Dim Admin_Name
Dim FormName, FormName_0

FormName   = "None"
FormName_0 = "None_0"

Dim Page_Timer_Begin

Dim Border_Toggle
Border_Toggle = 0

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->  
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_SiteWide

' Check for specific Site ID else Check for Multiple Accounts if Site_ID is not known

%>
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%

if Freeze = True and (Admin_Access = 8 or Admin_Access = 9) then
  Freeze = False
end if  

if Request.QueryString("Site_ID") = "800" then
    Response.redirect "discount_edit.asp"
end if

if Request.QueryString("Subscription_Option_ID") = "4" then
  response.redirect "/SW-Administrator/site_utility.asp?ID=site_utility&Site_ID=" & Request.QueryString("Site_ID") & "&Utility_ID=54"
end if

' --------------------------------------------------------------------------------------
' Determine Site Code and Description based on all Admin Accounts for User
' --------------------------------------------------------------------------------------

if Site_ID = 100 then

  ' --------------------------------------------------------------------------------------
  ' Update Records for Archive
  ' --------------------------------------------------------------------------------------

  if isblank(Session("Archive")) then
  
    ' --------------------------------------------------------------------------------------
    ' The following code does some cleanup for LOS items with POD Fulfillment since DCG
    ' nor Oracle Literature DB can report this status change.
    ' --------------------------------------------------------------------------------------  

    SQL = "UPDATE dbo.Shopping_Cart_Lit SET Order_Ship_Date = Submit_Date WHERE Order_Ship_Date IS NULL AND Order_Number IS NOT NULL AND Order_Status=16"
    conn.execute SQL

    SQL = "SELECT DISTINCT Order_Number FROM dbo.Shopping_Cart_Lit WHERE Submit_Date < '" & DateAdd("d",-10,date()) & "' AND Order_Status <=5"
    Set rsStatus = Server.CreateObject("ADODB.Recordset")
    rsStatus.Open SQL, conn, 3, 3

    do while not rsStatus.EOF
      SQL = "UPDATE dbo.Shopping_Cart_Lit SET Order_Ship_Date='" & Date() & "', Order_Status=16 WHERE Order_Number='" & rsStatus("Order_Number") & "'"
      conn.execute SQL
      rsStatus.MoveNext
    loop
    rsStatus.close
    set rsStatus = nothing
    
    ' If Submit_Date is Null and Status on Item Number is not LIVE, auto delete from shopping cart because it is not orderable.

    SQL = "SELECT DISTINCT dbo.Shopping_Cart_Lit.ID AS SC_ID " &_
          "FROM            dbo.Shopping_Cart_Lit LEFT OUTER JOIN " &_
          "                dbo.Calendar ON dbo.Shopping_Cart_Lit.Item_Number = dbo.Calendar.Item_Number " &_
          "WHERE          (dbo.Calendar.Status <> 1) AND (dbo.Shopping_Cart_Lit.Submit_Date IS NULL) " &_
          "ORDER BY SC_ID"
    Set rsStatus = Server.CreateObject("ADODB.Recordset")
    rsStatus.Open SQL, conn, 3, 3

    do while not rsStatus.EOF
      SQL = "DELETE FROM dbo.Shopping_Cart_Lit WHERE ID=" & rsStatus("SC_ID")
      conn.execute SQL
      rsStatus.MoveNext
    loop
    rsStatus.close
    set rsStatus = nothing
    
    ' If Asset has been deleted, then delete shopping cart items from cart where for that asset
    SQL = "SELECT     dbo.Shopping_Cart_Lit.ID AS SC_ID " &_
          "FROM       dbo.Shopping_Cart_Lit LEFT OUTER JOIN " &_
          "           dbo.Calendar ON dbo.Shopping_Cart_Lit.Asset_ID = dbo.Calendar.ID " &_
          "WHERE     (dbo.Calendar.ID IS NULL) AND (dbo.Shopping_Cart_Lit.Submit_Date IS NULL)"
    Set rsStatus = Server.CreateObject("ADODB.Recordset")
    rsStatus.Open SQL, conn, 3, 3

    do while not rsStatus.EOF
      SQL = "DELETE FROM dbo.Shopping_Cart_Lit WHERE ID=" & rsStatus("SC_ID")
      conn.execute SQL
      rsStatus.MoveNext
    loop
    rsStatus.close
    set rsStatus = nothing
    set SQL      = nothing

    Session("Archive") = "Updated"
    
  end if
  
  ' Last bit of house cleaning.  For those records assigned a non exsistant FCM_ID, delete reference.

  SQLFCM =  "SELECT DISTINCT FCM_ID " &_
            "FROM dbo.UserData " &_
            "WHERE Fcm_ID <> 0 AND FCM_ID not in( " &_
            "SELECT ID " &_
            "FROM dbo.UserData " &_
            "WHERE ID IN(SELECT DISTINCT FCM_ID " &_
            "FROM dbo.UserData " &_
            "WHERE Fcm_ID <> 0))"
                  
  Set rsFCM = Server.CreateObject("ADODB.Recordset")
  rsFCM.Open SQLFCM, conn, 3, 3
    
  do while not rsFCM.EOF
    
    SQLFCM = "UPDATE UserData SET Fcm_ID=0 WHERE Fcm_ID=" & rsFCM("FCM_ID")
    conn.execute SQLFCM
      
    rsFCM.MoveNext
      
  loop
    
  rsFCM.close
  set rsFCM  = nothing
  set SQLFCM = nothing
  
  ' Synchronize Oracle Cost Centers with Asset Item Numbers
  
  SQLCC = "SELECT DISTINCT dbo.Calendar.Item_Number, dbo.Calendar.Revision_Code, dbo.Literature_Items_US.ITEM, dbo.Literature_Items_US.COST_CENTER " &_
          "FROM            dbo.Calendar LEFT OUTER JOIN " &_
          "                dbo.Literature_Items_US ON dbo.Calendar.Item_Number = dbo.Literature_Items_US.ITEM " &_
          "WHERE          (LEN(dbo.Calendar.Item_Number) = 7) AND (LEN(dbo.Literature_Items_US.ITEM) = 7) AND len(dbo.Literature_Items_US.ITEM)=4" &_
          "ORDER BY dbo.Calendar.Revision_Code"

  Set rsCC = Server.CreateObject("ADODB.Recordset")
  rsCC.Open SQLCC, conn, 3, 3

  do while not rsCC.EOF
  
    SQLCC = "UPDATE Calendar SET Cost_Center=" & rsCC("Cost_Center") & " WHERE Item_Number='" & rsCC("Item_Number") & "'"
    conn.execute SQLCC
  
    rsCC.MoveNext
  
  loop
  
  rsCC.close
  set rsCC  = nothing
  set SQLCC = nothing
  
  ' --------------------------------------------------------------------------------------  

  ' If only one valid site then redirect
  
  SQL = "SELECT dbo.UserData.Site_ID, dbo.UserData.NTLogin, dbo.UserData.NewFlag, dbo.UserData.SubGroups " &_
        "FROM   dbo.UserData LEFT OUTER JOIN " &_
        "dbo.Site ON dbo.UserData.Site_ID = dbo.Site.ID " &_
        "WHERE (dbo.UserData.NTLogin = '" & Admin_Name & "') AND (dbo.UserData.NewFlag = 0) AND " &_
        "      (dbo.UserData.SubGroups LIKE '%administrator%' OR " &_
        "       dbo.UserData.SubGroups LIKE '%domain%' OR " &_
        "       dbo.UserData.SubGroups LIKE '%content%' OR " &_
        "       dbo.UserData.SubGroups LIKE '%submitter%') " &_
        "      AND (dbo.Site.Enabled = - 1)"

  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3

  if rsSite.RecordCount = 1 then
    response.redirect "/sw-administrator/default.asp?Site_ID=" & rsSite("Site_ID")
  end if
  
  rsSite.close
  set rsSite = nothing
  
  Screen_Title   = Translate("Extranet Support Site",Alt_Language,conn) & " - " & Translate("Administrator&acute;s Site Selection",Alt_Language,conn)
  Bar_Title      = Translate("Extranet Support Site",Login_Language,conn) & "<BR><FONT Class=MediumBoldGold>" & Translate("Administrator&acute;s Site Selection",Login_Language,conn) & "</FONT>"
  Navigation     = false
  Top_Navigation = false
  Content_Width  = 95  ' Percent

  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
  <%
  
  response.write "<FONT CLASS=NormalBold>" & FormatFullName(Admin_FirstName, Admin_MiddleName, Admin_LastName) & "</FONT><BR><BR>"

  SQL = "SELECT *, Site.Site_Code FROM Site WHERE Site.Enabled=" & CInt(True) & " ORDER BY Site.Site_Description"
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3
   ''response.write sql 
  response.write "<FORM NAME=""Dummy-0"">"
  Call Nav_Border_Begin
  response.write "<TABLE CELLPADDING=2 CELLSPACING=4>"
  response.write "<TR>"
  response.write "<TD BGCOLOR=""#FFCC00"" CLASS=Medium>"
  response.write Translate("Extranet Support Site",Login_Language,conn) & ":"
  response.write "</TD>"
  
  response.write "<TD CLASS=Medium>"
  %>
  <SELECT  Class=Medium LANGUAGE="JavaScript" ONCHANGE="window.location.href='default.asp?Site_ID='+this.options[this.selectedIndex].value" NAME="Site_ID">
  <%
  response.write "<OPTION Class=Medium VALUE="""">" & Translate("Select Site from List",Login_Language,conn) & "</OPTION>"
  
  Do while not rsSite.EOF
  
    SQL = "SELECT Site_ID, NTLogin, NewFlag, SubGroups FROM UserData WHERE Site_ID=" & rsSite("ID") & " AND NTLogin='" & Admin_Name & "' AND NewFlag=" & CInt(False)
    Set rsAccess = Server.CreateObject("ADODB.Recordset")
    rsAccess.Open SQL, conn, 3, 3

    if not rsAccess.EOF then
      if instr(1,rsAccess("SubGroups"),"domain") > 0 or _
         instr(1,rsAccess("SubGroups"),"administrator") > 0 or _
         instr(1,rsAccess("SubGroups"),"account") > 0 or _
         instr(1,rsAccess("SubGroups"),"content") > 0 or _
         instr(1,rsAccess("SubGroups"),"branch") > 0 or _
         instr(1,rsAccess("SubGroups"),"submitter") > 0 or _
         instr(1,rsAccess("SubGroups"),"literature") > 0 or _         
         instr(1,rsAccess("SubGroups"),"forum") > 0 then
           response.write "<OPTION Class=Medium VALUE=""" & rsSite("ID") & """>" & rsSite("Site_Description") & "</OPTION>"
      end if
    end if  
      
    rsAccess.close
    set rsAccess = nothing
                       
    rsSite.MoveNext 
    
  loop
            
  rsSite.close
  set rsSite=nothing

  response.write "</SELECT>"
  response.write "</TD>"
  response.write "</TR>"
  response.write "</TABLE>"
  Call Nav_Border_End
  response.write "</FORM>"
  
  response.write "<UL>"
  response.write "<LI>" & Translate("Since you have logged on directly to the Extranet Administrator Tool Kit as opposed to accessing this site through your administrator&acute; link provided for via the site&acute;s navigation button, please select the appropriate Extranet Site that you wish to administrative access to, from drop-down above.",Login_Language,conn) & "</LI><BR><BR>"
  response.write "<LI>" & Translate("Problems or questions about this site or site tools should be directed to",Login_Language,conn) & " <A HREF=""mailto:Webmaster@fluke.com"">" & Translate("Webmaster - Fluke Extranet Sites",Login_Language,conn) & "</A>.</LI><BR><BR>"
  response.write "<LI>" & Translate("If you are reporting a problem, please provide the URL and a complete description of the problem by using a copy of the error message or a screen capture.",Login_Language,conn) & "</A></LI>"
  response.write "</UL>"

else 

  ' --------------------------------------------------------------------------------------
  ' Determine Site Code and Description based on Site_ID Number 
  ' --------------------------------------------------------------------------------------
  
  SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID  'CInt(request("Site_ID"))
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3

  Site_Code        = rsSite("Site_Code")      
  Site_Description = rsSite("Site_Description")
  Logo             = rsSite("Logo")  
  Logo_Left        = rsSite("Logo_Left")

  rsSite.close
  set rsSite=nothing
  
  ' --------------------------------------------------------------------------------------
    
  Screen_Title   = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Site Administration Main Menu",Alt_Language,conn)
  Bar_Title      = Translate(Site_Description,Login_Language,conn) & "<BR><FONT Class=MediumBoldGold>" & Translate("Site Administration Main Menu",Login_Language,conn) & "</FONT>"
  Navigation     = false
  Top_Navigation = false
  Content_Width  = 95  ' Percent
  
  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
  <%

  response.write "<FONT CLASS=MediumBoldRed>"
  select case Admin_Access
    case 2
      response.write Translate("Content Submitter",Login_Language,conn)
    case 3
      response.write Translate("Literature Order System Administrator",Login_Language,conn)
    case 4
      response.write Translate("Content Administrator",Login_Language,conn)
    case 6
      response.write Translate("Account Administrator",Login_Language,conn)
    case 8
      response.write Translate("Site Administrator",Login_Language,conn)
    case 9
      response.write Translate("Domain Administrator",Login_Language,conn)
  end select
  response.write "</B></FONT><FONT CLASS=MediumBold><BR>" & Admin_FirstName & " " & Admin_LastName & "<BR>" & Admin_Company & "</FONT><BR><BR>"  

  ' --------------------------------------------------------------------------------------
  ' Administration Tools Container - Begin
  ' --------------------------------------------------------------------------------------

  ' Logoff
  Call Nav_Border_Begin
  response.write "<A HREF=""/register/default.asp"" onclick=""location.href='/register/default.asp';"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Logoff",Login_language,conn) & "&nbsp;&nbsp;</A>"
  response.write "&nbsp;&nbsp;<A HREF=""default.asp?Site_ID=100"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Site Menu",Login_language,conn) & "&nbsp;&nbsp;</A>"
  Call Nav_Border_End
  
  if Freeze = True then

    response.write "<FONT CLASS=NormalBoldRed>The Site Administration Tools are unavilable at this time.</FONT><BR><BR>"

  else

    FormName_0 = "Main_Select"
    response.write "<FORM NAME=""" & FormName_0 & """>"
    Call Table_Begin
    response.write "<TABLE WIDTH=""100%"" CELLPADDING=4 BORDER=0 CLASS=TableBackground>"
          
    ' --------------------------------------------------------------------------------------
    ' Open Close Site
    ' --------------------------------------------------------------------------------------
  
    if Admin_Access = 9 or Admin_Access = 8 then
            
      response.write "<TR>"
      response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      response.write "<B>" & Translate("Site",Login_Language,conn) & "</B> - " & Translate("Open / Closed Status",Login_language,conn) & ":"
      response.write "</TD>"
  
      if not isblank(request("Site_Closed")) then
        Site_Closed = CInt(request("Site_Closed"))
        SQL = "UPDATE Site SET Site.Closed=" & CInt(Site_Closed) & " WHERE (((Site.ID)=" & CInt(Site_ID) & "))"
        conn.execute (SQL)
      end if  
  
      if isblank(Site_Closed) then              
        SQL = "SELECT Site.Closed FROM Site WHERE Site.ID=" & CInt(Site_ID)
        Set rsSite = Server.CreateObject("ADODB.Recordset")
        rsSite.Open SQL, conn, 3, 3
        Site_Closed = CInt(rsSite("Closed"))                
        rsSite.Close
        set rsSite = nothing
      end if
  
      response.write "<TD BGCOLOR="""
      if Site_Closed = True then
        response.write "Red"
      else
        response.write Contrast
      end if
      response.write """ CLASS=Medium>"                 
  
      %>
      <SELECT Class=Medium LANGUAGE="JavaScript" ONCHANGE="window.location.href='default.asp?Site_ID=<%=Site_ID%>&Site_Closed='+this.options[this.selectedIndex].value" NAME="Site_Closed">    
      <%                                                    
  
  	  response.write "<OPTION CLASS=SmallGreen VALUE=""" & CInt(False) & """"
      if Site_Closed = False then
        response.write " SELECTED"
      end if
      response.write ">" & Translate("Open",Login_language,conn) & "</OPTION>"
  
   	  response.write "<OPTION CLASS=SmallRed VALUE=""" & CInt(True) & """"
      if Site_Closed = True then
        response.write " SELECTED"
      end if
      response.write ">" & Translate("Closed",Login_language,conn) & "</OPTION>"   
  
      response.write "</SELECT>"
      response.write "</TD>"
      response.write "</TR>"
  
    end if
  
    ' --------------------------------------------------------------------------------------
    ' NT Accounts Edit
    ' --------------------------------------------------------------------------------------
  
    if Admin_Access >= 6 then
    
      response.write "<TR>"
      response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
  
      response.write "<B>" & Translate("Edit",Login_Language,conn) & "</B> - " & Translate("Users Account Profile",Login_Language,conn) & ":"
      response.write "</TD>"
  
      SQL = "SELECT UserData.NewFlag, UserData.ExpirationDate FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND (UserData.NewFlag<>" & CInt(False) & " OR UserData.ExpirationDate<='" & Date & "')"
      
      ' Limit Alert Notification to Region Only if User is Only an Account Administrator
      
      if Admin_Access = 6 then
        SQLRegion = "SELECT Approvers_Account.* FROM Approvers_Account WHERE Approvers_Account.Site_ID=" & CInt(Site_ID)
        Set rsRegion = Server.CreateObject("ADODB.Recordset")
        rsRegion.Open SQLRegion, conn, 3, 3
        
        Region_Toggle = False
        if not rsRegion.EOF then
          do while not rsRegion.EOF
            if CLng(rsRegion("Approver_ID")) = CLng(Admin_ID) then
              if Region_Toggle = False then       
                SQL = SQL & " AND (Region=" & rsRegion("Region")
                Region_Toggle = True
              else  
                SQL = SQL & " OR Region=" & rsRegion("Region")
              end if
            end if
            rsRegion.MoveNext
          loop
          if Region_Toggle = True then
            SQL = SQL & ")"
          end if  
        end if
        rsRegion.close
        set rsRegion = nothing
      end if        
  
      ' Now Check for Active Accounts needing Approval or Expired
      
      ' response.write SQL & "<BR><BR>"
      
      Set rsAccount = Server.CreateObject("ADODB.Recordset")
      rsAccount.Open SQL, conn, 3, 3
  
      if not rsAccount.EOF then
        response.write "<TD BGCOLOR=""Red"" CLASS=Medium>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
      
      rsAccount.close
      set rsAccount = nothing
      
      SC2 = -1  ' Set to bogus number to envoke Select from list message
      %>
      <!--#include virtual="/SW-Administrator/Account_List_Query_Criteria.asp"-->
      <%
  
      response.write "</TD>"
      response.write "</TR>"
  
    end if
    
    ' --------------------------------------------------------------------------------------
    ' Add Content or Event
    ' --------------------------------------------------------------------------------------
  
    if Admin_Access = 9 or Admin_Access = 8 or Admin_Access = 4 or Admin_Access = 2 then
  
      response.write "<TR>"
      response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      response.write "<B>" & Translate("Add",Login_language,conn) & "</B> - " & Translate("Content or Event into Category",Login_language,conn) & ":"
      response.write "</TD>"
      
      response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      %>
      <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='Calendar_edit.asp?ID=add&Site_ID=<%=Site_ID%>&Category_ID='+this.options[this.selectedIndex].value" NAME="Category_ID">    
      <%
      
      if Admin_Access <> 2 then
        SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=" & CInt(True) & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
      else
        SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=" & CInt(True) & " AND Calendar_Category.Code<8000 ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
      end if
      
      Set rsCategory = Server.CreateObject("ADODB.Recordset")
      rsCategory.Open SQL, conn, 3, 3
        
      response.write "<OPTION Class=Small VALUE="""">" & Translate("Select from list",Login_language,conn) & "</OPTION>"
                                  
      Do while not rsCategory.EOF
        select case rsCategory("Code")
          case 8000
            response.write "<OPTION Class=Region1"          
          case 8001
            response.write "<OPTION Class=Region2"          
          case else
            response.write "<OPTION Class=Medium"
        end select
            
          response.write " VALUE=""" & rsCategory("ID") & """>" & Translate(RestoreQuote(rsCategory("Title")),Login_Language,conn) & "</OPTION>"                    
      	  rsCategory.MoveNext 
      loop
                
      rsCategory.close
      set rsCategory=nothing
     
      response.write "</SELECT>"
      response.write "</TD>"
      response.write "</TR>"
  
    end if
    
    ' --------------------------------------------------------------------------------------
    ' Edit Content or Event Category Selection
    ' --------------------------------------------------------------------------------------
  
    if Admin_Access = 2 or Admin_Access = 4 or Admin_Access = 8  or Admin_Access = 9 then
    
      response.write "<TR>"
  
      if request("ID") = "edit_record" then
        response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
  
      response.write "<B>" & Translate("Edit",Login_language,conn) & "</B> - " & Translate("Content or Event from Category",Login_language,conn) & ":"
      response.write "</TD>"
  
      if Admin_Access = 4 OR Admin_Access = 8 OR Admin_Access = 9 then
  
        'Check Category Items/Events Pending Approval
  
        SQL =       "SELECT Calendar.* "
        SQL = SQL & "FROM Calendar "
        SQL = SQL & "WHERE Calendar.Site_ID=" & CInt(Site_ID) & " "
        SQL = SQL & "AND Calendar.Status=0" & " "
        SQL = SQL & "AND Calendar.Review_By=" & CLng(Admin_ID)
  
        Set rsApproval = Server.CreateObject("ADODB.Recordset")
        rsApproval.Open SQL, conn, 3, 3  
  
        if not rsApproval.EOF then
          response.write "<TD BGCOLOR=""#FF0000"" CLASS=Medium>"      ' Red
        else
          if request("ID") = "edit_record" then
            response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"  ' Black
          else
            response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"    ' Gold
          end if  
        end if
    
        rsApproval.close
        set rsApproval = nothing
        
      else
          if request("ID") = "edit_record" then
            response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"  ' Black
          else
            response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
          end if
      end if
      %>        
      <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='default.asp?ID=edit_record&Site_ID=<%=Site_ID%>&Category_ID='+this.options[this.selectedIndex].value+'#Results'" NAME="Category_ID">    
      <%
      if request("ID") <> "edit_record" then    
        response.write "<OPTION Class=Small VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>"
      end if
  
      ' Individual Approval Queues
      
      if isblank(request("Category_ID")) then
        CatID = 0
      else
        CatID = request("Category_ID")
        response.write CatID
      end if
      
      if request("ID") = "edit_record" and CLng(CatID) = 9998 then
        if Admin_Access = 2 then
          response.write "<OPTION VALUE=""9998"" CLASS=NavLeftHighlight1 SELECTED>" & Translate("View Submit Queue",Login_Language,conn) & "</OPTION>"
        else
          response.write "<OPTION VALUE=""9998"" CLASS=NavLeftHighlight1 SELECTED>" & Translate("View Approval Queue",Login_Language,conn) & "</OPTION>"                  
        end if  
      else
        if Admin_Access = 2 then
          response.write "<OPTION VALUE=""9998"" CLASS=NavLeftHighlight1>" & Translate("View Submit Queue",Login_Language,conn) & "</OPTION>"
        else
          response.write "<OPTION VALUE=""9998"" CLASS=NavLeftHighlight1>" & Translate("View Approval Queue",Login_Language,conn) & "</OPTION>"                  
        end if  
      end if
  
      ' All Approval Queues
      
      if request("ID") = "edit_record" and CLng(CatID) = 9999 then
        if Admin_Access = 4 or Admin_Access = 8 or Admin_Access = 9 then
          response.write "<OPTION VALUE=""9999"" CLASS=NavLeftHighlight1 SELECTED>" & Translate("View Approval Queue - All",Login_Language,conn) & "</OPTION>"
        end if
      else
        if Admin_Access = 4 or Admin_Access = 8 or Admin_Access = 9 then
          response.write "<OPTION VALUE=""9999"" CLASS=NavLeftHighlight1>" & Translate("View Approval Queue - All",Login_Language,conn) & "</OPTION>"                  
        end if  
      end if
  
      Select Case Admin_Access
        Case 2,4,8,9
  
          SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=" & CInt(True) & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
          Set rsCategory = Server.CreateObject("ADODB.Recordset")
          rsCategory.Open SQL, conn, 3, 3
                              
          Do while not rsCategory.EOF

            select case rsCategory("Code")
              case 8000
                response.write "<OPTION Class=Region1"          
              case 8001
                response.write "<OPTION Class=Region2"          
              case else
                response.write "<OPTION Class=Medium"
            end select
          
            if request("ID") = "edit_record" and CLng(CatID) = rsCategory("ID") then
           	  response.write " SELECTED VALUE=""" & rsCategory("ID") & """>" & Translate(RestoreQuote(rsCategory("Title")),Login_Language,conn) & "</OPTION>"
            else
           	  response.write " VALUE=""" & rsCategory("ID") & """>" & Translate(RestoreQuote(rsCategory("Title")),Login_Language,conn) & "</OPTION>"
            end if                
        	  rsCategory.MoveNext 
          loop
            
          rsCategory.close
          set rsCategory=nothing
  
      end select

      response.write "</SELECT>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
      
      if request("ID") = "edit_record" then
        response.write "<SPAN Class=MediumGold>"
      else
        response.write "<SPAN CLASS=MediumBold>"
      end if
      
      response.write Translate("Search",Login_Language,conn) & ": "
      response.write "</SPAN>"
      response.write "<INPUT TYPE=TEXT CLASS=Small NAME=""Search_Parameter_0"" Value="""" MAXLENGTH=""50"" SIZE=""20"" onkeypress=""return handleEnter(this, event);"">&nbsp;"
      response.write "<INPUT TYPE=BUTTON Class=NavLeftHighlight1 Value="" " & Translate("Go",Login_Language,conn) & " "" LANGUAGE=""JavaScript"" ONCLICK=""SearchForItem_0();"" TITLE=""Search for Asset by ID, Item Number or Keyword"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">&nbsp;" & vbCrLf
      
      response.write "</TD>"
      response.write "</TR>"
  
    end if
  
    ' --------------------------------------------------------------------------------------
    ' Category Options
    ' --------------------------------------------------------------------------------------
  
    if Admin_Access = 9 or Admin_Access = 8 then
  
      response.write "<TR>"
  
      if request("ID") = "edit_category" then
        response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
  
      response.write "<B>" & Translate("Edit",Login_Language,conn) & "</B> - " & Translate("Category Options",Login_Language,conn) & ":"
      response.write "</TD>"
  
      if request("ID") = "edit_category" then
        response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
      %>
      <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='default.asp?ID=edit_category&Site_ID=<%=Site_ID%>&Category_ID='+this.options[this.selectedIndex].value+'#Results'" NAME="Category_ID">    
      <%
      SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
      Set rsCategory = Server.CreateObject("ADODB.Recordset")
      rsCategory.Open SQL, conn, 3, 3
      
      if request("ID") <> "edit_category" then
        response.write "<OPTION Class=Small VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>"                              
      end if
        
      Do while not rsCategory.EOF
        if CInt(request("Category_ID")) = rsCategory("ID") and request("ID") = "edit_category" then
       	  response.write "<OPTION SELECTED VALUE=""" & rsCategory("ID") & """"
          if rsCategory("Enabled") = True then
            response.write " CLASS=Region1NavMedium>+ "
          else
            response.write " CLASS=RegionXNavMedium>o "
          end if
          response.write Translate(RestoreQuote(rsCategory("Title")),Login_Language,conn) & "</OPTION>"
  
        else              
       	  response.write "<OPTION VALUE=""" & rsCategory("ID") & """"
          if rsCategory("Enabled") = True then                  
            response.write " CLASS=Region1NavMedium>+ "
          else
            response.write " CLASS=RegionXNavMedium>o "
          end if
          response.write Translate(RestoreQuote(rsCategory("Title")),Login_Language,conn) & "</OPTION>"
        end if              
    	  rsCategory.MoveNext 
      loop
      
      rsCategory.close
      set rsCategory=nothing
            
      response.write "</SELECT>"
      response.write "</TD>"
      response.write "</TR>"
                      
      ' --------------------------------------------------------------------------------------
      ' Subgroups
      ' --------------------------------------------------------------------------------------
    
      response.write "<TR>"
  
      if request("ID") = "edit_group" then
        response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"          
      end if
  
      response.write "<B>" & Translate("Edit",Login_Language,conn) & "</B> - " & Translate("Group Options",Login_Language,conn) & ":"
      response.write "</TD>"
  
      if request("ID") = "edit_group" then
        response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
      %>
      <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='default.asp?ID=edit_group&Site_ID=<%=Site_ID%>&Group_ID='+this.options[this.selectedIndex].value+'#Results'" NAME="Group_ID">    
      <%
      SQL = "SELECT SubGroups.* FROM SubGroups WHERE SubGroups.Site_ID=" & CInt(Site_ID) & " AND SubGroups.Order_Num <> 99 ORDER BY SubGroups.Order_Num"
      Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
      rsSubGroups.Open SQL, conn, 3, 3
  
      if request("ID") <> "edit_group" then    
        response.write "<OPTION Class=Small VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
      end if
                        
      Do while not rsSubGroups.EOF            
        if CInt(request("Group_ID")) = rsSubGroups("ID") and request("ID") = "edit_group" then
       	  response.write "<OPTION SELECTED VALUE=""" & rsSubGroups("ID") & """"
          if rsSubGroups("Enabled") = True then
            response.write " CLASS=Region" & Trim(rsSubGroups("Region")) & "NavMedium>+ "
          else
            response.write " CLASS=RegionXNavMedium>o "
          end if
          response.write RestoreQuote(rsSubGroups("X_Description")) & "</OPTION>" & vbCrLf
        else
       	  response.write "<OPTION VALUE=""" & rsSubGroups("ID") & """"
          if rsSubGroups("Enabled") = True then
            response.write " CLASS=Region" & Trim(rsSubGroups("Region")) & "NavMedium>+ "                  
          else
            response.write " CLASS=RegionXNavMedium>o "
          end if
          response.write RestoreQuote(rsSubGroups("X_Description")) & "</OPTION>" & vbCrLf
        end if
    	  rsSubGroups.MoveNext 
      loop
      
      rsSubGroups.close
      Set rsSubGroups = Nothing            
  
      response.write "</SELECT>"
      response.write "</TD>"
      response.write "</TR>"
          
    ' --------------------------------------------------------------------------------------
    ' Subscription Options
    ' --------------------------------------------------------------------------------------
  
    if Admin_Access = 9 or Admin_Access = 8 then
  
      response.write "<TR>"
  
      if request("ID") = "edit_subscription" then
        response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
  
      response.write "<B>" & Translate("Edit",Login_Language,conn) & "</B> - " & Translate("Subscription Options",Login_Language,conn) & ":"
      response.write "</TD>"
  
      if request("ID") = "edit_subscription" then
        response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
      %>
      <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='default.asp?ID=edit_subscription&Site_ID=<%=Site_ID%>&Subscription_Option_ID='+this.options[this.selectedIndex].value+'#Results'" NAME="Subscription_Option_ID#Subscription_Edit">    
      <%
      if request("ID") <> "edit_subscription" then
        response.write "<OPTION Class=Small VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>"                              
      end if
      
      for x = 1 to 4
        response.write "<OPTION VALUE=""" & x & """ CLASS="
        
        select case x
          case 4
            response.write "Region2NavMedium"
          case else
            response.write "Region1NavMedium"
        end select
            
        if CInt(request("Subscription_Option_ID")) = x and request("ID") = "edit_subscription" then
          response.write " SELECTED"
        end if
        response.write ">"
        select case x
          case 1
            response.write Translate("Subject Line",Login_Language,conn)
          case 2
            response.write Translate("Body Alternate-Header",Login_Language,conn)
          case 3
            response.write Translate("Body Alternate-Footer",Login_Language,conn)
          case 4
            response.write Translate("Subscription Service - Queue",Login_Language,conn)
        end select
        response.write "</OPTION>" & vbCrLf
      next
      response.write "</SELECT>" & vbCrLf
      response.write "</TD>"
      response.write "</TR>"
      
    end if

      ' --------------------------------------------------------------------------------------
      ' Auxiliary Fields
      ' --------------------------------------------------------------------------------------
  
      response.write "<TR>"
  
      if request("ID") = "edit_aux" then
        response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
  
      response.write "<B>" & Translate("Edit",Login_Language,conn) & "</B> - " & Translate("Account Profile Auxiliary Fields",Login_Language,conn) & ":"
      response.write "</TD>"
  
      if request("ID") = "edit_aux" then
        response.write "<TD BGCOLOR=""#666666"" Class=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
      %>
      <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='default.asp?ID=edit_aux&Site_ID=<%=Site_ID%>&Aux_ID='+this.options[this.selectedIndex].value+'#Results'" NAME="Aux_ID">    
      <%
      SQL = "SELECT Auxiliary.* FROM Auxiliary WHERE Auxiliary.Site_ID=" & CInt(Site_ID) & " ORDER BY Auxiliary.Order_Num"
      Set rsAux = Server.CreateObject("ADODB.Recordset")
      rsAux.Open SQL, conn, 3, 3
  
      if request("ID") <> "edit_aux" then    
        response.write "<OPTION Class=Small VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>"
      end if
  
      Do while not rsAux.EOF
        tempStr = RestoreQuote(rsAux("Description"))
  
        if len(tempStr) > 40 then tempStr = Mid(tempStr,1,40) & " ..."
  
        if CInt(request("Aux_ID")) = rsAux("ID") and request("ID") = "edit_aux" then
       	  response.write "<OPTION SELECTED VALUE=""" & rsAux("ID") & """"
          if rsAux("Enabled") = True then
            response.write " CLASS=Region1NavMedium>+ "
          else
            response.write " CLASS=RegionXNavMedium>o "
          end if
          response.write tempStr
        else
       	  response.write "<OPTION VALUE=""" & rsAux("ID") & """"
          if rsAux("Enabled") = True then
            response.write " CLASS=Region1NavMedium>+ "
          else
            response.write " CLASS=RegionXNavMedium>o "
          end if
          response.write tempStr                  
        end if
        response.write "</OPTION>"                              
    	  rsAux.MoveNext 
      loop
      
      rsAux.close
      Set rsAux = Nothing            
  
      response.write "</SELECT>"
      response.write "</TD>"
      response.write "</TR>"
        
    end if
  
    ' --------------------------------------------------------------------------------------
    ' Misc Site Utilities
    ' --------------------------------------------------------------------------------------
  
    if Admin_Access = 9 or Admin_Access = 8 or Admin_Access = 4 or Admin_Access = 3 or Admin_Access = 2 then
  
      response.write "<TR>"
  
      if request("ID") = "misc_utility" then
        response.write "<TD BGCOLOR=""#666666"" CLASS=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
  
      response.write "<B>" & Translate("Misc",Login_Language,conn) & "</B> - " & Translate("Site Utilities",Login_Language,conn) & ":"
      response.write "</TD>"
  
      if request("ID") = "misc_utility" then
        response.write "<TD BGCOLOR=""#666666"" CLASS=MediumGold>"
      else
        response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Medium>"
      end if
      %>
      <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID='+this.options[this.selectedIndex].value" NAME="Utility_ID">    
      <%      
      if request("ID") <> "misc_utility" then    
        response.write "<OPTION Class=Small VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>"
      end if
  
  
  	 response.write "<OPTION CLASS=Region1NavMedium VALUE=""44"""
  	if SC2 = 44 then response.write " SELECTED"
  	response.write ">" & Translate("Associate Price List Access Code with Oracle Customer Number",Login_Language,conn) & "</OPTION>" & vbCrLf

      if Admin_Access = 2 or Admin_Access = 3 or Admin_Access = 4 or Admin_Access >= 8 then
      	if site_id = 11 then
	        response.write "<OPTION Class=Region3NavMedium VALUE=""53"">" & Translate("Metcal Procedure Administration",Login_Language,conn) & "</OPTION>"
      	end if
        if Admin_Access = 2 or Admin_Access = 4 or Admin_Access >= 8 then        
          response.write "<OPTION Class=Region1NavMedium VALUE=""50"">" & Translate("List Assets (Content or Events) - All",Login_Language,conn) & "</OPTION>"
          'response.write "<OPTION Class=Region1NavMedium VALUE=""51"">" & Translate("Literature Fulfillment - All",Login_Language,conn) & "</OPTION>"
          'response.write "<OPTION Class=Region1NavMedium VALUE=""52"">" & Translate("Literature Fulfillment - Active",Login_Language,conn) & "</OPTION>"
        end if
        
        'Modified by zensar on 25-01-2007 for adding Associate Assets option for Portweb Only.
        if Site_ID=82 then
		  if Admin_Access = 4 or Admin_Access >= 8 then
  			  response.write "<OPTION Class=Region1NavMedium VALUE=""2222"">" & Translate("List Assets Associated with Product",Login_Language,conn) & "</OPTION>"
	  	  end if
	  	  if Admin_Access = 4 or Admin_Access >= 8 then
		  	  response.write "<OPTION Class=Region1NavMedium VALUE=""1111"">" & Translate("Bulk Association of Assets to New Product",Login_Language,conn) & "</OPTION>"
		  end if
		  if Admin_Access = 4 or Admin_Access >= 8 then
		  	  response.write "<OPTION Class=Region1NavMedium VALUE=""3333"">" & Translate("List Assets",Login_Language,conn) & "</OPTION>"
		  end if
  	    end if
          
        if Admin_Access = 3 or Admin_Access = 4 or Admin_Access >= 8 then
          response.write "<OPTION Class=Region1NavMedium VALUE=""73"">" & Translate("Literature Order Activity Detail (Metrics)",Login_Language,conn) & "</OPTION>"
        end if  
      end if
        
      if Admin_Access = 4 or Admin_Access >= 8 then
        response.write "<OPTION Class=Region1NavMedium VALUE=""60"">" & Translate("Thumbnail Requests",Login_Language,conn) & "</OPTION>"
      end if
      
      if Admin_Access = 4 or Admin_Access >= 8 then
        response.write "<OPTION Class=Region1NavMedium VALUE=""98"">" & Translate("Upload File Monitor Status",Login_Language,conn) & "</OPTION>"
        response.write "<OPTION Class=Region2NavMedium VALUE=""54"">" & Translate("Subscription Service - Queue",Login_Language,conn) & "</OPTION>"                                          
        response.write "<OPTION Class=Region3NavMedium VALUE=""70"">" & Translate("Asset Activity Detail (Metrics)",Login_Language,conn) & "</OPTION>"
'EVII              
      end if

      if Admin_Access = 4 or Admin_Access = 6 or Admin_Access >= 8 then
        response.write "<OPTION Class=Region3NavMedium VALUE=""71"">" & Translate("Site Activity Summary (Metrics)",Login_Language,conn) & "</OPTION>"
        response.write "<OPTION Class=Region3NavMedium VALUE=""72"">" & Translate("WWW Document Activity Detail (Metrics)",Login_Language,conn) & "</OPTION>"
      end if
      
      response.write "</SELECT>"
      response.write "</TD>"
      response.write "</TR>"
  
    end if
  
    ' --------------------------------------------------------------------------------------
    ' Language Select
    ' --------------------------------------------------------------------------------------
  
    if Admin_Access >= 2 then
  
      response.write "<TR>"
  
      response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Small>"
      response.write "<B>" & Translate("Language",Login_Language,conn) & "</B>:<BR>"
      response.write "</TD>"
  
      response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Small>"
      SQL = "SELECT * FROM Language WHERE Language.Enable=" & CInt(True) & " ORDER BY Language.Sort"
      Set rsLanguage = Server.CreateObject("ADODB.Recordset")
      rsLanguage.Open SQL, conn, 3, 3
  
      response.write "<SELECT NAME=""Language"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='/SW-Administrator/Default.asp" & "?Site_ID=" & Site_ID & "&Language='+this.options[this.selectedIndex].value"">" & vbCrLf
  
      Do while not rsLanguage.EOF
       	  response.write "<OPTION VALUE=""" & rsLanguage("Code") & """"
          if LCase(rsLanguage("Code")) = LCase(Login_Language) then
           response.write " SELECTED"
           Language_ID = rsLanguage("ID")
          end if 
          if LCase(rsLanguage("Code")) <> "eng" then
            response.write " CLASS=Region2"
          else
            response.write " CLASS=Small"
          end if  
          response.write ">" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>"              
    	  rsLanguage.MoveNext 
      loop
      
      rsLanguage.close
      set rsLanguage=nothing
      
      Select case Admin_Access
        case 8,9
          if Session("ShowTranslation") = True then
            response.write "<OPTION VALUE=""XOF"" CLASS=Translate>Translation View Off</OPTION>"
          elseif Session("ShowTranslation") = False then
            response.write "<OPTION VALUE=""XON"" CLASS=Translate>Translation View On</OPTION>"
          end if
      end select
                              
      response.write "</SELECT>"
      response.write "</TD>"
      response.write "</TR>"
  
    end if
  
    ' --------------------------------------------------------------------------------------
    ' Quick Site Statistics
    ' --------------------------------------------------------------------------------------
  
    if Admin_Access >= 6 then
    
      response.write "<TR>"
  
      response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Small>"
      response.write "<B>" & Translate("Statistics",Login_Language,conn) & "</B>:<BR>"
      response.write "</TD>"
  
      response.write "<TD BGCOLOR=""" & Contrast & """ CLASS=Small>"
      
      ' Total
      SQL = "SELECT ID, Site_ID FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND UserData.NewFlag=" & CInt(False)
      Set rsStat = Server.CreateObject("ADODB.Recordset")
      rsStat.Open SQL, conn, 3, 3    
      response.write "&nbsp;" & Translate("Total Accounts",Login_Language,conn) & ": "
      if not rsStat.EOF then
        response.write CLng(rsStat.RecordCount)
      else
        response.write "0"  
      end if    
      rsStat.close
  
      response.write "&nbsp;&nbsp;&nbsp;"    
  
      ' Active
      SQL = "SELECT Logon, Site_ID FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND UserData.Logon IS NOT NULL"
      Set rsStat = Server.CreateObject("ADODB.Recordset")
      rsStat.Open SQL, conn, 3, 3
      response.write Translate("Active",Login_Language,conn) & ": "
      if not rsStat.EOF then
        response.write CLng(rsStat.RecordCount)
      else
        response.write "0"  
      end if    
      rsStat.close
  
      response.write "&nbsp;&nbsp;&nbsp;"    
          
      ' Pending
      SQL = "SELECT NewFlag, Site_ID FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND UserData.NewFlag=" & CInt(True)
      Set rsStat = Server.CreateObject("ADODB.Recordset")
      rsStat.Open SQL, conn, 3, 3  
      response.write Translate("Pending",Login_Language,conn) & ": "
      if not rsStat.EOF then
        response.write CLng(rsStat.RecordCount)
      else
        response.write "0"  
      end if    
      rsStat.close
      
      response.write "&nbsp;&nbsp;&nbsp;"    
          
      ' Expired
      SQL = "SELECT Site_ID FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND UserData.ExpirationDate <='" & Date() & "'"
      Set rsStat = Server.CreateObject("ADODB.Recordset")
      rsStat.Open SQL, conn, 3, 3  
      response.write Translate("Expired",Login_Language,conn) & ": "
      if not rsStat.EOF then
        response.write CLng(rsStat.RecordCount)
      else
        response.write "0"  
      end if    
      rsStat.close
  
      set rsStat = nothing
      
      response.write "</TD>"
      response.write "</TR>"
  
    end if
   
    response.write "</TABLE>"
    Call Table_End
    response.write "</FORM>"
   
    response.write "<BR>"

  end if

  ' --------------------------------------------------------------------------------------
  ' Administration Tools Container - End
  ' --------------------------------------------------------------------------------------
  
  ' --------------------------------------------------------------------------------------
  ' List / Edit Content or Event Record
  ' --------------------------------------------------------------------------------------
  
  if request("ID") = "edit_record" then
  
    if CInt(request("Category_ID")) = 9998 then                         ' View Users Queue

      SQL = "SELECT Calendar.* FROM Calendar " &_
            "WHERE Calendar.Site_ID=" & CInt(Site_ID) & " " &_
            "AND Calendar.Status=" & CInt(0) & " "

      if Admin_Access = 2 then                                          ' Submitter Only
        SQL = SQL & "AND  Calendar.Submitted_By=" & CLng(Admin_ID) & " "
      else
        SQL = SQL & "AND  Calendar.Review_By=" & CLng(Admin_ID) & " "
      end if

    elseif CInt(request("Category_ID")) = 9999 then                     ' View All Queue

      SQL = "SELECT Calendar.* FROM Calendar " &_
            "WHERE Calendar.Site_ID=" & CInt(Site_ID) & " " &_
            "AND Calendar.Status=" & CInt(0) & " "

    elseif CInt(request("SortBy")) = 5 then                             ' Parent / Child ID
    
      SQL = "SELECT *, CASE clone WHEN 0 THEN [ID] ELSE [clone] END AS PC_Order " &_
            "FROM  Calendar " &_
            "WHERE Site_ID=" & CInt(Site_ID) &_
            "AND Calendar.Category_ID=" & CInt(Request("Category_ID")) & " "
    
      if Admin_Access = 2 then                                           ' Limit User View for Submitter Only
        SQL = SQL & "AND Calendar.Submitted_By=" & CLng(Admin_ID) & " "
      end if

    else                                                                ' Individual Catagories

      SQL =       "SELECT Calendar.* "
      SQL = SQL & "FROM Calendar "
      SQL = SQL & "WHERE Calendar.Site_ID=" & CInt(Site_ID) & " "
      SQL = SQL & "AND Calendar.Category_ID=" & CInt(Request("Category_ID")) & " "

      if Admin_Access = 2 then                                           ' Limit User View for Submitter Only
        SQL = SQL & "AND Calendar.Submitted_By=" & CLng(Admin_ID) & " "
      end if

    end if
      
    if isblank(request("SortBy")) or CInt(request("SortBy")) = 3 then
      SQL = SQL & "ORDER BY Calendar.Status, Calendar.ID DESC"
    elseif CInt(request("SortBy")) = 0 then
      SQL = SQL & "ORDER BY Calendar.ID DESC"
    elseif CInt(request("SortBy")) = 1 then
      SQL = SQL & "ORDER BY Calendar.BDATE DESC"
    elseif CInt(request("SortBy")) = 2 then
      SQL = SQL & "ORDER BY Calendar.Product, Calendar.Title"
    elseif CInt(request("SortBy")) = 4 then
      SQL = SQL & "ORDER BY Item_Number DESC"
    elseif CInt(request("SortBy")) = 5 then
      SQL = SQL & "ORDER BY PC_Order, Language"
    elseif CInt(request("SortBy")) = 6 then
      SQL = SQL & "ORDER BY Calendar.Title"
    end if
                    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, conn, 3, 3  

    if not rs.EOF then

      FormName = "SearchSort"
      response.write "<A NAME=""Results""></A>"
      response.write "<FORM NAME=""" & FormName & """>" & vbCrLf
      Call Table_Begin
      
      response.write "<TABLE WIDTH=""100%"" BORDER=0 CLASS=TableBackground>"
      response.write "<TR>"
      response.write "<TD WIDTH=""5%"" BGCOLOR=""White"" CLASS=SmallBold ALIGN=CENTER>" & Translate("Status",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD WIDTH=""5%"" BGCOLOR=""Yellow"" CLASS=Small ALIGN=CENTER>" & Translate("Review",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD WIDTH=""5%"" BGCOLOR=""#00CC00"" CLASS=Small ALIGN=CENTER>" & Translate("Live",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD WIDTH=""5%"" BGCOLOR=""#AAAAFF"" CLASS=Small ALIGN=CENTER>" & Translate("Archive",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD WIDTH=""25%"" BGCOLOR=""#FFFFFF"" CLASS=SmallBold ALIGN=CENTER NOWRAP VALIGN=""top"">&nbsp;" & Translate("Search",Login_Language,conn) & ": "
      response.write "<INPUT TYPE=TEXT CLASS=Small NAME=""Search_Parameter"" Value="""" MAXLENGTH=""50"" SIZE=""20"" onkeypress=""return handleEnter(this, event);"">&nbsp;"
      response.write "<INPUT TYPE=BUTTON Class=NavLeftHighlight1 Value="" " & Translate("Go",Login_Language,conn) & " "" LANGUAGE=""JavaScript"" ONCLICK=""SearchForItem();"" TITLE=""Search for Asset by ID, Item Number or Keyword"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">&nbsp;" & vbCrLf
      response.write "</TD>" & vbCrLf
      response.write "<TD WIDTH=""55%"" CLASS=Small ALIGN=LEFT>"  & vbCrLf
      response.write "<B>&nbsp;&nbsp;" & Translate("Sort By",Login_Language,conn) & ":</B>&nbsp;&nbsp;"
      %>
      <INPUT Class=Small TYPE="RADIO" NAME="SortBy" VALUE=0 onClick="location.href='default.asp?ID=edit_record&Site_ID=<%=Site_ID%>&Category_ID=<%=Request("Category_ID")%>&SortBy=0#Results'"<%if request("SortBy") = "0" then response.write " CHECKED"%>>
      &nbsp;<%=Translate("ID",Login_Language,conn)%>&nbsp;&nbsp;
      <INPUT Class=Small TYPE="RADIO" NAME="SortBy" VALUE=5 onClick="location.href='default.asp?ID=edit_record&Site_ID=<%=Site_ID%>&Category_ID=<%=Request("Category_ID")%>&SortBy=5#Results'"<%if request("SortBy") = "5" then response.write " CHECKED"%>>
      &nbsp;<%=Translate("PID/ID",Login_Language,conn)%>&nbsp;&nbsp;
      <INPUT Class=Small TYPE="RADIO" NAME="SortBy" VALUE=1 onClick="location.href='default.asp?ID=edit_record&Site_ID=<%=Site_ID%>&Category_ID=<%=Request("Category_ID")%>&SortBy=1#Results'"<%if request("SortBy") = "1" then response.write " CHECKED"%>>
      &nbsp;<%=Translate("Date",Login_Language,conn)%>&nbsp;&nbsp;
      <INPUT Class=Small TYPE="RADIO" NAME="SortBy" VALUE=2 onClick="location.href='default.asp?ID=edit_record&Site_ID=<%=Site_ID%>&Category_ID=<%=Request("Category_ID")%>&SortBy=2#Results'"<%if request("SortBy") = "2" then response.write " CHECKED"%>>
      &nbsp;<%=Translate("Product",Login_Language,conn)%>/<%=Translate("Date",Login_Language,conn)%>&nbsp;&nbsp;&nbsp;
      <INPUT Class=Small TYPE="RADIO" NAME="SortBy" VALUE=6 onClick="location.href='default.asp?ID=edit_record&Site_ID=<%=Site_ID%>&Category_ID=<%=Request("Category_ID")%>&SortBy=6#Results'"<%if request("SortBy") = "6" then response.write " CHECKED"%>>
      &nbsp;<%=Translate("Title",Login_Language,conn)%>&nbsp;&nbsp;
      <INPUT Class=Small TYPE="RADIO" NAME="SortBy" VALUE=3 onClick="location.href='default.asp?ID=edit_record&Site_ID=<%=Site_ID%>&Category_ID=<%=Request("Category_ID")%>&SortBy=3#Results'"<%if isblank(request("SortBy")) or request("SortBy") = "3" then response.write " CHECKED"%>>
      &nbsp;<%=Translate("Status",Login_Language,conn)%>&nbsp;&nbsp;
      <INPUT Class=Small TYPE="RADIO" NAME="SortBy" VALUE=4 onClick="location.href='default.asp?ID=edit_record&Site_ID=<%=Site_ID%>&Category_ID=<%=Request("Category_ID")%>&SortBy=4#Results'"<%if request("SortBy") = "4" then response.write " CHECKED"%>>
      &nbsp;<%=Translate("Item #",Login_Language,conn)%>

      <%
      response.write "</TD>"
      response.write "</TR>"
      response.write "</TABLE>"
      Call Table_End
      
      response.write "<BR>"

      Call Table_Begin
      
            
      response.write "<DIV ID=""ContentTableStart"" STYLE=""position: absolute;"">" & vbCrLf
      response.write "</DIV>" & vbCrLf

      response.write "<TABLE WIDTH=""100%"" CELLPADDING=2 BORDER=0>"

      response.write "<TR ID=""ContentHeader1"">"
'     response.write "<TD BGCOLOR=""#FF0000"" WIDTH=""2%"" ALIGN=CENTER Class=SmallBoldWhite>" & Translate("Action",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""3%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("Thumb",Login_language,conn) & "</TD>"      
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""3%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("ID",Login_language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""3%"" ALIGN=CENTER Class=SmallBoldGold>PID</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("MAC",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""13%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("Sub-Category",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""30%"" Class=SmallBoldGold>" & Translate("Product or Product Series",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""30%"" Class=SmallBoldGold>" & Translate("Title of Content or Event",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("Item #",Login_Language,conn) & "</TD>"      
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("CC",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("POD",Login_Language,conn) & "</TD>"      
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("Language",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("Live",Login_language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("Embargo",Login_Language,conn) & "</TD>"
      response.write "</TR>"
      
      Old_PC_Order = 0
      PC_Color     = "#DDDDDD"

      Do while not rs.EOF

        ' Edit
        response.write "<TR>" & vbCrLf
        response.write "<TD BGCOLOR=""#666666"" ALIGN=CENTER VALIGN=CENTER Class=Small>"  & vbCrLf
        %>
        <INPUT TYPE="Button" onClick="location.href='Calendar_Edit.asp?ID=<%=rs("ID")%>&Site_ID=<%=Site_ID%>'" VALUE="<%=Translate("Edit",Login_Language,conn)%>" CLASS=Navlefthighlight1 onmouseover="this.className='NavLeftButtonHover'" onmouseout="this.className='Navlefthighlight1'"></A>
        <%
        response.write "</TD>"  & vbCrLf
        
        ' Thumbnail
'        response.write "<TD BGCOLOR=""#666666"" ALIGN=CENTER VALIGN=CENTER Class=Small>"  & vbCrLf
'        if not isblank(rs("Thumbnail")) then
'          response.write "<IMG SRC=""" & "/" & Site_Code & "/" & rs("Thumbnail") & """ BORDER=0 WIDTH=40>"
'        else  
'          response.write "&nbsp;"
'        end if  
'        response.write "</TD>" & vbCrLf

        ' ID
        response.write "<TD ALIGN=RIGHT BGCOLOR="
        if rs("Status") = 1 then
          response.write """#00CC00"""
        elseif rs("Status") = 2 then
          response.write """#AAAAFF"""
        else
          response.write """Yellow"""
        end if
        response.write " CLASS=Small>"
        
        response.write rs("ID")
        response.write "</TD>" & vbCrLf

        if CInt(request("SortBY")) = 5 then
          if rs("PC_Order") <> Old_PC_Order then
            if PC_Color = "#FFFFFF" then
              PC_Color = "#DDDDDD"
            else
              PC_Color = "#FFFFFF"
            end if
            Old_PC_Order = rs("PC_Order")
          end if
        end if  
            
        response.write "<TD BGCOLOR=""" & PC_Color & """ ALIGN=RIGHT Class=Small>"
        if rs("Clone") > 0 then
          response.write rs("Clone")
        else
          response.write "&nbsp;"
        end if
        response.write "</TD>" & vbCrLf
        
        response.write "<TD BGCOLOR="""
        select case CInt(rs("Content_Group"))
          case 1,2
            response.write "#FFFFFF"" Class=Region1"
          case 3,4
            response.write "#FFFFFF"" Class=Region2"
          case else
            response.write "#EEEEEE"" Class=Small"
        end select
        response.write " ALIGN=Center>"
        response.write "<SPAN CLASS=Small>"
        select case CInt(rs("Content_Group"))
          case 1  ' Product Introduciton + Individual  
            response.write "P+I"
          case 2  ' Product Introduciton Only
            response.write "P"
          case 3  ' Campaign + Individual  
            response.write "C+I"
          case 4  ' Campaign Only
            response.write "C"
          case else
            if CInt(rs("Code")) >= 8000 and CInt(rs("Code")) <= 8999 then
              response.write "<SPAN CLASS=Small>MAC</SPAN>"
            else  
              response.write "I"
            end if
        end select      
        response.write "</SPAN>"
        response.write "</TD>" & vbCrLf

        response.write "<TD BGCOLOR=""#FFFFFF"" Class=Small>" & RestoreQuote(rs("Sub_Category")) & "</TD>" & vbCrLf
        response.write "<TD BGCOLOR=""#FFFFFF"" Class=Small>" & RestoreQuote(rs("Product")) & "</TD>" & vbCrLf
        response.write "<TD BGCOLOR=""#FFFFFF"" Class=Small>" & RestoreQuote(rs("Title"))
        response.write "</TD>" & vbCrLf

        response.write "<TD BGCOLOR=""#FFFFFF"" Class=Small NOWRAP>"
        if not isblank(rs("Item_Number")) then
          response.write rs("Item_Number")
          if not isblank(rs("Revision_Code")) then
            response.write " " & rs("Revision_Code")
          end if  
        else
          response.write "&nbsp;"
        end if  
        response.write "</TD>" & vbCrLf
        
        response.write "<TD BGCOLOR=""#EEEEEE"" Class=Small ALIGN=Center>"
        if rs("Cost_Center") > 0 then
          response.write rs("Cost_Center")
        else
          response.write "&nbsp;"
        end if
        response.write "</TD>" & vbCrLf
        
        if not isblank(rs("File_Name_POD")) then
          response.write "<TD BGCOLOR=""#EEEEEE"" Class=Small ALIGN=Center><B>Y</B>"
        else
          response.write "<TD BGCOLOR=""#EEEEEE"" Class=Small ALIGN=Center>&nbsp;"
        end if  
        response.write "</TD>" & vbCrLf

        if UCase(rs("Language")) = "ENG" then
          response.write "<TD BGCOLOR=""#FFFFFF"" Class=Small ALIGN=Center>" & UCase(rs("Language")) & "</TD>" & vbCrLf
        else
          response.write "<TD BGCOLOR=""#EEEEEE"" Class=Small ALIGN=Center>" & UCase(rs("Language")) & "</TD>" & vbCrLf
        end if  

        response.write "<TD BGCOLOR=""#EEEEEE"" Class=Small ALIGN=Center>"
        if CDate(rs("BDATE")) <> CDate(rs("LDATE")) and isdate(rs("LDate")) then
          response.write FormatDate(1,rs("LDate"))
        else
          response.write FormatDate(1,rs("BDate"))
        end if  
        response.write "</TD>" & vbCrLf

        response.write "<TD BGCOLOR=""#EEEEEE"" Class=SmallRed ALIGN=CENTER>"
        if isdate(rs("PEDate")) and not isblank(rs("PEDate")) then
          if CDate(Date()) < CDate(rs("PEDate")) then
            response.write FormatDate(1,rs("PEDate"))
          else
            response.write "&nbsp;"
          end if  
        else
          response.write "&nbsp;"
        end if  
        response.write "</TD>" & vbCrLf                         
        response.write "</TR>"
        
    	  rs.MoveNext 
        
      loop

      response.write "</TABLE>"
      Call Table_End
      response.write "</FORM>"

    else

      Call Table_Begin
      response.write "<TABLE WIDTH=""100%"" CELLPADDING=4 BORDER=0>"
      response.write "<TR><TD COLSPAN=7 Class=NavLeft1>" & Translate("There are no records available to display for this category, please select another category to [Edit] or select [Add] category to add a new Content or Event item.",Login_Language,conn) & "</TD></TR>"
      response.write "</TABLE>"
      Call Table_End

    end if  

    rs.close
    set rs=nothing
    
  end if

  ' --------------------------------------------------------------------------------------  
  ' List / Edit Category Title
  ' --------------------------------------------------------------------------------------  

  if request("ID") = "edit_category" then

    SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.ID=" & CInt(request("Category_ID"))
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, conn, 3, 3  

    response.write "<FORM NAME=""Category_Edit"" ACTION=""category_admin.asp"" METHOD=""POST"">"
    response.write "<INPUT TYPE=""Hidden"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>"
    response.write "<INPUT TYPE=""Hidden"" NAME=""Category_ID"" VALUE=""" & request("Category_ID") & """>"

    response.write "<A Name=""Results""></A>"
    Call Table_Begin
    response.write "<TABLE WIDTH=""100%"" CELLPADDING=4 BORDER=0>"

    response.write "<TR><TD COLSPAN=5 BGCOLOR=""#EEEEEE"" Class=Medium>" & Translate("Note: Global change affects all records related to category. Title and Category Description must be in English.",Login_Language,conn) & "</TD></TR>"

    if not rs.EOF then

      response.write "<TR>"
      response.write "<TD BGCOLOR=""#FF0000"" WIDTH=""3%"" ALIGN=CENTER Class=SmallBoldWhite>" & Translate("Action",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""3%"" ALIGN=RIGHT Class=SmallBoldGold>" & Translate("ID",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""22%"" Class=SmallBoldGold>" & Translate("Title",Login_language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""22%"" Class=SmallBoldGold>" & Translate("Directory",Login_Language,conn) & " / " & Translate("Folder",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""50%"" Class=SmallBoldGold>" & Translate("Customize Category Attributes",Login_Language,conn) & "</TD>"
      response.write "</TR>"

      Do while not rs.EOF
        response.write "<TR>" & vbCrLf
        response.write "<TD BGCOLOR=""#666666"" ALIGN=CENTER VALIGN=TOP Class=Medium>"  & vbCrLf
        %>
        <INPUT TYPE="Submit" NAME="Submit" VALUE="Save" CLASS=Navlefthighlight1>
        <%
        response.write "</TD>"  & vbCrLf
        response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=RIGHT VALIGN=TOP Class=Medium><FONT COLOR=""#666666"">" & rs("ID") & "</FONT>"
        response.write "<INPUT TYPE=""Hidden"" NAME=""ID"" VALUE=""" & rs("ID") & """></TD>" & vbCrLf

        response.write "<TD BGCOLOR=""#FFFFFF"" VALIGN=TOP Class=Medium>"
        response.write "<INPUT Class=Medium TYPE=""TEXT"" NAME=""Title"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & RestoreQuote(rs("Title")) & """>"
        response.write "<BR><BR>" & Translate("Category Description",Login_Language,conn) & "<BR>" & vbCrLf
        response.write "<TEXTAREA onKeyPress='return maxLength(this,""3000"");' onpaste='return maxLengthPaste(this,""3000"");' Class=Medium NAME=""Description""  MAXLENGTH=""3000"" COLS=33 ROWS=15  TITLE=""Category Description"">" & RestoreQuote(rs("Description")) & "</TEXTAREA>"

	'Validation to check max length on description text area starts
        'added on 19aug09
        'related to item#611
        response.write "<script language=""javascript"">" & vbCrLf
		response.write "function maxLength(field,maxChars)" & vbCrLf
		response.write "{if(field.value.length >= maxChars) {" & vbCrLf
		response.write "event.returnValue=false;" & vbCrLf
		response.write "alert(""Contents exceed maximum limit of "" +maxChars+"" characters"");" & vbCrLf
		response.write "return false;}}" & vbCrLf
		
		response.write "function maxLengthPaste(field,maxChars)" & vbCrLf
		response.write "{event.returnValue=false;" & vbCrLf
		response.write "if((field.value.length +  window.clipboardData.getData(""Text"").length) > maxChars) {" & vbCrLf
		response.write "alert(""Contents exceed maximum limit of "" +maxChars+"" characters"");" & vbCrLf
		response.write "return false;" & vbCrLf
		response.write "}event.returnValue=true;}" & vbCrLf
		response.write "</script>"
		'Validation ends.
	
        ' Rich Text Editor for TextArea
        if Admin_Access >=8 then
          response.write "<P ALIGN=CENTER>"
          RTE_Length = 4000
          RTE_Cols   = 33
          RTE_Rows   = 15
          FormName   = "Category_Edit"
          Element    = "Description"
          response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" VALUE=""HTML"" LANGUAGE=""JavaScript"" ONCLICK=""RTEditor_Open('" & FormName & "','" & Element & "','" & Site_ID & "','" & Site_Code & "','" & RTE_Length & "','" & RTE_Cols & "','" & RTE_Rows & "');"" TITLE=""Edit Field with the HTML Editor"">" & vbCrLf
          response.write "</P>"
        end if
          
        response.write "</TD>"

        response.write "<TD BGCOLOR=""#FFFFFF"" VALIGN=TOP Class=Medium>" & RestoreQuote(rs("Category")) & "</TD>" & vbCrLf

        response.write "<TD BGCOLOR=""#EEEEEE"">" & vbCrLf
        response.write "<TABLE WIDTH=""100%"" BGCOLOR=""EEEEEE"" CELLPADDING=0 CELLSPACING=0 BORDER=0>" & vbCrLf

        response.write "<TR>"
        response.write "<TD WIDTH=""60%"" Class=MediumBold>" & Translate("Category Enabled",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD WIDTH=""40%"" ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Enabled"""
        if rs("Enabled") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Navigation Sort Order (number)",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""TEXT"" NAME=""Sort"" SIZE=""9"" MAXLENGTH=""3"" VALUE=""" & rs("Sort") & """></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf
        
        response.write "<TR><TD Class=Medium>" & Translate("Category Sort By",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium>" & vbCrLf
        response.write "<SELECT NAME=""SortBy"" ClASS=Medium>" & vbCrLf
        response.write "<OPTION VALUE=""0"""
        if rs("SortBy") = 0 then response.write " SELECTED"
        response.write ">"  & Translate("Product",Login_Language,conn) & "</OPTION>" & vbCrLf
        response.write "<OPTION VALUE=""1"""
        if rs("SortBy") = 1  or rs("SortBy") = 999 then response.write " SELECTED"
        response.write ">"  & Translate("Category",Login_Language,conn) & "</OPTION>" & vbCrLf
        response.write "<OPTION VALUE=""2"""
        if rs("SortBy") = 2 then response.write " SELECTED"
        response.write ">"  & Translate("Date",Login_Language,conn) & "</OPTION>" & vbCrLf

        response.write "</TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf
        
        response.write "<TR><TD Class=Medium>" & Translate("Sub-Category",Login_language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Sub_Category"""
        if rs("Sub_Category") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf        
        
        SQL       = "SELECT Content_Sub_Category.Site_ID, Content_Sub_Category.Sub_Category, Content_Sub_Category.Code, Content_Sub_Category.Language "
        SQL = SQL & "FROM Content_Sub_Category "
        SQL = SQL & "GROUP BY Content_Sub_Category.Site_ID, Content_Sub_Category.Sub_Category, Content_Sub_Category.Code, Content_Sub_Category.Language "
        SQL = SQL & "HAVING Content_Sub_Category.Site_ID=" & Site_ID & " "
        SQL = SQL & "AND Content_Sub_Category.Sub_Category IS NOT NULL "
        SQL = SQL & "AND Content_Sub_Category.Code=" & rs("Code") & " "
        SQL = SQL & "AND Content_Sub_Category.Language='eng'"

        Set rsSubCategoryPreset = Server.CreateObject("ADODB.Recordset")
        rsSubCategoryPreset.Open SQL, conn, 3, 3              

        if not rsSubCategoryPreset.EOF then
          response.write "<TR><TD Class=Medium>" & Translate("Sub-Category Preset (for new asset)",Login_Language,conn) & ":</TD>" & vbCrLf
          response.write "<TD ALIGN=CENTER Class=Medium>"
          
          response.write "<SELECT NAME=""Preset_Sub_Category"" CLASS=Medium>"
          response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>"

          Do while not rsSubCategoryPreset.EOF            
         	  response.write "<OPTION CLASS=Medium VALUE=""" & rsSubCategoryPreset("Sub_Category") & """"
            if LCase(rs("Preset_Sub_Category")) = LCase(rsSubCategoryPreset("Sub_Category")) then
              response.write " SELECTED"
            end if  
            response.write ">" & rsSubCategoryPreset("Sub_Category") & "</OPTION>" & vbCrLF                 
        	  rsSubCategoryPreset.MoveNext 
          loop

          response.write "</SELECT>"
          response.write "</TD></TR>"  & vbCrLf

        end if
        
        rsSubCategoryPreset.close
        set rsSubCategoryPreset = nothing
        
        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Product or Product Series",Login_language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Product_Series"""
        if rs("Product_Series") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Content or Event Title",Login_Language,conn) & ":</TD>"
        response.write "<TD ALIGN=CENTER Class=Small>" & Translate("Required",Login_Language,conn) & "</TD></TR>" & vbCrLf
        response.write "<TR><TD Class=Medium>" & Translate("Description",Login_Language,conn) & ":</TD>"
        response.write "<TD ALIGN=CENTER Class=Small>" & Translate("Optional",Login_Language,conn) & "</TD></TR>" & vbCrLf
        
        response.write "<TR><TD Class=Medium>" & Translate("Allow MAC Containers",Login_language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Content_Group"""
        if rs("Content_Group") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Special Instructions",Login_language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Instructions"""
        if rs("Instructions") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR>"
        response.write "<TD Class=Medium>" & Translate("Append Library Subcategory to Title",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Title_View"""
        if rs("Title_View") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf
        
        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Item Reference Number 1",Login_language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Item_Number"""
        if rs("Item_Number") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Item Reference Number 2",Login_language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Item_Number_2"""
        if rs("Item_Number_2") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("PCat Relationships",Login_language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""PID"""
        if rs("PID") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf
        
        response.write "<TR><TD Class=Medium>" & Translate("Location",Login_language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Location"""
        if rs("Location") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf
        
        response.write "<TR><TD Class=Medium>" & Translate("Web Page URL",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Link"""
        if rs("Link") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Web Page URL Pop-Up Window Disable",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Link_PopUp_Disabled"""
        if rs("Link_PopUp_Disabled") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Web Asset File Upload",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""File_Name"""
        if rs("File_Name") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("POD Asset File Upload",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""File_Name_POD"""
        if rs("File_Name_POD") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        'response.write "<TR><TD Class=Medium>" & Translate("Include File Upload",Login_Language,conn) & ":</TD>" & vbCrLf
        'response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Include"""
        'if rs("Include") = True then response.write " CHECKED"
        'response.write "></TD>" & vbCrLf
        'response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Image File Upload",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Thumbnail"""
        if rs("Thumbnail") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Preserve Web Asset File Path for Clone",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Preserve_Path_Clone"""
        if rs("Preserve_Path_Clone") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf
        
        response.write "<TR><TD Class=Medium>" & Translate("Image Store Locator ID",Login_language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""ImageStore"""
        if rs("ImageStore") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf
        
        response.write "<TR><TD Class=Medium>" & Translate("Begin Date Only",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Date_Basic"""
        if rs("Date_Basic") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Public Release Date",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Date_PRD"""
        if rs("Date_PRD") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Mark as Confidential",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Mark_Confidential"""
        if rs("Mark_Confidential") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Subscription Notification",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Subscription"""
        if rs("Subscription") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Preset Checkbox for EEF",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Preset_EEF"""
        if rs("Preset_EEF") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Preset Checkbox for FDL",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Preset_FDL"""
        if rs("Preset_FDL") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf
        
        response.write "<TR><TD Class=Medium>" & Translate("Shopping Cart",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Shopping_Cart"""
        if rs("Shopping_Cart") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Country Restrictions",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Country_Restrictions"""
        if rs("Country_Restrictions") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf
        
        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Portal Site View",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Site_View"""
        if rs("Site_View") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Reassign Asset Owner",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Reassign_Owner"""
        if rs("Reassign_Owner") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Submission Approval",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Submission_Approve"""
        if rs("Submission_Approve") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf
        
        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Calendar View",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Calendar_View"""
        if rs("Calendar_View") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>" & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Forum View",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Forum"""
        if rs("Forum") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>" & vbCrLf
        
        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Menu Separator",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Separator"""
        if rs("Separator") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf


        ''added on 28th Oct 2009 to display Marketing Automation category
        response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf

        response.write "<TR><TD Class=Medium>" & Translate("Marketing Automation",Login_Language,conn) & ":</TD>" & vbCrLf
        response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Marketing_Auto"""
        if rs("Marketing_Automation") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"  & vbCrLf
        ''end
        
        ''Added on 6th Aug 2010 to Enable CDN 
        ''Currentlly it is only for Fnet
        if Site_ID=82 then
            response.write "<TR><TD COLSPAN=2 BGCOLOR=""#666666"" Class=Medium></TD></TR>" & vbCrLf
            response.write "<TR><TD Class=Medium>" & Translate("Enable CDN",Login_Language,conn) & ":</TD>" & vbCrLf
            response.write "<TD ALIGN=CENTER Class=Medium><INPUT Class=Medium TYPE=""Checkbox"" id=""Enable_CDN"" NAME=""Enable_CDN"""
            if rs("CDN_Implementation") = True then response.write " CHECKED"
            response.write " onclick=""enableCDNState(this)""></TD>" & vbCrLf
            response.write "</TR>"  & vbCrLf
            
            ''For Note
            response.write "<TR><TD Class=Medium colspan=2>" & Translate("(CDN will be enabled only in case of File Upload)",Login_Language,conn)& "</TD>" & vbCrLf
            response.write "</TR>"  & vbCrLf
            
            ''Added on 16th Aug 2010, for CDN check box state
            response.write "<TR><TD Class=Medium colspan=2><br>" & Translate("CDN Checkbox State:",Login_Language,conn)& "</TD>" & vbCrLf
            response.write "</TR>"  & vbCrLf
            response.write "<TR><TD Class=Medium colspan=2>" & vbCrLf
            response.write "<INPUT Class=Medium TYPE=""RADIO"" NAME=""CDN_State"" id=""CheckedDisabled"" VALUE= ""CheckedDisabled"""
            if rs("CDN_DefaultState") = "CheckedDisabled" then response.write " CHECKED" 
            response.write ">" & vbCrLf
            response.write "<label for=""CheckedDisabled"">Checked/Disabled</label>" & vbCrLf
            response.write "<INPUT Class=Medium TYPE=""RADIO"" NAME=""CDN_State"" id= ""CheckedEnabled"" VALUE= ""CheckedEnabled""" 
            if rs("CDN_DefaultState") = "CheckedEnabled" then response.write " CHECKED"
            response.write ">" & vbCrLf
            response.write "<label for=""CheckedEnabled"">Checked/Enabled</label>" & vbCrLf
            response.write "<INPUT Class=Medium TYPE=""RADIO"" NAME=""CDN_State"" id= ""UncheckedEnabled"" VALUE= ""UncheckedEnabled"""
            if rs("CDN_DefaultState") = "UncheckedEnabled" then response.write " CHECKED"
            response.write ">" & vbCrLf
            response.write "<label for=""UncheckedEnabled"">Unchecked/Enabled</label>" & vbCrLf
            response.write "</TD></TR>"
            ''End
                       
            
        end if
        ''End
        
'        response.write "</TD>" & vbCrLf
        response.write "</TABLE>" & vbCrLf

    	  rs.MoveNext 
      loop
    else
      response.write "<TR><TD COLSPAN=5 Class=Medium>" & Translate("No data is available for this category, please see the Domain Administrator for assistance.",Login_Language,conn) & "</TD></TR>" & vbCrLf
    end if  

    response.write "</TABLE>"
    Call Table_End
    response.write "</FORM>"
    
    response.write "<INDENT>"
    response.write "<UL>"
    response.write "<LI>" & Translate("Click on [Save] to save changes or select another site administration function.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Title - User viewable name of category.  This name must be in English.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Customize Category Attributes - Determines whether or not this option is available for data input and/or display to the user.",Login_Language,conn) & "</LI>"
    response.write "</UL>"

    rs.close
    set rs=nothing
  end if

  ' --------------------------------------------------------------------------------------  
  ' Edit Subscription Titles
  ' --------------------------------------------------------------------------------------  

  if request("ID") = "edit_subscription" then
  
    SQL = "SELECT "
    select case request.querystring("Subscription_Option_ID")
    
      case 1
        SQL = SQL & "Subscription_Subject"
      case 2
        SQL = SQL & "Subscription_Header"
      case 3
        SQL = SQL & "Subscription_Footer"
      case else
    end select
    
    SQL = SQL & " AS Subscription_Text FROM Site WHERE ID=" & Site_ID
                           
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, conn, 3, 3  

    response.write "<FORM NAME=""Subscription_Edit"" ACTION=""subscription_admin.asp"" METHOD=""POST"">"
    response.write "<INPUT TYPE=""Hidden"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>"
    response.write "<INPUT TYPE=""Hidden"" NAME=""Subscription_Option_ID"" VALUE=""" & request("Subscription_Option_ID") & """>"
    
    response.write "<A Name=""Results""></A>"
    Call Table_Begin
    response.write "<TABLE WIDTH=""100%"" CELLPADDING=4 BORDER=0>"

    response.write "<TR><TD COLSPAN=3 BGCOLOR=""#EEEEEE"" Class=Medium>" & Translate("Note: This global change affects all future subscription emails. Text must be in English.",Login_Language,conn) & "</TD></TR>"

    response.write "<TR>"
    response.write "<TD BGCOLOR=""#FF0000"" WIDTH=""1%"" ALIGN=CENTER Class=SmallBoldWhite>" & Translate("Action",Login_Language,conn) & "</TD>"
    response.write "<TD BGCOLOR=""#666666"" WIDTH=""1%"" Class=SmallBoldGold>"
    response.write Translate("Subscription Service Email ",Login_Language,conn) & ": "
    select case request.querystring("Subscription_Option_ID")
      case 1
        response.write Translate("Subject Line",Login_Language,conn)
      case 2
        response.write Translate("Alternate Header",Login_Language,conn)
      case 3
        response.write Translate("Alternate Footer",Login_Language,conn)      
    end select
    response.write Translate(" Text",Login_language,conn) & "</TD>"
    response.write "<TD BGCOLOR=""#666666"" WIDTH=""98%"" Class=SmallBoldGold>"
    response.write Translate("Comments",Login_Language,conn)
    response.write "</TD>"    
    response.write "</TR>"
  
    response.write "<TR>" & vbCrLf
    response.write "<TD ALIGN=CENTER VALIGN=TOP Class=Medium>"  & vbCrLf
    response.write "<INPUT TYPE=""Submit"" NAME=""Submit"" VALUE=""Save"" CLASS=Navlefthighlight1>"
    response.write "</TD>"  & vbCrLf

    response.write "<TD BGCOLOR=""White"" VALIGN=TOP CLASS=Medium>" & vbCrLf
    
    select case request.querystring("Subscription_Option_ID")
      case 1
        response.write "<INPUT TYPE=TEXT CLASS=Medium NAME=""Subscription_Text"" MAXLENGTH=""80"" SIZE=""60"" VALUE="""  & RestoreQuote(rs("Subscription_Text")) & """>"
      case else  
        response.write "<TEXTAREA Class=Medium NAME=""Subscription_Text""  MAXLENGTH=4000 COLS=60 ROWS=15 TITLE=""Description"">" & RestoreQuote(rs("Subscription_Text")) & "</TEXTAREA>"
    end select
    ' Rich Text Editor for TextArea
    if 1=1 and Admin_Access >=8 and request.querystring("Subscription_Option_ID") <> "1" then
      response.write "<P ALIGN=LEFT>"
      RTE_Length = 4000
      RTE_Cols   = 33
      RTE_Rows   = 15
      FormName   = "Subscription_Edit"
      Element    = "Subscription_Text"
      response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" VALUE=""HTML"" LANGUAGE=""JavaScript"" ONCLICK=""RTEditor_Open('" & FormName & "','" & Element & "','" & Site_ID & "','" & Site_Code & "','" & RTE_Length & "','" & RTE_Cols & "','" & RTE_Rows & "');"" TITLE=""Edit Field with the HTML Editor"">" & vbCrLf
      response.write "</P>"
    end if
      
    response.write "</TD>"& vbCrLf
    
    response.write "<TD BGCOLOR=""White"" VALIGN=TOP CLASS=Medium>" & vbCrLf
    select case request.querystring("Subscription_Option_ID")
      case 1
        response.write "<B>" & Translate("Caution",Login_Language,conn) & "</B>: " & Translate("Do not include any HTML &lt;TAG&gt; in the Subject Line, except the &lt;SITENAME&gt; and or &lt;DATE&gt; tags defined below in the help section of this page.",Login_Language,conn) & " "     
      case else
        response.write "<B>" & Translate("Caution",Login_Language,conn) & "</B>: " & Translate("The Subscription Service Email Service, sends both HTML (Rich Text) and Plain Text versions in the same email.",Login_Language,conn) & "<P>"     
        response.write Translate("Therefore, it is important to limit the use of embedded HTML &lt;TAG&gt; attributes within the text because it may not render as expected in the Plain Text version.",Login_Language,conn) & "<P>"
        response.write Translate("You can view both the HTML and Plain Text rendering of your custom text below.",Login_Language,conn) & " "
        response.write Translate("Remember to [SAVE] your edits to see the updated rendering.",Login_Language,conn)
    end select   
    response.write "</TD>" & vbCrLf
  
    RenderText = RestoreQuote(rs("Subscription_Text"))
    if not isblank(RenderText) then
      RenderText = Replace(RenderText,"<SITENAME>",Site_Description)
      RenderText = Replace(RenderText,"<DATE>",FormatDateText(2, Date()))
    end if  
    response.write "<TR><TD>&nbsp;</TD><TD COLSPAN=3 BGCOLOR=""#EEEEEE"" Class=Medium>" & Translate("Rendered as HTML",Login_Language,conn) & "</TD></TR>"
    response.write "<TR><TD>&nbsp;</TD><TD COLSPAN=3 BGCOLOR=""White"" Class=Medium>" & RenderText & "</TD></TR>"    
    response.write "<TR><TD>&nbsp;</TD><TD COLSPAN=3 BGCOLOR=""#EEEEEE"" Class=Medium>" & Translate("Rendered as Plain Text",Login_Language,conn) & "</TD></TR>"
    if not isblank(RenderText) then
      RenderText = Replace(Un_HTML(RenderText),vbCrLf,"<BR>")
    end if  
    response.write "<TR><TD>&nbsp;</TD><TD COLSPAN=3 BGCOLOR=""White"" Class=Medium>" & RenderText & "</TD></TR>"    
    
    response.write "</TABLE>"
    Call Table_End
    response.write "</FORM>"
    
    response.write "<INDENT>"
    response.write "<UL>"

    select case request.querystring("Subscription_Option_ID")
      case 1
        response.write "<LI>" & Translate("Subject Line Text appears as the Subject Line of the Email.",Login_Language,conn) & " " & Translate("You can embedded the tags &lt;SITENAME&gt; to include the Site's name and/or &lt;DATE&gt; (that will appear as: dd Month yyyy) to include the Date of the Subscription Email.",Login_Language,conn) & "</LI>"      
        response.write "<LI>" & Translate("Example: ""Today`s News from &lt;SITENAME&gt; - &lt;DATE&gt;"" would produce the following Subject Line:",Login_Language,conn)
        response.write Translate(Site_Description,Login_Language,conn) & " - " & Translate("Today`s News",Login_Language,conn) & " - " & FormatDateText(2,Date) & "</LI>"              
        response.write "<LI>" & Translate("To revert back to the default Subject Line, delete the existing text then click [SAVE].",Login_Language,conn) & "</LI>" & vbCrLf
        response.write "<LI>" & Translate("No language translations are provided for custom text for either Subject Line, Alternate Header or Alternate Footer.  Custom text will appear in the Subscription Email as English only.",Login_Language,conn)
      case 2
        response.write "<LI>" & Translate("Alternate Header Text appears following the default ""What`s New"" Header.",Login_Language,conn) & "</LI>"
        response.write "<LI>" & Translate("To revert back to the default Header without the Alternate Header text, delete the existing text then click [SAVE].",Login_Language,conn) & "</LI>" & vbCrLf        
        response.write "<LI>" & Translate("No language translations are provided for custom text for either Subject Line, Alternate Header or Alternate Footer.  Custom text will appear in the Subscription Email as English only.",Login_Language,conn)
        response.write Translate("To force the subscription service to render in English Language Only for all emails sent, embed the HTML tag, &lt;ENGLISHONLY&gt;.",Login_Language,conn) & "</LI>"
        if Admin_Access >= 9 then
          response.write "<LI>Domain Administrator Tags: &lt;NOHEADER&gt; &lt;NOFOOTER&gt; &lt;NOUNSUBSCRIBE&gt; &lt;NOSIGNATURE&gt; &lt;NOCOPYRIGHT&gt;" & "</LI>"        
        end if      
      case 3
        response.write "<LI>" & Translate("Alternate Footer Text appears prior to the default Footer.",Login_Language,conn) & "</LI>"
        response.write "<LI>" & Translate("To revert back to the default Footer without the Alternate Footer text, delete the existing text then click [SAVE].",Login_Language,conn) & "</LI>" & vbCrLf        
        response.write "<LI>" & Translate("No language translations are provided for custom text for either Subject Line, Alternate Header or Alternate Footer.  Custom text will appear in the Subscription Email as English only.",Login_Language,conn)
        response.write Translate("To force the subscription service to render in English Language Only for all emails sent, embed the HTML tag, &lt;ENGLISHONLY&gt;.",Login_Language,conn) & "</LI>"    
        if Admin_Access >= 9 then
          response.write "<LI>Domain Administrator Tags: &lt;NOHEADER&gt; &lt;NOFOOTER&gt; &lt;NOUNSUBSCRIBE&gt; &lt;NOSIGNATURE&gt; &lt;NOCOPYRIGHT&gt;" & "</LI>"        
        end if      

    end select
    response.write "<LI>" & Translate("Click on [Save] to save changes or select another site administration function.",Login_Language,conn) & "</LI>"
    response.write "</UL>"

    rs.close
    set rs=nothing
  end if

  ' --------------------------------------------------------------------------------------  
  ' List / Edit Group Title
  ' --------------------------------------------------------------------------------------  

  if request("ID") = "edit_group" then
  
    SQL = "SELECT SubGroups.* FROM SubGroups WHERE SubGroups.ID=" & CInt(request("Group_ID"))
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, conn, 3, 3  
  

    response.write "<FORM NAME=""Group_Admin"" ACTION=""group_admin.asp"" METHOD=""POST"">"
    response.write "<INPUT TYPE=""Hidden"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>"
    response.write "<INPUT TYPE=""Hidden"" NAME=""Group_ID"" VALUE=""" & request("Group_ID") & """>"
  
    response.write "<A Name=""Results""></A>"
    Call Table_Begin
    response.write "<TABLE WIDTH=""100%"" CELLPADDING=2 BORDER=0>"
    response.write "<TR><TD COLSPAN=7 BGCOLOR=""#EEEEEE"" Class=Medium>" & Translate("Note: Global change affects all records related to groups and categories.  The Displayed Title must be in English.",Login_Language,conn) & "</TD></TR>"

    if not rs.EOF then

      response.write "<TR>"
      response.write "<TD BGCOLOR=""#FF0000"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldWhite>" & Translate("Action",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("ID",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""40%"" Class=SmallBoldGold>" & Translate("Displayed Title",Login_language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""30%"" Class=SmallBoldGold>" & Translate("Default Defintion",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""10%"" Class=SmallBoldGold>" & Translate("Group Code",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" Class=SmallBoldGold>" & Translate("Enabled",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" Class=SmallBoldGold>" & Translate("Default",Login_Language,conn) & "</TD>"
      response.write "</TR>"

      Do while not rs.EOF
        response.write "<TR>" & vbCrLf
        response.write "<TD BGCOLOR=""#666666"" ALIGN=CENTER VALIGN=TOP Class=Medium>"  & vbCrLf
        %>
        <INPUT TYPE="Submit" NAME="Submit" VALUE="Save" CLASS=Navlefthighlight1>
        <%
        response.write "</TD>"  & vbCrLf
 
        response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=RIGHT Class=Medium><FONT COLOR=""#666666"">" & rs("ID") & "</FONT>"
        response.write "<INPUT TYPE=""Hidden"" NAME=""ID"" VALUE=""" & rs("ID") & """></TD>" & vbCrLf
        response.write "<TD BGCOLOR=""#FFFFFF"" Class=Medium>"
        response.write "<INPUT Class=Medium TYPE=""TEXT"" NAME=""X_Description"" WIDTH=""50"" MAXLENGTH=""50"" VALUE=""" & RestoreQuote(rs("X_Description")) & """></TD>" & vbCrLf
        response.write "<TD BGCOLOR=""#FFFFFF"" Class=Medium>" & RestoreQuote(rs("Description")) & "</TD>" & vbCrLf
        response.write "<TD BGCOLOR=""#FFFFFF"" Class=Medium>" & rs("Code") & "</TD>" & vbCrLf
  
        response.write "<TD ALIGN=CENTER Class=Medium BGCOLOR=""#EEEEEE""><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Enabled"""
        if rs("Enabled") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
  
        response.write "<TD ALIGN=CENTER Class=Medium BGCOLOR=""#EEEEEE""><INPUT Class=Medium TYPE=""Checkbox"" NAME=""Default_Select"""
        if rs("Default_Select") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        response.write "</TR>"
                  
        rs.MoveNext
                  
      loop
                  
    end if

    response.write "</TABLE>"
    Call Table_End
    response.write "</FORM>"

    response.write "<INDENT>"
    response.write "<UL>"
    response.write "<LI>" & Translate("Click on [Save] to save changes or select another site administration function.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Displayed Title - Customizable user viewable name of the default internal group name.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Enabled - Determines whether or not this group is available for use.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Default - Determines whether or not this group is automatically selected for a new Content or Event record.",Login_Language,conn) & "</LI>"
    response.write "</UL>"
    rs.close
    set rs=nothing
  end if

  ' --------------------------------------------------------------------------------------  
  ' Auxiliary Field Options
  ' --------------------------------------------------------------------------------------  

  if request("ID") = "edit_aux" then
  
    SQL = "SELECT Auxiliary.* FROM Auxiliary WHERE Auxiliary.ID=" & CInt(request("Aux_ID"))
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, conn, 3, 3  
  
    response.write "<FORM ACTION=""auxiliary_admin.asp"" METHOD=""POST"">"
    response.write "<INPUT TYPE=""Hidden"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>"
    response.write "<INPUT TYPE=""Hidden"" NAME=""Aux_ID"" VALUE=""" & request("Aux_ID") & """>"
  
    response.write "<A Name=""Results""></A>"
    Call Table_Begin
    response.write "<TABLE WIDTH=""100%"" CELLPADDING=2 BORDER=0>"
    response.write "<TR><TD COLSPAN=9 BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Note: Global change affects all records related to Auxiliary Information Field.  &quot;Displayed Text&quot; and &quot;Text Choices&quot; must be in English.",Login_Language,conn) & "</TD></TR>"

    if not rs.EOF then

      response.write "<TR>"
      response.write "<TD BGCOLOR=""#FF0000"" WIDTH=""5%"" ALIGN=CENTER CLASS=SmallBoldWhite>" & Translate("Action",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("ID",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""40%"" Class=SmallBoldGold>" & Translate("Displayed Text",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""25%"" Class=SmallBoldGold>" & Translate("Defintion",Login_language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""10%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("Method",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("Enabled",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("Required",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("User Edit",Login_Language,conn) & "</TD>"
      response.write "<TD BGCOLOR=""#666666"" WIDTH=""5%"" ALIGN=CENTER Class=SmallBoldGold>" & Translate("Reg Form",Login_Language,conn) & "</TD>"
      response.write "</TR>"

      Do while not rs.EOF
        response.write "<TR>" & vbCrLf
        response.write "<TD BGCOLOR=""#666666"" ALIGN=CENTER VALIGN=TOP ROWSPAN=2 CLASS=Medium>"  & vbCrLf
        response.write "<INPUT TYPE=""Submit"" NAME=""Submit"" VALUE=""Save"" CLASS=Navlefthighlight1>"
        response.write "</TD>"  & vbCrLf
        
        response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP ALIGN=RIGHT ROWSPAN=2 CLASS=Medium><FONT COLOR=""#666666"">" & rs("ID") & "</FONT>"
        response.write "<INPUT TYPE=""Hidden"" NAME=""ID"" VALUE=""" & rs("ID") & """ CLASS=Medium></TD>" & vbCrLf
        
        response.write "<TD BGCOLOR=""#FFFFFF"" VALIGN=TOP ROWSPAN=2 CLASS=Medium>"
        response.write "<TEXTAREA CLASS=Medium NAME=""Description"" COLS=35 ROWS=10 TITLE=""Category Description"" MAXLENGTH=""255"" CLASS=Medium>" & RestoreQuote(rs("Description")) & "</TEXTAREA></TD>"
        
        response.write "<TD BGCOLOR=""#FFFFFF"" VALIGN=TOP ROWSPAN=2 CLASS=Medium>"
        response.write "Auxiliary " & Trim(rs("Order_Num"))
        response.write "</TD>" & vbCrLf                  
        
        response.write "<TD ALIGN=CENTER VALIGN=TOP CLASS=Medium BGCOLOR=""#EEEEEE"">"
        response.write "<SELECT NAME=""Input_Method"" CLASS=Medium>"
        response.write "<OPTION CLASS=Medium VALUE=""0"""
        if rs("Input_Method") = 0 then response.write " SELECTED"
        response.write ">" & Translate("Text",Login_Language,conn) & "</OPTION>"
        response.write "<OPTION CLASS=Medium VALUE=""1"""
        if rs("Input_Method") = 1 then response.write " SELECTED"                  
        response.write ">" & Translate("Drop-Down",Login_Language,conn) & "</OPTION>"
        response.write "<OPTION CLASS=Medium VALUE=""2"""
        if rs("Input_Method") = 2 then response.write " SELECTED"                  
        response.write ">" & Translate("Checkbox",Login_Language,conn) & "</OPTION>"
        
        response.write "</SELECT>"
        response.write "</TD>"

        response.write "<TD ALIGN=CENTER VALIGN=TOP CLASS=Medium BGCOLOR=""#EEEEEE"">"
        response.write "<INPUT CLASS=Medium TYPE=""Checkbox"" NAME=""Enabled"""
        if rs("Enabled") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf

        response.write "<TD ALIGN=CENTER VALIGN=TOP CLASS=Medium BGCOLOR=""#EEEEEE"">"
        response.write "<INPUT CLASS=Medium TYPE=""Checkbox"" NAME=""Required"""
        if rs("Required") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf

        response.write "<TD ALIGN=CENTER VALIGN=TOP CLASS=Medium BGCOLOR=""#EEEEEE"">"
        response.write "<INPUT CLASS=Medium TYPE=""Checkbox"" NAME=""User_Edit"""
        if rs("User_Edit") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf

        response.write "<TD ALIGN=CENTER VALIGN=TOP CLASS=Medium BGCOLOR=""#EEEEEE"">"
        response.write "<INPUT CLASS=Medium TYPE=""Checkbox"" NAME=""Registration"""
        if rs("Registration") = True then response.write " CHECKED"
        response.write "></TD>" & vbCrLf
        
        Response.write "</TR>"
        response.write "<TR>"
        
        response.write "<TD BGCOLOR=""#FFFFFF"" VALIGN=TOP CLASS=Medium COLSPAN=5>"
        response.write "<HR NOSHADE SIZE=1>"
        response.write Translate("Method for Checkbox or Drop-Down<BR><BR>Supply the &quot;Text Choices&quot; separated by a comma in the input box below",Login_Language,conn) & ":<BR><BR>"
        response.write "<INPUT CLASS=Medium TYPE=""Text"" Name=""Radio_Text"" VALUE=""" & rs("Radio_Text") & """ WIDTH=""50"" MAXLENGTH=""255""><BR><BR>" & Translate("Example",Login_Language,conn) & ": Yes, No"
        response.write "</TD>" & vbCrLf

        response.write "</TR>"
        
        rs.MoveNext
        
      loop
              
    end if

    response.write "</TABLE>"
    Call Table_End
    response.write "</FORM>"
    
    response.write "<INDENT>"
    response.write "<UL>"
    response.write "<LI>" & Translate("Click on [Save] to save changes or select another Auxiliary Information field.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Displayed Text - User viewable text description of the Auxiliary Information field. This could be in the form of a Field Name, or a complete question requiring an answer from the user.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Method = Text -- User supples a text answer, Method = Button -- User selects from one or more choices as defined by the Text Choice field, Method = Drop-Down -- Same as Button, except user choices are displayed in a drop-down menu.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Drop-Down / Button Choice Text - Text choices for restricted input selection.  Each choice is separated by a comma. (Ensure that the only commas used in this definition are to separate the selection text otherwise the text containing a comma will be treated as another choice.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Enabled - Determines whether or not this Auxiliary Information field is available for use.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Required - Requires an answer from the user.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("User Edit - Determines whether or not the Auxiliary Information field is available to the user for edit or just for view only.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Registration Form - The Auxiliary Information field will appear on the initial site registration form and the Users Account Profile.",Login_Language,conn) & "</LI>"
    response.write "</UL>"

    rs.close
    set rs=nothing

  end if

  if Admin_Access >= 4 then
    response.write "<UL>"
  end if  
  select case Admin_Access
    case 6, 8, 9
      response.write "<LI><A HREF=""/SW-Help/AA-UG.pdf"">" & Translate("Account Administrator",Login_language,conn) & " - " & Translate("Users Guide",Login_Language,conn) & "</A></LI>"
  end select
  
  select case Admin_Access
    case 4, 8, 9
      if Site_ID <> 82 then   ' Exclude FNET
        response.write "<LI><A HREF=""/SW-Help/CA-UG.pdf"">" & Translate("Content Administrator",Login_language,conn) & " - " & Translate("Users Guide",Login_Language,conn) & "</A></LI>"
      else
        response.write "<LI><A HREF=""/SW-Help/FNet-CA-UG.pdf"">" & Translate("Content Administrator",Login_language,conn) & " - " & Translate("Users Guide",Login_Language,conn) & "</A></LI>"        
      end if  
  end select
  if Admin_Access >= 4 then
    response.write "</UL>"
  end if
  
  if isblank(request("ID")) then
    response.write "<UL>"
    response.write "<LI>" & Translate("If you are experiencing problems or have questions about this site or site tools, please email them to",Login_Language,conn) & ": <A HREF=""mailto:Webmaster@fluke.com"">" & Translate("Webmaster - Partner Portal Sites",Login_Language,conn) & "</A>.<BR><BR></LI>"
    response.write "<LI>" & Translate("When reporting a problem, please provide the URL and a complete description of the problem by using a copy of the error message or a screen capture.",Login_Language,conn) & "</LI>"
    response.write "</UL>"
  end if   
  
end if

%>     
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

%>

<SCRIPT LANGUAGE="JAVASCRIPT">
<!--//

var FormName   = document.<%=FormName%>
var FormName_0 = document.<%=FormName_0%>

function SearchForItem() {

  var mystyle        = "height=500,width=700,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no,resizable=yes";
  
  if (FormName.Search_Parameter.value != "") {
    var CkItemWindow = window.open('/sw-administrator/SW-Ck_Item_Number.asp?Search=on&Asset_ID=' + FormName.Search_Parameter.value + '&Site_ID=<%=Site_ID%>' + '&Language=<%=Login_Language%>' + '&FormName=<%=FormName%>','CkItemWinddow',mystyle);
    FormName.Search_Parameter.focus();
    FormName.Search_Parameter.value = "";
    
  }
}

function SearchForItem_0() {

  var mystyle        = "height=500,width=700,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no,resizable=yes";
  
  if (FormName_0.Search_Parameter_0.value != "") {
    var CkItemWindow = window.open('/sw-administrator/SW-Ck_Item_Number.asp?Search=on&Asset_ID=' + FormName_0.Search_Parameter_0.value + '&Site_ID=<%=Site_ID%>' + '&Language=<%=Login_Language%>' + '&FormName=<%=FormName_0%>','CkItemWinddow',mystyle);
    FormName_0.Search_Parameter_0.focus();
    FormName_0.Search_Parameter_0.value = "";
    
  }
}

function handleEnter (field, event) {
  var keyCode = event.keyCode ? event.keyCode : event.which ? event.which : event.charCode;
	if (keyCode == 13) {
	  var i;
		for (i = 0; i < field.form.elements.length; i++)
		  if (field == field.form.elements[i])
			break;
			i = (i + 1) % field.form.elements.length;
			field.form.elements[i].focus();
			return false;
	} 
	else
	return true;
}      

//-->
</script>

<SCRIPT Language="Javascript">

  var headTop = -1;
  var FloatHead1;

  function processScroll() {
    if (headTop < 0) {
      saveHeadPos();
    }
    if (headTop > 0) {
      if (document.documentElement && document.documentElement.scrollTop)
        theTop = document.documentElement.scrollTop;
      else if (document.body)
        theTop = document.body.scrollTop;

    if (theTop > headTop)
      FloatHead1.style.top = (theTop-headTop) + 'px';
    else
      FloatHead1.style.top = '0px';
  }
}

function saveHeadPos() {
  parTable = document.getElementById("ContentTableStart");
  if (parTable != null) {
    headTop = parTable.offsetTop + 3;
    FloatHead1 = document.getElementById("ContentHeader1");
    FloatHead1.style.position = "relative";
  }
}

window.onscroll = processScroll;

//Added on 16th Aug 2010, for Enable CDN  checkbox functionality
function enableCDNState(frmCDN)
{
    var cbCDN = document.getElementById("Enable_CDN");
    if (cbCDN != null)
    {
        if (cbCDN.checked)
        {
             document.getElementById("CheckedDisabled").disabled = false;
             document.getElementById("CheckedEnabled").disabled = false;
             document.getElementById("UncheckedEnabled").disabled = false;
        }
        else
        {
             document.getElementById("CheckedDisabled").disabled = true;
             document.getElementById("CheckedEnabled").disabled = true;
             document.getElementById("UncheckedEnabled").disabled = true;
        }
    }
    
}
enableCDNState();
</SCRIPT>


  
<!--#include virtual="/include/RTEditor/RTE_Editor_Launch.asp"-->