<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     K. D. Whitlock
' Date:       06/1/2000
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------
 
Dim Site_ID
Dim Site_ID_Change
Dim Site_Clicks
Dim Region
Dim Utility_ID
Dim Unique_Item_Numbers

Dim Border_Toggle
Border_Toggle = 0

Dim Begin_Date, Country_Code, Local_Code

Site_ID        = request("Site_ID")
Site_ID_Change = Site_ID

if isnumeric(request("Utility_ID")) then
  Utility_ID     = CInt(request("Utility_ID"))
else
  Utility_ID     = 999
end if
  
if site_id = 11 and utility_id = 53 then
  response.redirect("/calibrators/metcal_admin.asp")
end if

Login_Language = "eng"

%>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/SW-Order_Inquiry_Literature_OStatus.asp"-->  
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_FormData.asp"-->
<%

Call Connect_SiteWide

%>
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%

' --------------------------------------------------------------------------------------
' Determine Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

select case Utility_ID
  case 0
    Screen_TitleX = "Accounts by Name"
  case 1
    Screen_TitleX = "Accounts by Company - All"
  case 2
    Screen_TitleX = "Accounts by Company - Non Fluke"

  case 20
    Screen_TitleX = "Main / Branch Location Administrators"
  case 21
    Screen_TitleX = "Account Managers"
  case 22
    Screen_TitleX = "Content Submitters"
  case 23
    Screen_TitleX = "Content Administrators"
  case 24
    Screen_TitleX = "Content Administrators Matrix"
    
  case 30
    Screen_TitleX = "Account Administrators"
  case 31
    Screen_TitleX = "Account Administrators Matrix"

  case 39
    Screen_TitleX = "Site Administrators"

  case 40  
    Screen_TitleX = "Order Inquiry Administrator"    
  case 41  
    Screen_TitleX = "Order Inquiry Search"    

  case 50
    Screen_TitleX = "Content or Event - All"

  case 51
    Screen_TitleX = "Literature Fulfillment - All"

  case 52
    Screen_TitleX = "Literature Fulfillment - Active"
    
  case 53
    Screen_TitleX = "Met/Cal Procedure Administration"

  Case 54
    Screen_TitleX = "Subscription Service - Queue"      

  case 60
    Screen_TitleX = "Thumbnail Requests"
          
  case 70
    Screen_TitleX = "Asset Activity Detail (Metrics)"

  case 71
    Screen_TitleX = "Site Activity Summary (Metrics)"

  case 72
    Screen_TitleX = "WWW Asset Activity Detail (Metrics)"

  case 73
    Screen_TitleX = "Literature Order Activity Detail (Metrics)"

  case 90
    Screen_TitleX = "File - Utility Program"
    
  case 98
    Screen_TitleX = "File Upload Monitor"    

  case 99  
    Screen_TitleX = "List - Directory Contents"
    
  case else
    Screen_TitleX = "Undefined Site Utility Option"
    
end select

Dim RegionColor(3)
RegionColor(0) = "#0000CC"
RegionColor(1) = "#99FFCC"
RegionColor(2) = "#66CCFF"
RegionColor(3) = "#FFCCFF"                                                    

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code        = rsSite("Site_Code")     
Screen_Title     = rsSite("Site_Description") & " - " & Screen_TitleX
Bar_Title        = rsSite("Site_Description") & "<BR><FONT CLASS=SmallBoldGold>" & Screen_TitleX & "</FONT>"
Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<%

rsSite.close
set rsSite=nothing

response.flush

if Admin_Access = 1 or Admin_Access >= 3 then
 
  ' --------------------------------------------------------------------------------------
  ' List Accounts by Name
  ' --------------------------------------------------------------------------------------
  
  if Utility_ID = 0 and Admin_Access >= 4 then
  
    UserNumber = 0
    
    SQL =  "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & Site_ID & " AND UserData.Fcm<>" & CInt(True)
  
    if Admin_Access = 6 then
      SQL = SQL & " AND UserData.Region=" & Admin_Account_Region
    end if
      
    SQL = SQL & " ORDER BY Userdata.LastName"
  
    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQL, conn, 3, 3
    
    if rsUser.EOF and rsUser.BOF then
      response.write "There are no User Accounts for this site.<BR><BR>"
      TableOn = false
    else   
      TableOn = true
  
      Call Nav_Border_Begin
      Call Main_Menu
      response.write "&nbsp;&nbsp;&nbsp;"
      Call Group_Code_Table
      Call Nav_Border_End
      response.write "<BR>"
      %>
  
    	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
        <TR>
          <TD>
            <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
              <TR>
                <TD BGCOLOR="Red" ALIGN="CENTER" CLASS=SmallBoldWhite>Action</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>ID</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Users Name</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Company</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>City</TD>              
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>State</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Country</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Phone Number</TD>            
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Groups</TD>            
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Account Status</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Last Logon</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Account Expiration</TD>                                          
          		</TR>
     <%
     end if
     
     Do while not rsUser.EOF
       if instr(1,lcase(rsUser("SubGroups")),"admin") = 0 then
      %>
          		<TR>
          			<TD BGCOLOR="Silver" ALIGN="CENTER" CLASS=Small>
                  <A HREF="account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>" CLASS=NavLeftHighlight1 onClick="location.href='account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>'" VALUE=" Edit ">&nbsp;&nbsp;Edit&nbsp;&nbsp;</A>
                </TD>
                
          			<TD BGCOLOR="#FFFFFF" ALIGN="RIGHT" CLASS=Small>
                  <% response.write rsUser("ID") %>
                </TD>
  
           			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%
                  response.write "<B>" & rsUser("LastName") & "</B>, "
                  response.write rsUser("FirstName")
                  if not isblank(rsUser("MiddleName")) then response.write " " & rsUser("MiddleName")
                  if not isblank(rsUser("Prefix")) then response.write " " & rsUser("Prefix") & ". "
                  %>               
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Company") %>
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Business_City") %>
                </TD>
         
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Business_State") %>
                </TD>
          			
                <%              
                response.write "<TD BGCOLOR=""" & RegionColor(rsUser("Region")) & """ ALIGN=""LEFT"" CLASS=Small>"
                response.write rsUser("Business_Country")
                response.write "</TD>"
                %>
       
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write FormatPhone(rsUser("Business_Phone")) %>
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%
                  Call Write_SubGroups
                  %>
                </TD>
  
          			<TD BGCOLOR="#FFFFFF" ALIGN="CENTER" CLASS=Small>
                  <%
                  Call Write_Account_Status
                  %>
                </TD>
  
                <TD BGCOLOR="#FFFFFF" ALIGN="CENTER" CLASS=Small>
                  <%
                  Call Write_Last_Logon
                  %>                
                </TD>
  
                  <%
                  Call Write_Expiration_Date
                  %>
                </TD>              
              </TR>
    <%
      end if 
      rsUser.MoveNext
    
    loop
                         
    rsUser.close
    set rsManager=nothing
  
    if TableOn then
      %>
            </TABLE>
          </TD>
        </TR>
      </TABLE>        
      <%
      
    end if                     
  
  ' --------------------------------------------------------------------------------------
  ' List Accounts by Company - All / Non-Fluke
  ' --------------------------------------------------------------------------------------
    
  elseif (Utility_ID = 1  or Utility_ID = 2) and Admin_Access >= 4 then
  
    UserNumber = 0
    
    if Utility_ID = 1 then
      SQL =  "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & Site_ID
    else  
      SQL =  "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & Site_ID & " AND UserData.Company NOT LIKE '%Fluke%'"
    end if
        
    if Admin_Access = 6 then
      SQL = SQL & " AND UserData.Region=" & Admin_Account_Region
    end if
    
    SQL = SQL & " ORDER BY UserData.Company, UserData.LastName"
    
    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQL, conn, 3, 3
    
    if rsUser.EOF and rsUser.BOF then
      response.write "There are no User Accounts established for this site.<BR><BR>"
      TableOn = false
    else   
      TableOn = true
  
      Call Nav_Border_Begin
      Call Main_Menu
      response.write "&nbsp;&nbsp;&nbsp;"
      Call Group_Code_Table
      Call Nav_Border_End
      response.write "<BR>"
    %>
    	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
        <TR>
          <TD>
            <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
              <TR>
                <TD BGCOLOR="Red" ALIGN="CENTER" CLASS=SmallBoldWhite>Action</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>ID</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Company</TD>              
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Users Name</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>City</TD>              
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>State</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Country</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Phone Number</TD>            
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Groups</TD>            
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Account Status</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Last Logon</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Account Expiration</TD>                                          
          		</TR>
     <%
     end if
     
     Do while not rsUser.EOF
       if instr(1,lcase(rsUser("SubGroups")),"admin") = 0 then
      %>
          		<TR>
          			<TD BGCOLOR="Silver" ALIGN="CENTER" CLASS=Small>
                  <A HREF="account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>" CLASS=NavLeftHighlight1 onClick="location.href='account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>'" VALUE=" Edit ">&nbsp;&nbsp;Edit&nbsp;&nbsp;</A>
                </TD>
                
          			<TD BGCOLOR="#FFFFFF" ALIGN="RIGHT" CLASS=Small>
                  <% response.write rsUser("ID") %>
                </TD>
  
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Company") %>
                </TD>
           			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%
                  response.write "<B>" & rsUser("LastName") & "</B>, "
                  response.write rsUser("FirstName")
                  if not isblank(rsUser("MiddleName")) then response.write " " & rsUser("MiddleName")
                  if not isblank(rsUser("Prefix")) then response.write " " & rsUser("Prefix") & ". "
                  %>               
                </TD>
          			        			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Business_City") %>
                </TD>
         
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Business_State") %>
                </TD>
          			
                <%              
                response.write "<TD BGCOLOR=""" & RegionColor(rsUser("Region")) & """ ALIGN=""LEFT"" CLASS=Small>"
                response.write rsUser("Business_Country")
                response.write "</TD>"
                %>
       
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write FormatPhone(rsUser("Business_Phone")) %>
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%
                  Call Write_SubGroups
                  %>
                </TD>
  
          			<TD BGCOLOR="#FFFFFF" ALIGN="CENTER" CLASS=Small>
                  <%
                  Call Write_Account_Status
                  %>
                </TD>
  
                <TD BGCOLOR="#FFFFFF" ALIGN="CENTER" CLASS=Small>
                  <%
                  Call Write_Last_Logon
                  %>                
                </TD>
  
                  <%
                  Call Write_Expiration_Date
                  %>
                </TD>
              </TR>
    <%
      end if 
      rsUser.MoveNext
    
    loop
                         
    rsUser.close
    set rsManager=nothing
  
    if TableOn then
      %>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
      <%
      
    end if                       
  
  ' --------------------------------------------------------------------------------------
  ' List Accounts by Account Manager Name
  ' --------------------------------------------------------------------------------------
  
  elseif Utility_ID = 21 and Admin_Access >= 4 then
  
    UserNumber = 0
    
    SQL =  "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & Site_ID & " AND UserData.Fcm=" & CInt(True)
    
    if Admin_Access = 6 then
      SQL = SQL & " AND UserData.Region=" & Admin_Account_Region
    end if
      
    SQL = SQL & " ORDER BY UserData.Region, Userdata.LastName"
  
    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQL, conn, 3, 3
    
    if rsUser.EOF then
      response.write "There are no Account Manager Accounts for this site.<BR><BR>"
      TableOn = false
    else   
      TableOn = true
  
      Call Nav_Border_Begin
      Call Main_Menu
      response.write "&nbsp;&nbsp;&nbsp;"
      Call Group_Code_Table
      Call Nav_Border_End
      response.write "<BR>"
    %>
    	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
        <TR>
          <TD>
            <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
              <TR>
                <TD BGCOLOR="Red" ALIGN="CENTER" CLASS=SmallBoldWhite>Action</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>ID</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Users Name</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Company</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>City</TD>              
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>State</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Country</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Phone Number</TD>            
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Groups</TD>            
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Account Status</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Last Logon</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Account Expiration</TD>                                          
          		</TR>
     <%
     end if
     
     Do while not rsUser.EOF
       if instr(1,lcase(rsUser("SubGroups")),"admin") = 0 then
      %>
          		<TR>
          			<TD BGCOLOR="Silver" ALIGN="CENTER" CLASS=Small>
                  <A HREF="account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>" CLASS=NavLeftHighlight1 onClick="location.href='account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>'" VALUE=" Edit ">&nbsp;&nbsp;Edit&nbsp;&nbsp;</A>              
                </TD>
                
          			<TD BGCOLOR="#FFFFFF" ALIGN="RIGHT" CLASS=Small>
                  <% response.write rsUser("ID") %>
                </TD>
  
           			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%
                  response.write "<B>" & rsUser("LastName") & "</B>, "
                  response.write rsUser("FirstName")
                  if not isblank(rsUser("MiddleName")) then response.write " " & rsUser("MiddleName")
                  if not isblank(rsUser("Prefix")) then response.write " " & rsUser("Prefix") & ". "
                  %>               
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Company") %>
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Business_City") %>
                </TD>
         
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Business_State") %>
                </TD>
          			
                <%              
                response.write "<TD BGCOLOR=""" & RegionColor(rsUser("Region")) & """ ALIGN=""LEFT"" CLASS=Small>"
                response.write rsUser("Business_Country")
                response.write "</TD>"
                %>
       
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write FormatPhone(rsUser("Business_Phone")) %>
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%
                  Call Write_SubGroups
                  %>
                </TD>
  
          			<TD BGCOLOR="#FFFFFF" ALIGN="CENTER" CLASS=Small>
                  <%
                  Call Write_Account_Status
                  %>
                </TD>
  
                <TD BGCOLOR="#FFFFFF" ALIGN="CENTER" CLASS=Small>
                  <%
                  Call Write_Last_Logon
                  %>                
                </TD>
  
                  <%
                  Call Write_Expiration_Date
                  %>
                </TD>
              </TR>
    <%
      end if 
      rsUser.MoveNext
    
    loop
                         
    rsUser.close
    set rsManager=nothing
  
    if TableOn then
      %>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
      <%
          
    end if
    
  ' --------------------------------------------------------------------------------------
  ' Content Administrator's Matrix
  ' --------------------------------------------------------------------------------------
  
  elseif Utility_ID = 24 and Admin_Access >= 4 then
  
    if request("Toggle") = "True" and isnumeric(request("Approver_ID")) then
    
      SQL = "UPDATE Approvers SET Approvers.Approver_ID =" & request("Approver_ID") & " WHERE (((Approvers.ID)=" & request("Group_ID") & "))"
   	  conn.Execute(SQL)
  
    end if
    
    Toggle  = false    
    TableOn = True
    
    SQL = "SELECT Approvers.* FROM Approvers WHERE Approvers.Site_ID=" & CInt(Site_ID) & " ORDER BY Approvers.Order_Num, Approvers.Description"
    
    Set rsApproverGroups = Server.CreateObject("ADODB.Recordset")
    rsApproverGroups.Open SQL, conn, 3, 3    
    
    if rsApproverGroups.EOF then
      response.write "There are no Content Administrator Groups established for this site.<BR><BR>"
      TableOn = false
    end if
   
    SQL = "Select UserData.* FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND (UserData.Subgroups LIKE '%content%' OR UserData.Subgroups LIKE '%administrator%') ORDER BY UserData.LastName"
    Set rsApproverNames = Server.CreateObject("ADODB.Recordset")
    rsApproverNames.Open SQL, conn, 3, 3
    
    if rsApproverNames.EOF then
      response.write "There are no Content Administrators established for this site.<BR><BR>"
      TableOn = false
    end if
  
    if TableOn = True then  
  
      Call Nav_Border_Begin
      Call Main_Menu
      Call Nav_Border_End
      response.write "<BR>"
      
    %>
      <FORM NAME="Dummy-5">
    	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
        <TR>
          <TD>
            <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
              <TR>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Region / Group / Sub-Region or Description</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Content Administrator&acute; Name</TD>
          		</TR>
     <%
        
      Do while not rsApproverGroups.EOF
       
        response.write "<TR>"
        response.write "<TD BGCOLOR=""" & RegionColor(rsApproverGroups("Region")) & """ ALIGN=""LEFT"" CLASS=Medium VALIGN=MIDDLE>"
        response.write rsApproverGroups("Description")
        response.write "</TD>"
    
        rsApproverNames.MoveFirst
        
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Medium VALIGN=MIDDLE>"
        %>
        <SELECT LANGUAGE="JavaScript" ONCHANGE="window.location.href='site_utility.asp?Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Toggle=True&Group_ID=<%=rsApproverGroups("ID")%>&Approver_ID='+this.options[this.selectedIndex].value" NAME="Approver_ID">
        <%                        
  
        response.write "<OPTION VALUE=""0"" CLASS=NavLeftHighlight1>Select from list</OPTION>" & vbCrLf
        
        Do while not rsApproverNames.EOF    
          response.write "<OPTION CLASS=Region" & rsApproverNames("Region") & "NavMedium VALUE=""" & rsApproverNames("ID") & """"
          if isnumeric(rsApproverGroups("Approver_ID")) then
            if CInt(rsApproverNames("ID")) = CInt(rsApproverGroups("Approver_ID")) then
              response.write " SELECTED"
            end if
          end if    
          response.write ">" & rsApproverNames("LastName") & ", " & rsApproverNames("FirstName") & "</OPTION>" & vbCrLf
          rsApproverNames.MoveNext
        Loop
        
        response.write "</SELECT>" & vbCrLf
        response.write "</TD>" & vbCrLf
        response.write "</TR>" & vbCrLf
        
        rsApproverGroups.MoveNext
       
      Loop
            
    end if  
    
    rsApproverGroups.close
    set rsApproverGroups = nothing
    
    rsApproverNames.close
    set rsApproverNames = nothing  
    
    if TableOn then
      %>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
      </FORM>
      <%
              
    end if

  ' --------------------------------------------------------------------------------------
  ' Account Administrator's Matrix
  ' --------------------------------------------------------------------------------------
  
  elseif Utility_ID = 31 and Admin_Access >= 4 then
  
    if request("Toggle") = "True" and isnumeric(request("Approver_ID")) then
    
      SQL = "UPDATE Approvers_Account SET Approvers_Account.Approver_ID =" & request("Approver_ID") & " WHERE (((Approvers_Account.ID)=" & request("Group_ID") & "))"
   	  conn.Execute(SQL)
  
    end if
    
    Toggle  = false    
    TableOn = True
    
    SQL = "SELECT Approvers_Account.* FROM Approvers_Account WHERE Approvers_Account.Site_ID=" & CInt(Site_ID) & " ORDER BY Approvers_Account.Order_Num, Approvers_Account.Description"
    
    Set rsApproverGroups = Server.CreateObject("ADODB.Recordset")
    rsApproverGroups.Open SQL, conn, 3, 3    
    
    if rsApproverGroups.EOF then
      response.write "There are no Account Administrator Groups established for this site.<BR><BR>"
      TableOn = false
    end if
   
    SQL = "Select UserData.* FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND (UserData.Subgroups LIKE '%account%' OR UserData.Subgroups LIKE '%administrator%') ORDER BY UserData.LastName"
    Set rsApproverNames = Server.CreateObject("ADODB.Recordset")
    rsApproverNames.Open SQL, conn, 3, 3
    
    if rsApproverNames.EOF then
      response.write "There are no Account Administrators established for this site.<BR><BR>"
      TableOn = false
    end if
  
    if TableOn = True then  
  
      Call Nav_Border_Begin
      Call Main_Menu
      Call Nav_Border_End
      response.write "<BR><BR>"
      
    %>
      <FORM NAME="Dummy-6">
    	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
        <TR>
          <TD>
            <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
              <TR>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Region</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Account Administrator&acute; Name</TD>
          		</TR>
     <%
        
      Do while not rsApproverGroups.EOF
       
        response.write "<TR>"
        response.write "<TD BGCOLOR=""" & RegionColor(rsApproverGroups("Region")) & """ ALIGN=""LEFT"" CLASS=Medium VALIGN=MIDDLE>"
        response.write rsApproverGroups("Description")
        response.write "</TD>"
    
        rsApproverNames.MoveFirst
        
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Medium VALIGN=MIDDLE>"
        %>
        <SELECT LANGUAGE="JavaScript" ONCHANGE="window.location.href='site_utility.asp?Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Toggle=True&Group_ID=<%=rsApproverGroups("ID")%>&Approver_ID='+this.options[this.selectedIndex].value" NAME="Approver_ID">
        <%                        
  
        response.write "<OPTION VALUE=""0"" CLASS=NavLeftHighlight1>Select from list</OPTION>" & vbCrLf
        
        Do while not rsApproverNames.EOF    
          response.write "<OPTION CLASS=Region" & rsApproverNames("Region") & "NavMedium VALUE=""" & rsApproverNames("ID") & """"
          if isnumeric(rsApproverGroups("Approver_ID")) then
            if CInt(rsApproverNames("ID")) = CInt(rsApproverGroups("Approver_ID")) then
              response.write " SELECTED"
            end if
          end if    
          response.write ">" & rsApproverNames("LastName") & ", " & rsApproverNames("FirstName") & "</OPTION>" & vbCrLf
          rsApproverNames.MoveNext
        Loop
        
        response.write "</SELECT>" & vbCrLf
        response.write "</TD>" & vbCrLf
        response.write "</TR>" & vbCrLf
        
        rsApproverGroups.MoveNext
       
      Loop
            
    end if  
    
    rsApproverGroups.close
    set rsApproverGroups = nothing
    
    rsApproverNames.close
    set rsApproverNames = nothing  
    
    if TableOn then
      %>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
      </FORM>
      <%
              
    end if                     
  
  ' --------------------------------------------------------------------------------------
  ' List Accounts by Branch / Submitter / Content / Account / Site Administrator
  ' --------------------------------------------------------------------------------------
  
  elseif (Utility_ID = 20 or Utility_ID = 22 or Utility_ID = 23 or Utility_ID = 30 or Utility_ID = 39 or Utility_ID = 40 or Utility_ID = 41) and Admin_Access >= 4 then
  
    UserNumber = 0
    
    SQL =  "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & Site_ID
    
    select case Utility_ID
      case "20"   ' Main/Branch Location Administrator
        SQL = SQL & " AND UserData.SubGroups LIKE '%branch%'"
      case "22"   ' Submitter
        SQL = SQL & " AND UserData.SubGroups LIKE '%submitter%'"
      case "23"   ' Content Administrator
        SQL = SQL & " AND UserData.SubGroups LIKE '%content%'"
      case "30"   ' Account Administrator
        SQL = SQL & " AND UserData.SubGroups LIKE '%account%'"
      case "40"   ' Order Inquiry Administrator
        SQL = SQL & " AND UserData.SubGroups LIKE '%ordadt%'"
      case "41"   ' Order Inquiry Search
        SQL = SQL & " AND UserData.SubGroups LIKE '%order%'"
      case "39"   ' Site Administrator
        SQL = SQL & " AND (UserData.SubGroups LIKE '%administrator%' OR UserData.NTLogin LIKE 'whitlock')"
    end select    

    SQL = SQL & " ORDER BY UserData.Region, Userdata.LastName"
  
    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQL, conn, 3, 3
    
    if rsUser.EOF then
      response.write "There are no Site Administrator Accounts established for this site.<BR><BR>"
      TableOn = false
    else   
      TableOn = true
  
      Call Nav_Border_Begin
      Call Main_Menu
      response.write "&nbsp;&nbsp;&nbsp;"
      Call Group_Code_Table
      Call Nav_Border_End
      response.write "<BR>"
    %>
    	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
        <TR>
          <TD>
            <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
              <TR>
                <TD BGCOLOR="Red" ALIGN="CENTER" CLASS=SmallBoldWhite>Action</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>ID</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Users Name</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Company</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>City</TD>              
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>State</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Country</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Phone Number</TD>            
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Groups</TD>            
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Account Status</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Last Logon</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Account Expiration</TD>                                          
          		</TR>
     <%
     end if
     
     Do while not rsUser.EOF
      %>
          		<TR>
          			<TD BGCOLOR="Silver" ALIGN="CENTER" CLASS=Small>
                  <%if Admin_Access = 9 or (Utility_ID=20 or Utility_ID=22 or Utility_ID=23 ) then%>
                  <A HREF="account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>" CLASS=NavLeftHighlight1 onClick="location.href='account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>'" VALUE=" Edit ">&nbsp;&nbsp;Edit&nbsp;&nbsp;</A>
                  <%else%>
                  No Edit
                  <%end if%>
                </TD>
                
          			<TD BGCOLOR="#FFFFFF" ALIGN="RIGHT" CLASS=Small>
                  <% response.write rsUser("ID") %>
                </TD>
  
           			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%
                  response.write "<B>" & rsUser("LastName") & "</B>, "
                  response.write rsUser("FirstName")
                  if not isblank(rsUser("MiddleName")) then response.write " " & rsUser("MiddleName")
                  if not isblank(rsUser("Prefix")) then response.write " " & rsUser("Prefix") & ". "
                  %>               
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Company") %>
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Business_City") %>
                </TD>
         
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write rsUser("Business_State") %>
                </TD>
          			
                <%              
                response.write "<TD BGCOLOR=""" & RegionColor(rsUser("Region")) & """ ALIGN=""LEFT"" CLASS=Small>"
                response.write rsUser("Business_Country")
                response.write "</TD>"
                %>
       
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <% response.write FormatPhone(rsUser("Business_Phone")) %>
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%
                  Call Write_SubGroups
                  %>
                </TD>
  
          			<TD BGCOLOR="#FFFFFF" ALIGN="CENTER" CLASS=Small>
                  <%
                  Call Write_Account_Status
                  %>
                </TD>
                
          			<TD BGCOLOR="#FFFFFF" ALIGN="CENTER" CLASS=Small>              
                  <%
                  Call Write_Last_Logon
                  %>                              
                </TD>
  
                  <%
                  Call Write_Expiration_Date
                  %>              
                </TD>
              </TR>
    <%
      rsUser.MoveNext
    
    loop
                         
    rsUser.close
    set User = nothing
  
    if TableOn then
      %>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
      <%

    end if

  ' --------------------------------------------------------------------------------------
  ' List - Content / Event - All
  ' View - Thumbnail Requests
  ' --------------------------------------------------------------------------------------
  
  elseif (Utility_ID = 50 or Utility_ID = 51 or Utility_ID = 52 or Utility_ID=54 or Utility_ID = 60) and (Admin_Access = 4 or Admin_Access >= 8) then

    if request("Category_ID") <> "" and Utility_ID <> 54 then
      Category_ID = CInt(request("Category_ID"))
    elseif Utility_ID = 54 then
      Category_ID = -1      
    else
      Category_ID = 0
    end if  
    
    if request("View") <> "" then
      View = CInt(request("View"))
    else
      View = 0
    end if
    
    if request("Campaign") <> "" then
      Campaign = CInt(request("Campaign"))
    else
      Campaign = 0
    end if
    
    if request("Group_ID") <> "" then
      Group_ID = LCase(request("Group_ID"))
    else
      Group_ID = ""
    end if    
    
    if request("Country") <> "" then
      Country = UCase(request("Country"))
    else
      Country = ""
    end if
    
    if request("LDate") <> "" then
      LDate = request("LDate")
    else
      LDate = Date()
    end if
    
    if request("Sort_By") <> "" then
      Sort_By = request("Sort_By")
    else
      Sort_By = 0
    end if
    
    if request("Subject") <> "" and request("Subject") <> "Todays News" and request("Subject") <> "*" then
    
      Subject = request("Subject")
    
      SQLSubject = "SELECT * FROM Subscription_Subject WHERE Site_ID=" & Site_ID & " AND LDate='" & LDate & "'"
      Set rsSubject = Server.CreateObject("ADODB.Recordset")
      rsSubject.Open SQLSubject, conn, 3, 3
      
      if rsSubject.EOF then
        SQLSubject = "INSERT INTO Subscription_Subject (Site_ID, LDate, Subject_Text) VALUES (" & Site_ID & ", '" & LDate & "', '" & Replace(Subject,"'","") & "')"
        conn.execute SQLSubject
      else
        SQLSubject = "UPDATE Subscription_Subject SET Subject_Text='" & Replace(Subject,"'","") & "' WHERE Site_ID=" & Site_ID & " AND LDate='" & LDate & "'"  
        conn.execute SQLSubject
      end if

      rsSubject.close
      set rsSubject  = nothing      

    elseif request("Subject") = "*" then
      SQLSubject = "DELETE FROM Subscription_Subject WHERE Site_ID=" & Site_ID & " AND LDate='" & LDate & "'"
      conn.execute SQLSubject
      Subject = ""
    else
    
      SQLSubject = "SELECT * FROM Subscription_Subject WHERE Site_ID=" & Site_ID & " AND LDate='" & LDate & "'"
      Set rsSubject = Server.CreateObject("ADODB.Recordset")
      rsSubject.Open SQLSubject, conn, 3, 3
    
      if not rsSubject.EOF then
        Subject = rsSubject("Subject_Text")
      end if
      
      rsSubject.close
      set rsSubject  = nothing      
        
    end if
    
    set SQLSubject = nothing
  
    response.write "<TABLE WIDTH=""100%"" border=0>"
    response.write "<TR>"
    response.write "<TD CLASS=Small WIDTH=""50%"" ROWSPAN=5 VALIGN=TOP>"
        
    Call Nav_Border_Begin    
    Call Main_Menu
'    response.write "&nbsp;"
    Call Group_Code_Table
'    response.write "&nbsp;"
    Call Status_Colors
'    response.write "&nbsp;"
    Call Element_Names

    if Group_ID <> "" or Country <> "" then
      response.write "&nbsp;&nbsp;"
      %>      
      <A HREF="site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&View=<%=View%>&Group_ID=&Country=&Category_ID=<%=Category_ID%>&LDate=<%=LDate%>" ONCLICK="window.location.href='site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&View=<%=View%>&Group_ID=&Country=&Category_ID=<%=Category_ID%>&LDate=<%=LDate%>'"><SPAN CLASS=NavLeftHighlight1>&nbsp;Clear&nbsp;Filters&nbsp;</SPAN></A>      
      <%
    end if
    
    Call Nav_Border_End
    
    if Utility_ID = 54 then
      response.write "<P>"
      response.write "<SPAN CLASS=SmallBold>" & Translate("Subscription Email Date",Login_Language,conn) & "</SPAN>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
      response.write "<INPUT TYPE=""TEXT"" VALUE=""" & LDate & """ MAXLENGTH=""10"" SIZE=""7"" CLASS=Small NAME=""LDate"" "
      response.write "ONCHANGE=""window.location.href='site_utility.asp?ID=site_utility&Site_ID=" & Site_ID & "&Utility_ID=" & Utility_ID & "&View=" & View & "&LDate=' + this.value + '" & "&Group_ID=" & Group_ID & "&Country=" & Country & "&Category_ID=" & Category_ID & "&Subject=" & Subject & "'"">"
      response.write "&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=BUTTON VALUE=""" & Translate("Go",Login_Language,conn) & """>"

'EVII  (remove 1=2 when live)      
      if 1=2 and Admin_Access >=8 then
        response.write "<BR>"
        response.write "<SPAN CLASS=SmallBold>" & Translate("Subscription Subject",Login_Language,conn) & "</SPAN>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        response.write "<SPAN CLASS=SmallBoldRed>" 
        if Subject = "" then
          response.write Translate("Todays News",Login_Language,conn)
        else
          response.write Subject
        end if
        response.write "</SPAN>"
        response.write "<BR>"
        response.write "<SPAN CLASS=SmallBold>" & Translate("Subscription Subject",Login_Language,conn) & " (" & Translate("Change",Login_Language,conn) & ")</SPAN>: "        
        response.write "<INPUT TYPE=""TEXT"" VALUE="""" MAXLENGTH=""60"" SIZE=""20"" CLASS=Small NAME=""Subject"" "
        response.write "ONFOCUS=""alert('Changing the Subject text of the Subscription Email will be in English only for all recipients.\n\nTo Change back to the default text, \'Todays News\', just enter an asterisk *.');"" "
        response.write "ONCHANGE=""window.location.href='site_utility.asp?ID=site_utility&Site_ID=" & Site_ID & "&Utility_ID=" & Utility_ID & "&View=" & View & "&Subject=' + this.value + '" & "&LDate=" & LDate & "&Group_ID=" & Group_ID & "Sort_By=" & Sort_By & "&Country=" & Country & "&Category_ID=" & Category_ID & "'"">"
      end if  
    end if
    response.write "<P>"     
    response.write "</TD>"

    response.write "<TD CLASS=Small>"
    response.write "<SPAN CLASS=SmallBold>" & Translate("Category",Login_Language,conn) & ":</SPAN>"
    response.write "</TD>"

    response.write "<TD CLASS=Small>"
    %>        
    <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&View=<%=View%>&LDate=<%=LDate%>&Group_ID=<%=Group_ID%>&Country=<%=Country%>&Sort_By=<%=Sort_By%>&Category_ID='+this.options[this.selectedIndex].value" NAME="Category_ID">
    <%

    if Admin_Access <> 2 then
      SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=" & CInt(True) & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
    else
      SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=" & CInt(True) & " AND (Calendar_Category.Code < 8000 OR Calendar_Category > 8999) ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
    end if

    Set rsCategory = Server.CreateObject("ADODB.Recordset")
    rsCategory.Open SQL, conn, 3, 3

    if Campaign <> 0 then
      response.write "<OPTION CLASS=Medium SELECTED VALUE="""">" & Translate("PI / C Listing",Login_Language,conn) & "</OPTION>" & vbCrLF
    end if
      
    Do while not rsCategory.EOF

      select case rsCategory("Code")
        case 8000
          response.write "<OPTION Class=Region1"          
        case 8001
          response.write "<OPTION Class=Region2"          
        case else
          response.write "<OPTION Class=Medium"
      end select
      
      if CInt(request("Category_ID")) = rsCategory("ID") then
     	  response.write " SELECTED"
      end if   

  	  response.write " VALUE=""" & rsCategory("ID") & """>"
        
      SQL = "SELECT Count(*)AS Count FROM Calendar WHERE Code=" & rsCategory("Code") & " AND Site_ID=" & Site_ID
      if Group_ID <> "" then
        SQL = SQL & " AND SubGroups LIKE '%" & Group_ID & "%'"
      end if
      if Country <> "" then
        SQL = SQL & " AND (Country = 'none' OR Country LIKE '%0%' AND Country NOT LIKE '%" & Login_Country & "%' OR Country NOT LIKE '%0%' AND Country LIKE '%" & Login_Country & "%')"
'       SQL = SQL & " AND (Country='none' OR Country LIKE '%" & Country & "%')"
      end if  
      Set rsCount = Server.CreateObject("ADODB.Recordset")
      rsCount.Open SQL, conn, 3, 3
      
      if rsCount("Count") > 0 then
        response.write "+ "
      else
        response.write "o "
      end if
      
      rsCount.Close
      set rsCount = nothing  
      
      response.write Translate(RestoreQuote(rsCategory("Title")),Login_Language,conn) & "</OPTION>"
      if Category_ID = 0 then
        Category_ID   = rsCategory("ID")
        Category_Code = rsCategory("Code")
      end if  
               
  	  rsCategory.MoveNext 
    loop
        
    rsCategory.close
    set rsCategory=nothing

    response.write "<OPTION Class=MediumRed VALUE=""-1"""
    if Category_ID = -1 then
      response.write " SELECTED"
    end if
    response.write ">+ " & Translate ("List All",Login_Language,conn) & " (" & Translate("Long Listing",Login_Language,conn) & ")</OPTION>"

    response.write "</SELECT>"  
    response.write "</TD>"
    response.write "</TR>"
    
    response.write "<TR>"
    response.write "<TD CLASS=Small>"
    response.write "<SPAN CLASS=SmallBold>" & Translate("View",Login_Language,conn) & ":</SPAN>"

    response.write "<TD CLASS=Small>"
    %>        
    <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Campaign=<%=Campaign%>&Category_ID=<%=Category_ID%>&LDate=<%=LDate%>&Group_ID=<%=Group_ID%>&Country=<%=Country%>&Sort_By=<%=Sort_By%>&View='+this.options[this.selectedIndex].value" NAME="View">
    <%
    response.write "<OPTION CLASS=Medium VALUE=0 "
    if View = 0 then response.write " SELECTED"
    response.write ">" & Translate("Condensed",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "<OPTION CLASS=Medium VALUE=1 "
    if View = 1 then response.write " SELECTED"
    response.write ">" & Translate("Groups",Login_Language,conn) & "</OPTION>" & vbCrLf
    
    response.write "<OPTION CLASS=Medium VALUE=2 "
    if View = 2 then response.write " SELECTED"
    response.write ">" & Translate("Country Restrictions",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "<OPTION CLASS=Medium VALUE=3 "
    if View = 3 then response.write " SELECTED"
    response.write ">" & Translate("Groups",Login_Language,conn) & " + " & Translate("Country Restrictions",Login_Language,conn) & "</OPTION>" & vbCrLf
    
    response.write "<OPTION CLASS=Medium VALUE=4 "
    if View = 4 then response.write " SELECTED"
    response.write ">" & Translate("Product Introduction",Login_Language,conn) & " / " & Translate("Campaign",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "</SELECT>"
    response.write "</TD>"
    response.write "</TR>"
    
    ' Filter Group
    response.write "<TR>"
    response.write "<TD CLASS=Small>"
    response.write "<SPAN CLASS=SmallBold>" & Translate("Filter by Group",Login_Language,conn) & ": </SPAN>"
    response.write "</TD>"
    
    response.write "<TD CLASS=Small>"
    %>
    <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='/SW-Administrator/site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Campaign=<%=Campaign%>&LDate=<%=LDate%>&Category_ID=<%=Category_ID%>&View=<%=View%>&Country=<%=Country%>&Sort_By=<%=Sort_By%>&Group_ID='+this.options[this.selectedIndex].value" NAME="Group_ID">
    <%
    SQL = "SELECT SubGroups.* FROM SubGroups WHERE SubGroups.Site_ID=" & CInt(Site_ID) & " AND SubGroups.Order_Num <> 99 AND SubGroups.Enabled=" & CInt(True) & " ORDER BY SubGroups.Order_Num"
    Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
    rsSubGroups.Open SQL, conn, 3, 3

    response.write "<OPTION Class=Small VALUE="""">" & Translate("No Filter",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""""></OPTION>"
                          
    Do while not rsSubGroups.EOF            
      if request("Group_ID") = rsSubGroups("Code") then
     	  response.write "<OPTION SELECTED VALUE=""" & rsSubGroups("Code") & """"
        if rsSubGroups("Enabled") = True then
          response.write " CLASS=Region" & Trim(rsSubGroups("Region")) & "NavMedium>+ "
        else
          response.write " CLASS=RegionXNavMedium>o "
        end if
        response.write RestoreQuote(rsSubGroups("X_Description")) & "</OPTION>" & vbCrLf
      else
     	  response.write "<OPTION VALUE=""" & rsSubGroups("Code") & """"
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
    
    ' Filter by Country or no restrictions

    response.write "<TR>"
    response.write "<TD Class=Small>"
    response.write "<SPAN CLASS=SmallBold>" & Translate("Filter by Country",Login_Language,conn) & ": </SPAN>"
    response.write "</TD>"

    response.write "<TD CLASS=Small>"
    %>
    <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='/SW-Administrator/site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Campaign=<%=Campaign%>&LDate=<%=LDate%>&Category_ID=<%=Category_ID%>&View=<%=View%>&Group_ID=<%=Group_ID%>&Sort_By=<%=Sort_By%>&Country='+this.options[this.selectedIndex].value" NAME="Country">
    <%
    
    response.write "<OPTION Class=Small VALUE="""">" & Translate("No Filter",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""""></OPTION>"
    
    'Call Connect_FormDatabase
    Call DisplayCountrySelect(Country, "Small")
    'Call Disconnect_FormDatabase

    response.write "<SELECT>"

    response.write "</TD>"
    response.write "</TR>"

    ' Sort By

    response.write "<TR>"
    response.write "<TD Class=Small>"
    response.write "<SPAN CLASS=SmallBold>" & Translate("Sort By",Login_Language,conn) & ": </SPAN>"
    response.write "</TD>"

    response.write "<TD CLASS=Small>"
    %>
    <SELECT CLASS=Small LANGUAGE="JavaScript" ONCHANGE="window.location.href='/SW-Administrator/site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Campaign=<%=Campaign%>&LDate=<%=LDate%>&Category_ID=<%=Category_ID%>&View=<%=View%>&Group_ID=<%=Group_ID%>&Country=<%=Country%>&Sort_By='+this.options[this.selectedIndex].value" NAME="Sort_By">
    <%
    
    response.write "<OPTION Class=Small VALUE="""">" & Translate("No Filter",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""""></OPTION>"
    response.write "<OPTION VALUE=""1"""
    if Sort_By = 1 then response.write " SELECTED"
    response.write ">" & Translate("Asset ID",Login_language,conn) & "</OPTION>"        
    response.write "<OPTION VALUE=""2"""
    if Sort_By = 2 then response.write " SELECTED"
    response.write ">" & Translate("Item Number",Login_language,conn) & " + " & Translate("Revision",Login_Language,conn) & "</OPTION>"    
    response.write "<OPTION VALUE=""3"""
    if Sort_By = 3 then response.write " SELECTED"
    response.write ">" & Translate("Begin Date",Login_language,conn) & "</OPTION>"
    response.write "<OPTION VALUE=""4"""                
    if Sort_By = 4 then response.write " SELECTED"
    response.write ">" & Translate("Category",Login_language,conn) & "</OPTION>"            

    
    response.write "<SELECT>"

    response.write "</TD>"
    response.write "</TR>"

    response.write "</TABLE>"
    
    ' Display Items for Selected Category

    SQL =       "SELECT dbo.Calendar.*, dbo.Calendar_Category.Title AS Category, dbo.Literature_Items_US.STATUS AS Lit_Status, dbo.Literature_Items_US.STATUS_Name AS Lit_Status_Name, " &_
                "       dbo.Literature_Items_US.[ACTION] AS Lit_Action, dbo.Literature_Items_US.Inventory_Rule AS Lit_Inventory_Rule, dbo.Literature_Items_US.Revision AS Lit_Revision, dbo.Literature_Items_US.POD AS Lit_POD, dbo.Literature_Items_US.[PRINT] AS Lit_Print, " &_
                "       dbo.Literature_Items_US.CD AS Lit_CD, dbo.Literature_Items_US.DISPLAY AS Lit_Display, dbo.Literature_Items_US.VIDEO_NTSC AS Lit_Video_NTSC, " &_
                "       dbo.Literature_Items_US.PDF AS Lit_PDF, dbo.Literature_Items_US.FAX AS Lit_FAX,dbo.Literature_Items_US.WEB AS Lit_WEB, dbo.Literature_Items_US.VIDEO_PAL AS Lit_Video_PAL, UserData.FirstName AS FirstName, UserData.LastName AS LastName " &_

                "FROM   dbo.Calendar LEFT OUTER JOIN " &_
                "       dbo.Literature_Items_US ON dbo.Calendar.Revision_Code = dbo.Literature_Items_US.REVISION AND  " &_
                "       dbo.Calendar.Item_Number = dbo.Literature_Items_US.ITEM LEFT OUTER JOIN " &_
                "       dbo.UserData ON dbo.Calendar.Submitted_By = dbo.UserData.ID LEFT OUTER JOIN " &_
                "       dbo.Calendar_Category ON dbo.Calendar.Category_ID = dbo.Calendar_Category.ID " &_
                "WHERE  dbo.Calendar.Site_ID=" & Site_ID & " "

'                "FROM   Calendar LEFT OUTER JOIN " &_
'                "       UserData ON Calendar.Submitted_By = UserData.ID LEFT OUTER JOIN " &_
'                "       Literature_Items_US ON Calendar.Item_Number = Literature_Items_US.ITEM LEFT OUTER JOIN " &_
'                "       Calendar_Category ON Calendar.Category_ID = Calendar_Category.ID " &_
                
                
    if Campaign <> 0 then
      SQL = SQL  & "AND Calendar.Campaign=" & Campaign & " OR Calendar.ID=" & Campaign & " "
    else  
      if Category_ID > 0 then
        SQL = SQL  & "AND Calendar.Category_ID=" & Category_ID & " "
      end if
    end if  

    if Utility_ID = 51 or Utility_ID = 52 then
      SQL = SQL & "AND Calendar.Subgroups LIKE '%view%'" & " "
    elseif Utility_ID = 60 then
      SQL = SQL & "AND Calendar.Thumbnail_Request=" & CInt(True) & " "
    end if     

    if Group_ID <> "" then
      SQL = SQL & "AND Calendar.Subgroups LIKE '%" & Group_ID & "%'" & " "    
    end if
      
    if Country <> "" then
      SQL = SQL & " AND (Calendar.Country = 'none' OR Calendar.Country NOT LIKE '0%' AND Calendar.Country LIKE '%" & Country & "%')" & " "   
'     SQL = SQL & "AND (Calendar.Country='none' OR Calendar.Country LIKE '%" & Country & "%')" & " "    
    end if

    if Utility_ID = 54 then
      SQL = SQL & " AND Calendar.Subscription=-1 AND Calendar.LDate='" & LDate & "' "
    end if  

    if Campaign <> 0 then
      SQL = SQL & "ORDER BY Calendar_Category.Sort, Calendar.Status, Calendar_Category.Title, Calendar.Sub_Category, Calendar.Product"
    else
      select case Sort_By
        case 1  ' Asset ID
          SQL = SQL & "ORDER BY Calendar.Status, Calendar.ID "
        case 2  ' Item Number + Revision
          SQL = SQL & "ORDER BY Calendar.Item_Number, Calendar.Revision_Code "
        case 3  ' Begin Date
          SQL = SQL & "ORDER BY Calendar.BDate "        
        case 4  ' Category, Sub Category, Begin Date, ID
          SQL = SQL & "ORDER BY Calendar.Code, Calendar.Sub_Category, Calendar.BDate, Calendar.ID "        

        case else
      SQL = SQL & "ORDER BY Calendar.Status, Calendar.Sub_Category, Calendar.Product, Calendar.Item_Number, Calendar.Revision_Code" 'BDate Desc"
      end select    
    end if
    
'response.write SQL & "<P>"

    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQL, conn, 3, 3

    if rsUser.EOF then

      select case Utility_ID
        case 54
          response.write "<BR><HR COLOR=Gray SIZE=3 WIDTH=""100%""><BR>" & Translate("There are no Content or Event Items scheduled for tonight's Subscription Service Email.",Login_Language,conn) & "<BR><BR>"
        case else
          response.write "<BR><HR COLOR=Gray SIZE=3 WIDTH=""100%""><BR>" & Translate("There are no Content or Event Items for Category. Please Select Another Category.",Login_Language,conn) & "<BR><BR>"
          TableOn = false
      end select

    else   

      TableOn = true

      Columns = 0
    	with response
        .write "<SPAN CLASS=SMALL>" & Translate("Total Content or Event Items",Login_Language,conn) & ": " & rsUser.recordcount
        .write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Translate("Current Date/Time",Login_Language,conn) & ": " & Now() & " PST</SPAN><P>"

        Call Table_Begin

        .write "      <TABLE CELLPADDING=2 CELLSPACING=1 BORDER=0  WIDTH=""100%"">"
        .write "        <TR>"
        .write "          <TD BGCOLOR=""Red"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Action",Login_Language,conn) & "</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("ID",Login_Language,conn) & "</TD>"
      end with    
        
      if View = 4 then
        response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("PIC",Login_Language,conn) & "</TD>"
        Columns = Columns + 1
      end if  

      with response  
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>" & Translate("Category",Login_Language,conn) & "</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>" & Translate("Product",Login_Language,conn) & "</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>" & Translate("Title",Login_Language,conn) & "</TD>"
'       .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>I</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>L</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>A</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>Z</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>P</TD>"
'        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>A</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>T</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>U</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>D</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>S</TD>"
      end with

      Columns = Columns + 11
              
      if Utility_ID = 50 or Utility_ID = 54 or Utility_ID = 60 then
        if View = 1 or View = 3 then
          response.write "<TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>" & Translate("Groups",Login_Language,conn) & "</TD>"
          Columns = Columns + 1
        end if  
        if View = 2 or View = 3 then 
          response.write "<TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>" & Translate("Country",Login_Language,conn) & "</TD>"
          Columns = Columns + 1
        end if
      end if  

      if Utility_ID = 50 or Utility_ID = 51 or Utility_ID = 52 or Utility_ID=54 or Utility_ID = 60 then
        with response
          .write "          <TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=SmallBoldGold>" & Translate("Item",Login_Language,conn) & "</TD>"
          .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>" & Translate("Rev",Login_Language,conn) & "</TD>"
        end with
        Columns = Columns + 2
      end if
      
      if Utility_ID = 51 or Utility_ID = 52 or Utility_ID = 60 then
        with response
          .write "          <TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>" & Translate("Legacy",Login_Language,conn) & "</TD>"
        end with
        Columns = Columns + 2
      end if

      with response            
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("LNG",Login_Language,conn) & "</TD>"
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Announce<BR>Date",Login_Language,conn) & "</TD>"
          .write "          <TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Begin",Login_Language,conn) & "<BR>" & Translate("Date",Login_Language,conn) & "</TD>"
        if Utility_ID = 54 then
          .write "          <TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Subscription",Login_Language,conn) & "<BR>" & Translate("Time",Login_Language,conn) & "</TD>"
          .write "          <TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Embargo<BR>Date",Login_Language,conn) & "</TD>"
        else
          .write "          <TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("End<BR>Date",Login_Language,conn) & "</TD>"
          .write "          <TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Embargo<BR>Date",Login_Language,conn) & "</TD>"
        end if
        .write "          <TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Archive<BR>Date",Login_Language,conn) & "</TD>"
        .write "    		</TR>"
      end with  
      Columns = Columns + 6
    end if
     
    Category_Old = ""

    do while not rsUser.EOF
    
      if rsUser("Category") <> Category_Old then
        response.write "<TR>"
        response.write "<TD BGCOLOR=""Silver"" CLASS=SmallBoldGold COLSPAN=2>"
        response.write "&nbsp;"
        response.write "</TD>"

        response.write "<TD BGCOLOR="""
        if CInt(rsUser("Code")) >=8000 and CInt(rsUser("Code")) <=8999 then
          response.write "Red""     CLASS=SmallBoldWhite "
        else  
          response.write "#666666"" CLASS=SmallBoldGold "
        end if
        response.write "COLSPAN=" & Columns & ">"  

        response.write rsUser("Category")
        response.write "</TD>"
        response.write "</TR>"
        Category_Old = rsUser("Category")
      end if
      
      ' ID
      response.write "<TR>"
      response.write "<TD BGCOLOR="""
      if Campaign = rsUser("ID") then
        response.write "#00CC00"
      else
        response.write "#FFFFFF"
      end if    
      response.write """ ALIGN=""CENTER"" CLASS=Small>"      
      response.write "<A HREF=""/sw-administrator/Calendar_Edit.asp?ID=" & rsUser("ID") & "&Site_ID=" & Site_ID & """ CLASS=Navlefthighlight1>&nbsp;&nbsp;" & Translate("Edit",Login_Language,conn) & "&nbsp;&nbsp;</A>" & vbCrLf
      if (Category_Code >= 8000 and Category_Code <= 8999) and Campaign = 0 then
        response.write "<HR NOSHADE COLOR=""BLACK"" SIZE=1>"
        response.write "<A HREF=""/sw-administrator/Site_Utility.asp?ID=Site_Utility" & "&Campaign=" & rsUser("ID") & "&Site_ID=" & Site_ID & "&Utility_ID=" & Utility_ID & "&View=4" & """ CLASS=Navlefthighlight1>&nbsp;&nbsp;" & Translate("List",Login_Language,conn) & "&nbsp;&nbsp;</A>" & vbCrLf
      end if  
      response.write "</TD>"
    
      ' Status
      Status = rsUser("Status")
      select case Status
        Case 1        
          response.write "<TD BGCOLOR=""#00CC00"" ALIGN=""CENTER"" CLASS=Small>"
        case 2
          response.write "<TD BGCOLOR=""#AAAAFF"" ALIGN=""CENTER"" CLASS=Small>"
        case else
          response.write "<TD BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
      end select
      response.write rsUser("ID")
      response.write "</TD>"

      ' PI/C
      if View = 4 then
        response.write "<TD BGCOLOR=""White"" ALIGN=CENTER CLASS=SmallBold NOWRAP>"
        select case CInt(rsUser("Content_Group"))
          case 0
            response.write "I"
          case 1
            response.write "P+I"
          case 2
            response.write "P"
          case 3
            response.write "C+I"
          case 4
            response.write "C"
          case else
            response.write "&nbsp;"  
        end select
        response.write "</TD>"
        Columns = Columns + 1
      end if      

      with response  

        ' Sub-Category
        .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
        .write rsUser("Sub_Category")
        .write "</TD>"

        ' Product or Product Series
        .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
        .write rsUser("Product")
        .write "</TD>"

        ' Title
        .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
        .write rsUser("Title")
        if not isblank(rsUser("LastName")) then
          .write "<BR><SPAN CLASS=Smallest><SPAN STYLE=""COLOR=#999999"">" & Translate("Owner",Login_Language,conn) & ": "
          if not isblank(rsUser("FirstName")) then
            .write Mid(rsUser("FirstName"),1,1) & ". "
          end if
          .write rsUser("LastName") & "</SPAN></SPAN>"
        end if    
        
        .write "</TD>"
      end with

      ' Missing Assets - Warning
      Missing_Assets = False
      if (Category_Code < 8000 or Category_Code > 8999) then
        if isblank(rsUser("Link")) and isblank(rsUser("File_Name")) and isblank(rsUser("File_Name_POD")) then
          Missing_Assets = True
        elseif instr(1,LCase(rsUser("SubGroups")),"view") > 0 and (isblank(rsUser("File_Name")) and isblank(rsUser("File_Name_POD"))) then
          Missing_Assets = True
        elseif instr(1,LCase(rsUser("SubGroups")),"view") > 0 and (isblank(rsUser("File_Name")) or isblank(rsUser("File_Name_POD"))) then
          Missing_Assets = True
        end if
      end if

      ' Include File
'      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"      
'      if not isblank(rsUser("Include")) then
'        response.write "Y"
'      else
'        response.write "&nbsp;"  
'      end if
'      response.write "</TD>"

      ' Link
      if not isblank(rsUser("Link")) then
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "Y"
      elseif Missing_Assets = True and (isblank(rsUser("File_Name")) and isblank(rsUser("File_Name_POD"))) then
        response.write "<TD BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      else
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      end if
      response.write "</TD>"

      ' Asset File
      if not isblank(rsUser("File_Name")) then
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"      
        response.write "Y"
      elseif Missing_Assets = True then
        response.write "<TD BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      else
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      end if
      response.write "</TD>"

      ' Archive File
      if not isblank(rsUser("Archive_Name")) then
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"      
        response.write "Y"
      elseif Missing_Assets = True then
        response.write "<TD BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      else
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      end if
      response.write "</TD>"

      ' POD Asset File
      if not isblank(rsUser("File_Name_POD")) then
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"      
        response.write "Y"
      elseif not isnull(rsUser("Lit_POD")) then
        if CInt(rsUser("Lit_POD")) = CInt(True) then
          response.write "<TD BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        else  
          response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        end if
      else
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      end if
      response.write "</TD>"

if 1=2 then
      ' POD Archive File
      if not isblank(rsUser("Archive_Name_POD")) then
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"      
        response.write "Y"
      elseif Missing_Assets = True then
        response.write "<TD BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      else
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      end if
      response.write "</TD>"
end if
      
      ' Thumbnail
      if not isblank(rsUser("Thumbnail")) then
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"            
        response.write "Y"
      elseif Missing_Assets = True then
        response.write "<TD BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      elseif Missing_Assets = False and isblank(rsUser("Link")) then
        response.write "<TD BGCOLOR=""Red"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      else
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"  
      end if  
      response.write "</TD>"

      ' Electronic Email Fulfillment Viewable

      if instr(1,rsUser("SubGroups"),"view") > 0 then
        response.write "<TD BGCOLOR=""#FF9966"" ALIGN=""CENTER"" CLASS=Small>"            
        response.write "Y"
      else
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"            
        response.write "&nbsp;"  
      end if
      response.write "</TD>"
      
      ' Digital Library Viewable

      if instr(1,rsUser("SubGroups"),"fedl") > 0 then
        response.write "<TD BGCOLOR=""#FF9966"" ALIGN=""CENTER"" CLASS=Small>"            
        response.write "Y"
      else
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"            
        response.write "&nbsp;"  
      end if
      response.write "</TD>"

      ' Shopping Cart

      Missing_Item = ""
      
      Lit_Inventory_Rule = ""
      if not isblank(rsUser("Lit_Inventory_Rule")) then
        Lit_Inventory_Rule = LCase(Replace(rsUser("Lit_Inventory_Rule")," ","_"))
      end if  

      if LCase(rsUser("Lit_Status")) = "retired" then   ' Retired (This should never happen, since retired status is prechecked
        Missing_Item = "R"
        response.write "<TD BGCOLOR=""#AAAAFF"" ALIGN=""CENTER"" CLASS=Small>"      
      
      elseif ((LCase(rsUser("Lit_Status")) = "active" and LCase(rsUser("Lit_Action")) = "complete" and LCase(rsUser("Lit_Status_Name")) = "final loaded") or _
              (LCase(rsUser("Lit_Status")) = "active" and LCase(rsUser("Lit_Action")) = "complete" and LCase(rsUser("Lit_Status_Name")) = "reprint")) then

         if instr(1,rsUser("SubGroups"),"shpcrt") > 0 then
           response.write "<TD BGCOLOR=""#FF8000"" ALIGN=""CENTER"" CLASS=Small>"      
           Missing_Item = "E"
         else
           Missing_Item = "Y"
           response.write "<TD BGCOLOR=""#90EE90"" ALIGN=""CENTER"" CLASS=Small>"      
         end if  

      elseif LCase(rsUser("Lit_Status")) = "active" and LCase(rsUser("Lit_Action")) = "complete" and LCase(rsUser("Lit_Status_Name")) = "web" then
        Missing_Item = "P"
        response.write "<TD BGCOLOR=""#F5DEB3"" ALIGN=""CENTER"" CLASS=Small>"

      ' Unknown Status
      
      elseif LCase(rsUser("Lit_Status")) = "active" and LCase(rsUser("Lit_Action")) <> "complete" and _
        LCase(rsUser("Lit_Action")) <> "n/a" then

        Missing_Item = "N"
        response.write "<TD BGCOLOR=""#0099FF"" ALIGN=""CENTER"" CLASS=Small>"

      elseif not isblank(rsUser("Item_Number")) and not isblank(rsUser("Revision_Code")) and isblank("Lit_Item") then 
        Missing_Item = "?"
        response.write "<TD BGCOLOR=""#FFFF00"" ALIGN=""CENTER"" CLASS=Small>"

      else
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      end if
      
      if Missing_Item <> "Y" and Missing_Item <> "" then
        response.write "<A HREF=""JavaScript:void(0);"" ONCLICK=""var MyPop2 = window.open('/sw-administrator/Calendar_Element_Names.asp','MyPop2','fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,directories=no,location=no,width=250,height=300,left=600,top=200'); MyPop2.focus(); return false;"" CLASS=Small>"
        response.write Missing_Item
        response.write "</A>"
      else
        response.write Missing_Item
      end if

      response.write "</TD>"    

      if Utility_ID = 50 or Utility_ID = 54 or Utility_ID = 60 then
        ' Groups Allowed to View
        if View = 1 or View = 3 then
          response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
          if Group_ID <> "" then
            response.write Highlight_Keyword(replace(rsUser("SubGroups"),"view, ",""),Group_ID, "#FF0000")          
          else
            response.write replace(rsUser("SubGroups"),"view, ","")
          end if  
          response.write "</TD>"
        end if  
        if View = 2 or View = 3 then  
          response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
          if rsUser("Country") <> "none" then
            if instr(1,rsUser("Country"),"0, ") = 1 then
              response.write "<SPAN CLASS=SmallRed>" & Translate("Exclude",Login_Language,conn) & ":</SPAN> " & mid(rsUser("Country"),3)
            else
              response.write "<SPAN CLASS=SmallRed>" & Translate("Limit to",Login_Language,conn) & ":</SPAN> " & rsUser("Country")  
            end if
          else
            response.write "&nbsp;"  
          end if
          response.write "</TD>"
        end if  
      end if  
      
      if Utility_ID = 50 or Utility_ID = 51 or Utility_ID = 52 or Utility_ID = 54 or Utility_ID = 60 then
      
        ' Item Reference Number
        with response
          .write "<TD BGCOLOR="""
          if Missing_Item = "?" then
            .write "#FFFF00"
          else
            .write "#FFFFFF"
          end if
          .write """ ALIGN=""LEFT"" CLASS=Small>"
          .write rsUser("Item_Number")
          .write "</TD>"
        end with

        ' Item Reference Revision Code
        with response
          .write "<TD BGCOLOR="""
          if Missing_Item = "?" then
            .write "#FFFF00"
          else
            .write "#FFFFFF"
          end if
          .write """ ALIGN=""LEFT"" CLASS=Small>"
          .write rsUser("Revision_Code")
          .write "</TD>"
        end with

      end if

      if Utility_ID = 51 or Utility_ID = 52 or Utility_ID = 60 then      
        with response
          .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
          .write rsUser("Item_Number_2")
          .write "</TD>"          
        end with
      end if        

      ' Language
      if UCase(rsUser("Language")) = "ENG" then
        response.write "<TD BGCOLOR=""#EFEFEF"" ALIGN=""LEFT"" CLASS=Small>"
      else  
        response.write "<TD BGCOLOR=""#CFCFCF"" ALIGN=""LEFT"" CLASS=Small>"
      end if
      response.write UCase(rsUser("Language"))
      response.write "</TD>"

      ' Pre-Announce Date
      response.write "<TD BGCOLOR="""
      if CDate(rsUser("LDate")) < CDate(rsUser("BDate")) then
        if (CDate(rsUser("LDate")) > Date) then
          response.write "Yellow"
        elseif (CDate(rsUser("EDate")) < CDate(rsUser("XDate"))) and (CDate(rsUser("XDate")) >= Date) then
          if Status > 0 then
            response.write "#00CC00"
          else
            response.write "Orange"
          end if    
        elseif (CDate(rsUser("EDate")) < CDate(rsUser("XDate"))) and (CDate(rsUser("XDate")) < Date) then
          response.write "#AAAAFF"
        elseif CDate(rsUser("BDate")) <> CDate(rsUser("EDate")) and CDate(rsUser("XDate")) < Date then
          response.write "#AAAAFF"        
        elseif (CDate(rsUser("EDate")) = CDate(rsUser("XDate"))) and (CDate(rsUser("LDate")) <= Date) then
          if Status > 0 then
            response.write "#00CC00"
          else
            response.write "Orange"
          end if    
        else
          response.write "#FFFFFF"
        end if
        response.write """ ALIGN=""CENTER"" CLASS=Small>"
        response.write Replace(FormatDate(0,rsUser("LDate")),"/","&nbsp;&nbsp;")
      else
        response.write "#FFFFFF"
        response.write """ ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      end if    
      response.write "</TD>"

      ' Begin Date
      
      response.write "<TD BGCOLOR="""
      if (CDate(rsUser("LDate")) > Date) and (CDate(rsUser("BDate")) > Date) then
        response.write "Yellow"
      elseif (CDate(rsUser("EDate")) < CDate(rsUser("XDate"))) and (CDate(rsUser("XDate")) >= Date) then
        if Status > 0 then
          response.write "#00CC00"
        else
          response.write "Orange"
        end if    
      elseif (CDate(rsUser("EDate")) < CDate(rsUser("XDate"))) and (CDate(rsUser("XDate")) < Date) then
        response.write "#AAAAFF"
      elseif CDate(rsUser("BDate")) <> CDate(rsUser("EDate")) and CDate(rsUser("XDate")) < Date then
        response.write "#AAAAFF"        
      elseif (CDate(rsUser("EDate")) = CDate(rsUser("XDate"))) then
        if Status > 0 then
          response.write "#00CC00"
        else
          response.write "Orange"
        end if    
      else
        response.write "#FFFFFF"
      end if
      response.write """ ALIGN=""CENTER"" CLASS=Small>"
      response.write Replace(FormatDate(0,rsUser("BDate")),"/","&nbsp;&nbsp;")
      response.write "</TD>"
      
      ' Subscription Service Time
      
      if Utility_ID = 54 then
        response.write "<TD BGCOLOR="""
        if CInt(rsUser("Subscription_Early")) = CInt(True) then
          if CDate(rsUser("LDate") & " " & "12:05:00 PM") > CDate(Now()) then
            response.write "Yellow"
          else
            response.write "#00CC00"
          end if
        else
          if CDate(rsUser("LDate") & " " & "09:05:00 PM") > CDate(Now()) then
            response.write "Yellow"
          else
            response.write "#00CC00"
          end if
        end if
        response.write """ ALIGN=""CENTER"" CLASS=Small>"
        if CInt(rsUser("Subscription_Early")) = CInt(True) then
          response.write "12:05 PM PST"
        else
          response.write "09:05 PM PST"
        end if
        response.write "</TD>"
      else
      
      ' End Date      
      
        response.write "<TD BGCOLOR="""
        if CDate(rsUser("BDate")) < CDate(rsUser("EDate")) then
          if (CDate(rsUser("LDate")) > Date) and (CDate(rsUser("EDate")) > Date) then
            response.write "Yellow"
          elseif (CDate(rsUser("EDate")) < CDate(rsUser("XDate"))) and (CDate(rsUser("XDate")) >= Date) then
            if Status > 0 then
              response.write "#00CC00"
            else
              response.write "Orange"
            end if    
          elseif (CDate(rsUser("EDate")) < CDate(rsUser("XDate"))) and (CDate(rsUser("XDate")) < Date) then
            response.write "#AAAAFF"
          elseif CDate(rsUser("BDate")) <> CDate(rsUser("EDate")) and CDate(rsUser("XDate")) < Date then
            response.write "#AAAAFF"        
          elseif (CDate(rsUser("EDate")) = CDate(rsUser("XDate"))) then
            if Status > 0 then
              response.write "#00CC00"
            else
              response.write "Orange"
            end if    
          else
            response.write "#FFFFFF"
          end if
          response.write """ ALIGN=""CENTER"" CLASS=Small>"
          response.write Replace(FormatDate(0,rsUser("EDate")),"/","&nbsp;&nbsp;")
        else  
          response.write "#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        end if      
        response.write "</TD>"
      end if
      
      ' Public Embargo Date
      response.write "<TD BGCOLOR="""
      if isdate(rsUser("PEDate")) then
        if (CDate(rsUser("PEDate")) >= CDate(Date)) then
          if (CDate(rsUser("PEDate")) >= CDate(Date)) then
            response.write "Yellow"
          elseif (CDate(rsUser("PEDate")) < CDate(rsUser("XDate"))) and (CDate(rsUser("XDate")) >= CDate(Date)) then
            if Status > 0 then
              response.write "#00CC00"
            else
              response.write "Orange"
            end if    
          elseif (CDate(rsUser("PEDate")) < CDate(rsUser("XDate"))) and (CDate(rsUser("XDate")) < Date) then
            response.write "#AAAAFF"
          elseif (CDate(rsUser("PEDate")) = CDate(rsUser("XDate"))) then
            if Status > 0 then
              response.write "#00CC00"
            else
              response.write "Orange"
            end if    
          else
            response.write "#FFFFFF"
          end if
          response.write """ ALIGN=""CENTER"" CLASS=Small>"
          response.write Replace(FormatDate(0,rsUser("PEDate")),"/","&nbsp;&nbsp;")
        else  
          response.write "#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        end if      
      else  
        response.write "#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
        response.write "&nbsp;"
      end if
      response.write "</TD>"

      ' Expiration Date
      response.write "<TD BGCOLOR="""
      
      if CDate(rsUser("BDate")) = CDate(rsUser("EDate")) and CDate(rsUser("EDate")) = CDate(rsUser("XDate")) then
        response.write "White"" ALIGN=""CENTER"" CLASS=Small>"      
        response.write "Never"
      elseif CDate(rsUser("EDate")) < CDate(rsUser("XDate")) and CDate(rsUser("XDate")) > Date then
        if Status > 0 then
          response.write "#00CC00"
        else
          response.write "Orange"
        end if    
        response.write """ ALIGN=""CENTER"" CLASS=Small>"      
        response.write Replace(FormatDate(0,rsUser("XDate")),"/","&nbsp;&nbsp;")
      elseif CDate(rsUser("EDate")) < CDate(rsUser("XDate")) and CDate(rsUser("XDate")) <= Date then
        response.write "#AAAAFF"" ALIGN=""CENTER"" CLASS=Small>"      
        response.write Replace(FormatDate(0,rsUser("XDate")),"/","&nbsp;&nbsp;")
      elseif CDate(rsUser("EDate")) = CDate(rsUser("XDate")) and CDate(rsUser("XDate")) <= Date then
        response.write "#AAAAFF"" ALIGN=""CENTER"" CLASS=Small>"      
        response.write Replace(FormatDate(0,rsUser("XDate")),"/","&nbsp;&nbsp;")
      else
        response.write "WHITE"" ALIGN=""CENTER"" CLASS=Small>&nbsp;"
      end if
      response.write "</TD>"
      response.write "</TR>"

      rsUser.MoveNext
    
    loop
                         
    rsUser.close
    set rsUser=nothing
  
    if TableOn then
      response.write "</TABLE>"
      Call Table_End
    end if
        
  ' --------------------------------------------------------------------------------------
  ' Asset Activity
  ' --------------------------------------------------------------------------------------
  
  elseif Utility_ID = 70 and Admin_Access <> 3 then

    TableOn = False
    
    if Admin_Access <> 1 then
      Call Nav_Border_Begin
      Call Main_Menu
    else
      Call Nav_Border_Begin    
      Call Metrics_Menu
    end if

    Call Activity_Methods
    Call Nav_Border_End

    response.write "<FORM NAME=""Activity"" METHOD=""POST"" ACTION=""Site_Utility.asp"">" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""ID"" VALUE=""site_utility"">" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Utility_ID"" VALUE=""" & Utility_ID & """>" & vbCrLf

    ' Search for Item Numbers (Not Null will exclude certain other filters)
    if not isblank(request.form("Item_Numbers")) then
      item_numbers = replace(request.form("Item_Numbers")," ","")
    else
      item_numbers = ""
    end if

    Call Table_Begin
    response.write "<SPAN CLASS=SmallBoldWhite>" & Translate("Begin Date",Login_Language,conn) & ": "
    if isblank(request.form("Begin_Date")) then
      Begin_Date = Date()
    elseif isdate(request.form("Begin_Date")) then
      Begin_Date = request.form("Begin_Date")
    else
      response.write Translate("Invalid Date - Reseting to Today's Date",Login_Language,conn) & "<P>" & vbCrLf
      Begin_Date = Date()      
    end if
    response.write "</SPAN>" & vbCrLf
    response.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""Begin_Date"" VALUE=""" & Begin_Date & """ SIZE=""6"">" & vbCrLf
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Span",Login_Langugage,conn) & ":</SPAN> " & vbCrLf

    if isblank(request.form("Interval")) then
      Interval = 1
    elseif isnumeric(request.form("Interval")) then
      if FormatDateTime(Begin_Date) = FormatDateTime(Date()) and request.form("Interval") >= 0 then
        Interval = 1
      elseif request.form("Interval") > 90 then
        Interval = 90
      elseif request.form("Interval") < -90 then
        Interval = -90
      else  
        Interval = request.form("Interval")
      end if  
    end if

    ' Interval in Days
    response.write "<SELECT CLASS=Small NAME=""Interval"">"

    for x = 90 to -90 step -1
      select case x
        case 90,60,30,14,7          ' Future Days
          response.write "<OPTION VALUE=""" & x & """"
          if CInt(Interval) = x then response.write " SELECTED"
          response.write ">+" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
        case 1
          response.write "<OPTION CLASS=Region5NavSmall VALUE=""" & x & """"
          if CInt(Interval) = 1 then response.write " SELECTED"
          response.write ">+1" & " " & Translate("Day",Login_Language,conn) & "</OPTION>" & vbCrLf                
        case -90,-60,-30,-14,-7     ' Past Days
          response.write "<OPTION VALUE=""" & x & """"
          if CInt(Interval) = x then response.write " SELECTED"
          response.write ">-" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
      end select
    next

    response.write "</SELECT>" & vbCrLf

    ' Order Inquiry Filter

    if not isblank(request.form("OI")) and CInt(request.form("OI")) = CInt(True) then
      Order_Inquiry = CInt(True)
    else
      Order_Inquiry = CInt(False)
    end if
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Order Inquiry",Login_Language,conn) & ": </SPAN>"
    response.write "<SELECT CLASS=Small NAME=""OI"">"
    response.write "<OPTION VALUE=""0"""
    if CInt(Order_Inquiry) = 0 then response.write " SELECTED"
    response.write ">" & Translate("Excluded",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""-1"""
    if CInt(Order_Inquiry) = -1 then response.write " SELECTED"
    response.write ">" & Translate("Included",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "</SELECT>" & vbCrLf

    ' Change Region / Site

    Call Change_Region
    if Admin_Access = 1 then
      Site_ID_Change = 0
    else
      Call Change_Site
    end if  

    ' Submit Button
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE="" " & Translate("GO",Login_Language,conn) & " "">" & vbCrLf

    ' Reset Button
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""RESET"" VALUE="" " & Translate("Reset",Login_Language,conn) & " "" LANGUAGE=""Javascript"" ONCLICK=""this.form.Begin_Date.value='" & Date() & "';this.form.Interval.options[0].selected=true;this.form.OI.options[0].selected=true;this.form.Region.options[0].selected=true;this.form.Item_Numbers.value='';"">" & vbCrLf
    
    response.write "<BR>"
    response.write "<SPAN CLASS=SmallBoldWhite>" & Translate("Asset ID or Item Number",Login_Language,conn) & ": "
    response.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""Item_Numbers"" VALUE=""" & item_numbers & """ SIZE=""30"">&nbsp;&nbsp;" & vbCrLf
    response.write "<SPAN CLASS=SmallWhite>(" & Translate("Separate multiple Asset ID and Item Numbers with a comma.",Login_Language,conn) & ")</SPAN>"
    Call Table_End
    response.write "</FORM>"
    
    ' Filter Specific WHERE clause
    
    if Interval >= 0 and isblank(Item_Numbers) then
      SQLWhere    = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Activity.View_Time, 102), 102) <= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102))) "
      SQLWhereLOS = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date, 102), 102) <= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102))) "      
    elseif not isblank(Item_Numbers) then
      SQLWhere    = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "
      SQLWhereLOS = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "              
    else
      SQLWhere    = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102)) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Activity.View_Time,102), 102) <= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "
      SQLWhereLOS = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102)) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) <= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "      
    end if

    if Site_ID_Change > 0 then
      SQLWhere = SQLWhere & "AND dbo.Activity.Site_ID=" & Site_ID_Change & " "
    else
      SQLWhere = SQLWhere & " "
    end if  
      
    select case Region
      case 1, 2, 3
        SQLWhere = SQLWhere & "AND dbo.UserData.Region=" & Region & " "
    end select
   
    ' Site Clicks
   
    if isblank(Item_Numbers) then
      SQL = "SELECT  COUNT(dbo.Activity.View_Time) AS Clicks " &_
            "FROM         dbo.Activity LEFT OUTER JOIN " &_
            "             dbo.UserData ON dbo.Activity.Account_ID = dbo.UserData.ID "
    
      SQL = SQL & SQLWhere
    
      if CInt(Order_Inquiry) = CInt(False) then
        SQL = SQL & "AND dbo.Activity.Calendar_ID <> 101 AND dbo.Activity.Calendar_ID <> 102 "
      end if  

      Set rsActivity = Server.CreateObject("ADODB.Recordset")
      rsActivity.Open SQL, conn, 3, 3
    
      Site_Clicks = rsActivity("Clicks")
      rsActivity.close
      set rsActivity = nothing
    
      ' Unique Visitors
    
      Dim Site_Visitors
    
      SQL = "SELECT DISTINCT dbo.Activity.Account_ID " &_
              "FROM         dbo.Activity LEFT OUTER JOIN " &_
            "             dbo.UserData ON dbo.Activity.Account_ID = dbo.UserData.ID "
      SQL = SQL & SQLWhere
      Set rsActivity = Server.CreateObject("ADODB.Recordset")
      rsActivity.Open SQL, conn, 3, 3

      Site_Visitors = rsActivity.RecordCount
      rsActivity.close
      set rsActivity = nothing
      
    end if   
    
    ' Asset Activity    
     
    SQL = "SELECT    dbo.Calendar.Product AS Product, Calendar.Item_Number AS Item_Number, Calendar.Revision_Code, Calendar.Item_Number_2 AS Item_Number_2, dbo.Calendar.Title AS Title, dbo.Activity.Calendar_ID AS Asset_ID, dbo.Activity.Method AS Method, dbo.Activity.CID AS CID, dbo.Activity.Account_ID AS Account_ID, " &_
          "          dbo.Activity.View_Time AS View_Time, dbo.Activity.Site_ID AS Site_ID, dbo.Activity.[Language] AS Language, dbo.Activity.CMS_Site AS CMS_Site, " &_
          "          dbo.Calendar_Category.Title AS Category_Title, dbo.Content_Sub_Category.Sub_Category AS Sub_Category_Title, dbo.UserData.Region " &_
          "FROM      dbo.UserData RIGHT OUTER JOIN " &_
          "          dbo.Activity ON dbo.UserData.ID = dbo.Activity.Account_ID AND dbo.UserData.Site_ID = dbo.Activity.Site_ID LEFT OUTER JOIN " &_
          "          dbo.Content_Sub_Category RIGHT OUTER JOIN " &_
          "          dbo.Calendar ON dbo.Content_Sub_Category.Sub_Category = dbo.Calendar.Sub_Category AND " &_
          "          dbo.Content_Sub_Category.Site_ID = dbo.Calendar.Site_ID ON dbo.Activity.Calendar_ID = dbo.Calendar.ID LEFT OUTER JOIN " &_
          "          dbo.Calendar_Category ON dbo.Calendar.Code = dbo.Calendar_Category.Code AND dbo.Calendar.Site_ID = dbo.Calendar_Category.Site_ID "
      
    SQL = SQL & SQLWhere & "AND (dbo.Activity.Calendar_ID > 0) "
      
    item_numbers_max = 0
    if not isblank(item_numbers) then
      if instr(item_numbers,",") > 0 then
        item_number = Split(item_numbers,",")
        item_number_max = Ubound(item_number)
      else
        Dim item_number(0)
        item_number(0) = item_numbers
      end if

      for x = 0 to item_number_max
        if x = 0 then
          SQL = SQL & "AND ("
        else
          SQL = SQL & "OR "  
        end if
        if len(item_number(x)) < 6 then
          SQL = SQL & "dbo.Activity.Calendar_ID='" & item_number(x) & "' "
        else  
          SQL = SQL & "dbo.Calendar.Item_Number='" & item_number(x) & "' "
        end if  
      next
      SQL = SQL & ") "
    end if    

    SQL = SQL & "ORDER BY dbo.Calendar_Category.Title, dbo.Content_Sub_Category.Sub_Category, dbo.Activity.Calendar_ID"
        
    Set rsActivity = Server.CreateObject("ADODB.Recordset")
'response.write SQL & "<P>"
'response.flush
    rsActivity.Open SQL, conn, 3, 3
    
    ' Count Order Inquiry Assets to subtract from total if excluded
    ' Since OI Items do not have a Title, they are at the beginning of the RS, then sorted by Calendar_ID (101 & 102)

    Dim OI_Count
    OI_Count = 0
    Call Bypass_Assets()
    
    if isblank(Item_Numbers) then
      response.write "<SPAN CLASS=SmallBold>"
      response.write Translate("Unique Visitors",Login_Language,conn) & ": " & FormatNumber(Site_Visitors,0) & "&nbsp;&nbsp;&nbsp;"
      response.write Translate("Navigation Clicks",Login_Language,conn) & ": " & FormatNumber(Site_Clicks,0) & "&nbsp;&nbsp;&nbsp;"
      response.write Translate("Asset Activity",Login_Language,conn) & ": " & FormatNumber((rsActivity.RecordCount - OI_Count),0) & "&nbsp;&nbsp;&nbsp;"
      response.write "</SPAN><BR>"
    end if  
    
    if not rsActivity.EOF then
     
      %>
     	<!--TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
        <TR>
          <TD-->
            <%Call Table_Begin%>
            <TABLE CELLPADDING=2 CELLSPACING=1 BORDER=0  WIDTH="100%">
              <TR>
                <TD BGCOLOR="#000000" ALIGN="RIGHT" CLASS=SmallBoldGold><%=Translate("Asset ID",Login_Language,conn)%></TD>              
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold><%=Translate("Category",Login_Language,conn)%></TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold><%=Translate("Sub Category",Login_Language,conn)%></TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold><%=Translate("Product or Product Series",Login_Language,conn)%></TD>                
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold><%=Translate("Title",Login_Language,conn)%></TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold><%=Translate("Item Number",Login_Language,conn)%></TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold><%=Translate("LAN",Login_Language,conn)%></TD>
                
                <%
                '  xOLView         = 0  ' On-Line View (Default)
                '  xOLDownLoad     = 1  ' On-Line Download
                '  xOLSend         = 2  ' On-Line Send
                '  xSSView         = 3  ' Subscription Service View
                '  xSSDownload     = 4  ' Subscription Service Download
                '  xSSSend         = 5  ' Subscription Service Send
                '  xOLSendNoZip    = 6  ' On-Line Send Non-Zip File Version
                '  xOLLink         = 7  ' On-Line Link
                '  xOLLinkNoPop    = 8  ' On-Line Link No Pop-Up
                '  xOLGateway      = 9  ' On-Line Gateway to Site
                '  xOLGatewayNoPop = 10 ' On-Line Gateway to Site No Pop-Up
                '  xOLViewPOD      = 11 ' On-Line Download Print on Demand Doccument
                '  xOLDownLoadPOD  = 12 ' On-Line Download Print on Demand Doccument
                '                    13 ' WWW.Fluke.com via Find_It.asp
                %>
  
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>OLV</TD>            
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>OLD</TD>            
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>OLS</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>SSV</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>SSD</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>SSS</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>OLL</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>OLG</TD>              
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>EEF</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>EDL</TD>                
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>LOS</TD>                                
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold><%=Translate("Total",Login_Language,conn)%></TD>              
          		</TR>
        <%

      Dim Method(15)                ' Total items for Method per Individual Asset ID
      Dim Category(15)              ' Total items for Category
      Dim Period(15)                ' Total items for Reporting Period
      
      for x = 0 to 15       ' Initialize Method Category Counters
        Method(x)   = 0
        Category(x) = 0
        Period(x)   = 0
      next  
   
      Old_Asset_ID = rsActivity("Asset_ID")
      
      OI_Select = Translate("Order Inquiry",Login_Language,conn) & " - " & Translate("PO Select",Login_Language,conn)
      OI_Result = Translate("Order Inquiry",Login_Language,conn) & " - " & Translate("PO Result",Login_Language,conn)
      Retired   = Translate("Asset Retired",Login_Language,conn)
      
      select case Old_Asset_ID
        case 101
          Old_Category = OI_Select
          New_Category = OI_Select
        case 102
          Old_Category = OI_Result
          New_Category = OI_Result
        case else
          Old_Category = Translate(rsActivity("Category_Title"),Login_Language,conn)
          New_Category = Translate(rsActivity("Category_Title"),Login_Language,conn)
      end select
      
      if isblank(Old_Category) then Old_Category = Retired
      if isblank(New_Category) then New_Category = Retired
                 
      do while not rsActivity.EOF

        Call Bypass_Assets
  
        if CInt(Old_Asset_ID) = CInt(rsActivity("Asset_ID")) then

          Call Update_Methods

          Old_Asset_ID = rsActivity("Asset_ID")
          
          select case Old_Asset_ID
            case 101
              Old_Category = OI_Select
            case 102
              Old_Category = OI_Result
            case else
              Old_Category = Translate(rsActivity("Category_Title"),Login_Language,conn)
          end select
          
          if isblank(Old_Category) then Old_Category = Retired
          
        else

          rsActivity.MovePrevious
          
          response.write "<TR>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">" & rsActivity("Asset_ID") & "</TD>" & vbCrLf
          if not isblank(New_Category) then
            response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"">" & New_Category & "</TD>" & vbCrLf
            response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"">" & Translate(rsActivity("Sub_Category_Title"),Login_Language,conn) & "</TD>" & vbCrLf
            response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"">" & rsActivity("Product") & "</TD>" & vbCrLf
            response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"">" & rsActivity("Title") & "</TD>" & vbCrLf
          else  
            response.write "<TD CLASS=Small BGCOLOR=""LightSteelBlue"" COLSPAN=5>" & Retired & "</TD>" & vbCrLf
          end if  

          if not isblank(rsActivity("Item_Number")) then
            response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">"
            response.write rsActivity("Item_Number")
            if not isblank(rsActivity("Revision_Code")) then
              response.write " " & rsActivity("Revision_Code")
            end if

            if instr(1,Unique_Item_Numbers,rsActivity("Item_Number") & ",") = 0 then
  
              SQLLOS = "SELECT  Item_Number AS Item_Number, Quantity AS Quantity " &_
                       "FROM    dbo.Shopping_Cart_Lit "
              SQLLOS = SQLLOS & SQLWhereLOS & " AND Item_Number = '" & rsActivity("Item_Number") & "' "
  
              Set rsLOS = Server.CreateObject("ADODB.Recordset")
              rsLOS.Open SQLLOS, conn, 3, 3
            
              do while not rsLOS.EOF
                Method(14) = Method(14) + rsLOS("Quantity")
                Method(15) = Method(15) + rsLOS("Quantity")
                rsLOS.MoveNext
              loop
            
              rsLOS.close
              set rsLOS = nothing
              
              Unique_Item_Numbers = Unique_Item_Numbers & rsActivity("Item_Number") & ","
              
            end if  
            
            
          elseif not isblank(rsActivity("Item_Number_2")) then
            response.write "<TD CLASS=Small BGCOLOR=""#F1F1F1"" ALIGN=""CENTER"">"
            response.write rsActivity("Item_Number_2")
          else
            response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">"
            response.write "&nbsp;"
          end if
          response.write "</TD>" & vbCrLf        

          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & UCase(rsActivity("Language")) & "</TD>" & vbCrLf        

          Call Display_Methods
  
          response.write "</TR>" & vbCrLf & vbCrLf

          select case rsActivity("Asset_ID")
            case 101
              Old_Category = OI_Select
            case 102
              Old_Category = OI_Result
            case else
              Old_Category = Translate(rsActivity("Category_Title"),Login_Language,conn)
          end select

          if isblank(Old_Category) then Old_Category = Retired
          
          rsActivity.MoveNext

          select case rsActivity("Asset_ID")
            case 101
              New_Category = OI_Select
            case 102
              New_Category = OI_Result
            case else
              New_Category = Translate(rsActivity("Category_Title"),Login_Language,conn)
          end select

          if isblank(New_Category) then New_Category = Retired
  
          Call Update_Methods
  
          Old_Asset_ID = rsActivity("Asset_ID")

        end if
        
        ' Display Category Totals on Category Change
        
        if LCase(Old_Category) <> LCase(New_Category) then

          response.write "<TR>" & vbCrLf
          response.write "<TD CLASS=SmallBold BGCOLOR=""LightGrey"" COLSPAN=7 ALIGN=RIGHT>" & Translate("Category Totals",Login_Language,conn) & " : </TD>" & vbCrLf
          Call Display_Category  
          response.write "</TR>" & vbCrLf & vbCrLf

        end if
  
        rsActivity.MoveNext
              
      loop

      if method(15) > 0 then

        rsActivity.MovePrevious
        
        response.write "<TR>" & vbCrLf
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">" & rsActivity("Asset_ID") & "</TD>" & vbCrLf
        if not isblank(New_Category) then
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"">" & New_Category & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"">" & Translate(rsActivity("Sub_Category_Title"),Login_Language,conn) & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"">" & rsActivity("Product") & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"">" & rsActivity("Title") & "</TD>" & vbCrLf
        else  
          response.write "<TD CLASS=Small BGCOLOR=""LightSteelBlue"" COLSPAN=5>" & Retired & "</TD>" & vbCrLf
        end if
          
        if not isblank(rsActivity("Item_Number")) then
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">"
          response.write rsActivity("Item_Number")
          if not isblank(rsActivity("Revision_Code")) then
            response.write " " & rsActivity("Revision_Code")
          end if
        elseif not isblank(rsActivity("Item_Number_2")) then
          response.write "<TD CLASS=Small BGCOLOR=""#F1F1F1"" ALIGN=""CENTER"">"
          response.write rsActivity("Item_Number_2")
        else
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">"
          response.write "&nbsp;"
        end if
        response.write "</TD>" & vbCrLf        

        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & UCase(rsActivity("Language")) & "</TD>" & vbCrLf        
  
        Call Display_Methods

        response.write "</TR>" & vbCrLf & vbCrLf
        
      end if  

      ' Display Last Category Set
      
      if Category(15) > 0 then
        response.write "<TR>" & vbCrLf
        response.write "<TD CLASS=SmallBold BGCOLOR=""LightGrey"" COLSPAN=7 ALIGN=RIGHT>" & Translate("Category Totals",Login_Language,conn) & " : </TD>" & vbCrLf
        Call Display_Category
        response.write "</TR>" & vbCrLf & vbCrLf
      end if  

      ' Display Period Totals
      
      response.write "<TR>" & vbCrLf
      response.write "<TD CLASS=SMALLBOLDWHITE BGCOLOR=""SteelBlue"" COLSPAN=4 ALIGN=RIGHT>&nbsp;</TD>" & vbCrLf
      response.write "<TD CLASS=SMALLBOLD BGCOLOR=""#FFFFFF"" COLSPAN=3 ALIGN=RIGHT>" & Translate("Period Totals",Login_Language,conn) & " : </TD>" & vbCrLf

      for x = 0 to 15
        select case x
          case 0,1,2,3,4,5,7,9,11,13,14,15       ' Display only certain fields
            response.write "<TD CLASS=SMALLBOLDWHITE ALIGN=RIGHT BGCOLOR=""SteelBlue"">"
            if Period(x) = 0 then
              response.write "&nbsp;"
            else
              response.write FormatNumber(Period(x),0)
            end if
            response.write "</TD>" & vbCrLf
        end select
        Period(x) = 0
      next

      response.write "</TR>" & vbCrLf & vbCrLf

      rsActivity.close
      set rsActivity = nothing    
    
      response.write "</TABLE>" & vbCrLf & vbCrLf
      Call Table_End
      
    else

      response.write "<SPAN CLASS=SmallBold>" & Translate("There are 0 records that meet the filter criteria that you have specified.",Login_Language,conn) & "<SPAN><P>"

    end if
    
  ' --------------------------------------------------------------------------------------
  ' List - Asset Activity Summary
  ' --------------------------------------------------------------------------------------

  elseif Utility_ID = 71 and Admin_Access <> 3 then
  
    server.scripttimeout = 360
  
    Dim Summary_Class
    Summary_Class = 19
    Dim Summary(19,13)
    Dim Summary_Title(19)
    Dim Summary_Key(19)

    response.write "<FORM NAME=""Activity"" METHOD=""POST"" ACTION=""Site_Utility.asp"">" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""ID"" VALUE=""site_utility"">" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Utility_ID"" VALUE=""" & Utility_ID & """>" & vbCrLf
    if Admin_Access <> 1 then
      Call Nav_Border_Begin
      Call Main_Menu
      Call Change_Site
    else
      Call Nav_Border_Begin
      Call Metrics_Menu
      Site_ID_Change = 0
    end if  

    if not isblank(request.form("Count_Year")) and isnumeric(request.form("Count_Year")) then
      Count_Year = request.form("Count_Year")
    else
      Count_Year = 0
    end if
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Year Span",Login_Language,conn) & ":&nbsp;<INPUT TYPE=""RADIO"" NAME=""Count_Year"" VALUE=""0"""
    if Count_Year = 0 then response.write " CHECKED"
    response.write ">&nbsp;" & Year(Date()) & "&nbsp;&nbsp;" & vbCrLf
    response.write "<INPUT TYPE=""RADIO"" NAME=""Count_Year"" VALUE=""1"""
    if Count_Year = 1 then response.write " CHECKED"    
    response.write ">&nbsp;" & Year(Date()) & " + " & Year(Date()) -1 & "</SPAN>" & vbCrLf
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Region",Login_Language,conn) & ":&nbsp;"
    response.write "<SELECT NAME=""Region"" CLASS=Small>" & vbCrLf
    response.write "<OPTION VALUE=""0"""
    if isblank(request("Region")) or request("Region") = "0" then response.write " SELECTED"
    response.write ">" & Translate("All",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""1"""
    if request("Region") = "1" then response.write " SELECTED"
    response.write ">" & Translate("United States",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""2"""
    if request("Region") = "2" then response.write " SELECTED"
    response.write ">" & Translate("Europe",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""3"""
    if request("Region") = "3" then response.write " SELECTED"
    response.write ">" & Translate("Intercon",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "</SELECT>" & vbCrLf
    
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE="" " & Translate("GO",Login_Language,conn) & " "">" & vbCrLf
    Call Nav_Border_End
    response.write "</FORM>"

    for Summary_Year = Year(Date()) to Year(Date()) - Count_Year step -1

      for x = 0 to Summary_Class
        for y = 0 to 13       ' Do not change
          Summary(x,y) = 0
        next
      next
    
      for Summary_Month = 1 to 12        

          SQL = "SELECT "
          
          for z = 0 to 7
    
            SQL = SQL & "Count_" & z & " = (SELECT Count(Calendar_ID) FROM Activity WHERE "
            if Site_ID_Change > 0 then
              SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
            end if

            SQL = SQL & "(DATEPART(m, dbo.Activity.View_Time) = " & Summary_Month & ") AND (DATEPART(yyyy, dbo.Activity.View_Time) = " & Summary_Year & ") "
            
            select case z
              case 0      ' Order Query Criteria
                SQL = SQL & " AND Calendar_ID=101), "
              case 1      ' Order Inquiry Results
                SQL = SQL & " AND Calendar_ID=102), "
              case 2      ' Assets Accessed
                SQL = SQL & " AND Calendar_ID>200), "
              case 3      ' EDF  
                SQL = SQL & " AND Account_ID=1 AND CMS_Site IS NULL), "
              case 4      ' WWW  
                SQL = SQL & " AND Account_ID=1 AND CMS_Site IS NOT NULL), "
              case 5      ' Search
                SQL = SQL & " AND CID=9004), "
              case 6      ' Navigation clicks
                SQL = SQL & " AND Calendar_ID=0), "
              case 7      ' All Clicks
                SQL = SQL & ")"
            end select
          next

response.write SQL & "<P>"

          Set rsActivity = conn.Execute(SQL)
        
          for z = 0 to 7        ' Activity Data
            Summary(z,Summary_Month) = rsActivity("Count_" & z)
            Summary(z,13) = Summary(z,13) + Summary(z,Summary_Month)                ' Year Totals for Class
          next
          
          for z = 8 to 16       ' Account Data

            Multiplier = 1
            
            select case z
            
              case 8            ' New Registrations
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                SQL = SQL & "(DATEPART(m, UserData.Reg_Request_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, UserData.Reg_Request_Date) = " & Summary_Year & ") "

              case 9            ' New Registrations Pending
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CInt(True) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                SQL = SQL & "(DATEPART(m, UserData.Reg_Request_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, UserData.Reg_Request_Date) = " & Summary_Year & ") "

              case 10            ' New Accounts
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CInt(False) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                SQL = SQL & "(DATEPART(m, UserData.Reg_Approval_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, UserData.Reg_Approval_Date) = " & Summary_Year & ") "
                
              case 11            ' Expired Accounts
                Multiplier = -1
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CInt(False) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                if Summary_Month + 1 < 13 then
                  Last_Day_Month = DateAdd("d",-1,(Summary_Month + 1) & "/1/" & Summary_Year)
                else
                  Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
                end if  
                SQL = SQL & " (Reg_Approval_Date >= '1/1/2001' AND Reg_Approval_Date <= '" & Last_Day_Month & "') AND (ExpirationDate >= '" & Summary_Month & "/1/" & Summary_Year & "' AND ExpirationDate <= '" & Last_Day_Month & "') "

              case 12            ' Never Logon Accounts
                Multiplier = -1
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CInt(False) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                if Summary_Month + 1 < 13 then
                  Last_Day_Month = DateAdd("d",-1,(Summary_Month + 1) & "/1/" & Summary_Year)
                else
                  Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
                end if  
                SQL = SQL & " (Reg_Approval_Date >= '" & Summary_Month & "/1/" & Summary_Year & "' AND Reg_Approval_Date <= '" & Last_Day_Month & "') "
                SQL = SQL & "AND Logon IS NULL or Logon='' "

              case 13            ' Active Accounts
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CInt(False) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                if Summary_Month + 1 < 13 then
                  Last_Day_Month = DateAdd("d",-1,(Summary_Month + 1) & "/1/" & Summary_Year)
                else
                  Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
                end if  
                SQL = SQL & " (Reg_Approval_Date >= '1/1/2001' AND Reg_Approval_Date <= '" & Last_Day_Month & "') AND ExpirationDate >'" & Last_Day_Month & "' "

              case 14            ' Last Logons
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CInt(False) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                if Summary_Month + 1 < 13 then
                  Last_Day_Month = DateAdd("d",-1,(Summary_Month + 1) & "/1/" & Summary_Year)
                else
                  Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
                end if  
                SQL = SQL & " (Logon >= '" & Summary_Month & "/1/" & Summary_Year & "' AND Logon <= '" & Last_Day_Month & "') "

              case 15            ' Unique Logons
                SQL = "SELECT COUNT(DISTINCT Account_ID) AS Count_" & z & " FROM Activity WHERE "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                if Summary_Month + 1 < 13 then
                  Last_Day_Month = DateAdd("d",-1,(Summary_Month + 1) & "/1/" & Summary_Year)
                else
                  Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
                end if  
                SQL = SQL & " (View_Time >= '" & Summary_Month & "/1/" & Summary_Year & "' AND View_Time <= '" & Last_Day_Month & "') "

              case 16            ' Sessions
                SQL = "SELECT Count(DISTINCT Session_ID) AS Count_" & z & " FROM Activity WHERE "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                if Summary_Month + 1 < 13 then
                  Last_Day_Month = DateAdd("d",-1,(Summary_Month + 1) & "/1/" & Summary_Year)
                else
                  Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
                end if  
                SQL = SQL & " (View_Time >= '" & Summary_Month & "/1/" & Summary_Year & "' AND View_Time <= '" & Last_Day_Month & "') "

            end select
           
            Set rsActivity = conn.Execute(SQL)
            
            Summary(z,Summary_Month) = rsActivity("Count_" & z) * Multiplier

            select case z
              case 11           ' Total = Cummulative up to Current Month
                if Summary_Year = Year(Date()) and Summary_Month > Month(Date) then
                else
                  Summary (z,13) = Summary(z,13) + Summary(z,Summary_Month)
                end if
              case 13           ' Total = Current Month's Totals
                if Summary_Year = Year(Date()) and Summary_Month > Month(Date) then
                else
                  Summary (z,13) = Summary(z,Summary_Month)
                end if
              case else         ' Total = Cummulative for Year
                Summary(z,13) = Summary(z,13) + Summary(z,Summary_Month)     ' Year Totals for Class
            end select    
            
          next  

      next

      ' Literature Order Totals

      z = 17
      SQL = "SELECT "

      if Site_ID_Change > 0 then
        SQLWhere = "(Site_ID = " & Site_ID_Change & ") AND "
      else
        SQLWhere = " "
      end if  

      for Summary_Month = 1 to 12        

        SQL = SQL & "Count_" & Summary_Month & "=(SELECT COUNT(DISTINCT Order_Number) " &_
              "FROM dbo.Shopping_Cart_Lit " &_
              "WHERE " & SQLWhere & "(DATEPART(m, Submit_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, Submit_Date) = " & Summary_Year & "))"
        if Summary_Month < 12 then SQL = SQL & ", "

      next

      Set rsActivity = conn.Execute(SQL)

      for Summary_Month = 1 to 12
        if not isblank(rsActivity("Count_" & Summary_Month)) then
          Summary(z,Summary_Month) = rsActivity("Count_" & Summary_Month)
          Summary(z,13) = Summary(z,13) + Summary(z,Summary_Month)
        else
          Summary(z,Summary_Month) = 0
        end if
      next

      ' Literature Order Item Number Totals

      z = 18
      SQL = "SELECT "

      for Summary_Month = 1 to 12        

        SQL = SQL & "Count_" & Summary_Month & "=(SELECT COUNT(DISTINCT Item_Number) " &_
              "FROM dbo.Shopping_Cart_Lit " &_
              "WHERE " & SQLWhere & "(DATEPART(m, Submit_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, Submit_Date) = " & Summary_Year & "))"
        if Summary_Month < 12 then SQL = SQL & ", "

      next

      Set rsActivity = conn.Execute(SQL)

      for Summary_Month = 1 to 12
        if not isblank(rsActivity("Count_" & Summary_Month)) then
          Summary(z,Summary_Month) = rsActivity("Count_" & Summary_Month)
          Summary(z,13) = Summary(z,13) + Summary(z,Summary_Month)
        else
          Summary(z,Summary_Month) = 0
        end if
      next
      
      ' Literature Order Item Number Totals

      z = 19
      SQL = "SELECT "

      for Summary_Month = 1 to 12        

        SQL = SQL & "Count_" & Summary_Month & "=(SELECT SUM(Quantity) " &_
              "FROM dbo.Shopping_Cart_Lit " &_
              "WHERE " & SQLWhere & "(DATEPART(m, Submit_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, Submit_Date) = " & Summary_Year & "))"
        if Summary_Month < 12 then SQL = SQL & ", "

      next

      Set rsActivity = conn.Execute(SQL)

      for Summary_Month = 1 to 12
        if not isblank(rsActivity("Count_" & Summary_Month)) then
          Summary(z,Summary_Month) = rsActivity("Count_" & Summary_Month)
          Summary(z,13) = Summary(z,13) + Summary(z,Summary_Month)
        else
          Summary(z,Summary_Month) = 0
        end if
      next


      if Summary_Year = Year(Date()) then
        response.write "<SPAN CLASS=Small>" & Translate("Date of Snapshot",Login_Language,conn) & ": " & FormatDateTime(Now(),vbLongDate) & " " & FormatDateTime(Now(),vbLongTime) & " " & Translate("PST",Login_Language,conn) & "</SPAN><BR>" & vbCrLf
      end if  

      Call Table_Begin
      response.write "<TABLE CELLPADDING=2 CELLSPACING=1 BORDER=0  WIDTH=""100%"">" & vbCrLf
      
      for x = -1 to Summary_Class
        if x = -1 then
          response.write "<TR>" & vbCrLf
          response.write "<TD CLASS=SmallBold BGCOLOR=""#FFCC00"">" & Summary_Year & "</TD>" & vbCrLf
          for y = 1 to 12
            response.write "<TD CLASS=SMALLBOLDGOLD ALIGN=""CENTER"" BGCOLOR=""Black"" WIDTH=""6%""><FONT COLOR=""#FFCC00"">" & Translate(MonthName(y),Login_Language,conn) & "</FONT></TD>" & vbCrLf
          next
          response.write "<TD CLASS=SMALLBOLDGOLD ALIGN=""CENTER"" BGCOLOR=""Black"" WIDTH=""6%""><FONT COLOR=""#FFCC00"">" & Translate("Total",Login_Language,conn) & "</FONT></TD>" & vbCrLf            
          response.write "</TR>" & vbCrLf
        else          
          response.write "<TR>" & vbCrLf
          response.write "<TD CLASS=SMALL BGCOLOR=""#D6D6D6"">"
          if isblank(Summary_Title(x)) then
            select case x
              case 0
                Summary_Title(x) = Translate("Order Inquiry Criteria",Login_Language,conn)
                Summary_Key(x) = Translate("Form used by the user to enter order criteria parameters such as Fluke Order Number or Customer Order Number.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 1
                Summary_Title(x) = Translate("Order Inquiry Results",Login_Language,conn)
                Summary_Key(x) = Translate("Results of Order Inquiry Criteria, either viewed on-line, or printed by the user.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 2
                Summary_Title(x) = Translate("Asset Items Accessed",Login_Language,conn)
                Summary_Key(x) = Translate("Asset Items (physical files), viewed, download or sent as an email attachment by clicking on a link at the site or the subscription service newsletter.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 3
                Summary_Title(x) = Translate("Electronic Document Fulfillment",Login_Language,conn)
                Summary_Key(x) = Translate("Asset Items (physical files), viewed or downloaded by clicking on a link on the Electronic Document Fulfillment (EDF) email.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 4
                Summary_Title(x) = Translate("Electronic Document Fulfillment for www.Fluke.com",Login_Language,conn)
                Summary_Key(x) = Translate("Asset Items (physical files), viewed or downloaded by clicking on a link on at www.Fluke.com",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 5
                Summary_Title(x) = Translate("Search",Login_Language,conn)
                Summary_Key(x) = Translate("Form used to enter Site Search Criteria.  Results when clicked would be counted separtely as either a navigation click or an asset item click",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 6
                Summary_Title(x) = Translate("Navigation Clicks",Login_Language,conn)
                Summary_Key(x) = Translate("Site navigation clicks without an asset item as the target.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 7
                Summary_Title(x) = Translate("All Clicks",Login_Language,conn)
                Summary_Key(x) = Translate("Total of all site clicks; navigation, asset items, gateway to other applications or off-site links.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 8
                Summary_Title(x) = Translate("Registrations",Login_Language,conn)                
                Summary_Key(x) = Translate("Total of new site registrations.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 9
                Summary_Title(x) = Translate("Registrations - Pending",Login_Language,conn)                
                Summary_Key(x) = Translate("Total of pending registrations not yet approved or deleted (registration denied).",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 10
                Summary_Title(x) = Translate("Registrations - Approved",Login_Language,conn)                
                Summary_Key(x) = Translate("Total registrations approved in the current month",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 11
                Summary_Title(x) = Translate("Accounts - Expired",Login_Language,conn)
                Summary_Key(x) = Translate("Accounts that have expired and have not been reviewed for reinstatement or deletion. Total is cumulative through current month, with projection into the remaining months of the current year.",Login_Language,conn)
              case 12
                Summary_Title(x) = Translate("Accounts - Logon Never",Login_Language,conn)
                Summary_Key(x) = Translate("The number of accounts who have registered, been approved, but have never logged on or accessed an asset item through the subscription service email.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 13
                Summary_Title(x) = Translate("Accounts - Active",Login_Language,conn)                
                Summary_Key(x) = Translate("Total of all Active accounts (Previous Month + Approved - Expired).",Login_Language,conn) & " " & Translate("Active Account Total does not exclude accounts that are active but have never logged in.",Login_Language,conn)
              case 14
                Summary_Title(x) = Translate("Accounts - Logon Last",Login_Language,conn)
                Summary_Key(x) = Translate("The number of accounts whose last logon was in that month.  A good trend would note the current month as high and the previous months approaching 0 the farther you are away from the current month.",Login_Language,conn)
              case 15
                Summary_Title(x) = Translate("Accounts - Logon Unique",Login_Language,conn)                
                Summary_Key(x) = Translate("Total of all unique accounts accessing the site, or accessing an asset item through the subscription service email.",Login_Language,conn)
              case 16
                Summary_Title(x) = Translate("Accounts - Logon Count",Login_Language,conn)                
                Summary_Key(x) = Translate("Total of all account sessions within a given month or session established when clicking on an asset item link withing the subscription service email.  None of the above statistics include Administrative logons.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 17
                Summary_Title(x) = Translate("Shopping Cart - Orders",Login_Language,conn)                
                Summary_Key(x) = Translate("Total of Literature Fulfilment Orders sent by the Portal to DCG.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 18
                Summary_Title(x) = Translate("Shopping Cart - Unique Item Numbers",Login_Language,conn)                
                Summary_Key(x) = Translate("Total of Unique Item Numbers for Literature Fulfilment Orders sent by the Portal to DCG.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 19
                Summary_Title(x) = Translate("Shopping Cart - Item Number Quantity",Login_Language,conn)                
                Summary_Key(x) = Translate("Total Quantity Ordered of all Item Numbers for Literature Fulfilment Orders sent by the Portal to DCG.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
            end select
          end if
          
          response.write Summary_Title(x)
          response.write "</TD>" & vbCrLf
          
          for y = 1 to 12
            response.write "<TD CLASS=SMALL ALIGN=""RIGHT"" BGCOLOR="""
            if Summary_Year = Year(Date()) and y = Month(Date) then
              response.write "#FFCC99"
            elseif Summary_Year = Year(Date()) and y > Month(Date) then
              response.write "#99FF99"
            elseif x = 9 or x = 11 or x = 12 or x = 14 then
              if (Summary_Year = Year(Date()) and y <= Month(Date) - 3 and Abs(Summary(x,y)) > 0) or (Summary_Year < Year(Date()) and Abs(Summary(x,y)) > 0) then
                response.write "Yellow"
              elseif (Summary_Year = Year(Date()) and y <= Month(Date) - 1 and Abs(Summary(x,y)) > 0) or (Summary_Year < Year(Date()) and Abs(Summary(x,y)) > 0) then                
                response.write "#FFFF99"
              else
                response.write "White"
              end if
            elseif x = 13 then
              response.write "#F6F6F6"
            else
              response.write "White"
            end if
            response.write """>"
            response.write FormatNumber(Summary(x,y),0)
            response.write "</TD>" & vbCrLf
          next
          response.write "<TD CLASS=SMALL ALIGN=""RIGHT"" BGCOLOR=""#99CCFF"">" & FormatNumber(Summary(x,13),0) & "</TD>" & vbCrLf          
          response.write "</TR>" & vbCrLf
        end if  
      next

      response.write "</TABLE>" & vbCrLf
      Call Table_End

      response.write "<P>"
      
      response.flush

    next
    
    set Activity = nothing
    
    response.write "<UL>" & vbCrLf
    for x = 0 to Summary_Class
      response.write "<SPAN CLASS=SmallBold><LI>" & Summary_Title(x) & "</SPAN><SPAN Class=Small> - " & Summary_Key(x) & "</LI>" & vbCrLf
    next  
    response.write "</UL>" & vbCrLf

  ' --------------------------------------------------------------------------------------
  ' Activity by Item Number EEF/WWW
  ' --------------------------------------------------------------------------------------

  elseif Utility_ID = 72 and Admin_Access <> 3 then

    TableOn = False
    
    if Admin_Access <> 1 then
      Call Nav_Border_Begin
      Call Main_Menu
    else
      Call Nav_Border_Begin    
      Call Metrics_Menu
    end if
    Call Nav_Border_End

    response.write "<FORM NAME=""Activity_EEF"" METHOD=""POST"" ACTION=""Site_Utility.asp"">" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""ID"" VALUE=""site_utility"">" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Utility_ID"" VALUE=""" & Utility_ID & """>" & vbCrLf

    ' Search for Item Numbers (Not Null will exclude certain other filters)
    if not isblank(request.form("Item_Numbers")) then
      item_numbers = replace(request.form("Item_Numbers")," ","")
    else
      item_numbers = ""
    end if

    Call Table_Begin
    response.write "<SPAN CLASS=SmallBoldWhite>" & Translate("Begin Date",Login_Language,conn) & ": "
    if isblank(request.form("Begin_Date")) then
      Begin_Date = Date()
    elseif isdate(request.form("Begin_Date")) then
      Begin_Date = request.form("Begin_Date")
    else
      response.write Translate("Invalid Date - Reseting to Today's Date",Login_Language,conn) & "<P>" & vbCrLf
      Begin_Date = Date()      
    end if
    response.write "</SPAN>" & vbCrLf
    response.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""Begin_Date"" VALUE=""" & Begin_Date & """ SIZE=""6"">" & vbCrLf
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Span",Login_Langugage,conn) & ":</SPAN> " & vbCrLf

    if isblank(request.form("Interval")) then
      Interval = 1
    elseif isnumeric(request.form("Interval")) then
      if FormatDateTime(Begin_Date) = FormatDateTime(Date()) and request.form("Interval") >= 0 then
        Interval = 1
      elseif request.form("Interval") > 90 then
        Interval = 90
      elseif request.form("Interval") < -90 then
        Interval = -90
      else  
        Interval = request.form("Interval")
      end if  
    end if

    ' Interval in Days
    response.write "<SELECT CLASS=Small NAME=""Interval"">"

    for x = 90 to -90 step -1
      select case x
        case 90,60,30,14,7          ' Future Days
          response.write "<OPTION VALUE=""" & x & """"
          if CInt(Interval) = x then response.write " SELECTED"
          response.write ">+" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
        case 1
          response.write "<OPTION CLASS=Region5NavSmall VALUE=""" & x & """"
          if CInt(Interval) = 1 then response.write " SELECTED"
          response.write ">+1" & " " & Translate("Day",Login_Language,conn) & "</OPTION>" & vbCrLf                
        case -90,-60,-30,-14,-7     ' Past Days
          response.write "<OPTION VALUE=""" & x & """"
          if CInt(Interval) = x then response.write " SELECTED"
          response.write ">-" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
      end select
    next

    response.write "</SELECT>" & vbCrLf
    response.write "<BR>" & vbCrLf
    
    ' Category / Sub-Category
    
    if not isblank(request.form("Category_Code")) then
      Category_Code = request.form("Category_Code")
    else
      Category_Code = "all"
    end if  

    SQLCategory = "SELECT DISTINCT dbo.Calendar.Sub_Category " &_
                  "FROM dbo.Activity LEFT OUTER JOIN " &_
                  "     dbo.Calendar ON dbo.Activity.Calendar_ID = dbo.Calendar.ID CROSS JOIN " &_
                  "     dbo.Calendar_Category " &_
                  "WHERE     (dbo.Activity.CMS_Site IS NOT NULL) AND (dbo.Calendar.Sub_Category IS NOT NULL) " &_
                  "ORDER BY dbo.Calendar.Sub_Category"
                  
    Set rsCategory = Server.CreateObject("ADODB.Recordset")
    rsCategory.Open SQLCategory, conn, 3, 3

    response.write "<SPAN CLASS=SmallBoldWhite>" & Translate("Category",Login_Langugage,conn) & ":</SPAN>&nbsp;&nbsp;" & vbCrLf
    response.write "<SELECT CLASS=Small NAME=""Category_Code"">"

    response.write "<OPTION CLASS=Region5NavSmall VALUE=""all"""
    if LCase(Category_Code) = "all" then response.write " SELECTED"
    response.write ">" & " " & Translate("All",Login_Language,conn) & "</OPTION>" & vbCrLf                
   
    do while not rsCategory.EOF

      response.write "<OPTION CLASS=Small VALUE=""" & rsCategory("Sub_Category") & """"
      if LCase(Category_Code) = LCase(rsCategory("Sub_Category")) then response.write " SELECTED"
      response.write ">" & " " & rsCategory("Sub_Category") & "</OPTION>" & vbCrLf
      
      rsCategory.MoveNext

    loop
    
    rsCategory.close
    set rsCategory  = nothing
    set SQLCategory = nothing
    
    response.write "</SELECT>"
    
    ' Country/Region
    
    if not isblank(request.form("Country_Code")) then
      Country_Code = request.form("Country_Code")
    else
      Country_Code = "all"
    end if  

    SQLCountry = "SELECT DISTINCT SUBSTRING(CMS_Site, 1, 2) AS Country_Code " &_
                 "FROM Activity " &_
                 "WHERE (SUBSTRING(CMS_Site, 1, 2) IS NOT NULL) ORDER BY Country_Code"

    Set rsCountry = Server.CreateObject("ADODB.Recordset")
    rsCountry.Open SQLCountry, conn, 3, 3

    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Country",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
    response.write "<SELECT CLASS=Small NAME=""Country_Code"">"

    response.write "<OPTION CLASS=Region5NavSmall VALUE=""all"""
    if LCase(Country_Code) = "all" then response.write " SELECTED"
    response.write ">" & " " & Translate("All",Login_Language,conn) & "</OPTION>" & vbCrLf                
   
    do while not rsCountry.EOF
    
      response.write "<OPTION CLASS=Small VALUE=""" & rsCountry("Country_Code") & """"
      if LCase(Country_Code) = LCase(rsCountry("Country_Code")) then response.write " SELECTED"
      response.write ">" & " " & UCase(rsCountry("Country_Code")) & "</OPTION>" & vbCrLf
    
      rsCountry.MoveNext
    
    loop
    
    rsCountry.close
    set rsCountry  = nothing
    set SQLCountry = nothing
    
    response.write "</SELECT>"
    
    ' Country/Local/Language
    
    if not isblank(request.form("Local_Code")) then
      Local_Code = request.form("Local_Code")
    else
      Local_Code = "all"
    end if  

    SQLLocal = "SELECT DISTINCT SUBSTRING(CMS_Site, 3, 2) AS Local_Code " &_
               "FROM Activity " &_
               "WHERE (SUBSTRING(CMS_Site, 3, 2) IS NOT NULL) ORDER BY Local_Code"

    Set rsLocal = Server.CreateObject("ADODB.Recordset")
    rsLocal.Open SQLLocal, conn, 3, 3

    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Local",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
    response.write "<SELECT CLASS=Small NAME=""Local_Code"">"

    response.write "<OPTION CLASS=Region5NavSmall VALUE=""all"""
    if LCase(Local_Code) = "all" then response.write " SELECTED"
    response.write ">" & " " & Translate("All",Login_Language,conn) & "</OPTION>" & vbCrLf                
   
    do while not rsLocal.EOF
    
      response.write "<OPTION CLASS=Small VALUE=""" & rsLocal("Local_Code") & """"
      if LCase(Local_Code) = LCase(rsLocal("Local_Code")) then response.write " SELECTED"
      response.write ">" & " " & UCase(rsLocal("Local_Code")) & "</OPTION>" & vbCrLf
    
      rsLocal.MoveNext
    
    loop
    
    rsLocal.close
    set rsLocal  = nothing
    set SQLLocal = nothing
    
    response.write "</SELECT>" & vbCrLf
    
    ' Group by
    
    if not isblank(request.form("Group_By")) then
      Group_By = request.form("Group_By")
    else
      Group_By = 0
    end if  
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Group by",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
    response.write "<SELECT CLASS=Small NAME=""Group_By"">"

    response.write "<OPTION CLASS=Small VALUE=""0"""
    if Group_By = 0 then response.write " SELECTED"
    response.write ">" & " " & Translate("Category",Login_Language,conn) & "</OPTION>" & vbCrLf                
        
    response.write "<OPTION CLASS=Small VALUE=""1"""
    if Group_By = 1 then response.write " SELECTED"
    response.write ">" & " " & Translate("Item Number",Login_Language,conn) & "</OPTION>" & vbCrLf
    
    response.write "<OPTION CLASS=Small VALUE=""2"""
    if Group_By = 2 then response.write " SELECTED"
    response.write ">" & " " & Translate("Site Path",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "</SELECT>" & vbCrLf

    ' Sort by
    
    if not isblank(request.form("Sort_By")) then
      Sort_By = request.form("Sort_By")
    else
      Sort_By = 0
    end if  
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Sort by",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
    response.write "<SELECT CLASS=Small NAME=""Sort_By"">"

    response.write "<OPTION CLASS=Small VALUE=""0"""
    if Sort_By = 0 then response.write " SELECTED"
    response.write ">" & " " & Translate("Item Number",Login_Language,conn) & "</OPTION>" & vbCrLf                
        
    response.write "<OPTION CLASS=Small VALUE=""1"""
    if Sort_By = 1 then response.write " SELECTED"
    response.write ">" & " " & Translate("Category",Login_Language,conn) & "</OPTION>" & vbCrLf
    
    response.write "<OPTION CLASS=Small VALUE=""2"""
    if Sort_By = 2 then response.write " SELECTED"
    response.write ">" & " " & Translate("Site Path",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "</SELECT>" & vbCrLf

    ' Submit Button
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE="" " & Translate("GO",Login_Language,conn) & " "">" & vbCrLf

    ' Reset Button
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""RESET"" NAME=""RESET"" VALUE="" " & Translate("Reset",Login_Language,conn) & " "">" & vbCrLf

    response.write "<BR>" & vbCrLf

    ' Individual Item Numbers
    response.write "<SPAN CLASS=SmallBoldWhite>" & Translate("Asset ID or Item Number",Login_Language,conn) & ": "
    response.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""Item_Numbers"" VALUE=""" & item_numbers & """ SIZE=""30"">&nbsp;&nbsp;" & vbCrLf
    response.write "<SPAN CLASS=SmallWhite>(" & Translate("Separate multiple Asset ID and Item Numbers with a comma.",Login_Language,conn) & ")</SPAN><BR>"
    Call Table_End
    response.write "</FORM>"
    
    ' Filter Specific WHERE clause
    
    if Interval >= 0  then
      SQLWhere = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Activity.View_Time, 102), 102) <= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102))) "
    else
      SQLWhere = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102)) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Activity.View_Time,102), 102) <= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "
    end if
   
    SQL = "SELECT     dbo.Activity.Site_ID, dbo.Activity.Calendar_ID, dbo.Calendar.Item_Number, dbo.Calendar.Revision_Code, dbo.Calendar.Product, dbo.Calendar.Title, dbo.Calendar.Sub_Category, dbo.Calendar.Language, dbo.Activity.View_Time, " &_
          "                      dbo.Activity.CMS_Site, dbo.CMS_XReference.CMS_Path AS CMS_Path " &_
          "FROM         dbo.Activity LEFT OUTER JOIN " &_
          "                      dbo.Calendar ON dbo.Activity.Calendar_ID = dbo.Calendar.ID LEFT OUTER JOIN " &_
          "                      dbo.CMS_XReference ON dbo.Activity.CMS_ID = dbo.CMS_XReference.ID "

    SQL = SQL & SQLWhere & "AND (dbo.Activity.CMS_Site IS NOT NULL) "
    
    if Category_Code <> "all" then
      SQL = SQL & "AND (dbo.Calendar.Sub_Category='" & Category_Code & "') "
    end if
    
    if Country_Code <> "all" then
      SQL = SQL & "AND (SUBSTRING(dbo.Activity.CMS_Site,1,2)='" & Country_Code & "') "
    end if  
      
    if Local_Code <> "all" then
      SQL = SQL & "AND (SUBSTRING(dbo.Activity.CMS_Site,3,2)='" & Local_Code & "') "
    end if  

    item_numbers_max = 0
    if not isblank(item_numbers) then
      if instr(item_numbers,",") > 0 then
        item_number = Split(item_numbers,",")
        item_number_max = Ubound(item_number)
      else
        reDim item_number(0)
        item_number(0) = item_numbers
      end if

      for x = 0 to item_number_max
        if x = 0 then
          SQL = SQL & "AND ("
        else
          SQL = SQL & "OR "  
        end if
        if len(item_number(x)) < 6 then
          SQL = SQL & "dbo.Activity.Calendar_ID='" & Trim(item_number(x)) & "' "
        else  
          SQL = SQL & "dbo.Calendar.Item_Number='" & Trim(item_number(x)) & "' "
        end if  
      next
      SQL = SQL & ") "
    end if    

    select case Sort_By
      case 0
        SQL = SQL & "ORDER BY dbo.Calendar.Item_Number, dbo.Activity.CMS_Site, dbo.CMS_XReference.CMS_Path"
      case 1
        SQL = SQL & "ORDER BY dbo.Calendar.Sub_Category, dbo.Calendar.Item_Number"
      case 2
        SQL = SQL & "ORDER BY dbo.CMS_XReference.CMS_Path, dbo.Activity.CMS_Site, dbo.Calendar.Item_Number"
    end select
        
    Set rsActivity = Server.CreateObject("ADODB.Recordset")
'response.write SQL & "<P>"
'response.flush
    rsActivity.Open SQL, conn, 3, 3

    if not rsActivity.EOF then
      response.write "<SPAN CLASS=Small>" & Translate("Date Span",Login_Language,conn) & ": "
      if Interval = 1 then
        response.write FormatDateTime(Begin_Date,vbLongDate)
      elseif Interval > 0 then
        response.write FormatDateTime(Begin_Date,vbLongDate) & " - " & FormatDateTime(DateAdd("d",Interval,Begin_Date),vbLongDate)
      else
        response.write FormatDateTime(DateAdd("d",Interval,Begin_Date),vbLongDate) & " - " & FormatDateTime(Begin_Date,vbLongDate)
      end if  
      response.write " " & Translate("PST",Login_Language,conn) & "</SPAN><BR>" & vbCrLf

      response.write "<TABLE WIDTH=""100%"" BORDER=""1"" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR=""#666666"" BGCOLOR=""#666666"">" & vbCrLf
      response.write "<TR>" & vbCrLf
      response.write "<TD>" & vbCrLf
      response.write "<TABLE CELLPADDING=""4"" CELLSPACING=""1"" BORDER=""0""  WIDTH=""100%"">" & vbCrLf
      response.write "<TR>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Item Number",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Rev",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Asset ID",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>" & Translate("Title",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Category",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Language",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Country",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Local",Login_Language,conn) & "</TD>" & vbCrLf
      if CInt(Group_By) <> 1 then
        response.write "<TD BGCOLOR=""#000000"" ALIGN=""LEFT"" CLASS=SmallBoldGold>" & Translate("Path",Login_Language,conn) & "</TD>" & vbCrLf
      end if    
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Count",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "</TR>" & vbCrLf

      TableOn = True

      temp_old = rsActivity("Item_Number")
      if CInt(Group_By) <> 1 then
        temp_old = temp_old  & rsActivity("CMS_Site") & rsActivity("CMS_Path")
      end if  
      temp_cnt = 0
      temp_tot = 0
      do while not rsActivity.EOF

        temp_new = rsActivity("Item_Number")
        if CInt(Group_By) <> 1 then
          temp_new = temp_new  & rsActivity("CMS_Site") & rsActivity("CMS_Path")
        end if  

        if temp_old <> temp_new then
          rsActivity.MovePrevious
          response.write "<TR>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">" & rsActivity("Item_Number") & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & rsActivity("Revision_Code") & "</TD>" & vbCrLf          
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">" & rsActivity("Calendar_ID") & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"">" & rsActivity("Title") & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"">" & rsActivity("Sub_Category") & "</TD>" & vbCrLf          
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & rsActivity("Language") & "</TD>" & vbCrLf          
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & Mid(rsActivity("CMS_Site"),1,2) & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & Mid(rsActivity("CMS_Site"),3,2) & "</TD>" & vbCrLf        
          if CInt(Group_By) <> 1 then
            response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"">" & rsActivity("CMS_Path") & "</TD>" & vbCrLf            
          end if    
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">" & FormatNumber(temp_cnt,0) & "</TD>" & vbCrLf
          response.write "</TR>" & vbCrLf
          temp_cnt = 0
          temp_old = temp_new
        else
          temp_cnt = temp_cnt + 1
          temp_tot = temp_tot + 1
        end if  

        rsActivity.MoveNext

      loop
      
      if temp_cnt >= 1 then
        rsActivity.MovePrevious
          response.write "<TR>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">" & rsActivity("Item_Number") & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & rsActivity("Revision_Code") & "</TD>" & vbCrLf          
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">" & rsActivity("Calendar_ID") & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"">" & rsActivity("Title") & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"">" & rsActivity("Sub_Category") & "</TD>" & vbCrLf          
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & rsActivity("Language") & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & Mid(rsActivity("CMS_Site"),1,2) & "</TD>" & vbCrLf
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & Mid(rsActivity("CMS_Site"),3,2) & "</TD>" & vbCrLf        
          if Group_By <> 1 then
            response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"">" & rsActivity("CMS_Path") & "</TD>" & vbCrLf            
          end if    
          response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">" & FormatNumber(temp_cnt,0) & "</TD>" & vbCrLf
          response.write "</TR>" & vbCrLf
      end if
          
      rsActivity.close
      set rsActivity = nothing
      
      response.write "<TR>" & vbCrLf
      response.write "<TD CLASS=SmallBold BGCOLOR=""CCCCCC"" ALIGN=""Right"" COLSPAN=8>" & Translate("Period Total",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""Right"""
      if CInt(Group_By) <> 1 then
        response.write " COLSPAN=2"
      end if
      response.write ">" & FormatNumber(temp_tot,0) & "</TD>" & vbCrLf
      response.write "</TR>" & vbCrLf
          
      response.write "</TABLE>" & vbCrLf
      response.write "</TD>" & vbCrLf
      response.write "</TR>" & vbCrLf
      response.write "</TABLE>" & vbCrLf & vbCrLf

    else

      response.write "<SPAN CLASS=SmallBold>" & Translate("There are 0 records that meet the filter criteria that you have specified.",Login_Language,conn) & "<SPAN><P>"

    end if
    
    with response
      .write "<P>" & vbCrLf
      .write "<SPAN CLASS=Small>" & vbCrLf
      .write "<UL>" & vbCrLf
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Begin Date",Login_Language,conn) & "</SPAN> - " & Translate("Starting or Ending Date for report, depending on how Span is set.",Login_Language,conn) & "</LI>" & vbCrLf
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Span",Login_Language,conn) & "</SPAN> - " & Translate("Increment of day or days from Begin Date.  Span in combination with Begin Date allows the reporting of data up to a +/- 90 day window.  A Span of +1 Day is inclusive of the Begin Date only.",Login_Language,conn) & "</LI>" & vbCrLf
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Category",Login_Language,conn) & "</SPAN> - " & Translate("Categories are defined by the selection of Asset Sub-Categories by the Partner Portal Content Administrators.  Selection of one of these Categories, limits the report to only data with a match.",Login_Language,conn) & "</LI>" & vbCrLf
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Country",Login_Language,conn) & "</SPAN> - " & Translate("The country code represents the internal country/local designator used by the www.Fluke.com CMS system. Selection of one of these country codes, limits the report to only data with a match.",Login_Language,conn) & "</LI>" & vbCrLf
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Local",Login_Language,conn) & "</SPAN> - " & Translate("The local code represents the internal country/local designator used by the www.Fluke.com CMS system. Selection of one of these local codes, limits the report to only data with a match.",Login_Language,conn) & "</LI>" & vbCrLf
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Group by",Login_Language,conn) & "</SPAN> - " & Translate("Group by totals similar Category/Item Number, Item Number or Site Path/Item Number records as opposed to listing each Item Number individually.",Login_Language,conn) & "</LI>" & vbCrLf
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Sort by",Login_Language,conn) & "</SPAN> - " & Translate("Sorts the report in ascending order by Item Number, Category or Site Path.",Login_Language,conn) & "</LI>" & vbCrLf      
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Site Path",Login_Language,conn) & "</SPAN> - " & Translate("The Site Path designator corresponds to the www.Fluke.com path under which the document appeared and was clicked to view by the visitor.",Login_Language,conn) & "</LI>" & vbCrLf
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Item Number",Login_Language,conn) & "</SPAN> - " & Translate("Oracle Item Number of the document.",Login_Language,conn) & "</LI>" & vbCrLf
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Asset ID",Login_Language,conn) & "</SPAN> - " & Translate("Partner Portal Asset ID Number.",Login_Language,conn) & "</LI>" & vbCrLf      
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Title",Login_Language,conn) & "</SPAN> - " & Translate("Partner Portal Asset Title of the document.",Login_Language,conn) & "</LI>" & vbCrLf      
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Language",Login_Language,conn) & "</SPAN> - " & Translate("Partner Portal Asset Language of the document.",Login_Language,conn) & "</LI>" & vbCrLf      
      .write "<LI><SPAN CLASS=SmallBold>" & Translate("Item Number",Login_Language,conn) & "</SPAN> - " & Translate("Oracle Item Number of the document.",Login_Language,conn) & "</LI>" & vbCrLf
      .write "</UL>" & vbCrLf
      .write "</SPAN>" & vbCrLf
    end with
    
  ' --------------------------------------------------------------------------------------
  ' Literature Order System = DCG
  ' --------------------------------------------------------------------------------------

  elseif Utility_ID = 73 and Admin_Access >= 3 then
  
    ' --------------------------------------------------------------------------------------
    ' Update Status for Literature Order System
    ' --------------------------------------------------------------------------------------    
    
    SQL = "SELECT   DISTINCT Order_Number, Submit_Date, Order_Status " &_
          "FROM     Shopping_Cart_Lit " &_
          "WHERE    Submit_Date IS NOT NULL AND Order_Number IS NOT NULL AND Order_Status <> 10 AND Order_Status <> 15  AND Order_Status <> 16 AND Order_Status <> 20 AND Order_Status <> 99 " &_
          "ORDER BY Submit_Date DESC, Order_Number DESC"

    Set rsOrders = Server.CreateObject("ADODB.Recordset")
    rsOrders.Open SQL, conn, 3, 3

    Do while not rsOrders.EOF
      Call DCG_Status("Order_Status", rsOrders("Order_Number"), conn)
      rsOrders.MoveNext      
    loop

    rsOrders.close
    set rsOrders = nothing
  
    TableOn = False
    
    if Admin_Access <> 1 then
      Call Nav_Border_Begin
      Call Main_Menu
    else
      Call Nav_Border_Begin    
      Call Metrics_Menu
    end if
    Call Nav_Border_End

    response.write "<FORM NAME=""Activity_LOS"" METHOD=""POST"" ACTION=""Site_Utility.asp"">" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""ID"" VALUE=""site_utility"">" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Utility_ID"" VALUE=""" & Utility_ID & """>" & vbCrLf

    ' Search for Item Numbers (Not Null will exclude certain other filters)
    if not isblank(request("Item_Numbers")) then
      item_numbers = replace(request("Item_Numbers")," ","")
    else
      item_numbers = ""
    end if

    Call Table_Begin
    response.write "<SPAN CLASS=SmallBoldWhite>" & Translate("Begin Date",Login_Language,conn) & ": "
    if isblank(request("Begin_Date")) then
      Begin_Date = Date()
    elseif isdate(request("Begin_Date")) then
      Begin_Date = request("Begin_Date")
    else
      response.write Translate("Invalid Date - Reseting to Today's Date",Login_Language,conn) & "<P>" & vbCrLf
      Begin_Date = Date()      
    end if
    response.write "</SPAN>" & vbCrLf
    response.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""Begin_Date"" VALUE=""" & Begin_Date & """ SIZE=""6"">" & vbCrLf
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Span",Login_Langugage,conn) & ":</SPAN> " & vbCrLf

    if isblank(request("Interval")) then
      Interval = 1
    elseif isnumeric(request("Interval")) then
      if FormatDateTime(Begin_Date) = FormatDateTime(Date()) and request("Interval") >= 0 then
        Interval = 1
      elseif request("Interval") > 90 then
        Interval = 90
      elseif request("Interval") < -90 then
        Interval = -90
      else  
        Interval = request("Interval")
      end if  
    end if

    ' Interval in Days
    response.write "<SELECT CLASS=Small NAME=""Interval"">"

    for x = 90 to -90 step -1
      select case x
        case 90,60,30,14,7          ' Future Days
          response.write "<OPTION VALUE=""" & x & """"
          if CInt(Interval) = x then response.write " SELECTED"
          response.write ">+" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
        case 1
          response.write "<OPTION CLASS=Region5NavSmall VALUE=""" & x & """"
          if CInt(Interval) = 1 then response.write " SELECTED"
          response.write ">+1" & " " & Translate("Day",Login_Language,conn) & "</OPTION>" & vbCrLf                
        case -90,-60,-30,-14,-7     ' Past Days
          response.write "<OPTION VALUE=""" & x & """"
          if CInt(Interval) = x then response.write " SELECTED"
          response.write ">-" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
      end select
    next

    response.write "</SELECT>" & vbCrLf

    ' Country

    if not isblank(request("Country_Code")) then
      Country_Code = request("Country_Code")
    else
      Country_Code = "all"
    end if  

    SQLCountry =  "SELECT DISTINCT dbo.UserData.Business_Country, dbo.Country.Name AS Business_Country_Name " &_
                  "FROM            dbo.Country LEFT OUTER JOIN " &_
                  "                dbo.UserData ON dbo.Country.Abbrev = dbo.UserData.Business_Country RIGHT OUTER JOIN " &_
                  "                dbo.Shopping_Cart_Lit ON dbo.UserData.ID = dbo.Shopping_Cart_Lit.Account_ID " &_
                  "WHERE          (dbo.Shopping_Cart_Lit.Submit_Date IS NOT NULL) "

    Set rsCountry = Server.CreateObject("ADODB.Recordset")
    rsCountry.Open SQLCountry, conn, 3, 3

    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Country",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
    response.write "<SELECT CLASS=Small NAME=""Country_Code"">"

    response.write "<OPTION CLASS=Region5NavSmall VALUE=""all"""
    if LCase(Country_Code) = "all" then response.write " SELECTED"
    response.write ">" & " " & Translate("All",Login_Language,conn) & "</OPTION>" & vbCrLf                
   
    do while not rsCountry.EOF
    
      response.write "<OPTION CLASS=Small VALUE=""" & rsCountry("Business_Country") & """"
      if LCase(Country_Code) = LCase(rsCountry("Business_Country")) then response.write " SELECTED"
      response.write ">" & " " & rsCountry("Business_Country_Name") & "</OPTION>" & vbCrLf
    
      rsCountry.MoveNext
    
    loop
    
    rsCountry.close
    set rsCountry  = nothing
    set SQLCountry = nothing
    
    response.write "</SELECT>"
    
    ' Sort by
    
    if not isblank(request("Sort_By")) then
      Sort_By = request("Sort_By")
    else
      Sort_By = 0
    end if  
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Sort by",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
    response.write "<SELECT CLASS=Small NAME=""Sort_By"">"

    response.write "<OPTION CLASS=Small VALUE=""0"""
    if Sort_By = 0 then response.write " SELECTED"
    response.write ">" & " " & Translate("Item Number",Login_Language,conn) & "</OPTION>" & vbCrLf                
        
    response.write "<OPTION CLASS=Small VALUE=""1"""
    if Sort_By = 1 then response.write " SELECTED"
    response.write ">" & " " & Translate("Order Number",Login_Language,conn) & "</OPTION>" & vbCrLf
    
    response.write "<OPTION CLASS=Small VALUE=""2"""
    if Sort_By = 2 then response.write " SELECTED"
    response.write ">" & " " & Translate("Account",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "<OPTION CLASS=Small VALUE=""3"""
    if Sort_By = 3 then response.write " SELECTED"
    response.write ">" & " " & Translate("Order Date",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "<OPTION CLASS=Small VALUE=""4"""
    if Sort_By = 4 then response.write " SELECTED"
    response.write ">" & " " & Translate("Company",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "</SELECT>" & vbCrLf

    ' Submit Button
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE="" " & Translate("GO",Login_Language,conn) & " "">" & vbCrLf

    ' Reset Button
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""RESET"" NAME=""RESET"" VALUE="" " & Translate("Reset",Login_Language,conn) & " "" ONCLICK=""javascript: document.Activity_LOS.Item_Numbers.value=''; document.Activity_LOS.Begin_Date.value='" & Date() & "'; document.Activity_LOS.Country_Code[0].selected=true; document.Activity_LOS.Sort_By[0].selected=true; document.Activity_LOS.Interval[5].selected=true; return false;"">" & vbCrLf

    
    ' View CSV File
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""Button"" NAME=""CSV"" VALUE=""" & Translate("View CSV File",Login_Language,conn) & """ ONCLICK=""javascript: Ck_RecordCount();"">" & vbCrLf
    response.write "<BR>" & vbCrLf
    
    ' Individual Item Numbers

    
    response.write "<SPAN CLASS=SmallBoldWhite>" & Translate("Item or Order Number",Login_Language,conn) & ": "
    response.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""Item_Numbers"" VALUE=""" & item_numbers & """ SIZE=""30"">&nbsp;&nbsp;" & vbCrLf
    response.write "<SPAN CLASS=SmallWhite>(" & Translate("Separate multiple Asset ID and Item Numbers with a comma.",Login_Language,conn) & ")</SPAN><BR>"
    Call Table_End
    
    if Interval >= 0  then
      SQLWhere = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date, 102), 102) <= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102))) "
    else
      SQLWhere = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102)) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) <= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "
    end if
    
    SQL = "SELECT dbo.Shopping_Cart_Lit.Account_ID AS Common_ID, dbo.UserData.NTLogin AS NTLogin, dbo.UserData.FirstName AS FirstName, dbo.UserData.LastName AS LastName,  " &_
          "       dbo.UserData.Company AS Company, dbo.UserData.Region AS Region, dbo.UserData.Business_Country AS Business_Country,  " &_
          "       dbo.Shopping_Cart_Lit.Shipping_Address_ID, dbo.Shopping_Cart_Lit.Item_Number AS Item_Number, dbo.Shopping_Cart_Lit.Quantity AS Quantity,  " &_
          "       dbo.Literature_Items_US.COST_CENTER AS Cost_Center, dbo.Shopping_Cart_Lit.Asset_ID AS Asset_ID,  " &_
          "       dbo.Shopping_Cart_Lit.Cart_Type AS Cart_Type, dbo.Shopping_Cart_Lit.Submit_Date AS Submit_Date,  " &_
          "       dbo.Shopping_Cart_Lit.Order_Number AS Order_Number, dbo.Shopping_Cart_Lit.Order_Status AS Ship_Status,  " &_
          "       dbo.Shopping_Cart_Lit.Order_Ship_Date AS Ship_Date, dbo.Shopping_Cart_Lit.Ship_Tracking_No AS Ship_Tracking_No,  " &_
          "       dbo.Shopping_Cart_Ship_Tracking.Name AS Ship_Carrier, dbo.Shopping_Cart_Lit.Site_ID AS Site_ID " &_
          "FROM   dbo.Shopping_Cart_Lit LEFT OUTER JOIN " &_
          "       dbo.Literature_Items_US ON dbo.Shopping_Cart_Lit.Item_Number = dbo.Literature_Items_US.ITEM LEFT OUTER JOIN " &_
          "       dbo.Shopping_Cart_Ship_Tracking ON dbo.Shopping_Cart_Lit.Ship_Carrier = dbo.Shopping_Cart_Ship_Tracking.ID LEFT OUTER JOIN " &_
          "       dbo.UserData ON dbo.Shopping_Cart_Lit.Account_ID = dbo.UserData.ID " 

          SQL = SQL & SQLWhere & " AND (dbo.Shopping_Cart_Lit.Submit_Date IS NOT NULL) AND (dbo.Shopping_Cart_Lit.Cart_Type = 'lit_us') "
                  
    item_numbers_max = 0
    if not isblank(item_numbers) then
      if instr(item_numbers,",") > 0 then
        item_number = Split(item_numbers,",")
        item_number_max = Ubound(item_number)
      else
        reDim item_number(0)
        item_number(0) = item_numbers
      end if

      for x = 0 to item_number_max
        if x = 0 then
          SQL = SQL & "AND ("
        else
          SQL = SQL & "OR "  
        end if
        if len(item_number(x)) = 8 and instr(1,trim(item_number(x)),"0") = 1 then
          SQL = SQL & "Order_Number LIKE '%" & Trim(item_number(x)) & "' "
        else  
          SQL = SQL & "Item_Number='" & Trim(item_number(x)) & "' "
        end if  
      next
      SQL = SQL & ") "
    end if    
          
    if LCase(Country_Code) <> "all" then
        SQL = SQL & " AND (dbo.UserData.Business_Country='" & Country_Code & "') "
    end if  
    
    select case Sort_By
      case 0
        SQL = SQL & "ORDER BY dbo.Shopping_Cart_Lit.Item_Number, dbo.Shopping_Cart_Lit.Submit_Date Desc"
      case 1
        SQL = SQL & "ORDER BY dbo.Shopping_Cart_Lit.Order_Number DESC, dbo.Shopping_Cart_Lit.Item_Number"
      case 2
        SQL = SQL & "ORDER BY dbo.Shopping_Cart_Lit.LastName, dbo.Shopping_Cart_Lit.FirstName, dbo.Shopping_Cart.Item_Number"
      case 3
        SQL = SQL & "ORDER BY dbo.Shopping_Cart_Lit.Submit_Date Desc, dbo.Shopping_Cart_Lit.LastName, dbo.Shopping_Cart_Lit.FirstName, dbo.Shopping_Cart.Item_Number"
      case 4
        SQL = SQL & "ORDER BY dbo.UserData.Company, dbo.Shopping_Cart_Lit.Submit_Date Desc, dbo.Shopping_Cart_Lit.LastName, dbo.Shopping_Cart_Lit.FirstName, dbo.Shopping_Cart.Item_Number"
    end select

    Set rsActivity = Server.CreateObject("ADODB.Recordset")
    'response.write SQL
    'response.flush
    'response.end
    rsActivity.Open SQL, conn, 3, 3

    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Counter"" VALUE=""" & rsActivity.RecordCount & """>" & vbCrLf
    response.write "</FORM>"

    %>
    <SCRIPT LANGUAGE="JavaScript">
    function Ck_RecordCount() {    
      if (document.Activity_LOS.Counter.value != "0") {
        location.href="/sw-administrator/SW_Order_Inquiry_Literature_CSV.asp?Site_ID=<%=Site_ID%>&Site_Code=<%=Site_Code%>&Begin_Date=<%=Begin_Date%>&Interval=<%=Interval%>&Country_Code=<%=Country_Code%>&Sort_By=<%=Sort_By%>&Item_Numbers=<%=Item_Numbers%>&Language=<%=Login_Language%>";
      }
      else {
        alert("<%=Translate("There are no records to view based on your query criteria",Login_Language,conn)%>");
      }
      return false;
    }
    </SCRIPT>
    <%

    if not rsActivity.EOF then
      response.write "<SPAN CLASS=Small>" & Translate("Date Span",Login_Language,conn) & ": "
      if Interval = 1 then
        response.write FormatDateTime(Begin_Date,vbLongDate)
      elseif Interval > 0 then
        if DateDiff("d",Date(),DateAdd("d",Interval,Begin_Date)) > 0 then
          response.write FormatDateTime(Begin_Date,vbLongDate) & " - " & FormatDateTime(Date(),vbLongDate)
        else
          response.write FormatDateTime(Begin_Date,vbLongDate) & " - " & FormatDateTime(DateAdd("d",Interval,Begin_Date),vbLongDate)
        end if  
      elseif Interval < 0 then
        response.write FormatDateTime(DateAdd("d",Interval,Begin_Date),vbLongDate) & " - " & FormatDateTime(Begin_Date,vbLongDate)
      else
        response.write FormatDateTime(DateAdd("d",Interval,Begin_Date),vbLongDate) & " - " & FormatDateTime(Begin_Date,vbLongDate)
      end if  
      response.write " " & Translate("PST",Login_Language,conn) & "</SPAN><BR>" & vbCrLf

      response.write "<TABLE WIDTH=""100%"" BORDER=""1"" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR=""#666666"" BGCOLOR=""#666666"">" & vbCrLf
      response.write "<TR>" & vbCrLf
      response.write "<TD>" & vbCrLf
      response.write "<TABLE CELLPADDING=""4"" CELLSPACING=""1"" BORDER=""0""  WIDTH=""100%"">" & vbCrLf
      response.write "<TR>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Item Number",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Quantity",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Status",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Order Date",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Ship Date",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Days",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("CC",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""LEFT""   CLASS=SmallBoldGold>" & Translate("Company",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""LEFT""   CLASS=SmallBoldGold>" & Translate("Name",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Country",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Order Number",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SmallBoldGold>" & Translate("Order",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "</TR>" & vbCrLf

      TableOn = True

      Old_Item_Number = ""
      Item_Number_Count = 0
      Item_Number_Total = 0
      
      do while not rsActivity.EOF
      
        if Old_Item_Number = rsActivity("Item_Number") then
          Item_Number_Count = Item_Number_Count + rsActivity("Quantity")
          Item_Number_Total = Item_Number_Total + rsActivity("Quantity")                    
        else
          Old_Item_Number = rsActivity("Item_Number")
          Item_Number_Count = Item_Number_Count + rsActivity("Quantity")
          Item_Number_Total = Item_Number_Total + rsActivity("Quantity")          
        end if  
      
        response.write "<TR>" & vbCrLf
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">"  & rsActivity("Item_Number") & "</TD>" & vbCrLf
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & rsActivity("Quantity") & "</TD>" & vbCrLf          
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">"

        ' Order / Shippment Status
        
        ' 01 To Be Approved
        ' 02 Order Received
        ' 03 In Process
        ' 04 Order Accepted
        ' 05 In Process
        ' 10 Shipped
        ' 11 In Shipping
        ' 15 Shipped
        ' 16 POD Order
        ' 20 Invoiced
        ' 80 Back-Ordered
        ' 99 Cancelled
        
        select case rsActivity("Ship_Status")
          case 10, 15, 20
            response.write "<IMG SRC=""/Images/Check_Green.gif"" BORDER=0>&nbsp;" & Translate("Shipped",Login_Language,conn) & vbCrLf
          case 16
            response.write "<IMG SRC=""/Images/Check_Green.gif"" BORDER=0>&nbsp;" & Translate("See Note 1",Login_Language,conn) & vbCrLf
          case 80
            response.write Translate("Back-Ordered",Login_Language,conn) & vbCrLf
          case 99
            response.write Translate("Cancelled",Login_Language,conn) & vbCrLf
          case 100
            response.write Translate("Unknown",Login_Language,conn) & vbCrLf
          case 0
            response.write Translate("In Process",Login_Language,conn) & vbCrLf
          case else
            response.write Translate("In Process",Login_Language,conn) & vbCrLf
         end select
        response.write "</TD>" & vbCrLf
        
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">"   & FormatDate(1,rsActivity("Submit_Date")) & "</TD>" & vbCrLf          
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"">"
        if isblank(rsActivity("Ship_Date")) then
          response.write "&nbsp;"
        else  
          response.write FormatDate(1,rsActivity("Ship_Date"))
        end if
        response.write "</TD>" & vbCrLf
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">"
        Delta = rsActivity("Ship_Date")
        if isblank(Delta) then
          Delta = Date()
          response.write "<SPAN CLASS=SmallRed>"
        else
          response.write "<SPAN CLASS=Small>"          
        end if  
        response.write DateDiff("D",rsActivity("Submit_Date"),Delta)
        response.write "</SPAN></TD>" & vbCrLf                    
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">"   & rsActivity("Cost_Center") & "</TD>" & vbCrLf
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"">"   & rsActivity("Company") & "</TD>" & vbCrLf          
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"">" & rsActivity("LastName") & ", " & rsActivity("FirstName") & "</TD>" & vbCrLf          
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & rsActivity("Business_Country") & "</TD>" & vbCrLf
        response.write "<TD CLASS=Small BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"">" & replace(rsActivity("Order_Number"),"FLUKECO","") & "</TD>" & vbCrLf
        
        response.write "<TD CLASS=Small BGCOLOR=""#666666"" ALIGN=""CENTER"">"
        if not isblank(rsActivity("NTLogin")) then
          temp = "&Logon_Cart=" & Server.URLEncode(rsActivity("NTLogin"))
        else
          temp = ""
        end if
        Admin_URL = "/SW-Common/SW-Order_Inquiry_Literature.asp?Sync=True&Site_ID=" & Site_ID & "&Logon_User=" & Server.URLEncode(Session("Logon_User")) & temp & "&Admin_URL=" & Server.URLEncode("/sw-administrator/Site_Utility.asp?Site_ID=" & Site_ID & "&ID=site_utility&Utility_ID=73&Begin_Date=" & request("Begin_Date") & "&Interval=" & request("Interval") & "&Country_Code=" & request("Country_Code") & "&Sort_By=" & request("Sort_By") & "&Item_Numbers=" & request("Item_Numbers"))
        response.write "<A HREF=""javascript:void(0);"" TITLE=""View Shopping Cart History"
        response.write """ LANGUAGE=""JavaScript"" onclick=""location.href='" & Admin_URL & "'; return false;"">"
        response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;" & Translate("Detail",Login_Language,conn) & "&nbsp;</SPAN></A>"
        response.write "</TD>" & vbCrLf                            
        response.write "</TR>" & vbCrLf
        
        rsActivity.MoveNext
        
        if Sort_By = 0 then
          if not rsActivity.EOF then
            if Item_Number_Old <> rsActivity("Item_Number") then
              response.write "<TR>" & vbCrLf
              response.write "<TD CLASS=SmallBold BGCOLOR=""CCCCCC"" ALIGN=""Left"">" & Translate("Item Total",Login_Language,conn) & "</TD>" & vbCrLf
              response.write "<TD CLASS=Small BGCOLOR=""#EEEEEE"" ALIGN=""Center"""
              response.write ">" & FormatNumber(Item_Number_Count,0) & "</TD>" & vbCrLf
              response.write "<TD CLASS=SmallBold BGCOLOR=""CCCCCC"" ALIGN=""Right"" COLSPAN=10>&nbsp;</TD>" & vbCrLf
              response.write "</TR>" & vbCrLf
              Item_Number_Old = rsActivity("Item_Number")
              Item_Number_Count = 0
            end if  
          else
            response.write "<TR>" & vbCrLf
            response.write "<TD CLASS=SmallBold BGCOLOR=""CCCCCC"" ALIGN=""Left"">" & Translate("Item Total",Login_Language,conn) & "</TD>" & vbCrLf
            response.write "<TD CLASS=Small BGCOLOR=""#EEEEEE"" ALIGN=""Center"""
            response.write ">" & FormatNumber(Item_Number_Count,0) & "</TD>" & vbCrLf
            response.write "<TD CLASS=SmallBold BGCOLOR=""CCCCCC"" ALIGN=""Right"" COLSPAN=10>&nbsp;</TD>" & vbCrLf
            response.write "</TR>" & vbCrLf
            Item_Number_Count = 0            
          end if    
        end if


      loop
          
      rsActivity.close
      set rsActivity = nothing
      
      response.write "<TR>" & vbCrLf
      response.write "<TD CLASS=SmallBold BGCOLOR=""White"" ALIGN=""Left"">" & Translate("Period Total",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD CLASS=SmallBoldWhite BGCOLOR=""SteelBlue"" ALIGN=""CENTER"""
      response.write ">" & FormatNumber(Item_Number_Total,0) & "</TD>" & vbCrLf
      response.write "<TD CLASS=SmallBold BGCOLOR=""SteelBlue"" ALIGN=""Right"" COLSPAN=10>&nbsp;</TD>" & vbCrLf
      response.write "</TR>" & vbCrLf
          
      response.write "</TABLE>" & vbCrLf
      response.write "</TD>" & vbCrLf
      response.write "</TR>" & vbCrLf
      response.write "</TABLE>" & vbCrLf & vbCrLf

    else

      response.write "<SPAN CLASS=SmallBold>" & Translate("There are 0 records that meet the filter criteria that you have specified.",Login_Language,conn) & "<SPAN><P>"

    end if

  ' --------------------------------------------------------------------------------------
  ' List - Directory Contents
  ' --------------------------------------------------------------------------------------
  
  elseif Utility_ID = 90 and Admin_Access >= 4 then
    response.write "This tools is unavilable at this time.<BR>"

  ' --------------------------------------------------------------------------------------
  ' List - Directory Contents
  ' --------------------------------------------------------------------------------------

  elseif Utility_ID = 98 and Admin_Access >= 4 then

    if isblank(request("Interval")) then
      Interval = -7
    else
      Interval = CInt(request("Interval"))
    end if
   
    SQL =       "SELECT Calendar_Upload_Status.Site_ID AS Site_ID, Calendar_Upload_Status.BTime AS BTime, Calendar_Upload_Status.ETime AS ETime, "
    SQL = SQL & "Calendar_Upload_Status.Path_Source AS Path_Source, Calendar_Upload_Status.Path_Destination AS Path_Destination, "
    SQL = SQL & "Calendar_Upload_Status.Bytes AS Bytes, Calendar_Upload_Status.Status AS Status, "
    SQL = SQL & "Calendar_Upload_Status.Error_Number AS Error_Number, Calendar_Upload_Status.Error_Description AS Error_Description, "
    SQL = SQL & "UserData.FirstName, UserData.LastName, dbo.Site.Site_Code "
    SQL = SQL & "FROM  Calendar_Upload_Status "
    SQL = SQL & "INNER JOIN UserData ON Calendar_Upload_Status.Account_ID = UserData.ID "
    SQL = SQL & "INNER JOIN Site ON Calendar_Upload_Status.Site_ID = Site.ID "

    if Admin_Access < 9 then
      SQL = SQL & "WHERE Calendar_Upload_Status.Site_ID=" & Site_ID & " AND "
    else
      SQL = SQL & "WHERE "
    end if    

    SQL = SQL & "Calendar_Upload_Status.ETime >= '" & DateAdd("d",Interval,Date()) & "' "
    SQL = SQL & "ORDER BY Site_Code, Calendar_Upload_Status.ETime Desc"

    Set rsStatus = Server.CreateObject("ADODB.Recordset")
    rsStatus.Open SQL, conn, 3, 3
    
    if rsStatus.EOF then
      response.write "There are file upload records for this site based on the ""Limit to:"" criteria that you have selected.<BR><BR>"
      TableOn = false
    else   
      TableOn = true
    end if  

    Response.write "<FORM NAME=""Dummy-98"">" & vbCrLf

    if Admin_Access <> 1 then
      Call Nav_Border_Begin
      Call Main_Menu
    else
      Call Nav_Border_Begin
      Call Metrics_Menu
    end if
    Call Nav_Border_End
      
    response.write "&nbsp;&nbsp;<SPAN CLASS=Small>Limit Listing to:</SPAN>&nbsp;"

    response.write "<SELECT CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='/SW-Administrator/site_utility.asp?ID=site_utility&Site_ID=" & Site_ID & "&Utility_ID=" & Utility_ID & "&Interval='+this.options[this.selectedIndex].value"" NAME=""Interval"">"
    response.write "<OPTION VALUE=""0"""
    if Interval = 0 then response.write " SELECTED"
    response.write ">Today</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""-7"""
    if Interval = -7 then response.write " SELECTED"
    response.write ">Last 7 Days</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""-30"""
    if Interval = -30 then response.write " SELECTED"
    response.write ">Last 30 Days</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""-60"""
    if Interval = -60 then response.write " SELECTED"
    response.write ">Last 60 Days</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""-90"""
    if Interval = -90 then response.write " SELECTED"
    response.write ">Last 90 Days</OPTION>" & vbCrLf
    response.write "</SELECT>"
      
    if Not rsStatus.EOF then
    
      tcount = 0  ' Total   Counter
      fcount = 0  ' Failure Counter
            
      do while not rsStatus.EOF
        tcount = tcount + 1
        if CInt(rsStatus("Status")) = CInt(False) then
          fcount = fcount + 1
        end if
        rsStatus.MoveNext  
      loop
        
      rsStatus.MoveFirst
    
      response.write "&nbsp;&nbsp;<SPAN CLASS=Small>Total Uploads: "  & tcount & "</SPAN>"
      response.write "&nbsp;&nbsp;<SPAN CLASS=Small>Total Failures: " & fcount & "</SPAN>"
      
    end if
           
    response.write "<BR><BR>"
      
    if Not rsStatus.EOF then
      %>
    	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
        <TR>
          <TD>
            <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
              <TR>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Begin Time</TD>              
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>End Time</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Elapsed<BR>> 1 sec</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Uploaded<BR>By</TD>                
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Source</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Destination</TD>
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>File Size</TD>            
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Status</TD>            
                <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Error</TD>
                <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Error<BR>Description</TD>
          		</TR>
      <%
    end if
     
    Do while not rsStatus.EOF
      %>
          		<TR>
               
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%=rsStatus("BTime")%>
                </TD>
  
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%=rsStatus("ETime")%>
                </TD>
           			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%=ConvertTime(DateDiff("s",rsStatus("BTime"),rsStatus("ETime")))%>
                </TD>

                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%=rsStatus("Firstname") & " " & rsStatus("Lastname")%>
                </TD>
          			        			
                <TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%=Replace(Replace(rsStatus("Path_Source"),"\","<BR>\")," ","&nbsp;")%>
                </TD>
         
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" VALIGN="TOP" CLASS=Small>
                  <%
                  ThisSite = Replace(rsStatus("Path_Destination"),"D:\inetpub\extranet\","") %><%
                  ThisSite = Replace(ThisSite,"D:\InetPub\extranet\","") %><%
                  ThisSite = Replace(ThisSite," ","&nbsp;")
                  ThisSite = "<SPAN CLASS=SmallRed>" & Mid(ThisSite,1,Instr(1,ThisSite,"\")-1)  & "</SPAN>" & Replace(Mid(ThisSite, Instr(1,ThisSite,"\")),"\","<BR>\") %><%
                  response.write ThisSite                  
                  %>
                </TD>
          			     
          			<TD BGCOLOR="#FFFFFF" ALIGN="Right" CLASS=Small>
                  <%=FormatNumber((CDbl(rsStatus("Bytes"))/1024),0) & " KBytes" %>
                </TD>
          			
                <TD BGCOLOR="#FFFFFF" ALIGN="CENTER" CLASS=Small>
                  <%
                  if CInt(rsStatus("Status")) = CInt(True) then
                    response.write "Complete"
                  else
                    response.write "<SPAN CLASS=SmallRedBold>Failure</SPAN>"
                  end if  
                  %>
                </TD>
  
          			<TD BGCOLOR="#FFFFFF" ALIGN="CENTER" CLASS=Small>
                  <%
                  if rsStatus("Error_Number") = 0 then
                    response.write "None"
                  else
                    response.write "<SPAN CLASS=SmallRedBold>" & rsStatus("Error_Number") & "</SPAN>"
                  end if
                  %>
                </TD>
  
          			<TD BGCOLOR="#FFFFFF" ALIGN="LEFT" CLASS=Small>
                  <%=rsStatus("Error_Description")%>
                </TD>
              </TR>
    <%
      rsStatus.MoveNext
    
    loop
                         
    rsStatus.close
    set rsStatus=nothing
  
    if TableOn then
      %>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
      <%
      
    end if                       
  
  ' --------------------------------------------------------------------------------------
  ' File - Upload Utility
  ' --------------------------------------------------------------------------------------
  
  elseif Utility_ID = 99 and Admin_Access >= 4 then
    response.write "This tools is unavilable at this time.<BR>"


  ' --------------------------------------------------------------------------------------
  '
  ' --------------------------------------------------------------------------------------
  
  else
    response.write "This tools is unavilable at this time.<BR>"
  end if

' Invalid Access Level Credentials

else
  response.write "This tools is unavilable at this time.<BR>"
end if

response.write "<BR>"

if Admin_Access <> 1 then
  Call Nav_Border_Begin
  Call Main_Menu
else
  Call Nav_Border_Begin
  Call Metrics_Menu
end if  
Call Nav_Border_End

response.write "<BR><BR>"

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"--> 
<%

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------

sub Main_Menu()

  response.write "<A HREF=""/sw-administrator/default.asp?Site_ID=" & Site_ID & """ CLASS=NavLeftHighlight1>&nbsp;Main Menu&nbsp;</A>"

end sub

' --------------------------------------------------------------------------------------

sub Metrics_Menu()

  response.write "<A HREF=""/sw-administrator/Site_Metrics.asp"" CLASS=NavLeftHighlight1>&nbsp;Main Menu&nbsp;</A>"

end sub

' --------------------------------------------------------------------------------------

sub Group_Code_Table()

  response.write "&nbsp;&nbsp;&nbsp;<A HREF=""/sw-administrator/SubGroup_Codes.asp?Site_ID=" & Site_ID & """ onclick=""var MyPop1 = window.open('/sw-administrator/SubGroup_Codes.asp?Site_ID=" & Site_ID & "','MyPop1','fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=no,directories=no,location=no,width=525,height=410,left=400,top=200'); MyPop1.focus(); return false;"" CLASS=NavLeftHighlight1>&nbsp;Group Codes&nbsp;</A>"
'  response.write "&nbsp;&nbsp;&nbsp;<A HREF=""/sw-administrator/SubGroup_Codes.asp?Site_ID=" & Site_ID & """ onclick=""openit('/sw-administrator/SubGroup_Codes.asp?Site_ID=" & Site_ID & "','Vertical');return false;"" CLASS=NavLeftHighlight1>&nbsp;Group Codes&nbsp;</A>"
  
end sub

' --------------------------------------------------------------------------------------


sub Status_Colors()

  response.write "&nbsp;&nbsp;&nbsp;<A HREF=""/sw-administrator/Calendar_Status_Colors.asp"" onclick=""var MyPop2 = window.open('/sw-administrator/Calendar_Status_Colors.asp','MyPop2','fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,directories=no,location=no,width=200,height=220,left=600,top=200'); MyPop2.focus(); return false;"" CLASS=NavLeftHighlight1>&nbsp;Status Colors&nbsp;</A>"
  
end sub

' --------------------------------------------------------------------------------------

sub Element_Names()

  response.write "&nbsp;&nbsp;&nbsp;<A HREF=""/sw-administrator/Calendar_Element_Names.asp"" onclick=""var MyPop2 = window.open('/sw-administrator/Calendar_Element_Names.asp','MyPop2','fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,directories=no,location=no,width=250,height=300,left=600,top=200'); MyPop2.focus(); return false;"" CLASS=NavLeftHighlight1>&nbsp;Status Codes&nbsp;</A>"
  
end sub

' --------------------------------------------------------------------------------------

sub Activity_Methods()

  response.write "&nbsp;&nbsp;&nbsp;<A HREF=""/sw-administrator/Activity_Methods.asp"" onclick=""var MyPop2 = window.open('/sw-administrator/Activity_Methods.asp','MyPop2','fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,directories=no,location=no,width=300,height=380,left=500,top=200'); MyPop2.focus(); return false;"" CLASS=NavLeftHighlight1>&nbsp;" & Translate("Activity Method Codes",Login_Language,conn) & "&nbsp;</A>"
  
end sub

' --------------------------------------------------------------------------------------

sub Change_Region()

    if isblank(request.form("Region")) then
      Region = 0
    else
      Region = request.form("Region")
    end if
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Region",Login_Language,conn) & " : </SPAN>" & vbCrLf
    response.write "<SELECT NAME=""Region"" CLASS=SMALL>" & vbCrLf
    response.write "<OPTION VALUE=""0"""
    if CInt(Region) = 0 then response.write " SELECTED"    
    response.write ">" & Translate("All",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""1"""
    if CInt(Region) = 1 then response.write " SELECTED"
    response.write ">" & Translate("USA",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""2"""
    if CInt(Region) = 2 then response.write " SELECTED"
    response.write ">" & Translate("Europe",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""3"""
    if CInt(Region) = 3 then response.write " SELECTED"
    response.write ">" & Translate("Intercon",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "</SELECT>" & vbCrLf

end sub

' --------------------------------------------------------------------------------------

sub Change_Site

    if isblank(request.form("Site_ID_Change")) then
      Site_ID_Change = Site_ID
    else
      Site_ID_Change = request.form("Site_ID_Change")
    end if
    if Admin_Access >= 9 then
      SQLSite = "SELECT ID, Site_Description FROM Site WHERE Enabled=" & CInt(True) & " ORDER BY Site_Description"
      Set rsSite = Server.CreateObject("ADODB.Recordset")
      rsSite.Open SQLSite, conn, 3, 3
      response.write "&nbsp;&nbsp;&nbsp;&nbsp;<SPAN CLASS=SMALLBoldWhite>" & Translate("Site",Login_Language,conn) & " : </SPAN>" & vbCrLf
      response.write "<SELECT NAME=""Site_ID_Change"" CLASS=SMALL>" & vbCrLf
      response.write "<OPTION CLASS=Small VALUE=""0"""
      if Site_ID_Change = 0 then response.write " SELECTED"
      response.write ">" & Translate("All Sites",Login_Language,conn) & "</OPTION>" & vbCrLf
      do while not rsSite.EOF
        if rsSite("ID") > 0 then
          response.write "<OPTION CLASS=Small VALUE=""" & rsSite("ID") & """"
          if CInt(Site_ID_Change) = CInt(rsSite("ID")) then response.write " SELECTED"
          response.write ">" & Translate(rsSite("Site_Description"),Login_Language,conn) & "</OPTION>" & vbCrLf
        end if  
        rsSite.MoveNext
      loop
      rsSite.close
      set rsSite = nothing
      response.write "</SELECT>" & vbCrLf
    else
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID_Change"" VALUE=""" & Site_ID_Change & """>" & vbCrLf
    end if

end sub

' --------------------------------------------------------------------------------------

sub Write_SubGroups()

  SubGroups = rsUser("SubGroups")
  SubGroups = Replace(SubGroups,"submitter","<FONT COLOR=""Red"">submitter</FONT>")
  SubGroups = Replace(SubGroups,"content","<FONT COLOR=""Red"">content</FONT>")
  SubGroups = Replace(SubGroups,"account","<FONT COLOR=""Red"">account</FONT>")
  SubGroups = Replace(SubGroups,"PIK, Users,","")                                
  response.write SubGroups

end sub

' --------------------------------------------------------------------------------------

sub Write_Account_Status()

  if rsUser("NewFlag") = True then                
    response.write "<FONT COLOR=""Red"">Pending Approval</FONT>"
  else
    response.write "Active"
  end if    

end sub

' --------------------------------------------------------------------------------------

sub Write_Last_Logon()

    if isblank(rsUser("Logon")) or instr(1,rsUser("Logon"),"9999") > 0 or not isDate(rsUser("Logon")) then
      response.write "Never"
    else  
      response.write FormatDate(1,rsUser("Logon"))
    end if  

end sub

' --------------------------------------------------------------------------------------

sub Write_Expiration_Date()

  response.write "<TD BGCOLOR="
  if isdate(rsUser("ExpirationDate")) then
    if DateDiff("d",CDate(rsUser("ExpirationDate")),Date) >=20 then
      response.write """Red"""
    else
      response.write """#FFFFFF"""
    end if
  end if  
  response.write " ALIGN=""CENTER"" CLASS=Small>"
  
  if CDate(rsUser("ExpirationDate")) = CDate("09/09/9999") then                
    response.write "Never"
  else
    response.write FormatDate(1,rsUser("ExpirationDate"))
  end if    

end sub

' --------------------------------------------------------------------------------------

sub Write_Logon_History()

    SQL = "SELECT Logon.Account_ID, Logon.Logon, Logon.Logoff FROM Logon WHERE Logon.Site_ID=" & Site_ID & " Logon.Account_ID=" & rsUser("ID") & " ORDER BY Logon.Logon DESC"
    Set rsLogon = Server.CreateObject("ADODB.Recordset")
    rsLogon.Open SQL, conn, 3, 3                    
    
    if not rsLogon.EOF then
       
      do while not rsLogon.EOF
        response.write FormatDate(1,rsLogon("Logon")) & "<BR>"
        rsLogon.MoveNext
      loop  
    else
      response.write "No History"
    end if  

    rsLogon.close
    set rsLogon = nothing                       

end sub

' --------------------------------------------------------------------------------------

sub Bypass_Assets
  do while not rsActivity.EOF AND CInt(Order_Inquiry) = CInt(False)
    if rsActivity("Asset_ID") = 101 or rsActivity("Asset_ID") = 102 then
      OI_Count = OI_Count + 1
      rsActivity.MoveNext
    else
      exit do  
    end if
  loop
end sub

' --------------------------------------------------------------------------------------

sub Update_Methods
  select case CInt(rsActivity("Method"))
    case 0
      select case CInt(rsActivity("Account_ID"))
        case 1          ' Electronic Document Fulfillment EDF or WWW
          if isblank(rsActivity("CMS_Site")) then ' EDF
            Method(11) = Method(11) + 1
          else  
            Method(13) = Method(13) + 1           ' WWW
          end if  
        case else
          Method(0)  = Method(0) + 1
       end select     
    case 2, 6 
      Method(2)  = Method(2)  + 1          
    case 7, 8
      Method(7)  = Method(7)  + 1
    case 9,10
      Method(9)  = Method(9)  + 1
    case 11, 12
      Method(11) = Method(11) + 1
    case else
      Method(CInt(rsActivity("Method"))) = Method(CInt(rsActivity("Method"))) + 1
  end select
  Method(15) = Method(15)     + 1      ' Increment Totals Counter  
end sub  

' --------------------------------------------------------------------------------------

sub Display_Methods
  for x = 0 to 15
    select case x
      case 0,1,2,3,4,5,7,9,11,13,14,15         ' Display only certain fields
        response.write "<TD CLASS=SMALL ALIGN=RIGHT BGCOLOR="""
        select case x
          case 0,3
            response.write "PaleTurquoise"
          case 1,4
            response.write "Wheat"
          case 2,5
            response.write "LightPink"
          case 7
            response.write "Aquamarine"
          case 9
            response.write "LightGreen"
          case 11,13
            response.write "White"
          case 14
            response.write "Wheat"          
          case 15
            response.write "LightSteelBlue"
          case else
            response.write "Yellow"                  
        end select
        response.write """>"
        if Method(x) = 0 then
          response.write "&nbsp;"
        else
          if x = 15 then response.write "<B>"
          response.write FormatNumber(Method(x),0)
          if x = 15 then response.write "</B>"          
          Category(x) = Category(x) + Method(x)
          Period(x)   = Period(x)   + Method(x)
        end if
        response.write "</TD>" & vbCrLf
    end select
    Method(x) = 0
  next
end sub

' --------------------------------------------------------------------------------------

sub Display_Category
  for x = 0 to 15
    select case x
      case 0,1,2,3,4,5,7,9,11,13,14,15         ' Display only certain fields
        response.write "<TD CLASS=SMALL ALIGN=RIGHT BGCOLOR=""LightGrey"">"
        if Category(x) = 0 then
          response.write "&nbsp;"
        else
          if x = 15 then response.write "<B>"
          response.write FormatNumber(Category(x),0)
          if x = 15 then response.write "</B>"
        end if
        response.write "</TD>" & vbCrLf
    end select
    Category(x) = 0
  next
end sub

' --------------------------------------------------------------------------------------

Function convertTime(seconds)
   ConvSec = seconds mod 60
   if Len(ConvSec) = 1 then
         ConvSec = "0" & ConvSec
   end if
   ConvMin = (seconds mod 3600) \ 60
   if Len(ConvMin) = 1 then
         ConvMin = "0" & ConvMin
   end if
   ConvHour =  seconds \ 3600
   if Len(ConvHour) = 1 then
         ConvHour = "0" & ConvHour
   end if
   convertTime = ConvHour & ":" & ConvMin & ":" & ConvSec
end Function

'--------------------------------------------------------------------------------------

%>
  
<!--#include virtual="/include/core_countries_select.inc"-->