<%@  language="VBScript" codepage="65001" %>
<%
' --------------------------------------------------------------------------------------
' Author:     K. D. Whitlock
' Date:       06/1/2000
' --------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

' Release Task  :   688
' Updated on    :   08 Dec 2009
' Updated by    :   Amol Jagtap
' Description   :   To change all the CInt value conversion to CLng to avoid Interger range violation in future


if Utility_ID =50 then
    Response.Buffer=false   
end if

Session("BackURL_Calendar") = ""
 
Dim Site_ID
Dim Site_ID_Change
Dim Site_Clicks
Dim Region
Dim Utility_ID
Dim Unique_Item_Numbers

Dim Page_QS
Dim Record_Count
Dim Record_Limit
Dim Record_Pages
Dim ltEnabled
Dim PCID

Record_Limit  = 500
PCID     = CLng(request("PCID")) ' Page Sequence
if PCID  = Null or PCID = 0 then PCID = 0

Dim Border_Toggle
Border_Toggle = 0
'gpd RFC#1819
Dim Begin_Date,End_Date, Country_Code, Local_Code

Site_ID        = request("Site_ID")
Site_ID_Change = Site_ID

if isnumeric(request("Utility_ID")) then
  Utility_ID     = CLng(request("Utility_ID"))
else
  Utility_ID     = 999
end if
  
'' Updated for RI 579
if Utility_ID=44 then
      response.redirect "/SW-Administrator/edit_pricelists.asp?Site_ID=" & Site_ID '& "&Logon_user=" & Session("LOGON_USER")
end if

if site_id = 11 and utility_id = 53 then
	response.redirect("/Met-Support-Gold/Admin/Default.asp")
elseif site_id=82 and utility_id = 1111 then
	response.redirect("SW-PCAT_FNET_PROD_ASSET_REL.asp?site_id=" & site_id & "&Associate=True")
elseif site_id=82 and utility_id = 2222 then
	response.redirect("SW-PCAT_FNET_PROD_ASSET_REL.asp?site_id=" & site_id & "&Associate=false")	
elseif site_id=82 and utility_id = 3333 then
	response.redirect("SW-PCAT_FNET_ASSET_LIST.asp?site_id=" & site_id)	
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

Session.timeout = 60            ' Set to 1 Hour
'RFC1819
'Server.ScriptTimeout = 6 * 60   ' Set to 6 Minutes
Server.ScriptTimeout = 999   ' Set to 6 Minutes
Call Connect_SiteWide

%>
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%
''response.write Session("LOGON_USER")
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
  Case 55
    Screen_TitleX = "Reset / Align Asset Groups"      
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

Dim RegionColor(4)
RegionColor(0) = "#0000CC"
RegionColor(1) = "#99FFCC"
RegionColor(2) = "#66CCFF"
RegionColor(3) = "#FFCCFF"
RegionColor(4) = "#FFCC99"

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code        = rsSite("Site_Code")     
Screen_Title     = rsSite("Site_Description") & " - " & Screen_TitleX
Bar_Title        = rsSite("Site_Description") & "<BR><FONT CLASS=MediumBoldGold>" & Screen_TitleX & "</FONT>"
Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent

Logo             = rsSite("Logo")  
Logo_Left        = rsSite("Logo_Left")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<% 

rsSite.close
set rsSite=nothing

if Admin_Access = 1 or Admin_Access = 2 or Admin_Access >= 3 then
 
  ' --------------------------------------------------------------------------------------
  ' List Accounts by Name
  ' --------------------------------------------------------------------------------------
  
  if Utility_ID = 0 and Admin_Access >= 4 then
  
    UserNumber = 0
    
    SQL =  "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & Site_ID & " AND UserData.Fcm<>" & CLng(True)
  
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
<div id="ContentTableStart" style="position: absolute;">
</div>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#666666"
    bgcolor="#666666" id="Table1">
    <tr>
        <td>
            <table cellpadding="4" cellspacing="1" border="0" width="100%" id="Table2">
                <tr id="ContentHeader1">
                    <td bgcolor="Red" align="center" class="SmallBoldWhite">
                        Action</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        ID</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Users Name</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Company</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        City</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        State</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Country</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Phone Number</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Groups</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Account Status</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Last Logon</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Account Expiration</td>
                </tr>
                <%
     end if
     
     Do while not rsUser.EOF
       if instr(1,lcase(rsUser("SubGroups")),"admin") = 0 then
                %>
                <tr>
                    <td bgcolor="Silver" align="CENTER" class="Small">
                        <a href="account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>"
                            class="NavLeftHighlight1" onclick="location.href='account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>'"
                            value=" Edit ">&nbsp;&nbsp;Edit&nbsp;&nbsp;</a>
                    </td>
                    <td bgcolor="#FFFFFF" align="RIGHT" class="Small">
                        <% response.write rsUser("ID") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <%
                  response.write "<B>" & rsUser("LastName") & "</B>, "
                  response.write rsUser("FirstName")
                  if not isblank(rsUser("MiddleName")) then response.write " " & rsUser("MiddleName")
                  if not isblank(rsUser("Prefix")) then response.write " " & rsUser("Prefix") & ". "
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Company") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Business_City") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Business_State") %>
                    </td>
                    <%              
                response.write "<TD BGCOLOR=""" & RegionColor(rsUser("Region")) & """ ALIGN=""LEFT"" CLASS=Small>"
                response.write rsUser("Business_Country") 
                response.write "</TD>"
                    %>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write FormatPhone(rsUser("Business_Phone")) %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <%
                  Call Write_SubGroups
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="CENTER" class="Small">
                        <%
                  Call Write_Account_Status
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="CENTER" class="Small">
                        <%
                  Call Write_Last_Logon
                        %>
                    </td>
                    <%
                  Call Write_Expiration_Date
                    %>
        </td>
    </tr>
    <%
      end if 
      rsUser.MoveNext
    
    loop
                         
    rsUser.close
    set rsManager=nothing
  
    if TableOn then
    %>
</table>
</TD> </TR> </TABLE>
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
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#666666"
    bgcolor="#666666" id="Table3">
    <tr>
        <td>
            <table cellpadding="4" cellspacing="1" border="0" width="100%" id="Table4">
                <tr>
                    <td bgcolor="Red" align="CENTER" class="SmallBoldWhite">
                        Action</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        ID</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Company</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Users Name</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        City</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        State</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Country</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Phone Number</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Groups</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhiteoldGold">
                        Account Status</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Last Logon</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Account Expiration</td>
                </tr>
                <%
     end if
     
     Do while not rsUser.EOF
       if instr(1,lcase(rsUser("SubGroups")),"admin") = 0 then
                %>
                <tr>
                    <td bgcolor="Silver" align="CENTER" class="Small">
                        <a href="account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>"
                            class="NavLeftHighlight1" onclick="location.href='account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>'"
                            value=" Edit ">&nbsp;&nbsp;Edit&nbsp;&nbsp;</a>
                    </td>
                    <td bgcolor="#FFFFFF" align="RIGHT" class="Small">
                        <% response.write rsUser("ID") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Company") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <%
                  response.write "<B>" & rsUser("LastName") & "</B>, "
                  response.write rsUser("FirstName")
                  if not isblank(rsUser("MiddleName")) then response.write " " & rsUser("MiddleName")
                  if not isblank(rsUser("Prefix")) then response.write " " & rsUser("Prefix") & ". "
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Business_City") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Business_State") %>
                    </td>
                    <%              
                response.write "<TD BGCOLOR=""" & RegionColor(rsUser("Region")) & """ ALIGN=""LEFT"" CLASS=Small>"
                response.write rsUser("Business_Country")
                response.write "</TD>"
                    %>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write FormatPhone(rsUser("Business_Phone")) %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <%
                  Call Write_SubGroups
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="CENTER" class="Small">
                        <%
                  Call Write_Account_Status
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="CENTER" class="Small">
                        <%
                  Call Write_Last_Logon
                        %>
                    </td>
                    <%
                  Call Write_Expiration_Date
                    %>
        </td>
    </tr>
    <%
      end if 
      rsUser.MoveNext
    
    loop
                         
    rsUser.close
    set rsManager=nothing
  
    if TableOn then
    %>
</table>
</TD> </TR> </TABLE>
<%
      
    end if                       
  
  ' --------------------------------------------------------------------------------------
  ' List Accounts by Account Manager Name
  ' --------------------------------------------------------------------------------------
  
  elseif Utility_ID = 21 and Admin_Access >= 4 then
  
    UserNumber = 0
    
    SQL =  "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & Site_ID & " AND UserData.Fcm=" & CLng(True)
    
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
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#666666"
    bgcolor="#666666" id="Table5">
    <tr>
        <td>
            <table cellpadding="4" cellspacing="1" border="0" width="100%" id="Table6">
                <tr>
                    <td bgcolor="Red" align="CENTER" class="SmallBoldWhite">
                        Action</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        ID</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Users Name</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Company</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        City</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        State</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Country</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Phone Number</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Groups</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Account Status</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Last Logon</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Account Expiration</td>
                </tr>
                <%
     end if
     
     Do while not rsUser.EOF
       if instr(1,lcase(rsUser("SubGroups")),"admin") = 0 then
                %>
                <tr>
                    <td bgcolor="Silver" align="CENTER" class="Small">
                        <a href="account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>"
                            class="NavLeftHighlight1" onclick="location.href='account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>'"
                            value=" Edit ">&nbsp;&nbsp;Edit&nbsp;&nbsp;</a>
                    </td>
                    <td bgcolor="#FFFFFF" align="RIGHT" class="Small">
                        <% response.write rsUser("ID") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <%
                  response.write "<B>" & rsUser("LastName") & "</B>, "
                  response.write rsUser("FirstName")
                  if not isblank(rsUser("MiddleName")) then response.write " " & rsUser("MiddleName")
                  if not isblank(rsUser("Prefix")) then response.write " " & rsUser("Prefix") & ". "
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Company") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Business_City") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Business_State") %>
                    </td>
                    <%              
                response.write "<TD BGCOLOR=""" & RegionColor(rsUser("Region")) & """ ALIGN=""LEFT"" CLASS=Small>"
                response.write rsUser("Business_Country")
                response.write "</TD>"
                    %>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write FormatPhone(rsUser("Business_Phone")) %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <%
                  Call Write_SubGroups
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="CENTER" class="Small">
                        <%
                  Call Write_Account_Status
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="CENTER" class="Small">
                        <%
                  Call Write_Last_Logon
                        %>
                    </td>
                    <%
                  Call Write_Expiration_Date
                    %>
        </td>
    </tr>
    <%
      end if 
      rsUser.MoveNext
    
    loop
                         
    rsUser.close
    set rsManager=nothing
  
    if TableOn then
    %>
</table>
</TD> </TR> </TABLE>
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
    
    SQL = "SELECT Approvers.* FROM Approvers WHERE Approvers.Site_ID=" & CLng(Site_ID) & " ORDER BY Approvers.Order_Num, Approvers.Description"
    
    Set rsApproverGroups = Server.CreateObject("ADODB.Recordset")
    rsApproverGroups.Open SQL, conn, 3, 3    
    
    if rsApproverGroups.EOF then
      response.write "There are no Content Administrator Groups established for this site.<BR><BR>"
      TableOn = false
    end if
   
    SQL = "Select UserData.* FROM UserData WHERE UserData.Site_ID=" & CLng(Site_ID) & " AND (UserData.Subgroups LIKE '%content%' OR UserData.Subgroups LIKE '%administrator%') ORDER BY UserData.LastName"
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
<form name="Dummy-5" id="Form1">
    <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#666666"
        bgcolor="#666666" id="Table7">
        <tr>
            <td>
                <table cellpadding="4" cellspacing="1" border="0" width="100%" id="Table8">
                    <tr>
                        <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                            Region / Group / Sub-Region or Description</td>
                        <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                            Content Administrator&acute; Name</td>
                    </tr>
                    <%
        
      Do while not rsApproverGroups.EOF
       
        response.write "<TR>"
        response.write "<TD BGCOLOR=""" & RegionColor(rsApproverGroups("Region")) & """ ALIGN=""LEFT"" CLASS=Medium VALIGN=MIDDLE>"
        response.write rsApproverGroups("Description")
        response.write "</TD>"
    
        rsApproverNames.MoveFirst
        
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Medium VALIGN=MIDDLE>"
                    %>
                    <select language="JavaScript" onchange="window.location.href='site_utility.asp?Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Toggle=True&Group_ID=<%=rsApproverGroups("ID")%>&Approver_ID='+this.options[this.selectedIndex].value"
                        name="Approver_ID" id="Select1">
                        <%                        
  
        response.write "<OPTION VALUE=""0"" CLASS=NavLeftHighlight1>Select from list</OPTION>" & vbCrLf
        
        Do while not rsApproverNames.EOF    
          response.write "<OPTION CLASS=Region" & rsApproverNames("Region") & "NavMedium VALUE=""" & rsApproverNames("ID") & """"
          if isnumeric(rsApproverGroups("Approver_ID")) then
            if CLng(rsApproverNames("ID")) = CLng(rsApproverGroups("Approver_ID")) then
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
                </table>
            </td>
        </tr>
    </table>
</form>
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
    
    SQL = "SELECT Approvers_Account.* FROM Approvers_Account WHERE Approvers_Account.Site_ID=" & CLng(Site_ID) & " ORDER BY Approvers_Account.Order_Num, Approvers_Account.Description"
    
    Set rsApproverGroups = Server.CreateObject("ADODB.Recordset")
    rsApproverGroups.Open SQL, conn, 3, 3    
    
    if rsApproverGroups.EOF then
      response.write "There are no Account Administrator Groups established for this site.<BR><BR>"
      TableOn = false
    end if
   
    SQL = "Select UserData.* FROM UserData WHERE UserData.Site_ID=" & CLng(Site_ID) & " AND (UserData.Subgroups LIKE '%account%' OR UserData.Subgroups LIKE '%administrator%') ORDER BY UserData.LastName"
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
<form name="Dummy-6" id="Form2">
    <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#666666"
        bgcolor="#666666" id="Table9">
        <tr>
            <td>
                <table cellpadding="4" cellspacing="1" border="0" width="100%" id="Table10">
                    <tr>
                        <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                            Region</td>
                        <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                            Account Administrator&acute; Name</td>
                    </tr>
                    <%
        
      Do while not rsApproverGroups.EOF
       
        response.write "<TR>"
        response.write "<TD BGCOLOR=""" & RegionColor(rsApproverGroups("Region")) & """ ALIGN=""LEFT"" CLASS=Medium VALIGN=MIDDLE>"
        response.write rsApproverGroups("Description")
        response.write "</TD>"
    
        rsApproverNames.MoveFirst
        
        response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Medium VALIGN=MIDDLE>"
                    %>
                    <select language="JavaScript" onchange="window.location.href='site_utility.asp?Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Toggle=True&Group_ID=<%=rsApproverGroups("ID")%>&Approver_ID='+this.options[this.selectedIndex].value"
                        name="Approver_ID" id="Select2">
                        <%                        
  
        response.write "<OPTION VALUE=""0"" CLASS=NavLeftHighlight1>Select from list</OPTION>" & vbCrLf
        
        Do while not rsApproverNames.EOF    
          response.write "<OPTION CLASS=Region" & rsApproverNames("Region") & "NavMedium VALUE=""" & rsApproverNames("ID") & """"
          if isnumeric(rsApproverGroups("Approver_ID")) then
            if CLng(rsApproverNames("ID")) = CLng(rsApproverGroups("Approver_ID")) then
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
                </table>
            </td>
        </tr>
    </table>
</form>
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
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#666666"
    bgcolor="#666666" id="Table11">
    <tr>
        <td>
            <table cellpadding="4" cellspacing="1" border="0" width="100%" id="Table12">
                <tr>
                    <td bgcolor="Red" align="CENTER" class="SmallBoldWhite">
                        Action</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        ID</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Users Name</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Company</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        City</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        State</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Country</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Phone Number</td>
                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                        Groups</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Account Status</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Last Logon</td>
                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                        Account Expiration</td>
                </tr>
                <%
     end if
     
     Do while not rsUser.EOF
                %>
                <tr>
                    <td bgcolor="Silver" align="CENTER" class="Small">
                        <%if Admin_Access = 9 or (Utility_ID=20 or Utility_ID=22 or Utility_ID=23 ) then%>
                        <a href="account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>"
                            class="NavLeftHighlight1" onclick="location.href='account_edit.asp?Site_ID=<%=Site_ID%>&ID=edit_account&account_ID=<%=rsUser("ID")%>'"
                            value=" Edit ">&nbsp;&nbsp;Edit&nbsp;&nbsp;</a>
                        <%else%>
                        No Edit
                        <%end if%>
                    </td>
                    <td bgcolor="#FFFFFF" align="RIGHT" class="Small">
                        <% response.write rsUser("ID") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <%
                  response.write "<B>" & rsUser("LastName") & "</B>, "
                  response.write rsUser("FirstName")
                  if not isblank(rsUser("MiddleName")) then response.write " " & rsUser("MiddleName")
                  if not isblank(rsUser("Prefix")) then response.write " " & rsUser("Prefix") & ". "
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Company") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Business_City") %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write rsUser("Business_State") %>
                    </td>
                    <%              
                response.write "<TD BGCOLOR=""" & RegionColor(rsUser("Region")) & """ ALIGN=""LEFT"" CLASS=Small>"
                response.write rsUser("Business_Country")
                response.write "</TD>"
                    %>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <% response.write FormatPhone(rsUser("Business_Phone")) %>
                    </td>
                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                        <%
                  Call Write_SubGroups
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="CENTER" class="Small">
                        <%
                  Call Write_Account_Status
                        %>
                    </td>
                    <td bgcolor="#FFFFFF" align="CENTER" class="Small">
                        <%
                  Call Write_Last_Logon
                        %>
                    </td>
                    <%
                  Call Write_Expiration_Date
                    %>
        </td>
    </tr>
    <%
      rsUser.MoveNext
    
    loop
                         
    rsUser.close
    set User = nothing
  
    if TableOn then
    %>
</table>
</TD> </TR> </TABLE>
<%

    end if

  ' --------------------------------------------------------------------------------------
  ' List - Content / Event - All
  ' View - Thumbnail Requests
  ' --------------------------------------------------------------------------------------
  
  elseif (Utility_ID = 50 or Utility_ID = 51 or Utility_ID = 52 or Utility_ID=54 or Utility_ID = 60) and (Admin_Access = 2 or Admin_Access = 4 or Admin_Access >= 8) then
    if request("Category_ID") <> "" and Utility_ID <> 54 then
      Category_ID = CLng(request("Category_ID"))
    elseif Utility_ID = 54 then
      Category_ID = -1      
    else
      Category_ID = 0
    end if  
    
    if request("View") <> "" then
      View = CLng(request("View"))
    else
      View = 0
    end if
    
    if request("Campaign") <> "" then
      Campaign = CLng(request("Campaign"))
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
    
    if request("Submitted_By") <> "" then
      Submitted_By = request("Submitted_By")
    else
      Submitted_By = 0
    end if

    if request("FLanguage") <> "" then
      FLanguage = request("FLanguage")
    else
      FLanguage = ""
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
    
    if request("Subject") <> "" and request("Subject") <> "Today's News" and request("Subject") <> "*" then
    
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
    response.write "<TD CLASS=Small WIDTH=""50%"" ROWSPAN=7 VALIGN=TOP>"
        
    Call Nav_Border_Begin    
    Call Main_Menu
    Call Group_Code_Table
    Call Status_Colors
    Call Element_Names

    if Group_ID <> "" or Country <> "" or Submitted_By > 0 or FLanguage <> "" then
      response.write "&nbsp;&nbsp;"
%>
<a href="site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&View=<%=View%>&Group_ID=&Country=&Category_ID=<%=Category_ID%>&LDate=<%=LDate%>"
    onclick="window.location.href='site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Submitted_By=0&FLanguage=&View=<%=View%>&Group_ID=&Country=&Category_ID=<%=Category_ID%>&LDate=<%=LDate%>'">
    <span class="NavLeftHighlight1">&nbsp;Clear&nbsp;Filters&nbsp;</span></a>
<%
    end if
    
    Call Nav_Border_End
  
    
    if Utility_ID = 54 then
    
      FormName = "Subscription"
      response.write "<FORM NAME=""" & FormName & """>" & vbCrLf

      response.write "<SPAN CLASS=SmallBold>" & Translate("Subscription Email Date",Login_Language,conn) & "</SPAN>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
      response.write "<INPUT TYPE=""TEXT"" VALUE=""" & LDate & """ MAXLENGTH=""10"" SIZE=""7"" CLASS=Small NAME=""LDate"" "
      response.write "ONCHANGE=""window.location.href='site_utility.asp?ID=site_utility&Site_ID=" & Site_ID & "&Utility_ID=" & Utility_ID & "&View=" & View & "&LDate=' + this.value + '" & "&Group_ID=" & Group_ID & "&Country=" & Country & "&Category_ID=" & Category_ID & "&Subject=" & Subject & "'"">"
      response.write "&nbsp;&nbsp;"
%>
<a href="javascript:void()" language="JavaScript" onclick="window.dateField = document.<%=FormName%>.LDate;calendar = window.open('/sw-common/sw-calendar_picker.asp','cal','WIDTH=200,HEIGHT=250');return false">
    <img src="/images/calendar/calendar_icon.gif" border="0" height="21" align="TOP"></a>&nbsp;&nbsp;
<%      
      response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=BUTTON VALUE=""" & Translate("Go",Login_Language,conn) & """ "
      response.write "ONCLICK=""window.location.href='site_utility.asp?ID=site_utility&Site_ID=" & Site_ID & "&Utility_ID=" & Utility_ID & "&View=" & View & "&LDate=' + document." & FormName & ".LDate.value + '" & "&Group_ID=" & Group_ID & "&Country=" & Country & "&Category_ID=" & Category_ID & "&Subject=" & Subject & "'"">"
      response.write "</FORM>" & vbCrLf
      
      if Admin_Access >=8 then
        response.write "<BR>"
        response.write "<SPAN CLASS=SmallBold>" & Translate("Subscription Subject",Login_Language,conn) & "</SPAN>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        response.write "<SPAN CLASS=SmallBoldRed>" 
        if Subject = "" then
          response.write Translate("Today's News",Login_Language,conn)
        else
          response.write Subject
        end if
        response.write "</SPAN>"
        response.write "<BR>"
        response.write "<SPAN CLASS=SmallBold>" & Translate("Subscription Subject",Login_Language,conn) & " (" & Translate("Change",Login_Language,conn) & ")</SPAN>: "        
        response.write "<INPUT TYPE=""TEXT"" VALUE="""" MAXLENGTH=""60"" SIZE=""20"" CLASS=Small NAME=""Subject"" "
        response.write "ONFOCUS=""alert('Changing the Subject text of the Subscription Email will be in English only for all recipients.\n\nTo Change back to the default text, \'Today\'s News\', just enter an asterisk *.');"" "
        response.write "ONCHANGE=""window.location.href='site_utility.asp?ID=site_utility&Site_ID=" & Site_ID & "&Utility_ID=" & Utility_ID & "&View=" & View & "&Subject=' + this.value + '" & "&LDate=" & LDate & "&Group_ID=" & Group_ID & "Sort_By=" & Sort_By & "&Country=" & Country & "&Category_ID=" & Category_ID & "'"">"
      end if  
    end if
    
    response.write "</TD>"

    response.write "<TD CLASS=Small>"
    response.write "<SPAN CLASS=SmallBold>" & Translate("Category",Login_Language,conn) & ":</SPAN>"
    response.write "</TD>"

    response.write "<TD CLASS=Small>"
    
    if Sort_By = 4 or (Category_ID = -1 and (Sort_By <> 8 and Sort_By <> 9)) then 
        Sort_By = 4
    end if   
    
%>
<select class="Small" language="JavaScript" onchange="window.location.href='site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&View=<%=View%>&Submitted_By=<%=Submitted_By%>&FLanguage=<%=FLanguage%>&LDate=<%=LDate%>&Group_ID=<%=Group_ID%>&Country=<%=Country%>&Sort_By=<%=Sort_By%>&Category_ID='+this.options[this.selectedIndex].value;document.getElementById('Standby').style.visibility = 'visible';"
    name="Category_ID" id="Select3">
    <%

    SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CLng(Site_ID) & " AND Calendar_Category.Enabled=" & CLng(True) & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"

    Set rsCategory = Server.CreateObject("ADODB.Recordset")
    rsCategory.Open SQL, conn, 3, 3

    if Campaign <> 0 then
      response.write "<OPTION CLASS=Medium SELECTED VALUE="""">" & Translate("PI / C Listing",Login_Language,conn) & "</OPTION>" & vbCrLF
    end if
      
    do while not rsCategory.EOF

      select case rsCategory("Code")
        case 8000
          response.write "<OPTION Class=Region1"          
        case 8001
          response.write "<OPTION Class=Region2"          
        case else
          response.write "<OPTION Class=Medium"
      end select
      
      if CLng(request("Category_ID")) = rsCategory("ID") then
     	  response.write " SELECTED"                                  
        Category_Code = rsCategory("Code")
      end if   

  	  response.write " VALUE=""" & rsCategory("ID") & """>"
        
      SQL = "SELECT Count(*)AS Count FROM Calendar WHERE Code=" & rsCategory("Code") & " AND Site_ID=" & Site_ID
      if Group_ID <> "" then
        SQL = SQL & " AND SubGroups LIKE '%" & Group_ID & "%'"
      end if
      if Country <> "" then
        SQL = SQL & " AND (Country = 'none' OR Country LIKE '%0%' AND Country NOT LIKE '%" & Login_Country & "%' OR Country NOT LIKE '%0%' AND Country LIKE '%" & Login_Country & "%')"
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
    <select class="Small" language="JavaScript" onchange="window.location.href='site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Campaign=<%=Campaign%>&Submitted_By=<%=Submitted_By%>&FLanguage=<%=FLanguage%>&Category_ID=<%=Category_ID%>&LDate=<%=LDate%>&Group_ID=<%=Group_ID%>&Country=<%=Country%>&Sort_By=<%=Sort_By%>&View='+this.options[this.selectedIndex].value;document.getElementById('Standby').style.visibility = 'visible';"
        name="View" id="Select4">
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
    response.write ">"
    response.write Translate("PI/C (MAC)",Login_Language,conn) & " or " & Translate("Individual",Login_Language,conn)
    response.write "</OPTION>" & vbCrLf

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
        <select class="Small" language="JavaScript" onchange="window.location.href='/SW-Administrator/site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Campaign=<%=Campaign%>&Submitted_By=<%=Submitted_By%>&FLanguage=<%=FLanguage%>&LDate=<%=LDate%>&Category_ID=<%=Category_ID%>&View=<%=View%>&Country=<%=Country%>&Sort_By=<%=Sort_By%>&Group_ID='+this.options[this.selectedIndex].value;document.getElementById('Standby').style.visibility = 'visible';"
            name="Group_ID" id="Select5">
            <%
    SQL = "SELECT SubGroups.* FROM SubGroups WHERE SubGroups.Site_ID=" & CLng(Site_ID) & " AND SubGroups.Order_Num <> 99 AND SubGroups.Enabled=" & CLng(True) & " ORDER BY SubGroups.Order_Num"
    Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
    rsSubGroups.Open SQL, conn, 3, 3

    response.write "<OPTION Class=Small VALUE="""">" & Translate("No Filter",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""""></OPTION>"
                          
    Do while not rsSubGroups.EOF            
      if request("Group_ID") = rsSubGroups("Code") then
     	  response.write "<OPTION SELECTED VALUE=""" & rsSubGroups("Code") & """"
        if rsSubGroups("Enabled") = True then
          response.write ">+ "
        else
          response.write ">o "
        end if
        if instr(1,RestoreQuote(rsSubGroups("X_Description"))," (") > 0 then
          response.write mid(RestoreQuote(rsSubGroups("X_Description")),1, instr(1,RestoreQuote(rsSubGroups("X_Description"))," (") -1) 
        else
          response.write RestoreQuote(rsSubGroups("X_Description"))
        end if
        response.write "</OPTION>" & vbCrLf
      else
     	  response.write "<OPTION VALUE=""" & rsSubGroups("Code") & """"
        if rsSubGroups("Enabled") = True then
          response.write ">+ "                  
        else
          response.write ">o "
        end if
        if instr(1,RestoreQuote(rsSubGroups("X_Description"))," (") > 0 then
          response.write mid(RestoreQuote(rsSubGroups("X_Description")),1, instr(1,RestoreQuote(rsSubGroups("X_Description"))," (") -1) 
        else
          response.write RestoreQuote(rsSubGroups("X_Description"))
        end if
        response.write "</OPTION>" & vbCrLf
      end if
  	  rsSubGroups.MoveNext 
    loop
    
    rsSubGroups.close
    Set rsSubGroups = Nothing
    
    response.write "</SELECT>" & vbCrLf
    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf

    ' Filter by Country or no restrictions
    response.write "<TR>" & vbCrLf
    response.write "<TD Class=Small>" & vbCrLf
    response.write "<SPAN CLASS=SmallBold>" & Translate("Filter by Country",Login_Language,conn) & ": </SPAN>"
    response.write "</TD>" & vbCrLf

    response.write "<TD CLASS=Small>" & vbCrLf
            %>
            <select class="Small" language="JavaScript" onchange="window.location.href='/SW-Administrator/site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Campaign=<%=Campaign%>&Submitted_By=<%=Submitted_By%>&FLanguage=<%=FLanguage%>&LDate=<%=LDate%>&Category_ID=<%=Category_ID%>&View=<%=View%>&Group_ID=<%=Group_ID%>&Sort_By=<%=Sort_By%>&Country='+this.options[this.selectedIndex].value;document.getElementById('Standby').style.visibility = 'visible';"
                name="Country" id="Select6">
                <%

    response.write "<OPTION Class=Small VALUE="""">" & Translate("No Filter",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""""></OPTION>"  & vbCrLf
    response.write "<OPTION VALUE=""US"""
    if UCase(Country) = "US" then response.write " SELECTED"
    response.write ">United States</OPTION>"  & vbCrLf
    response.write "<OPTION VALUE=""UM"""
    if UCase(Country) = "UM" then response.write " SELECTED"
    response.write ">United States Minor Outlying Islands</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""""></OPTION>" & vbCrLf
    
    'Call Connect_FormDatabase
    Call DisplayCountrySelect(Country, "Small")
    'Call Disconnect_FormDatabase

    response.write "<SELECT>" & vbCrLf

    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf
    
    ' Filter by Language or no restrictions
    response.write "<TR>" & vbCrLf
    response.write "<TD Class=Small>" & vbCrLf
    response.write "<SPAN CLASS=SmallBold>" & Translate("Filter by Language",Login_Language,conn) & ": </SPAN>"
    response.write "</TD>" & vbCrLf

    response.write "<TD CLASS=Small>" & vbCrLf
                %>
                <select class="Small" language="JavaScript" onchange="window.location.href='/SW-Administrator/site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Campaign=<%=Campaign%>&Submitted_By=<%=Submitted_By%>&LDate=<%=LDate%>&Category_ID=<%=Category_ID%>&View=<%=View%>&Group_ID=<%=Group_ID%>&Sort_By=<%=Sort_By%>&FLanguage='+this.options[this.selectedIndex].value;document.getElementById('Standby').style.visibility = 'visible';"
                    name="FLanguage" id="Select7">
                    <%

    response.write "<OPTION Class=Small VALUE="""">" & Translate("No Filter",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""""></OPTION>" & vbCrLf

    SQL = "SELECT * FROM Language WHERE Language.Enable=" & CLng(True) & " ORDER BY Language.Sort"
    Set rsLanguage = Server.CreateObject("ADODB.Recordset")
    rsLanguage.Open SQL, conn, 3, 3
    
    Do while not rsLanguage.EOF
     	  response.write "<OPTION VALUE=""" & rsLanguage("Code") & """"
        if LCase(rsLanguage("Code")) = LCase(FLanguage) then
         response.write " SELECTED"
        end if 
        response.write " CLASS=Small"

        response.write ">" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>" & vbCrLf
  	  rsLanguage.MoveNext 
    loop
    
    rsLanguage.close
    set rsLanguage=nothing

    response.write "</SELECT>" & vbCrLf

    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf


    ' Filter by Owner
    
    response.write "<TR>" & vbCrLf
    response.write "<TD CLASS=Small>" & vbCrLf
    response.write "<SPAN CLASS=SmallBold>" & Translate("Filter by Owner",Login_Language,conn) & ": </SPAN>" & vbCrLf
    response.write "</TD>" & vbCrLf
    
    response.write "<TD CLASS=Small>" & vbCrLf
                    %>
                    <select class="Small" language="JavaScript" onchange="window.location.href='/SW-Administrator/site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&Campaign=<%=Campaign%>&LDate=<%=LDate%>&Category_ID=<%=Category_ID%>&View=<%=View%>&Country=<%=Country%>&Sort_By=<%=Sort_By%>&FLanguage=<%=FLanguage%>&Group_ID=<%=Group_ID%>&Submitted_By='+this.options[this.selectedIndex].value;document.getElementById('Standby').style.visibility = 'visible';"
                        name="Submitted_By" id="Select8">
                        <%
    SQL = "SELECT DISTINCT dbo.Calendar.Submitted_By as ID, dbo.UserData.FirstName, dbo.UserData.LastName " &_
          "FROM            dbo.Calendar LEFT OUTER JOIN " &_
          "                dbo.UserData ON dbo.Calendar.Site_ID = dbo.UserData.Site_ID AND dbo.Calendar.Submitted_By = dbo.UserData.ID " &_
          "WHERE           dbo.Calendar.Site_ID=" & Site_ID & "  AND  dbo.UserData.FirstName IS NOT NULL AND dbo.UserData.LastName IS NOT NULL "
          
    if Category_ID > 0 then
      SQL = SQL & " AND Category_ID=" & Category_ID & " "
    end if
          
    SQL = SQL & "ORDER BY dbo.UserData.LastName, dbo.UserData.FirstName"
    
    Set rsSubmitted_By = Server.CreateObject("ADODB.Recordset")
    rsSubmitted_By.Open SQL, conn, 3, 3

    response.write "<OPTION Class=Small VALUE="""">" & Translate("No Filter",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""""></OPTION>" & vbCrLf
                          
    do while not rsSubmitted_By.EOF            
      if CDbl(Submitted_By) = CDbl(rsSubmitted_By("ID")) then
     	  response.write "<OPTION SELECTED VALUE=""" & rsSubmitted_By("ID") & """>"
        response.write rsSubmitted_By("LastName") & ", " & rsSubmitted_By("FirstName")
        response.write "</OPTION>" & vbCrLf
      else
     	  response.write "<OPTION VALUE=""" & rsSubmitted_By("ID") & """>"
        response.write rsSubmitted_By("LastName") & ", " & rsSubmitted_By("FirstName")
        response.write "</OPTION>" & vbCrLf
      end if
  	  rsSubmitted_By.MoveNext 
    loop
    
    rsSubmitted_By.close
    Set rsSubmitted_By = Nothing
    
    response.write "</SELECT>" & vbCrLf
    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf

    ' Sort By
    response.write "<TR>" & vbCrLf
    response.write "<TD Class=Small>" & vbCrLf
    response.write "<SPAN CLASS=SmallBold>" & Translate("Sort By",Login_Language,conn) & ": </SPAN>" & vbCrLf
    response.write "</TD>" & vbCrLf

    response.write "<TD CLASS=Small>" & vbCrLf
    
                        %>
                        <select class="Small" language="JavaScript" onchange="window.location.href='/SW-Administrator/site_utility.asp?ID=site_utility&Site_ID=<%=Site_ID%>&Utility_ID=<%=Utility_ID%>&FLanguage=<%=FLanguage%>&Campaign=<%=Campaign%>&Submitted_By=<%=Submitted_By%>&LDate=<%=LDate%>&Category_ID=<%=Category_ID%>&View=<%=View%>&Group_ID=<%=Group_ID%>&Country=<%=Country%>&Sort_By='+this.options[this.selectedIndex].value;document.getElementById('Standby').style.visibility = 'visible';"
                            name="Sort_By" id="Select9">
                            <%
    
    if Category_ID <> -1 then
        response.write "<OPTION Class=Small VALUE="""">" & Translate("No Filter",Login_Language,conn) & "</OPTION>" & vbCrLf
        response.write "<OPTION VALUE=""""></OPTION>" & vbCrLf
        
        response.write "<OPTION VALUE=""6"""
        if Sort_By = 6 then response.write " SELECTED"
        response.write ">" & Translate("Title",Login_language,conn) & "</OPTION>" & vbCrLf

        response.write "<OPTION VALUE=""7"""
        if Sort_By = 7 then response.write " SELECTED"
        response.write ">" & Translate("Product",Login_Language,conn) & " / " & Translate("Title",Login_language,conn) & "</OPTION>" & vbCrLf

        response.write "<OPTION VALUE=""1"""
        if Sort_By = 1 then response.write " SELECTED"
        response.write ">" & Translate("Asset ID",Login_language,conn) & "</OPTION>" & vbCrLf
        
        response.write "<OPTION VALUE=""5"""
        if Sort_By = 5 then response.write " SELECTED"
        response.write ">" & Translate("Asset ID Parent / Clone",Login_language,conn) & "</OPTION>" & vbCrLf
        
        response.write "<OPTION VALUE=""2"""
        if Sort_By = 2 then response.write " SELECTED"
        response.write ">" & Translate("Item Number",Login_language,conn) & " + " & Translate("Revision",Login_Language,conn) & "</OPTION>" & vbCrLf
        
        response.write "<OPTION VALUE=""3"""
        if Sort_By = 3 then response.write " SELECTED"
        response.write ">" & Translate("Begin Date",Login_language,conn) & "</OPTION>" & vbCrLf
    end if
    
    response.write "<OPTION VALUE=""4"""                
    if Sort_By = 4 or (Category_ID = -1 and (Sort_By <> 8 and Sort_By <> 9)) then
        response.write " SELECTED"
        Sort_By = 4
    end if
    response.write ">" & Translate("Category",Login_language,conn) & "</OPTION>" & vbCrLf
    
    response.write "<OPTION VALUE=""8"""                
    if Sort_By = 8 then 
        response.write " SELECTED"
        Sort_By = 8
    end if
    response.write ">" & Translate("Category + SubCategory",Login_language,conn) & "</OPTION>" & vbCrLf

    response.write "<OPTION VALUE=""9"""                
    if Sort_By = 9 then 
        response.write " SELECTED"
        Sort_By = 9
    end if
    response.write ">" & Translate("Category + Product",Login_language,conn) & "</OPTION>" & vbCrLf
    response.write "<SELECT>" & vbCrLf
    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf
    
    if Utility_ID = 50 then        
        Response.Write "<TR>"
        Response.Write "<TD class=small colspan =3 align=right>"
        response.write "<a href=""javascript:ExportToExcel()""><b>" & Translate("Export to Excel",Login_language,conn) & "</b>" & vbCrLf
        Response.Write "</TD>"
        Response.Write "</TR>"	
    end if
    
    response.write "</TABLE>" & vbCrLf
	   Set rsUser = Server.CreateObject("ADODB.Recordset")    
    rsUser.cursorlocation = adUseClient 
    set objcmd=server.createobject("ADODB.Command")
	   set objcmd.ActiveConnection=conn
	   objcmd.CommandText="getAssetReport"

	   Set objPara1 = objcmd.CreateParameter("@siteid", 3, 1)
	   Set objPara2 = objcmd.CreateParameter("@language",200, 1,3)	
	   Set objPara3 = objcmd.CreateParameter("@categoryid", 3, 1)
	   Set objPara4 = objcmd.CreateParameter("@GroupId", 200, 1,2048)
	   Set objPara5 = objcmd.CreateParameter("@Country", 200, 1,10)
	   Set objPara6 = objcmd.CreateParameter("@Submitted_By", 3, 1)
	   Set objPara7 = objcmd.CreateParameter("@Campaign", 3, 1)
	   Set objPara8 = objcmd.CreateParameter("@SortBy", 3, 1)
    Set objPara9 = objcmd.CreateParameter("@Subscription", 3, 1)
    Set objPara10 = objcmd.CreateParameter("@ldate", 135, 1)
    Set objPara11 = objcmd.CreateParameter("@ThumbnailRequest", 3, 1)
    
	   objcmd.Parameters.append objPara1
	   objcmd.Parameters.append objPara2
	   objcmd.Parameters.append objPara3
	   objcmd.Parameters.append objPara4
	   objcmd.Parameters.append objPara5
	   objcmd.Parameters.append objPara6
	   objcmd.Parameters.append objPara7
	   objcmd.Parameters.append objPara8
    objcmd.Parameters.append objPara9
    objcmd.Parameters.append objPara10
    objcmd.Parameters.append objPara11

	   objPara1.value  = site_id

    if Campaign <> 0 then
      SQL = SQL  & "AND Calendar.Campaign=" & Campaign & " OR Calendar.ID=" & Campaign & " "
      objPara7.value = Campaign
    else  
      if Category_ID > 0 then
        objPara3.value = Category_ID 
      end if
    end if  

    if Utility_ID = 51 or Utility_ID = 52 then
      objPara4.value =  "view"
    elseif Utility_ID = 60 then
      objPara11.value = CLng(true)
      'SQL = SQL & "AND Calendar.Thumbnail_Request=" & CLng(True) & " "
    end if     

    if Submitted_By > 0 then
      objPara6.value =  Submitted_By 
    end if
    
    if FLanguage <> "" then
      objPara2.value = FLanguage 
    end if
   'response.write Group_ID
   'response.end     
    if Group_ID <> "" then
      objPara4.value =   Group_ID 
    end if
      
    if Country <> "" then
      'SQL = SQL & " AND (Calendar.Country = 'none' OR Calendar.Country NOT LIKE '0%' AND Calendar.Country LIKE '%" & Country & "%')" & " "   
      objPara5.value = Country 
    end if
	
    if Utility_ID = 54 then
      'SQL = SQL & " AND Calendar.Subscription=-1 AND Calendar.LDate='" & LDate & "' "
      objPara9.value = -1
      objPara10.value = LDate
    end if
   	
	   rsUser.open objcmd
	      
    'response.write Sort_By
    'response.end
    if Campaign <> 0 then
      rsUser.sort="Sort, Status, Category, Sub_Category, Product, Revision_Code Desc, ID, Lit_ACTIVE_FLAG "
    else
      select case Sort_By
        case 1  ' Asset ID
          rsUser.sort="Status, ID, Revision_Code Desc, Lit_Active_Flag "
        case 2  ' Item Number + Revision
          rsUser.sort="Item_Number,Lit_Active_Flag, Revision_Code Desc "
        case 3  ' Begin Date
          rsUser.sort="BDate, Lit_Active_Flag, Revision_Code Desc "        
        case 4  ' Category, Sub Category, Begin Date, ID
           rsUser.sort="Sort,Category,Product,Title, BDate, ID,Lit_Active_Flag, Revision_Code Desc "        
        case 5  ' Parent / Clone, Language
          ' Modified by zensar on 09-03-2007 as adding sort to select statement is returning duplicate rows.
          'SQL = SQL & "ORDER BY PC_Order, dbo.[Language].Sort"
          '>>>>>>>>>>>>
          rsUser.sort="PC_Order,languagesort"
        case 6  ' Title
          rsUser.sort="Title"
        case 7  ' Product Title
          rsUser.sort="Product, Title"
        case 8  ' Category, Sub Category
          rsUser.sort="sort, Category, Sub_Category, Product,Title, BDate, ID, Lit_Active_Flag, Revision_Code Desc "        
        case 9          
          rsUser.sort="sort, Category, Product, Sub_Category, Title, BDate, ID,Lit_Active_Flag, Revision_Code Desc "        
        case else
          rsUser.sort="Status, Sub_Category, Product, Item_Number, Lit_Active_Flag, Revision_Code Desc "
      end select    
    end if
  
    'Modified by zensar.Added functionality to extract the data into Excel file.
    if Utility_ID = 50 then
        'response.Write "<input type=""text"" name = ""txtQuery"" value=""" & SQL & """>"
        'response.Write "<input type=""text"" name = ""txtOperation"" value="""">"
    end if
	'Response.Write SQL
    
    'rsUser.Open SQL, conn, 3, 3
    

    if rsUser.EOF then
      select case Utility_ID
        case 54
          response.write "<BR><HR COLOR=Gray SIZE=3 WIDTH=""100%""><BR>" & Translate("There are no Content or Event Items scheduled for tonight's Subscription Service Email.",Login_Language,conn) & "<BR><BR>"
        case else
          response.write "<BR><HR COLOR=Gray SIZE=3 WIDTH=""100%""><BR>" & Translate("There are no Content or Event Items for Category. Please Select Another Category.",Login_Language,conn) & "<BR><BR>"
          TableOn = false
      end select
    else 
      Asset_ID_Old = 0
      Asset_Record_Count = 0
      'dim rsUserClone
      'dim iUserfieldCount
      'set rsUserClone=Server.CreateObject("adodb.recordset")
      'rsUserClone.CursorLocation=3
      'for iUserfieldCount=0 to rsUser.fields.count-1
'		 rsUserClone.Fields.append rsUser.fields(iUserfieldCount).name,200,500,64
 '     next
  '    rsUserClone.Open()
      'do while not rsUser.EOF
        'if rsUser("ID") <> Asset_ID_Old then
          'Asset_Record_Count = Asset_Record_Count + 1
          'Asset_ID_Old = rsUser("ID")
     '     rsUserClone.AddNew
   '       for iUserfieldCount=0 to rsUser.fields.count-1
	'		rsUserClone.Fields(iUserfieldCount).Value = rsUser.fields(iUserfieldCount)
	'	  next
	'	  rsUserClone.Update
        'end if
        'rsUser.MoveNext
      'loop
      'rsUser.close
     ' set rsUser = rsUserClone
      'Response.Write rsUser.recordcount
      Asset_Record_Count = rsUser.recordcount
      Record_Count = Asset_Record_Count
      
      Record_Pages  = Record_Count \ Record_Limit
      if Record_Count mod Record_Limit > 0 then Record_Pages = Record_Pages + 1

      Page_QS = "ID=site_utility&Site_ID=" & Site_ID & "&Language=" & Login_Language & "&NS=" & Top_Navigation & "&Utility_ID=" & Utility_ID & "&PCID=0" & _
      "&Campaign=" & Campaign & "&Submitted_By=" & Submitted_By & "&LDate="& LDate &"&Category_ID="& Category_ID  &"&View="& View &"&Group_ID="& Group_ID &"&Country=" & Country & "&Sort_By=" & Sort_By
      xPCID = 1
      Record_Number = 0
      rsUser.MoveFirst  
      
      if Record_Limit * (PCID - 1) > 0 then
        rsUser.Move (Record_Limit * (PCID - 1))
      end if  
      Record_Number = 1

      Old_ID        = 0

      TableOn = true
              
      with response
        .write "<SPAN CLASS=SMALL>" & Translate("Total Content or Event Items",Login_Language,conn) & ": " & Asset_Record_Count
        .write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Translate("Current Date/Time",Login_Language,conn) & ": " & Now() & " PST</SPAN><P>"
        
        Call RS_Page_Navigation
        .write "<br>"
        Call Table_Begin
        
        .write "<DIV ID=""ContentTableStart"" STYLE=""position: absolute;"">" & vbCrLf
        .write "</DIV>" & vbCrLf
        'RI#1028-gpd
        if Site_ID = 82 and Utility_ID = 50 then
        .write "    <div style=""overflow-y:scroll;overflow-x:scroll; width:960;overflow: -moz-scrollbars-horizontal;"" >" & vbCrLf
        end if
        .write "      <TABLE CELLPADDING=2 CELLSPACING=1 BORDER=0  WIDTH=""100%"" ID=""ContentTable"">"
        .write "        <TR ID=""ContentHeader1"">"
        .write "          <TD BGCOLOR=""Red"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Action",Login_Language,conn) & "</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("ID",Login_Language,conn) & "</TD>"
        
        Columns = 0
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("PID",Login_Language,conn) & "</TD>"
        Columns = Columns + 1
      end with
        
      if View = 4 then
        response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("MAC",Login_Language,conn) & "</TD>"
        Columns = Columns + 1
      end if  

      with response  
        'Modified by zensar on 08-03-2004
        if sort_by <> 8 then
            .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>" & Translate("Category",Login_Language,conn) & "</TD>"
        end if
        if sort_by <> 9 then
            .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>" & Translate("Product",Login_Language,conn) & "</TD>"
        end if
        'End
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>" & Translate("Title",Login_Language,conn) & "</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>L</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>A</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>Z</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>P</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>T</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>U</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>D</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>S</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>O</TD>"
        Columns = Columns + 12
      end with

             
      if Utility_ID = 50 or Utility_ID = 54 or Utility_ID = 60 then
        if View = 1 or View = 3 then
          response.write "<TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>" & Translate("Groups",Login_Language,conn) & "</TD>"
          Columns = Columns + 1
        end if  
        if View = 2 or View = 3 then 
          response.write "<TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>" & Translate("Country",Login_Language,conn) & "</TD>"
          Columns = Columns + 1
        end if
      end if  

      if Utility_ID = 50 or Utility_ID = 51 or Utility_ID = 52 or Utility_ID=54 or Utility_ID = 60 then
        with response
          .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Item",Login_Language,conn) & "</TD>"
          .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Rev",Login_Language,conn) & "</TD>"
          .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("CC",Login_Language,conn) & "</TD>"          
        end with
        Columns = Columns + 3
      end if
      
      if (Utility_ID = 51 or Utility_ID = 52 or Utility_ID = 60) then
        with response
          .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>" & Translate("Legacy",Login_Language,conn) & "</TD>"
        end with
        Columns = Columns + 1
      end if

      with response            
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("LNG",Login_Language,conn) & "</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Announce<BR>Date",Login_Language,conn) & "</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Begin",Login_Language,conn) & "<BR>" & Translate("Date",Login_Language,conn) & "</TD>"
        if Utility_ID = 54 then
          .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Subscription",Login_Language,conn) & "<BR>" & Translate("Time",Login_Language,conn) & "</TD>"
          .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Embargo<BR>Date",Login_Language,conn) & "</TD>"
        else
          .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("End<BR>Date",Login_Language,conn) & "</TD>"
          .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Embargo<BR>Date",Login_Language,conn) & "</TD>"
        end if
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Archive<BR>Date",Login_Language,conn) & "</TD>"
         'RI#1028-gpd
         if Site_ID = 82 and Utility_ID = 50 then
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>" & Translate("File<BR>Name",Login_Language,conn) & "</TD>"
        .write "          <TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>" & Translate("File<BR>Size",Login_Language,conn) & "</TD>"
        
         end if
        .write "    		</TR>"
      end with  
      if Site_ID = 82 then
      Columns = Columns + 8
      else
      Columns = Columns + 6
      end if
    end if
     
    Category_Old     = ""
    Sub_Category_Old = ""
    Product_Old      = ""
    Asset_ID_Old     = 0
    
    PC_Color         = "#FFFFFF"
    Old_PC_Order     = 0
    
    do while not rsUser.EOF and Record_Number <= Record_Limit
      if rsUser("ID") <> Asset_ID_Old then
        if rsUser("Category") <> Category_Old then
          response.write "<TR>"
          response.write "<TD BGCOLOR=""Silver"" CLASS=SmallBoldGold COLSPAN=2>"
          response.write "&nbsp;"
          response.write "</TD>"
  
          response.write "<TD TITLE=""Library or Calendar Category"" BGCOLOR="""

          response.write "Red""     CLASS=SmallBoldWhite "
          response.write "COLSPAN=" & Columns & ">"
  
          response.write rsUser("Category")
          response.write "</TD>"
          response.write "</TR>"
          Category_Old = rsUser("Category")
        end if
        
        if Sort_By = 8 then
          if trim(rsUser("Sub_Category")& "") <> "" then
            if trim(rsUser("Sub_Category")& "") <> trim(Sub_Category_Old)  then
              response.write "<TR>"
              response.write "<TD BGCOLOR=""Silver"" CLASS=SmallBoldGold COLSPAN=3>"
              response.write "&nbsp;"
              response.write "</TD>"
             
              response.write "<TD TITLE=""Category - " & rsUser("Category") & "&#13Library or Calendar Sub Category - " & rsUser("Sub_Category") & """ BGCOLOR="""
              response.write "#666666"" CLASS=SmallBoldWhite "

              response.write "COLSPAN=" & Columns & ">"
              if (rsUser("Sub_Category") & "" ="") then
                Sub_Category_Old = ""
              else
                response.write rsUser("Sub_Category") & ""
                Sub_Category_Old = rsUser("Sub_Category") & ""
              end if
              response.write "</TD>"
              response.write "</TR>"
            end if
          end if
        end if
        
        if Sort_By = 9 then
          if trim(rsUser("Product")& "") <> "" then
            if trim(rsUser("Product")& "") <> trim(Product_Old)  then
              response.write "<TR>"
              response.write "<TD BGCOLOR=""Silver"" CLASS=SmallBoldGold COLSPAN=3>"
              response.write "&nbsp;"
              response.write "</TD>"
      
              response.write "<TD TITLE=""Category - " & rsUser("Category") & "&#13Library or Calendar Product - " & rsUser("Product") & """ BGCOLOR="""
              response.write "#666666"" CLASS=SmallBoldWhite "
              response.write "COLSPAN=" & Columns & ">"
              if (rsUser("Product") & "" ="") then
                    Product_Old = ""
              else
                    response.write rsUser("Product") & ""
                    Product_Old = rsUser("Product") & ""
              end if
              response.write "</TD>"
              response.write "</TR>"
            end if
          end if
        end if
        
        ' Action
        response.write "<TR>"
        response.write "<TD BGCOLOR="""
        if Campaign = rsUser("ID") then
          response.write "#00CC00"
        elseif request("LastID") = CStr(rsUser("ID")) then
          response.write "#F5DEB3"
        else
          response.write "#FFFFFF"
        end if    
        response.write """ ALIGN=""CENTER"" CLASS=Small>"
        response.write "<A NAME=""ID" & rsUser("ID") & """></A>"  
        response.write "<A HREF=""/sw-administrator/Calendar_Edit.asp?ID=" & rsUser("ID") & "&Site_ID=" & Site_ID & """ Title=""Edit Asset"" CLASS=Navlefthighlight1>&nbsp;&nbsp;" & Translate("Edit",Login_Language,conn) & "&nbsp;&nbsp;</A>" & vbCrLf
        if (Category_Code >= 8000 and Category_Code <= 8999) and Campaign = 0 then
          response.write "<HR NOSHADE COLOR=""BLACK"" SIZE=1>"
          response.write "<A HREF=""/sw-administrator/Site_Utility.asp?ID=Site_Utility" & "&Campaign=" & rsUser("ID") & "&Site_ID=" & Site_ID & "&Utility_ID=" & Utility_ID & "&View=4" & """ Title=""List MAC Assets"" CLASS=Navlefthighlight1>&nbsp;&nbsp;" & Translate("List",Login_Language,conn) & "&nbsp;&nbsp;</A>" & vbCrLf
        end if
        response.write "</TD>"
      
        ' Status and ID
        Status = rsUser("Status")
        select case Status
          Case 1        
            response.write "<TD TITLE=""Asset ID Status=LIVE"" BGCOLOR=""#00CC00"" ALIGN=""CENTER"" CLASS=Small>"
          case 2
            response.write "<TD TITLE=""Asset ID Status=ARCHIVE"" BGCOLOR=""#AAAAFF"" ALIGN=""CENTER"" CLASS=Small>"
          case else
            response.write "<TD TITLE=""Asset ID Status=REVIEW"" BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
        end select
        response.write rsUser("ID")
        response.write "</TD>"
  
        ' Parent ID
        Status = rsUser("Clone")
        response.write "<TD TITLE=""Parent ID"" BGCOLOR="""
        
        if Sort_By = 5 then
          if Old_PC_Order <> rsUser("PC_Order") then
            if PC_Color="#FFFFFF" then
              PC_Color = "#EEEEEE"
            else
              PC_Color = "#FFFFFF"
            end if
            Old_PC_Order = rsUser("PC_Order")
          end if
        end if  
         
        response.write PC_Color & """ ALIGN=""CENTER"" CLASS=Small>"
        
        if not isblank(rsUser("Clone")) and rsUser("Clone") <> "0" then
          response.write rsUser("Clone")
        else
          response.write "&nbsp;"
        end if  
        response.write "</TD>"
  
        ' PI/C
        if View = 4 then
          response.write "<TD TITLE=""Content Grouping MAC or Individual"" BGCOLOR=""white"" Class="""
          select case CLng(rsUser("Content_Group"))
            case 1,2
              response.write "Region1"
            case 3,4
              response.write "Region2"
            case else
              response.write "Small"
          end select          
          
          response.write """ ALIGN=CENTER NOWRAP>"
          response.write "<SPAN CLASS=""Small"">"
          
          select case CLng(rsUser("Content_Group"))
            case 0
              response.write "I"
            case 1
              response.write "P + I"
            case 2
              response.write "P"
            case 3
              response.write "C + I"
            case 4
              response.write "C"
            case else
              response.write "&nbsp;"  
          end select
          if not isblank(rsUser("Campaign")) and rsUser("Campaign") <> 0 then
            response.write "<BR>"
            response.write "<A HREF=""/sw-administrator/Calendar_Edit.asp?ID=" & rsUser("Campaign") & "&Site_ID=" & Site_ID & """ Title=""Edit MAC"">" & rsUser("Campaign") & "</A>" & vbCrLf
            'response.write rsUser("Campaign")
          end if
          response.write "</SPAN>"  
          response.write "</TD>"
        end if      

        with response  
  
          ' Sub-Category
          if sort_by <> 8 then
              .write "<TD TITLE=""Sub-Category"" BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
              .write rsUser("Sub_Category")
              .write "</TD>"
          end if
          
          ' Product or Product Series
          if sort_by <> 9 then
              .write "<TD TITLE=""Product or Series"" BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
              .write rsUser("Product")
              .write "</TD>"
          end if

          ' Title
          .write "<TD TITLE=""Title / Owner"" BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
          .write rsUser("Title")
          if not isblank(rsUser("LastName")) then
            .write "<BR><SPAN CLASS=Smallest><SPAN STYLE=""COLOR=#666666"">" & Translate("Owner",Login_Language,conn) & ": "
            if not isblank(rsUser("FirstName")) then
              .write Mid(rsUser("FirstName"),1,1) & ". "
            end if
            .write rsUser("LastName") & "</SPAN></SPAN>"
          end if    
          
          .write "</TD>"
        end with
  
        ' Missing Assets - Warning
        Missing_Assets = False
        if (Category_Code < 8000 and Category_Code > 8999) then
          if isblank(rsUser("Link")) and isblank(rsUser("File_Name")) and isblank(rsUser("File_Name_POD")) then
            Missing_Assets = True
          elseif instr(1,LCase(rsUser("SubGroups")),"view") > 0 and (isblank(rsUser("File_Name")) and isblank(rsUser("File_Name_POD"))) then
            Missing_Assets = True
          elseif instr(1,LCase(rsUser("SubGroups")),"view") > 0 and (isblank(rsUser("File_Name")) or isblank(rsUser("File_Name_POD"))) then
            Missing_Assets = True
          end if
        end if
  
        ' Include File
        '      response.write "<TD TITLE=""Include File"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"      
        '      if not isblank(rsUser("Include")) then
        '        response.write "Y"
        '      else
        '        response.write "&nbsp;"  
        '      end if
        '      response.write "</TD>"                                     
        
        ' Link
        if not isblank(rsUser("Link")) then
          response.write "<TD TITLE=""Link to Web Page URL"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "Y"
        elseif Missing_Assets = True and (isblank(rsUser("File_Name")) and isblank(rsUser("File_Name_POD"))) then
          response.write "<TD TITLE=""Link to Web Page URL - Missing"" BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        else
          response.write "<TD TITLE=""Link to Web Page URL"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        end if
        response.write "</TD>"
  
        ' Asset File
        if not isblank(rsUser("File_Name")) then
          response.write "<TD TITLE=""Low Resolution Asset File"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"      
          response.write "Y"
        elseif Missing_Assets = True and isblank(rsUser("Link")) then
          response.write "<TD TITLE=""Low Resolution Asset File - Missing"" BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        else
          response.write "<TD TITLE=""Low Resolution Asset File"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        end if
        response.write "</TD>"
  
        ' Archive File
        if not isblank(rsUser("Archive_Name")) then
          response.write "<TD TITLE=""Low Resolution ZIP File"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"      
          response.write "Y"
        elseif Missing_Assets = True and isblank(rsUser("Link")) then
          response.write "<TD TITLE=""Low Resolution ZIP File - Missing"" BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        else
          response.write "<TD TITLE=""Low Resolution ZIP File"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        end if
        response.write "</TD>"
  
        ' POD Asset File
        if not isblank(rsUser("File_Name_POD")) then
          response.write "<TD TITLE=""Medium Resolution POD File"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"      
          response.write "Y"
        elseif not isnull(rsUser("Lit_POD")) then
          if CLng(rsUser("Lit_POD")) = CLng(True) then
            response.write "<TD TITLE=""Medium Resolution POD File - Missing"" BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
            response.write "&nbsp;"
          else  
            response.write "<TD TITLE=""Medium Resolution POD File"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
            response.write "&nbsp;"
          end if
        else
          response.write "<TD TITLE=""Medium Resolution POD File"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        end if
        response.write "</TD>"
  
        ' Thumbnail
        if not isblank(rsUser("Thumbnail")) then
          response.write "<TD TITLE=""Thumbnail File"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"            
          response.write "Y"
        elseif Missing_Assets = True then
          response.write "<TD TITLE=""Thumbnail File - Missing"" BGCOLOR=""#FFFFCC"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        elseif Missing_Assets = False and isblank(rsUser("Link")) then
          response.write "<TD TITLE=""Thumbnail File - Missing"" BGCOLOR=""#FFFFCC"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        else
          response.write "<TD TITLE=""Thumbnail File"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"  
        end if  
        response.write "</TD>"
  
        ' Electronic Email Fulfillment Viewable
        if     instr(1,rsUser("SubGroups"),"view") > 0 and (Category_Code < 8000 or Category_Code > 8999) then
          response.write "<TD TITLE=""Email / POD Fulfillment - Enabled"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"            
          response.write "Y"
        elseif instr(1,rsUser("SubGroups"),"view") = 0 and (Category_Code < 8000 or Category_Code > 8999) then
          response.write "<TD TITLE=""Email / POD Fulfillment - Disabled"" BGCOLOR=""#FFFFCC"" ALIGN=""CENTER"" CLASS=Small>"            
          response.write "&nbsp;"
        else
          response.write "<TD TITLE=""Email / POD Fulfillment"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"            
          response.write "&nbsp;"
        end if
        response.write "</TD>"
        
        ' Digital Library Viewable
        if instr(1,rsUser("SubGroups"),"fedl") > 0 and (Category_Code < 8000 or Category_Code > 8999) then
          response.write "<TD TITLE=""Digital Library Fulfillment - Enabled"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"            
          response.write "Y"
        elseif instr(1,rsUser("SubGroups"),"fedl") = 0 and (Category_Code < 8000 or Category_Code > 8999) then
          response.write "<TD TITLE=""Digital Library Fulfillment - Disabled"" BGCOLOR=""#FFFFCC"" ALIGN=""CENTER"" CLASS=Small>"            
          response.write "&nbsp;"  
        else
          response.write "<TD TITLE=""Digital Library Fulfillment"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"            
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
          response.write "<TD TITLE=""Shopping Cart Status"" BGCOLOR=""#AAAAFF"" ALIGN=""CENTER"" CLASS=Small>"      
        
        elseif ((LCase(rsUser("Lit_Status")) = "active" and LCase(rsUser("Lit_Action")) = "complete" and LCase(rsUser("Lit_Status_Name")) = "final loaded") or _
                (LCase(rsUser("Lit_Status")) = "active" and LCase(rsUser("Lit_Action")) = "complete" and LCase(rsUser("Lit_Status_Name")) = "reprint")) then
  
           if instr(1,rsUser("SubGroups"),"shpcrt") > 0 then
             response.write "<TD TITLE=""Shopping Cart Status - Excluded"" BGCOLOR=""#FF8000"" ALIGN=""CENTER"" CLASS=Small>"      
             Missing_Item = "E"
           else
             Missing_Item = "Y"
             response.write "<TD TITLE=""Shopping Cart Status - Enabled"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"      
           end if  
  
        elseif LCase(rsUser("Lit_Status")) = "active" and LCase(rsUser("Lit_Action")) = "complete" and LCase(rsUser("Lit_Status_Name")) = "web only" then
          Missing_Item = "P"
          response.write "<TD TITLE=""Shopping Cart Status - Enabled"" BGCOLOR=""#F5DEB3"" ALIGN=""CENTER"" CLASS=Small>"
  
        ' Unknown Status
        elseif LCase(rsUser("Lit_Status")) = "active" and LCase(rsUser("Lit_Action")) <> "complete" and _
          LCase(rsUser("Lit_Action")) <> "n/a" and LCase(rsUser("Lit_Action")) <> "web only" then
  
          Missing_Item = "N"
          response.write "<TD TITLE=""Shopping Cart Status"" BGCOLOR=""#0099FF"" ALIGN=""CENTER"" CLASS=Small>"
  
        elseif not isblank(rsUser("Item_Number")) and not isblank(rsUser("Revision_Code")) and isblank("Lit_Item") then 
          Missing_Item = "?"
          response.write "<TD TITLE=""Shopping Cart Status"" BGCOLOR=""#FFFF00"" ALIGN=""CENTER"" CLASS=Small>"
  
        else
          response.write "<TD TITLE=""Shopping Cart Status"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        end if
        
        if Missing_Item <> "Y" and Missing_Item <> "" then
          response.write "<A HREF=""JavaScript:void(0);"" ONCLICK=""var MyPop2 = window.open('/sw-administrator/Calendar_Element_Names.asp','MyPop2','fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,directories=no,location=no,width=250,height=360,left=600,top=200'); MyPop2.focus(); return false;"" CLASS=Small>"
          response.write Missing_Item
          response.write "</A>"
        else
          response.write Missing_Item
        end if
  
        response.write "</TD>"
        
        ' Oracle Status
        if not isblank(rsUser("Status_Comment")) and not isblank(rsUser("Item_Number")) then
          response.write "<TD TITLE=""Oracle Deliverable Status - Problem"" BGCOLOR=""#FF0000"" ALIGN=""CENTER"" CLASS=Small>"      
          response.write "?"
        elseif not isblank(rsUser("Status_Comment")) and isblank(rsUser("Item_Number")) then
          response.write "<TD TITLE=""Asset Container Status - Problem"" BGCOLOR=""" & RegionColor(3) & """ ALIGN=""CENTER"" CLASS=Small>"      
          response.write "?"
        else
          response.write "<TD TITLE=""Asset Container Status"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        end if
        response.write "</TD>"
  
        if Utility_ID = 50 or Utility_ID = 54 or Utility_ID = 60 then
          ' Groups Allowed to View
          if View = 1 or View = 3 then
            response.write "<TD TITLE=""Groups Allowed to View Asset"" BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
            if Group_ID <> "" then
              response.write Highlight_Keyword(replace(rsUser("SubGroups"),"view, ",""),Group_ID, "#FF0000")          
            else
              response.write replace(rsUser("SubGroups"),"view, ","")
            end if  
            response.write "</TD>"
          end if  
          if View = 2 or View = 3 then  
            response.write "<TD TITLE=""Country Restrictions"" BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small>"
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
            if rsUser("Lit_Active_Flag") = "-1" or isblank(rsUser("Lit_Active_Flag")) then
              .write "<TD TITLE=""Oracle Item Number"" BGCOLOR="""
              .write "#FFFFFF"
            else
              .write "<TD TITLE=""Oracle Item Number - Not Active"" BGCOLOR="""            
              .write "#FFFF00"
            end if
            .write """ ALIGN=""CENTER"" CLASS=Small>"
            .write rsUser("Item_Number")
            .write "</TD>"
          end with
  
          ' Item Reference Revision Code
          with response
            if LCase(rsUser("Revision_Code")) <> LCase(rsUser("Lit_Revision")) then
            .write "<TD TITLE=""Oracle Revision - Mismatch"" BGCOLOR="""
            .write "#FFFF00"
            elseif rsUser("Lit_Active_Flag") <> "-1" and not isblank(rsUser("Lit_Active_Flag")) then              
              .write "<TD TITLE=""Oracle Revision - Not Active"" BGCOLOR="""
              .write "#FFFF00"            
            else
              .write "<TD TITLE=""Oracle Revision"" BGCOLOR="""
              .write "#FFFFFF"
            end if
            .write """ ALIGN=""CENTER"" CLASS=Small>"
            .write rsUser("Revision_Code")
            .write "</TD>"
          end with
  
          ' Cost Center
          with response
            .write "<TD TITLE=""Oracle Cost Center"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
            if rsUser("Cost_Center") > 0 then
              .write rsUser("Cost_Center")
            else
              .write "&nbsp;"
            end if
            .write "</TD>"
          end with
  
        end if
  
        if Utility_ID = 51 or Utility_ID = 52 or Utility_ID = 60 then      
          with response
            .write "<TD TITLE=""Legacy Item Number"" BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"" CLASS=Small>"
            .write rsUser("Item_Number_2")
            .write "</TD>"          
          end with
        end if        
  
        ' Language
        if UCase(rsUser("Language")) = "ENG" then
          response.write "<TD TITLE=""Language"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
        else  
          response.write "<TD TITLE=""Language"" BGCOLOR=""#CFCFCF"" ALIGN=""CENTER"" CLASS=Small>"
        end if
        
        response.write UCase(rsUser("Language"))
        response.write "</TD>"
  
        ' Pre-Announce Date
        response.write "<TD TITLE=""Pre-Announce Date"" BGCOLOR="""
        
        if (CDate(rsUser("LDate")) < CDate(rsUser("BDate"))) then
  
          select case rsUser("Status")
          
            case 0 ' Pending
              if (CDate(rsUser("LDate")) > Date()) or (CDate(rsUser("BDate")) > Date()) then
                response.write "Yellow"
              elseif (CDate(rsUser("BDate")) < Date()) then
                response.write "Orange"
              else  
                response.write "Orange"
              end if
              
            case 1 ' Live
              if CDate(rsUser("LDate")) > Date() then
                response.write "#99FF99"
              else
                response.write "#00CC00"
              end if
              
            case 2 ' Archive
              response.write "#AAAAFF"
              
            case else
              response.write "#FFFFFF"
              
          end select
          
          response.write """ ALIGN=""CENTER"" CLASS=Small>"
    
          if CLng(rsUser("Status")) = 1 and (CDate(rsUser("LDate")) > Date()) and (CDate(rsUser("LDate")) < CDate(rsUser("BDate"))) then
            response.write Translate("Go Live",Logon_Language,conn) & "<BR>"
          end if
          response.write Replace(FormatDate(0,rsUser("LDate")),"/","&nbsp;&nbsp;")
          
        else
        
          response.write "#FFFFFF"
          response.write """ ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        
        end if
  
        response.write "</TD>"
  
        ' Begin Date
        response.write "<TD TITLE=""Begin Date - GO LIVE"" BGCOLOR="""
        
        select case rsUser("Status")
        
          case 0 ' Pending
            if (CDate(rsUser("LDate")) > Date()) or (CDate(rsUser("BDate")) > Date()) then
              response.write "Yellow"
            elseif (CDate(rsUser("BDate")) < Date()) then
              response.write "Orange"
            else  
              response.write "Orange"
            end if
            
          case 1 ' Live
            if (CDate(rsUser("LDate")) >= Date()) or (CDate(rsUser("BDate")) >= Date()) then
              if (CDate(rsUser("LDate")) >= Date()) and (CDate(rsUser("BDate")) < Date()) then
                response.write "#99FF99"
              else
                response.write "#00CC00"
              end if  
            else
              response.write "#00CC00"
            end if
            
          case 2 ' Archive
            response.write "#AAAAFF"
            
          case else
            response.write "#FFFFFF"
            
        end select
        
        response.write """ ALIGN=""CENTER"" CLASS=Small>"
  
        if CLng(rsUser("Status")) = 1 and (CDate(rsUser("LDate")) > Date()) and (CDate(rsUser("LDate")) = CDate(rsUser("BDate"))) then
          response.write Translate("Go Live",Logon_Language,conn) & "<BR>"
        end if
        response.write Replace(FormatDate(0,rsUser("BDate")),"/","&nbsp;&nbsp;")
        response.write "</TD>"
        
        ' Subscription Service Time
        if Utility_ID = 54 then
          response.write "<TD TITLE=""Subscription Service Date/Time"" BGCOLOR="""
          if CLng(rsUser("Subscription_Early")) = CLng(True) then
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
          if CLng(rsUser("Subscription_Early")) = CLng(True) then
            response.write "12:05 PM PST"
          else
            response.write "09:05 PM PST"
          end if
          response.write "</TD>"
        else
        
          ' End Date
          response.write "<TD TITLE=""End Date"" BGCOLOR="""
    
          if (CDate(rsUser("BDate")) < CDate(rsUser("XDate"))) then
          
            select case rsUser("Status")
            
              case 0 ' Pending
                if CDate(rsUser("XDate")) > Date() then
                  response.write "Yellow"
                elseif CDate(rsUser("XDate")) < Date() then
                  response.write "Orange"
                else  
                  response.write "Orange"
                end if
                
              case 1 ' Live
                if CDate(rsUser("EDate")) > Date() then
                  response.write "#99FF99"
                else
                  response.write "#00CC00"
                end if
    
                
              case 2 ' Archive
                response.write "#AAAAFF"
                
              case else
                response.write "#FFFFFF"
                
            end select
            
            response.write """ ALIGN=""CENTER"" CLASS=Small>"
            response.write Replace(FormatDate(0,rsUser("EDate")),"/","&nbsp;&nbsp;")
            response.write "</TD>"
            
          else  
            response.write "#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
            response.write "&nbsp;"
          end if
  
        end if
  
        ' Public Embargo Date
        response.write "<TD TITLE=""Public Embargo Date"" BGCOLOR="""
        
        if isdate(rsUser("PEDate")) then
  
          select case rsUser("Status")
          
            case 0 ' Pending
              if CDate(rsUser("PEDate")) < Date() then
                response.write "Yellow"
              elseif CDate(rsUser("PDate")) >= Date() then
                response.write "Orange"
              end if
              
            case 1 ' Live
              if CDate(rsUser("PEDate")) > Date() then
                response.write "#99FF99"
              else
                response.write "#00CC00"
              end if
              
            case 2 ' Archive
              response.write "#AAAAFF"
              
            case else
              response.write "#FFFFFF"
              
          end select
          
          response.write """ ALIGN=""CENTER"" CLASS=Small>"
          response.write Replace(FormatDate(0,rsUser("PEDate")),"/","&nbsp;&nbsp;")
          response.write "</TD>"
  
        else  
          response.write "#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
          response.write "&nbsp;"
        end if
        response.write "</TD>"
  
        ' Expiration Date
        response.write "<TD TITLE=""Expiration Date"" BGCOLOR="""
        
        select case rsUser("Status")
        
          case 0 ' Pending
          
            if CDate(rsUser("EDate")) = CDate(rsUser("XDate")) then
              response.write "#FFFFFF"
            elseif (CDate(rsUser("EDate")) < CDate(rsUser("XDate"))) and (CDate(rsUser("XDate")) <  Date()) then
              response.write "Orange"
            elseif (CDate(rsUser("EDate")) < CDate(rsUser("XDate"))) and (CDate(rsUser("XDate")) >= Date()) then
              response.write "Yellow"
            end if
            
          case 1 ' Live
          
            if CDate(rsUser("EDate")) = CDate(rsUser("XDate")) then
              response.write "#FFFFFF"
            elseif CDate(rsUser("XDate")) > Date() then
              response.write "#99FF99"
            else
              response.write "#00CC00"
            end if
            
          case 2 ' Archive
            response.write "#AAAAFF"
            
          case else
            response.write "#FFFFFF"
            
        end select
        
        response.write """ ALIGN=""CENTER"" CLASS=Small>"
  
        if CDate(rsUser("EDate")) = CDate(rsUser("XDate")) then
          if not isblank(rsUser("Item_Number")) then
            if isnumeric(rsUser("Item_Number")) and len(rsUser("Item_Number")) = 7 then
              if CDbl(rsUser("Item_Number")) < CDbl(9000000) then
                response.write "<FONT COLOR=""#666666"">" & Translate("Oracle", Login_Language,conn) & "</FONT>"
              else
                response.write Translate("Never",Login_Language,conn)
              end if
            else
              response.write Translate("Never",Login_Language,conn)
            end if
          else
             response.write Translate("Never",Login_Language,conn)
          end if
        else
          response.write Replace(FormatDate(0,rsUser("XDate")),"/","&nbsp;&nbsp;")
        end if
          
        response.write "</TD>"        
        'RI#1028-gpd
        if Site_ID = 82 and Utility_ID = 50 then
            Dim startIndex   
            startIndex = InStr(1,rsUser("File_Name"), "Asset/", 1)   
            startIndex = startIndex +5 
           
         response.write "<TD  TITLE=""File Name"" BGCOLOR="""
         response.write "white"
        response.write """ ALIGN=""LEFT"" CLASS=Smallest>"
       ' response.write rsUser("File_Name")
         response.write  "<textarea rows=""3"" CLASS=Small  cols=""25"" readonly=""readonly"">" 
        if startIndex >0 then
        Response.Write(Right(rsUser("File_Name"),Len(rsUser("File_Name"))- startIndex))        
        else
        response.write rsUser("File_Name")
        end if
       response.write "</textarea>"
        response.write "</TD>"
         response.write "<TD TITLE=""File Size"" BGCOLOR="""
         response.write "white"
        response.write """ ALIGN=""CENTER"" CLASS=Smallest>"
        response.write rsUser("File_size")
        response.write "</TD>"
        end if
        
        response.write "</TR>"
        
        Asset_ID_Old = rsUser("ID")
        Record_Number = Record_Number + 1
      end if
      
      '----------------------------------------------------
      rsUser.MoveNext
    loop
    rsUser.close
    set rsUser=nothing

    if TableOn then
      response.write "</TABLE>"
       'RI#1028-gpd
         if Site_ID = 82 and Utility_ID = 50 then
         response.write "</DIV>"
         end if
      Call Table_End
    end if
    
    response.write "<br>"
    Call RS_Page_Navigation 
    
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
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Begin Date",Login_Language,conn) & ": "
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
    'gpd rfc#1819
    if  Site_ID = 3 then
        response.write "<SPAN CLASS=SmallBoldGold>" & Translate("End Date",Login_Language,conn) & ": "
        if isblank(request.form("End_Date")) then
          End_Date = Date()
        elseif isdate(request.form("End_Date")) then
          End_Date = request.form("End_Date")
        else
          response.write Translate("Invalid End Date - Reseting to Today's Date",Login_Language,conn) & "<P>" & vbCrLf
          End_Date = Date()      
        end if
        response.write "</SPAN>" & vbCrLf
        response.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""End_Date"" VALUE=""" & End_Date & """ SIZE=""6"">" & vbCrLf
     end if
    'gpd rfc#1819
   if  Site_ID <> 3 then 
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Span",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
   end if
  'Below condition and code updated as 365 and 120 option needs to enable for fnet AMS Span
  if Site_ID = 82 then
      if isblank(request.form("Interval")) then
      Interval = 1
    elseif isnumeric(request.form("Interval")) then
      if FormatDateTime(Begin_Date) = FormatDateTime(Date()) and request.form("Interval") >= 0 then
        Interval = 1
      elseif request.form("Interval") > 365 then
        Interval = 365
      elseif request.form("Interval") < -365 then
        Interval = -365
      else  
        Interval = request.form("Interval")
      end if  
    end if
    
  else
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
    
    
  end if
  
  'gpd RFC#1819
  if  Site_ID <> 3 then 
    ' Interval in Days
        response.write "<SELECT CLASS=Small NAME=""Interval"">"

        
      if Site_ID = 82 then
        for x = 365 to -365 step -1
          select case x
            case 365,120,90,60,30,14,7          ' Future Days
              response.write "<OPTION VALUE=""" & x & """"
              if CLng(Interval) = x then response.write " SELECTED"
              response.write ">+" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
            case 1
              response.write "<OPTION CLASS=Region5NavSmall VALUE=""" & x & """"
              if CLng(Interval) = 1 then response.write " SELECTED"
              response.write ">+1" & " " & Translate("Day",Login_Language,conn) & "</OPTION>" & vbCrLf                
            case -365,-120,-90,-60,-30,-14,-7     ' Past Days
              response.write "<OPTION VALUE=""" & x & """"
              if CLng(Interval) = x then response.write " SELECTED"
              response.write ">-" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
          end select
        next
      else
      for x = 90 to -90 step -1
          select case x
            case 90,60,30,14,7          ' Future Days
              response.write "<OPTION VALUE=""" & x & """"
              if CLng(Interval) = x then response.write " SELECTED"
              response.write ">+" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
            case 1
              response.write "<OPTION CLASS=Region5NavSmall VALUE=""" & x & """"
              if CLng(Interval) = 1 then response.write " SELECTED"
              response.write ">+1" & " " & Translate("Day",Login_Language,conn) & "</OPTION>" & vbCrLf                
            case -90,-60,-30,-14,-7     ' Past Days
              response.write "<OPTION VALUE=""" & x & """"
              if CLng(Interval) = x then response.write " SELECTED"
              response.write ">-" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
          end select
        next
     
     end if
        response.write "</SELECT>" & vbCrLf
  end if
  
    ' Order Inquiry Filter

    if not isblank(request.form("OI")) and CLng(request.form("OI")) = CLng(True) then
      Order_Inquiry = CLng(True)
    else
      Order_Inquiry = CLng(False)
    end if
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Order",Login_Language,conn) & ": </SPAN>"
    response.write "<SELECT CLASS=Small NAME=""OI"">"
    response.write "<OPTION VALUE=""0"""
    if CLng(Order_Inquiry) = 0 then response.write " SELECTED"
    response.write ">" & Translate("Excluded",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""-1"""
    if CLng(Order_Inquiry) = -1 then response.write " SELECTED"
    response.write ">" & Translate("Included",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "</SELECT>" & vbCrLf

    ' Change Region / Site

    Call Change_Region
    if Admin_Access = 1 then
      Site_ID_Change = 0
    else
      Call Change_Site
    end if  
   response.write "<BR>"
    ' Submit Button
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE="" " & Translate("GO",Login_Language,conn) & " "">" & vbCrLf

    ' Reset Button
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""RESET"" VALUE="" " & Translate("Reset",Login_Language,conn) & " "" LANGUAGE=""Javascript"" ONCLICK=""this.form.Begin_Date.value='" & Date() & "';this.form.End_Date.value='" & Date() & "';this.form.Interval.options[0].selected=true;this.form.OI.options[0].selected=true;this.form.Region.options[0].selected=true;this.form.Item_Numbers.value='';"">" & vbCrLf
    
    response.write "<BR>"
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Asset ID or Item Number",Login_Language,conn) & ": "
    response.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""Item_Numbers"" VALUE=""" & item_numbers & """ SIZE=""30"">&nbsp;&nbsp;" & vbCrLf
    response.write "<SPAN CLASS=Small>(" & Translate("Separate multiple Asset ID and Item Numbers with a comma.",Login_Language,conn) & ")</SPAN>"
    Call Table_End
    response.write "</FORM>"
    
    ' Filter Specific WHERE clause
    'RFC#1819 gpd
    if  Site_ID = 3 then 
       if isblank(Item_Numbers) then
          SQLWhere    = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Activity.View_Time, 102), 102) <=  CONVERT(DATETIME, '" & End_Date & "', 102)) "
          SQLWhereLOS = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date, 102), 102) <=  CONVERT(DATETIME, '" & End_Date & "', 102)) "      
       else
          SQLWhere    = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "
          SQLWhereLOS = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "              
       end if
    else
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
    end if
    if Site_ID_Change > 0 then
      SQLWhere = SQLWhere & "AND dbo.Activity.Site_ID=" & Site_ID_Change & " "
    else
      SQLWhere = SQLWhere & " "
    end if  
      
    select case Region
      case 1, 2, 3
        'SQLWhere = SQLWhere & "AND dbo.UserData.Region=" & Region & " "
	SQLWhere = SQLWhere & "AND dbo.Activity.Region=" & Region & " "
    end select
   
    ' Site Clicks
   
    if isblank(Item_Numbers) then
      SQL = "SELECT  COUNT(dbo.Activity.View_Time) AS Clicks " &_
            "FROM         dbo.Activity LEFT OUTER JOIN " &_
            "             dbo.UserData ON dbo.Activity.Account_ID = dbo.UserData.ID "
    
      SQL = SQL & SQLWhere
    
      if CLng(Order_Inquiry) = CLng(False) then
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
     
    SQL = "SELECT   dbo.Calendar.Product AS Product, Calendar.Item_Number AS Item_Number, Calendar.Revision_Code, Calendar.Item_Number_2 AS Item_Number_2, dbo.Calendar.Title AS Title, dbo.Activity.Calendar_ID AS Asset_ID, dbo.Activity.Method AS Method, dbo.Activity.CID AS CID, dbo.Activity.Account_ID AS Account_ID, " &_
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
'response.write "length is " & len(item_number(x))
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
response.write "jj" & SQL
'response.end
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
                            <%Call Table_Begin
              response.write "<DIV ID=""ContentTableStart"" STYLE=""position: absolute;"">" & vbCrLf
              response.write "</DIV>" & vbCrLf
                            %>
                            <table cellpadding="2" cellspacing="1" border="0" width="100%" id="Table13">
                                <tr id="ContentHeader1">
                                    <td bgcolor="#666666" align="RIGHT" class="SmallBoldWhite">
                                        <%=Translate("Asset ID",Login_Language,conn)%>
                                    </td>
                                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                                        <%=Translate("Category",Login_Language,conn)%>
                                    </td>
                                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                                        <%=Translate("Sub Category",Login_Language,conn)%>
                                    </td>
                                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                                        <%=Translate("Product or Product Series",Login_Language,conn)%>
                                    </td>
                                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                                        <%=Translate("Title",Login_Language,conn)%>
                                    </td>
                                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                                        <%=Translate("Item Number",Login_Language,conn)%>
                                    </td>
                                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                                        <%=Translate("LAN",Login_Language,conn)%>
                                    </td>
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
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        OLV</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        OLD</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        OLS</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        SSV</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        SSD</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        SSS</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        OLL</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        OLG</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        EEF</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        EDL</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        LOS</td>
                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                        <%=Translate("Total",Login_Language,conn)%>
                                    </td>
                                </tr>
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
     
            
      'do while not rsActivity.EOF      

        Call Bypass_Assets
  
        if Clng(Old_Asset_ID) = Clng(rsActivity("Asset_ID")) then

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
            if rsActivity("Asset_ID")=28059 then
                       'response.write "Old_Category "& Old_Category 
            end if
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
            


             ' do while not rsLOS.EOF
                Method(14) = Method(14) + rsLOS("Quantity")
                Method(15) = Method(15) + rsLOS("Quantity")
			
                rsLOS.MoveNext
             ' loop
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

          'Call Display_Methods

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
        'RFC 1819 gpd
        Response.Flush()         
        rsActivity.MoveNext
              
      'loop

      if method(15) > 0 then

        rsActivity.MovePrevious
        response.write "New Cat is " & New_Category
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
  
        'Call Display_Methods

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
  
    if isblank(request("Region")) then
      response.write "<EMBED src=""/images/jeopardy.wav"" autostart=true loop=false volume=1 hidden=true><NOEMBED><BGSOUND src=""/images/jeopardy.wav.wav""></NOEMBED>"& vbCrLf
      response.flush
    end if
    
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
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Year Span",Login_Language,conn) & ":&nbsp;<INPUT TYPE=""RADIO"" NAME=""Count_Year"" VALUE=""0"""
    if Count_Year = 0 then response.write " CHECKED"
    response.write ">&nbsp;" & Year(Date()) & "&nbsp;&nbsp;" & vbCrLf
    response.write "<INPUT TYPE=""RADIO"" NAME=""Count_Year"" VALUE=""1"""
    if Count_Year = 1 then response.write " CHECKED"    
    response.write ">&nbsp;" & Year(Date()) & " + " & Year(Date()) -1 & "</SPAN>" & vbCrLf
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Region",Login_Language,conn) & ":&nbsp;"
    response.write "<SELECT NAME=""Region"" CLASS=Small>" & vbCrLf
    response.write "<OPTION VALUE=""0"""
    if isblank(request("Region")) or request("Region") = "0" then response.write " SELECTED"
    response.write ">" & Translate("Region",Login_Language,conn) & ": " & Translate("Worldwide",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION CLASS=Region1NavSmall VALUE=""1"""
    if request("Region") = "1" then response.write " SELECTED"
    response.write ">" & Translate("Region",Login_Language,conn) & ": " & Translate("United States",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION Class=Region2NavSmall VALUE=""2"""
    if request("Region") = "2" then response.write " SELECTED"
    response.write ">" & Translate("Region",Login_Language,conn) & ": " & Translate("Europe",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION CLASS=Region3NavSmall VALUE=""3"""
    if request("Region") = "3" then response.write " SELECTED"
    response.write ">" & Translate("Region",Login_Language,conn) & ": " & Translate("Intercon",Login_Language,conn) & "</OPTION>" & vbCrLf
    
    'conn.CommandTimeout = 240

    SQLCountry = "SELECT DISTINCT dbo.Activity.Country AS ISO2, dbo.Country.Name " &_
                 "FROM dbo.Activity LEFT OUTER JOIN dbo.Country ON dbo.Activity.Country = dbo.Country.Abbrev " &_
                 "WHERE (dbo.Country.Name IS NOT NULL) " &_
                 "ORDER BY dbo.Country.Name"

    Set rsCountry = Server.CreateObject("ADODB.Recordset")
    rsCountry.Open SQLCountry, conn, 3, 3
response.write sqlCountry

    do while not rsCountry.EOF
    
      response.write "<OPTION VALUE=""" & rsCountry("ISO2") & """"
      if request("Region") = rsCountry("ISO2") then response.write " SELECTED"
      response.write ">" & rsCountry("Name") & "</OPTION>" & vbCrLf
      
      rsCountry.MoveNext
    
    loop
    rsCountry.close
    set rsCountry = nothing
    
    response.write "</SELECT>" & vbCrLf
    
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE="" " & Translate("GO",Login_Language,conn) & " "">" & vbCrLf
    Call Nav_Border_End
    response.write "</FORM>"
    response.flush

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
            
            select case UCase(request("Region"))     ' Filter Results by Region

              case "0", ""  ' Worldwide no filter
              case "1"
                SQL = SQL & " (Region=0 OR Region=1) AND "
              case "2"
                SQL = SQL & " (Region=2) AND "
              case "3"
                SQL = SQL & " (Region=3) AND "
              case else
                if not isblank(request("Region")) then              
                  SQL = SQL & " (Country='" & UCase(request("Region")) & "') AND "
                end if  
            end select

            SQL = SQL & "(DATEPART(m, dbo.Activity.View_Time) = " & Summary_Month & ") AND (DATEPART(yyyy, dbo.Activity.View_Time) = " & Summary_Year & ") "
		'''SQL=SQL & " (dbo.Activity.View_Time >= '11/1/2009' and dbo.Activity.View_Time <= '11/30/2009') "

            'response.end
            select case z
              case 0      ' Order Query Criteria
                SQL = SQL & " AND Calendar_ID=101), "
              case 1      ' Order Inquiry Results
                SQL = SQL & " AND Calendar_ID=102), "


              case 2      ' Assets Accessed jigar
                SQL = SQL & " AND Calendar_ID>200), "
'response.write SQL
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

           
	
	'response.end
          Set rsActivity = conn.Execute(SQL)
        
          for z = 0 to 7        ' Activity Data
            Summary(z,Summary_Month) = rsActivity("Count_" & z)
            Summary(z,13) = Summary(z,13) + Summary(z,Summary_Month)                ' Year Totals for Class
          next

'		Response.write "0 to 7" & SQL



          for z = 8 to 16       ' Account Data

            Multiplier = 1
            
            select case z
            
              case 8            ' New Registrations
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                
                select case UCase(request("Region"))     ' Filter Results by Region
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') AND "
                    end if  
                end select

                SQL = SQL & "(DATEPART(m, UserData.Reg_Request_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, UserData.Reg_Request_Date) = " & Summary_Year & ") "

              case 9            ' New Registrations Pending
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CLng(True) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                  
                select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then                  
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') AND "
                    end if  
                end select

                SQL = SQL & "(DATEPART(m, UserData.Reg_Request_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, UserData.Reg_Request_Date) = " & Summary_Year & ") "

              case 10            ' New Accounts
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CLng(False) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                  
                select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then                  
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') AND "
                    end if  
                end select

                SQL = SQL & "(DATEPART(m, UserData.Reg_Approval_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, UserData.Reg_Approval_Date) = " & Summary_Year & ") "
                
              case 11            ' Expired Accounts
                Multiplier = -1
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CLng(False) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                  
                select case UCase(request("Region"))     ' Filter Results by Region
                  
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then                  
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') AND "
                    end if  
                end select

                if Summary_Month + 1 < 13 then
                  Last_Day_Month = DateAdd("d",-1,(Summary_Month + 1) & "/1/" & Summary_Year)
                else
                  Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))

                end if  
	
	'response.end
                SQL = SQL & " (Reg_Approval_Date >= '1/1/2001' AND Reg_Approval_Date <= '" & Last_Day_Month & "') AND (ExpirationDate >= '" & Summary_Month & "/1/" & Summary_Year & "' AND ExpirationDate <= '" & Last_Day_Month & "') "

              case 12            ' Never Logon Accounts
                Multiplier = -1
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CLng(False) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                  
                select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then                  
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') AND "
                    end if  
                end select

                if Summary_Month + 1 < 13 then
                  Last_Day_Month = DateAdd("d",-1,(Summary_Month + 1) & "/1/" & Summary_Year)
                else
                  Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
                end if  
                SQL = SQL & " (Reg_Approval_Date >= '" & Summary_Month & "/1/" & Summary_Year & "' AND Reg_Approval_Date <= '" & Last_Day_Month & "') "
                SQL = SQL & "AND Logon IS NULL or Logon='' "

              case 13            ' Active Accounts
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CLng(False) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                  
                select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0"
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') AND "
                    end if  
                end select
                

                if Summary_Month + 1 < 13 then
                  Last_Day_Month = DateAdd("d",-1,(Summary_Month + 1) & "/1/" & Summary_Year)
		
			'response.write Last_Day_Month
                else
                  Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
                end if  
                SQL = SQL & " (Reg_Approval_Date >= '1/1/2001' AND Reg_Approval_Date <= '" & Last_Day_Month & "') AND ExpirationDate >'" & Last_Day_Month & "' "

              case 14            ' Last Logons
                SQL = "SELECT Count(ID) AS Count_" & z & " FROM UserData WHERE NewFlag=" & CLng(False) & " AND "
                if Site_ID_Change > 0 then
                  SQL = SQL & "Site_ID=" & Site_ID_Change & " AND "
                end if
                  
                select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') AND "
                    end if  
                end select

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
                  
                select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then
                      SQL = SQL & " (Country='" & UCase(request("Region")) & "') AND "
                    end if  
                end select

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
                  
                select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then
                      SQL = SQL & " (Country='" & UCase(request("Region")) & "') AND "
                    end if                      
                end select

                if Summary_Month + 1 < 13 then
                  Last_Day_Month = DateAdd("d",-1,(Summary_Month + 1) & "/1/" & Summary_Year)
                else
                  Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
                end if  
                SQL = SQL & " (View_Time >= '" & Summary_Month & "/1/" & Summary_Year & "' AND View_Time <= '" & Last_Day_Month & "') "

            end select
            
            ' response.write SQL
            ' response.end
           ' Response.Write  SQL & vbCrLf & "<BR><BR>"

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
              "WHERE " & SQLWhere & "(DATEPART(m, Submit_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, Submit_Date) = " & Summary_Year & ")"

              select case UCase(request("Region"))     ' Filter Results by Region

                case "0", ""
                  SQL = SQL & ") "
                case "1"
                  SQL = SQL & " AND(Region=0 OR Region=1)) "
                case "2"
                  SQL = SQL & " AND(Region=2)) "
                case "3"
                  SQL = SQL & " AND(Region=3)) "
                case else
                  if not isblank(request("Region")) then                
                    SQL = SQL & " AND(Country='" & UCase(request("Region")) & "')) "
                  end if  
              end select

        if Summary_Month < 12 then SQL = SQL & ", "

      next
     ' Response.Write  SQL & vbCrLf & "<BR><BR>"
'response.end
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
              "WHERE " & SQLWhere & "(DATEPART(m, Submit_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, Submit_Date) = " & Summary_Year & ")"
              
              select case UCase(request("Region"))     ' Filter Results by Region
  
                case "0", ""
                  SQL = SQL & ") "
                case "1"
                  SQL = SQL & " AND(Region=0 OR Region=1)) "
                case "2"
                  SQL = SQL & " AND(Region=2)) "
                case "3"
                  SQL = SQL & " AND(Region=3)) "
                case else
                  if not isblank(request("Region")) then                
                    SQL = SQL & " AND(Country='" & UCase(request("Region")) & "')) "
                  end if  
              end select

        if Summary_Month < 12 then SQL = SQL & ", "

      next
 '     Response.Write  SQL & vbCrLf & "<BR><BR>"
'	response.end
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
              "WHERE " & SQLWhere & "(DATEPART(m, Submit_Date) = " & Summary_Month & ") AND (DATEPART(yyyy, Submit_Date) = " & Summary_Year & ")"
              
              select case UCase(request("Region"))     ' Filter Results by Region
  
                case "0", ""
                  SQL = SQL & ") "
                case "1"
                  SQL = SQL & " AND(Region=0 OR Region=1)) "
                case "2"
                  SQL = SQL & " AND(Region=2)) "
                case "3"
                  SQL = SQL & " AND(Region=3)) "
                case else
                  if not isblank(request("Region")) then                
                    SQL = SQL & " AND(Country='" & UCase(request("Region")) & "')) "
                  end if  
              end select
 
        if Summary_Month < 12 then SQL = SQL & ", "

      next
      
'     Response.Write  SQL & vbCrLf & "<BR><BR>"
'response.end
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
          if  isblank(Summary_Title(x)) then
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
                Summary_Key(x) = Translate("Total of all account sessions within a given month or session established when clicking on an asset item link within the subscription service email.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 17
                Summary_Title(x) = Translate("Shopping Cart - Orders",Login_Language,conn) & " *"
                Summary_Key(x) = Translate("Total of Literature Fulfilment Orders worldwide sent by the Portal to DCG.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 18
                Summary_Title(x) = Translate("Shopping Cart - Unique Item Numbers",Login_Language,conn) & " *"
                Summary_Key(x) = Translate("Total of Unique Item Numbers for Literature Fulfilment Orders worldwide sent by the Portal to DCG.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
              case 19
                Summary_Title(x) = Translate("Shopping Cart - Item Number Quantity",Login_Language,conn) & " *"
                Summary_Key(x) = Translate("Total Quantity Ordered of all Item Numbers for Literature Fulfilment Orders worldwide sent by the Portal to DCG.",Login_Language,conn) & " " & Translate("Total is cumulative through current month.",Login_Language,conn)
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
	''RI-648
if Site_ID_Change=11  then	
		if x=2 or x=8 or x=9 or x=10 or x=11 or x=12 or x=13 or x=14 or x=15 or x=16 then
			dim url1
			url1=""
			url1="SiteSummaryReport.asp?y=" & y & "&z=" & x & "&site_id=" & Site_ID & "&Region=" & request("Region") & "&Summary_Year=" & Summary_Year

			if FormatNumber(Summary(x,y),0)<> 0 then
		           response.write "<a href=" & url1 & ">" & FormatNumber(Summary(x,y),0) & "</a>"
			else
				response.write FormatNumber(Summary(x,y),0)
			end if
		else
			response.write FormatNumber(Summary(x,y),0)
		end if
else
			response.write FormatNumber(Summary(x,y),0)
end if
            response.write "</TD>" & vbCrLf
          next

''' jj add logic 648''
	'response.write "Summary_Year " & Summary_Year
	dim url
	url=""
	url="SiteSummaryReport.asp?z=" & x & "&site_id=" & Site_ID_Change & "&Region=" & request("Region") & "&Summary_Year=" & Summary_Year
	'response.write url

	if Site_ID_Change=11  then	
		if FormatNumber(Summary(x,13),0) <> 0  and (x=2 or x=8 or x=9 or x=10 or x=11 or x=12 or x=13 or x=14 or x=15 or x=16) then
	      	    response.write "<TD CLASS=SMALL ALIGN=""RIGHT"" BGCOLOR=""#99CCFF""><a href=" & url & ">" & FormatNumber(Summary(x,13),0) & "</a></TD>" & vbCrLf          
		else
			response.write "<TD CLASS=SMALL ALIGN=""RIGHT"" BGCOLOR=""#99CCFF"">" & FormatNumber(Summary(x,13),0) & "</TD>" & vbCrLf          		      	      
		end if
	else
          response.write "<TD CLASS=SMALL ALIGN=""RIGHT"" BGCOLOR=""#99CCFF"">" & FormatNumber(Summary(x,13),0) & "</TD>" & vbCrLf          
	end if
'''' end logic'''
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
    ' jjActivity by Item Number EEF/WWW
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
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Begin Date",Login_Language,conn) & ": "
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
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Span",Login_Langugage,conn) & ":</SPAN> " & vbCrLf

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
          if CLng(Interval) = x then response.write " SELECTED"
          response.write ">+" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
        case 1
          response.write "<OPTION CLASS=Region5NavSmall VALUE=""" & x & """"
          if CLng(Interval) = 1 then response.write " SELECTED"
          response.write ">+1" & " " & Translate("Day",Login_Language,conn) & "</OPTION>" & vbCrLf                
        case -90,-60,-30,-14,-7     ' Past Days
          response.write "<OPTION VALUE=""" & x & """"
          if CLng(Interval) = x then response.write " SELECTED"
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

    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Category",Login_Langugage,conn) & ":</SPAN>&nbsp;&nbsp;" & vbCrLf
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

    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Country",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
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

    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Local",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
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
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Group by",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
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
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Sort by",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
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
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Asset ID or Item Number",Login_Language,conn) & ": "
    response.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""Item_Numbers"" VALUE=""" & item_numbers & """ SIZE=""30"">&nbsp;&nbsp;" & vbCrLf
    response.write "<SPAN CLASS=Small>(" & Translate("Separate multiple Asset ID and Item Numbers with a comma.",Login_Language,conn) & ")</SPAN><BR>"
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
''       response.write "jj" &  SQL
	'response.end
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

      response.write "<DIV ID=""ContentTableStart"" STYLE=""position: absolute;"">" & vbCrLf
      response.write "</DIV>" & vbCrLf

      response.write "<TABLE CELLPADDING=""4"" CELLSPACING=""1"" BORDER=""0""  WIDTH=""100%"">" & vbCrLf
      response.write "<TR ID=""ContentHeader1"">" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Item Number",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Rev",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Asset ID",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""LEFT""   CLASS=SmallBoldWhite>" & Translate("Title",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Category",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Language",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Country",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Local",Login_Language,conn) & "</TD>" & vbCrLf
      if CLng(Group_By) <> 1 then
        response.write "<TD BGCOLOR=""#666666"" ALIGN=""LEFT"" CLASS=SmallBoldWhite>" & Translate("Path",Login_Language,conn) & "</TD>" & vbCrLf
      end if    
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Count",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "</TR>" & vbCrLf

      TableOn = True

      temp_old = rsActivity("Item_Number")
      if CLng(Group_By) <> 1 then
        temp_old = temp_old  & rsActivity("CMS_Site") & rsActivity("CMS_Path")
      end if  
      temp_cnt = 0
      temp_tot = 0
      'do while not rsActivity.EOF

        temp_new = rsActivity("Item_Number")
        if CLng(Group_By) <> 1 then
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
          if CLng(Group_By) <> 1 then
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

      'loop
      
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
      if CLng(Group_By) <> 1 then
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
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Begin Date",Login_Language,conn) & ": "
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
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Span",Login_Langugage,conn) & ":</SPAN> " & vbCrLf

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
          if CLng(Interval) = x then response.write " SELECTED"
          response.write ">+" & " " & ABS(x) & " " & Translate("Days",Login_Language,conn) & "</OPTION>" & vbCrLf
        case 1
          response.write "<OPTION CLASS=Region5NavSmall VALUE=""" & x & """"
          if CLng(Interval) = 1 then response.write " SELECTED"
          response.write ">+1" & " " & Translate("Day",Login_Language,conn) & "</OPTION>" & vbCrLf                
        case -90,-60,-30,-14,-7     ' Past Days
          response.write "<OPTION VALUE=""" & x & """"
          if CLng(Interval) = x then response.write " SELECTED"
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

    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Country",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
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
    
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Sort by",Login_Langugage,conn) & ":</SPAN> " & vbCrLf
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
     response.write "<BR>" & vbCrLf
    ' Submit Button
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE="" " & Translate("GO",Login_Language,conn) & " "">" & vbCrLf

    ' Reset Button
    response.write "&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""RESET"" NAME=""RESET"" VALUE="" " & Translate("Reset",Login_Language,conn) & " "" ONCLICK=""javascript: document.Activity_LOS.Item_Numbers.value=''; document.Activity_LOS.Begin_Date.value='" & Date() & "'; document.Activity_LOS.Country_Code[0].selected=true; document.Activity_LOS.Sort_By[0].selected=true; document.Activity_LOS.Interval[5].selected=true; return false;"">" & vbCrLf

    
    ' View CSV File
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""Button"" NAME=""CSV"" VALUE=""" & Translate("View CSV File",Login_Language,conn) & """ ONCLICK=""javascript: Ck_RecordCount();"">" & vbCrLf
    response.write "<BR>" & vbCrLf
    
    ' Individual Item Numbers

    
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Item or Order Number",Login_Language,conn) & ": "
    response.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""Item_Numbers"" VALUE=""" & item_numbers & """ SIZE=""30"">&nbsp;&nbsp;" & vbCrLf
    response.write "<SPAN CLASS=Small>(" & Translate("Separate multiple Asset ID and Item Numbers with a comma.",Login_Language,conn) & ")</SPAN><BR>"
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

                                <script language="JavaScript">
    function Ck_RecordCount() {    
      if (document.Activity_LOS.Counter.value != "0") {
        location.href="/sw-administrator/SW_Order_Inquiry_Literature_CSV.asp?Site_ID=<%=Site_ID%>&Site_Code=<%=Site_Code%>&Begin_Date=<%=Begin_Date%>&Interval=<%=Interval%>&Country_Code=<%=Country_Code%>&Sort_By=<%=Sort_By%>&Item_Numbers=<%=Item_Numbers%>&Language=<%=Login_Language%>";
      }
      else {
        alert("<%=Translate("There are no records to view based on your query criteria",Login_Language,conn)%>");
      }
      return false;
    }
                                </script>

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

      response.write "<DIV ID=""ContentTableStart"" STYLE=""position: absolute;"">" & vbCrLf
      response.write "</DIV>" & vbCrLf
      
      response.write "<TABLE CELLPADDING=""4"" CELLSPACING=""1"" BORDER=""0""  WIDTH=""100%"">" & vbCrLf
      response.write "<TR ID=""ContentHeader1"">" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Item Number",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Quantity",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Status",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Order Date",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Ship Date",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Days",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("CC",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""LEFT""   CLASS=SmallBoldWhite>" & Translate("Company",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""LEFT""   CLASS=SmallBoldWhite>" & Translate("Name",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Country",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Order Number",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "<TD BGCOLOR=""#666666"" ALIGN=""CENTER"" CLASS=SmallBoldWhite>" & Translate("Order",Login_Language,conn) & "</TD>" & vbCrLf      
      response.write "</TR>" & vbCrLf

      TableOn = True

      Old_Item_Number = ""
      Item_Number_Count = 0
      Item_Number_Total = 0
      
      'do while not rsActivity.EOF
      
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


      'loop
          
      rsActivity.close
      set rsActivity = nothing
      
      response.write "<TR>" & vbCrLf
      response.write "<TD CLASS=SmallBold BGCOLOR=""White"" ALIGN=""Left"">" & Translate("Period Total",Login_Language,conn) & "</TD>" & vbCrLf
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""SteelBlue"" ALIGN=""CENTER"""
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
  ' List - File Upload Status
  ' --------------------------------------------------------------------------------------

  elseif Utility_ID = 98 and Admin_Access >= 4 then

    if isblank(request("Interval")) then
      Interval = -7
    else
      Interval = CLng(request("Interval"))
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
      
    response.write "<BR>&nbsp;&nbsp;<SPAN CLASS=Small>Limit Listing to:</SPAN>&nbsp;"

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
        if CLng(rsStatus("Status")) = CLng(False) then
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
                                <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#666666"
                                    bgcolor="#666666" id="Table14">
                                    <tr>
                                        <td>
                                            <table cellpadding="4" cellspacing="1" border="0" width="100%" id="Table15">
                                                <tr>
                                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                                        Begin Time</td>
                                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                                        End Time</td>
                                                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                                                        Elapsed<br>
                                                        > 1 sec</td>
                                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                                        Uploaded<br>
                                                        By</td>
                                                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                                                        Source</td>
                                                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                                                        Destination</td>
                                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                                        File Size</td>
                                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                                        Status</td>
                                                    <td bgcolor="#666666" align="CENTER" class="SmallBoldWhite">
                                                        Error</td>
                                                    <td bgcolor="#666666" align="LEFT" class="SmallBoldWhite">
                                                        Error<br>
                                                        Description</td>
                                                </tr>
                                                <%
    end if
     
    Do while not rsStatus.EOF
                                                %>
                                                <tr>
                                                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                                                        <%=rsStatus("BTime")%>
                                                    </td>
                                                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                                                        <%
          if rsStatus("BTime") = rsStatus("ETime") then
            response.write "<SPAN STYLE=""Color:#BBBBBB"">"
            response.write rsStatus("ETime")
            response.write "</SPAN>"
          else
            response.write rsStatus("ETime")
          end if
                                                        %>
                                                    </td>
                                                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                                                        <%
          if rsStatus("BTime") = rsStatus("ETime") then
            response.write "<SPAN STYLE=""Color:#BBBBBB"">"
            response.write ConvertTime(DateDiff("s",rsStatus("BTime"),rsStatus("ETime")))
            response.write "</SPAN>"
          else
            response.write ConvertTime(DateDiff("s",rsStatus("BTime"),rsStatus("ETime")))
          end if
                                                        %>
                                                    </td>
                                                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                                                        <%=rsStatus("Firstname") & " " & rsStatus("Lastname")%>
                                                    </td>
                                                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                                                        <%=Replace(Replace(rsStatus("Path_Source"),"\","<BR>\")," ","&nbsp;")%>
                                                    </td>
                                                    <td bgcolor="#FFFFFF" align="LEFT" valign="TOP" class="Small">
                                                        <%
          if instr(1,LCase(rsStatus("Path_Destination")),"\extranet$\") > 0 then
            ThisSite = mid(rsStatus("Path_Destination"),instr(LCase(rsStatus("Path_Destination")),"\extranet$\") + 11)
          elseif instr(1,LCase(rsStatus("Path_Destination")),"\extranet\") > 0 then
            ThisSite = mid(rsStatus("Path_Destination"),instr(LCase(rsStatus("Path_Destination")),"\extranet\")  + 10)
          end if
          ThisSite = Replace(ThisSite," ","&nbsp;")
          ThisSite = "<SPAN CLASS=SmallRed>" & UCase(Mid(ThisSite,1,Instr(1,ThisSite,"\")-1))  & "</SPAN>" & Replace(Mid(ThisSite, Instr(1,ThisSite,"\")),"\","<BR>\") %>
                                                        <%                  
          response.write ThisSite                  
                                                        %>
                                                    </td>
                                                    <td bgcolor="#FFFFFF" align="Right" class="Small">
                                                        <%=FormatNumber((CDbl(rsStatus("Bytes"))/1024),0) & " KBytes" %>
                                                    </td>
                                                    <td bgcolor="#FFFFFF" align="CENTER" class="Small">
                                                        <%
          if CLng(rsStatus("Status")) = CLng(True) then
            response.write "Complete"
          else
            response.write "<SPAN CLASS=SmallRedBold>Failure</SPAN>"
          end if  
                                                        %>
                                                    </td>
                                                    <td bgcolor="#FFFFFF" align="CENTER" class="Small">
                                                        <%
          if rsStatus("Error_Number") = 0 then
            response.write "None"
          else
            response.write "<SPAN CLASS=SmallRedBold>" & rsStatus("Error_Number") & "</SPAN>"
          end if
                                                        %>
                                                    </td>
                                                    <td bgcolor="#FFFFFF" align="LEFT" class="Small">
                                                        <%=rsStatus("Error_Description")%>
                                                    </td>
                                                </tr>
                                                <%
      rsStatus.MoveNext
    
    loop
                         
    rsStatus.close
    set rsStatus=nothing
  
    if TableOn then
                                                %>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
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
  
end sub

' --------------------------------------------------------------------------------------


sub Status_Colors()

  response.write "&nbsp;&nbsp;&nbsp;<A HREF=""/sw-administrator/Calendar_Status_Colors.asp"" onclick=""var MyPop2 = window.open('/sw-administrator/Calendar_Status_Colors.asp','MyPop2','fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,directories=no,location=no,width=200,height=220,left=600,top=200'); MyPop2.focus(); return false;"" CLASS=NavLeftHighlight1>&nbsp;Status Colors&nbsp;</A>"
  
end sub

' --------------------------------------------------------------------------------------

sub Element_Names()

  response.write "&nbsp;&nbsp;&nbsp;<A HREF=""/sw-administrator/Calendar_Element_Names.asp"" onclick=""var MyPop2 = window.open('/sw-administrator/Calendar_Element_Names.asp','MyPop2','fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,directories=no,location=no,width=250,height=360,left=600,top=200'); MyPop2.focus(); return false;"" CLASS=NavLeftHighlight1>&nbsp;Status Codes&nbsp;</A>"
  
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
    response.write "&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>" & Translate("Region",Login_Language,conn) & " : </SPAN>" & vbCrLf
    response.write "<SELECT NAME=""Region"" CLASS=SMALL>" & vbCrLf
    response.write "<OPTION VALUE=""0"""
    if CLng(Region) = 0 then response.write " SELECTED"    
    response.write ">" & Translate("All",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""1"""
    if CLng(Region) = 1 then response.write " SELECTED"
    response.write ">" & Translate("USA",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""2"""
    if CLng(Region) = 2 then response.write " SELECTED"
    response.write ">" & Translate("Europe",Login_Language,conn) & "</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""3"""
    if CLng(Region) = 3 then response.write " SELECTED"
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
      SQLSite = "SELECT ID, Site_Description FROM Site WHERE Enabled=" & CLng(True) & " ORDER BY Site_Description"
      Set rsSite = Server.CreateObject("ADODB.Recordset")
      rsSite.Open SQLSite, conn, 3, 3
      response.write "&nbsp;&nbsp;&nbsp;&nbsp;<SPAN CLASS=SMALLBoldGold>" & Translate("Site",Login_Language,conn) & " : </SPAN>" & vbCrLf
      response.write "<SELECT NAME=""Site_ID_Change"" CLASS=SMALL>" & vbCrLf
      response.write "<OPTION CLASS=Small VALUE=""0"""
      if Site_ID_Change = 0 then response.write " SELECTED"
      response.write ">" & Translate("All Sites",Login_Language,conn) & "</OPTION>" & vbCrLf
      do while not rsSite.EOF
        if rsSite("ID") > 0 then
          response.write "<OPTION CLASS=Small VALUE=""" & rsSite("ID") & """"
          if CLng(Site_ID_Change) = CLng(rsSite("ID")) then response.write " SELECTED"
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
  do while not rsActivity.EOF AND CLng(Order_Inquiry) = CLng(False)
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

  select case CLng(rsActivity("Method"))
    case 0
      select case CLng(rsActivity("Account_ID"))
            case 1          ' Electronic Document Fulfillment EDF or WWW
                if isblank(rsActivity("CMS_Site")) then ' EDF // EEF
                    Method(11) = Method(11) + 1 
                else           
                    Method(13) = Method(13) + 1          ' WWW
                end if  
            case else
                    Method(0)  = Method(0) + 1 'OLV
            end select     
    case 2, 6 
      Method(2)  = Method(2)  + 1 'OLS
    
    case 7, 8
      Method(7)  = Method(7)  + 1
    case 9,10
      Method(9)  = Method(9)  + 1
    case 11, 12
      Method(11) = Method(11) + 1
    case else

      Method(CLng(rsActivity("Method"))) = Method(CLng(rsActivity("Method"))) + 1
  end select
  Method(15) = Method(15)     + 1      ' Increment Totals Counter  

end sub  

' --------------------------------------------------------------------------------------

sub Display_Methods
''response.write rsActivity("Site_ID")

'' Updated for 648 ''
if not isblank(item_numbers) then
	dim itemno
	itemno=rsActivity("Asset_ID")
end if

dim title
title=""
'title= Replace(rsActivity("Category_Title")," ","%20")
'if(rsActivity("Category_Title")) <> "" then
'	title=Server.URLEncode(rsActivity("Category_Title"))
'end if
title=Replace(New_Category," ","%20")

'response.end
dim url
url=""
url="AssetActivityDetail.asp?id=" & rsActivity("Asset_ID") & "&Begin_Date=" & Begin_Date & "&Interval=" & Interval & "&Site_ID=" & Site_ID_Change & "&Region=" & Region & "&Item_Number=" & Itemno & "&title='" & title & "'"



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
'' Updated for 648 ''
if Site_ID_Change = 11 then
          if x = 15 then response.write "<B><a href=" & url & ">"
          response.write FormatNumber(Method(x),0)
          if x = 15 then response.write "</a></B>"          
else
if x = 15 then response.write "<B>"
          response.write FormatNumber(Method(x),0)
          if x = 15 then response.write "</B>"          
end if
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

function convertTime(seconds)

  Dim lHrs
  Dim lMinutes
  Dim lSeconds
  
  lSeconds = Seconds
  
  lHrs = Int(lSeconds / 3600)
  lMinutes = (Int(lSeconds / 60)) - (lHrs * 60)
  lSeconds = Int(lSeconds Mod 60)
  
  Dim sAns
  
  If lSeconds = 60 Then
      lMinutes = lMinutes + 1
      lSeconds = 0
  End If
  
  If lMinutes = 60 Then
      lMinutes = 0
      lHrs = lHrs + 1
  End If
  
  if Len(lHrs) = 1 then
    lHrs = "0" & lHrs
  end if

  if Len(lMinutes) = 1 then
    lMinutes = "0" & lMinutes
  end if
  
  if Len(lSeconds) = 1 then
    lSeconds = "0" & lSeconds
  end if  
  
  ConvertTime = lHrs & ":" & lMinutes & ":" & lSeconds
  
end function

'--------------------------------------------------------------------------------------

                                %>

                                <script language="Javascript">

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

                                </script>

                                <!--#include virtual="/include/core_countries_select.inc"-->
                                <%

' --------------------------------------------------------------------------------------
' Record Set Page Navigation
' --------------------------------------------------------------------------------------
Sub RS_Page_Navigation
  Page_QS = "ID=site_utility&Site_ID=" & Site_ID & "&FLanguage=" & FLanguage & "&NS=" & Top_Navigation & "&Utility_ID=" & Utility_ID & _
  "&Campaign=" & Campaign & "&Submitted_By=" & Submitted_By & "&LDate="& LDate &"&Category_ID="& Category_ID  &"&View="& View &"&Group_ID="& Group_ID &"&Country=" & Country & "&Sort_By=" & Sort_By
  if PCID = 0 then PCID = 1

  ltEnabled = 0
  
  if Record_Pages > 1 then

    Call Nav_Border_Begin
    
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Page", Login_Language, conn) & ": &nbsp;</SPAN>"

  	if PCID = 1 then
  		Call RS_Page_Numbers
    	response.write "<A HREF=""Site_Utility.asp?" & Page_QS & "&PCID=" & PCID + 1 & """ CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Next Page", Alt_Language, conn) & """>"
        response.write "&nbsp;&gt;&gt;&nbsp;</A>"
        response.write "&nbsp;&nbsp;"
  	else
  		if PCID = Record_Pages then
            ltEnabled = 1
  		    response.write "<A HREF=""Site_Utility.asp?" & Page_QS & "&PCID=" & PCID - 1 & """ CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Previous Page", Alt_Language, conn) & """>"
            response.write "&nbsp;&lt;&lt;&nbsp</A>&nbsp;&nbsp;"
    	    Call RS_Page_Numbers
  		else
            ltEnabled = 1
  			response.write "<A HREF=""Site_Utility.asp?" & Page_QS & "&PCID=" & PCID - 1 &  """ CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Previous Page", Alt_Language, conn) & """>"
            response.write "&nbsp;&lt;&lt;&nbsp;</A>&nbsp;&nbsp;"
    		Call RS_Page_Numbers
  			response.write "<A HREF=""Site_Utility.asp?" & Page_QS & "&PCID=" & PCID + 1 &  """ CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Next Page", Alt_Language, conn) & """>"
            response.write "&nbsp;&gt;&gt;&nbsp;</A>"
  		end if
  	end if
    
    Call Nav_Border_End

  end if

End Sub

' --------------------------------------------------------------------------------------
' Record Set Page Numbers
' --------------------------------------------------------------------------------------

Sub RS_Page_Numbers

  iBreak = 0
  for i = 1 to Record_Pages
  	if i = PCID then
	  	response.write "<A HREF=""Site_Utility.asp?" & Page_QS & "&PCID=" & i & """ CLASS=NAVLEFTHIGHLIGHT1>"
      response.write "&nbsp;"
      if i < 10 then response.write "&nbsp;&nbsp;"
      response.write CStr(i) & "&nbsp;</A>"
      if iBreak = 19 - (ltEnabled) then
        response.write "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        iBreak = -1
        ltEnabled = 0
      else
        response.write "&nbsp;&nbsp;"
      end if  
  	else
			response.write "<A HREF=""Site_Utility.asp?" & Page_QS & "&PCID=" & i & """ CLASS=NavTopHighLight onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='NavTopHighLight'"">"
      response.write  "&nbsp;"
      if i < 10 then response.write "&nbsp;&nbsp;"
      response.write CStr(i) & "&nbsp;</A>"
      if iBreak = 19 - (ltEnabled) then
        response.write "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        iBreak = -1
        ltEnabled = 0        
      else
        response.write "&nbsp;&nbsp;"
      end if  
  	end if
    iBreak = iBreak + 1
  next

end sub
if cstr(Group_ID)= "" then
    Group_ID=0
end if
if cstr(submitted_by) = "" then
    submitted_by = 0
end if  
if cstr(Campaign) ="" then
    Campaign=0
end if
if cstr(Sort_By) ="" then
    Sort_By=0
end if
                                %>

                                <script language="JavaScript">
  <!--
    document.getElementById("Standby").style.visibility = "hidden";
    function ExportToExcel()
    {
        alert("../ExcelExport/GenerateRows.aspx?Site_Id=" + <%=Site_Id%> + "&CategoryId=" + <%=Category_ID%>  + "&GroupId=" + '<%=Group_ID%>' + "&Language=" + '<%=FLanguage%>' + "&Country=" + '<%=Country%>' + "&Submitted_By=" + <%=Submitted_By%> + "&Sort_By=" + <%=Sort_By%> + "&Campaign=" + <%=Campaign%>);
        window.open("../ExcelExport/GenerateRows.aspx?Site_Id=" + <%=Site_Id%> + "&CategoryId=" + <%=Category_ID%>  + "&GroupId=" + '<%=Group_ID%>' + "&Language=" + '<%=FLanguage%>' + "&Country=" + '<%=Country%>' + "&Submitted_By=" + <%=Submitted_By%> + "&Sort_By=" + <%=Sort_By%> + "&Campaign=" + <%=Campaign%>);
//	  window.open("http://author.dev.fluke.com/ExcelExport/WebForm1.aspx?Site_Id=" + <%=Site_Id%> + "&CategoryId=" + <%=Category_ID%>  + "&GroupId=" + <%=Group_ID%> + "&Language=" + '<%=FLanguage%>' + "&Country=" + '<%=Country%>' + "&Submitted_By=" + <%=submitted_by%> + "&Sort_By=" + <%=Sort_By%> + "&Campaign=" + <%=Campaign%>);

    }
  //-->
                                </script>
