<!--#include virtual="/include/functions_string.asp"-->
<%

if isblank(Session("Site_ID")) or isblank(Session("Logon_User")) then

  response.write "Invalid Gateway Parameters."

else  

  %>
  <!--#include virtual="/connections/connection_SiteWide.asp"-->
  <%

  Call Connect_Sitewide
  
  SQL = "SELECT Site_Code FROM Site WHERE ID=" & Session("Site_ID")
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3
  
  BackURL   = "/"
  GoToURL   = ""
  tsmFlag   = 0
  tsmRegion = ""
  SourceRecID = 0
  SourceSysID = ""
  
  if not rsSite.EOF then
    BackURL = BackURL & rsSite("Site_Code")
  end if
  
  rsSite.close
  set rsSite = nothing

  SQL =  "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & Session("Logon_User") & "' AND Site_ID=" & Session("Site_ID")
  
  Set rsUser = Server.CreateObject("ADODB.Recordset")
  rsUser.Open SQL, conn, 3, 3

  if not rsUser.EOF then
  
    if not isblank(rsUser("Fluke_ID")) and not isblank(rsUser("Business_System")) and instr(1,rsUser("SubGroups"),"wtbdis") > 0 then

      GoToURL = "/WTBDistributor/SW_ShowDistLocations.aspx"
      SourceRecID = rsUser("Fluke_ID")
      SourceSysID = rsUser("Business_System")      
      
    elseif instr(1,rsUser("SubGroups"),"wtbtsm") > 0 then

      GoToURL = "/WTBDistributor/SW_ShowDistLocations.aspx"            
      tsmFlag   = 1
      tsmRegion = rsUser("Region")

    elseif (instr(1,rsUser("SubGroups"),"account") > 0 and instr(1,rsUser("SubGroups"),"wtbadm")) or (instr(1,rsUser("SubGroups"),"site") > 0  and instr(1,rsUser("SubGroups"),"wtbadm")) or instr(1,rsUser("SubGroups"),"domain") > 0 then

      'GoToURL = "http://locator.fluke.com/WTBFlukeAdmin/NewDistributorEdit.aspx"    
      'Modified as per Fluke Release Item No - 72
      GoToURL = "http://locator.fluke.com/wtbflukeadmin/default.aspx"
      '>>>>>>>>>>
    end if
    
    if not isblank(GoToURL) then
      response.write "<HTML>" & vbCrLf
      response.write "<HEAD>" & vbCrLf
      response.write "<TITLE>Gateway Account Verified</TITLE>" & vbCrLf
      response.write "</HEAD>" & vbCrLf
      'response.write "<BODY>" & vbCrLf
      response.write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
      response.write "<FORM ACTION=""" & GoToURL & """ METHOD=""POST"">"
      response.write "<INPUT TYPE=HIDDEN NAME=Site_ID VALUE=""" & rsUser("Site_ID") & """>" & vbCrLf
      response.write "<INPUT TYPE=HIDDEN NAME=txtLoginUserID VALUE=""" & rsUser("ID") & """>" & vbCrLf
      response.write "<INPUT TYPE=HIDDEN NAME=txtSourceRecID VALUE=""" & SourceRecID & """>" & vbCrLf            
      response.write "<INPUT TYPE=HIDDEN NAME=txtSourceSysID VALUE=""" & SourceSysID & """>" & vbCrLf
      response.write "<INPUT TYPE=HIDDEN NAME=Language VALUE=""" & rsUser("Language") & """>" & vbCrLf
      response.write "<INPUT TYPE=HIDDEN NAME=txtTsmFlag VALUE=""" & tsmFlag & """>" & vbCrLf
      response.write "<INPUT TYPE=HIDDEN NAME=txtTsmRegion VALUE=""" & tsmRegion & """>" & vbCrLf                  
      response.write "<INPUT TYPE=HIDDEN NAME=BackURL VALUE=""" & BackURL & """>" & vbCrLf            
      response.write "<INPUT TYPE=HIDDEN NAME=txtContentWidth VALUE=""95"">" & vbCrLf      
      response.write "</FORM>" & vbCrLf
      response.write "</BODY>" & vbCrLf
      response.write "</HTML>" & vbCrLf
    else
      response.write "Unable to Proceed."
    end if  
  end if
  
  rsUser.close
  set rsUser = nothing
  Call Disconnect_Sitewide

end if
%>
