<!--#include virtual="/include/functions_string.asp"-->
<%

if isblank(request("Site_ID")) or isblank(request("NTLogin")) or isblank(request("Password")) or isblank(request("BackURL")) then

  response.write "Invalid Gateway Parameters."

else  

  %>
  <!--#include virtual="/connections/connection_SiteWide.asp"-->
  <%

  Call Connect_Sitewide

  SQL =  "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & request("NTLogin") & "' AND Password='" & request("Password") & "' AND Site_ID=" & request("Site_ID")
  Set rsUser = Server.CreateObject("ADODB.Recordset")
  rsUser.Open SQL, conn, 3, 3
  
  if not rsUser.EOF and CInt(rsUser("NewFlag")) = 0 then
  
    Session("Logon_User")     = request("NTLogin")
    Session("Password")       = request("Password")
    Session("Language")       = rsUser("Language")
    Session("BackURL")        = request("BackURL")
    Session("Site_ID")        = request("Site_ID")
    Session("Session_ID")     = Session.SessionID
		
    response.write "<HTML>" & vbCrLf
    response.write "<HEAD>" & vbCrLf
    response.write "<TITLE>Gateway Account Verified</TITLE>" & vbCrLf
    response.write "</HEAD>" & vbCrLf
    response.write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
    response.write "<FORM ACTION=""/sw-common/sw-order_inquiry_form.asp"" METHOD=""POST"">"
    response.write "</FORM>" & vbCrLf
    response.write "</BODY>" & vbCrLf
    response.write "</HTML>" & vbCrLf

  elseif not rsUser.EOF and CInt(rsUser("NewFlag")) = -1 then
    response.write "Invalid User Account."
  else
    response.write "Invalid User Gateway parameters."
  end if
  
  rsUser.close
  set rsUser = nothing
  Call Disconnect_Sitewide   
  
end if
%>
