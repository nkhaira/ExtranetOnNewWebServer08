<%
' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       1/1/2001
' Purpose:    Security Module used by SiteWide Auxilliary Applications
' --------------------------------------------------------------------------------------

' If user logon is NT Authentication and not through the SiteWide Logon screen, Session("Logon_User") = null
' so if this is the case, we need to check NT Authentication and update the session variable.

if isblank(Session("Logon_User")) and not isblank(Request.ServerVariables("LOGON_USER")) then
  Session("Logon_User") = Request.ServerVariables("LOGON_USER")
  do while instr(1,Login_Name,"\") > 0
    Session("Logon_User") = mid(Session("Logon_User"),instr(1,Session("Logon_User"),"\")+1)
  loop  
end if

' Override with Request
  
if not isblank(request("Site_ID")) _
   and isnumeric(request("Site_ID")) _
   and not isblank(Session("Logon_User")) then

   Site_ID            = request("Site_ID")
   Session("Site_ID") = request("Site_ID")

' Default  

elseif not isblank(Session("Site_ID")) _
   and isnumeric(session("Site_ID")) _
   and not isblank(Session("Logon_User")) then

   Site_ID = session("Site_ID")

' Kick user back to Find It
   
else
   Call Disconnect_SiteWide
   Session("ErrorString") = "<LI>" & Translate("Your session has expired.",Login_Language,conn) & " " & Translate("For your protection, you have been automatically logged off of your extranet site account.",Login_Language,conn) & "</LI><LI>" & Translate("To establish another session, please type in the site's code in the &quot;Name of the Site where you want to go&quot;, then click on [ Login ] or",Login_Language,conn) & "</LI><LI>" & Translate("Use the Site Search feature below.",Login_Language,conn) & "</LI>"
   response.redirect "/register/default.asp"

end if
%>