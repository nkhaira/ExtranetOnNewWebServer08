<!--#include virtual="/include/functions_string.asp"-->
<%

Site_URL = "http://support.fluke.com/sw-administrator/site_utility.asp?ID=site_utility&Site_ID=3&Utility_ID=73"

if not isblank(Session("Logon_User")) and not isblank(Session("Site_ID")) then
  if Session("Site_ID") = 98 then
    response.redirect Site_URL
  end if
end if

response.redirect ("/Default.asp")  
%>
