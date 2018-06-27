<%@  language="VBScript" codepage="65001" %>
<%
' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       2/1/2000
'             Find it @ Support.Fluke.com
' --------------------------------------------------------------------------------------
Dim ShowTranslation
Dim ErrorString
Dim Reg_Freeze
Dim strBackURL
Dim Border_Toggle
Border_Toggle = 0

strBackURL = request("BackURL")

Reg_Freeze = False

ShowTranslation = False

if Session("ShowTranslation") = True or request("Language") = "XON" then
  ShowTranslation = True
elseif Session("ShowTranslation") = False or request("Language") = "XOF" then
  ShowTranslation = False
end if  



%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Dim Site
Dim Site_Alias  ' If not null, then replace site_id with site_alias
Dim Site_URL
Dim Site_ID
Dim Site_Found
Dim Site_Enabled

Site_ID = 0

Site_Select = request.form("Site_Select")
if isblank(Site_Select) then
	Site_Select = request.querystring("Site_Select")
else
    if LCase(Site_Select) = "met-support" then
             session.Abandon()
             response.Redirect "http://us.flukecal.com/support/my-met-support"
     end if 
end if

Site = request.form("Site")
if isblank(Site) then
	Site = request.querystring("Site")
end if

if not isblank(Site_Select) and isblank(Site) then
  if Site_Select <> "manually" then
    Site = Site_Select
  end if  
end if  

Session.abandon

if ShowTranslation = True then
  Session("ShowTranslation") = True
elseif ShowTranslation = False then
  Session("ShowTranslation") = False 
end if

Call Connect_SiteWide




































response.write("all good...")
%>

