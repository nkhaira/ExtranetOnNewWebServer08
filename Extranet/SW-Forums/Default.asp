<%@ LANGUAGE="vbScript"%>

<%

Script_Debug = False

if request("Site_ID") <> "" and isnumeric(request("Site_ID")) then
  Site_ID            = request("Site_ID")
  Session("Site_ID") = request("Site_ID")  
elseif session("Site_ID") <> "" and isnumeric(session("Site_ID")) then
  Site_ID = session("Site_ID")
else
  response.redirect "http://" & Request("SERVER_NAME") & "/register/default.asp"
end if

Session("Site_id") = Site_ID

if request("Asset_ID") <> "" and isnumeric(request("Asset_ID")) then
  Asset_ID           = request("Asset_ID")
  Session("Asset_ID") = request("Asset_ID")  
elseif session("Asset_ID") <> "" and isnumeric(session("Asset_ID")) then
  Asset_ID = session("Asset_ID")
else
  response.redirect "http://" & Request("SERVER_NAME") & "/register/default.asp"
end if

Dim BackURL
BackURL = Session("BackURL")    

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

if isblank(Session("Logon_User")) and not isblank(Request.ServerVariables("LOGON_USER")) then
  Login_Name = Request.ServerVariables("LOGON_USER")
elseif not isblank(Session("LOGON_USER")) then
  Login_Name = Session("LOGON_USER")
else
  Login_Name = ""
end if

if not isblank(Login_Name) then
  
  do while instr(1,Login_Name,"\") > 0  %><%
    Login_Name = mid(Login_Name,instr(1,Login_Name,"\")+1) %><%
  loop  

  SQL = "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & Login_Name & "'"
  Set rsLogin = Server.CreateObject("ADODB.Recordset")
  rsLogin.Open SQL, conn, 3, 3

  if not rsLogin.EOF then
    Session ("Username")     = UCase(Login_Name)
    Session ("FullName")     = rsLogin("FirstName") & " " & rsLogin("LastName")
    Session ("EmailAddress") = rsLogin("EMail")
    Session ("Language")     = rsLogin("Language")
    Session ("User_id")      = rsLogin("ID")
    if instr(1,LCase(rsLogin("SubGroups")),"administrator") > 0 or instr(1,LCase(rsLogin("SubGroups")),"forum") > 0 then
      Session ("IsAdministrator") = CInt(True)
    else
      Session ("IsAdministrator") = CInt(False)
    end if
  end if

  rsLogin.close
  set rsLogin = nothing
  
else
  Call Disconnect_SiteWide
  response.redirect "/Register/Default.asp"
end if

SQL = "SELECT Calendar.* FROM Calendar WHERE Calendar.ID=" & Asset_ID
Set rsForum = Server.CreateObject("ADODB.Recordset")
rsForum.Open SQL, conn, 3, 3

if not rsForum.EOF then
  Site_Description_Text      = rsForum("Description")  
  Session("Forum_Title")     = rsForum("Title")
  Session("Forum_ID")        = rsForum("Forum_ID")
  Session("Forum_Moderated") = rsForum("Forum_Moderated")
  Forum_Moderator_ID         = rsForum("Forum_Moderator_ID")
end if
  
rsForum.close
set rsForum = nothing

' --------------------------------------------------------------------------------------
' Get Site Information
' --------------------------------------------------------------------------------------
%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%
Session("Site_Code")        = Site_Code
Session("Site_Description") = Site_Description

if not isblank(Logo) then
  Session("Logo") = "http://" & Request("SERVER_NAME") & Logo
else
  Session("Logo") = "http://" & Request("SERVER_NAME") & "/images/FlukeLogo3.gif"
end if    
  
' --------------------------------------------------------------------------------------
' Get Forum Moderator Info
' --------------------------------------------------------------------------------------

if CInt(Session("Forum_Moderated")) = CInt(True) and not isblank(Forum_Moderator_ID) then

  SQL = "SELECT * FROM UserData WHERE ID=" &   Forum_Moderator_ID
  Set rsModerator = Server.CreateObject("ADODB.Recordset")
  rsModerator.Open SQL, conn, 3, 3

  if not rsModerator.EOF then
    Site_Admin_Name   = rsModerator("FirstName") & " " & rsModerator("LastName")
    Site_Admin_Email  = rsModerator("Email")
  end if  
  
  rsModerator.close
  set rsModerator = nothing
  
end if
  
Session("Moderator_Name")  = Site_Admin_Name
Session("Moderator_Email") = Site_Admin_Email

' --------------------------------------------------------------------------------------
' Main Screen
' --------------------------------------------------------------------------------------

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Active Forum Postings",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Forum",Login_Language,conn) & " / " & Translate("Discussion Group",Login_Language,conn) & " - " & Translate("Home",Login_Language,conn) & "</FONT>"
 
Side_Navigation = False
Top_Navigation  = False
Content_Width   = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<!--#include file="SWEForums.asp" -->
<%

response.write "<H3>"
response.write Session("Forum_Title")
response.write "<BR><FONT CLASS=SmallBold>" & Translate("Current Messages",Login_Language,conn) & "</FONT>"
response.write "</H3>"

'response.write "<FONT CLASS=Medium>" & Site_Description_Text & "</FONT>"

response.write "<TABLE WIDTH=""100%"" BORDER=0>"
response.write "<TR>"
response.write "<TD ALIGN=LEFT WIDTH=""25%"">"
response.write "<FORM NAME=""BackURL"">"
'response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" VALUE=""" & Translate("Exit Forum",Login_Languge,conn) & """ LANGUAGE=""JavaScript"" OnClick=""window.location.href='" & BackURL & "'"">" & vbCrLf
response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" VALUE=""" & Translate("Exit Forum",Login_Languge,conn) & """ LANGUAGE=""JavaScript"" OnClick=""window.close()"">" & vbCrLf


response.write "</FORM>"
response.write "</TD>"
response.write "<TD WIDTH=""25%"" ALIGN=CENTER>"
Call NEW_POST_BUTTON_DISC
response.write "</TD>"
response.write "<TD WIDTH=""25%"" ALIGN=CENTER>"
response.write "<FORM NAME=""Archive"">"
response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" VALUE=""" & Translate("Archive",Login_Languge,conn) & """ LANGUAGE=""JavaScript"" OnClick=""window.location.href='/sw-forums/archive.asp'"">" & vbCrLf
response.write "</FORM>"
response.write "</TD>"
response.write "<TD ALIGN=RIGHT WIDTH=""25%"">"
Call SEARCH_FORM_DISC
response.write "</TD>"
response.write "</TR>"
response.write "</TABLE>"

response.write "<FONT CLASS=Medium>"
Response.write "<LI>" & Translate("Listed below are the current messages for the last 90 days.",Login_Language,conn) & "</LI>"
response.write "<LI>" & Translate("Click on the",Login_Language,conn) & " " & "<IMG SRC=""plus.gif"" TITLE=""Expand Message Thread"" BORDER=0 WIDTH=9 HEIGHT=9>" & " " & Translate("icon to expand message threads or,",Login_Language,conn) & " "
response.write Translate("click on the",Login_Language,conn) & " " & "<IMG SRC=""minus.gif"" TITLE=""Collapse Message Thread"" BORDER=0 WIDTH=9 HEIGHT=9>" & " " & Translate("icon to collapse message threads.",Login_Language,conn) & "</LI>"
response.write "<LI>" & Translate("Click on [Post New Message] to post a question to the Forum / Discussion Group.",Login_Language,conn) & "</LI>"
response.write "<LI>" & Translate("Click on [Archive] to check the for older postings related to this topic, or use Keyword [Search]",Login_Language,conn) & ".</LI>"
response.write "</FONT>"
response.write "<BR><BR>"
response.write "<HR NOSHADE COLOR=""#000000"">"

response.write "<BR>"
Call CURRENT_POSTS_THREADED_DHTML_DISC
response.write "<BR>"

response.write "<HR NOSHADE COLOR=""#000000"">"

Call CLEANUP_DISC

if Script_Debug then
  response.write "<BR>"
  response.write "Site ID:"  & session("Site_ID") & "<BR>"
  response.write "Asset ID:" & Session("Asset_ID") & "<BR>"
  response.write "BackURL:" & Session("BackURL") & "<BR>"
  response.write "LOGON User:" & Session("Logon_User") & "<BR>"
  response.write "User Name:" & Session ("Username") & "<BR>"
  response.write "Full Name:" & Session ("FullName") & "<BR>"
  response.write "Email:" &  Session ("EmailAddress") & "<BR>"
  response.write "Language:" & Session ("Language") & "<BR>"
  response.write "Account ID:" & Session ("User_id") & "<BR>"
  response.write "Is Administrator:" & Session ("IsAdministrator") & "<BR>"
  response.write "Forum Title:" & Session("Forum_Title") & "<BR>"
  response.write "Forum_ID:" & Session("Forum_ID") & "<BR>"
  response.write "Forum Moderated" & Session("Forum_Moderated") & "<BR>"
  response.write "Forum Moderator ID:" & Session("Forum_Moderator_ID") & "<BR>"
  response.write "Moderator Name:" & Session("Moderator_Name") & "<BR>"
  response.write "Moderator Email:" & Session("Moderator_Email") & "<BR>"
  response.write "<BR>"
end if

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

%>
