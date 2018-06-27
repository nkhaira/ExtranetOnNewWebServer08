<%@ LANGUAGE="vbScript"%>

<%
if request("Site_ID") <> "" and isnumeric(request("Site_ID")) then
  Site_ID            = request("Site_ID")
  Session("Site_ID") = request("Site_ID")  
elseif session("Site_ID") <> "" and isnumeric(session("Site_ID")) then
  Site_ID = session("Site_ID")
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

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code = rsSite("Site_Code")
Site_Description = rsSite("Site_Description")
Logo = rsSite("Logo")
Footer_Disabled = rsSite("Footer_Disabled")
  
rsSite.close
set rsSite=nothing

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Forum",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Forum",Login_Language,conn) & " / " & Translate("Discussion Group",Login_Language,conn) & " - " & Translate("Message Thread",Login_Language,conn) & "</FONT>"
 
Side_Navigation = False
Top_Navigation  = False
Content_Width   = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->

<!--#include file="SWEForums.asp" -->
<%

response.write "<H3>"
Call FORUM_TITLE_DISC
response.write "<BR><FONT CLASS=SmallBold>Message Thread</FONT>"
response.write "</H3>"
Call FORUM_LINK_DISC
Call STD_MESSAGE_DISC

response.write "<FONT CLASS=Small>" & Translate("Other Message Threads",Login_Language,conn) & "</FONT> <FONT CLASS=Small>(" & Translate("The current message that you are viewing is indicated with right arrow symbol",Login_Language,conn) & " <IMG SRC=""Threadmarker.gif"" HEIGHT=""7"" WIDTH=""7"" BORDER=""0"">" & " )</FONT>"
Response.write "<HR NOSHADE COLOR=Black>"
response.write "<BR>"
Call THREAD_DISC
response.write "<BR>"
Response.write "<HR NOSHADE COLOR=Black>"

Call CLEANUP_DISC

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

%>