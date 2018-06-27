<%@ Language="VBScript" CODEPAGE="65001" %>
<% 

' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
'             Sandbox
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_SiteWide

Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Navigation
Dim Content_Width
Dim Site_ID

if isblank(request("Site_ID")) then
  Site_ID = 0
  Site_Description = "Support.Fluke.com"
else  
  Site_ID = CInt(request("Site_ID"))
  SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3
  
  Site_Code = rsSite("Site_Code")      
  Site_Description = rsSite("Site_Description")
  Logo = rsSite("Logo")
  Footer_Disabled = rsSite("Footer_Disabled")
end if  

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Site Closed",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Site Closed",Login_Language,conn) & "</FONT>"
Top_Navigation  = False 
Navigation      = False
Content_Width   = 95  ' Percent

%>  
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->  

<!-- BEGIN CONTENT -->
<%

response.write "<BR><BR>"
response.write "<DIV ALIGN=CENTER>"
response.write "<FONT CLASS=Heading3Black>" & Translate("Site Temporarily Closed",Login_Languge,conn) &  "</FONT>"
response.write "<BR><BR>"
response.write "<FONT CLASS=MediumRed>"
response.write Translate("We are sorry, but this site is temporarily closed while we are performing site software updates and server maintenance.",Login_Language,conn)
response.write "<BR>"
response.write Translate("We hope to have this site open again within the hour and apologize for any inconvenience this may have caused you.", Login_Language, conn)
response.write "</FONT>"
response.write "<FORM Action=""/register/default.asp"" METHOD=""Post"">"
response.write "<INPUT TYPE=""Submit"" Value=""" & Translate("Continue",Login_Language,conn) & """ CLASS=NavLeftHighlight1>"
response.write "</FORM>"
response.write "</DIV>"
%>
      
<!-- END CONTENT -->

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
%>

