<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/connections/connection_SiteWide.asp"-->

<% 

' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
'             Sandbox
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Navigation
Dim Content_Width
Dim Login_Language
Dim Site_ID

Site_ID = 0

if isblank(request("Language")) then
  Login_Language = "eng"
else
  Login_Language = request("Language")
end if
  
Call Connect_SiteWide

Screen_Title    = Translate(Request("Site_Description"),Login_Language,conn) & " - " & Translate("Extranet Support Site", Login_Language, conn)
Bar_Title       = Translate(Request("Site_Description"),Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Site Closed", Login_Language, conn) & "</FONT>"
 
Navigation = False
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

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->

