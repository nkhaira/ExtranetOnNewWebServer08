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

Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Side_Navigation
Dim Content_Width


Screen_Title    = Request("Site_Description") & " - " & Translate("Site Closed", Login_Language, conn)
Bar_Title       = Request("Site_Description") & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Site Closed", Login_Language, conn) & "</FONT>"
 
Side_Navigation = False
Content_Width   = 95  ' Percent

%>  
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/SW-Common/SW-Header.asp"-->  

<!-- BEGIN CONTENT -->
<%

response.write "<DIV ALIGN=CENTER>"
response.write Translate("We are sorry, but this site is temporarialy closed while we are performing site software updates and server maintenance.<BR><BR>Please check this site again in an hour.", Login_Language, conn)
response.write "</DIV>"

response.write "</TD></TR></TABLE>"

%>
      
<!-- END CONTENT -->

<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->

<%
