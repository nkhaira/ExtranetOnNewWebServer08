<%@ Language="VBScript" CODEPAGE="65001"%>
    
<%
Dim strBackURL

Dim Script_Debug
Script_Debug = false

  strBackURL = ""
  'Session("Site_ID")    = NULL
  'Session("LOGON_USER") = NULL
  'Session("Password")   = NULL 
  
  ' response.redirect "/register/login.asp"


%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
Call Connect_SiteWide

%>
<!--#include virtual="/include/functions_locator.asp"-->
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<% 
 

'  Login_Attempts = Login_Attempts + 1 
Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' Check if the user is blocking popups and advise them to enable popups for this site.
' --------------------------------------------------------------------------------------

if CInt(Session("PopUpCheck")) <> CInt(true) then
  %>
  <script type="text/JavaScript" language="JavaScript">
  var popUpsBlocked = false;
  var mine = window.open('','mine','width=1,height=1,left=0,top=0,scrollbars=no');
  if(mine) {
    //mine.blur();
    mine.close();
  }
  else {
    var popUpsBlocked = true;
  }

  if (popUpsBlocked == true || <%=CInt(Session("PopUpCheck"))%> == -1) {
    alert("Browser Compatiblity Notice:\r\n\nThis site requires that you allow popup windows to open from this site in order to use its advanced capabilities and delivery methods for the information that you have requested to view.  click on [OK] to continue, then allow popup windows for this site to open.");
  }
  </script>
  <%
end if

Session("PopUpCheck") = true

%>
	<form>
<TABLE id="Table1" border="1" cellSpacing="1" cellPadding="1" width="100%">
				
				<TR>
					<TD align="center"><FONT class="prodheadline" size="+0">You have tried to access a link 
							without authorization.
							<BR>
							<BR>
							<BR>
							Please contact <A href="mailto:support@fluke.com">support@fluke.com</A> for 
							assistance.</FONT>
						<BR>
					</TD>
				</TR>
</TABLE>
<BR><BR>
<DIV ALIGN="center"  >
 <!-- Begin Footer -->

  <TR>
           
    <% if Side_Navigation = True then %>
    <TD><IMG SRC="/Images/Spacer.gif" WIDTH=4 BGCOLOR=WHITE></TD>
    <TD BGCOLOR=WHITE>&nbsp;</TD>
    <% end if %>
    <TD ALIGN="CENTER" VALIGN="TOP" CLASS=Small BGCOLOR=WHITE>
      <%
      if not Footer_Disabled then
        response.write "&copy; 1995-" & DatePart("yyyy",Date) & " " & Translate("Fluke Corporation",Login_Language,conn) & " - " & Translate("All rights reserved",Login_Language,conn) & "."
      end if
      if Access_Level >= 8 then
        Page_Timer = Now() - Page_Timer_Begin
        response.write "<BR><SPAN CLASS=Small>Server Compilation Time: [" & FormatTime(Page_Timer) & "]</SPAN>" & vbCrLf & vbCrLf
      end if  
      %>
    </TD>
  </TR>

  <!-- End Footer -->
</DIV>
</form>
 	 
			  