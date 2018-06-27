<%@ LANGUAGE="VBSCRIPT"%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%
' --------------------------------------------------------------------------------------
' Setup Connection
' --------------------------------------------------------------------------------------
response.buffer = true

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------
%>

<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
<!-- #include virtual="/SW-Administrator/CK_Admin_Credentials.asp"-->
<!-- #include virtual="/met-support-gold/admin/CK_Credentials.asp"-->

<%
' --------------------------------------------------------------------------------------
' Check to see whether we should be redirecting
' --------------------------------------------------------------------------------------
Dim strProcType

strProcType = Trim(uCase(Request("ProcType")))
if strProcType = "METCAL" then
	response.redirect "MetCal_Admin.asp"
elseif strProcType = "PORTOCAL" then
	response.redirect "PortoCal_admin.asp"
end if

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

' --------------------------------------------------------------------------------------
' Build Nav and Header information
' --------------------------------------------------------------------------------------
Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_Width
Dim BackURL

Top_Navigation  = False 
Side_Navigation = True
Screen_Title    = Site_Description & " - " & "Calibration/PortoCal Procedure Admin"
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & "Calibration/PortoCal Procedure Admin" & "</FONT>"
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")
%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->

<%
response.write "<FONT CLASS=Heading3>Calibration/PortoCal Procedure Admin</FONT>"
response.write "<BR><BR>"
response.write "<FONT CLASS=Medium>"

' --------------------------------------------------------------------------------------
' Build Form for user to make the choice between administrating metcal or portocal procedures
' --------------------------------------------------------------------------------------
%>
<FORM ACTION="default.asp" METHOD="POST">
<INPUT TYPE="Hidden" NAME="lv" VALUE="<%=request("lv")%>">

<% Call Nav_Border_Begin %>
<TABLE BORDER=0 WIDTH="100%" BGCOLOR="#FFCC00" CELLPADDING=0 CELLSPACING=0>
  <TR>
    <TD>
      <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="#FFCC00" WIDTH="100%">
        
        <TR>
          <TD ROWSPAN=2 WIDTH="50%" CLASS=MediumBold valign=top>Select what procedures you'd like to administer:</TD>
          <TD width="50%" CLASS=MediumBold valign=top>
	    <INPUT type="radio" name="proctype" value="METCAL" checked=true>&nbsp;Met/Cal Procedures<br>
	  </TD>
	</TR>
	<TR>
	  <TD VALIGN=TOP CLASS=MediumBold>
	    <INPUT type="radio" name="proctype" value="PORTOCAL">&nbsp;PortoCal Procedures<br>
          </TD>
        </TR>                
        <TR>
          <TD COLSPAN=2 BGCOLOR="BLACK">
            <TABLE WIDTH="100%">
              <TR>
                <TD WIDTH="100%"><INPUT TYPE="submit" VALUE=" Edit Procedures " CLASS=NavLeftHighlight1></TD>
              </TR>
            </TABLE>
          </TD>   
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>
<% Call Nav_Border_End %>
</FORM>


<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
response.flush
%>