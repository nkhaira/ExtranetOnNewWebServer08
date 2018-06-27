<%@ LANGUAGE="VBSCRIPT"%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Setup Connection
' --------------------------------------------------------------------------------------
response.buffer = true

Call Connect_SiteWide

%>
<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
<!-- #include virtual="/SW-Administrator/CK_Admin_Credentials.asp"-->
<!-- #include virtual="/met-support-gold/admin/CK_Credentials.asp"-->
<%

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

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
Screen_Title    = Site_Description & " - " & "Calibration Procedure Administration"
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & "Calibration Procedure Administration" & "</FONT>"
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

response.write "<FONT CLASS=Heading3>MetCal Procedure Administration</FONT>"
response.write "<BR><BR>"
response.write "<FONT CLASS=Small>"

%>


<% Call Nav_Border_Begin %>

<TABLE BORDER=0 WIDTH="100%" BGCOLOR="#666666" CELLPADDING=0 CELLSPACING=0>
  <TR>
    <TD>
      <TABLE BORDER=0 WIDTH="100%" BGCOLOR="#FFCC00" CELLPADDING=0 CELLSPACING=0>
        <TR>
          <TD>
            <FORM ACTION="Metcal_Categories.asp" METHOD="POST">
            <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="#FFCC00" WIDTH="100%">
            	<TR>
            	  <TD COLSPAN=2>
                  <INPUT TYPE="button" onclick="return document.location='metcal_procedure.asp?new=true'" VALUE=" Create New MetCal Procedure " CLASS=NAVLEFTHIGHLIGHT1>
                  <hr COLOR="#000000>"
                </TD>
            	</TR>
            	<TR>
            	  <TD COLSPAN=2><SPAN CLASS=SMALLBOLD>Manage Pre-Populated Values</SPAN></TD>
            	</TR>
            	<TR>
            	  <TD WIDTH="30%">
                  <SELECT NAME="SubCategory" CLASS=SMALL STYLE="width:200px;height:20px;">
                  <OPTION VALUE="">Select from this list</OPTION>
              		<%=GetSubProcedures%>
            	  </TD>
            	  <TD WIDTH="70%">
                  <INPUT TYPE="submit" VALUE=" Edit Selection " CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&nbsp;&nbsp;
                  <INPUT TYPE="Button" Value="Check Price Points" CLASS=NAVLEFTHIGHLIGHT1 ONCLICK="window.location.href='/met-support-gold/admin/metcal_pricepoints.asp';">   
                </TD>		
            	</TR>
            </TABLE>
            </FORM>
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
  
  <TR>
    <TD>
      <TABLE BORDER=0 WIDTH="100%" BGCOLOR="#FFCC00" CELLPADDING=0 CELLSPACING=0>
        <TR>
          <TD>
            <FORM ACTION="MetCal_Procedures.asp" METHOD="POST">
            <INPUT TYPE="Hidden" NAME="lv" VALUE="<%=request("LV")%>">
            <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="#FFCC00" WIDTH="100%">
            	<TR>
            	  <TD COLSPAN=2>
                  <SPAN CLASS=SMALLBOLD>Search for an existing procedure</SPAN>
                  <SPAN CLASS=SMALL><I>(Leave blank list all procedures)</I><HR COLOR="#000000"></SPAN>
                </TD>
              </TR>
              <TR>
                <TD WIDTH="30%" CLASS=SMALL>Procedure Name Keyword:</TD>
                <TD WIDTH="70%" CLASS=SMALL>
                	<INPUT TYPE="Text" NAME="KeyWord" SIZE="25" CLASS=SMALL>
                </TD>
              </TR>
        
              <TR>
                <TD CLASS=SMALL>Primary Calibrator:</TD>
                <TD CLASS=SMALL>
                  <SELECT NAME="Calibrator" CLASS=SMALL>
                    <OPTION VALUE="">Select from this list</OPTION>
                    <OPTION VALUE="5011">5011</OPTION>
                    <OPTION VALUE="5100">5100</OPTION>
                    <OPTION VALUE="5500">5500</OPTION>
                    <OPTION VALUE="5520">5520</OPTION>
                    <OPTION VALUE="5700">5700</OPTION>
                    <OPTION VALUE="5720">5720</OPTION>
                    <OPTION VALUE="5790">5790</OPTION>
                    <OPTION VALUE="5800">5800</OPTION>
                    <OPTION VALUE="5800">5820</OPTION>
                    <OPTION VALUE="9100">9100</OPTION>
                    <OPTION VALUE="9500">9500</OPTION>
                    <OPTION VALUE="OTHER">Other</OPTION>
                  </SELECT>
                </TD>
              </TR>                
    
              <TR>
                <TD CLASS=SMALL>Procedure Filename Search:</TD>
                <TD CLASS=SMALL>
                	<INPUT TYPE="Text" NAME="FileName" SIZE="25" CLASS=SMALL>
                </TD>
              </TR>
    
              <TR>
                <TD CLASS=SMALL>Restricted:</TD>
                <TD CLASS=SMALL>
                	<INPUT TYPE="Checkbox" NAME="Restricted" Value="-1" CLASS=SMALL>
                </TD>
              </TR>

              <TR>
                <TD COLSPAN=2 BGCOLOR="BLACK">
                  <TABLE WIDTH="100%">
                    <TR>
                      <TD WIDTH="30%"><INPUT TYPE="reset" VALUE=" Clear " CLASS=NAVLEFTHIGHLIGHT1></TD>
                      <TD WIDTH="30%"><INPUT TYPE="submit" VALUE=" Begin Search " CLASS=NAVLEFTHIGHLIGHT1 NAME="submit" ID="submit"></TD>
                      <TD WIDTH="40%" ALIGN="right"><INPUT onclick="document.location='output_excel.asp';return false;" TYPE="submit" VALUE=" Dump Database " CLASS=NAVLEFTHIGHLIGHT1 ID="DumpDatabase" NAME="DumpDatabase"></TD>
                    </TR>
                  </TABLE>
                </TD>   
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

' --------------------------------------------------------------------------------------

function GetSubProcedures()
  Dim cmd, rsSubCategories
  Dim strOptionList
  
  Set cmd = Server.CreateObject("ADODB.Command")
  Set cmd.ActiveConnection = conn
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = "Admin_Metcal_SubCategories_GetList"

  Set rsSubCategories = Server.CreateObject("ADODB.Recordset")
  rsSubCategories.CursorLocation = adUseClient
  rsSubCategories.CursorType = adOpenDynamic
  rsSubCategories.open cmd

  set cmd = nothing

  do while not rsSubCategories.EOF
	strOptionList = strOptionList & "<option value=""" & rsSubCategories("Description")
	strOptionList = strOptionList & """>" & rsSubCategories("Description")
	rsSubCategories.movenext
  loop

  GetSubProcedures = strOptionList

end function
%>