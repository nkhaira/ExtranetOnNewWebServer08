<%@ LANGUAGE="VBSCRIPT"%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
'             Sandbox
' --------------------------------------------------------------------------------------

response.buffer = true

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------
' Get Recordset to populate the form
Dim cmd, rsValues, rsMake, rsModel, rsCalType, rsSoftVer, rsOptions

Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "PortoCal_GetSearchValues"

Set rsValues = Server.CreateObject("ADODB.Recordset")
rsValues.CursorLocation = adUseClient
rsValues.CursorType = adOpenDynamic
rsValues.open cmd


%>

<%
'#include virtual="/SW-Common/SW-Security_Module.asp" -->

site_id = 11

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

site_description = "Portocal"

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Site_Description & " - " & Translate("Portocal II Procedure Download",Login_Language,conn)
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Portocal II Procedure Download",Login_Language,conn) & "</FONT>"
Top_Navigation  = False 
Side_Navigation = True
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%
response.write "<FONT CLASS=Heading3>" & Translate("Portocal II Procedure Download",Login_Language,conn) & "</FONT>"
response.write "<BR><BR>"

response.write "<FONT CLASS=Medium>"

response.write Translate("Portocal II is available with thousands of procedures for hundreds of different makes and models of test and measurement instruments. The procedures that you can download from this site are ready to put to work as is.  Or you can use them as the basis for developing your own custom procedures to fit your operations specific requirements and calibration workload.  Many are even certified to meet the manufacturer's requirements.",Login_Language,conn)
response.write "<P>"
response.write Translate("You may choose to download from a wide selection of Portocal II calibration procedures. These procedures have been gathered from several sources. Most were created by Fluke and are fully supported by Fluke engineering. Some have been supplied by Fluke MET/CAL customers for the benefit of the MET/CAL user community. These &quot;User Contributed&quot; procedures are made available &quot;As-Is&quot; and have not been verified by Fluke in any way.",Login_Language,conn)

%>

<FORM ACTION="Portocal_Procedure_Results.asp" METHOD="POST">
<INPUT TYPE="Hidden" NAME="lv" VALUE="<%=request("lv")%>">
<CENTER>

<TABLE BORDER=1 WIDTH="90%" BORDERCOLOR="#666666" BGCOLOR="#FFCC00" CELLPADDING=0 CELLSPACING=0>
            <TR>
              <TD WIDTH="60%">
                <TABLE>
                  <TR>
                    <TD WIDTH="100%" CLASS=Medium><BR><%=Translate("To locate the procedure you require please make your choice from the drop down lists below:",Login_Language,conn)%>
                    <P>
                    </TD>
                  </TR>
                  <TR>
                    <TD WIDTH="100%" CLASS=Medium>
                    <%
                    response.write "<LI><B>" & Translate("Manufacturer",Login_Language,conn) & "</B> - "
                    response.write Translate("Select the manufacturer of the instrument to be calibrated.",Login_Language,conn)
                    %>
                    </LI><P>
                    </TD>
                  </TR>
                  <TR>
                    <TD WIDTH="100%" CLASS=Medium>
                    <%
                    response.write "<LI><B>" & Translate("Model",Login_Language,conn) & "</B> - "
                    response.write Translate("Select the model number of the instrument you wish to calibrate.",Login_Language,conn)
                    %>
                    </LI><P>
                    </TD>
                  </TR>
                  <TR>
                    <TD WIDTH="100%" CLASS=Medium>
                    <%
                    response.write "<LI><B>" & Translate("Calibrator Type",Login_Language,conn) & "</B> - "
                    response.write Translate("Select the calibrator to be used to be used to calibrate the instrument.",Login_Language,conn)
                    %>
                    </LI><P>
                    </TD>
                  </TR>
                  <TR>
                    <TD WIDTH="100%" CLASS=Medium>
                    <%
                    response.write "<LI><B>" & Translate("Portocal/9010 Software version",Login_Languge,conn) & "</B> - "
                    response.write Translate("Select the version of software you are using. Some calibration procedures will require a specific software version or later in order to operate correctly.  Please do not choose a version later than the one installed on your unit.",Login_Language,conn)
                    %>
                    </LI><P>
                    </TD>
                  </TR>
                  <TR>
                    <TD WIDTH="100%" CLASS=Medium>
                    <%
                    response.write "<LI><B>" & Translate("Option",Login_Language,conn) & "</B> - "
                    response.write Translate("Procedures are available that support both GPIB-IEE488 (option 20) interface and PCMCIA cards option 40). All Calibration procedures developed for option 40 will operate with option 20.",Login_Language,conn)
                    %>
                    </LI><P>
                    </TD>
                  </TR>
                </TABLE>
              </TD>
              <TD VALIGN="TOP">
	<TABLE BORDER="0" ALIGN="CENTER"width="40%">
<TR>
              <TD CLASS=Medium><P>&nbsp;<P><B><%=Translate("Manufacturer",Login_Language,conn)%></B></TD>
              <TD CLASS=Medium><P>&nbsp;<P><SELECT NAME="make">
			  		<OPTION VALUE=""><%=Translate("All",Login_Language,conn)%></OPTION>
<%
	set rsMake = rsValues
	do while not rsMake.eof
		response.write("<OPTION VALUE=""" & rsMake("Make") & """ CLASS=MEDIUM>" & rsMake("Make") & "</OPTION>")
		rsMake.movenext
	loop
	set rsMake = nothing
%>
					</SELECT>
				</TD>
            </TR>
<TR>
              <TD CLASS=Medium><B><%=Translate("Model",Login_Language,conn)%></B></TD>
              <TD CLASS=Medium>
			  	<SELECT NAME="model">
			  		<OPTION VALUE="" CLASS=Medium><%=Translate("All",Login_Language,conn)%></OPTION>
<%
	set rsModel = rsValues.NextRecordset
	do while not rsModel.eof
		response.write("<OPTION VALUE=""" & rsModel("Model") & """ CLASS=MEDIUM>" & rsModel("Model") & "</OPTION>")
		rsModel.movenext
	loop
	set rsModel = nothing
%>
				</SELECT>
			  </TD>
            </TR>
<TR>
              <TD NOWRAP CLASS=Medium><B><%=Translate("Calibrator type",Login_Language,conn)%></B></TD>
              <TD>
			  	<SELECT NAME="caltype"><OPTION VALUE="" CLASS=Medium><%=Translate("All",Login_Language,conn)%></OPTION>
<%
	set rsCalType = rsValues.NextRecordset
	do while not rsCalType.eof
		response.write("<OPTION VALUE=""" & rsCalType("caltype") & """ CLASS=MEDIUM>" & rsCalType("caltype") & "</OPTION>")
		rsCalType.movenext
	loop
	set rsCalType = nothing
%>
				</SELECT>
			  </TD>
            </TR>
<TR>
              <TD NOWRAP CLASS=Medium><B><%=Translate("Software version",Login_Language,conn)%></B></TD>
              <TD CLASS=Medium>
			  	<SELECT NAME="minsoftver"><OPTION VALUE="" CLASS=Medium><%=Translate("All",Login_Language,conn)%></OPTION>
<%
	set rsMinVer = rsValues.NextRecordset
	do while not rsMinVer.eof
		response.write("<OPTION VALUE=""" & rsMinVer("minsoftver") & """ CLASS=MEDIUM>" & rsMinVer("minsoftver") & "</OPTION>")
		rsMinVer.movenext
	loop
	set rsMinVer = nothing
%>
				</SELECT>
			  </TD>
            </TR>
<TR>
              <TD CLASS=Medium><B><%=Translate("Option",Login_Language,conn)%></B></TD>
              <TD CLASS=Medium CLASS=Medium>
			  	<SELECT NAME="option"><OPTION VALUE="" CLASS=Medium><%=Translate("All",Login_Language,conn)%></OPTION>
<%
	set rsOption = rsValues.NextRecordset
	do while not rsOption.eof
		response.write("<OPTION VALUE=""" & rsOption("options") & """ CLASS=MEDIUM>" & rsOption("options") & "</OPTION>")
		rsOption.movenext
	loop
	set rsOption = nothing
	set rsValues = nothing
%>
				</SELECT>
			  </TD>
            </TR>
<TR><TD COLSPAN=2>&nbsp;</TD></TR>
<TR>
              <TD COLSPAN="2" ALIGN="CENTER" CLASS=Medium><INPUT CLASS=Medium TYPE="SUBMIT" VALUE="<%=Translate("Show Procedures",Login_Language,conn)%>"</TD>
            </TR>
</TABLE>
</td>
</tr>
</table>
</FORM>

</CENTER>

<BR>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
response.flush
%>