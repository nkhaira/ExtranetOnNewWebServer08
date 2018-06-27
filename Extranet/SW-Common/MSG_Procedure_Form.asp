<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_estore.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_FormData.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->
<!--#include virtual="/SW-Common/SW-Security_Validate_Formdata_Sitewide.asp" -->

<%
'redirection as per FWS-583 dt. 08mar2018
Response.Redirect "https://us.flukecal.com/support/my-met-support/procedures/"

' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       2/1/2000
'             Sandbox
' --------------------------------------------------------------------------------------

Call Connect_SiteWide
Call Connect_eStoreDatabase
Call Validate_Security(g_strSitewideUser, g_strCore_ID, Site_ID, g_iSitewide_ID, g_bCoreExists, g_strEmail, g_strPassword) 

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------
%>

<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->

<%

site_description = "MET/CAL Plus"

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Site_Description & " - " & Translate("Calibration Procedure Download",Alt_Language,conn)
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Calibration Procedure Download",Login_Language,conn) & "</FONT>"
Top_Navigation  = False 
Side_Navigation = True
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

response.write "<FONT CLASS=Heading3>Calibration Procedure Download</FONT>"
response.write "<BR><BR>"

response.write "<SPAN CLASS=Medium>"

response.write Translate("<B>MET/CAL Plus </B>is provided with an assortment of procedures for many different makes and models of test and measurement instruments.",Login_Language,conn) & " "
response.write Translate(" The procedures that you can download from this site which are provided by Fluke are ready to put to work as is,or you can use them as the basis for developing your own custom procedures to fit your operations' specific requirements and calibration workload.",Login_Language,conn)
response.write "<P>"
response.write Translate("You may choose to download from a wide selection of MET/CAL calibration procedures.",Login_Language,conn) & " "
response.write Translate("Most were created by Fluke and are fully supported by Fluke engineering.",Login_Language,conn) & " "
response.write Translate("Some have been supplied by Fluke MET/CAL customers for the benefit of the MET/CAL user community.",Login_Language,conn) & " "
response.write Translate("These ""User Contributed"" procedures are made available ""As-Is"" and have not been verified by Fluke in any way.",Login_Language,conn)
response.write "<P>"
response.write "<FONT COLOR=""red"">"
response.write Translate("Please Note",Login_Language,conn) & ": " & Translate("4805 based procedures will also execute when any 47xx or 48xx calibrator is used.",Login_Language,conn) & " "
response.write Translate("Also 5700A based procedures will execute when a 5720A or 5700A/EP is used.",Login_Language,conn)
response.write "</FONT>"
response.write "<P>"

%>
<FORM ACTION="MSG_Procedure_Results.asp" METHOD="POST" NAME="Procedure_Results">
<INPUT TYPE="Hidden" NAME="lv" VALUE="<%=request("LV")%>">
<INPUT TYPE="Hidden" NAME="strCore_ID" VALUE="<%=g_strCore_ID%>">
<INPUT TYPE="Hidden" NAME="iSitewide_ID" VALUE="<%=g_iSitewide_ID%>">

<DIV ALIGN=CENTER>
      <%Call Nav_Border_Begin%>
      <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="#FFCC00" WIDTH="100%">
        <TR>
          <TD WIDTH="40%" CLASS=SMALLBOLD><%=Translate("Procedure Name Keyword",Login_Language,conn)%>:</TD>
          <TD WIDTH="60%" CLASS=SMALLBOLD>
          	<INPUT CLASS=SMALL TYPE="Text" NAME="KeyWord" SIZE="25">
          </TD>
        </TR>
        
        <TR>
          <TD CLASS=SMALLBOLD><%=Translate("Primary Calibrator",Login_Language,conn)%>:</TD>
          <TD CLASS=SMALLBOLD>
            <SELECT NAME="Calibrator" CLASS=SMALL>
              <OPTION VALUE=""><%=Translate("Select from this list",Login_Language,conn)%></OPTION>
              <OPTION VALUE="5011">5011</OPTION>
              <OPTION VALUE="5100">5100</OPTION>                                                
              <OPTION VALUE="5500">5320</OPTION>              
              <OPTION VALUE="5500">5500</OPTION>
              <OPTION VALUE="5520">5520</OPTION>
              <OPTION VALUE="5700">5700</OPTION>
              <OPTION VALUE="5720">5720</OPTION>
              <OPTION VALUE="5790">5790</OPTION>
              <OPTION VALUE="5800">5800</OPTION>
              <OPTION VALUE="5800">5820</OPTION>
              <OPTION VALUE="9100">9100</OPTION>
              <OPTION VALUE="9500">9500</OPTION>
              <OPTION VALUE="OTHER"><%=Translate("Other",Login_Language,conn)%></OPTION>
            </SELECT>
          </TD>
        </TR>                
        <TR>
          <TD COLSPAN=2 BGCOLOR="#666666">
            <TABLE WIDTH="100%">
              <TR>
                <TD WIDTH="50%" ALIGN=CENTER><INPUT TYPE="reset" VALUE=" <%=Translate("Clear Search Criteria",Login_Language,conn)%> " CLASS=NAVLEFTHIGHLIGHT1></TD>
                <TD WIDTH="50%" ALIGN=CENTER><INPUT TYPE="submit" VALUE=" <%=Translate("Begin Search",Login_Language,conn)%> " CLASS=NAVLEFTHIGHLIGHT1></TD>
<%
strQueryString = "lv=" & request("lv") & "&strCore_ID=" & g_strCore_ID & "&iSitewide_ID=" & g_iSitewide_ID & "&strAction=SHOWPURCHASED"
%>
<TD WIDTH="40%">
<%
  Dim aBoughtProcedures

'response.write("g_iSitewide_ID: " & g_iSitewide_ID & "<BR>")
'response.write("g_bCoreExists: " & g_bCoreExists & "<BR>")
'response.write("g_strEmail: " & g_strEmail & "<BR>")
'response.write("g_strPassword: " & g_strPassword & "<BR>")
'response.write("g_strCore_ID: " & g_strCore_ID & "<BR>")

  aBoughtProcedures = GetPurchasedProcedures(g_iSitewide_ID, g_bCoreExists, g_strEmail, g_strPassword, g_strCore_ID)
  
  if IsArray(aBoughtProcedures) then
%>
    <INPUT TYPE="button" NAME="purchased" onclick="return document.location='msg_procedure_results.asp?<%=strQueryString%>'" VALUE="<%=Translate("Show Purchased Procedures",Login_Language,conn)%>" CLASS=NavLeftHighlight1>
<% 	
  end if 
%>
		</TD>
              </TR>
            </TABLE>
          </TD>   
        </TR>
      </TABLE>
      <%Call Nav_Border_End%>
  </FORM>

</DIV>

<BR>
<B><%=Translate("How to search",Login_Language,conn)%>:</B><BR><BR>
<OL><LI><%=Translate("Enter a keyword to search for in the procedure names. By convention, Fluke written procedures denote the UUT model and type calibration performed in the procedure name.",Login_Language,conn)%> </LI>
<LI><%=Translate("If you want to limit your search to a particular primary calibrator, select a primary calibrator model number from the drop down list.",Login_Language,conn)%><BR>
	<B><%=Translate("or",Login_Language,conn)%></B>... <BR><%=Translate("Leave both search criteria blank to list all available procedures.",Login_Language,conn)%></LI>
<LI><%=Translate("Click on Begin Search",Login_Language,conn)%></FONT></LI>
</OL>
<%=Translate("All procedures can be executed with MET/CAL if your system has the necessary configuration of instruments.  The 5500/CAL column indicates all procedures designed to run with 5500/CAL.",Login_Language,conn)%><BR>
<BR>
</SPAN>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
response.flush
%>

<%
' --------------------------------------------------------------------------------------
' Functions
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Get a list of procedures this user has bought
' --------------------------------------------------------------------------------------
Function GetPurchasedProcedures(iSitewide_ID, bCoreExists, strEmail, strPassword, strCore_ID)
	Dim cmd, prm
	Dim rsShopper, rsBoughtProcedures
	Dim iShopperID
	Dim iCounter
	Dim aBoughtProcedures

	if iSitewide_ID > 0 or bCoreExists = true then

	  Set cmd = Server.CreateObject("ADODB.Command")
	  Set cmd.ActiveConnection = eConn
	  cmd.CommandType = adCmdStoredProc

	  ' First, get the eStore shopper id, if it exists
	  if iSitewide_ID > 0 then
		cmd.CommandText = "sp_GetShopper_By_Email_Password"
		Set prm = cmd.CreateParameter("@strEmail", adVarchar, adParamInput, 50, strEmail & "")
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("@strPassword", adVarchar, adParamInput, 50, strPassword & "")
		cmd.Parameters.Append prm
	  else
		cmd.CommandText = "sp_GetShopper_By_CoreID"
		Set prm = cmd.CreateParameter("@strCoreID", adVarchar, adParamInput, 25, strCore_ID & "")
		cmd.Parameters.Append prm
	  end if

	  Set rsShopper = Server.CreateObject("ADODB.Recordset")
	  rsShopper.CursorLocation = adUseClient
	  rsShopper.CursorType = adOpenStatic
	  rsShopper.open cmd
	  set prm = nothing
	  set cmd = nothing

	  if not rsShopper.EOF then
		iShopperID = rsShopper("shopper_ID")

		Set cmd = Server.CreateObject("ADODB.Command")
		Set cmd.ActiveConnection = eConn
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "sp_GetMetcalReceiptItems"
		set prm = cmd.CreateParameter("@strShopperID", adVarChar, adParamInput, 32, iShopperID)
		cmd.Parameters.Append prm

		set rsBoughtProcedures = Server.CreateObject("ADODB.Recordset")
		rsBoughtProcedures.CursorLocation = adUseClient
		rsBoughtProcedures.CursorType = adOpenStatic
		rsBoughtProcedures.open cmd
		set prm = nothing
		set cmd = nothing
	
		iCounter = 0
		ReDim aBoughtProcedures(iCounter)

		do while not rsBoughtProcedures.eof
			ReDim Preserve aBoughtProcedures(iCounter)
			aBoughtProcedures(iCounter) = rsBoughtProcedures("procedure_id")
			rsBoughtProcedures.movenext
			iCounter = iCounter + 1
		loop

		rsBoughtProcedures.close
		set rsBoughtProcedures = nothing
	  end if
	end if
	
	rsShopper.close
	set rsShopper = nothing

	if iCounter > 0 then
		GetPurchasedProcedures = aBoughtProcedures
	else
		GetPurchasedProcedures = ""
	end if
End Function

'--------------------------------------------------------------------------------------
  
%>