<%@ Language="VBScript" CODEPAGE="65001" %>


<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->
<!--#include virtual="/SW-Common/SW-Security_Validate_Formdata_Sitewide.asp" -->

<%
' --------------------------------------------------------------------------------------
' Setup Connection
' --------------------------------------------------------------------------------------
response.buffer = true

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Build Nav and Header information
' --------------------------------------------------------------------------------------

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_Width
Dim BackURL
Dim Border_Toggle

Border_Toggle = 0

Top_Navigation  = False 
Side_Navigation = True
Screen_Title    = Site_Description & " - " & Translate("Calibration Procedure Details",Alt_Language,conn)
Bar_Title       = Site_Description & "<BR><SPAN CLASS=SmallBoldGold>" & Translate("Calibration Procedure Details",Login_Language,conn) & "</SPAN>"
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")
%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->

<%
' --------------------------------------------------------------------------------------
' Get Posted values, build various variables needed later
' --------------------------------------------------------------------------------------
Dim strNewRecord
Dim iID
Dim strSearchKeyword
Dim strSearchCalibrator

strSearchKeyword = request("keyword")
strSearchCalibrator = request("calibrator")

iID = Trim(Request("iProcedure_ID"))

' --------------------------------------------------------------------------------------
' Get the record if ID is a number and it's not a new record
' --------------------------------------------------------------------------------------
if IsNumeric(iID) then
  Dim rsProcedure
  
  set rsProcedure = GetProcedureDetails(iID)
end if

' --------------------------------------------------------------------------------------
' Start building form to display procedure details
' --------------------------------------------------------------------------------------

response.write "<DIV ALIGN=center>"
response.write "<SPAN CLASS=Heading3>" & Translate("Calibration Procedure Details",Login_Language,conn) & "</SPAN>"
response.write "</DIV>"
response.write "<BR><BR>"

%>

<DIV ALIGN=CENTER>
    <%Call Nav_Border_Begin%>
      <TABLE CELLPADDING=1 CELLSPACING=4 BORDER=0 BGCOLOR="#FFCC00" WIDTH="100%">
    		<TR>
    		  <TD CLASS=MediumBold><%=Translate("Instrument",Login_Language,conn)%></TD>
    		  <TD CLASS=MediumBold>
  		    <%=ProcValue(rsProcedure, "INSTRUMENT")%>
    		  </TD>
	      </TR>                
    		<TR>
    		  <TD CLASS=MediumBold><%=Translate("Adj Threshold",Login_Language,conn)%></TD>
    		  <TD CLASS=MediumBold><%=ProcValue(rsProcedure, "ADJTHRESHOLD")%></TD>
	      </TR>                
    		<TR>
    		  <TD CLASS=MediumBold><%=Translate("Author",Login_Language,conn)%></TD>
    		  <TD CLASS=MediumBold>
    			<%=ProcValue(rsProcedure, "AUTHOR")%>
    		  </TD>
  	    </TR>                
    		<TR>
    		  <TD CLASS=MediumBold><%=Translate("Company",Login_Language,conn)%></TD>
    		  <TD CLASS=MediumBold>
    			<%=ProcValue(rsProcedure, "COMPANY")%>
    		  </TD>
        </TR>                
    		<TR>
    		  <TD CLASS=MediumBold><%=Translate("Date",Login_Language,conn)%></TD>
    		  <TD CLASS=MediumBold><%=ProcValue(rsProcedure, "DATE")%></TD>
        </TR>                
    		<TR>
		      <TD CLASS=MediumBold><%=Translate("Primary Calibrator",Login_Language,conn)%></TD>
    		  <TD CLASS=MediumBold>
    			<%=ProcValue(rsProcedure, "PRIMCALIBRATOR")%>
    		  </TD>
	      </TR>                
    		<TR>
		      <TD CLASS=MediumBold><%=Translate("Revision",Login_Language,conn)%></TD>
    		  <TD CLASS=MediumBold><%=ProcValue(rsProcedure, "REVISION")%></TD>
    	    </TR>                
  	    <TR>
	    	  <TD CLASS=MediumBold><%=Translate("Procedure Type",Login_Language,conn)%></TD>
    		  <TD CLASS=MediumBold>
    			<%=ProcValue(rsProcedure, "TYPE")%>
    		  </TD>
	      </TR>                
    		<TR>
    		  <TD CLASS=MediumBold><%=Translate("5500 CAL_Ready",Login_Language,conn)%></TD>
    		  <TD CLASS=MediumBold>
    			<% if rsProcedure("CAL_READY") = 0 then
      	 			response.write("NO")
    	 	     else
      				response.write("YES")
			       end if
    			%>
    		  </TD>
  	    </TR>                
	      <TR>
    		  <TD CLASS=MediumBold><%=Translate("Source",Login_Language,conn)%></TD>
    		  <TD CLASS=MediumBold>
    			<%=rsProcedure("SOURCE")%>
		      </TD>
  	    </TR>                
  	    <TR>
      		<TD class="MediumBold" valign="top"><%=Translate("Description",Login_Language,conn)%></TD>
      		<TD class="MediumNormal"><FONT FACE=Courier>
    			<%=ProcValue(rsProcedure, "DESCRIPTION")%>
      		</FONT></TD>
  	    </TR>
	    </TABLE>
    <%Call Nav_Border_End%>
</DIV>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->


<%
Call Disconnect_SiteWide
response.flush
%>

<%
'--------------------------------------------------------------------------------------

Function GetProcedureDetails(iID)
  Dim cmd, prm, rsProcedure

  Set cmd = Server.CreateObject("ADODB.Command")
  Set cmd.ActiveConnection = conn
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = "MetCal_GetProcedure"

  Set prm = cmd.CreateParameter("@iID", adInteger, adParamInput, , cInt(iID))
  cmd.Parameters.Append prm

  Set rsProcedure = Server.CreateObject("ADODB.Recordset")
  rsProcedure.CursorLocation = adUseClient
  rsProcedure.CursorType = adOpenDynamic
  rsProcedure.open cmd

  set prm = nothing
  set cmd = nothing
  
  set GetProcedureDetails = rsProcedure
End Function

'--------------------------------------------------------------------------------------

Function ProcValue(rsProcedure, strField)
  if IsObject(rsProcedure) then
  	FieldData = rsProcedure(strField)
    if not isblank(FieldData) then
      FieldData = Replace(FieldData," ","&nbsp;")
    end if
    ProcValue = FieldData
  end if
End Function

'--------------------------------------------------------------------------------------

%>