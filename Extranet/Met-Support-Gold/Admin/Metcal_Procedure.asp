<%@Language="VBScript" Codepage=65001%>
<%
' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock  (Rework of script originally developed by Jeff Patrick)
' Date:       2/7/2006
' Name:       Met/Cal Procedure Editor
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->
<%

' --------------------------------------------------------------------------------------
' Setup DB Connection
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

%>
<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
<!-- #include virtual="/SW-Administrator/CK_Admin_Credentials.asp"-->
<!-- #include virtual="/met-support-gold/admin/CK_Credentials.asp"-->
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

' --------------------------------------------------------------------------------------
' Get QueryString Values before FileUpEE
' --------------------------------------------------------------------------------------

Dim iID, strNewRecord, strSearchKeyword, strSearchCalibrator, strAction, CurrPage
Dim UpdateBy

iID                 = Trim(request.querystring("ID"))
strSearchKeyword    = request.querystring("Keyword")
strSearchCalibrator = request.querystring("Calibrator")
strSearchFileName   = request.querystring("FileName")
CurrPage            = request.querystring("CurrPage")
strNewRecord        = UCase(Trim(request.querystring("New")))
strAction           = UCase(Trim(request.querystring("Action")))
UpdateBy            = Admin_ID

'for each item in request.querystring
'  response.write item & ": " & request.querystring(item) & "<BR>"
'next
'response.flush
'response.end

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
Screen_Title    = Site_Description & " - " & "Calibration Procedure Download Admin"
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & "Calibration Procedure Download Admin" & "</FONT>"
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

' --------------------------------------------------------------------------------------
' Get the record if ID is a number and it's not a new record
' --------------------------------------------------------------------------------------
if IsNumeric(iID) and (strNewRecord = "FALSE" or strAction = "CLONE") then
  Dim rsProcedure
  set rsProcedure = GetProcedureDetails(iID)
end if

' --------------------------------------------------------------------------------------
' Start building form to display procedure details
' --------------------------------------------------------------------------------------

response.write "<FONT CLASS=Heading3>Calibration Procedure Download Administration</FONT>"
response.write "<BR><BR>"

response.write "<FONT CLASS=Small>"

Dim FormName
FormName = "Procedure"

%>
<FORM NAME="<%=FormName%>" ACTION="Metcal_Procedure_Admin.asp?PostFlag=-1" METHOD="POST" ENCTYPE="MULTIPART/FORM-DATA" ONSUBMIT="return CheckRequiredFields(this.form);">
<INPUT TYPE="HIDDEN" NAME="ID" VALUE="<%=ProcValue(rsProcedure, "PROCEDURE_ID")%>">
<INPUT TYPE="HIDDEN" NAME="Keyword" VALUE="<%=strSearchKeyWord%>">
<INPUT TYPE="HIDDEN" NAME="Calibrator" VALUE="<%=strSearchCalibrator%>">
<INPUT TYPE="HIDDEN" NAME="Action" VALUE="<%=strAction%>">
<INPUT TYPE="HIDDEN" NAME="UpdateBy" VALUE="<%=UpdateBy%>">
<INPUT TYPE="HIDDEN" NAME="UpdateDate" VALUE="<%=Date()%>">
<INPUT TYPE="HIDDEN" NAME="CurrPage" VALUE="<%=CurrPage%>">
<%

Call Nav_Border_Begin
response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" ONCLICK=""return document.location='Metcal_Admin.asp?KeyWord=" & strSearchkeyword & "&Calibrator=" & strSearchCalibrator & "'"" value=""MetCal Administration Menu"">"
response.write "&nbsp;&nbsp;&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" ONCLICK=""return document.location='metcal_procedures.asp?KeyWord=" & strSearchkeyword & "&Calibrator=" & strSearchCalibrator & "&FileName=" & strSearchFileName & "&CurrPage=" & CurrPage & "&ID=" & iID & "#PID" & iID & "'"" value=""Last Search Results"">"

if strAction <> "CLONE" and UCase(strNewRecord) <> "TRUE" then
  response.write "&nbsp;&nbsp;&nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" ONCLICK=""return document.location='metcal_procedure.asp?action=clone&id=" & iID & "'"" value=""Clone New Procedure"">"
end if

Call Nav_Border_End

response.write "<P>"
%>

<CENTER>
<% Call Table_Begin %>
<TABLE BORDER=0 WIDTH="100%" BORDERCOLOR="#666666" BGCOLOR="#FFCC00" CELLPADDING=0 CELLSPACING=0>
  <TR>
    <TD>
      <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="#FFCC00" WIDTH="100%">
      	<TR>
      	  <TD CLASS=SMALLBOLD WIDTH="20%">Instrument:</TD>
      	  <TD CLASS=SMALLBOLD WIDTH="80%">
      	    <INPUT CLASS=SMALL TYPE=TEXT NAME="Instrument" SIZE=100 MAXLENGTH=100 VALUE="<%=ProcValue(rsProcedure, "INSTRUMENT")%>" LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);">
      	  </TD>
        </TR>                
      	<TR>
      	  <TD CLASS=SMALLBOLD>Adj Threshold:</TD>
      	  <TD CLASS=SMALLBOLD><INPUT CLASS=SMALL TYPE=TEXT NAME="AdjThreshold" SIZE=10 MAXLENGTH=4 VALUE="<%=ProcValue(rsProcedure, "ADJTHRESHOLD")%>" LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);"></TD>
        </TR>                
      	<TR>
      	  <TD CLASS=SMALLBOLD>Author:</TD>
      	  <TD CLASS=SMALLBOLD>
      	    <SELECT NAME="Author" CLASS=SMALL LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);">
      		  <%=GetTypeInfo("AUTHORS", rsProcedure, "AUTHOR_ID")%>
      	    </SELECT>
      	  </TD>
        </TR>                
      	<TR>
      	  <TD CLASS=SMALLBOLD>Company:</TD>
      	  <TD CLASS=SMALLBOLD>
      	    <SELECT NAME="Company" CLASS=SMALL LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);">
      		  <%=GetTypeInfo("COMPANIES", rsProcedure, "COMPANY_ID")%>
      	    </SELECT>
      	  </TD>
        </TR>                
      	<TR>
      	  <TD CLASS=SMALLBOLD>Creation Date:</TD>
      	  <TD CLASS=SMALLBOLD>
          <%
          if LCase(strNewRecord) = "true" then
            response.write "<INPUT CLASS=SMALL TYPE=TEXT NAME=""Date"" SIZE=10 MAXLENGTH=10 VALUE=""" & FormatDateTime(Date(),"0") & """ LANGUAGE=""JavaScript"" ONCHANGE=""UnHighlight(this);"">"
          else
            response.write "<INPUT CLASS=SMALL TYPE=TEXT NAME=""Date"" SIZE=10 MAXLENGTH=10 VALUE=""" & FormatDateTime(ProcValue(rsProcedure, "Date"),"0") & """  LANGUAGE=""JavaScript"" ONCHANGE=""UnHighlight(this);"">"
            response.write "&nbsp;&nbsp;&nbsp;Add Date&nbsp;&nbsp;" & FormatDateTime(ProcValue(rsProcedure, "CreateDate"),"0")
          end if
          %>
          </TD>
        </TR>                
      	<TR>
      	  <TD CLASS=SMALLBOLD>Primary Calibrator:</TD>
      	  <TD CLASS=SMALLBOLD>
      	    <SELECT NAME="PrimCalibrators" CLASS=SMALL LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);">
        		<%=GetTypeInfo("PRIMARY CALIBRATORS", rsProcedure, "PRIMCALIBRATOR_ID")%>
      	    </SELECT>
      	  </TD>
        </TR>                
      	<TR>
      	  <TD CLASS=SMALLBOLD>Revision:</TD>
      	  <TD CLASS=SMALLBOLD><INPUT CLASS=SMALL TYPE=TEXT NAME="Revision" SIZE=10 MAXLENGTH=50 VALUE="<%=ProcValue(rsProcedure, "REVISION")%>" LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);"></TD>
        </TR>                
        <TR>
      	  <TD CLASS=SMALLBOLD>Procedure Type:</TD>
      	  <TD CLASS=SMALLBOLD>
      	    <SELECT NAME="Types" CLASS=SMALL>
        		<%=GetTypeInfo("TYPES", rsProcedure, "TYPE_ID")%>
      	    </SELECT>
      	  </TD>
        </TR>                
      	<TR>
      	  <TD CLASS=SMALLBOLD>5500 CAL Ready:</TD>
      	  <TD CLASS=SMALLBOLD>
      	    <INPUT TYPE=RADIO NAME="b5500Cal_Ready" VALUE="1" <%=RADIOVALUE(RSPROCEDURE, "5500CAL_READY", "YES")%>>Yes
      	    <INPUT TYPE=RADIO NAME="b5500Cal_Ready" VALUE="0" <%=RADIOVALUE(RSPROCEDURE, "5500CAL_READY", "NO")%>>No
      	  </TD>
        </TR>                
      	<TR>
      	  <TD CLASS=SMALLBOLD>Restricted:</TD>
      	  <TD CLASS=SMALLBOLD>
      	    <INPUT TYPE=RADIO NAME="Restricted" VALUE="1" <%=RADIOVALUE(RSPROCEDURE, "RESTRICTED", "YES")%>>Yes
      	    <INPUT TYPE=RADIO NAME="Restricted" VALUE="0" <%=RADIOVALUE(RSPROCEDURE, "RESTRICTED", "NO")%>>No
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If [Yes]&nbsp;&nbsp;Date&nbsp;
            <% if isdate(ProcValue(rsProcedure, "RESTRICTEDDATE")) then %>
            <INPUT TYPE=TEXT NAME="RestrictedDate" VALUE="<%=FormatDateTime(ProcValue(rsProcedure, "RESTRICTEDDATE"),"0")%>" SIZE=10 MAXLENGTH=10 CLASS=SMALL LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);">
            <% else %>
            <INPUT TYPE=TEXT NAME="RestrictedDate" VALUE="<%=ProcValue(rsProcedure, "RESTRICTEDDATE")%>" SIZE=10 MAXLENGTH=10 CLASS=SMALL LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);">
            <% end if %>      
            &nbsp;&nbsp;Reason&nbsp;
            <INPUT TYPE=TEXT NAME="RestrictedNote" VALUE="<%=ProcValue(rsProcedure, "RESTRICTEDNOTE")%>" SIZE=37 MAXLENGTH=50 CLASS=SMALL LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);">
      	  </TD>
        </TR>                
        <TR>
      	  <TD CLASS=SMALLBOLD>Sources:</TD>
      	  <TD CLASS=SMALLBOLD>
      	    <SELECT NAME="Sources" CLASS=SMALL LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);">
      		  <%=GetTypeInfo("SOURCES", rsProcedure, "SOURCE_ID")%>
      	    </SELECT>
      	  </TD>
        </TR>                
        <TR>
      	  <TD CLASS=SMALLBOLD>Price Point:</TD>
      	  <TD CLASS=SMALLBOLD>
	          <SELECT NAME="PricePoint" CLASS=SMALL LANGUAGE="JavaScript" ONCHANGE="UnHighlight(this);">
      		  <%=GetTypeInfo("PRICE POINTS", rsProcedure, "PRICE_POINT_ID")%>
      	    </SELECT>
      	  </TD>
        </TR>                
        <TR>
      	  <TD CLASS="SmallBold" VALIGN="Top">
      	    Zip File:
      	  </TD>
      	  <TD CLASS="Small" VALIGN=TOP>
      	    <%
            if strNewRecord="FALSE" or strAction="CLONE" then
      	  		response.write "<input type=""checkbox"" name=""KeepCurrentFileName"" CHECKED>&nbsp;&nbsp;"
              response.write "<SPAN class=""smallboldred"">Keep Current File Name:"
              response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
              response.write "<SPAN CLASS=SMALLBOLD>" & ProcValue(rsProcedure, "ZIPFILENAME") & "</SPAN></SPAN><BR>"
        		end if
        	  %>
      	    <INPUT CLASS=SMALL TYPE="CHECKBOX" NAME="File_Not_Available">&nbsp;&nbsp;<SPAN CLASS="smallboldred">Mark as N/A</SPAN><BR>
        	  &nbsp;&nbsp;<B>or</B>&nbsp;&nbsp;&nbsp;<SPAN CLASS=SMALLBoldRed>Upload New Zip File:</SPAN>&nbsp;&nbsp;&nbsp;<INPUT CLASS=SMALL TYPE="FILE" NAME="File_Name" SIZE="30" MAXLENGTH="50">
      	    <INPUT TYPE="HIDDEN" NAME="File_Name_Current" VALUE="<%=ProcValue(rsProcedure, "ZipFileName")%>">
      	  </TD>
        </TR>
        <TR>
        	<TD CLASS="SmallBold" VALIGN="top">Description:</TD>
      	  <TD CLASS="Small">
		        <TEXTAREA CLASS=SMALL NAME="Description" COLS=74 ROWS=8><%=ProcValue(rsProcedure, "DESCRIPTION")%></TEXTAREA>
        	</TD>
        </TR>
    	  <TR>
	        <TD CLASS=SMALLBOLD WIDTH="20%">ID:</TD>
      	  <TD CLASS=SMALLBOLD WIDTH="80%">
	          <%=ProcValue(rsProcedure, "PROCEDURE_ID")%>
      	  </TD>
        </TR>                
    
        <TR>
        	<TD CLASS="SmallBold" VALIGN="top">Last Update by:</TD>
      	  <TD CLASS="Small">
            <%
            if not isblank(ProcValue(rsProcedure, "UpdateBy")) then
    
              SQLUser = "SELECT FirstName, Lastname from UserData where ID=" & ProcValue(rsProcedure, "UpdateBy")
              Set rsUser = Server.CreateObject("ADODB.Recordset")
              rsUser.Open SQLUser, conn, 3, 3
      
              if not rsUser.EOF then
                response.write rsUser.Fields("FirstName").Value & " " & rsUser.Fields("LastName").Value
              else
                response.write "Met/Support Staff"
              end if
              rsUser.close
              set rsUser = nothing
      
            else
              response.write "Initial Record"
            end if
          
            if not isblank(ProcValue(rsProcedure, "UpdateDate")) then
              if CDate(ProcValue(rsProcedure, "UpdateDate")) <> CDate(ProcValue(rsProcedure, "CreateDate")) then
                response.write "&nbsp;&nbsp;&nbsp;on&nbsp;&nbsp;&nbsp;" & FormatDateTime(ProcValue(rsProcedure, "UpdateDate"),"0")
              end if
            end if
            %>
      	  </TD>
        </TR>

        <TR>
          <TD COLSPAN=2 BGCOLOR="BLACK">
            <TABLE WIDTH="100%">
              <TR>
                <TD WIDTH="20%">
									<INPUT CLASS=NAVLEFTHIGHLIGHT1 TYPE="Submit" VALUE=" Delete " onclick="return verifyDelete();" CLASS=NAVLEFTHIGHLIGHT1 ID="Delete" NAME="DoWhat">
								</TD>
                <TD WIDTH="80%">
									<INPUT CLASS=NAVLEFTHIGHLIGHT1 TYPE="Submit" VALUE=" Save " NAME="DoWhat" CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&nbsp;&nbsp;
                  <%
                  response.write  "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" ONCLICK=""return document.location='metcal_procedures.asp?KeyWord=" & strSearchkeyword & "&Calibrator=" & strSearchCalibrator & "&FileName=" & strSearchFileName & "&CurrPage=" & CurrPage & "&ID=" & iID & "#PID" & iID & "'"" value=""Cancel"">"
                  %>
								</TD>
              </TR>
            </TABLE>
          </TD>   
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>

<% Call Table_End %>

</FORM>
</CENTER>


<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
response.flush

' --------------------------------------------------------------------------------------

function GetProcedureDetails(iID)

  Dim cmd, prm, rsProcedure

  Set cmd = Server.CreateObject("ADODB.Command")
  Set cmd.ActiveConnection = conn
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = "Admin_MetCal_Procedure_Get"

  Set prm = cmd.CreateParameter("@iID", adInteger, adParamInput, , cInt(iID))
  cmd.Parameters.Append prm

  Set rsProcedure = Server.CreateObject("ADODB.Recordset")
  rsProcedure.CursorLocation = adUseClient
  rsProcedure.CursorType = adOpenDynamic
  rsProcedure.open cmd

  set prm = nothing
  set cmd = nothing
  
  set GetProcedureDetails = rsProcedure

end function

' --------------------------------------------------------------------------------------

function GetTypeInfo(strSubCategory, rsProcedure, strFieldName)

  Dim cmd, prm, rsSubCategory, strOptionList, strOptionFlag
  Dim strTarget, strTargetDefault
  
  if IsObject(rsProcedure) then
  	if not rsProcedure.eof then
		  strTarget = rsProcedure(strFieldName)
	  end if
  end if

  Set cmd = Server.CreateObject("ADODB.Command")
  Set cmd.ActiveConnection = conn
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = "Admin_MetCal_Categories_GetList"

  Set prm = cmd.CreateParameter("@strSubCategory", adVarchar,adParamInput ,50, uCase(strSubcategory) & "")
  cmd.Parameters.Append prm

  Set rsSubCategory = Server.CreateObject("ADODB.Recordset")
  rsSubCategory.CursorLocation = adUseClient
  rsSubCategory.CursorType = adOpenDynamic
  rsSubCategory.open cmd

  set prm = nothing
  set cmd = nothing

  strOptionFlag = false
  do while not rsSubCategory.EOF
    if strTarget = rsSubCategory("Category_ID") then
      strOptionFlag = true
      exit do
    end if
    rsSubCategory.MoveNext
  loop
  
  if strOptionFlag = false then
    select case UCase(strSubCategory)
      case "AUTHORS"
        strTarget = 6
      case "SOURCES"
        strTarget = 76
      case "COMPANIES"
        strTarget = 12
      case "TYPES"
        strTarget = 69
    end select
  end if
  
  rsSubCategory.MoveFirst

  do while not rsSubCategory.EOF
   	strOptionList = strOptionList & "<OPTION VALUE=""" & rsSubCategory("Category_ID") & """ "
   	if CInt(strTarget) = CInt(rsSubCategory("Category_ID")) then
   		strOptionList = strOptionList & " SELECTED"
      strOptionFlag = true
   	end if
   	strOptionList = strOptionList & ">" & rsSubCategory("Description") & "</OPTION>" & vbCrLf
  	rsSubCategory.movenext
  loop

  if CInt(strOptionFlag) = CInt(false) then
    strOptionList = "<OPTION VALUE=""00"" SELECTED>Select from list</OPTION>" & vbCrLf & strOptionList
  else
    strOptionList = "<OPTION VALUE=""00"">Select from list</OPTION>" & vbCrLf & strOptionList  
  end if
  
  GetTypeInfo = strOptionList

end function

' --------------------------------------------------------------------------------------

function ProcValue(rsProcedure, strField)

  if IsObject(rsProcedure) then
  
    select case UCase(strField)
      case "ADJTHRESHOLD"
        if isblank(rsProcedure(strField)) then
          ProcValue = "70%"
        elseif instr(1,rsProcedure(strField),"%") > 0 then
        	ProcValue = Trim(rsProcedure(strField))
        else
        	ProcValue = Trim(rsProcedure(strField) & "%")
        end if
      case "REVISION"
        if isblank(rsProcedure(strField)) then
          ProcValue = "0"
        else
        	ProcValue = Trim(Replace(UCase(rsProcedure(strField)),"V ",""))
        end if  
      case else
        ProcValue = Trim(rsProcedure(strField))
    end select
    
  else
    
    select case UCase(strField)
      case "DATE"
        ProcValue = Date()
      case "ADJTHRESHOLD"
        ProcValue = "70%"
      case "REVISION"
        ProcValue = "0"
    end select
  
  end if  

end function

' --------------------------------------------------------------------------------------

function RadioValue(rsProcedure, strFieldName, strRadioValue)

  Dim strValue

  if IsObject(rsProcedure) then
  	strValue = rsProcedure(strFieldName)
	  if UCase(strRadioValue) = "YES" then
		  if strValue <> 0 then
			  RadioValue = "CHECKED"
  		end if
	  elseif UCase(strRadioValue) = "NO" then
		  if strValue = 0 then
			  RadioValue = "CHECKED"
  	  end if
	  end if
  end if
  
end function

' --------------------------------------------------------------------------------------
%>

<SCRIPT LANGUAGE="JavaScript">

var ErrorMsg = "";
var FormName = document.<%=FormName%>
var DeleteOverride = false;

function verifyDelete(){
	var bValidate = confirm("Are you sure you want to delete this procedure?");
	if (bValidate == true){
    DeleteOverride = true;
		return (true);
	}
  else{
		return (false);
	}
}

function UnHighlight(myField) {
  myField.style.backgroundColor = "#FFFFFF";
}  

function CheckRequiredFields() {

  if (DeleteOverride == true) {
    return (true);
  }
  else {
    var ErrorMsg = "";
    var RadioChecked = 0;
    var RadioValue = "";
    var ctr = 0;
    var TestName = "";
    var TestRadio = "";
    var CheckSts;
        
    // Instrument
    if (FormName.Instrument.value.length == 0) {
        FormName.Instrument.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "Instrument Title\r\n";        
    }
  
    // Adjustment Threshold
    if (FormName.AdjThreshold.value.length == 0) {
      FormName.AdjThreshold.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "Adjustment Threshold Value\r\n";        
    }
    
    // Author
    CheckSts = false;
    TestName = "Author";
    TestRadio = FormName.Author;
    for (ctr=0; ctr < TestRadio.length; ctr++) {
      if (TestRadio[ctr].selected == true && TestRadio.value != "00") {
        CheckSts = true;
        break
      }  
    }
    if (CheckSts == false) {
      TestRadio.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "You have not selected a " + TestName + "\r\n";        
    }
    
    // Company
    CheckSts = false;
    TestName = "Company";
    TestRadio = FormName.Company;
    for (ctr=0; ctr < TestRadio.length; ctr++) {
      if (TestRadio[ctr].selected == true && TestRadio.value != "00") {
        CheckSts = true;
        break
      }  
    }
    if (CheckSts == false) {
      TestRadio.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "You have not selected a " + TestName + "\r\n";        
    }
  
    // Primary Calibrator
    CheckSts = false;
    TestName = "Primary Calibrator";
    TestRadio = FormName.PrimCalibrators;
    for (ctr=0; ctr < TestRadio.length; ctr++) {
      if (TestRadio[ctr].selected == true && TestRadio.value != "00") {
        CheckSts = true;
        break
      }  
    }
    if (CheckSts == false) {
      TestRadio.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "You have not selected a " + TestName + "\r\n";        
    }
  
    // Procedure Type
    CheckSts = false;
    TestName = "Procedure Type";
    TestRadio = FormName.Types;
    for (ctr=0; ctr < TestRadio.length; ctr++) {
      if (TestRadio[ctr].selected == true && TestRadio.value != "00") {
        CheckSts = true;
        break
      }  
    }
    if (CheckSts == false) {
      TestRadio.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "You have not selected a " + TestName + "\r\n";        
    }
    
    // Source
    CheckSts = false;
    TestName = "Source";
    TestRadio = FormName.Sources;
    for (ctr=0; ctr < TestRadio.length; ctr++) {
      if (TestRadio[ctr].selected == true && TestRadio.value != "00") {
        CheckSts = true;
        break
      }  
    }
    if (CheckSts == false) {
      TestRadio.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "You have not selected a " + TestName + "\r\n";        
    }
    
    // Price Point
    CheckSts = false;
    TestName = "Price Point";
    TestRadio = FormName.PricePoint;
    for (ctr=0; ctr < TestRadio.length; ctr++) {
      if (TestRadio[ctr].selected == true && TestRadio.value != "00") {
        CheckSts = true;
        break
      }  
    }
    if (CheckSts == false) {
      TestRadio.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "You have not selected a " + TestName + "\r\n";        
    }
  
    // Creation Date
    if (FormName.Date.value.length == 0) {
      FormName.Date.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "Missing Creation Date\r\n";
    }
    else if (! IsDate(FormName.Date.value)) {
      FormName.Date.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "Invalid Creation Date (Use: mm/dd/yyyy)\r\n";
    }
    
    // Revision
    if (FormName.Revision.value.length == 0) {
      FormName.Revision.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "Revision Value Missing or use 0\r\n";        
    }
  
    // Revision
    if (!IsNumeric(FormName.Revision.value)) {
      FormName.Revision.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "Invalid Revision Value (must be numeric) or use 0\r\n";        
    }


    // 5500 Cal Ready
    RadioChecked = 0;
    TestName = "5500 CAL Ready";
    TestRadio = FormName.b5500Cal_Ready;
    for (ctr=0; ctr < TestRadio.length; ctr++) {
      if (TestRadio[ctr].checked) {
        RadioChecked = 1;
        break;
      }
    }
    if (RadioChecked == 0) {
      ErrorMsg = ErrorMsg + "Checkmark for " + TestName + "\r\n";
    }
  
    // Restricted
    RadioChecked = 0;
    RadioValue = 0;
    TestName = "Restricted";
    TestRadio = FormName.Restricted;
    for (ctr=0; ctr < TestRadio.length; ctr++) {
      if (TestRadio[ctr].checked) {
        RadioValue = TestRadio[ctr].value;
        RadioChecked = 1;
        break;
      }
    }
    if (RadioChecked == 0) {
      ErrorMsg = ErrorMsg + "Checkmark for " + TestName + "\r\n";
    }
    
    // Restricted Date, Note
    if (RadioValue == "1") {
      if (FormName.RestrictedDate.value.length == 0) {
        FormName.RestrictedDate.style.backgroundColor = "#FFB9B9";
        ErrorMsg = ErrorMsg + "Missing Restricted Date\r\n";
      }
      if (! IsDate(FormName.RestrictedDate.value)) {
        FormName.RestrictedDate.style.backgroundColor = "#FFB9B9";
        ErrorMsg = ErrorMsg + "Invalid Restricted Date (Use: mm/dd/yyyy)\r\n";
      }
      if (FormName.RestrictedNote.value.length == 0) {
        FormName.RestrictedNote.style.backgroundColor = "#FFB9B9";    
        ErrorMsg = ErrorMsg + "Restricted Reason Note\r\n";
      }
    }   
    
    // Error Message
    if (ErrorMsg.length) {
      ErrorMsg = "Please Correct the Missing or Invalid Information Listed Below:\r\n\n" + ErrorMsg;
      alert(ErrorMsg);
      return (false);
    }
    
    // Check OK
    else {
      return (true);
    }
  }
}

function IsNumeric(sText){
   var ValidChars = "0123456789.";
   var IsNumber = true;
   var Char;
 
   for (i = 0; i < sText.length && IsNumber == true; i++) { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) {
         IsNumber = false;
      }
   }
   return IsNumber;
}

function isInteger(s){
	var i;
    for (i = 0; i < s.length; i++){   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
	var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){   
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
	// February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
		if (i==2) {this[i] = 29}
   } 
   return this;
}

var minYear=1900;
var maxYear=2100;

function IsDate(dtStr){

  var dtCh = "/";
	if (dtStr.indexOf("-") != -1) {
    dtCh = "-";
  }

	var daysInMonth = DaysArray(12)
	var pos1=dtStr.indexOf(dtCh)
	var pos2=dtStr.indexOf(dtCh,pos1+1)
	var strMonth=dtStr.substring(0,pos1)
	var strDay=dtStr.substring(pos1+1,pos2)
	var strYear=dtStr.substring(pos2+1)
	strYr=strYear
	if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
	if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
	for (var i = 1; i <= 3; i++) {
		if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
	}
	month=parseInt(strMonth)
	day=parseInt(strDay)
	year=parseInt(strYr)
	if (pos1==-1 || pos2==-1){
		return false
	}
	if (strMonth.length<1 || month<1 || month>12){
		return false;
	}
	if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
		return false;
	}
	if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
		return false;
	}
	if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
		return false;
	}
  return true;
}

  
</SCRIPT>