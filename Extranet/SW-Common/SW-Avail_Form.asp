<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_estore.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->
<%

' --------------------------------------------------------------------------------------
' Author:     P. Barbee
' Date:       4/11/2002
' --------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

Dim bShowSwitch, bShowModels, bSentMail
bShowSwitch = False
bSentMail = False
bShowModels = False

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

' get additional user data
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "Order_GetUserInfo"
' create all the parameters we want
cmd.Parameters.Append cmd.CreateParameter("@login",adVarChar,adParamInput,50,Session("logon_user"))
cmd.Parameters.Append cmd.CreateParameter("@site",adInteger,adParamInput,,Site_id)

set dbRS = cmd.Execute
set cmd = nothing

if dbRS.EOF then
	response.write "Something is seriously wrong "
	disconnect_SiteWide
	response.end
else
	' deal with recordset in it's order
	Login_ID = dbRS("ID")
	
	user_TypeCode = cint(dbRS("Type_Code"))
	if user_TypeCode = 5 then
		bShowSwitch = True
        bShowSearch = True
	end if
	
	user_first = dbRS("FirstName")
	user_last = dbRS("LastName")
	subgrp = Ucase(dbRS("SubGroups"))
	strCustomerName = dbRS("Company")
	user_region = cint(dbRS("Region"))
	
	if len(dbRS("Fluke_ID")) then
		strCustomerNum = dbRS("Fluke_ID")
	else
		' maybe the user doesn't have a customer number
		strCustomerNum = 0
	end if
	
	strSrcSystem = dbRS("Business_System")
	strLang = dbRS("Description")
	if strLang = "Chinese (Simplified)" then
		strLang = "Simplified Chinese"
	elseif strLang = "Chinese (Traditional)" then
		strLang = "Traditional Chinese"
	elseif strLang = "Spanish(South American)" then
		strLang = "Spanish"
	elseif strLang = "Portuguese(South American)" then
		strLang = "Portuguese"
	end if
	
end if
set dbRS = Nothing
mydebug = False

' this block of code allows the customer selection page to work for internal users
if request.form("customer_Number") <> "" then
	strCustomerNum = request.form("customer_Number")
	mydebug = False
	
	if request.form("srcsystem") <> "" then
		strSrcSystem = request.form("srcsystem")
	else
		strSrcSystem = "ORA"
	end if
	
	' if we're coming in this way we'll want to know the customer name
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "Order_Company"
	' create all the parameters we want
	cmd.Parameters.Append cmd.CreateParameter("@source",adVarChar,adParamInput,8,strSrcSystem)
  on error resume next  
	cmd.Parameters.Append cmd.CreateParameter("@iCust",adInteger,adParamInput,,strCustomerNum)
  if Err.Number <> 0 then
    response.write "Customer Number contains alpha characters and is not compatiable with this system - Report to P. Barbee"
    response.end
  end if
  on error goto 0

	set dbRS = cmd.Execute

	set cmd = nothing
	
	if dbRS.EOF then
		strCustomerName = "Not found"
	else
		strCustomerName = dbRS("Customer_Name")
	end if
	set dbRS = Nothing
	'response.redirect Session("BackURL")
end if

' grant permissions to subgroups -- FUTURE functionality, comments are copy from Order_
' ordad - get everything
' order - get search functions
'if Instr(subgrp,"ORDAD") > 0 then
'	bShowSwitch = True
'	bShowSearch = True
'elseif Instr(subgrp,"ORDER") > 0 then
'	bShowSearch = True
'end if

Dim BackURL

BackURL = Session("BackURL")

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

strTitle = "Availability"

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate(strTitle,Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) &_
                	"<BR><SPAN CLASS=SmallBoldGold>" &_
                  Translate(strTitle,Login_Language,conn) & " - " & Translate("Search",Login_Language,conn) & "</SPAN>"
Top_Navigation  = False 
Side_Navigation = True
Content_Width   = 95  ' Percent
BackURL = Replace(Session("BackURL"),"CID=9008","CID=9000")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->
<%
' Log the activity

ActivitySQL = "INSERT INTO Activity	(Account_ID,Site_ID,Session_ID,View_Time,CID,SCID,Language,Region,Country)" & vbcrlf &_
          		"Values(" &_
          		Login_ID & "," &_
          		Site_ID & "," &_
          		Session("Session_ID") & "," &_
          		"getdate()," &_
          		"9008," &_
          		"0," &_
          		"'" & Login_Language & "'," &_
              user_region & "," &_
              "'" & Session("Login_Country") &  "')"

conn.Execute (ActivitySQL)

' start writing content

response.write "<SPAN CLASS=Heading3>" & strTitle & " - " & Translate("Search",Login_Language,conn) & "</SPAN><BR>"
response.write "<BR>"

response.write "<SPAN CLASS=""SmallBold"">"

with response

  .write "<FORM NAME=""Menu_Bar"">" & vbCrLf
  Call Table_Begin
  .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Home"" VALUE=""" & " " & Translate("Home",Login_Language,conn) & """ "
  .write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & BackURL & "'; return false;"" TITLE=""Return to Order Inquiry Select"">"

  if bShowSwitch then
      .write "&nbsp;&nbsp;"
      .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""OICC"" VALUE=""" & " " & Translate("Change Company",Login_Language,conn) & """ "
      .write "LANGUAGE=""Javascript"" ONCLICK=""location.href='/sw-common/SW-Order_Inquiry_Change_Company.asp?caller=avail'; return false;"" TITLE=""Change Company"">"
  end if

  if strCustomerNum <> "0" then
    .write "&nbsp;&nbsp;"
    .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Help"" VALUE=""" & " " & Translate("Instructions",Login_Language,conn) & """ "
    .write "LANGUAGE=""Javascript"" ONCLICK=""location.href='/sw-common/SW-Avail_Form.asp#Instruct'; return false;"" TITLE=""Instructions"">"
  end if  

  Call Table_End
  .write "</FORM>" & vbCrLf

  .write "<P>"
  
end with

Response.write "<P>" & Translate("Customer",Login_Language,conn) & ":&nbsp;&nbsp;" & strCustomerName
if bShowSwitch then
  response.write "&nbsp;&nbsp;[" & strCustomerNum & "]"
  response.write "&nbsp;&nbsp;[" & strSrcSystem & "]"
end if
response.write "</SPAN><P>" & vbcrlf

if bSentMail then	' they're coming back in after requesting mail be sent
	strA = "Your request has been mailed to the account administrator.  Choose [Home] from " &_
		"the left side menu to return to the main menu."
		
	Response.write "<P class=""medium"">" & Translate(strA,Login_Language,conn) & vbcrlf
    
elseif strCustomerNum = 0 then   ' this is when they don't yet have a customer number
	Session("BackURL") = BackURL
	
	if bShowSwitch then    ' it's a fluke employee
		strA = "Choose ""Change Company"" to select the company for your inquiry"
	    Response.write "<P class=""medium"">" & Translate(strA,Login_Language,conn) & vbcrlf
    elseif user_TypeCode = 2 then   ' it's a fluke sales rep, this message is bad
        with response
            .write "Sales Representatives: please note that we are working on a "
            .write "capability for the Order Inquiry function that will allow you to "
            .write "look up order status for all of the distributors you represent. "
            .write "Until that is completed, <U>you will not have access</U> to the "
            .write "Order Inquiry function.  Your distributors may look up their own "
            .write "order status on the Partner Portal in the meantime."
            .write "<P>"
        end with
	else    ' give the distributor without a Customer Number the chance to request one
		strA = "Your profile must be updated to include the requisite " &_
			"information for this function.  Choose [Request] to have the administrator " &_
			"make that update."
		
		strB = "Our system currently supports orders placed directly with Fluke corporate " &_
			"headquarters in Everett, WA, USA.  If you place your orders with a different " &_
			"Fluke entity check back soon."
        
		with Response
			.write "<P class=""medium"">" & Translate(strB,Login_Language,conn) & vbcrlf
			.write "<P><Form name=""oi_input"" method=""POST""><input type=""button"" value="" "
			.write Translate("Request",Login_Language,conn) & " "" onclick=""Do_req();"">" & vbcrlf
			.write "<input type=hidden name=""do_mail"">" & vbcrlf
			.write "</form>" & vbcrlf
			.write "<script language=""Javascript"">" & vbcrlf
			.write "	function Do_req() {" & vbcrlf
			.write "		var df = document.oi_input;" & vbcrlf
			.write "		df.do_mail.value = 1;" & vbcrlf
			.write "		df.submit();" & vbcrlf
			.write "	}" & vbcrlf
			.write "</script>" & vbcrlf
		end with
	end if
	
else    ' go ahead and build the form
    if site_id = 14 then
        strA = "Fluke Part Numbers are completely numeric.  Enter a <B>Pomona Part Number</b> (either numeric or alpha-numeric) in the Model field"
        response.write "<P class=""medium"">" & Translate(strA,Login_Language,conn) & vbcrlf
    end if
    
    with response
    	.write "<FORM NAME=""swavail"" ACTION=""/SW-Common/SW-Avail_Results.asp"" METHOD=""POST"">" & vbcrlf
    	.write "<input type=""hidden"" name=""BackURL"" value=""" & BackURL & """>" & vbcrlf
      .write "<input type=""hidden"" name=""srcsystem"" value=""" & strSrcSystem & """>" & vbcrlf
      .write "<input type=""hidden"" name=""customer_number"" value=""" & strCustomerNum & """>" & vbcrlf
      
      .write "<DIV ALIGN=CENTER>"
      
      Call Table_Begin
      
    	.write "      <TABLE WIDTH=""600"" CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="""
    	.write Contrast & """ WIDTH=""100%"">" & vbcrlf
    	.write "        <TR>" & vbcrlf
    	.write "          <TD CLASS=""MediumBold"" NOWRAP>"
    	.write Translate("Part Number",Login_Language,conn) & " :</TD>" & vbcrlf
    	.write "          <TD CLASS=MediumBold>" & vbcrlf
    	.write "            <INPUT TYPE=""text"" NAME=""part"" SIZE=""30"" CLASS=""Medium"">" & vbcrlf
    	.write "          </TD>" & vbcrlf
    	.write "        </TR>" & vbcrlf
    	.write "        <TR>" & vbcrlf
    	.write "          <TD CLASS=""MediumBold"">"
    	.write Translate("Model",Login_Language,conn) & " :</TD>" & vbcrlf
    	.write "          <TD CLASS=MediumBold>" & vbcrlf
    	.write "            <INPUT TYPE=""text"" NAME=""model"" SIZE=""30"" CLASS=""Medium"">" & vbcrlf
    	.write "          </TD>" & vbcrlf
    	.write "        </TR>" & vbcrlf
    end with
    
    if bShowModels then
        Connect_eStoreDatabase
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = eConn
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "Avail_GetModelGroups"
        set rs = cmd.Execute
        set cmd = nothing
        
        if not rs.EOF then
        	with response
        		.write "        <TR>" & vbcrlf
        		.write "          <TD CLASS=""MediumBold"">"
        		.write Translate("Model Group",Login_Language,conn) & " :</TD>" & vbcrlf
        		.write "          <TD CLASS=MediumBold>" & vbcrlf
        		.write "            <SELECT MULTIPLE NAME=""sel_group"" SIZE=""8"" CLASS=""Medium"">"
                .write vbcrlf
        		.write "              <OPTION></option>" & vbcrlf
        	end with
        	
        	do until rs.EOF
        		response.write "              <OPTION value=""" & rs("Model_Group") & """>"
        		response.write rs("Display_Name") & "</option>" & vbcrlf
        		rs.MoveNext
        	loop
            rs.Close
        	
        	with response
        		.write "            </SELECT>" & vbcrlf
        		.write "          </TD>" & vbcrlf
        		.write "        </TR>" & vbcrlf
        	end with
        end if
        set rs = nothing
        Disconnect_eStoreDatabase
    end if
    
    with response
    	.write "        <TR>" & vbcrlf
    	.write "          <TD ALIGN=""center"" BGCOLOR=""#666666"" colspan=""2"">" & vbcrlf
    	.write "            <INPUT TYPE=""button"" VALUE="""
    	.write Translate("Search",Login_Language,conn) & """ CLASS=""NavLeftHighlight1"""
    	.write " onclick=""My_submit();"">"
    	.write "			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf
    	.write "            <INPUT TYPE=""reset"" VALUE="""
    	.write Translate("Clear",Login_Language,conn) & """ CLASS=""NavLeftHighlight1"">" & vbcrlf
    	.write "          </TD>" & vbcrlf
    	.write "        </TR>" & vbcrlf
    	.write "      </TABLE>" & vbcrlf
    	.write "<INPUT TYPE=""HIDDEN"" NAME=""group"">" & vbcrlf

      Call Table_End
    	.write "</FORM>" & vbcrlf     
    	.write "</DIV>" & vbcrlf
      
      
    	.write "<P><P CLASS=Medium>" & vbcrlf
    	.write "<A NAME=""Instruct""></A>" & vbcrlf
    	.write "<DIV ALIGN=CENTER>" & vbcrlf
    	.write "<TABLE BORDER=0 WIDTH=""100%"" CELLPADDING=0 CELLSPACING=0>" & vbcrlf
    	.write "<TR><TD CLASS=""Medium"">" & vbcrlf
    	.write "<SPAN CLASS=SmallBold>" & Translate("Instructions",Login_Language,conn) & "</SPAN><P>"
    end with
    
    strA = "You may enter either or both Part Numbers and Models"
    strB = "When entering more than one Part Number or Model seperate the entries with a comma"
    strC = "Part Numbers are completely numeric 6 or 7 digit Fluke Part Numbers"
    strD = "The wild card character is * and can be placed anywhere in either a Part Number or Model"
    strE = "You may select multiple Model Groups by using your [ctrl] or [shift] key"
    strF = "<B>Pomona Part Numbers</b> are entered in the Model field"
    
    with Response
    	.write "<A NAME=""Instruct""></A>" & vbcrlf
    	.write "<UL>" & vbcrlf
    	.write vbtab & "<LI>" & Translate(strA,Login_Language,conn) & "</LI>" & vbcrlf
    	.write vbtab & "<LI>" & Translate(strB,Login_Language,conn) & "</LI>" & vbcrlf
    	.write vbtab & "<LI>" & Translate(strC,Login_Language,conn) & "</LI>" & vbcrlf
    	.write vbtab & "<LI>" & Translate(strD,Login_Language,conn) & "</LI>" & vbcrlf
      if site_id = 14 then .write vbtab & "<LI>" & Translate(strF,Login_Language,conn) & "</LI>" & vbcrlf
    end with
    
    if bShowModels then
    	response.write vbtab & "<LI>" & Translate(strD,Login_Language,conn) & "</LI>" & vbcrlf
    end if
    
    with Response
    	.write "</LI></UL><BR>" & vbcrlf
    	.write "</TD></TR></TABLE></DIV>" & vbcrlf
    	.write "</FONT>"
    end with
%>
<script language="Javascript">
	
	function My_submit() {
		var df = document.swavail;
		var GoodtoGo = true;
		var mesg;
        
<%
    if bShowModels then %>
		var sel = df.sel_group;
		
		df.group.value = '';
		for (var i=0;i<sel.options.length;i++) {
			if ((sel.options[i].selected) && (sel.options[i].value.length)) {
				if (df.group.value.length) {
					df.group.value += ',' + sel.options[i].value;
				}
				else {
					df.group.value = sel.options[i].value;
				}
			}
		}
		
		if ((! df.part.value.length) && (! df.model.value.length) && (! df.group.value.length)) {
			mesg = 'Please enter Part Number, Model Number, or select a Model Group';
			GoodtoGo = false;
		}<%
    else %>
		if ((! df.part.value.length) && (! df.model.value.length)) {
			mesg = 'Please enter Part Number or Model Number';
			GoodtoGo = false;
		}<%
    end if %>
		else if ((df.part.value.length) && (df.part.value.search(/[^\d,\s*]/) > -1)) {
			mesg = 'Part Number must be fully numeric (or *)';
			GoodtoGo = false;
		}
		
		if (GoodtoGo) {
			//alert('submit part and model');
			df.submit();
		}
		else {
			alert(mesg);
		}
		
	}
</script>
<%
end if
%>

<!--#include file="SW-Footer.asp"-->
<%
Call Disconnect_Sitewide

'--------------------------------------------------------------------------------------
' Functions and Subroutines
'--------------------------------------------------------------------------------------

sub Table_Begin()
    response.write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" CLASS=TableBorder>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLr
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "        <TD VALIGN=""top"">" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Table_End()
    response.write "        </TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_BL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_BR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "    </TABLE>" & vbCrLf
end sub

'--------------------------------------------------------------------------------------

sub Nav_Border_Begin()
    response.write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" CLASS=NavBorder>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_TR_corner.gif""<IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLr
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "        <TD VALIGN=""top"" CLASS=NavBorder>" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Nav_Border_End()
    response.write "        </TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_BL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_BR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "    </TABLE>" & vbCrLf
end sub   

'--------------------------------------------------------------------------------------
%>

<SCRIPT TYPE="text/javascript" LANGUAGE="JavaScript">
function PrintIt(){

  var NS = (navigator.appName == "Netscape");
  var VERSION = parseInt(navigator.appVersion);
  
  if (window.print) {
    window.print() ;  
  }
  else {
    var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
    document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
    WebBrowser1.ExecWB(6, 2);   //Use a 1 vs. a 2 for a prompting dialog box    WebBrowser1.outerHTML = "";  
  }
}
</SCRIPT>

