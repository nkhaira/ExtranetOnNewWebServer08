<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/connections/connection_email_new.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_EMail.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Author:     P. Barbee and Kelly Whitlock
' Date:       11/15/2000
'             Dev
' --------------------------------------------------------------------------------------

response.buffer = true

Dim bShowSwitch,bShowSearch,bSentMail,g_sDisplayType
bShowSwitch = False
bShowSearch = False
bSentMail = False
g_sDisplayType = "Dist" ' "Rep" is the other valid value

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

' after initial testing I'll re-enable the security model

' the security module ensures we have Session("logon_user") and Site_id we will
' make a custom call to return more fields:
' source_system
' customer_number
' for right now set them hard:

if Session("Logon_user") <> "" then
	%>
	<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
	<%
else
	site_id = 3
end if
%>
<!-- #include virtual="/SW-Common/SW-Site_Information.asp"-->
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
    
  if user_TypeCode = 2 then
      g_sDisplayType = "Rep"
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
		strCustomerNum = "0"
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

if request.form("customer_Number") <> "" then
    ' this implies the ChangeCompany function was used
	strCustomerNum = request.form("customer_Number")
	mydebug = False
    g_sDisplayType = "Dist"
	
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
	cmd.Parameters.Append cmd.CreateParameter("@iCust",advarchar,adParamInput,15,strCustomerNum)
	set dbRS = cmd.Execute
	set cmd = nothing
	
	if dbRS.EOF then
		strCustomerName = "Not found"
	else
		strCustomerName = dbRS("Customer_Name")
	end if
	set dbRS = Nothing
elseif request.form("Rep_Number") <> "" then
    ' this implies the ChangeRep function was used
	strCustomerNum = request.form("Rep_Number")
	mydebug = False
    g_sDisplayType = "Rep"
	
	if request.form("srcsystem") <> "" then
		strSrcSystem = request.form("srcsystem")
	else
		strSrcSystem = "ORA"
	end if
	
	' if we're coming in this way we'll want to know the customer name
	
	sql = "SELECT TOP 1 Customer_Name" & vbcrlf &_
	    "FROM  Orders" & vbcrlf &_
	    "WHERE Source_System = '" & strSrcSystem & "'" & vbcrlf &_
	    "AND Salesperson_Rep = '" & strCustomerNum & "'"
    
	set dbRS = conn.Execute(sql)
	
	if dbRS.EOF then
		strCustomerName = "Not found"
	else
		strCustomerName = dbRS("Customer_Name")
	end if
	set dbRS = Nothing
elseif Request.Form("do_mail") = 1 then
	Send_admin_mail
	bSentMail = True
	'response.redirect Session("BackURL")
end if

' grant permissions to subgroups
' ordad - get everything
' order - get search functions
if Instr(subgrp,"ORDAD") > 0 then
	bShowSwitch = True
	bShowSearch = True
elseif Instr(subgrp,"ORDER") > 0 then
	bShowSearch = True
end if

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

strTitle = "Order Inquiry"

Screen_Title    = Site_Description & " - " & Translate(strTitle,Alt_Language,conn) & " - " & Translate("Search",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) &_
                  "<BR><FONT CLASS=SmallBoldGold>" & _
                  Translate("Order Inquiry",Login_Language,conn) & " - " & Translate("Search",Login_Language,conn) & "</FONT>"
Top_Navigation  = False 
Side_Navigation = True
Content_Width   = 95  ' Percent
BackURL = Replace(Session("BackURL"),"CID=9007","CID=9000")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->
<%

' Log the activity

ActivitySQL = "INSERT INTO Activity" & vbcrlf &_
              "(Account_ID,Site_ID,Session_ID,View_Time,CID,SCID,Language,Method,Calendar_ID,Region,Country)" & vbcrlf &_
              "Values(" &_
              Login_ID & "," &_
              Site_ID & "," &_
              Session("Session_ID") & "," &_
              "getdate()," &_
              "9007," &_
              "0," &_
              "'" & Login_Language & "'," & _
              "0," & _
              "101," &_
              user_region & "," &_
              "'" & Session("Login_Country") & "')"

	conn.Execute (ActivitySQL)

response.write "<SPAN CLASS=Heading3>" & Translate("Order Inquiry - Search",Login_Language,conn) & "</SPAN><BR>"
response.write "<BR>"

' Menu Bar

with response

  .write "<FORM NAME=""Menu_Bar"">" & vbCrLf
  Call Table_Begin
  .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Home"" VALUE=""" & " " & Translate("Home",Login_Language,conn) & """ "
  .write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & BackURL & "'; return false;"" TITLE=""Return to Order Inquiry Select"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"

  if bShowSwitch then
      .write "&nbsp;&nbsp;"
      .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""OICC"" VALUE=""" & " " & Translate("Change Company",Login_Language,conn) & """ "
      .write "LANGUAGE=""Javascript"" ONCLICK=""location.href='/sw-common/SW-Order_Inquiry_Change_Company.asp'; return false;"" TITLE=""Change Company"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
  end if

  if strCustomerNum <> "0" then
    .write "&nbsp;&nbsp;"
    .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Help"" VALUE=""" & " " & Translate("Instructions",Login_Language,conn) & """ "
    .write "LANGUAGE=""Javascript"" ONCLICK=""location.href='/sw-common/SW-Order_Inquiry_Form.asp#Instruct'; return false;"" TITLE=""Instructions"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
  end if  

  Call Table_End
  .write "</FORM>" & vbCrLf

  .write "<P>"
  
end with

' KDW Temporary Override
if user_TypeCode = 2 then
  response.write "Sales Representatives: please note that we are working on a "
  response.write "capability for the Order Inquiry function that will allow you to "
  response.write "look up order status for all of the distributors you represent. "
  response.write "Until that is completed, <U>you will not have access</U> to the "
  response.write "Order Inquiry function.  Your distributors may look up their own "
  response.write "order status on the Partner Portal in the meantime."
  response.write "<P>"
end if

if 1 = 2 then ' no longer necessary (11/13/02) - left it here as example of how to do this
if site_id = 14 then 'custom message for Pomona
    with Response
        .write "<SPAN CLASS=""Small"">We are currently transferring order history from the old"
        .write " business system to the new.  During this process the information here may be "
        .write "incomplete.  "
        .write "We currently have all unshipped orders in the new system.  "
        .write "We are adding orders which shipped prior to 5/26/2002.  "
        .write "This message will change when the status changes.</span><P>" & vbcrlf
    end with
end if
end if

Response.write "<SPAN CLASS=SmallBold>" & Translate("Customer",Login_Language,conn) & ":&nbsp;&nbsp;" & strCustomerName
if bShowSwitch then
  response.write "&nbsp;&nbsp;[" & strCustomerNum & "]"
  response.write "&nbsp;&nbsp;[" & strSrcSystem & "]"
end if
response.write "</SPAN><P>" & vbcrlf

'get the date this data is as of
	
set cmd1 = Server.CreateObject("ADODB.Command")
cmd1.ActiveConnection = conn
cmd1.CommandType = adCmdStoredProc
cmd1.CommandText = "Order_GetUpload"

cmd1.Parameters.Append cmd1.CreateParameter("@source",adVarChar,adParamInput,8,strSrcSystem)
cmd1.Parameters.Append cmd1.CreateParameter("@strLang",adVarChar,adParamInput,50,strLang)
set dbRS = cmd1.Execute
set cmd1 = nothing

if not dbRS.EOF then
	with Response
		.write "<SPAN CLASS=Small>" & Translate("Order data was refreshed",Login_Language,conn) & " "
		.write Replace(Replace(dbRS("Fupload_date"),"AM"," AM"),"PM"," PM") & " PST</SPAN><P>" & vbcrlf
	end with
end if
set dbRS = nothing


if mydebug then
	Response.write "<P>Customer Name " & strCustomerName & vbcrlf
	Response.write "<BR>Customer Number " & strCustomerNum
	Response.write "  System " & strSrcSystem & vbcrlf
end if
	
set cmd1 = Server.CreateObject("ADODB.Command")
cmd1.ActiveConnection = conn
cmd1.CommandType = adCmdStoredProc
cmd1.Parameters.Append cmd1.CreateParameter("@strLang",adVarChar,adParamInput,50,strLang)
'if g_sDisplayType = "Rep" then
'    cmd1.CommandText = "Order_GetQuickRepOrder"
'    cmd1.Parameters.Append cmd1.CreateParameter("@strCust",adVarChar,adParamInput,32,strCustomerNum)
'else
    cmd1.CommandText = "Order_GetQuickOrder"
    cmd1.Parameters.Append cmd1.CreateParameter("@strCust",advarchar,adParamInput,32,strCustomerNum)
'end if
' create all the parameters we want
cmd1.Parameters.Append cmd1.CreateParameter("@source",adVarChar,adParamInput,8,strSrcSystem)
set dbRS = cmd1.Execute
set cmd1 = nothing

pocnt = Cint(dbRS("mycnt")) + 2

if pocnt > 15 then 
	pocnt = 15
elseif pocnt = 2 then
	pocnt = 0
end if

if bSentMail then	
	strA = "Your request has been mailed to the account administrator.  Choose [Home] from " &_
		"the left side menu to return to the main menu."
		
	Response.write "<P class=""medium"">" & Translate(strA,Login_Language,conn) & vbcrlf
  
'KDW Temporary Override  user_TypeCode addition  
elseif strCustomerNum = "0" and user_TypeCode <> 2 then
	set dbRS = nothing
	' this is when they don't yet have a customer number
	if bShowSwitch then
		strA = "Choose ""Change Company"" to select the company for your inquiry"
	else
		strA = "Your profile must be updated to include the requisite " &_
			"information for this function.  Choose [Request] to have the administrator " &_
			"make that update."
		
		strB = "Our system currently supports orders placed directly with Fluke corporate " &_
			"headquarters in Everett, WA, USA.  If you place your orders with a different " &_
			"Fluke entity check back soon."
	end if
	
	Session("BackURL") = BackURL
	Response.write "<P class=""medium"">" & Translate(strA,Login_Language,conn) & vbcrlf
	if not bShowSwitch then
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
	
elseif pocnt = 0 then
	'Response.write "<P>NUm :" & strCustomerNum & ": Length = " & len(strCustomerNum)
	Response.write "<P>"
    Response.write "<SPAN CLASS=""MediumRed""><LI>"
	Response.write Translate("There are no current orders for this Company",Login_Language,conn)
    Response.write "</SPAN></LI>" & vbcrlf
	set dbRS = nothing
	
else
	with response
		.write "<FORM name=""oi_input"" ACTION=""SW-Order_Inquiry_Results.asp"" METHOD=""POST"" "
		.write "onsubmit=""return My_submit();"">" & vbcrlf
		.write "<input type=""hidden"" name=""srcsystem"" value=""" & strSrcSystem & """>" & vbcrlf
		.write "<input type=""hidden"" name=""customernum"" value=""" & strCustomerNum & """>" & vbcrlf
		.write "<input type=""hidden"" name=""inq_type"">" & vbcrlf
		.write "<input type=""hidden"" name=""BackURL"" value=""" & BackURL & """>" & vbcrlf
		.write "<input type=""hidden"" name=""bShowSwitch"" value=""" & CInt(bShowSwitch) & """>" & vbcrlf
		.write "<input type=""hidden"" name=""UserType"" value=""" & g_sDisplayType & """>" & vbcrlf
	end with
	
	if mydebug then
		Response.write "<input type=""hidden"" name=""debug"" value=""on"">" & vbcrlf
	end if

	with response
  
    .write "<DIV ALIGN=CENTER>"
    
    Call Table_Begin
    
		.write "<TABLE BORDER=1 CLASS=tablebackground CELLPADDING=0 CELLSPACING=0>" & vbcrlf
		.write "  <TR>" & vbcrlf
		.write "    <TD align=""center"">" & vbcrlf
		.write "      <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 CLASS=tablebackground WIDTH=""650"">" & vbcrlf
		.write "        <TR>" & vbcrlf
 		.write "         <TD CLASS=""SmallBold"" nowrap>"
		.write Translate("Your PO Number",Login_Language,conn) & ":</TD>" & vbcrlf
		.write "          <TD CLASS=""SmallBold"">"
		.write "          	<INPUT CLASS=""Medium"" TYPE=""Text"" NAME=""ponum"" SIZE=25>" & vbcrlf
		.write "          </TD>" & vbcrlf
		.write "		  <TD CLASS=""SmallItalic"">" & vbcrlf
		.write Translate("You may enter a comma separated list.",Login_Language,conn) & vbcrlf
	end with
	
	if bShowSearch then
		Response.write "<BR>"
		Response.write Translate("% is the wild card character for Search.",Login_Language,conn) & vbcrlf
	end if
	
	with response
		.write "		  </td>" & vbcrlf
		.write "        </TR>" & vbcrlf
		.write "		<TR>" & vbcrlf
		.write "			<TD class=SmallItalic>&nbsp;&nbsp;&nbsp;"
		.write Translate("or",Login_Language,conn) & "</td>" & vbcrlf
		.write "		</tr>" & vbcrlf
		.write "        <TR>" & vbcrlf
		.write "          <TD CLASS=SmallBold nowrap>"
		.write Translate("Fluke Sales Order Number",Login_Language,conn) & ":</TD>" & vbcrlf
		.write "          <TD CLASS=SmallBold>" & vbcrlf
		.write "          	<INPUT CLASS=Medium TYPE=""Text"" NAME=""ordernum"" SIZE=25>" & vbcrlf
		.write "          </TD>" & vbcrlf
		.write "		  <TD CLASS=""SmallItalic"">" & vbcrlf
		.write Translate("You may enter a comma separated list.",Login_Language,conn) & vbcrlf
	end with
	
	if bShowSearch then
		Response.write "<BR>"
		Response.write Translate("No wild cards in Order Number.",Login_Language,conn) & vbcrlf
	end if
	
	with response
		.write "		  </td>" & vbcrlf
		.write "        </TR>" & vbcrlf
		.write "		<TR>" & vbcrlf
		.write "		  <TD align=""center"" colspan=3>" & vbcrlf
		.write "		  	<input type=""submit"" value="""
		.write Translate("Retrieve",Login_Language,conn) & """ class=""NavLeftHighlight1"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">" & vbcrlf
	end with
	
	if bShowSearch then
		Response.write "			&nbsp;&nbsp;&nbsp;" & vbcrlf
		Response.write "		  	<input type=""button"" onclick=""Chk_search();"" value="""
		Response.write Translate("Search",Login_Language,conn) & """ class=""NavLeftHighlight1"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
	end if
	
	with response
		.write "			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf
		.write "		    <input type=""button"" onclick=""Clr_list();"" value="""
		.write Translate("Clear",Login_Language,conn) & """ class=""NavLeftHighlight1"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">" & vbcrlf
		.write "		  </td>" & vbcrlf
		.write "		</tr>" & vbcrlf
		.write "      </TABLE>" & vbcrlf
		.write "    </TD>" & vbcrlf
		.write "  </TR>" & vbcrlf
		.write "</TABLE>" & vbcrlf
		
    Call Table_End
    
    .write "<P>" & vbcrlf
	end with
	
	if bShowSearch then
		with response
			.write "<P><SPAN CLASS=SmallBold>" & Translate("Your Current Orders",Login_Language,conn)
			.write "*</SPAN> <SPAN CLASS=Small>("
			.write Translate("Orders not completely shipped are highlighted",Login_Language,conn)
			.write ")</span>" & vbcrlf
			.write "<BR>" & vbcrlf
      
      Call Table_Begin
      
			.write "<TABLE BORDER=1 CLASS=tablebackground CELLPADDING=0 CELLSPACING=0>" & vbcrlf
			.write "  <TR>" & vbcrlf
			.write "    <TD align=""center"">" & vbcrlf
			.write "		<TABLE WIDTH=""650"" CELLPADDING=4 CELLSPACING=2 BORDER=0 CLASS=tablebackground>" & vbcrlf
			.write "			<TR>" & vbcrlf
			.write "				<TD align=""center"">" & vbcrlf
			.write "					<select name=""order_list"" multiple size="""
			.write pocnt & """ ondblclick=""Chk_list();"" style=""font: normal 10pt Courier;"">" & vbcrlf
			.write "						<option value="""" style=""background-color: #CCCCCC"">"
			.write Translate("PO Num",Login_Language,conn)
		end with
		
		maxorder = Cint(dbRS("maxorder")) + 3
		maxpo = Cint(dbRS("maxpo")) + 3
		
		set dbRS1 = dbRS.NextRecordset
		
		tdiff = maxpo - len(Translate("PO Num",Login_Language,conn))
		for i=0 to tdiff
			Response.write "&nbsp;"
		next
		Response.write Translate("Fluke Num",Login_Language,conn)
		tdiff = maxorder - len(Translate("Fluke Num",Login_Language,conn))
		for i=0 to tdiff
			Response.write "&nbsp;"
		next
		Response.write Translate("Date Entered",Login_Language,conn) & "</option>" & vbcrlf
		
		do until dbRS1.EOF
			tpo = dbRS1("Customer_PO_Number")
			tord = dbRS1("Order_Number")
			if dbRS1("Completed") = "0" then
				obg = ""
			else
				obg = " style=""background-color: wheat"""
			end if
			Response.write "						<option" & obg & " value=""" & tord & """>" & tpo
			
			for i=0 to (maxpo - len(tpo))
				Response.write "&nbsp;"
			next
			Response.write tord
			for i=0 to (maxorder - len(tord))
				Response.write "&nbsp;"
			next
			Response.write dbRS1("Torder_date") & "</option>" & vbcrlf
			
			dbRS1.MoveNext
		loop
		dbRS1.Close
		set dbRS1 = nothing
		
		with Response
			.write "					</select>" & vbcrlf
			.write "				</td>" & vbcrlf
			.write "			<TR>" & vbcrlf
			.write "			  <TD colspan=2 align=center>" & vbcrlf
			.write "			  	<input type=""button"" onclick=""Chk_list2();"" value="""
			.write Translate("Retrieve",Login_Language,conn) & """ CLASS=""NavLeftHighlight1"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
			.write vbcrlf & "				&nbsp;&nbsp;&nbsp;" & vbcrlf
			.write "			  	<input type=""button"" onclick=""Chk_list3();"" value="""
			.write Translate("Retrieve All",Login_Language,conn) & """ CLASS=""NavLeftHighlight1"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
			.write vbcrlf & "			  </td>" & vbcrlf
			.write "			</tr>" & vbcrlf
			.write "		</table>" & vbcrlf
			.write "    </TD>" & vbcrlf
			.write "  </TR>" & vbcrlf
			.write "</TABLE>" & vbcrlf
      
      Call Table_End
      
		end with
	end if
  
  response.write "</DIV>" 
	
	strA = "Enter your &quot;PO Number&quot; or the &quot;Fluke Sales Order Number&quot; then " &_
	       "click on the [Retrieve] button to view the details of that order."
		   
	with response
		.write "</FORM>" & vbcrlf
		.write "</CENTER>" & vbcrlf
		.write "<P><FONT CLASS=Medium>" & vbcrlf
		.write "<A NAME=""Instruct""></A>" & vbcrlf
		.write "<DIV ALIGN=LEFT>" & vbcrlf
		.write "<TABLE BORDER=0 WIDTH=""700"" CELLPADDING=0 CELLSPACING=0>" & vbcrlf
		.write "<TR><TD CLASS=""Medium"">" & vbcrlf
		.write "<SPAN CLASS=SmallBold>" & Translate("Limitations",Login_Language,conn) & "</SPAN><P>"
		.write "<UL><LI>*" & Translate("Your Current Orders",Login_Language,conn) & " - " &  Translate("This application shows all open orders and closed orders less than 30 days old.",Login_Language,conn) & "</LI></UL>" & vbcrlf
		.write "<SPAN CLASS=SmallBold>" & Translate("Instructions",Login_Language,conn) & "</SPAN><P>"
		.write "<UL><LI>" & Translate(strA,Login_Language,conn) & "</LI>" & vbcrlf
	end with
	
	if bShowSearch then
		strAs = "Use the [Search] button when using wildcards."
		strB  = "Use the list of&nbsp;&nbsp;Your Current Orders:"
		strC  = "Double click on an order to immediately view the details or;"
		strD  = "Select multiple orders by using your [ctrl] or [shift] key, then click on [Retrieve] button."
		with Response
			.write "<LI>" & Translate(strAs,Login_Language,conn) & "</LI>" & vbcrlf & "<P>" & vbcrlf
			.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>" & Translate("or",Login_Language,conn) & vbcrlf & "</I><P>" & vbcrlf
			.write "<LI>" & Translate(strB,Login_Language,conn) & vbcrlf & "<UL>" & vbcrlf
			.write vbtab & "<LI>" & Translate(strC,Login_Language,conn) & "</LI>" & vbcrlf
			.write vbtab & "<LI>" & Translate(strD,Login_Language,conn) & "</LI>" & vbcrlf
			.write "</UL></LI><BR>" & vbcrlf
		end with
		
		if bShowSwitch then
			strE = "Use the [Change Company] link to investigate a different Company."
			Response.write "<LI>" & Translate(strE,Login_Language,conn) & "</LI>" & vbcrlf
		end if
	end if
	
	Response.write "</ul>" & vbcrlf & "</TD></TR></TABLE></DIV><P>" & vbcrlf
	%>

<script language="Javascript">
	function Clr_list() {
		var df = document.oi_input;
		df.ponum.value = '';
		df.ordernum.value = '';
	}
	
	function Chk_search() {
		var df = document.oi_input;
		df.ponum.value = df.ponum.value.replace(/^\s+/,"");
		df.ordernum.value = df.ordernum.value.replace(/^\s+/,"");
		if ((df.ponum.value.length) || (df.ordernum.value.length)) {
			df.inq_type.value = 'lk';
			df.submit();
			//alert('ponum ' + df.ponum.value + '\nordernum ' + df.ordernum.value);
		}
		else {
			alert('Please enter either your PO Number or a Fluke Sales Order Number');
		}
	}
	
	function My_submit() {
		var df = document.oi_input;
		df.ponum.value = df.ponum.value.replace(/^\s+/,"");
		df.ordernum.value = df.ordernum.value.replace(/^\s+/,"");
		if ((df.ponum.value.length) || (df.ordernum.value.length)) {
			df.inq_type.value = 'eq';
			return true;
			//alert('ponum ' + df.ponum.value + '\nordernum ' + df.ordernum.value);
		}
		else {
			alert('Please enter either your PO Number or a Fluke Sales Order Number');
			return false;
		}
	}
	
	function Chk_list() {
		var df = document.oi_input;
		var dfl = document.oi_input.order_list;
		
		for (var i=0;i<dfl.length;i++) {
			if (dfl.options[i].selected) {
				df.ponum.value = '';
				df.ordernum.value = dfl.options[i].value;
				df.inq_type.value = 'list';
				df.submit();
				//alert('ordernum ' + df.ordernum.value);
			}
		}
	}
	
	function Chk_list2() {
		var df = document.oi_input;
		var dfl = document.oi_input.order_list;
		
		if (dfl.selectedIndex > 0) {
			df.ponum.value = '';
			var poval = dfl.options[dfl.selectedIndex].value;
			for (var i=dfl.selectedIndex+1;i<dfl.length;i++) {
				if (dfl.options[i].selected) {
					poval += ',' + dfl.options[i].value;
				}
			}
			df.ordernum.value = poval;
			df.inq_type.value = 'list';
			df.submit();
			//alert('ordernum(s) ' + df.ordernum.value);
		}
	}
	
	function Chk_list3() {
		var df = document.oi_input;
		var dfl = document.oi_input.order_list;
		var poval = dfl.options[1].value;
		
		if (df.length > 2) {
			for (var i=2;i<dfl.length;i++) {
				poval += ',' + dfl.options[i].value;
			}
		}
		df.ordernum.value = poval;
		df.inq_type.value = 'list';
		df.submit();
	}
</script>

<%
end if

%>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
response.end

'--------------------------------------------------------------------------------------

sub Table_Begin()
    response.write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" CLASS=TableBorder>" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif""            BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLr
    response.write "  <TR>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD VALIGN=""top"">" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Table_End()
    response.write "    </TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_BL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_BR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLf
    response.write "</TABLE>" & vbCrLf
end sub  

'--------------------------------------------------------------------------------------

sub Send_admin_mail
	
	'Set Mailer = CreateObject("SMTPsvg.Mailer")
	'adding new email method
	%>
	
	<%

	'Mailer.ReturnReceipt = false
	'Mailer.ConfirmRead = false
	'Mailer.WordWrapLen = 80
	'Mailer.QMessage = False
	'Mailer.ClearAttachments
	'Mailer.RemoteHost = GetEmailServer()
	' send mail to Peter
	
	'Mailer.ClearRecipients
	'''Mailer.AddRecipient "Kelly Whitlock","Kelly.Whitlock@fluke.com"
	
	' who to mail to
	sSql = "SELECT aa.email_site_admin,ud.firstname,ud.lastname,ud.email" & vbcrlf &_
		"FROM approvers_account aa" & vbcrlf &_
			"inner join userdata ud on aa.approver_id = ud.id" & vbcrlf &_
		"WHERE aa.site_id = " & site_id & vbcrlf &_
			"and aa.region = " & user_region

  set dbRS = conn.Execute(sSql)
	
	if dbRS.EOF then
		'Mailer.AddRecipient Site_Admin_Name,Site_Admin_Email
		msg.To = """" & Site_Admin_Name & """" & Site_Admin_Email
	else
		admin_name = dbRS("firstname") & " " & dbRS("Lastname")
		'Mailer.AddRecipient admin_name,dbRS("email")
		msg.To = """" & admin_name & """" & dbRS("email")
		
		if dbRS("email_site_admin") < 0 then
			'Mailer.AddRecipient Site_Admin_Name,Site_Admin_Email
			msg.To = """" & Site_Admin_Name & """" & Site_Admin_Email
		end if
	end if
	
	set dbRS = nothing
	
	'Mailer.FromName    = "Extranet"
	'Mailer.FromAddress = "webmaster@fluke.com"
	'Mailer.WordWrap    = True
	'Mailer.ContentType = "text/plain"
	
	msg.From = """Extranet""" & "webmaster@fluke.com"

	strText = "The following user has requested their Fluke Customer Number and " &_
		"Business System be updated:" & vbcrlf & vbcrlf &_
		"UserID:    " & Login_ID & vbcrlf &_
		"LastName:  " & user_last & vbcrlf &_
		"FirstName: " & user_first & vbcrlf &_
		"Company:   " & strCustomerName & vbcrlf &_
		"Site:      " & Site_Description & vbcrlf
		
	'Mailer.ClearBodyText
	'Mailer.BodyText = strText

	msg.TextBody = strText
	
	'Mailer.Subject = "Extranet Order Inquiry Request"
	'Mailer.SendMail

	msg.Subject = "Extranet Order Inquiry Request"
	msg.Configuration = conf
		
	On Error Resume Next
	msg.Send
	If Err.Number = 0 then
		'Success
	Else
		'Fail
	End If

	'set Mailer = Nothing
end sub

'--------------------------------------------------------------------------------------
%>