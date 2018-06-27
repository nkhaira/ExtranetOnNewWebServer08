<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/Include/Adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Author:     Peter Barbee
' Date:       01/09/2002
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim HomeURL
Dim BackURL

Dim Site_ID

Dim SQL
Dim bShow_Main

%>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_Login_Admin.asp"-->
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%
'site_id = Request("Site_ID")
'logon_user = Request("logon_user")
fluke_id = Request("fluke_id")
business_system = Request("bus_sys")

Bar_Tag = Translate("Branch Account Edit",Login_Language,conn)

' --------------------------------------------------------------------------------------
' Determine Site Code and Description based on Site_ID Number
' --------------------------------------------------------------------------------------

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code        = rsSite("Site_Code")     
Site_Description = rsSite("Site_Description")
Screen_Title     = rsSite("Site_Description") & " - " & Translate("Account Administrator",Alt_Language,conn)
Bar_Title        = rsSite("Site_Description") & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Account Administrator",Login_Language,conn) & "</FONT>"
Bar_Title        = Bar_Title & "<BR><FONT CLASS=SmallBoldGold>" & Translate("List",Login_Language,conn) & " / " & Translate("Edit",Login_Language,conn) & " " & Translate("Group",Login_Language,conn) & ": " & Bar_Tag & "</FONT>"

Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<%

rsSite.close
set rsSite=nothing

' when we introduce distributor-based branch editing make sure those people
' have bShow_Main set to False
bShow_Main = True

response.write "<FONT CLASS=NormalBoldRed>"
select case Admin_Access
  case 2
    response.write Translate("Content Submitter",Login_Language,conn)
  case 4
    response.write Translate("Content Administrator",Login_Language,conn)
  case 6
    response.write Translate("Account Administrator",Login_Language,conn)
  case 8
    response.write Translate("Site Administrator",Login_Language,conn)
  case 9
    response.write Translate("Domain Administrator",Login_Language,conn)
end select

with response
	.write "</FONT><BR><FONT CLASS=MediumBold>" & Admin_FirstName & " " & Admin_LastName & "<BR>"
	.write Admin_Company & "</FONT><BR><BR>"

' put a tiny menu in
  Call Nav_Border_Begin
'	.write "<A HREF=""/" & site_code & "/default.asp?Site_ID=" & site_id & "&Language=" & Login_Language
'	.write "&NS=true&CID=9999&SCID=0&PCID=0&CIN=0&CINN=0"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;"
'	.write Translate("Logoff",Login_Language,conn) & "&nbsp;&nbsp;</A>&nbsp;&nbsp;"
end with

if bShow_Main then
	with response
		.write "<A HREF=""default.asp?Site_ID=" & site_id
		.write """ CLASS=NavLeftHighlight1>&nbsp;&nbsp;Main Menu&nbsp;&nbsp;</A>"
	end with
else

end if

if Request.Form("PostFlag") = "1" then
	Update_table
end if

' for now only admin_access > 5 (subgroup => account) accesses this so:
' this may need a change when we introduce distributor-based branch editing
if admin_access > 5 then
	with Response
		.write "&nbsp;&nbsp;"
    .write "<A HREF=""JavaScript:Sel_Comp();"" CLASS=""NavLeftHighlight1"">&nbsp;Change Company&nbsp;</A>"

	'	.write "&nbsp;&nbsp;&nbsp;<input type=""button"" value=""Refresh w/New Company"" "
	'	.write "class=""NavLeftHighlight1"" onclick=""Roll_Me();"">" & vbcrlf
  Call Nav_Border_End
	end with
  
	Create_form
else
  Call Nav_Border_End
  response.write "<P>"
	Response.write "<SPAN CLASS=SmallBoldRed>" & "You do not have access to this function" & "</SPAN>"
end if
%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide
Response.end

' ------------------------ end of main -------- subroutines below -------------

sub Update_table
	
	aids = split(Request.Form("sids"),", ")
	' Status and order enable based on the database
	strStatus = Request.Form("sStatus")
	strOrder = Request.Form("sOrder")
	' Status and order enable based on the Form elements
	oStatus = "|" & Replace(Request.Form("ostatus"),", ","|") & "|"
	oOrder = "|" & Replace(Request.Form("oenable"),", ","|") & "|"
	
	' create the command object we will use
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "Branch_Update"
	
	cmd.Parameters.Append cmd.CreateParameter("@userid", adInteger,adParamInput)
	cmd.Parameters.Append cmd.CreateParameter("@expire", adInteger,adParamInput)
	cmd.Parameters.Append cmd.CreateParameter("@order", adInteger,adParamInput)
	
	' we'll create a pair of tri-state variables, iStatus & iOrder
	' 0 => take no action
	' 1 => toggle off
	' 2 => toggle on
	for each uid in aids
		suid = "|" & uid & "|"
		iStatus = 0
		iOrder = 0
		if Instr(strStatus,suid) > 0 then ' the user is active in the database
			if Instr(oStatus,suid) = 0 then ' inactive on form
				iStatus = 1 ' toggle off
			end if
		else ' inactive in database
			if Instr(oStatus,suid) > 0 then ' active on form
				iStatus = 2 ' toggle on
			end if
		end if
		
		' same stuff for order
		if Instr(strOrder,suid) > 0 then
			if Instr(oOrder,suid) = 0 then
				iOrder = 1 ' toggle off
			end if
		else
			if Instr(oOrder,suid) > 0 then
				iOrder = 2 ' toggle on
			end if
		end if
		
		if (iStatus > 0) OR (iOrder > 0) then
			'Response.write uid & " call SP: " & iStatus & " " & iOrder & "<BR>" & vbcrlf
			cmd.Parameters("@userid").Value = uid
			cmd.Parameters("@expire").Value = iStatus
			cmd.Parameters("@order").Value = iOrder
			cmd.Execute
		end if
	next
	set cmd = Nothing
end sub

sub Create_form
	' first we need the javascript that will call change company and refresh this page
	%>
<SCRIPT language="JavaScript">
	function Sel_Comp() {
		var df = document.branch
		var href = '/sw-administrator/Company_select.asp?';
		href += 'site_id=' + df.site_id.value;
		href += '&logon_user=' + df.logon_user.value;
		href += '&Language=' + df.Language.value;
		var opts = 'status=no,scrollbars=yes,resizable=yes,toolbar=yes,links=no';
		var BE_Window = window.open(href,'BE_Window',opts);
		BE_Window.focus();
	}
	
	function Roll_Me() {
		if (document.branch.fluke_id.value.length) {
			document.branch.PostFlag.value = 0;
			document.branch.submit();
		}
		else {
			alert('There is not a current value for Customer Number.\nChoose "Change Company"');
		}
	}
	
	function Update() {
		document.branch.PostFlag.value = 1;
		document.branch.submit();
	}
</script>
	<%
	' now we need the basics of the form
	with Response
		.write "<form name=""branch"" method=""POST"">" & vbcrlf
		.write "<input type=""hidden"" name=""site_id"" value=""" & site_id & """>" & vbcrlf
		.write "<input type=""hidden"" name=""logon_user"" value=""" & logon_user & """>" & vbcrlf
		.write "<input type=""hidden"" name=""Language"" value=""" & Language & """>" & vbcrlf
		.write "<input type=""hidden"" name=""fluke_id"" value=""" & fluke_id & """>" & vbcrlf
		.write "<input type=""hidden"" name=""bus_sys"" value=""" & business_system & """>" & vbcrlf
		.write "<input type=""hidden"" name=""PostFlag"">" & vbcrlf
	end with
	
	' now, if we don't have a Fluke_ID we quit
	if fluke_id = "" then
		Response.write "</form>"
    response.write "<BR>" & Translate("There is not a current value for Customer Number.",Login_Language,conn) 
		Response.write "<BR>" & Translate("Choose",Login_Language,conn) & " &quot;" & Translate("Change Company",Login_Language,conn) & "&quot;<BR>" & vbcrlf
		exit sub
	end if
	
	' now we'll do the real work
	
	' let's get the company name - just to make everyone feel good
	sql = "select top 1 company from userdata where fluke_id = '" & fluke_id & "'" & vbcrlf &_
		"order by len(company) desc"
	
	set dbRS = conn.Execute(sql)
	Response.write "<SPAN CLASS=SmallBold>" & Translate("Company",Login_Language,conn) & ": " & "<SPAN CLASS=SmallBoldRed>" & dbRS("Company") & "</SPAN>"
	set dbRS = Nothing
	if bShow_Main then Response.write " (" & fluke_id & ") " & vbcrlf
  response.write "</SPAN><BR>"
	
	' now we'll do the table of Data
	sql = "select id" & vbcrlf &_
		",firstname" & vbcrlf &_
		",lastname" & vbcrlf &_
		",email" & vbcrlf &_
		",Subgroups" & vbcrlf &_
		",DateDiff(d,getdate(),ExpirationDate) as eDays" & vbcrlf &_
		"FROM UserData" & vbcrlf &_
		"WHERE site_id = " & site_id & vbcrlf &_
		"and fluke_id = '" & fluke_id & "'" & vbcrlf &_
		"and business_system = '" & business_system & "'" & vbcrlf &_
		"ORDER BY lastname,firstname"
	
	set dbRS = conn.Execute(sql)
	
	if dbRS.EOF then
		Response.write "</form>"
    response.write "<BR><SPAN CLASS=SmallBoldRed>" & Translate("There are new user accounts for this customer number",Login_Language,conn) & "</SPAN><BR>"
		set dbRS = Nothing
		exit sub
	end if
	
	%>
<BR>
<%Call Table_Begin%>
<TABLE BORDER=0 WIDTH="100%" BORDERCOLOR="#666666" BGCOLOR="#FFCC00" CELLPADDING=0 CELLSPACING=0>
  <TR>
    <TD align="center">
		<TABLE CELLPADDING=2 CELLSPACING=0 BORDER=0 BGCOLOR="#666666" WIDTH="100%">
			<TR>
	<%
	With Response
		.write "                <TD class=""SmallBoldGold"" BGColor=""Black"">&nbsp;&nbsp;"
		.write Translate("First Name",Login_Language,conn) & "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"" BGColor=""Black"">"
		.write Translate("Last Name",Login_Language,conn) & "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"" WIDTH=""60%"" BGColor=""Black"">"
		.write Translate("Email",Login_Language,conn) & "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"" ALIGN=CENTER BGColor=""Black"">"
		.write Translate("Active",Login_Language,conn) & "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"" ALIGN=CENTER BGColor=""Black"">"
		.write Translate("Order Inquiry",Login_Language,conn) & "<BR>" & Translate("Search",Login_Language,conn) &  "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"" BGColor=""Black"">"
    .write Translate("Action",Login_Language,conn)
    .write "</TD>" & vbcrlf
		.write "              </TR>" & vbcrlf
	end with
	
	Rnow = Now
	sids = ""
	strStatus = "|"
	strOrder = "|"
	do until dbRS.EOF
		uid = dbRS("ID")
		sids = sids & uid & ", "
		edays = dbRS("eDays")
		if edays > 0 then
			os = "CHECKED"
			strStatus = strStatus & uid & "|"
		else
			os = ""
		end if
		
		if Instr(ucase(dbRS("SubGroups")),"ORD") > 0 then
			oe = "CHECKED"
			strOrder = strOrder & uid & "|"
		else
			oe = ""
		end if
		
    if toggle = "#EAEAEA" then toggle = "#DBDBDB" else toggle = "#EAEAEA"
    
		With Response
			.write "              <TR>" & vbcrlf
			.write "                <TD class=""Small"" BGCOLOR=""" & toggle & """>&nbsp;&nbsp;"
			.write dbRS("FirstName") & "</td>" & vbcrlf
			.write "                <TD class=""Small"" BGCOLOR=""" & toggle & """>"
			.write dbRS("LastName") & "</td>" & vbcrlf
			.write "                <TD class=""Small"" BGCOLOR=""" & toggle & """>"
			.write dbRS("Email") & "</td>" & vbcrlf
			.write "                <TD class=""Small"" BGCOLOR=""" & toggle & """ ALIGN=CENTER>"
			.write "<input type=""checkbox"" name=""ostatus"" value=""" & uid & """" & os & "> "
			'.write edays & "</td>" & vbcrlf
			.write "</td>" & vbcrlf
			.write "                <TD class=""Small"" BGCOLOR=""" & toggle & """ ALIGN=CENTER>"
			.write "<input type=""checkbox"" name=""oenable"" value=""" & uid & """" & oe & ">"
			.write "</td>" & vbcrlf
		end with
		
		if bShow_Main then
			with response
				.write "                <TD BGCOLOR=""#006400"" ALIGN=CENTER>"
        .write "<A HREF=""/SW-Administrator/"
				.write "Account_Edit.asp?Site_ID=" & site_id & "&ID=edit_account&Account_ID=" & uid
				.write """ CLASS=""NavLeftHighlight1"" TARGET=""_BLANK"">"
				.write "&nbsp;"
        .write Translate("Edit",Login_Language,conn)
        .write "&nbsp;"
				.write "</a>"
        .write "</td>" & vbcrlf
			end with
		else
			response.write "                <TD class=""Small"">&nbsp;</td>" & vbcrlf
		end if
		
		response.write "              </TR>" & vbcrlf
		
		dbRS.MoveNext
	loop
	sids = Left(sids,(len(sids)-2))
	set dbRS = Nothing
	%>
			<TR>
			  <TD colspan=6 align=center bgcolor="#666666">
			  	<input type="button" onclick="Update();" value=" <%=Translate("Update",Login_Language,conn)%> " CLASS="NavLeftHighlight1">
			  </td>
			</tr>
		</TABLE>
    </TD>
  </TR>
</TABLE>
<%Call Table_End%>
<input type="hidden" name="sids" value="<%=sids%>">
<input type="hidden" name="sStatus" value="<%=strStatus%>">
<input type="hidden" name="sOrder" value="<%=strOrder%>">
</FORM>
<P>
	<%
end sub
%>
  
