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
' Author:     D. Whitlock
' Date:       2/1/2000
'             Sandbox
' --------------------------------------------------------------------------------------

response.buffer = true

' Globals (at leat those defined here...)
Dim ErrMessage, bNeedTable, Shading, strDiscountRate, strPriceListName

ErrMessage = ""
Shading = "#EFEFEF"
bNeedTable = True

' Locals (at least sort of...)

Call Connect_SiteWide

if Session("Logon_user") <> "" then
	%>
	<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
	<%
else
  response.redirect "/register/default.asp"
	site_id = 3
end if
%>
<!-- #include virtual="/SW-Common/SW-Site_Information.asp"-->
<%
strCustomerNum = Request.QueryString("customer_Number")
strSrcSystem = Request.QueryString("srcsystem")

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
	Login_ID = dbRS("ID")
	
	user_TypeCode =  cint(dbRS("Type_Code"))
	subgrp = Ucase(dbRS("SubGroups"))
	user_region = cint(dbRS("Region"))
    
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

' get their "official" company name
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "Order_Company"
' create all the parameters we want
cmd.Parameters.Append cmd.CreateParameter("@source",adVarChar,adParamInput,8,strSrcSystem)
cmd.Parameters.Append cmd.CreateParameter("@iCust",adInteger,adParamInput,,strCustomerNum)
set dbRS = cmd.Execute
set cmd = nothing

if dbRS.EOF then
    strCustomerName = "Not found"
else
    strCustomerName = dbRS("Customer_Name")
end if
set dbRS = Nothing

' now we need to get their pricelist...
if strCustomerNum <> 0 then
    set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Avail_ListRate_Get"
    ' create all the parameters we want
    cmd.Parameters.Append cmd.CreateParameter("@site",adInteger,adParamInput,,Site_id)
    cmd.Parameters.Append cmd.CreateParameter("@CustomerNum",adInteger,adParamInput,,strCustomerNum)
    set dbRS = cmd.Execute
    set cmd = nothing
    
    if not dbRS.EOF then
        strPriceListName = dbRS("pricelistname")
        strDiscountRate = dbRS("DiscountRate")
        dbRS.Close
        set dbRS = nothing
    else
        set dbRS = nothing
        response.write "Bogus customer - no number in system"
        response.end
    end if
else
    response.write "Bogus customer - no number in system"
    response.end
end if

strPartnum = Trim(Request.QueryString("partno"))

if Instr(strPartnum,"|") > 0 then
	aParts = split(strPartnum,"|")
else
	ReDim aParts(0)
	aParts(0) = strPartnum
end if

for i = 0 to ubound(aParts)
	aParts(i) = trim(aParts(i))
next

' --------------------------------------------------------------------------------------
' Start building the page
' --------------------------------------------------------------------------------------
Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_width	  ' Percent
strTitle = Translate("Net Price Information",Login_Language,conn)

Screen_Title    = Site_Description & " - " & strTitle
Bar_Title       = Site_Description & "<BR><SPAN CLASS=MediumBoldGold>" & strTitle & "</SPAN>"
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95

BackURL = Session("BackURL")
%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->

<%
' Log the activity

ActivitySQL = "INSERT INTO Activity (Account_ID,Site_ID,Session_ID,View_Time,CID,SCID,Language,Region,Country)" & vbcrlf &_
          		"Values(" &_
          		Login_ID & "," &_
          		Site_ID & "," &_
          		Session("Session_ID") & "," &_
          		"getdate()," &_
          		"9008," &_
          		"2," &_
          		"'" & Login_Language & "'," &_
              user_region & "," &_
              "'" & Session("Login_Country") & "')"
conn.Execute (ActivitySQL)

response.write "<SPAN CLASS=Heading3>" & strTitle & "</SPAN><BR>"
response.write "<P>"

Response.write "<SPAN CLASS=SmallBold>" & Translate("Customer",Login_Language,conn) & ":&nbsp;&nbsp;" & strCustomerName
if bShowSwitch then
  response.write "&nbsp;&nbsp;[" & strCustomerNum & "]"
  response.write "&nbsp;&nbsp;[" & strSrcSystem & "]"
end if
response.write "</SPAN>" & vbcrlf

Response.write "<P><P><CENTER>" & vbcrlf

mydebug = False
if Request.Form("debug") = "on" then
	mydebug = True
	with response
		.write "<Table><TR><TD>"
		.write "PO Num = " & strPOnum & "<BR>" & vbcrlf
	end with
	
	if len(strOrderNum) > 0 then
		response.write "Order Nums = " & join(OrderNums," - ") & "<BR>" & vbcrlf
	else
		response.write "Order Nums = <BR>" & vbcrlf
	end if
	
	with response
		.write "Customer Num = " & strCustomerNum & "<BR>" & vbcrlf
		.write "Source = " & strSrcSystem & "<BR>" & vbcrlf
		.write "</td><TD><EM>This information is shown in test mode only</em>" & vbcrlf
		.write "</td></tr></table>" & vbcrlf
	end with
end if

if strPartNum = "" then
	ErrMessage = ErrMessage & "<LI>" & Translate("No input == no results",Login_Language,conn) &_
		": <SPAN class=""MediumRed"">" & bad_parts & "</SPAN></LI>"
else
	Connect_eStoreDatabase
	
	bad_parts = Get_Results(aParts,eConn)
	Disconnect_eStoreDatabase
	
	if not bNeedTable then Response.write "</table><!-- end of results table -->" & vbcrlf
end if

response.write "<UL>" & vbcrlf

if not isblank(bad_models) then
	bad_models = left(bad_models,len(bad_models)-2)
	ErrMessage = "<LI>" & Translate("No results for Model",Login_Language,conn) &_
		": <SPAN class=""MediumRed"">" & bad_models & "</SPAN></LI>"
end if

if not isblank(bad_parts) then
	bad_parts = left(bad_parts,len(bad_parts)-2)
	ErrMessage = ErrMessage & "<LI>" & Translate("No results for Part",Login_Language,conn) &_
		": <SPAN class=""MediumRed"">" & bad_parts & "</SPAN></LI>"
end if

if not isblank(ErrMessage) then
  response.write "<SPAN class=""Medium"">"
  response.write ErrMessage
  response.write "</span>" & vbcrlf
end if

with response
	.write "</UL>" & vbcrlf
	.write "<P>" & Translate("This is a new window, you can close it",Login_Language,conn) & vbcrlf
	.write "<form name=np>" & vbcrlf
	.write "<input type=""button"" onclick=""CloseMe();"" value="""
	.write Translate("Close Window",Login_Language,conn) & """></form></center>" & vbcrlf
end with
  
%>
<SCRIPT language="JavaScript1.2">
	function CloseMe() {
		self.close();
	}
</script>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

' ------------- end of main --------------------------------------------------------------------

function CkBlank(myString)
  if isblank(myString) then
    CkBlank = "&nbsp;"
  else
    CkBlank = myString
  end if
end function

function Get_Results(aMine,lConn)
	dim cmd,lcmd,ecmd,fvar
	
	Get_Results = ""
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = lConn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "Avail_Net_Get"
	cmd.Parameters.Append cmd.CreateParameter("@input", adVarchar,adParamInput ,128)
	cmd.Parameters.Append cmd.CreateParameter("@Listname", adVarchar,adParamInput ,32,strPriceListName)
	
	for each fvar in aMine
		cmd.Parameters("@input").Value = fvar
		set dbRS = cmd.Execute
		
		if dbRS.EOF then
			Get_Results = Get_Results & fvar & ", "
		else
			if bNeedTable then bNeedTable = Start_Table()
			
			do until dbRS.EOF
				if Shading = "White" then 
					Shading = "#EFEFEF"
				else
					Shading = "White"
				end if
				
				with response
					.write "    <tr>" & vbcrlf
					.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & ">"
					.write dbRS("Model") & "</td>" & vbcrlf
					.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & ">"
					.write dbRS("pfid") & "</td>" & vbcrlf
					.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & " ALIGN=""right"">"
					.write Format_Discount(dbRS("list_price")) & "</td>" & vbcrlf
				end with
				response.write "    </tr>" & vbcrlf
				dbRS.MoveNext
			Loop
		end if
		
		set dbRS = nothing
	next
	set cmd = Nothing
end function

function Start_Table
	' build some static strings
	tmodel = Translate("Model",Login_Language,conn)
	tprod = Translate("Fluke P/N",Login_Language,conn)
	tnprice = Translate("Net Price",Login_Language,conn)
	
	with response
		.write vbcrlf & "<TABLE cellpadding=""3"" cellspacing=0 BORDER=0>" & vbcrlf
		.write "    <tr bgcolor=""" & Contrast & """>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"">"
		.write tmodel & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"">"
		.write tprod & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Right"">"
		.write tnprice & "</td>" & vbcrlf
	end with
	
	response.write "    </tr>" & vbcrlf
	Start_Table = False
end function

Function Format_Discount(num)
    Format_Discount = "??"
    
    if isblank(num) then
        exit function
    elseif not isnumeric(num) then
        exit function
    else
        Format_Discount = FormatNumber(((num - (num * (Cdbl(strDiscountRate)/100)) )/ 100))
    end if
end function

%>
