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
Dim ErrMessage, iShowMax, iTotal, dbRS, bNeedTable, bShowAll, Shading, strDiscountRate

ErrMessage = ""
iShowMax = 100
iTotal = 0
bNeedTable = True
bShowAll = True
Shading = "#EFEFEF"

' Locals (at least sort of...)
bad_models = ""
bad_parts = ""
bHaveGroups = False
bHaveModels = False
bHaveParts = False

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
' Customer Number and srcsystem come from the form
strCustomerNum = Request.Form("customer_number")
strsrcsystem = Request.Form("srcsystem")

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

' now we need to get their pricelist...
if strCustomerNum <> 0 then
    set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Avail_ListRate_Get"
    ' create all the parameters we want
    cmd.Parameters.Append cmd.CreateParameter("@site",adInteger,adParamInput,,Site_id)
    cmd.Parameters.Append cmd.CreateParameter("@CustomerNum",adInteger,adParamInput,,clng(strCustomerNum))
    set dbRS = cmd.Execute
    set cmd = nothing
    
    if not dbRS.EOF then
        strPriceListName = dbRS("pricelistname")
        strDiscountRate = dbRS("DiscountRate")
        dbRS.Close
        set dbRS = nothing
    else
        set dbRS = nothing
        'response.write site_id & " - " & strCustomerNum & "<BR>"
        response.write "<SPAN CLASS=SmallBoldRed>We apologize, your company's discount rate has not been entered into our system.</SPAN>"
        response.end
    end if
else
    response.write "<SPAN CLASS=SmallBoldRed>No customer number in system</SPAN>"
    response.end
end if

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

strPartnum = Trim(Request.Form("part"))
strModel = Trim(Request.Form("model"))
strGroup = Trim(Request.Form("group"))

' how are we getting results??

if len(strGroup) > 0 then
	bHaveGroups = True
	if Instr(strGroup,",") > 0 then
		aGroups = split(strGroup,",")
	else
		ReDim aGroups(0)
		aGroups(0) = strGroup
	end if
	
	for i = 0 to ubound(aGroups)
		aGroups(i) = trim(aGroups(i))
	next
end if
	
if len(strModel) > 0 then
	bHaveModels = True
	if Instr(strModel,",") > 0 then
		aModels = split(strModel,",")
	else
		ReDim aModels(0)
		aModels(0) = strModel
	end if
	
	for i = 0 to ubound(aModels)
		aModels(i) = trim(aModels(i))
	next
end if

if len(strPartNum) > 0 then
	bHaveParts = True
	if Instr(strPartnum,",") > 0 then
		aParts = split(strPartnum,",")
	else
		ReDim aParts(0)
		aParts(0) = strPartnum
	end if
	
	for i = 0 to ubound(aParts)
		aParts(i) = trim(aParts(i))
	next
end if

' --------------------------------------------------------------------------------------
' Start building the page
' --------------------------------------------------------------------------------------
Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_width	  ' Percent
strTitle = Translate("Price & Availability Information",Login_Language,conn)

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

ActivitySQL = "INSERT INTO Activity (Account_ID,Site_ID,Session_ID,View_Time,CID,SCID,Language,Region)" & vbcrlf &_
          		"Values(" &_
          		Login_ID & "," &_
          		Site_ID & "," &_
          		Session("Session_ID") & "," &_
          		"getdate()," &_
          		"9008," &_
          		"1," &_
          		"'" & Login_Language & "'," &_
              user_region & ")"
conn.Execute (ActivitySQL)

response.write "<SPAN CLASS=Heading3>" & strTitle & "</SPAN><BR>"
response.write "<BR>"

' Substitute Navigation (because we'r using ...No-Navigation)
with Response

	.write "<form name=""foodle"" action=""SW-Avail_Form.asp"" method=""POST"">" & vbcrlf
	.write "<input type=""hidden"" name=""customer_Number"" value=""" & strCustomerNum & """>" & vbcrlf
	.write "<input type=""hidden"" name=""srcsystem"" value=""" & strSrcSystem & """>" & vbcrlf
	.write "</form>" & vbcrlf
  
  Call Nav_Border_Begin
  .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Home"" VALUE=""" & " " & Translate("Home",Login_Language,conn) & """ "
  .write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & BackURL & "'; return false;"" TITLE=""Return to Order Inquiry Select"">"
  
  if CInt(Request("bShowSwitch")) then
    .write "&nbsp;&nbsp;"
    .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""OICC"" VALUE=""" & " " & Translate("Change Company",Login_Language,conn) & """ "
    .write "LANGUAGE=""Javascript"" ONCLICK=""location.href='/sw-common/SW-Order_Inquiry_Change_Company.asp?caller=avail'; return false;"" TITLE=""Change Company"">"
  end if

  .write "&nbsp;&nbsp;"    
 	.write "<INPUT TYPE=""button"" VALUE="""
 	.write Translate("New Search",Login_Language,conn) & """ CLASS=""NavLeftHighlight1"""
 	.write " onclick=""javascript:document.foodle.submit();"">"
  
  .write "&nbsp;&nbsp;&nbsp;&nbsp;"
  .write "<INPUT TYPE=""BUTTON"" CLASS=NavLeftHighlight1 VALUE=""" & Translate("Print",Login_Language,conn) & """ NAME=""Print"" LANGUAGE=""JavaScript"" onClick=""PrintIt()"" TITLE=""Print Page"">"
  
  Call Nav_Border_End

end with

response.write "<SPAN CLASS=SmallBold>"

' this is the start of "real" content

Response.write "<P>" & Translate("Customer",Login_Language,conn) & ":&nbsp;&nbsp;" & strCustomerName
if bShowSwitch then
  response.write "&nbsp;&nbsp;[" & strCustomerNum & "]"
  response.write "&nbsp;&nbsp;[" & strSrcSystem & "]"
end if
response.write "</SPAN>" & vbcrlf
Response.write "<P><P>" & vbcrlf

Connect_eStoreDatabase

' if we have Groups then we get all the parts from that group
if bHaveGroups then
	bad_Groups = Get_Results(1,aGroups,eConn)
end if

if bHaveModels then
	bad_models = Get_Results(2,aModels,eConn)
end if

if bHaveParts then
	bad_parts = Get_Results(3,aParts,eConn)
end if

if not bNeedTable then
	if iTotal > iShowMax then
		response.write "<TR><TD colspan=40 align=""left"" class=""MediumRed"" BGCOLOR=WHITE>"
		strMax = "Result count exceeded the maximum allowed, you may want to limit your search"
		response.write "<BR>" & Translate(strMax,Login_Language,conn) & "<P></td></tr>" & vbcrlf
	end if
	
    if 1 = 2 then
	if bShowAll then
		with response
			.write "<TR><TD colspan=40 align=""right"">"
			.write "<input type=""button"" value=""" & Translate("Get Net Prices",Login_Language,conn)
			.write """ onclick=""get_net();"" CLASS=""NavLeftHighlight1"""
			.write "</td></tr>" & vbcrlf
		end with
	end if
    end if
	Response.write "</table><!-- end of results table -->" & vbcrlf
  Call Table_End
	Response.write "</form>" & vbcrlf
end if

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
  response.write "<SPAN class=""Medium""><ul>"
  response.write ErrMessage
  response.write "</ul></span>" & vbcrlf
end if

response.write "<UL><!-- line 298 -->" & vbcrlf

' generic footers
if bShowAll then
	strOH = "On Hand Quantity does not reflect orders received today"
	response.write "<LI>" & Translate(strOH,Login_Language,conn) & "</LI>" & vbcrlf
end if

strATPe = "ATP Days is the number of calendar days from when we receive your order until we ship it."
response.write "<LI>" & Translate(strATPe,Login_Language,conn) & "</LI>" & vbcrlf
response.write "</UL>" & vbcrlf
  
strperror = Translate("Please check the Net Price boxes for which you want the price",Login_Language,conn)
%>
<script language="Javascript">
	function get_net() {
		var vHref = 'SW-Avail_Price.asp?partno=';
		var wName = 'FlukeNetPrice';
		var sOpts = 'status=no,scrollbars=1,resizable=1,toolbar=0,width=600,height=';
		var pwid = document.dist_info.nprice;
		var apl = new Array();
		var j = 0;
		
        if (pwid.length) {
    		for (var i=0;i<pwid.length;i++) {
    			if (pwid[i].checked) {
    				apl[j++] = pwid[i].value;
    			}
    		}
        }
        else {
            if (pwid.checked) {
                apl[j++] = pwid.value;
            }
        }
		
		if (j) {
			vHref += apl.join("|");
            vHref += '&customer_number=' + document.foodle.customer_Number.value;
            vHref += '&srcsystem=' + document.foodle.srcsystem.value;
			
			//j = Math.ceil((vHref.length - 26) / 7);
			j = (j * 20) + 380;
			sOpts += j;
				
			newWind = window.open(vHref,wName,sOpts);
		
			if (newWind.opener == null) {
			   newWind.opener = window;
			}
			newWind.focus();
		}
		else {
			alert("<%=strperror%>");
		}
	}
</script>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

' ------------- end of main --------------------------------------------------------------------


sub Write_Rows
	
	do until dbRS.EOF
		
		if Shading = "White" then 
			Shading = "#EFEFEF"
		else
			Shading = "White"
		end if
      	
		select case UCase(dbRS("weight_code"))
			case "KILOGRAM"
				strWC = " kg"
			case "GRAM"
				strWC = " gr"
			case "POUND"
				strWC = " lb"
			case else
				strWC = ""
		end select
		
		'response.write "LEN " & len(Cstr(cdbl(dbRS("weight"))))
		if isblank(dbRS("weight")) then
			strWC = "-"
		else
			strWC = FormatNumber(dbRS("weight"),3,-1) & strWC
		end if
		
		with response
			.write "    <tr>" & vbcrlf
			.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & ">"
			.write dbRS("Model") & "</td>" & vbcrlf
			.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & " ALIGN=""Right"">"
			.write dbRS("pfid") & "</td>" & vbcrlf
			.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & " ALIGN=""center"">"
			.write CkBlank(dbRS("unit_of_measure")) & "</td>" & vbcrlf
			.write "        <td class=""Small"" BGCOLOR=" & Shading & ">"
			.write CkBlank(dbRS("short_description")) & "</td>" & vbcrlf
			.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & " ALIGN=""Right"">"
			.write CkBlank(dbRS("upc_code")) & "</td>" & vbcrlf
			.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & " ALIGN=""Right"">"
			.write strWC & "</td>" & vbcrlf
			.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & " ALIGN=""Center"">"
			.write CkBlank(dbRS("ce_status")) & "</td>" & vbcrlf
			.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & ">"
			.write CkBlank(dbRS("origin")) & "</td>" & vbcrlf
			.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & " ALIGN=""right"">"
			.write CkBlank(dbRS("ATP")) & "</td>" & vbcrlf
			.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & " ALIGN=""right"">"
			.write CkBlank(FormatNumber((dbRS("list_price") / 100))) & "</td>" & vbcrlf
		end with
			
		if bShowAll then
			with response
				.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & " ALIGN=""Right"">"
				.write OnHand(dbRS("onhandqty")) & "</td>" & vbcrlf
				.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & " ALIGN=""Right"">"
			    .write Format_Discount(dbRS("list_price")) & "</td>" & vbcrlf
				'.write "<input type=""checkbox"" name=""nprice"" value=""" & dbRS("pfid")
				'.write """></td>" & vbcrlf
			end with
		end if
		response.write "    </tr>" & vbcrlf
		dbRS.MoveNext
		
		iTotal = iTotal + 1
		if iTotal > iShowMax then
			exit sub
		end if
	Loop
end sub

function CkBlank(myString)
  if isblank(myString) then
    CkBlank = "&nbsp;"
  else
    CkBlank = myString
  end if
end function

function OnHand(var)
    if isblank(var) then
        OnHand = "&nbsp;"
    elseif isNumeric(var) then
        if var < 0 then
            OnHand = "-"
        else
            OnHand = var
        end if
    else
        OnHand = var
    end if
end function

function Get_Results(itype,aMine,lConn)
	dim cmd,lcmd,ecmd,fvar
	
	Get_Results = ""
	
	if iTotal > iShowMax then
		exit function
	end if
	
	select case itype
		case 1 ' groups
			ecmd = "Avail_GetGroup"
		case 2 ' models
			ecmd = "Avail_GetModel"
			lcmd = "Avail_LikeModel"
		case 3 ' parts
			ecmd = "Avail_GetPart"
			lcmd = "Avail_LikePart"
	end select
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = lConn
	cmd.CommandType = adCmdStoredProc
	cmd.Parameters.Append cmd.CreateParameter("@input", adVarchar,adParamInput ,128)
	cmd.Parameters.Append cmd.CreateParameter("@Listname", adVarchar,adParamInput ,32,strPriceListName)
    
	for each fvar in aMine
	
		cmd.Parameters("@input").Value = fvar
		
		if Instr(fvar,"*") > 0 then
			fvar = Replace(fvar,"*","%")
			cmd.CommandText = lcmd
			'response.write "calling " & lcmd & " on " & fvar
		else
			cmd.CommandText = ecmd
			'response.write "calling " & ecmd & " on " & fvar
		end if
		cmd.Parameters("@input").Value = fvar
		set dbRS = cmd.Execute
		
		if dbRS.EOF then
			Get_Results = Get_Results & fvar & ", "
		else
			if bNeedTable then bNeedTable = Start_Table()
			
			Write_Rows()
		end if
		
		set dbRS = nothing
	next
	set cmd = Nothing
end function

function Start_Table
	
	' build some static strings
	tmodel = Translate("Model",Login_Language,conn)
	tprod = Replace(Translate("Fluke P/N",Login_Language,conn)," ","<BR>")
	tunit = Translate("Unit",Login_Language,conn)
	tdesc = Translate("Description",Login_Language,conn)
	tupc = Translate("UPC Code",Login_Language,conn)
	twght  = Translate("Weight",Login_Language,conn)
	tce   = Translate("CE",Login_Language,conn)
	torigin    = Translate("Origin",Login_Language,conn)
	tATP  =  " * " & Replace(Translate("ATP Days",Login_Language,conn)," ","<BR>")
	tonhand = Translate("On Hand<BR>Quantity",Login_Language,conn)
	tprice = Replace(Translate("List Price",Login_Language,conn)," ","<BR>")
	tnprice = Replace(Translate("Net Price",Login_Language,conn)," ","<BR>")
	
	with response
		.write vbcrlf & "<FORM name=""dist_info"" method=""POST"">" & vbcrlf
    Call Table_Begin
		.write "<TABLE width=""100%"" cellpadding=""3"" cellspacing=0 BORDER=0>" & vbcrlf
		.write "    <tr bgcolor=""" & Contrast & """>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"">"
		.write tmodel & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Center"">"
		.write tprod & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Center"">"
		.write tunit & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"">"
		.write tdesc & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Center"">"
		.write tupc & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Center"">"
		.write twght & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Center"">"
		.write tce & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Center"">"
		.write torigin & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Right"">"
		.write tATP & "</td>" & vbcrlf
		.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Right"">"
		.write tprice & "</td>" & vbcrlf
	end with
	
	if bShowAll then
		with response
			.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Right"">"
			.write tonhand & "</td>" & vbcrlf
			.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Right"">"
			.write tnprice & "</td>" & vbcrlf
		end with
	end if
	
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

'--------------------------------------------------------------------------------------
' Functions and Subroutines
'--------------------------------------------------------------------------------------

sub Table_Begin()
    response.write "<TABLE BORDER=""0"" WIDTH=""100%"" CELLPADDING=""0"" CELLSPACING=""0"" CLASS=TableBorder>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLr
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "        <TD VALIGN=""top"" WIDTH=""100%"">" & vbCrLf
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
    response.write "        <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
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


