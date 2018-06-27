<%@ Language="VBScript" CODEPAGE="65001" %>

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

Dim ErrMessage
ErrMessage = ""

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
  user_region   = cint(dbRS("Region"))
	
	strLang = dbRS("Description")
	if strLang = "Chinese (Simplified)" then
		strLang = "Simplified Chinese"
	elseif strLang = "Chinese (Traditional)" then
		strLang = "Traditional Chinese"
	end if
	
end if
set dbRS = Nothing

Dim strPOnum, strOrderNum, strCustomerNum, strSrcSystem,dbRS1,g_sDisplayType
Dim cmd
Dim prm

strPOnum = Trim(Request.Form("ponum"))
strOrderNum = Trim(Request.Form("ordernum"))
strCustomerNum = Trim(Request.Form("customernum"))
strSrcSystem = Trim(Request.Form("srcsystem"))
inq_type = Trim(Request.Form("inq_type"))
g_sDisplayType = Trim(Request.Form("UserType")) ' "Rep" and "Dist" are the 2 values


' build the top level display structure

if len(strOrderNum) < 1 then
	' Get OrderNum based on POnum
elseif Instr(strOrderNum,",") > 0 then
	OrderNums = split(strOrderNum,",")
else
	ReDim OrderNums(0)
	OrderNums(0) = strOrderNum
end if

' --------------------------------------------------------------------------------------
' Start building the page
' --------------------------------------------------------------------------------------
Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_width	  ' Percent

Screen_Title    = Site_Description & " - " & Translate("Order Status Information",Login_Language,conn)
Bar_Title       = Site_Description & "<BR><SPAN CLASS=MediumBoldGold>" & _
                  Translate("Order Status Information",Login_Language,conn) & "</SPAN>"
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95

BackURL = Session("BackURL")
%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->

<%
' Log the activity

ActivitySQL = "INSERT INTO Activity (Account_ID,Site_ID,Session_ID,View_Time,CID,SCID,Language,Method,Calendar_ID,Region,Country)" & vbcrlf &_
          		"Values(" &_
          		Login_ID & "," &_
          		Site_ID & "," &_
          		Session("Session_ID") & "," &_
          		"getdate()," &_
          		"9007," &_
          		"1," &_
          		"'" & Login_Language & "'," & _
              "0," & _
              "102," &_
              user_region & "," &_
              "'" & Session("Login_Country") & "')"
conn.Execute (ActivitySQL)

response.write "<SPAN CLASS=Heading3>" & Translate("Order Status Information",Login_Language,conn) & "</SPAN><BR>"
response.write "<BR>"

' Substitute Navigation (because we'r using ...No-Navigation)
with Response
	.write "<form name=""foodle"" action=""SW-Order_Inquiry_Form.asp"" method=""POST"">" & vbcrlf
	.write "<input type=""hidden"" name=""customer_Number"" value=""" & strCustomerNum & """>" & vbcrlf
	.write "<input type=""hidden"" name=""srcsystem"" value=""" & strSrcSystem & """>" & vbcrlf
	.write "</form>" & vbcrlf
	.write "<SPAN CLASS=SmallBold>"
	.write "<A HREF=""" & Request("BackURL") & """>"
	.write Translate("Home",Login_Language,conn) & "</a> | " & vbcrlf
	.write "<A HREF=""javascript:document.foodle.submit();"">" & Translate("New Search",Login_Language,conn) & "</A>"
end with

if CInt(Request("bShowSwitch")) then
	Response.write " | <A HREF=""/sw-common/SW-Order_Inquiry_Change_Company.asp"">" & Translate("Change Company",Login_Language,conn) & "</a>" & vbcrlf
end if

'get the date this data is as of
	
set cmd1 = Server.CreateObject("ADODB.Command")
cmd1.ActiveConnection = conn
cmd1.CommandType = adCmdStoredProc
cmd1.CommandText = "Order_GetUpload"

cmd1.Parameters.Append cmd1.CreateParameter("@source",adVarChar,adParamInput,8,strSrcSystem)
cmd1.Parameters.Append cmd1.CreateParameter("@strLang",adVarChar,adParamInput,20,strLang)
set dbRS = cmd1.Execute
set cmd1 = nothing

if not dbRS.EOF then
	with Response
		.write "<P><SPAN CLASS=Small>" & Translate("Order data was refreshed",Login_Language,conn) & " "
		.write Replace(Replace(dbRS("Fupload_date"),"AM"," AM"),"PM"," PM") & " PST</SPAN><P>" & vbcrlf
	end with
end if
set dbRS = nothing

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

prev_cust = "foobar"
bad_ponums = ""
bad_ordernums = ""
bShowWayFoot = False
bShowDirect = False
bShowWayUL = False

' if we don't have an ordernum then we have to get one from a po search
if len(strPOnum) > 0 then
	'if mydebug then Response.write "strPOnum has length " & len(strPOnum) & "<BR>" & vbcrlf
	if Instr(strPOnum,",") > 0 then
		PONums = split(strPOnum,",")
	else
		ReDim PONums(0)
		PONums(0) = strPOnum
	end if
	
	for i = 0 to ubound(PONums)
		PONums(i) = trim(PONums(i))
	next
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.Parameters.Append cmd.CreateParameter("@source", adVarchar,adParamInput ,8, strSrcSystem)
        cmd.Parameters.Append cmd.CreateParameter("@iCust", adVarChar,adParamInput, 32,strCustomerNum)
	cmd.Parameters.Append cmd.CreateParameter("@ponum", adVarchar,adParamInput, 64)
	
	for each strPOnum in PONums
		'if mydebug then Response.write "looking for " & strPOnum & "<BR>" & vbcrlf
		
		' this sets us up to increment the size of OrderNums, if it exists or not
		if IsArray(OrderNums) then
			i = ubound(OrderNums) + 1
		else
			'if mydebug then Response.write "OrderNums array does not yet exist<BR>" & vbcrlf
			ReDim OrderNums(0)
			i = 0
		end if
		
		cmd.Parameters("@ponum").Value = strPOnum
		
        if g_sDisplayType = "Rep" then
    		if Instr(strPOnum,"%")>0 AND inq_type = "lk" then
    			cmd.CommandText = "Order_LikeRepOrder_PO"
    		else
    			cmd.CommandText = "Order_GetRepOrder_PO"
    		end if
        else
    		if Instr(strPOnum,"%")>0 AND inq_type = "lk" then
    			cmd.CommandText = "Order_LikeOrder_PO"
    		else
    			cmd.CommandText = "Order_GetOrder_PO"
    		end if
		end if
		
		set dbRS = cmd.Execute
		
		if dbRS.EOF then
			bad_ponums = bad_ponums & strPOnum & ", "
      'ErrMessage = "<LI>" & Translate("There are no orders for PO Number",Login_Language,conn) & ": <SPAN CLASS=MediumBold>" & strPOnum & "</SPAN></LI>"
		else
			do until dbRS.EOF
				ReDim Preserve OrderNums(i)
				OrderNums(i) = dbRS("Order_number")
				i = i + 1
				dbRS.MoveNext
			Loop
		end if
		
		set dbRS = nothing
	next
	set cmd = nothing
end if

'if mydebug then
'	Response.write "<BR>" & join(OrderNums," - ") & "<BR>" & vbcrlf
'end if

if not isblank(OrderNums(0)) then
  Call Get_Orders
end if

if not isblank(bad_ponums) then
	bad_ponums = left(bad_ponums,len(bad_ponums)-2)
	ErrMessage = "<LI>" & Translate("There are no orders for PO Number",Login_Language,conn) &_
		": <SPAN class=""MediumRed"">" & bad_ponums & "</SPAN></LI>"
end if

if not isblank(bad_ordernums) then
	bad_ordernums = left(bad_ordernums,len(bad_ordernums)-2)
	ErrMessage = ErrMessage & "<LI>" & _
		Translate("There are no orders for Fluke Order Number",Login_Language,conn) &_
		": <SPAN class=""MediumRed"">" & bad_ordernums & "</SPAN></LI>"
end if

if not isblank(ErrMessage) then
  response.write "<SPAN class=""Medium""><ul>"
  response.write ErrMessage
  response.write "</ul></span>" & vbcrlf
end if
  
%>
<script language="Javascript">
	function Track(num,carrier) {
		
		if (carrier == 'UPS') {
			var vHref = 'http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=5';
  			vHref += '&TypeOfInquiryNumber=T&HTMLVersion=4.0&InquiryNumber1=' + num;
		}
		else if (carrier == 'FDX') {
			var vHref = 'http://www.fedex.com/cgi-bin/tracking?action=track&language=english&';
        vHref += 'cntry_code=us&initial=x&tracknumbers=' + num;
		}
		else if (carrier == 'TNT') {
			var vHref = 'http://www.tntew.com/new_tracker/SaCGI.cgi/tracker.exe?';
        vHref += 'FNC=gotoresults__Adummy_html___conok='+num+'___ttype=R___lang=EN___page=0';
        vHref += '___laf=default';
		}
		else if (carrier == 'PUR') {
			var vHref = 'http://shipnow.purolator.com/shiponline/track/PurolatorTrackE.asp?';
        vHref += 'PINNO=' + num;
		}
		
		var wName = 'Track';
	
		newWind = window.open(vHref,wName);
	
		if (newWind.opener == null) {
		   newWind.opener = window;
		}
		// self.blur();
		newWind.focus();
	}
</script>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

' ------------- end of main --------------------------------------------------------------------

sub Get_Orders

  response.write "<DIV ALIGN=CENTER>" & vbCrLf   
  
  Call Table_Begin

	response.write "<TABLE width=""100%"" cellpadding=""3"" cellspacing=0 BORDER=0 BGCOLOR=""White"">" & vbcrlf
	
	' build some static strings
	tline   = Replace(Translate("Fluke Line",Login_Language,conn)," ","<BR>") & " #"
	tprod   = Replace(Translate("Fluke P/N",Login_Language,conn)," ","<BR>")
	tmod    = Replace(Translate("Model Number",Login_Language,conn)," ","<BR>")
	tqty    = Replace(Translate("Qty Ordered",Login_Language,conn)," ","<BR>")
	tcanqty = Replace(Translate("Qty Cancelled",Login_Language,conn)," ","<BR>")
	tschdt  = Replace(Replace(Translate("Scheduled Ship-Date",Login_Language,conn)," ","<BR>"),"-"," ")
	tmeth   = Replace(Translate("Ship Method",Login_Language,conn)," ","<BR>")
	tway    = Replace(Translate("Waybill Number",Login_Language,conn)," ","<BR>")
	tactdt  = Replace(Replace(Translate("Actual Ship-Date",Login_Language,conn)," ","<BR>"),"-"," ")
	tshpqty = Replace(Translate("Qty Shipped",Login_Language,conn)," ","<BR>")
	tpurch  = Replace(Translate("Purchased Services",Login_Language,conn)," ","<BR>")
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "Order_GetOrderStatus"
	
	cmd.Parameters.Append cmd.CreateParameter("@strLang", adVarchar,adParamInput ,50, strLang)
	cmd.Parameters.Append cmd.CreateParameter("@strOrder", adVarchar,adParamInput ,50)
	cmd.Parameters.Append cmd.CreateParameter("@source", adVarchar,adParamInput ,50, strSrcSystem)
	
	for each onum in OrderNums
		' build the command => the command brings back 1 order header with multiple lines
		onum = Trim(onum)
		prev_line = 0
		
		cmd.Parameters("@strOrder").Value = onum
		Set dbRS = cmd.Execute
		
		if not dbRS.EOF then
			cur_cust = dbRS("Customer_Name")
			if prev_cust <> cur_cust then
				Response.write "<TR><TD colspan=""20""><SPAN CLASS=SmallBold>" & Translate("Customer",Login_Language,conn) & ":&nbsp;&nbsp;" & cur_cust
		        if CInt(Request("bShowSwitch")) then
		        	response.write "&nbsp;&nbsp;[" & strCustomerNum & "]"
		        	response.write "&nbsp;&nbsp;[" & strSrcSystem & "]"
		        end if
		        response.write "</SPAN></td></tr>" & vbcrlf
				prev_cust = cur_cust
			end if
			
			with Response
				' big blank line
  				.write "<TR><TD class=""Heading1"" colspan=""20"">&nbsp;</td></tr>" & vbcrlf
				
	  			' Order Information
	  			.write "    <tr bgcolor=""" & Contrast & """>" & vbcrlf
	  			.write "       <td colspan=""2"" valign=""TOP"" NOWRAP class=""Small"">"
	        	.write "<SPAN CLASS=SmallBold>" & Translate("Order Detail",Login_Language,conn)
				.write ":</SPAN><BR>" & vbCrLf
	   			.write Translate("Order Number",Login_Language,conn) & ":<BR>" & vbcrlf
	        	.write Translate("PO Number",Login_Language,conn)    & ":<BR>" & vbcrlf
	  			.write Translate("Order Date",Login_Language,conn)   & ":<BR>" & vbcrlf
	  			.write "       </TD>" & vbCrLf	
	        	.write "       <td colspan=""1"" valign=""TOP"" NOWRAP class=""SmallBold"">"
	        	.write "       <BR>" 
	  			.write CkBlank(onum) & "<BR>" & vbcrlf
	  			.write CkBlank(dbRS("Customer_PO_Number")) & "<BR>" & vbcrlf
	  			.write CkBlank(dbRS("TOrder_Date")) & "<BR>" & vbcrlf
	  			.write "</td>" & vbcrlf
				
	     		' Billing Information
	  			.write "       <td colspan=""4"" class=""Small"" valign=""TOP"" NOWRAP><SPAN class=""SmallBold"">"
	  			.write Translate("Billing Information",Login_Language,conn) & ":</SPAN><BR>" & vbcrlf
			end with
			
			' get the billing address
			Set cmd1 = Server.CreateObject("ADODB.Command")
			Set cmd1.ActiveConnection = conn
			cmd1.CommandType = adCmdStoredProc
			cmd1.CommandText = "Order_GetAddress"
			
			cmd1.Parameters.Append cmd1.CreateParameter("@source", adVarchar,adParamInput ,50,_
					strSrcSystem)
			cmd1.Parameters.Append cmd1.CreateParameter("@CustomerID", advarchar,adParamInput,15,_
					dbRS("Bill_Cust_ID"))
			cmd1.Parameters.Append cmd1.CreateParameter("@AddressID", advarchar,adParamInput,10,_
					dbRS("Bill_To_Address_ID"))
			
			Set dbRS1 = cmd1.Execute
			
			if dbRS1.EOF then
				Response.write "Not found"
			else
				write_address(cur_cust)
				dbRS1.close
			end if
			set dbRS1 = nothing
			
			with response
				.write "</td>" & vbcrlf 
				.write "       <td colspan=""3"" class=""Small"" valign=""TOP"" NOWRAP><SPAN class=""SmallBold"">"
				.write Translate("Shipping Information",Login_Language,conn) & ":</SPAN><BR>" & vbcrlf
			end with
			
			' get the shipping address
			cmd1.Parameters("@CustomerID").Value = dbRS("Ship_Cust_ID")
			cmd1.Parameters("@AddressID").Value = dbRS("Ship_To_Address_ID")
			
			Set dbRS1 = cmd1.execute
			
			if dbRS1.EOF then
				Response.write "Not found"
			else
				write_address(cur_cust)
				dbRS1.close
			end if
			set dbRS1 = nothing
			set cmd1 = Nothing
			
			' finish the header row
			Response.write "</td>" & vbcrlf & "       <TD CLASS=Small>&nbsp;</TD>" & vbcrlf
      		Response.write "</TR>" & vbcrlf
			
			' now we're going to do the line information
			' line set headers
			with response
  				.write "    <tr>" & vbcrlf
  				.write "        <td colspan=6>&nbsp;</td>" & vbcrlf
  				.write "        <td valign=""Bottom"" align=""center"" colspan=""4"" "
				.write "bgcolor=""#CCCCCC"" nowrap class=""SmallBold"">" & vbcrlf
  				.write "        " & Translate("Shipment Information",Login_Language,conn)
				.write "</td>" & vbcrlf
  				.write "        <td nowrap valign=""Bottom"">&nbsp;</td>" & vbcrlf
  				.write "    </tr>" & vbcrlf
  				.write "    <tr>" & vbcrlf
  				.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Right"">"
				.write tline & "</td>" & vbcrlf
  				.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Right"">"
				.write tprod & "</td>" & vbcrlf
  				.write "        <td nowrap class=""SmallBold"" valign=""Bottom"">" & tmod & "</td>"
				.write vbcrlf
  				.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Right"">"
				.write tqty & "</td>" & vbcrlf
  				.write "        <td nowrap class=""SmallBold"" valign=""Bottom"" ALIGN=""Right"">"
				.write tcanqty & "</td>" & vbcrlf
  				.write "        <td nowrap class=""SmallBold"" valign=""Bottom"">"
				.write tschdt & "</td>" & vbcrlf
  				.write "        <td bgcolor=""#CCCCCC"" nowrap class=""SmallBold"" valign=""Bottom"">"
				.write tmeth & "</td>" & vbcrlf
  				.write "        <td bgcolor=""#CCCCCC"" nowrap class=""SmallBold"" valign=""Bottom"">"
				.write tway & "</td>" & vbcrlf
  				.write "        <td bgcolor=""#CCCCCC"" nowrap class=""SmallBold"" valign=""Bottom"">"
				.write tactdt & "</td>" & vbcrlf
  				.write "        <td bgcolor=""#CCCCCC"" nowrap class=""SmallBold"" valign=""Bottom"" "
				.write "ALIGN=""Right"">" & tshpqty & "</td>" & vbcrlf
  				.write "        <td nowrap class=""SmallBold"" valign=""Bottom"">"
				.write tpurch & "</td>" & vbcrlf
  				.write "    </tr>" & vbcrlf
			end with
			
      		Shading = "White"
			
			do until dbRS.EOF
				' start by inspecting waybill and ship method
				waybill = dbRS("Waybill_Number")
				shp_method = dbRS("Ship_Method")
				
				if not bShowWayFoot then 
					if Instr(waybill,"+") > 1 then 
						bShowWayFoot = True
					end if
				end if
				
				if len(waybill) > 0 then
					if Instr(lcase(waybill),"b/o") > 0 then
						waybill = "back order"
						shp_method = "&nbsp;"
					elseif (Instr(ucase(shp_method),"UPS") = 1) then
						bShowWayUL = True
						waybill = "<a href=""javascript:Track('" &_
  					  Trim(Replace(waybill,"+","")) & "','UPS');"">" & waybill & "</a>"
					elseif (Instr(ucase(shp_method),"TNT") <> 0) then
						bShowWayUL = True
						waybill = "<a href=""javascript:Track('" &_
  					  Trim(Replace(waybill,"+","")) & "','TNT');"">" & waybill & "</a>"
          ' this shows all MFG orders with shipped dates as links to TNT
					elseif (ucase(strSrcSystem) = "MFG" and len(dbRS("TShip_Date")&"") > 0) then
						bShowWayUL = True
						waybill = "<a href=""javascript:Track('" &_
  					  Trim(Replace(waybill,"+","")) & "','TNT');"">" & waybill & "</a>"
					elseif (Instr(ucase(shp_method),"FDX") = 1) then
						bShowWayUL = True
						waybill = "<a href=""javascript:Track('" &_
  					  Trim(Replace(waybill,"+","")) & "','FDX');"">" & waybill & "</a>"
					elseif (Instr(ucase(shp_method),"PUR") = 1) then
						bShowWayUL = True
						waybill = "<a href=""javascript:Track('" & Trim(waybill) & "','PUR');"">" &_
              waybill & "</a>"
					elseif (Instr(ucase(shp_method),"DIR.SH") = 1) then
						waybill = ""
						bShowDirect = True
					end if
				end if
				
				'this block switches the color, don't do that if the line number hasn't changed
				line_num = dbRS("Line_number")
				
				if line_num <> prev_line then
	        		if Shading = "White" then 
						Shading = "#EFEFEF" 
					else
						Shading = "White"
					end if
					prev_line = line_num
				end if
        		
				with response
  					.write "    <tr>" & vbcrlf
  					.write "        <td nowrap class=""Small"" ALIGN=RIGHT BGCOLOR=" & Shading & ">"
					.write CkBlank(line_num)
  					.write "</td>" & vbcrlf
  					.write "        <td nowrap class=""Small"" ALIGN=RIGHT BGCOLOR=" & Shading & ">"
					.write CkBlank(dbRS("Item_Number"))
  					.write "</td>" & vbcrlf
  					.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & ">"
					.write CkBlank(dbRS("Model"))
  					.write "</td>" & vbcrlf
  					.write "        <td nowrap class=""Small"" ALIGN=RIGHT BGCOLOR=" & Shading
					.write ">" & CkBlank(dbRS("Quantity_Ordered"))
  					.write "</td>" & vbcrlf
  					.write "        <td nowrap class=""Small"" ALIGN=RIGHT BGCOLOR=" & Shading
					.write ">" & CkBlank(dbRS("Quantity_Cancelled"))
  					.write "</td>" & vbcrlf
  					.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & ">"
					.write CkBlank(dbRS("TScheduled_Ship_Date"))
  					.write "</td>" & vbcrlf
  					.write "        <td bgcolor=""" & Shading & """ nowrap class=""Small"">"
  					.write CkBlank(shp_method) & "</td>" & vbcrlf
  					.write "        <td bgcolor=""" & Shading & """ nowrap class=""Small"">"
  					.write CkBlank(waybill) & "</td>" & vbcrlf
  					.write "        <td bgcolor=""" & Shading & """ nowrap class=""Small"">"
  					.write CkBlank(dbRS("TShip_Date")) & "</td>" & vbcrlf
  					.write "        <td bgcolor=""" & Shading & """ nowrap class=""Small"" ALIGN=RIGHT>"
  					.write CkBlank(dbRS("Shipped_Quantity")) & "</td>" & vbcrlf
					.write "        <td nowrap class=""Small"" BGCOLOR=" & Shading & ">"
					.write CkBlank(dbRS("Note1"))
					.write "</td>" & vbcrlf
					.write "    </tr>" & vbcrlf
				end with
				dbRS.MoveNext
			Loop
		else
			bad_ordernums = bad_ordernums & onum & ", "
    	end if
    
	Next
	set cmd = nothing
	
  	response.write "<TR><TD colspan=""20"" class=""Small""><BR>"
	
	if bShowWayFoot OR bShowWayUL OR bShowDirect then
		Response.write "<UL>" & vbcrlf
		
		if bShowWayUL then _
			response.write "<LI>" & Translate("An underlined <u>Waybill Number</U> indicates a link to Shipment Tracking information.",Login_Language,conn) & "</LI>"
			
		if bShowWayFoot then _
			response.write "<LI>" & Translate("The + sign in Waybill Number indicates there is more than one waybill for this Line Item.",Login_Language,conn) & "</LI>"
			
		if bShowDirect then _
			response.write "<LI>" & Translate("DIR.SHP indicates the product on this line was shipped from our supplier directly to you.  The shipment date is the day we received notification of the shipment.",Login_Language,conn) & "</LI>"
			
		Response.write "</UL>" & vbcrlf
	else
	    response.write "<P>&nbsp;" 
	end if
  response.write "<P>&nbsp;</td></tr>" & vbcrlf
	Response.write "</table>" & vbcrlf
  
  Call Table_End

  response.write "</DIV>" & vbCrLf
  
  response.write "<P>"
  
end sub

'--------------------------------------------------------------------------------------

sub write_address(cust)
	if dbRS1("Customer_Name") <> cust then
		Response.write dbRS1("Customer_Name") & "<BR>" & vbcrlf & vbtab
	end if
	
	if len(trim(dbRS1("Address_1"))) > 0 then
  	Response.write dbRS1("Address_1") & "<BR>" & vbcrlf & vbtab
	end if
	
	if len(trim(dbRS1("Address_2"))) > 0 then
		Response.write dbRS1("Address_2") & "<BR>" & vbcrlf & vbtab
	end if
	
	if len(trim(dbRS1("Address_3"))) > 0 then
		Response.write dbRS1("Address_3") & "<BR>" & vbcrlf & vbtab
	end if
	
	if len(trim(dbRS1("Address_4"))) > 0 then
		Response.write dbRS1("Address_4") & "<BR>" & vbcrlf & vbtab
	end if
	
	country = trim(dbRS1("Country"))
	
	if ucase(country) = "US" then
	
		with response
			.write dbRS1("City") & ", " & vbcrlf & vbtab
			.write dbRS1("State") & "&nbsp;&nbsp;" & vbcrlf & vbtab
			.write dbRS1("Postal_Code") & "<BR>" & vbcrlf & vbtab
			.write country & "<BR>" & vbcrlf & vbtab
		end with
	else
		if len(trim(dbRS1("City"))) > 0 then
			Response.write dbRS1("City") & "<BR>" & vbcrlf & vbtab
		end if
		
		if len(trim(dbRS1("State"))) > 0 then
			Response.write dbRS1("State") & "<BR>" & vbcrlf & vbtab
		end if
		
		if len(trim(dbRS1("Province"))) > 0 then
			Response.write dbRS1("Province") & "<BR>" & vbcrlf & vbtab
		end if
		
		if len(trim(dbRS1("Postal_Code"))) > 0 then
			Response.write dbRS1("Postal_Code") & "<BR>" & vbcrlf & vbtab
		end if
		
		if len(country) > 0 then
			Response.write country & "<BR>" & vbcrlf & vbtab
		end if
	end if
end sub

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
    response.write "    <TD VALIGN=""top"" WIDTH=""98%"">" & vbCrLf
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

function CkBlank(myString)
  if isblank(myString) then
    CkBlank = "&nbsp;"
  else
    CkBlank = myString
  end if
end function

'--------------------------------------------------------------------------------------
%>
