<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_FormData.asp"-->
<!--#include virtual="/connections/connection_estore.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->
<!--#include virtual="/SW-Common/SW-Security_Validate_Formdata_Sitewide.asp" -->

<%
response.buffer = true

' --------------------------------------------------------------------------------------
' Globals
' --------------------------------------------------------------------------------------

Dim g_strEmail		' User Email
Dim g_strPassword	' User Password

Dim g_iLimit		  ' Maximum number of records to display on a page
Dim g_iCurrPage		' Current page of the recordset we are on\want to show

Dim g_strLevel		' Gold versus silver

Dim g_strKeyword	' Keyword for search criteria
Dim g_strMainCalibrator	' Main calibrator for search criteria
Dim g_strAction		' Action to take - currently only for getting purchased procedures

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False

Dim Border_Toggle
Border_Toggle = 0

' --------------------------------------------------------------------------------------
' Initialize Globals
' --------------------------------------------------------------------------------------

Site_ID = 11	' Implicitly declared and used in /SW-Common/SW-Site_Information.asp

g_strKeyword = Request("Keyword")
g_strMainCalibrator = Request("Calibrator")
g_strAction = Request("strAction")

g_iLimit = 25
g_iCurrPage = Request("CurrPage")

if IsNumeric(g_iCurrPage) then
	if g_iCurrPage = 0 then
		g_iCurrPage = 1
	end if
else
	g_iCurrPage = 1
end if

g_strLevel = ucase(trim(request("lv")))

' --------------------------------------------------------------------------------------
' Execute program
' --------------------------------------------------------------------------------------

Call GetConnections
Call Main(g_strKeyword, g_strMainCalibrator, g_strLevel, g_iCurrPage, g_iLimit, site_ID, g_strEmail, g_strPassword, g_strAction)

%><!--#include virtual="/SW-Common/SW-Footer.asp"--><%  

Call Disconnect_Connections()

' --------------------------------------------------------------------------------------
' Everything is executed in Sub Main() context
' --------------------------------------------------------------------------------------

Sub Main(strKeyword, strMainCalibrator, g_strLevel, g_iCurrPage, g_iLimit, site_ID, g_strEmail, g_strPassword, g_strAction)
	Dim rsSearchResult_Procedures
	Dim iSearchResults_Recordcount
	Dim strEStoreShopper
	Dim iNumPages


  
	Call Validate_Security(g_strSitewideUser, g_strCore_ID, Site_ID, g_iSitewide_ID, g_bCoreExists, g_strEmail, g_strPassword)
	strEStoreShopper = GetEStoreShopper(g_iSitewide_ID, g_bCoreExists, g_strEmail, g_strPassword, g_strCore_ID)
	set rsSearchResults_Procedures = Get_SearchResult_Procedures(strKeyword, strMainCalibrator, iSearchResult_Recordcount, g_strAction, strEStoreShopper)

	call StartPage(strKeyword, strPrimaryCalibrator, iSearchResult_Recordcount)

	if not rsSearchResults_Procedures.EOF then
		rsSearchResults_Procedures.PageSize = g_iLimit
		rsSearchResults_Procedures.AbsolutePage = g_iCurrPage

		iNumPages = rsSearchResults_Procedures.PageCount
		if iNumPages > 25 then iNumPages = 25

		Call Write_Nav_Buttons(iNumPages, g_iCurrPage, strKeyword, strMainCalibrator, g_strLevel, g_strCore_ID)
		Call WriteTableHeader
		Call DisplaySearchResults(rsSearchResults_Procedures, aBoughtProcedures, g_iLimit, g_strCore_ID, g_iSitewide_ID)
    Call OutputFooter    
		Call Write_Nav_Buttons(iNumPages, g_iCurrPage, strKeyword, strMainCalibrator, g_strLevel, g_strCore_ID)
	else
		Output_SearchEmpty(g_strCore_ID)
	end if

'	Call OutputFooter
	Call OutputJavascript()
End Sub

' ################################## Begin Supporting Functions ############################################

' --------------------------------------------------------------------------------------
' Get Connections
' --------------------------------------------------------------------------------------

Function GetConnections()
	Call Connect_SiteWide
	Call Connect_eStoreDatabase
End Function

' --------------------------------------------------------------------------------------
' Disconnect Connections
' --------------------------------------------------------------------------------------

Function Disconnect_Connections()
	Call Disconnect_SiteWide
	Call Disconnect_eStoreDatabase()
End function

' --------------------------------------------------------------------------------------
' Get recordset of procedures based on search criteria
' --------------------------------------------------------------------------------------

Function Get_SearchResult_Procedures(strKeyword, strMainCalibrator, iSearchResult_Recordcount, strAction, strEStoreShopper)
	Dim cmd,prm
	Dim rsProcedures

	strKeyword = Trim(Replace(Replace(UCase(strKeyword),"FLUKE-","FLUKE "), "*", "%"))
	strCalibrator = trim(Replace(Replace(strMainCalibrator, "'", ""), "*", "%"))

'	if strKeyword = "" and strCalibrator = "" then
'		response.redirect "msg_procedure_form_new.asp?strCore_ID=" & strCore_ID
'	end if

	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "MetCal_GetProcedures"

	'response.write("strKeyword: " & strKeyword & "<BR>")
	'response.write("strCalibrator: " & strCalibrator & "<BR>")
	'response.write("strAction: " & strAction & "<BR>")
	'response.write("strESToreShopper: " & strEStoreShopper & "<BR>")

	Set prm = cmd.CreateParameter("@strKeyword", adVarchar,adParamInput ,50, strKeyword & "")
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@strCalibrator", adVarchar,adParamInput ,50, strCalibrator & "")
	cmd.Parameters.Append prm

	Set prm = cmd.CreateParameter("@strAction", adVarchar,adParamInput ,50, strAction & "")
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@strEStoreShopper", adVarchar,adParamInput ,255, strEStoreShopper & "")
	cmd.Parameters.Append prm

	Set rsProcedures = Server.CreateObject("ADODB.Recordset")
	rsProcedures.CursorLocation = adUseClient
	rsProcedures.CursorType = adOpenDynamic
	rsProcedures.open cmd

'	response.write("eof: " & rsprocedures.eof & "<BR>")
'	response.write("recordcount: " & rsprocedures.recordcount & "<BR>")

	if not rsProcedures.eof then
		iSearchResult_Recordcount = rsProcedures.RecordCount
	else
		iSearchResult_Recordcount = -1
	end if

	set prm = nothing
	set cmd = nothing

	set Get_SearchResult_Procedures = rsProcedures
End Function

' --------------------------------------------------------------------------------------
' Get a list of procedures this user has bought
' --------------------------------------------------------------------------------------

Function GetEStoreShopper(iSitewide_ID, bCoreExists, strEmail, strPassword, strCore_ID)
	Dim cmd, prm
	Dim rsShopper, rsBoughtProcedures
	Dim strShopperID

	if iSitewide_ID > 0 or bCoreExists = true then

	  Set cmd = Server.CreateObject("ADODB.Command")
	  Set cmd.ActiveConnection = eConn
	  cmd.CommandType = adCmdStoredProc

		'response.write "iSitewide_ID: " & iSitewide_ID & "<BR>"
		'response.write "bCoreExists: " & bCoreExists & "<BR>"
		'response.write "strEmail: " & strEmail & "<BR>"
		'response.write "strPassword: " & strPassword & "<BR>"
		'response.write "strCore_ID: " & strCore_ID & "<BR>"
		
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
		strShopperID = rsShopper("shopper_ID")
		  rsShopper.close
		  set rsShopper = nothing
	  end if
	end if
	
	' Finally, now that we have the procedures they've purchased, translate this to the zipfilenames
	'	since procedure_id's are unique but there could be multiple procedures with the same zipfilename
	
	GetEStoreShopper = strShopperID

End Function

' --------------------------------------------------------------------------------------
' Build the top nav and header
' --------------------------------------------------------------------------------------

Function StartPage(strKeyword, strPrimaryCalibrator, iProcedureCount)
  %>
	<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
  <%

	Dim Screen_Title          ' Window Title
	Dim Bar_Title             ' Black Bar Title
	Dim Content_width	  ' Percent

	BackURL = Session("BackURL")

	Top_Navigation  = False
	Side_Navigation = True
	Screen_Title    = Site_Description & " - " & Translate("Calibration Procedure Search Results",Alt_Language,conn)
	Bar_Title       = Site_Description & "<BR><SPAN CLASS=MediumBoldGold>" & Translate("Calibration Procedure Search Results",Login_Language,conn) & "</SPAN>"
	Content_Width   = 95
  %>

	<!--#include virtual="/SW-Common/SW-Header.asp"-->
	<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->

  <%
	response.write("<!-- Start Content -->" & vbcrlf)
	'response.write("<table border=1 width=100% ><tr><td>")
	response.write "<SPAN CLASS=Heading3>" & Translate("Calibration Procedure Download",Login_Language,conn) & "</SPAN><BR>"
	response.write "<SPAN CLASS=Heading4>" & Translate("Search Results",Login_Language,conn) & "</SPAN><BR><BR>"

	response.write "<LI>" & Translate("Search for",Login_Language,conn) & ": "
	if strKeyword = "" then
	  response.write Translate("Not Specified",Login_Language,conn)
	else
	  response.write "<SPAN CLASS=NormalRed>" & strKeyword & "</SPAN>"
	end if
	response.write "</LI>"
	response.write "<LI>" & Translate("Primary Calibrator",Login_Language,conn) & ": "
	if strPrimaryCalibrator = "" then
	  response.write Translate("Not Specified",Login_Language,conn)
	else
	  response.write "<SPAN CLASS=NormalRed>" & strPrimaryCalibrator & "</SPAN>"
	end if
	response.write "</LI>"

	if iProcedureCount > 650 then
		Response.write "<LI><SPAN CLASS=NormalRed>" & iProcedureCount & " " & Translate("calibration procedures have been found; not all of them can be listed here.  Please refine your search.",Login_Language,conn) & "</LI>"
	end if

	response.write "<BR>"
End Function

' --------------------------------------------------------------------------------------
' Build the search results table header
' --------------------------------------------------------------------------------------

Function WriteTableHeader()
	Dim strOS

	strOS = Request.ServerVariables("HTTP_USER_AGENT")
	if instr(strOS, "MSIE 5.5") then
		response.write "<SPAN class=""smallboldred"">" & Translate("To download a procedure, right-click <u>Download</u> and choose ""Save Target As...""",Login_Language,conn) & "</SPAN><BR><BR>"
	end if

	response.write "<SPAN class=""smallred"">" & Translate("Click on the Instrument Name to see more details about that procedure.",Login_Language,conn) & "</SPAN><P>"
  Call Table_Begin
    response.write "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=2 WIDTH=""100%"">"

    response.write "<TR>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"">" & Translate("Instrument",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"">" & Translate("Date",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER>" & Translate("Rev",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"">" & Translate("Calibrator",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER>" & Translate("Source",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER>" & Translate("5500/CAL",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER>" & Translate("Download",Login_Language,conn) & "</TD>"
	  if request("lv") = "gold" then
	      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER>" & Translate("Value",Login_Language,conn) & "</TD>"
	  else
	      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER>" & Translate("Price",Login_Language,conn) & "</TD>"
	  end if
    response.write "</TR>"

End Function

' --------------------------------------------------------------------------------------
' Output the search results
' --------------------------------------------------------------------------------------

Function DisplaySearchResults(rsProcedures, aBoughtProcedures, iLimit, strCore_ID, iSitewide_ID)

  Dim strOrderStatus
  Dim strOrderDisplay
  Dim strMoreInfoURL
  Dim strEStorePurchase
  Dim strEStoreStatus
  Dim ESTORE_TOKEN
  Dim iCounter

  iCounter = 0

  ESTORE_TOKEN = "<ESTORE_ADDRESS>"

  do while (not rsProcedures.EOF) AND (iCounter < iLimit)
  
  	strOrderStatus = ""
  	'response.write("Order Status: " & rsprocedures("OrderStatus") & "<BR>")
  	if rsProcedures("OrderStatus") <> "" then
  		strOrderStatus = rsProcedures("OrderStatus") & ""
  	end if
  
    strEStorePurchase = "<a href=""javascript:OpenWindow('http://<ESTORE_ADDRESS>/estore/metcal_order.asp?iSitewide_ID=" & iSitewide_ID & "&strCore_ID=" & strCore_ID & "&iProcedure_ID=" &  rsProcedures("Procedure_ID") & "&proc_type=metcal')""><SPAN CLASS=NavLeftHighlight1>&nbsp;" & Translate("Purchase On-Line",Login_Language,conn) & "&nbsp;</SPAN></a>"
  	strEStoreStatus = "<a href=""javascript:OpenWindow('http://<ESTORE_ADDRESS>/estore/welcome_lookup.asp?goWhere=stat')""><SPAN CLASS=NavLeftHighlight1>&nbsp;" & strOrderStatus & "&nbsp;</SPAN></a>"

    strEStorePurchase = replace(strEStorePurchase, ESTORE_TOKEN, "buy.fluke.com")
    strEStoreStatus = replace(strEStoreStatus, ESTORE_TOKEN, "buy.fluke.com")
  
    response.write "<TR VALIGN=TOP>"
  
    rsdoc = Server.HTMLEncode(rsProcedures.Fields("Instrument").Value)
  
    if InStr(1, rsDoc, ":") > 1 then
    	bNewInstrument = true
    	if Mid(rsDoc, 1, InStr(1, rsDoc, ":")) = ToggleStr then
      bNewInstrument = false
    	end if
    	ToggleStr = Mid(rsDoc, 1, InStr(1, rsDoc, ":"))
    else
    	bNewInstrument = true
    end if
  
    if bNewInstrument = true then
    	if ToggleColor = "#EEEEEE" then
      ToggleColor = "#FFFFFF"
    	else
      ToggleColor = "#EEEEEE"
    	end if
    end if
    
    strMoreInfoURL = "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """>"
    strMoreInfoURL = strMoreInfoURL & "<a href=""javascript:OpenWindow('msg_Procedure_MoreInfo.asp?iSitewide_ID=" & iSitewide_ID & "&Language=" & Login_Language & "&strCore_ID=" & strCore_ID & "&iProcedure_ID=" &  rsProcedures("Procedure_ID") & "')"" title=""Click for more detailed information."">"
    strMoreInfoURL = strMoreInfoURL & Server.HTMLEncode(rsProcedures.Fields("Instrument").Value & "") & "</A></TD>"
    response.write strMoreInfoURL
  
    response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """ NOWRAP>" & Server.HTMLEncode(rsProcedures.Fields("ProcDate").Value & "") & "</TD>"
  
    response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("Rev").Value & "") & "</TD>"
    response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("PrimCalibrator").Value & "") & "</TD>"
  
    response.write "<TD CLASS=Small ALIGN=CENTER NOWRAP BGCOLOR="""
    if instr(1,LCase(Server.HTMLEncode(rsProcedures.Fields("Source").Value & "")),"fluke") then
      response.write "#CCFFCC"">"
    elseif instr(1,LCase(Server.HTMLEncode(rsProcedures.Fields("Source").Value & "")),"warrant") then
      response.write "#FFCC00"">"
    else  
      response.write "Yellow"">"
    end if  
    response.write Server.HTMLEncode(rsProcedures.Fields("Source").Value & "") & "</TD>"
  
    response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""" & ToggleColor & """>"
    if Server.HTMLEncode(rsProcedures.Fields("5500/CAL").Value & "") = "True" then
      response.write "<IMG SRC=""/images/required.gif"" WIDTH=10 HEIGHT=10 ALIGN=BOTTOM>"
    else
      response.write "&nbsp;" 'Server.HTMLEncode(rsProcedures.Fields("5500/CAL").Value)
    end if
    response.write "</TD>"
  
    response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""" & ToggleColor & """>"
    strListPrice = rsProcedures.Fields("ListPrice").Value & ""
    if Trim(strListPrice) = "" then strListPrice = "0"
  
  	'response.write "strListPrice: " & strListPrice & "<BR>"
	  'response.write "lv: " & request("lv") & "<BR>"
  	'response.write "strOrderStatus: " & strOrderStatus & "<BR>"

    if (strListPrice > 0 and uCase(request("lv")) = "GOLD") or (strListPrice = 0) then
	  	'response.write "check 1" & "<BR>"
      response.write "<A HREF=""/SW-Common/MSG_Download_File.asp?Category=MetCal&Filename=" & Server.URLEncode(rsProcedures.Fields("DOWNLOAD").Value) & """>" & Translate("Download",Login_Language,conn) & "</A>"
    elseif strListPrice > 0 and isblank(request("lv")) and strOrderStatus = "" then
  		'response.write "check 2" & "<BR>"
      response.write "<SPAN CLASS=Navlefthighlight1>&nbsp;" & strEStorePurchase & "&nbsp;</SPAN>"
	  else
	  	'response.write("strOrderStatus: " & strOrderStatus & "<BR>")
		  'response.write "check 3" & "<BR>"
    	if uCase(strOrderStatus) = "ORDER SHIPPED" then
	  		'response.write "check 4" & "<BR>"
  	    response.write "<A HREF=""/SW-Common/MSG_Download_File.asp?Category=MetCal&Filename=" & Server.URLEncode(rsProcedures.Fields("DOWNLOAD").Value) & """>" & Translate("Download",Login_Language,conn) & "</A>"
  		else
	  		'response.write "check 5" & "<BR>"
	      response.write strEStoreStatus
		  end if
    end if                  
    response.write "</TD>"

  	response.write("<TD CLASS=Small ALIGN=right BGCOLOR=""" & ToggleColor & """>")
	
    if IsNumeric(strListPrice) then response.write("$" & FormatCurrency(cLng(strListPrice))/100 & " (USD)")
  	response.write("</td>")
    response.write "</TR>"

    iCounter = iCounter + 1
    rsProcedures.moveNext
       
  Loop

  response.write "</TABLE>"

End Function

' --------------------------------------------------------------------------------------
' Output a friendly message saying to results were found
' --------------------------------------------------------------------------------------

Function Output_SearchEmpty(strCore_ID)

  response.write "<LI><SPAN CLASS=NormalBoldRed>Sorry.</SPAN>&nbsp;&nbsp;" & Translate("There are no calibration procedures found that match your search critera.  Please adjust your criteria and try again.",Login_Language,conn) & "</LI><BR><BR>"
 	response.write "<table><tr><td>"
	response.write "<A HREF=""MSG_Procedure_Form.asp?strCore_ID=" & strCore_ID
  if request("lv") <> "" then
    response.write "?lv=" & request("lv")
  end if  
  response.write """><SPAN CLASS=Navlefthighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</SPAN></A>" 
  response.write "</td></TR></TABLE>"
 	response.write "<BR>"
  
End Function


' --------------------------------------------------------------------------------------
' Write out the recordset page buttons
' --------------------------------------------------------------------------------------

Sub Write_Nav_Buttons(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel, strCore_ID)

  Dim strLevelOutput1
  Dim strLevelOutput2
  Dim strResults
  Dim strPreviousPage
  Dim strNextPage
  Dim strNewSearch

  if strLevel <> "" then
	strLevelOutput1 = "?lv=" & strLevel
	strLevelOutput2 = "&lv=" & strLevel
  else
  strLevelOutput1 = "?lv="
  strLevelOutput2 = ""
  end if

  strPreviousPage = "<a href=""MSG_Procedure_Results.asp?CurrPage=" & iCurrPage - 1 & "&Keyword=" & strKeyword & "&Calibrator=" & strCalibrator & strLevelOutput2 & "&strCore_ID=" & strCore_ID & """>"
  strPreviousPage = strPreviousPage & "<SPAN CLASS=Navlefthighlight1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</A></SPAN>&nbsp;&nbsp;"

  strNextPage = "<a href=""MSG_Procedure_Results.asp?CurrPage=" & iCurrPage + 1 & "&Keyword=" & strKeyword & "&Calibrator=" & strCalibrator & strLevelOutput2 & "&strCore_ID=" & strCore_ID & """>"
  strNextPage = strNextPage & "<SPAN CLASS=Navlefthighlight1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</A></SPAN>&nbsp;&nbsp;"

  strNewSearch = "<A HREF=""MSG_Procedure_Form.asp" & strLevelOutput1 & "&strCore_ID=" & strCore_ID & """><SPAN CLASS=Navlefthighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</SPAN></A>"

  if iNumPages > 1 then
  	response.write "<BR>"
  
    'response.write("iCurrPage: " & iCurrPage & "<BR>")
    'response.write("iNumPages: " & iNumPages & "<BR>")

  	if cInt(iCurrPage) = 1 then
      Call Nav_Border_Begin
  		response.write "<table border=0><tr>"
  		Call Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel, strCore_ID)
  		response.write "<td>" & strNextPage & "</td>"
  		response.write("<td>&nbsp;</td>")
	  	response.write "<td>" & strNewSearch & "</td>"
  		response.write "</tr>"
      response.write "</table>"
      Call Nav_Border_End
  	else
  		if cInt(iCurrPage) = cInt(iNumPages) then
        Call Nav_Border_Begin
  			response.write "<table BORDER=0><tr>"
  			response.write "<TD>" & strPreviousPage & "</td>"
	  		Call Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel, strCore_ID)
	  		response.write "<td>" & strNewSearch & "</td>"
  			response.write "</tr></table>"
        Call Nav_Border_End
  		else
        Call Nav_Border_Begin
  			response.write "<table BORDER=0><tr>"
  			response.write "<td>" & strPreviousPage & "</td>"
    		Call Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel, strCore_ID)
  			response.write "<td>" & strNextPage  & "</td>"
	  		response.write "<td>" & strNewSearch & "</td>"
  			response.write "</tr></table>"
        Call Nav_Border_End
  		end if
  	end if
  
  	response.write "<BR>"
        
  else

  	response.write "<BR>"
    Call Nav_Border_Begin
   	response.write "<table><tr><td>"
  	response.write strNewSearch
  	response.write "</td></TR></TABLE>"
    Call Nav_Border_End
   	response.write "<BR>"

  end if

end sub

' --------------------------------------------------------------------------------------
' Write out the individual recordset navigation pages
' --------------------------------------------------------------------------------------

sub Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel, strCore_ID)

  Dim iCounter

  response.write "<TD CLASS=MediumBold>"
  
  for iCounter = 1 to iNumPages
	if iCounter = 26 then
	  	exit for
	else
	  	response.write "<A HREF=""MSG_Procedure_Results.asp?currpage=" & iCounter & "&Keyword=" & strKeyword & "&Calibrator=" & strCalibrator & "&strCore_ID=" & strCore_ID

		if strLevel <> "" then
			response.write "&lv=" & strLevel
		end if
 		if iCounter = CInt(iCurrPage) then
			response.write  """><SPAN CLASS=NavTopHighLight>&nbsp;"
		else
			response.write """><SPAN CLASS=Navlefthighlight1>&nbsp;"
		end if
		if iCounter < 10 then response.write "&nbsp;"
		response.write cstr(iCounter) & "&nbsp;</SPAN></A>&nbsp;&nbsp;"
	end if
  next
  
  response.write "</TD>"

end sub

' --------------------------------------------------------------------------------------

Function OutputFooter()
  %>
  <!--/td></tr></table-->
  <%Call Table_End%>
	</DIV>
  <%
End Function

' --------------------------------------------------------------------------------------

Function OutputJavaScript()
%>
	<script language="javascript">
	var oEStore

	function OpenWindow(strTarget){
	  oEStore = window.open(strTarget, 'windowname', 'status=yes,height=400,width=800,scrollbars=1,resizable=1,toolbar=1');
	  oEStore.focus();
	}
	</script>
<%
End Function

'--------------------------------------------------------------------------------------
  
%>

