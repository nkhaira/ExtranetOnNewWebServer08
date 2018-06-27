<%@ LANGUAGE="VBSCRIPT"%>

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

Call Connect_SiteWide

%>

<%
'#include virtual="/SW-Common/SW-Security_Module.asp" -->

site_id = 11

' --------------------------------------------------------------------------------------
' Get recordset of procedures based on search criteria
' --------------------------------------------------------------------------------------
Dim strMake, strModel, strCalType, strMinSoftVer, strOption
Dim cmd
Dim prm
Dim rsProcedures
Dim iRecordCount

strMake = Trim(Replace(Replace(UCase(Request("make")),"FLUKE-","FLUKE "), "*", "%"))
strModel = Trim(Replace(Replace(UCase(Request("model")),"FLUKE-","FLUKE "), "*", "%"))
strCalType = Trim(Replace(Replace(UCase(Request("caltype")),"FLUKE-","FLUKE "), "*", "%"))
strMinSoftVer = Trim(Replace(Replace(UCase(Request("minsoftver")),"FLUKE-","FLUKE "), "*", "%"))
strOption = Trim(Replace(Replace(UCase(Request("option")),"FLUKE-","FLUKE "), "*", "%"))

'response.write("strMake; " & strMake & "<BR>")
'response.write("strModel: " & strModel & "<BR>")
'response.write("strCalType: " & strCalType & "<BR>")
'response.write("strMinSoftVer: " & strMinSoftVer & "<BR>")
'response.write("strOption: " & strOption & "<BR>")

Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "PortoCal_GetMainProcedures"

Set prm = cmd.CreateParameter("@strMake", adVarchar,adParamInput ,50, strMake & "")
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("@strModel", adVarchar,adParamInput ,50, strModel & "")
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("@strCalType", adVarchar,adParamInput ,50, strCalType & "")
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("@strMinSoftVer", adVarchar,adParamInput ,50, strMinSoftVer & "")
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("@strOption", adVarchar,adParamInput ,50, strOption & "")
cmd.Parameters.Append prm

Set rsProcedures = Server.CreateObject("ADODB.Recordset")
rsProcedures.CursorLocation = adUseClient
rsProcedures.CursorType = adOpenDynamic
rsProcedures.open cmd

'response.write("eof: " & rsprocedures.eof & "<BR>")
'response.write("recordcount: " & rsprocedures.recordcount & "<BR>")

set prm = nothing
set cmd = nothing

iRecordCount = rsProcedures.RecordCount

' --------------------------------------------------------------------------------------
' Figure out how many pages there are in this recordset based upon 25 records per page
' --------------------------------------------------------------------------------------

Dim iLimit
Dim iNumPages
Dim iExtraPage
Dim iCurrPage

iLimit = 25
iCurrPage = Request("CurrPage")
if IsNumeric(iCurrPage) then
	if iCurrPage = 0 then
		iCurrPage = 1
	end if
else
	iCurrPage = 1
end if

rsProcedures.PageSize = iLimit
rsProcedures.AbsolutePage = cint(iCurrPage)

iNumPages = rsProcedures.PageCount

if iNumPages > 25 then
	iNumPages = 25
end if

' --------------------------------------------------------------------------------------
' Misc Declarations
' --------------------------------------------------------------------------------------
Dim strLevel
Dim iCounter
Dim strHREF

strLevel = ucase(trim(request("lv")))

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

site_description = "Portocal"

' --------------------------------------------------------------------------------------
' Start building the page
' --------------------------------------------------------------------------------------
Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_width	  ' Percent

Screen_Title    = Site_Description & " - " & Translate("Portocal II Procedure Search Results",Login_Language,conn)
Bar_Title       = Site_Description & "<BR><FONT CLASS=MediumBoldGold>" & Translate("Portocal II Procedure Search Results",Login_Language,conn) & "</FONT>"
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95

BackURL = Session("BackURL")
%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->

<%
response.write "<FONT CLASS=Heading3>" & Translate("Portocal II Procedure Download",Login_Language,conn) & "</FONT><BR>"
response.write "<FONT CLASS=Heading4>" & Translate("Search Results",Login_Language,conn) & "</FONT><BR><BR>"

response.write "<LI>" & Translate("Manufacturer",Login_Language,conn) & ": "
if strMake = "" then
  response.write Translate("Not Specified",Login_Language,conn)
else
  response.write "<FONT CLASS=NormalRed>" & strMake & "</FONT>"
end if
response.write "</LI>"
response.write "<LI>" & Translate("Model",Login_Language,conn) & ": "
if strModel = "" then
  response.write Translate("Not Specified",Login_Language,conn)
else
  response.write "<FONT CLASS=NormalRed>" & strModel & "</FONT>"
end if
response.write "</LI>"
response.write "<LI>" & Translate("Calibrator Type",Login_Language,conn) & ": "
if strCalType = "" then
  response.write Translate("Not Specified",Login_Language,conn)
else
  response.write "<FONT CLASS=NormalRed>" & strCalType & "</FONT>"
end if
response.write "</LI>"
response.write "<LI>" & Translate("Minimum Software Version",Login_Language,conn) & ": "
if strMinSoftVer = "" then
  response.write Translate("Not Specified",Login_Language,conn)
else
  response.write "<FONT CLASS=NormalRed>" & strMinSoftVer & "</FONT>"
end if
response.write "</LI>"
response.write "<LI>" & Translate("Option",Login_Language,conn) & ": "
if strOption = "" then
  response.write Translate("Not Specified",Login_Language,conn)
else
  response.write "<FONT CLASS=NormalRed>" & strOption & "</FONT>"
end if
response.write "</LI>"

if iRecordCount > 650 then
    Response.write "<LI><FONT CLASS=NormalRed>" & iRecordCount & " " & Translate("calibration procedures have been found; not all of them can be listed here.  Please refine your search.",Login_Language,conn) & "</LI>"
end if

response.write "<BR><BR>"

if not rsProcedures.EOF then

    Call Write_Nav_Buttons(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)

    response.write "<TABLE BORDER=1 BORDERCOLOR=""Gray"" CELLSPACING=0 CELLPADDING=2 WIDTH=""100%"" BGCOLOR=""#EEEEEE"">"

    response.write "<TR>"
      response.write "<TD CLASS=MEDIUMBOLDGOLD BGCOLOR=""#000000"">" & Translate("Procedure Name",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=MEDIUMBOLDGOLD BGCOLOR=""#000000"">" & Translate("Manufacturer",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=MEDIUMBOLDGOLD BGCOLOR=""#000000"">" & Translate("Model",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=MEDIUMBOLDGOLD BGCOLOR=""#000000"">" & Translate("Calibrator Type",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=MEDIUMBOLDGOLD BGCOLOR=""#000000"">" & Translate("Option",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=MEDIUMBOLDGOLD BGCOLOR=""#000000"">" & Translate("Download",Login_Language,conn) & "</TD>"
    response.write "</TR>"
    response.write "<TBODY>"

    ToggleColor = "#FFFFFF"
    ToggleStr = ""

  for iCounter = 1 to rsProcedures.PageSize
    response.write "<TR VALIGN=TOP>"

      response.write "<TD CLASS=MEDIUM BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("ProcName").Value & "") & "</TD>"

      response.write "<TD CLASS=MEDIUM BGCOLOR=""" & ToggleColor & """ NOWRAP>" & Server.HTMLEncode(rsProcedures.Fields("Make").Value & "") & "</TD>"

      response.write "<TD CLASS=MEDIUM ALIGN=CENTER BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("Model").Value & "") & "</TD>"
      response.write "<TD CLASS=MEDIUM BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("CalType").Value & "") & "</TD>"
      response.write "<TD CLASS=MEDIUM BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("Options").Value & "") & "</TD>"
      strHREF = "<a href=""/upload/portocal/procedures/" & rsProcedures.Fields("FileName").value & """>"
      response.write "<TD CLASS=MEDIUM BGCOLOR=""" & ToggleColor & """>" & strHREF & Server.HTMLEncode(rsProcedures.Fields("FileName").Value & "") & "</a></TD>"
    response.write "</TR>"

    rsProcedures.moveNext
	if rsProcedures.EOF then exit for
  Next

  Response.write "</TBODY>"
  response.write "</TABLE>"

'response.write("iCounter: " & iCounter & "<BR>")

  Call Write_Nav_Buttons(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  
else

  response.write "<LI><FONT CLASS=NormalBoldRed>Sorry.</FONT>&nbsp;&nbsp;" & Translate("There are no calibration procedures found that match your search critera.  Please adjust your criteria and try again.",Login_Language,conn) & "</LI><BR><BR>"
 	response.write "<table><tr><td>"
	response.write "<A HREF=""Portocal_Procedure_Form.asp"
  if request("lv") <> "" then
    response.write "?lv=" & request("lv")
  end if  
  response.write """><FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>" 
  response.write "</td></TR></TABLE>"
 	response.write "<BR>"
  
end if

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------

Sub Write_Nav_Buttons(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  Dim strLevelOutput1
  Dim strLevelOutput2
  Dim strResults
  Dim strPreviousPage
  Dim strNextPage
  Dim strNewSearch

  if strLevel <> "" then
	strLevelOutput1 = "?lvl=" & strLevel
	strLevelOutput2 = "&lvl=" & strLevel
  end if

strMake = Trim(Replace(Replace(UCase(Request("make")),"FLUKE-","FLUKE "), "*", "%"))
strModel = Trim(Replace(Replace(UCase(Request("model")),"FLUKE-","FLUKE "), "*", "%"))
strCalType = Trim(Replace(Replace(UCase(Request("caltype")),"FLUKE-","FLUKE "), "*", "%"))
strMinSoftVer = Trim(Replace(Replace(UCase(Request("minsoftver")),"FLUKE-","FLUKE "), "*", "%"))
strOption = Trim(Replace(Replace(UCase(Request("option")),"FLUKE-","FLUKE "), "*", "%"))

  strPreviousPage = "<a href=""Portocal_Procedure_Results.asp?CurrPage=" & iCurrPage - 1 & "&Make=" & strMake & "&Model=" & strModel & "&CalType=" & strCalType & "&MinSoftVer=" & strMinSoftVer & "&Option=" & strOption & strLevelOutput2 & """>"
  strPreviousPage = strPreviousPage & "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</A></FONT>&nbsp;&nbsp;"

  strNextPage = "<a href=""Portocal_Procedure_Results.asp?CurrPage=" & iCurrPage + 1 & "&Keyword=" & "&Make=" & strMake & "&Model=" & strModel & "&CalType=" & strCalType & "&MinSoftVer=" & strMinSoftVer & "&Option=" & strOption & strLevelOutput2 & """>"
  strNextPage = strNextPage & "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</A></FONT>&nbsp;&nbsp;"

  strNewSearch = "<A HREF=""Portocal_Procedure_Form.asp" & strLevelOutput1 & """><FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;New Search&nbsp;&nbsp;</FONT></A>"

  if iNumPages > 1 then
  	response.write "<BR>"
  
'response.write("iCurrPage: " & iCurrPage & "<BR>")
'response.write("iNumPages: " & iNumPages & "<BR>")

  	if cInt(iCurrPage) = 1 then
		response.write "<table><tr>"
  		Call Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  		response.write "<td>"
		response.write strNextPage
		response.write strNewSearch
  		response.write "</td></tr></table>"
	else
  		if cInt(iCurrPage) = cInt(iNumPages) then
  			response.write "<table><tr><td>"
			response.write strPreviousPage
  			response.write "</td>"
	  		Call Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
			response.write "<td>" & strNewSearch & "</td>"
  			response.write "</tr></table>"
  		else
  			response.write "<table><tr><td>"
			response.write strPreviousPage
  			response.write "</td>"
	    		Call Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  			response.write "<td>"
			response.write strNextPage  			
			response.write strNewSearch
  			response.write "</td></tr></table>"
  		end if
  	end if
  
  	response.write "<BR>"
        
  else

	response.write "<BR>"
     	response.write "<table><tr><td>"
	response.write strNewSearch
	response.write "</td></TR></TABLE>"
     	response.write "<BR>"

  end if

End Sub

' --------------------------------------------------------------------------------------

Sub Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  Dim iCounter

  response.write "<TD CLASS=MediumBold>"
  
  for iCounter = 1 to iNumPages
	if iCounter = 26 then
	  	exit for
	else
	  	response.write "<A HREF=""Portocal_Procedure_Results.asp?currpage=" & iCounter & "&Make=" & strMake & "&Model=" & strModel & "&CalType=" & strCalType & "&MinSoftVer=" & strMinSoftVer & "&Option=" & strOption 

		if strLevel <> "" then
			response.write "&lv=" & strLevel
		end if
  		if iCounter = CInt(iCurrPage) then
			response.write  """><FONT CLASS=NavTopHighLight>&nbsp;"
		else
			response.write """><FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;"
		end if
		if iCounter < 10 then response.write "&nbsp;"
		response.write cstr(iCounter) & "&nbsp;</FONT></A>&nbsp;&nbsp;"
	end if
  next
  
  response.write "</TD>"

end sub
%>
