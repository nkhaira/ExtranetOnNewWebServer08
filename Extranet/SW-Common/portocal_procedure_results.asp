<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
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

Dim Border_Toggle
Border_Toggle = 0

strMake = Trim(Replace(Replace(UCase(Request("make")),"FLUKE-","FLUKE "), "*", "%"))
strModel = Trim(Replace(Replace(UCase(Request("model")),"FLUKE-","FLUKE "), "*", "%"))
strCalType = Trim(Replace(Replace(UCase(Request("caltype")),"FLUKE-","FLUKE "), "*", "%"))
strMinSoftVer = Trim(Replace(Replace(UCase(Request("minsoftver")),"FLUKE-","FLUKE "), "*", "%"))
strOption = Trim(Replace(Replace(UCase(Request("option")),"FLUKE-","FLUKE "), "*", "%"))

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

if not rsProcedures.EOF then
	rsProcedures.PageSize = iLimit
	rsProcedures.AbsolutePage = cint(iCurrPage)
	iNumPages = rsProcedures.PageCount
end if

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

Screen_Title    = Site_Description & " - " & Translate("Portocal II Procedure Search Results",Alt_Language,conn)
Bar_Title       = Site_Description & "<BR><SPAN CLASS=MediumBoldGold>" & Translate("Portocal II Procedure Search Results",Login_Language,conn) & "</SPAN>"
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95

BackURL = Session("BackURL")
%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->

<%
response.write "<SPAN CLASS=Heading3>" & Translate("Portocal II Procedure Download",Login_Language,conn) & "</SPAN><BR>"
response.write "<SPAN CLASS=Heading4>" & Translate("Search Results",Login_Language,conn) & "</SPAN><BR><BR>"

response.write "<LI>" & Translate("Manufacturer",Login_Language,conn) & ": "
if strMake = "" then
  response.write Translate("Not Specified",Login_Language,conn)
else
  response.write "<SPAN CLASS=NormalRed>" & strMake & "</SPAN>"
end if
response.write "</LI>"
response.write "<LI>" & Translate("Model",Login_Language,conn) & ": "
if strModel = "" then
  response.write Translate("Not Specified",Login_Language,conn)
else
  response.write "<SPAN CLASS=NormalRed>" & strModel & "</SPAN>"
end if
response.write "</LI>"
response.write "<LI>" & Translate("Calibrator Type",Login_Language,conn) & ": "
if strCalType = "" then
  response.write Translate("Not Specified",Login_Language,conn)
else
  response.write "<SPAN CLASS=NormalRed>" & strCalType & "</SPAN>"
end if
response.write "</LI>"
response.write "<LI>" & Translate("Minimum Software Version",Login_Language,conn) & ": "
if strMinSoftVer = "" then
  response.write Translate("Not Specified",Login_Language,conn)
else
  response.write "<SPAN CLASS=NormalRed>" & strMinSoftVer & "</SPAN>"
end if
response.write "</LI>"
response.write "<LI>" & Translate("Option",Login_Language,conn) & ": "
if strOption = "" then
  response.write Translate("Not Specified",Login_Language,conn)
else
  response.write "<SPAN CLASS=NormalRed>" & strOption & "</SPAN>"
end if
response.write "</LI>"

if iRecordCount > 650 then
    Response.write "<LI><SPAN CLASS=NormalRed>" & iRecordCount & " " & Translate("calibration procedures have been found; not all of them can be listed here.  Please refine your search.",Login_Language,conn) & "</SPAN></LI>"
end if

response.write "<BR><BR>"

if not rsProcedures.EOF then

    Call Write_Nav_Buttons(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)

    Call Table_Begin
    response.write "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=2 WIDTH=""100%"">"

    response.write "<TR>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"">" & Translate("Manufacturer",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER>" & Translate("Model",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"">" & Translate("Procedure Name",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER>" & Translate("Calibrator Type",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER>" & Translate("Option",Login_Language,conn) & "</TD>"
      response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER>" & Translate("Download",Login_Language,conn) & "</TD>"
    response.write "</TR>"
    response.write "<TBODY>"

    ToggleColor = "#FFFFFF"
    ToggleStr = ""

    Old_Make  = ""
    Old_Model = ""
    
    for iCounter = 1 to rsProcedures.PageSize
  
    if (Old_Make <> rsProcedures.Fields("Make").Value) or (Old_Model <> rsProcedures.Fields("Model").Value) then
      if ToggleColor = "#FFFFFF" then
        ToggleColor = "#EEEEEE"
      else
        ToggleColor = "#FFFFFF"
      end if
    end if
    Old_Make  = rsProcedures.Fields("Make").Value
    Old_Model = rsProcedures.Fields("Model").Value    
    
    response.write "<TR VALIGN=TOP>"

      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """ NOWRAP>" & Server.HTMLEncode(rsProcedures.Fields("Make").Value & "") & "</TD>"

      response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("Model").Value & "") & "</TD>"
      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("ProcName").Value & "") & "</TD>"
      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """ ALIGN=CENTER>"
      if not isblank(rsProcedures.Fields("CalType").Value) then
        response.write Server.HTMLEncode(rsProcedures.Fields("CalType").Value & "")
      else
        response.write "&nbsp;"
      end if    
      response.write "</TD>"

      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """ ALIGN=CENTER>"
      if not isblank(rsProcedures.Fields("Options").Value) then
        response.write Server.HTMLEncode(rsProcedures.Fields("Options").Value & "")
      else
        response.write "&nbsp;"
      end if    
      strHREF = "<a href=""/upload/portocal/procedures/" & rsProcedures.Fields("FileName").value & """>"
      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """ ALIGN=CENTER>" & strHREF & "<SPAN CLASS=NavLeftHighlight1>&nbsp;" & Translate("Download",Login_Language,conn) & "&nbsp;</SPAN></A></TD>"
    response.write "</TR>"

    rsProcedures.moveNext
	if rsProcedures.EOF then exit for
  Next

  Response.write "</TBODY>"
  response.write "</TABLE>"
  Call Table_End

'response.write("iCounter: " & iCounter & "<BR>")

  Call Write_Nav_Buttons(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  
else

  response.write "<LI><SPAN CLASS=NormalBoldRed>Sorry.</SPAN>&nbsp;&nbsp;" & Translate("There are no calibration procedures found that match your search critera.  Please adjust your criteria and try again.",Login_Language,conn) & "</LI><BR><BR>"
 	response.write "<table><tr><td>"
	response.write "<A HREF=""Portocal_Procedure_Form.asp"
  if request("lv") <> "" then
    response.write "?lv=" & request("lv")
  end if  
  response.write """><SPAN CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</SPAN></A>" 
  response.write "</td></TR></TABLE>"
 	response.write "<BR>"
  
end if

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' Subroutines and Functions
'--------------------------------------------------------------------------------------

sub Write_Nav_Buttons(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
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
  strPreviousPage = strPreviousPage & "<SPAN CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</A></SPAN>&nbsp;&nbsp;"

  strNextPage = "<a href=""Portocal_Procedure_Results.asp?CurrPage=" & iCurrPage + 1 & "&Keyword=" & "&Make=" & strMake & "&Model=" & strModel & "&CalType=" & strCalType & "&MinSoftVer=" & strMinSoftVer & "&Option=" & strOption & strLevelOutput2 & """>"
  strNextPage = strNextPage & "<SPAN CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</A></SPAN>&nbsp;&nbsp;"

  strNewSearch = "<A HREF=""Portocal_Procedure_Form.asp" & strLevelOutput1 & """><SPAN CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;New Search&nbsp;&nbsp;</SPAN></A>"

  if iNumPages > 1 then
  	response.write "<BR>"
  
'response.write("iCurrPage: " & iCurrPage & "<BR>")
'response.write("iNumPages: " & iNumPages & "<BR>")

  	if cInt(iCurrPage) = 1 then
    Call Nav_Border_Begin
		response.write "<table><tr>"
  		Call Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  		response.write "<td>"
		response.write strNextPage
		response.write strNewSearch
  		response.write "</td></tr></table>"
    Call Nav_Border_End  
	else
  		if cInt(iCurrPage) = cInt(iNumPages) then
        Call Nav_Border_Begin
  			response.write "<table><tr><td>"
		  	response.write strPreviousPage
  			response.write "</td>"
    		Call Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  			response.write "<td>" & strNewSearch & "</td>"
  			response.write "</tr></table>"
        Call Nav_Border_End  
  		else
        Call Nav_Border_Begin
  			response.write "<table><tr><td>"
  			response.write strPreviousPage
  			response.write "</td>"
	  		Call Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  			response.write "<td>"
  			response.write strNextPage  			
	  		response.write strNewSearch
  			response.write "</td></tr></table>"
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

end Sub

' --------------------------------------------------------------------------------------

sub Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
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
			response.write  """><SPAN CLASS=NavTopHighLight>&nbsp;"
		else
			response.write """><SPAN CLASS=NAVLEFTHIGHLIGHT1>&nbsp;"
		end if
		if iCounter < 10 then response.write "&nbsp;"
		response.write cstr(iCounter) & "&nbsp;</SPAN></A>&nbsp;&nbsp;"
	end if
  next
  
  response.write "</TD>"

end sub

'--------------------------------------------------------------------------------------

%>
