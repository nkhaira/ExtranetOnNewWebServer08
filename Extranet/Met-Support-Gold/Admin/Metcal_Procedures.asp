<%@ LANGUAGE="VBSCRIPT"%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Author:     Kelly. Whitlock
' Date:       2/1/2000
'             Sandbox
' --------------------------------------------------------------------------------------

response.buffer = true

Call Connect_SiteWide

%>

<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
<!-- #include virtual="/SW-Administrator/CK_Admin_Credentials.asp"-->
<!-- #include virtual="/met-support-gold/admin/CK_Credentials.asp"-->

<%
' --------------------------------------------------------------------------------------
' Get recordset of procedures based on search criteria
' --------------------------------------------------------------------------------------

Dim strKeyword
Dim strCalibrator
Dim strFileName
Dim bRestricted
Dim cmd
Dim prm
Dim rsProcedures
Dim iRecordCount

strKeyword    = trim(Replace(Replace(UCase(Request("KeyWord")),"FLUKE-","FLUKE "), "*", "%"))
strCalibrator = trim(Replace(Replace(Request("Calibrator"), "'", ""), "*", "%"))
strFileName   = trim(replace(replace(request("Filename"), "'", ""), "*", "%"))

if request("Restricted") = "-1" then
  bRestricted = true
else
  bRestricted = false
end if  

Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "Admin_MetCal_Procedures_GetList"

'response.write("strKeyword: " & strKeyword & "<BR>")
'response.write("strCalibrator: " & strCalibrator & "<BR>")

Set prm = cmd.CreateParameter("@strKeyword", adVarchar, adParamInput, 50, strKeyword & "")
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("@strCalibrator", adVarchar, adParamInput, 50, strCalibrator & "")
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("@strFileName", adVarchar, adParamInput, 50, strFileName & "")
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

if not rsProcedures.eof then
  rsProcedures.PageSize = iLimit
  rsProcedures.AbsolutePage = iCurrPage
  iNumPages = rsProcedures.PageCount
end if

if iNumPages > 25 then
	iNumPages = 25
end if

' --------------------------------------------------------------------------------------
' Misc Declarations
' --------------------------------------------------------------------------------------
Dim Procedure_ID
Dim strLevel
Dim iCounter

strLevel = ucase(trim(request("lv")))

if not isblank(request("ID")) then
  Procedure_ID = request("ID")
else
  Procedure_ID = 0
end if

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_width	  ' Percent

BackURL = Session("BackURL")

Screen_Title    = Site_Description & " - " & "Calibration Procedure Search Results"
Bar_Title       = Site_Description & "<BR><SPAN CLASS=MediumBoldGold>" & "Calibration Procedure Search Results" & "</SPAN>"
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95

%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->

<%
response.write "<SPAN CLASS=Heading3>Calibration Procedure Download</SPAN><BR>"
response.write "<SPAN CLASS=Heading4>Search Results</SPAN><BR><BR>"

response.write "<LI>Search for: <SPAN CLASS=NormalRed>"
if strKeyword = "" then
  response.write "Not Specified"
else
  response.write strKeyword
end if
response.write "</SPAN></LI>"

response.write "<LI>Primary Calibrator: <SPAN CLASS=NormalRed>"
if strCalibrator = "" then
  response.write "Not Specified"
else
  response.write strCalibrator
end if
response.write "</SPAN></LI>"

response.write "<LI>Zip Filename: <SPAN CLASS=NormalRed>"
if strFilename = "" then
  response.write "Not Specified"
else
  response.write strFilename
end if
response.write "</SPAN></LI>"

if iRecordCount > 650 then
    Response.write "<LI><SPAN CLASS=NormalRed>A total of <B>" & iRecordCount & "</B> calibration procedures have been found; not all of them can be listed here.  Please refine your search.</SPAN></LI>"
end if

response.write "<BR>"

if not rsProcedures.EOF then

    Call Write_Nav_Buttons(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)

    Call Table_Begin
    response.write "<TABLE BORDER=0 BORDERCOLOR=""#666666"" CELLSPACING=0 CELLPADDING=2 WIDTH=""100%"" BGCOLOR=""#EEEEEE"">"

    response.write "<THEAD>"
    response.write "<TR>"
      response.write "<TD CLASS=SMALLBOLDWHITE BGCOLOR=""#FF0000"" ALIGN=CENTER>Action</TD>"
      response.write "<TD CLASS=SMALLBOLDGOLD BGCOLOR=""#000000"" ALIGN=LEFT>Instrument</TD>"
      response.write "<TD CLASS=SMALLBOLDGOLD BGCOLOR=""#000000"" ALIGN=CENTER>Create<BR>Date</TD>"
      response.write "<TD CLASS=SMALLBOLDGOLD BGCOLOR=""#000000"" ALIGN=CENTER>Rev</TD>"
      response.write "<TD CLASS=SMALLBOLDGOLD BGCOLOR=""#000000"" ALIGN=LEFT>Calibrator</TD>"
      response.write "<TD CLASS=SMALLBOLDGOLD BGCOLOR=""#000000"" ALIGN=CENTER>Source</TD>"
      response.write "<TD CLASS=SMALLBOLDGOLD BGCOLOR=""#000000"" ALIGN=CENTER>R</TD>"
      response.write "<TD CLASS=SMALLBOLDGOLD BGCOLOR=""#000000"" ALIGN=CENTER>Add<BR>Date</TD>"
      response.write "<TD CLASS=SMALLBOLDGOLD BGCOLOR=""#000000"" ALIGN=CENTER>Update<BR>By</TD>"
      response.write "<TD CLASS=SMALLBOLDGOLD BGCOLOR=""#000000"" ALIGN=CENTER>Update<BR>Date</TD>"      
    response.write "</TR>"
    response.write "</THEAD>"
    response.write "<TBODY>"

    ToggleColor = "#FFFFFF"
    ToggleStr = ""

  do while (not rsProcedures.EOF) AND (iCounter < iLimit)

    response.write "<TR VALIGN=TOP>"

    rsdoc = Server.HTMLEncode(rsProcedures.Fields("Instrument").Value)
    if Mid(rsdoc,1,instr(1,rsdoc,":")) <> ToggleStr then
      ToggleStr = Mid(rsdoc,1,instr(1,rsdoc,":"))
      if ToggleColor = "#DDDDDD" then
        ToggleColor = "#FFFFFF"
      else
        ToggleColor = "#DDDDDD"
      end if 
    end if
  
      if CInt(Procedure_ID) = CInt(rsProcedures.Fields("Procedure_ID").Value) then
        response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""#FF0000"" VALIGN=MIDDLE>"
      else  
        response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""#666666"" VALIGN=MIDDLE>"
      end if        
      response.write "<A NAME=""PID" & rsProcedures.Fields("Procedure_ID").Value & """></A>"
      response.write "<A HREF=""metcal_procedure.asp?new=false&id=" & rsProcedures.Fields("Procedure_ID") & "&keyword=" & strKeyword & "&calibrator=" & strCalibrator & "&FileName=" & strFileName & "&CurrPage=" & iCurrPage & """><SPAN CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;Edit&nbsp;&nbsp;</A></SPAN>"
      response.write "</TD>"

      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("Instrument").Value & "") & "</TD>"

      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """ NOWRAP ALIGN=RIGHT>" & Server.HTMLEncode(FormatDateTime(rsProcedures.Fields("Date").Value,"0") & "") & "</TD>"
      
      if not isblank(rsProcedures.Fields("Rev").Value) then
        response.write "<TD CLASS=Small NOWRAP ALIGN=CENTER BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("Rev").Value & "") & "</TD>"
      else
        response.write "<TD CLASS=Small NOWRAP ALIGN=CENTER BGCOLOR=""" & ToggleColor & """>?</TD>"      
      end if
      
      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """>" & Server.HTMLEncode(rsProcedures.Fields("Calibrator").Value & "") & "</TD>"

      response.write "<TD CLASS=Small ALIGN=CENTER NOWRAP BGCOLOR="""
      if instr(1,LCase(Server.HTMLEncode(rsProcedures.Fields("Source").Value & "")),"fluke") then
        response.write "#CCFFCC"">"
      elseif instr(1,LCase(Server.HTMLEncode(rsProcedures.Fields("Source").Value & "")),"warant") then
        response.write "#FFCC00"">"
      else  
        response.write "Yellow"">"
      end if  
      response.write Server.HTMLEncode(rsProcedures.Fields("Source").Value & "") & "</TD>"

      response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""" & ToggleColor & """>"
      response.write "<SPAN CLASS=NAVLEFTHIGHLIGHT1>"
      if rsProcedures.Fields("Restricted") then response.write("<img src=""\images\r_square.gif"">")
      response.write "</SPAN></TD>"
      
      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """ ALIGN=RIGHT>" & Server.HTMLEncode(FormatDateTime(rsProcedures.Fields("CreateDate").Value,"0") & "") & "</TD>"      

      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """ ALIGN=CENTER>"
      if not isblank(rsProcedures.Fields("UpdateBy").Value) then
      
        SQLUser = "SELECT Lastname from UserData where ID=" & rsProcedures.Fields("UpdateBy").Value
        Set rsUser = Server.CreateObject("ADODB.Recordset")
        rsUser.Open SQLUser, conn, 3, 3
        
        response.write Server.HTMLEncode(rsUser.Fields("Lastname").Value & "")
        rsUser.close
        set rsUser = nothing
        
      else
        response.write "&nbsp;"
      end if
      response.write "</TD>"
      
      response.write "<TD CLASS=Small BGCOLOR=""" & ToggleColor & """ NOWRAP ALIGN=RIGHT>"
      if FormatDateTime(rsProcedures.Fields("UpdateDate").Value) <> FormatDateTime(rsProcedures.Fields("CreateDate").Value) then
        response.write Server.HTMLEncode(FormatDateTime(rsProcedures.Fields("UpdateDate").Value,"0") & "")
      else
        response.write "&nbsp;"
      end if
      response.write "</TD>"
      
    response.write "</TR>"

    iCounter = iCounter + 1
    rsProcedures.moveNext
  Loop

  Response.write "</TBODY>"
  response.write "</TABLE>"

  Call Table_End
  
  Call Write_Nav_Buttons(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  
else

  response.write "<LI><SPAN CLASS=NormalRed>There are no calibration procedures found that match your search critera.  Please adjust your criteria and try again.</SPAN></LI><BR><BR>"
 	response.write "<table><tr><td>"
	Call Nav_Border_Begin
  response.write "<A HREF=""/met-support-gold/admin/metcal_admin.asp"
  if request("lv") <> "" then
    response.write "?lv=" & request("lv")
  end if
  
  response.write """>"
  response.write "<SPAN CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;New Search&nbsp;&nbsp;</SPAN></A>"
	Call Nav_Border_End
  response.write "</td></TR></TABLE>"
 	response.write "<BR>"
  
end if

Call Disconnect_SiteWide

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

' --------------------------------------------------------------------------------------

Sub Write_Nav_Buttons(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  Dim strLevelOutput1
  Dim strLevelOutput2
  Dim strResults
  Dim strPreviousPage
  Dim strNextPage
  Dim strNewSearch

  if strLevel <> "" then
	strLevelOutput1 = "?lv=" & strLevel
	strLevelOutput2 = "&lv=" & strLevel
  end if

  strPreviousPage = "<a href=""metcal_Procedures.asp?CurrPage=" & iCurrPage - 1 & "&Keyword=" & strKeyword & "&Calibrator=" & strCalibrator & strLevelOutput2 & """>"
  strPreviousPage = strPreviousPage & "<SPAN CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</A></SPAN>&nbsp;&nbsp;"

  strNextPage = "<a href=""metcal_Procedures.asp?CurrPage=" & iCurrPage + 1 & "&Keyword=" & strKeyword & "&Calibrator=" & strCalibrator & strLevelOutput2 & """>"
  strNextPage = strNextPage & "<SPAN CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</A></SPAN>&nbsp;&nbsp;"

  strNewSearch = "<A HREF=""metcal_admin.asp" & strLevelOutput1 & """><SPAN CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;New Search&nbsp;&nbsp;</SPAN></A>"

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
      Call Nav_Border_END
	else
  		if cInt(iCurrPage) = cInt(iNumPages) then
        Call Nav_Border_Begin      
  			response.write "<table><tr><td>"
	  		response.write strPreviousPage
  			response.write "</td>"
    		Call Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  			response.write "<td>"
        response.write strNewSearch
        response.write "</td>"
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

End Sub

' --------------------------------------------------------------------------------------

Sub Write_Pages(iNumPages, iCurrPage, strKeyword, strCalibrator, strLevel)
  Dim iCounter

  response.write "<TD CLASS=SmallBold>"

  for iCounter = 1 to iNumPages
  	if iCounter = 26 then
	  	exit for
	  else

   	response.write "<A HREF=""metcal_Procedures.asp?currpage=" & iCounter & "&Keyword=" & strKeyword & "&Calibrator=" & strCalibrator

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
%>
