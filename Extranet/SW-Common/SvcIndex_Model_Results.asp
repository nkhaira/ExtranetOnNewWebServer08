<%@ Language="VBScript" CODEPAGE="65001" %>

<%
response.buffer = true

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->
<%

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Variable Declarations
' --------------------------------------------------------------------------------------
' Variables to handle recordset pages
Dim iCurrPage
Dim iNumPages

' Variables to handle navigation issues
Dim BackURL
Dim LimitView
Dim ErrorString
Dim strQueryString
Dim strTargetPage
Dim strTargetNewSearch

' Variables to handle handle user search values
Dim strModel
Dim iMaxRows
Dim strView

' --------------------------------------------------------------------------------------
' Initialize variables from values in the request object
' --------------------------------------------------------------------------------------
iCurrPage = request("CurrPage")
iMaxRows = request("Rows")
strModel = request("model")
strView = request("view")

' --------------------------------------------------------------------------------------
' Get recordset of service doc listings
' --------------------------------------------------------------------------------------
set rsResults = GetResults(strModel)

' --------------------------------------------------------------------------------------
' Figure out how many pages there are in this recordset based upon 25 records per page
' --------------------------------------------------------------------------------------
if iMaxRows <> "" then
	iMaxRows = Cint(iMaxRows)
else
	iMaxRows = 10
end if

if IsNumeric(iCurrPage) then
	if iCurrPage = 0 then
		iCurrPage = 1
	end if
else
	iCurrPage = 1
end if

if not rsResults.eof then
	rsResults.PageSize = iMaxRows
	rsResults.AbsolutePage = iCurrPage

	iNumPages = rsResults.PageCount
end if

if iNumPages > 25 then
	iNumPages = 25
end if

' --------------------------------------------------------------------------------------
' Initialize Misc variables
' --------------------------------------------------------------------------------------
iCounter = 1
strTargetPage = request.servervariables("path_info")
strTargetNewSearch = "svcindex_model_form.asp"

BackURL = Session("BackURL")    

if lcase(request("lv")) = "false" then      ' Inital or Reset
  Session("LimitView")  = CInt(False)
elseif lcase(request("lv")) = "true" then
  Session("LimitView")  = CInt(True)
else
  if isblank(Session("LimitView")) then     ' Continue Existing View or Default
    Session("LimitView") = CInt(True)
  end if  
end if

LimitView    = Session("LimitView")

' Variables to handle error string
if isblank(Session("ErrorString")) then
  ErrorString = ""
else
  ErrorString = Session("ErrorString")
  Session("ErrorString") = ""
end if

%>

<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->

<%
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

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Product Service Support Information",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Product Service Support Information",Login_Language,conn) & "</FONT>" 
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

response.write "<FONT CLASS=Heading3>" & Translate("Product Service Support Information",Login_Language,conn) & "</FONT><BR>"
response.write "<FONT CLASS=Heading4>" & Translate("Search Results",Login_Language,conn) & "</FONT>"
response.write "<BR><BR>"

response.write "<FONT CLASS=Medium>"

if not isblank("ErrorString") then
  response.write "<UL>"
  response.write "<FONT COLOR=""Red"">" & ErrorString & "</FONT>"
  response.write "</UL>"
  Session("ErrorString") = ""
end if

' --------------------------------------------------------------------------------------
' Begin Main Section
' --------------------------------------------------------------------------------------

if (not rsResults.EOF) then

  response.write "<UL>"
  response.write "<LI>"
  response.write Translate("Key",Login_Language,conn) & ":&nbsp;&nbsp;"
  response.write "<IMG SRC=""/images/balls/g_square.gif"">=" & Translate("Active",Login_Language,conn) & "&nbsp;&nbsp;"
  response.write "<IMG SRC=""/images/balls/r_square.gif"">=" & Translate("Retired",Login_Language,conn) & "&nbsp;&nbsp;"
  response.write "<IMG SRC=""/images/balls/b_square.gif"">=" & Translate("Expired",Login_Language,conn) & "&nbsp;&nbsp;"
  response.write "<FONT CLASS=NavLeftHighLight1>EPD</FONT>=" & Translate("End Production Date",Login_Language,conn) & "&nbsp;&nbsp;"
  response.write "<FONT CLASS=NavLeftHighLight1>SP&nbsp;</FONT>=" & Translate("Support Period",Login_Language,conn) & "&nbsp;&nbsp;"
  response.write "<FONT CLASS=NavLeftHighLight1>ESP</FONT>=" & Translate("End Support Period Date",Login_Language,conn)
  response.write "</LI>"
  response.write "<BR><BR>"
  response.write "<LI>" & Translate("Click on Model to View Service Support Documentation",Login_Language,conn) & "</LI>"
  response.write "</UL>"


' --------------------------------------------------------------------------------------
' Write the navigation buttons if there are multiple pages of results
' --------------------------------------------------------------------------------------
  if iNumPages > 1 then
	strQueryString = "Rows=" & iMaxRows & "&model=" & strModel
	Write_Nav_Buttons iNumPages, iCurrPage, strLevel, strQuerystring, strTargetPage, strTargetNewSearch
  end if

' --------------------------------------------------------------------------------------
' Write header row
' --------------------------------------------------------------------------------------
	%>
  <TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
    <TR>
      <TD>
      
        <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
          <TR>

        	  <TH NOWRAP ALIGN="left" BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Model",Login_Language,conn)%></TH>
          	<TH NOWRAP ALIGN="left" BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Model Type",Login_Language,conn)%></TH>
          	<TH NOWRAP ALIGN="Middle" BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Warranty",Login_Language,conn)%></TH>
          	<TH NOWRAP ALIGN="Middle" BGCOLOR="#000000" Class=SmallBoldGold>A</TH>
          	<TH NOWRAP ALIGN="middle" BGCOLOR="#000000" Class=SmallBoldGold>R</TH>
          	<TH NOWRAP ALIGN="middle" BGCOLOR="#000000" Class=SmallBoldGold>X</TH>
          	<TH NOWRAP ALIGN="Middle" BGCOLOR="#000000" Class=SmallBoldGold>EPD</TH>
          	<TH NOWRAP ALIGN="Middle" BGCOLOR="#000000" Class=SmallBoldGold>SP</TH>
          	<TH NOWRAP ALIGN="Middle" BGCOLOR="#000000" Class=SmallBoldGold>ESP</TH>
	
	<%
' --------------------------------------------------------------------------------------
' Start looping through the result set and output information
' --------------------------------------------------------------------------------------
	while (not rsResults.EOF) and (iCounter <= iMaxRows)
		strModel = rsResults.Fields("Model") & ""
		strModelType = rsResults.Fields("Model_Type") & ""
		strWarranty = rsResults.Fields("Warranty") & ""
		dEndProdDate = rsResults.Fields("End_Production_Date") & ""
		strSupportPeriod = rsResults.Fields("Support_Period") & ""

		response.write "<tr>"
		if strModel <> "" then %>
			<TD Class=Small ALIGN="left" BGCOLOR="#FFFFFF"><A HREF="/sw-common/SvcIndex_Results.asp?view=<%=request("VIEW")%>&Model=<%=strModel%>&Rows=<%=limit%>&Highlight=<%=Session("Highlight")%>"><FONT SIZE=1 FACE="Arial"><%=strModel%></FONT></A></TD>
		<%else%>
			<TD Class=Small ALIGN="left" BGCOLOR="#EEEEEE">&nbsp;</TD>
		<%end if%>

		<%if strModelType <> "" then %>
			<TD Class=Small ALIGN="left" BGCOLOR="#FFFFFF"><%=strModelType%></TD>
		<%else%>
			<TD Class=Small ALIGN="left" BGCOLOR="#EEEEEE">&nbsp;</TD>
		<%end if%>
		
		<%if strWarranty <> "" then %>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF"><%=strWarranty%>&nbsp;<%=Translate("years",Login_Language,conn)%></TD>
		<%else%>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#EEEEEE">&nbsp;</TD>
		<%end if%>
		
		<%if dEndProdDate = "" then%>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF"><IMG SRC="/images/balls/g_square.gif"></TD>
		<%else%>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF">-</TD>
		<%end if%>
		
		<%if dEndProdDate <> "" and strSupportPeriod <> "" then %>
			<%
			yr = mid(Cstr(dEndProdDate),len(Cstr(dEndProdDate))-1,2)
			if len(yr) = 2 then
			        if Cint(yr) < 50 then
				         yr = 2000 + Cint(yr)
			        else
	  				yr = 1900 + Cint(yr)
			        end if
			end if

			strMonth = mid(Cstr(dEndProdDate),1,2)
			if mid(strMonth,2,1)="/" then
				strMonth= "0" & mid(strMonth,1,1)
			end if
			addto = Cint(strSupportPeriod)
			present = mid(Cstr(date()),len(Cstr(date()))-1,2)
			if len(Cstr(present)) = 2 then
				if Cint(present) < 50 then        
			          present = 2000 + Cint(present)        
			        else          
			          present = 1900 + Cint(present)
			        end if  
			end if
			cutdate = addto + Cint(yr)
			if cutdate > Cint(present) then %>
				<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF"><IMG SRC="/images/balls/r_square.gif"></TD>
			<%else%>
				<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF">-</TD>
			<%end if%>
		<%else%>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF">-</TD>
		<%end if%>
		
		<%if dEndProdDate <> "" and strSupportPeriod <> "" then %>			
			<%if cutdate < Cint(present) then %>
				<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF"><IMG SRC="/images/balls/b_square.gif"></TD>
			<%else%>
				<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF">-</TD>
			<%end if%>
		<%else%>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF">-</TD>
		<%end if%>
		
		<%if dEndProdDate <> "" then %>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF"><%=PrintFileDate(1, dEndProdDate)%></TD>
		<%else%>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#EEEEEE">&nbsp;</TD>
		<%end if%>
	
		<%if strSupportPeriod <> "" then %>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF"><%=strSupportPeriod%>&nbsp;<%=Translate("years",Login_Language,conn)%></TD>
		<%else%>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#EEEEEE">&nbsp;</TD>
		<%end if%>

		<%if dEndProdDate <> "" and strSupportPeriod <> "" then %>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#FFFFFF"><%response.write strMonth & "/" & cutdate%></TD>
		<%else%>
			<TD Class=Small ALIGN="Middle" BGCOLOR="#EEEEEE">&nbsp;</TD>
		<%end if%>

		<%
		Response.write "</tr>"
		rsResults.moveNext
		iCounter = iCounter + 1
	Wend
	response.write "</table></TD></TR></TABLE>"
	
else
  %>
	<UL><B><%=Translate("Sorry, No records match the search criteria you have entered.",Login_Language,conn)%></B></UL>
	<UL><B><%=Translate("Click on [New Search] to enter new search criteria.",Login_Language,conn)%></B></UL>	
	<%
end if

if iNumPages > 1 then
	strQueryString = "Rows=" & iMaxRows & "&model=" & strModel
	Write_Nav_Buttons iNumPages, iCurrPage, strLevel, strQuerystring, strTargetPage, strTargetNewSearch
end if

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' Functions
' --------------------------------------------------------------------------------------

Function GetResults(strModel)
	Dim rsResults
	Dim cmd, prm
	Dim bLikeSearch
	Dim strModel_Clean

	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "SVC_Models_SearchModels"

	strModel_Clean = changecard(strModel)
	if InStr(strModel_Clean,"%") then
		bLikeSearch = true
	else
		bLikeSearch = false
	end if

'response.write("bLikeSearch: " & bLikeSearch & "<BR>")
'response.write("strModel: " & strModel & "<BR>")

	Set prm = cmd.CreateParameter("@strModel", adVarchar, adParamInput, 50, strModel_Clean)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@bLikeSearch", adInteger, adParamInput, , bLikeSearch)
	cmd.Parameters.Append prm

	Set rsResults = Server.CreateObject("ADODB.Recordset")
	rsResults.CursorLocation = adUseClient
	rsResults.CursorType = adOpenDynamic
	rsResults.open cmd

	set prm = nothing
	set cmd = nothing

	set GetResults = rsResults

End Function

Sub Write_Nav_Buttons(iNumPages, iCurrPage, strLevel, strQuerystring, strTargetPage, strTargetNewSearch)
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

  strPreviousPage = "<a href=""" & strTargetPage & "?CurrPage=" & iCurrPage - 1 & "&" & strQueryString & strLevelOutput2 & """>"
  strPreviousPage = strPreviousPage & "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</A></FONT>&nbsp;&nbsp;"

  strNextPage = "<a href=""" & strTargetPage & "?CurrPage=" & iCurrPage + 1 & "&" & strQueryString & strLevelOutput2 & """>"
  strNextPage = strNextPage & "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</A></FONT>&nbsp;&nbsp;"

  strNewSearch = "<A HREF=""" & strTargetNewSearch & strLevelOutput1 & """><FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;New Search&nbsp;&nbsp;</FONT></A>"

  if iNumPages > 1 then
  	response.write "<BR>"
  
'response.write("iCurrPage: " & iCurrPage & "<BR>")
'response.write("iNumPages: " & iNumPages & "<BR>")

  	if cInt(iCurrPage) = 1 then
		response.write "<table><tr>"
  		Call Write_Pages(iNumPages, iCurrPage, strLevel, strQueryString, strTargetPage)
  		response.write "<td>"
		response.write strNextPage
		response.write strNewSearch
  		response.write "</td></tr></table>"
	else
  		if cInt(iCurrPage) = cInt(iNumPages) then
  			response.write "<table><tr><td>"
			response.write strPreviousPage
  			response.write "</td>"
	  		Call Write_Pages(iNumPages, iCurrPage, strLevel, strQueryString, strTargetPage)
			response.write "<td>" & strNewSearch & "</td>"
  			response.write "</tr></table>"
  		else
  			response.write "<table><tr><td>"
			response.write strPreviousPage
  			response.write "</td>"
	  		Call Write_Pages(iNumPages, iCurrPage, strLevel, strQueryString, strTargetPage)
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

Sub Write_Pages(iNumPages, iCurrPage, strLevel, strQueryString, strTargetPage)
  Dim iCounter

  response.write "<TD CLASS=MediumBold>"
  
  for iCounter = 1 to iNumPages
	if iCounter = 26 then
	  	exit for
	else
	  	response.write "<A HREF=""" & strTargetPage & "?currpage=" & iCounter & "&" & strQueryString

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

' --------------------------------------------------------------------------------------  

Function PrintFileDate(iDateFormat, TempDate)

  if Len(DatePart("d", TempDate)) = 1 then
  	strDay = "0" & DatePart("d", TempDate)
  else
  	strDay = DatePart("d", TempDate)
  end if
  
  if Len(DatePart("m", TempDate)) = 1 then
  	strMonth = "0" & DatePart("m", TempDate)
  else
  	strMonth = DatePart("m", TempDate)
  end if
  
  if iDateFormat = 1 then
  	PrintFileDate = strMonth & "/" & strDay & "/" & DatePart("yyyy", TempDate)
  elseif iDateFormat = 2 then
  	PrintFileDate = strDay & "/" & strMonth & "/" & DatePart("yyyy", TempDate)
  else
  	PrintFileDate = DatePart("yyyy", TempDate) & "/" & strMonth & "/" & strDay
  end if

end function
	
Function changecard(str)
	dim i
	dim tempstr
	for i = 1 to len(str)
		if Mid(str,i,1) = "*" then
			tempstr=tempstr & "%"
			if i <> 1 then
				exit for	
			end if
			if i = len(str)-1 then
				exit for
			end if
		else if Mid(str,i,1) = "'" then
			tempstr=tempstr
		else
			tempstr=tempstr & mid(str,i,1)
		end if
		end if
	next
	changecard = tempstr
end function

%>