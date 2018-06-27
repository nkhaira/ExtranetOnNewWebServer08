<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
response.buffer = true

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Variable Declarations
' --------------------------------------------------------------------------------------
' Misc vars
Dim BackURL
Dim bLimitView
Dim strErrorString
Dim DOC_PATH
Dim strQueryString
Dim strTargetPage
Dim strTargetPage_New
Dim iCounter
Dim strLastModel
Dim strFilter

' Variables to handle recordset pages
Dim iMaxRows
Dim iCurrPage
Dim iNumPages

' Variables to hold user-chosen search criteria
Dim strDoc_Type
Dim iDoc_Num_Min
Dim iDoc_Num_Max
Dim strModel
Dim rsResults
Dim dDate_Month
Dim dDate_Day
Dim dDate_Year
Dim dDate
Dim iSort

' --------------------------------------------------------------------------------------
' Initialize variables from values in the request object
' --------------------------------------------------------------------------------------
iCurrPage = request("CurrPage")
iMaxRows = request("Rows")
strDoc_Type = uCase(request("Doc_Type")) & ""
iDoc_Num_Min = request("Doc_Num_Min")
iDoc_Num_Max = request("Doc_Num_Max")
strModel = uCase(request("model")) & ""
dDate_Month = request("Date_Month")
dDate_Day = request("Date_Day")
dDate_Year = request("Date_Year")
dDate = request("date")
iSort = request("sort")

if iMaxRows = "" or not IsNumeric(iMaxRows) then
	iMaxRows = 25
end if
if iSort = "" or not IsNumeric(iSort) then
	iSort = 1
end if

' --------------------------------------------------------------------------------------
' Initialize misc variables
' --------------------------------------------------------------------------------------
iCounter = 1

BackURL = Session("BackURL")    

DOC_PATH = "/Service-Center/Download/Documents/"

strFilter = "dummy All Q UD SSU SD MSU MET SPF IM ESU DTE CIS USU GENERAL OBSOLETE"

if lcase(request("lv")) = "false" then      ' Inital or Reset
  strLimitView  = CInt(False)
else
  strLimitView  = CInt(True)
end if

if dDate = "" or not IsDate(dDate) then
	if dDate_Month <> "" and dDate_Day <> "" and dDate_Year <> "" then
		if IsNumeric(dDate_Month) and IsNumeric(dDate_Day) and IsNumeric(dDate_Year) then
			dDate = cDate(dDate_Month & "/" & dDate_Day & "/" & dDate_Year)
		end if
	end if
end if

strQueryString = "Rows=" & iMaxRows & "&Doc_Type=" & strDoc_Type
strQueryString = strQueryString & "&Doc_Num_Min=" & iDoc_Num_Min & "&Doc_Num_Max=" & iDoc_Num_Max
strQueryString = strQueryString & "&model=" & strModel & "&date=" & cStr(dDate) & "&sort=" & iSort

strTargetPage = request.servervariables("path_info")
strTargetPage_New = "svcindex_form.asp"

' --------------------------------------------------------------------------------------
' Get recordset of service doc listings
' --------------------------------------------------------------------------------------
set rsResults = GetResults(strDoc_Type, iDoc_Num_Min, iDoc_Num_Max, strModel, dDate, iSort)

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

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Service Documents",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Service Documents",Login_Language,conn) & "</FONT>" 
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

response.write "<FONT CLASS=Heading3>" & Translate("Service Documents",Login_Language,conn) & "</FONT><BR>"
response.write "<FONT CLASS=Heading4>" & Translate("Search Results",Login_Language,conn) & "</FONT>"
response.write "<BR><BR>"

response.write "<FONT CLASS=Medium>"

if not strErrorString = "" then
  response.write "<UL>"
  response.write "<FONT COLOR=""Red"">" & strErrorString & "</FONT>"
  response.write "</UL>"
  strErrorString = ""
end if

' --------------------------------------------------------------------------------------
' Begin Main section
' --------------------------------------------------------------------------------------
%>
<UL>
<% if request("view") <> "2" then %>
<LI><%=Translate("Click on the <B>Model</B> link to view Service Support Information related to that model number or click on the <B>Doc Number</B> link to view or download the document.",Login_Language,conn)%></LI>
<% else %>
<LI><%=Translate("Click on the <B>Doc Number</B> link to view or download the document.",Login_Language,conn)%></LI>
<% end if %>
<LI><%=Translate("Documents that have additional support files are indicated by an underlined link in the <B>Support File or Order Number</B> column.",Login_Language,conn)%></LI>
<LI><%=Translate("Documents that reference order numbers are indicated by a 6 or 7 digit order number appearing in the <B>Support File or Order Number</B> column.",Login_Language,conn)%></LI>
<LI><%=Translate("This document is best printed in <B>Landscape</B> orientation.",Login_Language,conn)%></LI>
</UL>

<%
if rsResults.EOF then

	response.write "<UL><B>" & Translate("Sorry, No records match the search criteria you have entered.",Login_Language,conn) & "</B></UL>"
	response.write "<UL><B>" & Translate("Click on <FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;New Search&nbsp;&nbsp;</FONT> to enter new search criteria.",Login_Language,conn) & "</B></UL>"

else

' --------------------------------------------------------------------------------------
' Write the navigation buttons if there are multiple pages of results
' --------------------------------------------------------------------------------------
Write_Nav_Buttons iNumPages, iCurrPage, strLevel, strQuerystring, strTargetPage, strTargetPage_New

' --------------------------------------------------------------------------------------
' Write header row
' --------------------------------------------------------------------------------------

%>
	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
    <TR>
      <TD>   
        <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
          <TR>
            <TD NOWRAP ALIGN="left"   BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Model",Login_Language,conn)%></TD>
          	<TD NOWRAP ALIGN="left"   BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Assembly",Login_Language,conn)%></TD>
          	<TD NOWRAP ALIGN="left"   BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Board<BR>Revision",Login_Language,conn)%></TD>
          	<TD NOWRAP ALIGN="left"   BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Serial #",Login_Language,conn)%></TD>
          	<TD NOWRAP ALIGN="left"   BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Doc<BR>Number",Login_Language,conn)%></TD>
          	<TD NOWRAP ALIGN="left"   BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Order #<BR>File",Login_Language,conn)%></TD>  
          	<TD NOWRAP ALIGN="middle" BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Doc<BR>Revision",Login_Language,conn)%></TD>
          	<TD NOWRAP ALIGN="left"   BGCOLOR="#000000" Class=SmallBoldGold><%=Translate("Description",Login_Language,conn)%></TD>
<%


  strLastModel = ""

' --------------------------------------------------------------------------------------
' Start looping through the result set and output information
' --------------------------------------------------------------------------------------

  do while not rsResults.EOF and iCounter <= iMaxRows
    
    ' --------------------------------------------------------------------------------------
    ' We display information grouped by model so if this is a new model, make a row to show this new model
    ' --------------------------------------------------------------------------------------
    if rsResults.Fields("Model") <> strLastModel then
	response.write("<TR>")

        if rsResults.Fields("Model_Type") <> "" then 
%>
          <TD BGCOLOR="#666666" CLASS=SmallBoldGold>
            <%=rsResults.Fields("Model")%>
          </TD>           
          <TD COLSPAN=9 Class=SmallBoldGold>
            <%=rsResults.Fields("Model_Type")%>
          </TD>

     <% else %>

          <TD COLSPAN=10 BGCOLOR="#666666" CLASS=MediumBoldGold>
            <%=rsResults.Fields("Model")%>
          </TD>           

     <% end if%>                                
      </TR>
      <%
    end if
    strLastModel = rsResults.Fields("Model")    

    ' --------------------------------------------------------------------------------------
    ' Write model information
    ' --------------------------------------------------------------------------------------
    response.write "<tr>"
    if rsResults.Fields("Model") <> "" then 
	response.write("<TD ALIGN=""left"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=Small>")
	if request("view") <> "2" then
		response.write("<A HREF=""svcindex_model_results.asp?view=" & request("VIEW") & "&Model=" & rsResults.Fields("MODEL") & "&Returns=" & limit & "&Highlight=" & Highlight & """>" & rsResults.Fields("Model") & "</A>")
	else	
		response.write rsResults.Fields("Model")
	end if
	response.write("</TD>")
    else
	response.write("<TD ALIGN=""left"" BGCOLOR=""#F8F8F8"" Class=Small>&nbsp;</TD>")
    end if

    ' --------------------------------------------------------------------------------------
    ' Write assembly information
    ' --------------------------------------------------------------------------------------
    if rsResults.Fields("Assembly") <> "" then
	response.write("<TD ALIGN=""left"" BGCOLOR=""#FFFFFF"" Class=Small>" & rsResults.Fields("Assembly") & "</TD>")
    else
	response.write("<TD ALIGN=""left"" BGCOLOR=""#F8F8F8"" Class=Small>&nbsp;</TD>")
    end if
		
    ' --------------------------------------------------------------------------------------
    ' Write Board Revision information
    ' --------------------------------------------------------------------------------------
    if rsResults.Fields("Board_Revision") <> "" then
	response.write("<TD ALIGN=""left"" BGCOLOR=""#FFFFFF"" Class=Small>" & rsResults.Fields("Board_Revision") & "</TD>")
    else
	response.write("<TD ALIGN=""left"" BGCOLOR=""#F8F8F8"" Class=Small>&nbsp;</TD>")
    end if
		
    ' --------------------------------------------------------------------------------------
    ' Write Serial Number information
    ' --------------------------------------------------------------------------------------
    if rsResults.Fields("Serial_Num") <> "" then
	response.write("<TD ALIGN=""left"" BGCOLOR=""#FFFFFF"" Class=Small>" & rsResults.Fields("Serial_Num") & "</TD>")
    else
	response.write("<TD ALIGN=""left"" BGCOLOR=""#F8F8F8"" Class=Small>&nbsp;</TD>")
    end if

    ' --------------------------------------------------------------------------------------
    ' Write Doc Number information
    ' --------------------------------------------------------------------------------------
    if rsResults.Fields("Doc_Num") <> "" then
	if InStr(strFilter,rsResults.Fields("Doc_Type")) > 0 then
		response.write("<TD ALIGN=""left"" BGCOLOR=""#FFFFFF"" Class=Small>" & rsResults.Fields("Doc_Num") & "</TD>")
	else
		path = makepath(rsResults.Fields, DOC_PATH)
		' new functionality to use .pdf files if exist instead of old .tif files
		tempPath = request.servervariables("APPL_PHYSICAL_PATH") & reverseslash(mid(path & ".pdf", 2))

		Set fso = CreateObject("Scripting.FileSystemObject")

		If (fso.FileExists(tempPath)) Then
			path = path & ".pdf"
		Else
			path = path & ".tif"
		End If
		response.write("<TD ALIGN=""left"" BGCOLOR=""#FFFFFF"" NOWRAP Class=Small><A HREF=""" & path & """>" & rsResults.Fields("Doc_Type") & "-" & rsResults.Fields("Doc_Num") & "</A></TD>")
	end if
    else
	response.write("<TD ALIGN=""left"" BGCOLOR=""#FFFFFF"" Class=Small>" & rsResults.Fields("Doc_Type") & "</TD>")
    end if
		
    ' --------------------------------------------------------------------------------------
    ' Write Order Number File information
    ' --------------------------------------------------------------------------------------
    if rsResults.Fields("Order_Code") <> "" then
	if rsResults.Fields("Doc_Type") <> "MET" and rsResults.Fields("Doc_Type") <> "SPF" then
		response.write("<TD ALIGN=""left"" BGCOLOR=""#FFFFFF"" Class=Small>")
          	order_code = rsResults.Fields("Order_Code")
          	if isnumeric(order_code) then
			response.write("<A HREF=""/SW-Common/Part_Query.asp?view=2&part=" & Order_Code & "&whatpage=1"" TARGET=""codes"" onclick=""openit('/SW-Common/Part_Query.asp?view=2&part=" & Order_Code & "&whatpage=1','Vertical');return false;"">" & Order_Code & "</A>")
          	else
              		response.write order_code
             	end if  
        	response.write("</TD>")
      	else
  		if rsResults.Fields("Doc_Type") = "MET" then       
    			path = "/Service-Center/Download/documents/Service_MetCal/" & rsResults.Fields("Order_Code")
    	  	elseif rsResults.Fields("Doc_Type") = "SPF" then
	    	  	path = "/Service-Center/Download/documents/Support_Files/" & rsResults.Fields("Order_Code")
          	end if
		response.write("<TD ALIGN=""left"" BGCOLOR=""#FFFFFF"" NOWRAP Class=Small><A HREF=""" & path & """>" & rsResults.Fields("Order_Code") & "</A></TD>")
  	end if
    else
		response.write("<TD ALIGN=""left"" BGCOLOR=""#F8F8F8"" Class=Small>&nbsp;</TD>")
    end if

    ' --------------------------------------------------------------------------------------
    ' Write Document Revision information
    ' --------------------------------------------------------------------------------------
    if rsResults.Fields("Doc_Rev") <> "" then
	response.write("<TD ALIGN=""Center"" BGCOLOR=""#FFFFFF"" Class=Small>" & PrintFileDate(1, rsResults.Fields("Doc_Rev")) & "</TD>")
    else
	if rsResults.Fields("Doc_Date") <> "" then
		response.write("<TD ALIGN=""Center"" BGCOLOR=""#FFFFFF"" Class=Small>" & PrintFileDate(1, rsResults.Fields("Doc_Date")) & "</TD>")
	else
		response.write("<TD ALIGN=""Center"" BGCOLOR=""#F8F8F8"" Class=Small>&nbsp;</TD>")
	end if
    end if

    ' --------------------------------------------------------------------------------------
    ' Write Description information
    ' --------------------------------------------------------------------------------------
    if rsResults.Fields("Description") <> "" then
	response.write("<TD ALIGN=""left"" BGCOLOR=""#FFFFFF"" Class=Small>")
	response.write rsResults.Fields("Description")
	response.write("</TD>")
    else
	response.write("<TD ALIGN=""left"" BGCOLOR=""#F8F8F8"" Class=Small>&nbsp;</TD>")
    end if

    Response.write "</tr>"

    rsResults.MoveNext

    iCounter = iCounter + 1
  loop
	
  response.write "</TABLE>"  
  response.write "</TD></TR></TABLE>"

end if

Write_Nav_Buttons iNumPages, iCurrPage, strLevel, strQuerystring, strTargetPage, strTargetPage_New

response.write "</FONT>"
response.write "<BR><BR>"

%>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' Functions
' --------------------------------------------------------------------------------------

Function GetResults(strDoc_Type, iDoc_Num_Min, iDoc_Num_Max, strModel, dDate, iSort)
	Dim strSQL
	Dim strFilter
	Dim rsResults
	Dim cmd, prm
	Dim bLikeSearch

	strFilter = "dummy All Q UD SSU SD MSU MET SPF IM ESU DTE CIS USU GENERAL OBSOLETE"

	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "SVC_Master_SearchDocs"

	if iDoc_Num_Min = "" or not IsNumeric(iDoc_Num_Min) or InStr(strFilter, strDoc_Type) <> 0 then
		iDoc_Num_Min = -1
	end if
	if iDoc_Num_Max = "" or not IsNumeric(iDoc_Num_Min) or InStr(strFilter, strDoc_Type) <> 0 then
		iDoc_Num_Max = -1
	end if

	if dDate = "" then
		dDate = ""
	end if
	
	if strDoc_Type <> "ALL" and strDoc_Type <> "" then
		if strDoc_Type = "OBSOLETE" then
			strModel = "OBSOLETE"
		elseif strDoc_Type = "GENERAL" then
			strModel = "GENERAL-INFORMATION"
		end if
	else
		strDoc_Type = ""
	end if

	strModel = changecard(strModel)
	if InStr(strModel,"%") then
		bLikeSearch = true
	else
		bLikeSearch = false
	end if

'response.write("strDoc_Type: " & strDoc_Type & "<BR>")
'response.write("iDoc_Num_MIN: " & iDoc_Num_Min & "<BR>")
'response.write("iDoc_Num_Max: " & iDoc_Num_Max & "<BR>")
'response.write("strModel: " & strModel & "<BR>")
'response.write("bLikeSearch: " & bLikeSearch & "<BR>")
'response.write("bit value of bLikeSearch: " & cInt(bLikeSearch) & "<BR>")
'response.write("dDate: " & dDate & "<BR>")
'response.write("iSort: " & iSort & "<BR>")

	Set prm = cmd.CreateParameter("@strDocType", adVarChar, adParamInput, 50, strDoc_Type & "")
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@iDoc_Num_Min", adInteger, adParamInput, , iDoc_Num_Min)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@iDoc_Num_Max", adInteger, adParamInput, , iDoc_Num_Max)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@strModel", adVarchar, adParamInput, 50, strModel)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@bLikeSearch", adBoolean, adParamInput, , bLikeSearch)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@dDate", adVarChar, adParamInput, 10, dDate)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@iSort", adInteger, adParamInput, , iSort)
	cmd.Parameters.Append prm

	Set rsResults = Server.CreateObject("ADODB.Recordset")
	rsResults.CursorLocation = adUseClient
	rsResults.CursorType = adOpenDynamic
	rsResults.open cmd

	set prm = nothing
	set cmd = nothing

	set GetResults = rsResults

End Function

' --------------------------------------------------------------------------------------  

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
    else
      if Mid(str,i,1) = "'" then
    		tempstr=tempstr
     	else
     		tempstr=tempstr & mid(str,i,1)
     	end if
    end if
  next
  changecard = tempstr
end function

' --------------------------------------------------------------------------------------  

Sub Write_Nav_Buttons(iNumPages, iCurrPage, strLevel, strQuerystring, strTargetPage, strTargetPage_New)
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

  strNewSearch = "<A HREF=""" & strTargetPage_New & strLevelOutput1 & """><FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;New Search&nbsp;&nbsp;</FONT></A>"

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

Function ReverseSlash(strToParse)

for i = 1 to Len(strToParse)
	if Mid(strToParse, i, 1) = "/" then
		strReturnValue = strReturnValue & "\" %><%
	elseif Mid(strToParse, i, 1) = "\" then %><%
		strReturnValue = strReturnValue & "/"
	else
		strReturnValue = strReturnValue & Mid(strToParse, i, 1)
	end if
next
ReverseSlash = strReturnValue
End Function

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

' --------------------------------------------------------------------------------------    

Function makepath(queryobject, DOC_PATH)
  dim doc
  dim kind
  dim str
  dim padded

  str = DOC_PATH
  doc = trim(queryobject("Doc_Num"))
  kind = trim(queryobject("Doc_Type"))

  Select Case trim(cstr(kind))
  	case "CIH"
  		str = str & "CIH00000/"
  		padded = pad(doc,kind)
  		str = str & padded
  	case "OSC"
  		str = str & "OSC00000/"
  		padded = pad(doc,kind)
  		str = str & padded
  	case "SBU"
  		str = str & "SBU00000/"
  		padded = pad(doc,kind)
  		str = str & padded
  	case "SI"
  		str = str & "SI000000/"
  		padded = pad(doc,kind)
  		str = str & padded
  	case "SME"
  		str = str & "SME00000/"
  		padded = pad(doc,kind)
  		str = str & padded
  	case "SPC"
  		str = str & "SPC00000/"
  		padded = pad(doc,kind)
  		str = str & padded
  	case "SRE"
  		str = str & "SRE00000/"
  		padded = pad(doc,kind)
  		str = str & padded
  	case "SSY"
  		str = str & "SSY00000/"
  		padded = pad(doc,kind)
  		str = str & padded
  	case "PCN"
  		if cint(doc) < 1000 then str = str & "PCN00000/"
  		if cint(doc) < 2000 and cint(doc) > 1000 then str = str & "PCN01000/"
  		if cint(doc) < 3000 and cint(doc) > 2000 then str = str & "PCN02000/"	
  		if cint(doc) < 4000 and cint(doc) > 3000 then str = str & "PCN03000/"
  		if cint(doc) < 5000 and cint(doc) > 4000 then str = str & "PCN04000/"
  		if cint(doc) < 6000 and cint(doc) > 5000 then str = str & "PCN05000/"
  		if cint(doc) < 7000 and cint(doc) > 6000 then str = str & "PCN06000/"
  		if cint(doc) < 8000 and cint(doc) > 7000 then str = str & "PCN07000/"
  		if cint(doc) < 9000 and cint(doc) > 8000 then str = str & "PCN08000/"
  		if cint(doc) < 10000 and cint(doc) > 9000 then str = str & "PCN09000/"
  		padded = pad(doc,kind)
  		str = str & padded
  	case "SA"
  		if cint(doc) < 1000 then str = str & "SA000000/"
  		if cint(doc) < 2000 and cint(doc) > 1000 then str = str & "SA001000/"
  		if cint(doc) < 3000 and cint(doc) > 2000 then str = str & "SA002000/"	
  		if cint(doc) < 4000 and cint(doc) > 3000 then str = str & "SA003000/"
  		if cint(doc) < 5000 and cint(doc) > 4000 then str = str & "SA004000/"
  		if cint(doc) < 6000 and cint(doc) > 5000 then str = str & "SA005000/"
  		if cint(doc) < 7000 and cint(doc) > 6000 then str = str & "SA006000/"
  		if cint(doc) < 8000 and cint(doc) > 7000 then str = str & "SA007000/"
  		if cint(doc) < 9000 and cint(doc) > 8000 then str = str & "SA008000/"
  		if cint(doc) < 10000 and cint(doc) > 9000 then str = str & "SA009000/"
  		padded = pad(doc,kind)
  		str = str & padded
  	case "SH", "KIT"
  		str = str & "SH000000/"
        if cstr(kind)= "KIT" then
          padded = pad(doc,"SH")
        else          
  				padded = pad(doc,kind)          
        end if          
  		str = str & padded
  end select
  makepath = str
end Function

' --------------------------------------------------------------------------------------    

function pad(topad,tokind)
  dim extra
  dim i
  extra = 8 - len(tokind) - len(topad)
  pad = tokind
  if extra > 0 then
  	for i = 1 to extra
  		pad = pad & "0"
    next
  end if
  pad = pad & topad
end function

' --------------------------------------------------------------------------------------    
%>

<SCRIPT LANGUAGE=JAVASCRIPT>

<!--

function Checkitout(){

//      Gets Browser and Version

        var appver = "null";
        var browser = navigator.appName;
        var version = navigator.appVersion;
        if ((browser == "Netscape")) version = navigator.appVersion.substring(0, 3);
        if ((browser == "Microsoft Internet Explorer")) version = navigator.appVersion.substring(22, 25);

//      Gives AppVersion (appver) for Detect Strings

        if ((browser == "Microsoft Internet Explorer") && (version >= 3)) appver = "ie3+";
        if ((browser == "Netscape") && (version >= 3)) appver = "ns3+";
        if ((browser == "Netscape") && (version < 3)) appver = "ns2";


       if ((appver == "ie3+")) {
                return 0;
        }  else {
                return 1;
                }
}

function PopoffWindow(DaURL, orient) {

	var ItsTheWindow;
	if (Checkitout())  {
		if (orient == "Horizontal")  {
			ItsTheWindow = window.open(DaURL,"himom","status,height=400,width=400,scrollbars=yes,resizable=no,toolbar=0");
		} else if (orient == "Vertical")  {
		    ItsTheWindow = window.open(DaURL,"himom","status,height=400,width=400,scrollbars=yes,resizable=no,toolbar=0");
		}


	} else {
		if (orient == "Horizontal")  {
	        ItsTheWindow = window.open(DaURL,"himom","scrollbars=yes,menubar=no,toolbar=no,links=no,status=no,height=400,width=400,resizable=no");
		} else if (orient == "Vertical")  {
	        ItsTheWindow = window.open(DaURL,"himom","scrollbars=yes,menubar=no,toolbar=no,links=no,status=no,height=400,width=400,resizable=no");
		}
			if (parseInt(navigator.appVersion) >= 3){
       		ItsTheWindow.focus();
        }
	}

}
function openit(DaURL, orient) {
		var ItsTheWindow;
        if (Checkitout())  {
                if (orient == "Horizontal")  {
                        ItsTheWindow = window.open(DaURL,"codes","status,height=400,width=600,scrollbars=1,resizable=1,toolbar=0");
                } else if (orient == "Vertical")  {
                    ItsTheWindow = window.open(DaURL,"codes","status,height=580,width=545,scrollbars=1,resizable=1,toolbar=0");
                }


        } else {
                if (orient == "Horizontal")  {
                ItsTheWindow = window.open(DaURL,"codes","scrollbars=1,menubar=0,toolbar=0,links=0,status=1,height=400,width=600,resizable=1");
                } else if (orient == "Vertical")  {
                ItsTheWindow = window.open(DaURL,"codes","scrollbars=1,menubar=0,toolbar=0,links=0,status=1,height=580,width=545,resizable=1");
                }
                        if (parseInt(navigator.appVersion) >= 3){
                ItsTheWindow.focus();
        }
        }
}

//-->

</SCRIPT>
