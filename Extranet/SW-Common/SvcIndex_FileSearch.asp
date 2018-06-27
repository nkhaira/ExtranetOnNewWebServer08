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
' Declarations
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->
<%

Dim BackURL
Dim ErrorString

BackURL = Session("BackURL")    

if isblank(Session("ErrorString")) then
  ErrorString = ""
else
  ErrorString = Session("ErrorString")
  Session("ErrorString") = ""
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
response.write "<FONT CLASS=Heading4>" & Translate("What&acute;s New",Login_Language,conn) & "</FONT><BR><BR>"
response.write "<FONT CLASS=MediumBold>" & Translate("Search Results within the last",Login_Language,conn) & " " & Request.Form("DP") & " " & Translate("days",Login_Language,conn) & ".</FONT>"
response.write "<BR><BR>"

response.write "<FONT CLASS=Medium>"

response.write "<A HREF=""/sw-common/SvcIndex_Whatsnew_Form.asp""><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A><BR><BR>"
if not isblank("ErrorString") then
  response.write "<UL>"
  response.write "<FONT COLOR=""Red"">" & ErrorString & "</FONT>"
  response.write "</UL>"
  Session("ErrorString") = ""
end if

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

Server.ScriptTimeOut = 600
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
dim vPath
dim vDB
dim i
i=0
	
vPath = Split(Request.form("Path"),"&")

If Request.Form("PostTest") <> 1 then NonPostError

' Response.write ":" & VarType(vPath) & ":<br><br>" 
' Key for vPath:
' vPath(0) is path
' vPath(1) is title of result page
' vPath(2) is descend directories toggle
' vPath(3) is use database toggle (1=use database, 0=use tag file, 2=use nothing)
	
vDB = Split(vPath(3), "=")

MyServerPath = "/service-center/" & vPath(0)
wDir = Server.MapPath(MyServerPath)
  
if vDB(1) = 0 then
  Set fsObject = CreateObject("Scripting.FileSystemObject")

	Set tfIn = fsObject.OpenTextFile(wDir & "\index.tag", 1, False)
	strTag = tfIn.ReadAll()
	tfIn.Close
	set tfIn = nothing
	set fsObject = nothing
end if
	

set oFolder = FSO.GetFolder(wDir)
set oFiles = oFolder.Files

for each oFile in oFiles

	fname = oFile.name
	if FName <> "" then
		strExt = Ucase(GetExtension(FName))
		if strExt <> "GIF" AND strExt <> "JPG" AND strExt <> "BMP" AND strExt <> "CGI" AND strExt <> "TAG" AND strExt <> "ASP" AND strExt <> "HTM" then 

		if CInt(DateDiff("d",oFile.DateLastModified,date)) < CInt(Request.form("DP")) AND i = 0 then
			' First file that matches, print header and first file
			PrintHeader
	
		elseif CInt(DateDiff("d",oFile.DateLastModified,date)) < CInt(Request.form("DP")) then

			Response.write "<TR><TD ALIGN=MIDDLE BGCOLOR=""White""><FONT CLASS=Small>"
			Response.write GetExtension(FName)
			Response.write "</FONT></TD><TD ALIGN=CENTER BGCOLOR=""White"">"
			Response.write PrintSquare(GetExtension(FName))
			Response.write "</TD><TD BGCOLOR=""White""><FONT CLASS=Small>"
			Response.write "<B>" & PrintHREF(FName) & PrintFileName(FName) & "</a></B>"
			Response.write PrintDesc(FName)
			Response.write "</FONT></TD>"
			Response.write "<TD ALIGN=RIGHT NOWRAP BGCOLOR=""White""><FONT CLASS=Small>"
			Response.write DateDiff("d",oFile.DateLastModified,date)
			Response.write "</FONT></TD><TD NOWRAP BGCOLOR=""White"" ALIGN=RIGHT><FONT CLASS=Small>"
			Response.write PrintFileDate(Request.form("DF"))
			Response.write "</FONT></TD><TD NOWRAP BGCOLOR=""White""><FONT CLASS=Small>"
			Response.write mid(oFile.DateLastModified, InStr(1, oFile.DateLastModified, " "))
			Response.write "</FONT></TD><TD ALIGN=RIGHT NOWRAP BGCOLOR=""White""><FONT CLASS=Small>"
			Response.write PrintFileSize(FName)
			Response.write "</FONT></TD></TR>" & vbCrLf
	
			i=i+1			
		end if
	end if    ' strExt

	end if
  
next

  if i > 0 then response.write "</TABLE></TD></TR></TABLE>"	
	
response.write "<BR><BR>"

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide
Set FSO = Nothing

' --------------------------------------------------------------------------------------
' Subroutines
' --------------------------------------------------------------------------------------
	
sub NonPostError

  response.write "<FONT CLASS=MediumBold>"
  response.write "<UL>"
  response.write "<LI>" & Translate("The results of your request have expired.",Login_Language,conn) & "</LI>"
  response.write "</UL>"
  response.write "</FONT>"   
  response.write "<BR><BR>"

end Sub

' --------------------------------------------------------------------------------------

sub PrintNoFind()
  
  response.write "<FONT CLASS=MediumBold>"
  response.write "<UL>"
  response.write "<LI>" & Translate("There have been no new or revised documents posted to this site within the last",Login_Language,conn) & " " & Request.Form("DP") & " " & Translate("days",Login_Language,conn) & ".</LI>"
  response.write "</UL>"
  response.write "</FONT>"
  response.write "<BR><BR>"
  
end sub	

' --------------------------------------------------------------------------------------

sub PrintHeader()
%>

  <UL>
	<LI><%=Translate("Below are the documents or support files that are new or have been revised within the last",Login_Language,conn)%> <FONT COLOR="#FF0000"><B><%= Request.Form("DP") %></B></FONT> <%=Translate("days",Login_Language,conn)%>.<BR><BR></LI>
  </UL>

	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
    <TR>
      <TD>
      
        <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
          <TR>
        		<TH ALIGN=MIDDLE BGCOLOR="Black"><FONT CLASS=SmallBoldGold><%=Translate("File",Login_Language,conn)%></FONT></TH>
        		<TH ALIGN=MIDDLE BGCOLOR="Black"><FONT CLASS=SmallBoldGold><%=Translate("Type",Login_Language,conn)%></FONT></TH>
        		<TH ALIGN=LEFT BGCOLOR="Black"><FONT CLASS=SmallBoldGold><%=Translate("Title or File Name",Login_Language,conn)%> -- <%=Translate("Click left to view right to download",Login_Language,conn)%>)</FONT></TH>
        		<TH ALIGN=LEFT BGCOLOR="Black"><FONT CLASS=SmallBoldGold><%=Translate("Days",Login_Language,conn)%></FONT></TH>
        		<TH COLSPAN=2 BGCOLOR="Black"><FONT CLASS=SmallBoldGold><%=Translate("Last Modified",Login_Language,conn)%></FONT></TH>
        		<TH BGCOLOR="Black"><FONT CLASS=SmallBoldGold><%=Translate("File Size",Login_Language,conn)%></FONT></TH>
      		</TR>
<%
		Response.write "<TR><TD ALIGN=MIDDLE BGCOLOR=""White""><FONT CLASS=Small>"
		Response.write GetExtension(FName)
		Response.write "</FONT></TD><TD ALIGN=CENTER BGCOLOR=""White"">"
		Response.write PrintSquare(GetExtension(FName))
		Response.write "</TD><TD BGCOLOR=""White""><FONT CLASS=Small>"
		Response.write "<B>" & PrintHREF(FName) & PrintFileName(FName) & "</a></B>"
		Response.write PrintDesc(FName)
		Response.write "</FONT></TD>"
		Response.write "<TD ALIGN=RIGHT NOWRAP BGCOLOR=""White""><FONT CLASS=Small>"
		Response.write DateDiff("d",oFile.DateLastModified,date)
		Response.write "</FONT></TD><TD NOWRAP BGCOLOR=""White"" ALIGN=RIGHT><FONT CLASS=Small>"
		Response.write PrintFileDate(Request.form("DF"))
		Response.write "</FONT></TD><TD NOWRAP BGCOLOR=""White""><FONT CLASS=Small>"
'		Response.write TSO.FileTime(FName, 2)
		Response.write mid(oFile.DateLastModified, InStr(1, oFile.DateLastModified, " "))
		Response.write "</FONT></TD><TD ALIGN=RIGHT NOWRAP BGCOLOR=""White""><FONT CLASS=Small>"
		Response.write PrintFileSize(FName)
		Response.write "</FONT></TD></TR>" & vbCrLf
		
		i=i+1

	end sub
	
' --------------------------------------------------------------------------------------
  
function GetExtension(strFileName)
	GetExtension = UCase(Right(strFileName, 3))
end function

' --------------------------------------------------------------------------------------
	
function PrintFileDate(iDateFormat)
	if Len(DatePart("d", oFile.DateLastModified)) = 1 then
		strDay = "0" & DatePart("d", oFile.DateLastModified)
	else
		strDay = DatePart("d", oFile.DateLastModified)
	end if
	
	if Len(DatePart("m", oFile.DateLastModified)) = 1 then
		strMonth = "0" & DatePart("m", oFile.DateLastModified)
	else
		strMonth = DatePart("m", oFile.DateLastModified)
	end if
		
	if iDateFormat = 1 then
		PrintFileDate = strMonth & "/" & strDay & "/" & DatePart("yyyy", oFile.DateLastModified)
	elseif iDateFormat = 2 then
		PrintFileDate = strDay & "/" & strMonth & "/" & DatePart("yyyy", oFile.DateLastModified)
	else
		PrintFileDate = DatePart("yyyy", oFile.DateLastModified) & "/" & strMonth & "/" & strDay
	end if
end function

' --------------------------------------------------------------------------------------
	
function PrintFileSize(strFileName)
	iTemp = oFile.size
	if iTemp < 1024 then
		PrintFileSize = "1 K"
	else
		PrintFileSize = iTemp \ 1024 & " K"
	end if
end function

' --------------------------------------------------------------------------------------	

function PrintSquare(strExtension)
	if UCase(strExtension) = "TIF" or UCase(strExtension) = "PDF" then
		PrintSquare = "<IMG SRC=""/images/balls/r_square.gif"" BORDER=0>"
	elseif Ucase(strExtension) = "ZIP" or UCase(strExtension) = "EXE" then
		PrintSquare = "<IMG SRC=""/images/balls/o_square.gif"" BORDER=0>"
	elseif Ucase(strExtension) = "TXT" or UCase(strExtension) = "DOC" then
		PrintSquare = "<IMG SRC=""/images/balls/k_square.gif"" BORDER=0>"
	elseif Ucase(strExtension) = "HTM" or UCase(strExtension) = "ASP" then
		PrintSquare = "<IMG SRC=""/images/balls/b_square.gif"" BORDER=0>"
	else
		PrintSquare = "<IMG SRC=""/images/balls/y_square.gif"" BORDER=0>"
	end if
end function

' --------------------------------------------------------------------------------------
	
Function PrintDesc(strFileName)
  Dim strDoc_Num
  Dim strDoc_Type

	if UCase(GetExtension(FName)) = "HTM" or UCase(GetExtension(FName)) = "ASP" then
'		PrintDesc = " - " & TSO.ReadTitle(strFileName)
		PrintDesc = " - " & oFile.name
	else
		if vDB(1) = 2 then
			PrintDesc = ""
		elseif vDB(1) = 0 then
			iStart = InStr(UCase(strTag), UCase(strFileName))
			if iStart <> 0 then
				iEnd = InStr(iStart, UCase(strTag), Chr(10))
				iTest = InStr(iStart, UCase(strTag), Chr(13))
				iStart = iStart + Len(strFileName)
				iLength = iEnd - iStart
				if iLength > 0 then
					PrintDesc = " - " & Mid(strTag, iStart, iLength)
				end if
			else
				'Couldnt find the file name in tag file
				PrintDesc = ""
			end if '   iStart
		else
			strDoc_Num = GetNumber(strFileName)
			strDoc_Type = GetPrefix(strFileName)

			'use database
			Set cmd = Server.CreateObject("ADODB.Command")
			Set cmd.ActiveConnection = conn
			cmd.CommandType = adCmdStoredProc
			cmd.CommandText = "SVC_Master_WhatsNew"
			Set prm = cmd.CreateParameter("@strDoc_Num", adVarChar, adParamInput, 10, strDoc_Num)
			cmd.Parameters.Append prm
			Set prm = cmd.CreateParameter("@strDoc_Type", adVarChar, adParamInput, 50, strDoc_Type)
			cmd.Parameters.Append prm

			Set rsResults = Server.CreateObject("ADODB.Recordset")
			rsResults.CursorLocation = adUseClient
			rsResults.CursorType = adOpenDynamic
			rsResults.open cmd
			set prm = nothing
			set cmd = nothing

			If not rsResults.eof then
				strDesc = " - " & rsResults("Description") & "<br>" & Translate("Model(s)",Login_Language,conn) & ": "
				strModel = ""
				do while not rsResults.EOF
					if strModel <> "" then strModel = strModel & ", "
					strModel = strModel & rsResults("Model")
					rsResults.MoveNext
				loop
			else
				PrintDesc = ""
			end if    ' rs.eof
			rsResults.Close
			set rsResults = nothing
			PrintDesc = strDesc & strModel
		end if    'vDB
	end if   ' if HTML
end function

' --------------------------------------------------------------------------------------
	
Function GetPrefix(strFileName)
	for i = 1 to len(strFileName) - 4
		if IsNumeric(Mid(strFileName, i, 1)) then
			'SKIP
		else
			strPrefix = strPrefix & Mid(strFileName, i, 1)
		end if
	next
	GetPrefix = UCase(strPrefix	)
End Function

Function GetNumber(strFileName)
	for i = 1 to len(strFileName) - 4
		if IsNumeric(Mid(strFileName, i, 1)) then
			strNum = strNum & Mid(strFileName, i, 1)
		else
			'SKIP
		end if
	next
	GetNumber = CLng(strNum)
End Function

' --------------------------------------------------------------------------------------
	
Function PrintFileName(strFileName)
	if UCase(GetExtension(FName)) = "TIF" or UCase(GetExtension(FName)) = "PDF" then
		PrintFileName = UCase(GetPrefix(strFileName) & "-" & GetNumber(strFileName))
	else
		PrintFileName = UCase(strFileName)
	end if
end Function

' --------------------------------------------------------------------------------------
	
Function PrintHREF(strFileName)
	PrintHREF = "<a href=""/service-center/" & ReverseSlash(vPath(0)) & strFileName & """>"
End Function

' --------------------------------------------------------------------------------------
	
Function ReverseSlash(strToParse)
	strToParse = CStr(strToParse)
	for i = 1 to Len(strToParse)
		if Mid(strToParse, i, 1) = "/" then
			strReturnValue = strReturnValue & "\"   %><%
		elseif Mid(strToParse, i, 1) = "\" then   %><%
			strReturnValue = strReturnValue & "/"
		else
			strReturnValue = strReturnValue & Mid(strToParse, i, 1)
		end if
	next
	ReverseSlash = strReturnValue
End Function

' --------------------------------------------------------------------------------------
%>
