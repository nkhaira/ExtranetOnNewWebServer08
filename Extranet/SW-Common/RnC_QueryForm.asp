<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
' --------------------------------------------------------------------------------------

response.buffer = true

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connections_parts.asp" -->
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
Dim LimitView
Dim ErrorString

BackURL = Session("BackURL")    

Set Session("rs") = nothing

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

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("US Repair and Calibration Service Options Database",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("US Repair and Calibration Service Options Database",Login_Language,conn) & "</FONT>"
Top_Navigation  = False 
Side_Navigation = True
Content_Width   = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

response.write "<FONT CLASS=Heading3>" & Translate("US Repair and Calibration Service Options Database",Login_Language,conn)
response.write "</FONT>"
response.write "<BR><BR>"

response.write "<FONT CLASS=Medium>"

if not isblank("ErrorString") then
  response.write "<UL>"
  response.write "<FONT COLOR=""Red"">" & ErrorString & "</FONT>"
  response.write "</UL>"
  Session("ErrorString") = ""
end if

Dim Model_Noun
Model_Noun = request("Model_Noun")

if not isblank(Model_Noun) then

  Call Connect_Parts

  strPart = Replace(Model_Noun, "*", "")
  strPart = Replace(strPart,chr(34),"")           'Remove Double Quotes
  strPart = Replace(strPart,chr(39),"")           'Remove Single Quotes    
       
  SQL = "SELECT catalog_rcnc.pfid AS Model_Number, catalog_rcnc.name AS Model_Noun "
  SQL = SQL & "FROM catalog_rcnc "
  SQL = SQL & "WHERE (((catalog_rcnc.name) Like '" & strPart & "%'"
  if ucase(mid(strPart,1,5)) = "FLUKE" then
   SQL = SQL & " Or (catalog_rcnc.name) Like 'FLK" & mid(strPart,6) & "%'"
  end if
  if ucase(mid(strPart,1,3)) = "FLK" then
   SQL = SQL & " Or (catalog_rcnc.name) Like 'FLUKE" & mid(strPart,4) & "%'"
  end if
  SQL = SQL & ")) "
  SQL = SQL & "ORDER BY catalog_rcnc.name "

  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.open SQL, dbconn, 3, 1, 1

  if rs.EOF then
    Disconnect_Parts
    Model_Noun = ""
    response.write "<UL><LI><FONT COLOR=""#FF0000"">" & Translate("Sorry, no records match the search criteria you have entered.",Login_Language,conn) & "</FONT></LI></UL>"
  end if

end if

response.write "<UL>"
response.write "<LI>"

if isblank(Model_Noun) then
  response.write Translate("Enter",Login_Language,conn) & " "
else
  response.write Translate("Select",Login_Language,conn) & " "
end if

response.write Translate("Product Model Name (noun)",Login_Language,conn) & " "
if isblank(Model_Noun) then
  response.write " " & Translate("or 6 or 7-digit Fluke Model Number",Login_Language,conn) & " : "
else
  response.write " : "
end if

response.write "<BR><BR>"

if isblank(Model_Noun) then
  response.write "<FORM ACTION=""/SW-Common/RnC_QueryForm.asp"" METHOD=""POST"">"
else
  response.write "<FORM ACTION=""/SW-Common/RnC_Query.asp"" METHOD=""POST"">"
end if

Call Nav_Border_Begin
response.write "      <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR=""" & Contrast & """ WIDTH=""100%"">"
response.write "        <TR>"
response.write "          <TD Class=MediumBold WIDTH=""40%"">"
response.write "            <NOBR>"

if isblank(Model_Noun) then
  response.write Translate("Enter",Login_Language,conn) & " "
else
  response.write Translate("Select",Login_Language,conn) & " "
end if

response.write Translate("Product Model Name",Login_Language,conn) & " "
if isblank(Model_Noun) then
  response.write "<BR>" & Translate("or Fluke Model Number",Login_Language,conn) & ":"
else
  response.write ":"
end if
response.write "</TD>"

response.write "<TD ClASS=Medium WIDTH=""60%"">"

if isblank(Model_Noun) then
  response.write "<INPUT Class=Medium TYPE=""text"" NAME=""Model_Noun"" SIZE=30 WIDTH=15 VALUE="""">"
else

  
  response.write "<SELECT Class=Small Name=""Part"">"
      
  Do While Not rs.EOF
    response.write "<OPTION Class=Small VALUE=""" & rs("Model_Number") & """>" & mid(rs("Model_Noun"),1,50) & "</OPTION>" & vbcrlf      
    rs.MoveNext
  Loop
      
  call disconnect_parts            

  response.write "</SELECT>"
end if

response.write "</TD>"
response.write "</TR>"

response.write "<TR>"
response.write "<TD CLASS=MediumBold>"
response.write "<NOBR>"
response.write Translate("Number of Results per Screen",Login_Language,conn) & ":"
response.write "</TD>"

response.write "<TD CLASS=Small>"

%>
  <SELECT Class=Small NAME="Limit">
    <OPTION Class=Small VALUE="10" <%  if request("limit") = "10" then response.write("SELECTED") end if %>>10</OPTION>
    <OPTION Class=Small VALUE="25" <%  if isblank(request("limit")) or request("limit") = "25" then response.write("SELECTED") end if %>>25</OPTION>
    <OPTION Class=Small VALUE="50" <%  if request("limit") = "50" then response.write("SELECTED") end if %>>50</OPTION>
    <OPTION Class=Small VALUE="100" <% if request("limit") = "100" then response.write("SELECTED") end if %>>100</OPTION>
    <OPTION Class=Small VALUE="250" <% if request("limit") = "250" then response.write("SELECTED") end if %>>250</OPTION>
  </SELECT>
<%

response.write "</TD>"
response.write "</TR>"

response.write "<TR>"
response.write "<TD CLASS=MediumBold>"
response.write "<NOBR>"
response.write Translate("Search Results View",Login_Language,conn) & ":"
response.write "</TD>"

response.write "<TD CLASS=Small>"
response.write "<NOBR>"

%>
  <SELECT Class=Small NAME="view">
    <OPTION Class=Small VALUE="0"><%response.write Translate("Standard",Login_Language,conn)%></OPTION>
    <% if request("view") = "1" then %>
    <OPTION Class=Small VALUE="1" SELECTED><%response.write Translate("Results Only",Login_Language,conn)%></OPTION>
    <% elseif request("view") = "2" then %>
    <OPTION Class=Small VALUE="2" SELECTED><%response.write Translate("Results Only",Login_Language,conn)%></OPTION>
    <% else %>
    <OPTION Class=Small VALUE="1"><%response.write Translate("Results Only",Login_Language,conn)%></OPTION>  
    <% end if %>
  </SELECT>
<%

response.write "</TD>"
response.write "</TR>"
        
response.write "<TR>"
response.write "<TD COLSPAN=2 BGCOLOR=""#666666"">"
response.write "<TABLE WIDTH=""100%"">"
response.write "<TR>"
response.write "<TD ClASS=Small WIDTH=""40%"" ALIGN=CENTER>"
if isblank(Model_Noun) then
  response.write "<INPUT Class=NavLeftHighlight1 TYPE=""reset"" VALUE=""" & Translate("Clear Form",Login_Language,conn) & """>"
else
  response.write "&nbsp;"
end if
response.write "</TD>"
                
response.write "<TD Class=Small WIDTH=""60%"" ALIGN=CENTER>"
if isblank(Model_Noun) then
  response.write "<INPUT Class=NavLeftHighlight1 TYPE=""submit"" VALUE=""" & Translate("Find Models",Login_Language,conn) & """>"
else
  response.write "<INPUT Class=NavLeftHighlight1 TYPE=""submit"" VALUE=""" & Translate("Display US Repair and Calibration Pricing",Login_Language,conn) & """>"
end if

response.write "</TD>"
response.write "</TR>"
response.write "</TABLE>"
response.write "</TD>"
response.write "</TR>"
response.write "</TABLE>"
Call Nav_Border_End
response.write "</FORM>"

response.write "<BR>"
response.write "</LI>"
response.write "</UL>"

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide
response.flush

%>
