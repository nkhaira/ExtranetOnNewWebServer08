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

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("US Mainframe / Option / Accessory Database",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><SPAN CLASS=SmallBoldGold>" & Translate("US Mainframe / Option / Accessory Database",Login_Language,conn) & "</SPAN>"
Top_Navigation  = False 
Side_Navigation = True
Content_Width   = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

response.write "<SPAN CLASS=Heading3>" & Translate("US Mainframe / Option / Accessory Database",Login_Language,conn) & "</SPAN>"
response.write "<BR><BR>"

response.write "<SPAN CLASS=Medium>"

if not isblank("ErrorString") then
  response.write "<UL>"
  response.write "<FONT COLOR=""Red"">" & ErrorString & "</FONT>"
  response.write "</UL>"
  Session("ErrorString") = ""
end if

response.write "<UL>"
if LimitView = false then
  response.write "<LI>" & Translate("Enter a 6, 7 or 12-digit Fluke Part Number or Model Noun",Login_Language,conn) & ":"
else
  response.write "<LI>" & Translate("Enter a 6, 7 or 12-digit Fluke Part Number",Login_Language,conn) & ":"
end if

response.write "<BR><BR>"

%>
<FORM ACTION="/SW-Common/Part_Query.asp" METHOD="POST">
<INPUT TYPE="Hidden" NAME="dept_id" VALUE="20">

  <%Call Nav_Border_Begin%>
  <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="<%=Contrast%>" WIDTH="100%">
    <TR>
      <% if LimitView = false then %>
        <TD WIDTH="40%" CLASS=MediumBold NOWRAP><%response.write Translate("Part Number or Model Noun",Login_Language,conn)%> :</TD>
      <% else %>
        <TD WIDTH="40%" CLASS=MediumBold><%response.write Translate("Part Number",Login_Language,conn)%> :</TD>
      <% end if %>              
      <TD WIDTH="60%" CLASS=MediumBold>
        <INPUT TYPE="text" NAME="Part" SIZE="30" WIDTH=15 CLASS=Small VALUE="<%=request("part")%>">
      </TD>
    </TR>
    
    <TR>
      <TD CLASS=MediumBold><%response.write Translate("Discount",Login_Language,conn)%> :</TD>
      <TD CLASS=MediumBold>
        <INPUT TYPE="text" NAME="discount" SIZE="5" CLASS=Small VALUE="<%if request("discount") <> "100" then response.write request("discount")%>">&nbsp;&nbsp;&nbsp;<B>%</B>
      </TD>
    </TR>
    <TR>
      <TD CLASS=MediumBold><%response.write Translate("Exchange Rate",Login_Language,conn)%> :</TD>
      <TD CLASS=MediumBold>
        <INPUT TYPE="Text" NAME="rate" SIZE="5" CLASS=Small VALUE="<%if request("rate") <> "1" then response.write request("rate")%>">&nbsp;&nbsp;&nbsp;<B>0.00</B>
      </TD>
    </TR>
    
    <TR>
      <TD CLASS=MediumBold><%response.write Translate("Number of Results per Screen",Login_Language,conn)%> :</TD>
      <TD CLASS=Medium>
        <SELECT NAME="Limit" CLASS=Small>
          <OPTION VALUE="10"<%  if request("Limit")="10" then response.write(" SELECTED") end if%>>10</OPTION>
          <OPTION VALUE="25"<%  if isblank(request("Limit")) or request("limit") = "25" then response.write("SELECTED") end if%>>25</OPTION>
          <OPTION VALUE="50"<%  if request("Limit")="50" then response.write(" SELECTED") end if %>>50</OPTION>
          <OPTION VALUE="100"<% if request("Limit")="100" then response.write(" SELECTED") end if %>>100</OPTION>
          <OPTION VALUE="250"<% if request("Limit")="250" then response.write(" SELECTED") end if %>>250</OPTION>
        </SELECT>
      </TD>
    </TR>
            
    <TR>
      <TD COLSPAN=2 BGCOLOR="#666666">
        <TABLE WIDTH="100%">
          <TR>
            <TD WIDTH="40%" CLASS=Small ALIGN=CENTER><INPUT TYPE="reset" VALUE="<%response.write Translate("Clear Search Criteria",Login_Language,conn)%>" CLASS=NavLeftHighlight1></TD>
            <TD WIDTH="60%" Class=Small ALIGN=CENTER><INPUT TYPE="submit" VALUE="<%response.write Translate("Begin Search",Login_Language,conn)%>" CLASS=NavLeftHighlight1></TD>
          </TR>
        </TABLE>
      </TD>   
    </TR>
  </TABLE>
  <%Call Nav_Border_End%>
</FORM>

<%

Response.write "<BR>"
response.write "</LI>"

if LimitView = false then
  response.write "<LI>" & Translate("A wild card * can be used only as a prefix or suffix to the part number e.g., *9999 or 9999* and not for use with a key word.",Login_Language,conn)
  response.write "&nbsp;&nbsp;" & Translate("Note",Login_Language,conn) & ": "
  response.write Translate("Use of a wild card * in a search take extra time to complete since the database contains thousands of records and each record must be compared to your wild card search criteria.",Login_Language,conn) & "<BR><BR></LI>"
  response.write "<LI>" & Translate("Model Noun",Login_Language,conn) & " " & Translate("search queries the part number&acute;s description field returning records that contain the Model Noun match, i.e., using the key word &quot;Fluke-12&quot;, would find &quot;Fluke-12&quot;, &quot;Fluke-12B&quot;, etc.",Login_Language,conn)
  response.write "<BR><BR></LI>"
end if
  
response.write "<LI>" & Translate("You can apply your standard Fluke discount from US list price to approximate your the cost of this item with your discount applied. For example, if your standard discount is equal to 10%, enter 10. (The default discount is equal to 0% or US List Price.)",Login_Language,conn) & "<BR><BR></LI>"
response.write "<LI>" & Translate("You can apply your local currency conversion to US dollars by specifying your local currency to US dollar exchange rate.  For example, if your local currency is equal to 1.62 US Dollars, enter 1.62. (The default currency exchange rate is equal to 1 or no conversion.)",Login_Language,conn) & "</LI>"

response.write "</UL>"
%>

<!--#include file="SW-Footer.asp"-->

<%
Call Disconnect_Sitewide

response.flush
%>
