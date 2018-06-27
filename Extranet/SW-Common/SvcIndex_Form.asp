<%@ Language="VBScript" CODEPAGE="65001" %>

<%
response.buffer = true

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/Pop-Up.asp"-->
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

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Service Documents",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Service Documents",Login_Language,conn) & "</FONT>" 
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

response.write "<FONT CLASS=Heading3>" & Translate("Service Documents",Login_Language,conn) & "</FONT>"
response.write "<BR><BR>"

response.write "<FONT CLASS=Medium>"

if not isblank("ErrorString") then
  response.write "<UL>"
  response.write "<FONT COLOR=""Red"">" & ErrorString & "</FONT>"
  response.write "</UL>"
  Session("ErrorString") = ""
end if

response.write "<UL>"
response.write "<LI>" & Translate("The Model Number can be either the complete Model Number, i.e., <B>Fluke-89</B> or if used with the wildcard suffix such as <B>Fluke-89*</B>, all model numbers beginning with <B>Fluke-89</B> will be returned by the query.",Login_Language,conn) & "<BR><BR></LI>"
response.write "<LI>" & Translate("Enter or select one or more search criteria below:",Login_Language,conn)

response.write "<BR><BR>"

%>
<FORM ACTION="SvcIndex_Results.asp" METHOD="POST">

<TABLE BORDER=1 WIDTH="90%" BORDERCOLOR="#666666" BGCOLOR="<%=Contrast%>" CELLPADDING=0 CELLSPACING=0>
  <TR>
    <TD>

<TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="<%=Contrast%>" WIDTH="100%">
  <TR>
    <TD WIDTH="40%" CLASS=MEDIUMBOLD><%=Translate("Document Type",Login_Language,conn)%> :</TD>
    <TD WIDTH="60%" CLASS=Small>
      <SELECT NAME="Doc_Type" ClASS=Small>
        <OPTION Class=Small VALUE="All" SELECTED><%=Translate("All Documents",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="GENERAL">GEN - <%=Translate("General Information",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="SA">SA_ - <%=Translate("Service Alert",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="PCN">PCN - <%=Translate("Change Notice",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="SPF">SPF - SA/PCN <%=Translate("Support File",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="KIT">KIT - <%=Translate("Parts Kit",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="MET">MET - <%=Translate("Met/Cal Procedure",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="SH">SH_ - <%=Translate("Instruction Sheets",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="UD">UD_ - <%=Translate("User Documentation",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="USU">USU - <%=Translate("User Supplement",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="IM">IM_ - <%=Translate("Instruction Manual",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="MSU">MSU - <%=Translate("Manual Supplement",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="SD">SD_ - <%=Translate("Service Documentation",Login_Language,conn)%></OPTION>
        <OPTION Class=Small VALUE="SSU">SSU - <%=Translate("Service Supplement",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="SI">SI_ - <%=Translate("Service Information",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="CIH">CIH - <%=Translate("Change Notice",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="CIS">CIS - <%=Translate("Change Notice",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="OSC">OSC - <%=Translate("Change Notice",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="SME">SME - <%=Translate("Change Notice",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="SPC">SPC - <%=Translate("Change Notice",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="SRE">SRE - <%=Translate("Change Notice",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="SSY">SSY - <%=Translate("Change Notice",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="DTE">DTE - <%=Translate("Information Sheet",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="SUP">SUP - <%=Translate("Service Supplement",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="SBU">SBU - <%=Translate("Service Bulletin Update",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="ESU">ESU - <%=Translate("Manual Supplement",Login_Language,conn)%> (DTD)</OPTION>
        <OPTION Class=Small VALUE="OBSOLETE">OBS - <%=Translate("Obsolete Documents",Login_Language,conn)%></OPTION>
      </SELECT>
    </TD>
  </TR>

<TR>
  <TD CLASS=MediumBold><%=Translate("Document Number or Numbers Between",Login_Language,conn)%> :</FONT></B></TD>
  <TD CLASS=MediumBold>
    <INPUT CLASS=Small NAME="Doc_Num_Min" SIZE=5 MAXLENGTH=5> <%=Translate("and",Login_Language,conn)%> <INPUT Class=Small NAME="Doc_Num_Max" SIZE=5 MAXLENGTH=5>
  </TD>
</TR>

<TR>
  <TD CLASS=MediumBold><%=Translate("Model Number",Login_Language,conn)%> :</TD>
  <TD CLASS=Small><INPUT Class=Small NAME="Model" SIZE="20" MAXLENGTH="20"></TD>
</TR>

<TR>
  <TD CLASS=MediumBold><%=Translate("New or Revised Documents After Date",Login_Language,conn)%> :</TD>
  <TD CLASS=Small><INPUT CLASS=Small NAME="Date_Month" SIZE=2 MAXLENGTH=2> / <INPUT NAME="Date_Day" SIZE=2 MAXLENGTH=2> / <INPUT NAME="Date_Year" SIZE=4 MAXLENGTH=4> <%=Translate("Use Format: [mm]/[dd]/[yyyy]",Login_language,conn)%>
</TR>

<TR>
  <TD Class=MediumBold><%=Translate("Sort By",Login_Language,conn)%> : </TD>
  <TD VALIGN="MIDDLE" CLASS=Small>
    <SELECT NAME="sort" CLASS=Small>
      <OPTION CLASS=Small VALUE="1" SELECTED><%=Translate("Document Number",Login_Language,conn)%></OPTION>
      <OPTION CLASS=Small VALUE="2"><%=Translate("Assembly",login_Language,conn)%></OPTION>
      <OPTION CLASS=Small VALUE="3"><%=Translate("Document Class Code",Login_Language,conn)%></OPTION>
    </SELECT>
  </TD>
</TR>

<TR>
  <TD CLASS=MediumBold><%=Translate("Number of Results per Screen",Login_Language,conn)%> :</TD>
  <TD CLASS=Small>
    <SELECT NAME="Rows" Class=Small>
      <OPTION CLASS=Small VALUE="10">10</OPTION>
      <OPTION CLASS=Small VALUE="25">25</OPTION>
      <OPTION CLASS=Small VALUE="50" Selected>50</OPTION>
      <OPTION CLASS=Small VALUE="100">100</OPTION>
      <OPTION CLASS=Small VALUE="250">250</OPTION>
    </SELECT>
</TD>
</TR>

<TR>
  <TD COLSPAN=2 BGCOLOR="Black">
    <TABLE WIDTH="100%">
      <TR>
        <TD WIDTH="40%"><INPUT CLASS=NavLeftHighlight1 TYPE="reset" VALUE="<%=Translate("Clear Form",Login_Language,conn)%>"></TD>
        <TD WIDTH="60%"><INPUT CLASS=NavLeftHighlight1 TYPE="submit" VALUE="<%=Translate("Begin Search",Login_Language,conn)%>"></TD>
      </TR>
    </TABLE>
  </TD>   
</TR>
</TABLE>
    </TD>
  </TR>
</TABLE>

</FORM>

</LI>
<LI><%=Translate("All documents are in English only.",Login_Language,conn)%></LI>
</UL>

<!--End Content -->
<BR><BR>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
%>