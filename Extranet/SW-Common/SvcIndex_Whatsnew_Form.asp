<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
' --------------------------------------------------------------------------------------

'response.buffer = true

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
response.write "<FONT CLASS=Heading4>" & Translate("What&acute;s New",Login_Language,conn) & "</FONT>"
response.write "<BR><BR>"

response.write "<FONT CLASS=Medium>"

if not isblank("ErrorString") then
  response.write "<UL>"
  response.write "<FONT COLOR=""Red"">" & ErrorString & "</FONT>"
  response.write "</UL>"
  Session("ErrorString") = ""
end if

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

Dim DocPath
DocPath = "Download\Documents"


response.write Translate("What's New provides a quick scan of all the time critical documents posted to the CSS Service Center Support Site.",Login_Language,conn) & "  " & Translate("The results of this scan can either be viewed on-line, or downloaded for local use.",Login_Language,conn) & " " & Translate("To use What's New, follow the instructions below:",Login_Language,conn)
response.write "<BR>"
response.write "<UL>"
response.write "<LI>" & Translate("Select the Service Center Support - Document / File Category (What's New Document Category:) from the list below.",Login_Language,conn) & "</LI>"
response.write "<LI>" & Translate("Select how many days prior to (Search Back:) today's date, you would like to begin your search.",Login_Language,conn) & "</LI>"
response.write "<LI>" & Translate("Select your prefered (Date Format:)",Login_Language,conn) & "</LI>"
response.write "<LI>" & Translate("Press the [Begin Search] button to start the What's New query.",Login_Language,conn) & "</LI>"
response.write "<BR><BR>"

%>
<FORM ACTION="/SW-Common/SvcIndex_FileSearch.asp" METHOD="POST">

<INPUT TYPE="Hidden" NAME="PostTest" VALUE="1">

<TABLE BORDER=1 WIDTH="90%" BORDERCOLOR="#666666" BGCOLOR="<%=Contrast%>" CELLPADDING=0 CELLSPACING=0>
  <TR>
    <TD>
      <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="<%=Contrast%>"  WIDTH="100%">
        <TR>
          <TD WIDTH="40%"><FONT FACE="ARIAL, Verdana, Helvetica" SIZE=2><B><%=Translate("What's New Document Category",Login_Language,conn)%> :</FONT></TD>
          <TD WIDTH="60%">
            <SELECT NAME="Path" Class=Small>

            <!-- Must always be in the order: path&title&dd&db -->
            <OPTION Class=Small VALUE="<%=DocPath%>\updates\&Title=CSS+Service+Support+Document+Updates&DD=1&DB=1"><%=Translate("Support Document Updates (PCNs, SA, etc.)",Login_Language,conn)%></OPTION>         
            <OPTION Class=Small VALUE="<%=DocPath%>\support_files\&Title=CSS+Service+Support+Files&DD=1&DB=0"><%=Translate("Support File Updates for (PCNs, SA, etc.)",Login_Language,conn)%></OPTION>            
            <OPTION Class=Small VALUE="<%=DocPath%>\updates_database\&Title=CSS+Service+Support+Database+Updates&DD=1&DB=0"><%=Translate("Support Database Updates",Login_Language,conn)%></OPTION>            
            <OPTION Class=Small VALUE="<%=DocPath%>\Service_Quality\&Title=CSS+Service+Quality+Documents&DD=1&DB=0"><%=Translate("CSS Service Quality Documents",Login_Language,conn)%></OPTION>            
            <OPTION Class=Small VALUE="<%=DocPath%>\Service_MetCal\&Title=CSS+Service+Support+MET/CAL+Procedure Updates&DD=1&DB=0"><%=Translate("Support MET/CAL Procedure Updates",Login_Language,conn)%></OPTION>      
            </SELECT>
          </TD>
        </TR>

        <TR>
          <TD Class=MediumBold><%=Translate("Search Back",Login_Language,conn)%> :</TD>
          <TD Class=Small>
            <SELECT NAME="DP" Class=Small>
              <OPTION Class=Small VALUE="7" SELECTED>7</OPTION>
              <OPTION Class=Small VALUE="14">14</OPTION>
              <OPTION Class=Small VALUE="30">30</OPTION>
              <OPTION Class=Small VALUE="60">60</OPTION>
              <OPTION Class=Small VALUE="90">90</OPTION>
              <OPTION Class=Small VALUE="180">180</OPTION>
              <OPTION Class=Small VALUE="360">360 (<%=Translate("Long Listing",Login_Language,conn)%>)</OPTION>
            </SELECT>
            <%=Translate("days",Login_Language,conn)%>.
          </TD>
        </TR>

        <TR>
        <TD Class=MediumBold><%=Translate("Preferred Date Format",Login_Language,conn)%> :</TD>
        <TD Class=SmallBold>
          <INPUT TYPE="RADIO" NAME="DF" VALUE="1" <% if request("DF") = "1" or request("DF") = "" then response.write "CHECKED"%>> <FONT FACE="ARIAL, Verdana, Helvetica" SIZE=2>mm/dd/yyyy</FONT><BR>
          <INPUT TYPE="RADIO" NAME="DF" VALUE="2" <% if request("DF") = "2" then response.write "CHECKED"%>> <FONT FACE="ARIAL, Verdana, Helvetica" SIZE=2>dd/mm/yyyy</FONT><BR>
          <INPUT TYPE="RADIO" NAME="DF" VALUE="3" <% if request("DF") = "3" then response.write "CHECKED"%>> <FONT FACE="ARIAL, Verdana, Helvetica" SIZE=2>yyyy/mm/dd</FONT>
        </TD>
      </TR>
      
      <TR>
        <TD Class=MediumBold><%=Translate("Search Results View",Login_Language,conn)%> :</TD>
        <TD Class=Small>
          <SELECT Class=Small NAME="view">
            <OPTION Class=Small VALUE="0"><%=Translate("Standard",Login_Language,conn)%></OPTION>
            <% if request("view") = "1" then %>
            <OPTION Class=Small VALUE="1" SELECTED><%=Translate("Results Only",Login_Language,conn)%></OPTION>
            <% elseif request("view") = "2" then %>
            <OPTION Class=Small VALUE="2" SELECTED><%=Translate("Results Only",Login_Language,conn)%></OPTION>
            <% else %>
            <OPTION Class=Small VALUE="1"><%=Translate("Results Only",Login_Language,conn)%></OPTION>            
            <% end if %>
          </SELECT>
        </TD>
      </TR>

      <TR>
        <TD COLSPAN=2 BGCOLOR="Black">
          <TABLE WIDTH="100%">
            <TR>
              <TD WIDTH="40%"><INPUT Class=NavLeftHighlight1 TYPE="reset" VALUE="<%=Translate("Clear Form",Login_Language,conn)%>"></TD>
              <TD WIDTH="60%"><INPUT Class=NavLeftHighlight1 TYPE="submit" VALUE="<%=Translate("Begin Search",Login_Language,conn)%>"></TD>
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
</FONT>

<!--End Content -->
<BR><BR>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
%>