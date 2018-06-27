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
%>

    <UL>
    <LI><%=Translate("The Model Number can be either the complete Model Number, i.e., Fluke-77 or if used with the wildcard suffix such as Fluke-77*, all model numbers beginning with Fluke-77 will be returned by the query.",Login_Language,conn)%><BR><BR></LI>
    </UL>

    <FORM METHOD="POST" ACTION="/SW-Common/SvcIndex_Model_Results.asp">
    <INPUT NAME="Verbose" TYPE=HIDDEN  VALUE="1">
    
    <UL>
    <LI><%=Translate("Enter your search criteria below",Login_Language,conn)%>: <BR><BR>
    
    <TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" WIDTH="90%"  BGCOLOR="<%=Contrast%>">
      <TR>
        <TD>
          <TABLE BORDER=0 CELLPADDING=4 CELLSPACING=2 BGCOLOR="<%=Contrast%>" WIDTH="100%">
            <TR>
              <TD WIDTH="40%" CLASS=MEDIUMBOLD><%=Translate("Model",Login_Language,conn)%>:</TD>
              <TD WIDTH="60%"><INPUT CLASS=Small NAME="Model" TYPE="text"></TD>
            </TR>

            <TR>
              <TD CLASS=MEDIUMBOLD><%=Translate("Number of Results per Screen",Login_Language,conn)%>:</TD>
              <TD>
                <SELECT Class=Small NAME="Rows">
                  <OPTION VALUE="10">10
                  <OPTION VALUE="25" SELECTED>25
                  <OPTION VALUE="50">50
                  <OPTION VALUE="100">100
                  <OPTION VALUE="250">250
                </SELECT>
              </TD>
            </TR>

            <TR>
            <TD CLASS=MEDIUMBOLD><%=Translate("Search Results View",Login_Language,conn)%></TD>
            <TD>
              <SELECT CLASS=Small NAME="view">
                <OPTION VALUE="0"><%=Translate("Standard",Login_Language,conn)%></OPTION>
                <OPTION VALUE="1" SELECTED><%=Translate("Results Only",Login_Language,conn)%></OPTION>
              </SELECT>
              </FONT>
            </TD>
          </TR>

          <TR>
            <TD BGCOLOR="Black" COLSPAN=2>
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
<%

response.write "<BR><BR>"

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' Subroutines
' --------------------------------------------------------------------------------------
	
%>
