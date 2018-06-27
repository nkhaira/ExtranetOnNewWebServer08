<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
'             Sandbox
' --------------------------------------------------------------------------------------

response.buffer = true

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

if request("Site_ID") <> "" and isnumeric(request("Site_ID")) then
  Site_ID            = request("Site_ID")
  Session("Site_ID") = request("Site_ID")  
elseif session("Site_ID") <> "" and isnumeric(session("Site_ID")) then
  Site_ID = session("Site_ID")
else
  response.redirect "/register/default.asp"
end if

Dim BackURL
BackURL = Session("BackURL")    

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connections_parts.asp"-->
<%

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code = rsSite("Site Code")
Site_Description = rsSite("Site Description")
Logo = rsSite("Logo")
Footer_Disabled = rsSite("Footer_Disabled")
  
rsSite.close
set rsSite=nothing

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("US Repair and Calibration Service Options Database",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("US Repair and Calibration Service Options Database",Login_Language,conn) & "</FONT>"
 
Side_Navigation = False
Content_Width   = 95  ' Percent

%>
<!--#include file="SW-Header.asp"-->
<!--#include file="SW-Common-Navigation.asp"-->
<%
response.write "<FONT CLASS=Heading3>" & Translate("US Repair and Calibration Service Options Database",Login_Language,conn)
response.write "</FONT><BR><BR>"
%>

<UL>
<LI>
<% if isnull(request("Model_Noun")) or request("Model_Noun") = "" then
    response.write Translate("Enter",Login_Language,conn) & " "
   else
    response.write Translate("Select",Login_Language,conn) & " "
   end if
   response.write Translate("Product Model Name (noun)",Login_Language,conn) & " "
   if isnull(request("Model_Noun")) or request("Model_Noun") = "" then
    response.write " " & Translate("or 6 or 7-digit Fluke Model Number",Login_Language,conn) & " : "
   else
    response.write " : "
   end if
%>

<BR><BR>

<% if isnull(request("Model_Noun")) or request("Model_Noun") = "" then
    response.write "<FORM ACTION=""rncQueryForm.asp"" METHOD=""POST"">"
   else
    response.write "<FORM ACTION=""rncQuery.asp"" METHOD=""POST"">"
   end if
%>

<TABLE BORDER=1 WIDTH="90%" BORDERCOLOR="#666666" BGCOLOR="<%=Contrast%>" CELLPADDING=0 CELLSPACING=0>
  <TR>
    <TD>
      <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="<%=Contrast%>" WIDTH="100%">
        <TR>
          <TD WIDTH="40%">
            <NOBR>
            <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica"><B>
            <% if isnull(request("Model_Noun")) or request("Model_Noun") = "" then
                response.write Translate("Enter",Login_Language,conn) & " "
               else
                response.write Translate("Select",Login_Language,conn) & " "
               end if
               response.write Translate("Product Model Name",Login_Language,conn) & " "
               if isnull(request("Model_Noun")) or request("Model_Noun") = "" then
                response.write "<BR>" & Translate("or Fluke Model Number",Login_Language,conn) & ":"
               else
                response.write ":"
               end if
            %>
            </B></FONT>
          </TD>
          <TD WIDTH="60%">
          
            <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica">

            <% if isnull(request("Model_Noun")) or request("Model_Noun") = "" then
                response.write "<INPUT TYPE=""text"" NAME=""Model_Noun"" SIZE=30 WIDTH=15 VALUE="""">"
               else
                call connect_parts

		            strPart = Replace(Request("Model_Noun"), "*", "")
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
                          
                response.write "<SELECT Name=""Part"">"
      
                Do While Not rs.EOF                                  
                   response.write "<OPTION VALUE=""" & rs("Model_Number") & """>" & mid(rs("Model_Noun"),1,50) & "</OPTION>" & vbcrlf      
              		 rs.MoveNext      
                Loop
      
                call disconnect_parts            

                response.write "</SELECT>"
              end if
            %>
            </FONT>
          </TD>
        </TR>
        
<!--
        <TR>
          <TD>
            <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica"><B>Discount :</B></FONT></TD>
          <TD>
            <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica">
            <INPUT TYPE="text" NAME="discount" SIZE="5" VALUE="<%if request("discount") <> "100" then response.write request("discount")%>">&nbsp;&nbsp;&nbsp;<B>%</B>
            </FONT>
          </TD>
        </TR>

        <TR>
          <TD>
            <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica"><B>Exchange Rate :</B></FONT> </B></TD>
          <TD>
            <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica">
            <INPUT TYPE="Text" NAME="rate" SIZE="5" VALUE="<%if request("rate") <> "1" then response.write request("rate")%>">&nbsp;&nbsp;&nbsp;<B>0.00</B>
            </FONT>
          </TD>
        </TR>
-->                
        <TR>
          <TD>
            <NOBR>
            <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica"><B><% response.write Translate("Number of Results per Screen",Login_Language,conn)%> :</B></FONT>
          </TD>
          <TD>
            <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica">
            <SELECT NAME="Returned">
              <OPTION VALUE="10" <% if request("Returned")="10" or request("limit") = "10" then response.write("SELECTED") end if %>>10</OPTION>
              <OPTION VALUE="25" <% if request("Returned")="" or request("limit") = "" or request("limit") = "25" then response.write("SELECTED")end if %>>25</OPTION>
              <OPTION VALUE="50" <% if request("Returned")="50" or request("limit") = "50" then response.write("SELECTED") end if %>>50</OPTION>
              <OPTION VALUE="100" <% if request("Returned")="100" or request("limit") = "100" then response.write("SELECTED") end if %>>100</OPTION>
              <OPTION VALUE="250" <% if request("Returned")="250" or request("limit") = "250" then response.write("SELECTED") end if %>>250</OPTION>
            </SELECT>
            </FONT>
          </TD>
        </TR>

        <TR>
          <TD>
            <NOBR>
            <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica"><B><%response.write Translate("Search Results View",Login_Language,conn)%> :</B></FONT>
          </TD>
          <TD><NOBR>
            <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica">
            <SELECT NAME="view">
              <OPTION VALUE="0"><%response.write Translate("Standard",Login_Language,conn)%></OPTION>
              <% if request("view") = "1" then %>
              <OPTION VALUE="1" SELECTED><%response.write Translate("Results Only",Login_Language,conn)%></OPTION>
              <% elseif request("view") = "2" then %>
              <OPTION VALUE="2" SELECTED><%response.write Translate("Results Only",Login_Language,conn)%></OPTION>
              <% else %>
              <OPTION VALUE="1"><%response.write Translate("Results Only",Login_Language,conn)%></OPTION>  
              <% end if %>
            </SELECT>
            </FONT>
          </TD>
        </TR>
        
        <TR>
          <TD COLSPAN=2 BGCOLOR="#666666">
            <TABLE WIDTH="100%">
              <TR>
                <TD WIDTH="40%">
                  <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica">    
                  <% if isnull(request("Model_Noun")) or request("Model_Noun") = "" then %>
                    <INPUT TYPE="reset" VALUE="<%response.write Translate("Clear Form",Login_Language,conn)%>">
                  <% else %>
                    &nbsp;
                  <% end if %>    
                  </FONT>
                </TD>
                
                <TD WIDTH="60%">
                  <FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica">
                  <% if isnull(request("Model_Noun")) or request("Model_Noun") = "" then %>
                    <INPUT TYPE="submit" VALUE="<%response.write Translate("Find Models",Login_Language,conn)%>">
                  <% else %>
                    <INPUT TYPE="submit" VALUE="<%response.write Translate("Display US Repair and Calibration Pricing",Login_Language,conn)%>">
                  <% end if %>
                  </FONT>
                </TD>
              </TR>
            </TABLE>
          </TD>   
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>

<BR></LI>
</UL>

<!--#include file="SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
%>