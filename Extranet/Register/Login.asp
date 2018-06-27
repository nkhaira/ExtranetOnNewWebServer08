<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author: K. Whitlock
' Date:   2/1/2000
' --------------------------------------------------------------------------------------

Dim strBackURL

Dim Script_Debug
Script_Debug = false

strBackURL = request("BackURL")

if instr(1,lcase(request("Action")),"cancel") > 0 or instr(1,lcase(request("Submit")),"cancel") > 0 then
  Session("Site_ID")    = NULL
  Session("LOGON_USER") = NULL
  Session("Password")   = NULL                              
  
  if LCase(request.ServerVariables("Script_Name")) = "/register/login.asp" then
    strBackURL = ""
  end if

  if not isblank(strBackURL) and (instr(1,LCase(strBackURL),"https://") > 0 or instr(1,LCase(strBackURL),"http://")  > 0) then
    response.redirect strBackURL
  else
    response.redirect "/register/default.asp"
  end if  
end if

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim Site_ID
Dim Login_Attempts
Dim User_Name
Dim Password
Dim Validated

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Security Module used by SiteWide Auxilliary Applications
' --------------------------------------------------------------------------------------

if not isblank(request("Locator")) then
  Locator = request("Locator")
else
  Locator = ""
end if

if not isblank(request("Screen_Width")) then
  Session("Screen_Width")  = request("Screen_Width")
  Session("Screen_Height") = request("Screen_Height")  
end if

%>
<!--#include virtual="/include/functions_locator.asp"-->
<%

if Script_Debug then
  response.write "Locator Values<BR>"
  for x = 0 to Parameter_Max
    response.write x & "&nbsp;&nbsp;&nbsp;" & Parameter(x) & "&nbsp;&nbsp;&nbsp;|"
    for y = 1 to len(Parameter(x))
      response.write Asc(Mid(Parameter(x),y,1)) & " | "
    next
    response.write "<BR>" & vbCrLf  
  next
  response.flush
  response.end
end if

if not isblank(request("Site_ID")) and isnumeric(request("Site_ID")) then
   Site_ID            = CInt(request("Site_ID"))
   Session("Site_ID") = CInt(request("Site_ID"))
elseif not isblank(session("Site_ID")) and isnumeric(session("Site_ID")) then
   Site_ID = CInt(Session("Site_ID"))
elseif not isblank(Parameter(xSite_ID)) then
   Site_ID            = CInt(Parameter(xSite_ID))
   Session("Site_ID") = CInt(Parameter(xSite_ID))
else
   Call Disconnect_Sitewide
   Session("ErrorString") = "<LI>" & Translate("Your session has expired.",Login_Language,conn) & " " & Translate("For your protection, you have been automatically logged off of your extranet site account.",Login_Language,conn) & "</LI><LI>" & Translate("To establish another session, please type in the site's code in the &quot;Name of the Site where you want to go&quot;, then click on [ Login ] or",Login_Language,conn) & "</LI><LI>" & Translate("Use the Site Search feature below.",Login_Language,conn) & "</LI>"
   response.redirect "/register/default.asp"
end if

' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

Validated = False

if not isblank(Session("Login_Attempts")) then
  Login_Attempts = CInt(Session("Login_Attempts"))
else
  Login_Attempts = 0
end if

User_Name = Trim(request("User_Name"))
Password  = Trim(request("Password"))

if Site_ID = 100 and instr(1,LCase(request.ServerVariables("SERVER_NAME")),"flukenetworks") > 0 then
  response.write "<LINK REL=STYLESHEET HREF=""/portweb/SW-Style.css"">" & vbCrLf
  Logo_Left = true
  Logo = "/images/FlukeNetworks-Logo.gif"
  Site_Description = "Support.FlukeNetworks.com"
end if

Screen_Title      = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Logon",Alt_Language,conn)
Bar_Title         = Translate(Site_Description,Login_Language,conn) &  "<BR><SPAN CLASS=SmallBoldBar>" & Translate("Logon",Login_Language,conn) & "</SPAN>"
Navigation        = False
Top_Navigation    = False
Content_Width     = 95  ' Percent

if not isblank(User_Name) and not isblank(Password) and Login_Attempts <= 3 then

  if Site_ID = 100 then
    SQL =  "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & User_Name & "' AND UserData.NewFlag=" & CInt(False)  
  else
    SQL =  "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & User_Name & "' AND UserData.Site_ID=" & Site_ID & " AND UserData.NewFlag=" & CInt(False)
  end if
    
  Set rsLogin = Server.CreateObject("ADODB.Recordset")
  rsLogin.Open SQL, conn, 3, 3
  
  ' Site Administrator
  
  if Site_ID = 100 then
  
    do while not rslogin.EOF

      ' Find any Administration Account

      if (LCase(rsLogin("NTLogin"))  = LCase(User_Name) _
        and rsLogin("Password") = Password) _
        and (instr(1,LCase(rsLogin("SubGroups")),LCase("domain"))        > 0 _
        or   instr(1,LCase(rsLogin("SubGroups")),LCase("administrator")) > 0 _
        or   instr(1,LCase(rsLogin("SubGroups")),LCase("account"))       > 0 _
        or   instr(1,LCase(rsLogin("SubGroups")),LCase("content"))       > 0 _
        or   instr(1,LCase(rsLogin("SubGroups")),LCase("submitter"))     > 0 _
        or   instr(1,LCase(rsLogin("SubGroups")),LCase("literature"))    > 0 _        
        or   instr(1,LCase(rsLogin("SubGroups")),LCase("forum"))         > 0) then

        Session("LOGON_USER")     = User_Name
        Session("Password")       = rsLogin("Password")        
        Session("Language")       = rsLogin("Language")
        Session("Login_Attempts") = Null

        response.write "<HTML>" & vbCrLf
        response.write "<HEAD>" & vbCrLf
        response.write "<TITLE>Account Verified</TITLE>" & vbCrLf
        response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8"">" & vbCrLf
        response.write "</HEAD>" & vbCrLf
        response.write "<BODY BGCOLOR=""White"" onLoad='document.FORM0.submit()'>" & vbCrLf
      
      	if trim(strBackURL) = "" then
          response.write "<FORM NAME=""FORM0"" ACTION=""" & Site_URL & "/default.asp"" METHOD=""POST"">" & vbCrLf
      	else
      	  response.redirect strBackURL
      	  response.write "<FORM NAME=""FORM0"" ACTION=""" & strBackURL & " METHOD=""POST"">" & vbCrLf
      	end if
        response.write "<INPUT TYPE=""HIDDEN"" NAME=""LOGON_USER"" VALUE=""" & User_Name & """>" & vbCrLf
        response.write "<INPUT TYPE=""HIDDEN"" NAME=""PASSWORD"" VALUE=""" & rsLogin("Password") & """>" & vbCrLf        
        response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""100"">" & vbCrLf
	      response.write "<INPUT TYPE=""HIDDEN"" NAME=""BackURL"" VALUE=""" & strBackURL & """>" & vbCrLf
        if Login_Language <> "eng" and Login_Language <> Session("Language") then
          response.write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
        end if  

        response.write "</FORM>" & vbCrLf
        response.write "</BODY>" & vbCrLf
        response.write "</HTML>" & vbCrLf
        
        Validated = True
        
        exit do

      end if
      
      rsLogin.MoveNext  
      
    loop     

  ' All other Sites
  
  else

    if not rsLogin.EOF then

      if LCase(rsLogin("NTLogin"))  = LCase(User_Name) _
        and rsLogin("Password") = Password then

        Session("LOGON_USER")     = User_Name
        Session("Password")       = rsLogin("Password")
        Session("Language")       = rsLogin("Language") 
        Session("Login_Attempts") = Null
  
        response.write "<HTML>" & vbCrLf
        response.write "<HEAD>" & vbCrLf
        response.write "<TITLE>Account Verified</TITLE>" & vbCrLf
        response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8"">" & vbCrLf
        response.write "</HEAD>" & vbCrLf
        response.write "<BODY BGCOLOR=""White"" onLoad='document.FORM1.submit()'>" & vbCrLf
      	if trim(strBackURL) = "" then
          response.write "<FORM NAME=""FORM1"" ACTION=""" & Site_URL & "/default.asp"" METHOD=""GET"">" & vbCrLf
      	else
      	  response.redirect strBackURL
      	  response.write "<FORM NAME=""FORM1"" ACTION=""" & strBackURL & " METHOD=""GET"">" & vbCrLf
      	end if
        response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
      	response.write "<INPUT TYPE=""HIDDEN"" NAME=""BackURL"" VALUE=""" & strBackURL & """>" & vbCrLf
        if Login_Language <> "eng" and Login_Language <> Session("Language") then
          response.write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
        end if
        
        if not isblank(Locator) then
          response.write "<INPUT TYPE=""HIDDEN"" NAME=""" & Parameter_Key(xCID)  & """ VALUE=""" & Parameter(xCID)  & """>" & vbCrLf
          response.write "<INPUT TYPE=""HIDDEN"" NAME=""" & Parameter_Key(xSCID) & """ VALUE=""" & Parameter(xSCID) & """>" & vbCrLf
          response.write "<INPUT TYPE=""HIDDEN"" NAME=""" & Parameter_Key(xPCID) & """ VALUE=""" & Parameter(xPCID) & """>" & vbCrLf
          response.write "<INPUT TYPE=""HIDDEN"" NAME=""" & Parameter_Key(xCIN)  & """ VALUE=""" & Parameter(xCIN)  & """>" & vbCrLf
          response.write "<INPUT TYPE=""HIDDEN"" NAME=""" & Parameter_Key(xCINN) & """ VALUE=""" & Parameter(xCINN) & """>" & vbCrLf          
        end if

        response.write "</FORM>" & vbCrLf
        response.write "</BODY>" & vbCrLf
        response.write "</HTML>" & vbCrLf

        Validated = True
      
      end if
      
    end if  
    
  end if
  
  ' Account not verified
  
  if Validated = False then

    response.write "<HTML>" & vbCrLf
    response.write "<HEAD>" & vbCrLf
    response.write "<TITLE>Account Not Verified</TITLE>" & vbCrLf
    response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8"">" & vbCrLf
    response.write "</HEAD>" & vbCrLf
    response.write "<BODY BGCOLOR=""White"" onLoad='document.FORM2.submit()'>" & vbCrLf
    response.write "<FORM NAME=""FORM2"" ACTION=""/register/login.asp"" METHOD=""POST"">" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""BackURL"" VALUE=""" & strBackURL & """>" & vbCrLf

    if not isblank(Locator) then
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Locator"" VALUE=""" & Locator & """>" & vbCrlf      
    end if
    response.write "</FORM>" & vbCrLf
    response.write "</BODY>" & vbCrLf
    response.write "</HTML>" & vbCrLf

    Session("LOGON_User")     = NULL
    Session("Password")       = NULL      
    Session("Login_Attempts") = Login_Attempts + 1
      
  end if
  
  rsLogin.Close
  set rsLogin = nothing
  
elseif Login_Attempts > 3 then

  with response
    .write "<%"
    .write "session.abandon"
    .write chr(asc("%")) & ">"
    .write "<HTML><HEAD><TITLE>Error 401.1</TITLE>"
    .write "<META NAME=""robots"" CONTENT=""noindex"">"
    .write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8""></HEAD>"
    .write "<BODY>"
    .write "<H2>HTTP Error 401</H2>"
    .write "<P><STRONG>401.1 Unauthorized: Logon Failed</STRONG></P>"
    .write "<P>This error indicates that the credentials passed to the server do not match the credentials required to log on to the server.</P>"
    .write "<P>Please contact the Web server's administrator to verify that you have permission to access the requested resource.</P>"
    .write "</BODY></HTML>"
   end with

else

  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <%

  response.write "<TABLE BORDER=0 WIDTH=""100%"" COLSPACING=0 CELLSPACING=0>"
  response.write "<TR>"
  response.write "<TD WIDTH=""100%"" HEIGHT=6 CLASS=TopColorBar><IMG SRC=""/images/1x1trans.gif"" HEIGHT=6 BORDER=0 VSPACE=0></TD>" & vbCrLf
  response.write "</TR>"
  response.write "<TR>"
  response.write "<TD CLASS=SMALL WIDTH=""100%"" BGCOLOR=WHITE>"
  
  response.write "<DIV ALIGN=CENTER>"
  response.write "<TABLE BORDER=0 WIDTH=""" & Content_Width & "%"">"
  response.write "<TR>"
  response.write "<TD WIDTH=""100%"" CLASS=Medium>"
  
  response.write "<BR><BR>"

  if not isblank(request("ErrorString")) then
    ErrorString = Request("ErrorString")
  elseif not isblank(Session("ErrorString")) then
    ErrorString = Session("ErrorString")
  else
    ErrorString = ""
  end if

  if not isblank(ErrorString) then
    response.write "<UL><FONT CLASS=MediumBold>" & ErrorString & "</FONT></UL>" & vbCrLf
    ErrorString = ""
    Session("ErrorString") = ""
  end if

  'response.write "<BR><BR><BR><BR>"
  Session("LOGON_User") = NULL
  Session("Password")   = NULL
  
  %>  
  
  <DIV ALIGN=CENTER>

  <FORM NAME="FORM4" ACTION="/register/login.asp" METHOD="POST">
  <INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
  <INPUT TYPE="Hidden" NAME="Language" VALUE="<%=Login_Language%>">
  <INPUT TYPE="Hidden" NAME="Screen_Width"  VALUE="">
  <INPUT TYPE="Hidden" NAME="Screen_Height" VALUE="">
  <%
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""BackURL"" VALUE=""" & strBackURL & """>" & vbCrLf

  if not isblank(Locator) then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Locator"" VALUE=""" & Locator & """>" & vbCrlf      
  end if
  %>
    
  <TABLE WIDTH=402 HEIGHT=254 BORDER=2 CELLSPACING=0 CELLPADDING=0 BORDER=0>
    <TR>
      <TD WIDTH="100%" BGCOLOR="#C2BFAF">
        <TABLE WIDTH=402 HEIGHT=254 BGCOLOR="#C2BFA5" CELLSPACING=0 CELLPADDING=2 BORDER=0>
          <TR>
            <TD COLSPAN=3 BGCOLOR=MAROON HEIGHT=12>
            <!--IMG SRC="/images/Button-Close.gif" ALIGN=RIGHT BORDER=0 ALT="Cancels Network Login Request."-->
            <SPAN STYLE="font-family:Arial;font-size:10pt;color:#E1E0D2;font-weight:bold">&nbsp;Enter Network Password</SPAN></TD>
          </TR>
  
          <TR>
            <TD ROWSPAN=8 ALIGN=CENTER VALIGN=TOP WIDTH=52><BR><IMG SRC="/images/IE-Logon-Key.jpg" BORDER=0></TD>
            <TD WIDTH=350 COLSPAN=2 HEIGHT=8><IMG SRC="/images/1x1trans.gif" HEIGHT=8 BORDER=0 VSPACE=0></TD>
          </TR>
  
          <TR>
            <TD COLSPAN=2 VALIGN=BOTTOM><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black">Please type your user name and password:</SPAN></TD>
          </TR>
  
          <TR>
            <TD WIDTH= 75 VALIGN=BOTTOM><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black">Site:</SPAN></TD>
            <TD WIDTH=275 VALIGN=BOTTOM><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><%=UCASE(request.ServerVariables("SERVER_NAME"))%></SPAN></TD>          
          </TR>
  
          <TR>
            <TD VALIGN=BOTTOM><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><U>U</U>ser Name</SPAN><BR><IMG SRC="/images/1x1trans.gif" HEIGHT=4 BORDER=0 VSPACE=0></TD>
            <TD VALIGN=BOTTOM><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><INPUT TYPE="TEXT" NAME="User_Name" MAXLENGTH=30></SPAN></TD>
          </TR>
  
          <TR>
            <TD VALIGN=BOTTOM><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><U>P</U>assword</SPAN><BR><IMG SRC="/images/1x1trans.gif" HEIGHT=4 BORDER=0 VSPACE=0></TD>
            <TD VALIGN=BOTTOM><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><INPUT TYPE="PASSWORD" NAME="Password"MAXLENGTH=30></SPAN></TD>
          </TR>
  
          <!--TR>
            <TD VALIGN=BOTTOM><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><U>D</U>omain:</SPAN><BR><IMG SRC="/images/1x1trans.gif" HEIGHT=4 BORDER=0 VSPACE=0></TD>
            <TD VALIGN=BOTTOM><SPAN STYLE="background-color:#C2BFA5;font-family:Arial;font-size:8.5pt;color:Black"><INPUT READONLY TYPE="TEXT" NAME="Domain"></SPAN></TD>
          </TR>
  
          <TR>
            <TD COLSPAN=2>
              <INPUT DISABLED TYPE="CHECKBOX" NAME="Save_Password">
              <SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><U>S</U>ave this password in your password list</SPAN>
            </TD>  
            
          </TR-->
  
          <TR>
            <TD COLSPAN=2>
              <TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0>
                <TR>
                  <TD ALIGN=LEFT HEIGHT=24 WIDTH="50%">
                    <%
                    if Site_ID <> 100 then
                      response.write "<INPUT TYPE=""BUTTON"" NAME=""Password"" VALUE=""" & Translate("Forgot Password?",Login_Language,conn) & """ STYLE=""width:150;height:24;background-color:#FF0000;color:#FFFFFF;font-family:Arial;font-size:9pt"" TITLE=""Click to retrieve your password."" LANGUAGE=""JavaScript"" ONCLICK=""window.location.href='/register/default.asp?Site=" & Site_Code & "&Language=" & Login_Language & "&SubmitPassword=Password'"">"
                    end if  
                    %>
                  </TD>
                  <TD ALIGN=RIGHT HEIGHT=24 WIDTH="25%">
                    <INPUT TYPE="submit" NAME="Action" VALUE="OK" STYLE="width:76;height:24;background-color:#C2BFA5;color:#000000;font-family:Arial;font-size:9pt" TITLE="Submits your User ID and Password for verification and selected site access.">&nbsp;&nbsp;&nbsp;&nbsp;
                  </TD>
                  <TD ALIGN=RIGHT HEIGHT=24 WIDTH="25%">
                    <INPUT TYPE="submit" NAME="Action" VALUE="Cancel" STYLE="width:76;height:24;background-color:#C2BFA5;color:#000000;font-family:Arial;font-size:9pt" TITLE="Cancels Network Login Request.">&nbsp;&nbsp;
                  </TD>
                </TR>
              </TABLE>  
            </TD>  
          </TR>
  
        </TABLE>
      </TD>
    </TR>
  </TABLE>
  <%
  if not isblank(Site_ID) and Site_ID <> 100 then
    response.write "<BR>"
    if Site_ID = 24 then
      response.write "<A HREF=""/register/default.asp?Site=" & Site_Code & "&Language=" & Login_Language & "&SubmitRegister=Register"">" & "<SPAN Class=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Register for an Account",Login_Language,conn) & " - " & Translate("Click here",Login_Language,conn) & "&nbsp;&nbsp;</A></SPAN><P>"
    end if  
  end if
  
  ' Favorites Shortcut
  if Instr(1,UCase(request.ServerVariables("SERVER_NAME")),".DEV.") = 0 then
    response.write "<P><A HREF=""Javascript:void(0);"" LANGUAGE=""JavaScript"" ONCLICK=""window.external.AddFavorite('http://Support.Fluke.com/" & Site_Code & "','" & Translate(Site_Description,Login_Language,conn) & "');"">" & "<SPAN Class=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Add this Site to your Favorites List",Login_Language,conn)   & " - " & Translate("Click here",Login_Language,conn) & "&nbsp;&nbsp;</A></SPAN>"
  end if  
  response.write "<br>"
  if Instr(1,UCase(request.ServerVariables("SERVER_NAME")),".DEV.") = 0 then
    response.write "<P><A HREF=""https://support.fluke.com/register/register.asp?Site_ID=" & Site_ID & "&Account_ID=new&Language=" & Login_Language & """>" & "<SPAN Class=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("If you do not have an account",Login_Language,conn)   & " - " & Translate("Click here",Login_Language,conn) & "&nbsp;&nbsp;</A></SPAN>"
  end if  
  
  
  %>
  <SCRIPT TYPE="text/javascript" LANGUAGE="JavaScript">
  	document.FORM4.User_Name.focus();
    document.FORM4.Screen_Width.value  = screen.width;
    document.FORM4.Screen_Height.value = screen.height;
  </SCRIPT>
  <%
  response.write "</FORM>" & vbCrLf
  response.write "</DIV>"  & vbCrLf
  
  response.write "<BR><BR>" & vbCrLf

  %>
  <!--#include virtual="/SW-Common/SW-Footer.asp"-->
  <%

'  Login_Attempts = Login_Attempts + 1
  Session("Login_Attempts") = Login_Attempts
  
end if
  
Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' Check if the user is blocking popups and advise them to enable popups for this site.
' --------------------------------------------------------------------------------------

if CInt(Session("PopUpCheck")) <> CInt(true) then
  %>
  <script type="text/JavaScript" language="JavaScript">
  var popUpsBlocked = false;
  var mine = window.open('','mine','width=1,height=1,left=0,top=0,scrollbars=no');
  if(mine) {
    //mine.blur();
    mine.close();
  }
  else {
    var popUpsBlocked = true;
  }

  if (popUpsBlocked == true || <%=CInt(Session("PopUpCheck"))%> == -1) {
    alert("Browser Compatiblity Notice:\r\n\nThis site requires that you allow popup windows to open from this site in order to use its advanced capabilities and delivery methods for the information that you have requested to view.  click on [OK] to continue, then allow popup windows for this site to open.");
  }
  </script>
  <%
end if

Session("PopUpCheck") = true

%>
