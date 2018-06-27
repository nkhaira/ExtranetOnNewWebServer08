<%

' --------------------------------------------------------------------------------------
' Author:     K. D. Whitlock
' Date:       2/1/2000
'             Find it @ Support.Fluke.com
' --------------------------------------------------------------------------------------

Dim ShowTranslation
Dim ErrorString
Dim Reg_Freeze

Reg_Freeze = False

ShowTranslation = False

if Session("ShowTranslation") = True or request("Language") = "XON" then
  ShowTranslation = True
elseif Session("ShowTranslation") = False or request("Language") = "XOF" then
  ShowTranslation = False
end if  

'if not isblank(request.QueryString("ErrorString")) then
'  ErrorString = request("ErrorString")
if not isblank(Session("ErrorString")) then
  ErrorString = Session("ErrorString")
else
  ErrorString = ""
end if

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Dim Site
Dim Site_URL
Dim Site_ID
Dim Site_Found
Dim Site_Enabled

Site_ID = 0
keep_site = ""

if not isblank(request.querystring("Site")) then
  Site = request.querystring("Site")
elseif not isblank(request.form("Site")) then
  Site = request.form("Site")
else
  site = ""
end if

Session.abandon

if ShowTranslation = True then
  Session("ShowTranslation") = True
elseif ShowTranslation = False then
  Session("ShowTranslation") = False 
end if

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

response.buffer = true

if not isblank(Site) then
  
  if instr(1,lcase(Site),"http://") = 0 and instr(1,lcase(Site),"https://") = 0 then
    
    if not isblank(request("SubmitRegister")) or not isblank(request("SubmitLogin")) then
    
      ' Use text [site code] to look up site_id so casual user cannot decode scheme
           
      SQL = "SELECT * FROM Site WHERE Site.Site_Code='" & Site & "'"
      Set rsSite = Server.CreateObject("ADODB.Recordset")
      rsSite.Open SQL, conn, 3, 3
      
      if not rsSite.EOF then
        Site_ID           = rsSite("ID")
        Site_Enabled      = rsSite("Enabled")
        Site_URL          = rsSite("URL")
        Site_URL_Page     = rsSite("URL_Page")
        Site_Login_Method = rsSite("Login_Method")
        Site_Found        = True
      else
        Site_Found        = False
        if instr(1,lcase(Site),"http://") = 0 and instr(1,lcase(Site),"https://") = 0 then
          ErrorString     = ErrorString & "<LI>" & Translate("We are sorry, &quot;Find It @ Support.Fluke.com&quot; was unable to find site by the name of",Login_Language,conn) & ":&nbsp;&nbsp;&nbsp;<FONT COLOR=""Black"">" & Site & "</FONT></LI>"
          Site            = ""
        end if  
      end if
          
      rsSite.Close
      Set rsSite = nothing
        
      if Site_Found = True then
        response.write "<HTML>" & vbCrLf
        response.write "<HEAD>" & vbCrLf
        response.write "<TITLE></TITLE>" & vbCrLf
        response.write "</HEAD>" & vbCrLf
        response.write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
        
        if Site_Enabled = True then
          
          if not isblank(request("SubmitRegister")) then
            
            response.write "<FORM ACTION=""/register/register.asp"" METHOD=""POST"">" & vbCrLf
            response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
            response.write "<INPUT TYPE=""HIDDEN"" NAME=""Account_ID"" VALUE=""new"">" & vbCrLf
            response.write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
            
          elseif not isblank(request("SubmitLogin")) then
            
            if LCase(Site_Login_Method) = "db" then
              response.write "<FORM ACTION=""login.asp"" METHOD=""POST"">" & vbCrLf
            else
              if isblank(Site_URL_Page) then
                response.write "<FORM ACTION=""" & Site_URL & "/default.asp"" METHOD=""POST"">" & vbCrLf
              else                
                response.write "<FORM ACTION=""" & Site_URL & "/" & Site_URL_Page & """ METHOD=""POST"">" & vbCrLf
              end if                
            end if
            
            response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
            response.write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
            
          end if
          
        else
          if isblank(Site_URL_Page) then
            response.write "<FORM NAME=""redo"" ACTION=""" & Site_URL & "/default.asp"" METHOD=""POST"">" & vbCrLf
          else            
            response.write "<FORM ACTION=""" & Site_URL & "/" & Site_URL_Page & """ METHOD=""POST"">" & vbCrLf
          end if  
          response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
          response.write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
        end if            
        
        response.write "</FORM>" & vbCrLf
        response.write "</BODY>" & vbCrLf
        response.write "</HTML>" & vbCrLf
        
      end if
    else
      '    response.write "You need to either click [Logon] or [Register]"
      '    response.flush
      'response.write("would be redirecting<BR>")
      'response.redirect "/register/default.asp?ErrorString=" & "<LI>" & Translate("After you enter the name of the site, you must either click on [Logon] or [Register] to proceed.",login_Language,conn) & "</LI>"
      ErrorString = "<LI>" & Translate("After you enter the name of the site, you must either click on [Logon] or [Register] to proceed.",login_Language,conn) & "</LI>"
	  keep_site = site
	  site = ""
    end if    
  end if
  
elseif instr(1,lcase(Site),"http://") > 0 or instr(1,lcase(Site),"https://") > 0 then
  
  response.redirect Site
  
else
  Site = ""
end if    

' --------------------------------------------------------------------------------------
' Default Main
' --------------------------------------------------------------------------------------

'if isblank(Site) then

  Screen_Title =  Translate("Find it @ Support.fluke.com",Alt_Language,conn)  
  Bar_Title = "<TABLE BGCOLOR=""Black""><TR><TD VALIGN=TOP CLASS=Heading3Gold>" & Translate("Find it",Login_Language,conn) & "<BR>&nbsp;</TD><TD CLASS=Heading1White VALIGN=MIDDLE>@</TD><TD CLASS=Heading3Gold VALIGN=BOTTOM>&nbsp;<BR>Support.Fluke.com</TD></TR></TABLE>"
  Navigation      = false
  Side_Navigation = false
  Content_Width = 95  ' Percent
  
  %>
  <!--#include virtual="/sw-common/sw-header.asp"-->
  <!--#include virtual="/sw-common/sw-navigation.asp"-->
  <!--#include virtual="/connections/connection_Search_Engine.asp"-->    
  <%
  
  if not isblank(ErrorString) then
    response.write "<FONT CLASS=MediumRed>" & ErrorString & "</FONT><BR><BR>" & vbCrLf
    ErrorString = ""
    Session("ErrorString") = ""
  end if
  
  response.write "<FORM NAME=""Language"">" & vbCrLf    
  response.write "<TABLE WIDTH=""100%"" BORDER=1 BORDERCOLOR=""GRAY"" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>" & vbCrLf
  response.write "  <TR>" & vbCrLf
  response.write "    <TD WIDTH=""100%"" BGCOLOR=""#EEEEEE"">" & vbCrLf
 	response.write "      <TABLE WIDTH=""100%"" CELLPADDING=4 BORDER=0>" & vbCrLf
  response.write "        <TR>" & vbCrLf
  response.write "          <TD BGCOLOR=""Black"" VALIGN=MIDDLE Width=""2%"">&nbsp;"
'  response.write "          <FONT CLASS=MediumBoldGold>&nbsp;" & Translate("Select",Login_Language,conn) & "</FONT>"
  response.write "          </TD>" & vbCrLf
  response.write "         	<TD BGCOLOR=""#FFCC00"" VALIGN=MIDDLE CLASS=MediumBold WIDTH=""30%"">" & Translate("Preferred Language",Login_Language,conn) & ":</TD>"
  response.write "          <TD BGCOLOR=""White"" CLASS=MEDIUM WIDTH=""68%"">"
  response.write "            <SELECT NAME=""Language"" CLASS=Medium LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='default.asp?Language='+this.options[this.selectedIndex].value"">" & vbCrLf                      

  SQL = "SELECT Language.* FROM Language WHERE Language.Enable=" & CInt(True) & " ORDER BY Language.Sort"
  Set rsLanguage = Server.CreateObject("ADODB.Recordset")
  rsLanguage.Open SQL, conn, 3, 3
                        
  Do while not rsLanguage.EOF
    if LCase(rsLanguage("Code")) = LCase(Login_Language) then
   	  response.write "<OPTION SELECTED VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>" & vbCrLf             
    else
   	  response.write "<OPTION VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>" & vbCrLf             
    end if
	  rsLanguage.MoveNext 
  loop
  
  rsLanguage.close
  set rsLanguage=nothing

  response.write "            </SELECT>" & vbCrLf
  response.write "          </TD>" & vbCrLf
  response.write "        </TR>" & vbCrLf
  response.write "      </TABLE>" & vbCrLf
  response.write "    </TD>" & vbCrLf
  response.write "  </TR>" & vbCrLf
  response.write "</TABLE>" & vbCrLf
  response.write "</FORM>"    
      
  response.write "<FORM NAME=""WhereToGo"" ACTION=""default_pb.asp"" METHOD=""POST"">" & vbCrLf         
  response.write "<INPUT TYPE=""Hidden"" NAME=""Language"" VALUE=""" & Login_Language & """>"
  response.write "<TABLE WIDTH=""100%"" BORDER=1 BORDERCOLOR=""GRAY"" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>" & vbCrLf
  response.write "  <TR>" & vbCrLf
  response.write "    <TD WIDTH=""100%"" BGCOLOR=""#EEEEEE"">" & vbCrLf
 	response.write "      <TABLE WIDTH=""100%"" CELLPADDING=4 BORDER=0>" & vbCrLf
 	response.write "        <TR>" & vbCrLf
  response.write "          <TD BGCOLOR=""Black"" VALIGN=MIDDLE Width=""2%"">&nbsp;"
  response.write "          </TD>" & vbCrLf  
  response.write "          <TD BGCOLOR=""#FFCC00"" VALIGN=MIDDLE WIDTH=""30%"" CLASS=MediumBold>" & Translate("Name of the Site where you want to go",Login_Langugage,conn) & ":</TD>" & vbCrLf
  response.write "          <TD BGCOLOR=""White"" ALIGN=LEFT WIDTH=""68%"" VALIGN=MIDDLE CLASS=Medium>" & vbCrLf
  response.write "            <INPUT TYPE=""Text"" NAME=""Site"" SIZE=""35"" MAXLENGTH=""50"" VALUE=""" & keep_site & """ CLASS=Medium>"
  response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
  response.write "            <INPUT TYPE=""Submit"" NAME=""SubmitLogin"" VALUE="" " & Translate("Logon",Login_Language,conn) & " "" CLASS=NavLeftHighlight1>"
  
  if Reg_Freeze = False then
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;" & Translate("or",Login_language,conn) & "&nbsp;&nbsp;&nbsp;"
    response.write "            <INPUT TYPE=""Submit"" NAME=""SubmitRegister"" VALUE="" " & Translate("Register",Login_Language,conn) & " "" CLASS=NavLeftHighlight1>" & vbCrLf
  end if
  
  response.write "           </TD>" & vbCrLf
  response.write "        </TR>" & vbCrLf
  response.write "      </TABLE>" & vbCrLf
  response.write "    </TD>" & vbCrLf
  response.write "  </TR>" & vbCrLf
  response.write "</TABLE>" & vbCrLf
  response.write "</FORM>" & vbCrLf

  %> 
   <script language="Javascript">
  	document.WhereToGo.Site.focus();
  </script>
  <%
 
  response.write "<FORM ACTION=""" & Search_Engine & "/search/default.asp"" METHOD=""POST"">" & vbCrLf
  response.write "<INPUT TYPE=""Hidden"" NAME=""Language"" VALUE=""" & Login_Language & """>"
  response.write "  <TABLE WIDTH=""100%"" BORDER=1 BORDERCOLOR=""GRAY"" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>" & vbCrLf
  response.write "    <TR>" & vbCrLf
  response.write "      <TD WIDTH=""100%"" BGCOLOR=""#EEEEEE"">" & vbCrLf
 	response.write "        <TABLE WIDTH=""100%"" CELLPADDING=4 BORDER=0>" & vbCrLf
 	response.write "          <TR>" & vbCrLf
  response.write "            <TD BGCOLOR=""Black"" VALIGN=MIDDLE Width=""2%"">&nbsp;"
  response.write "            </TD>" & vbCrLf
  response.write "            <TD BGCOLOR=""#FFCC00"" VALIGN=MIDDLE WIDTH=""30%"" CLASS=MediumBold>" & Translate("Search www.Fluke.com",Login_Language,conn) & ":</TD>" & vbCrLf
  response.write "            <TD BGCOLOR=""White"" ALIGN=LEFT WIDTH=""68%"" VALIGN=MIDDLE>" & vbCrLf
  response.write "              <INPUT TYPE=""Text"" NAME=""SearchString"" SIZE=""35"" MAXLENGTH=""65"" VALUE="""" CLASS=Medium>&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Submit"" NAME=""Action"" VALUE="" " & Translate("Find It",Login_Language,conn) & " "" CLASS=NavLeftHighlight1>" & vbCrLf
  response.write "            </TD>" & vbCrLf
  response.write "          </TR>" & vbCrLf
  response.write "        </TABLE>" & vbCrLf
  response.write "      </TD>" & vbCrLf
  response.write "    </TR>" & vbCrLf
  response.write "  </TABLE>" & vbCrLf
  response.write "</FORM>" & vbCrLf 

  response.write "<FONT CLASS=Medium>" & vbCrLf
  response.write "<A HREF=""" & Search_Engine & "/search/searchhelp.asp"">" & Translate("Click here",Login_Language,conn) & "</A> " & Translate("for help and more information on how to use the www.Fluke.com site search feature.",Login_Language,conn) & vbCrLf
  response.write "</FONT>" & vbCrLf
  
  if Reg_Freeze = True then
    response.write "<BR><BR><FONT CLASS=NormalBold COLOR=Red>" & vbCrLf
    response.write "We are sorry, but we are currently performing a site data backup.&nbsp;&nbsp;No new registrations are being accepted at this time.&nbsp;&nbsp;Please visit this site again in a few hours to register for a new account.<BR><BR>This backup does not affect your ability to LOGON to the site if you already have an account." & vbCrLf
    response.write "</FONT>" & vbCrLf
  end if
    

  %>   
  <!--#include virtual="/SW-Common/SW-Footer.asp"-->  
  <%

'end if

response.flush

Call Disconnect_SiteWide

%>


