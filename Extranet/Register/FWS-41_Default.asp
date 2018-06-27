<%@  language="VBScript" codepage="65001" %>
<%
' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       2/1/2000
'             Find it @ Support.Fluke.com
' --------------------------------------------------------------------------------------
Dim ShowTranslation
Dim ErrorString
Dim Reg_Freeze
Dim strBackURL
Dim Border_Toggle
Border_Toggle = 0

strBackURL = request("BackURL")

Reg_Freeze = False

ShowTranslation = False

if Session("ShowTranslation") = True or request("Language") = "XON" then
  ShowTranslation = True
elseif Session("ShowTranslation") = False or request("Language") = "XOF" then
  ShowTranslation = False
end if  

if not isblank(request("ErrorString")) then
  ErrorString = request("ErrorString")
elseif not isblank(Session("ErrorString")) then
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
Dim Site_Alias  ' If not null, then replace site_id with site_alias
Dim Site_URL
Dim Site_ID
Dim Site_Found
Dim Site_Enabled

Site_ID = 0

Site_Select = request.form("Site_Select")
if isblank(Site_Select) then
	Site_Select = request.querystring("Site_Select")
else
    if LCase(Site_Select) = "met-support" then
             session.Abandon()
             response.Redirect "http://us.flukecal.com/support/my-met-support"
     end if 
end if

Site = request.form("Site")
if isblank(Site) then
	Site = request.querystring("Site")
end if

if not isblank(Site_Select) and isblank(Site) then
  if Site_Select <> "manually" then
    Site = Site_Select
  end if  
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

    if not isblank(request("SubmitRegister")) or not isblank(request("SubmitLogin")) or not isblank(request("SubmitPassword"))then
    
      ' Use text [site code] to look up site_id so casual user cannot decode scheme
           
      SQL = "SELECT * FROM Site WHERE Site.Site_Code='" & Site & "'"
      Set rsSite = Server.CreateObject("ADODB.Recordset")
      rsSite.Open SQL, conn, 3, 3
      
      if not rsSite.EOF then
        Site_ID           = rsSite("ID")
        Site_Alias        = rsSite("Site_Alias")
        if not isblank(Site_Alias) then Site_ID = Site_Alias
        Site_Enabled      = rsSite("Enabled")
        Site_URL          = rsSite("URL")
        Site_URL_Page     = rsSite("URL_Page")
        Site_Login_Method = rsSite("Login_Method")
        Site_Found        = True
      else
        Site_Found        = False
        if instr(1,lcase(Site),"http://") = 0 and instr(1,lcase(Site),"https://") = 0 then
          ErrorString     = ErrorString & "<LI>" & Translate("We are sorry, &quot;Find It @ " & Replace(Replace(ProperCase(Replace(request.ServerVariables("SERVER_NAME"),"."," "))," ","."),".Com",".com") & "&quot; was unable to find site by the name of",Login_Language,conn) & ":&nbsp;&nbsp;&nbsp;<FONT COLOR=""Black"">" & Site & "</FONT></LI>"
          Site            = ""
        end if  
      end if
          
      rsSite.Close
      Set rsSite = nothing
        
      if Site_Found = True then

        response.write "<HTML>" & vbCrLf
        response.write "<HEAD>" & vbCrLf
        response.write "<TITLE></TITLE>" & vbCrLf
        response.write "<META HTTP-EQUIV=""Edge-Control"" CONTENT=""no-store"">" & vbCrLf  
%>
<!-- Below code added for RI#1595 (Google code update) -->

<script type="text/javascript">
var _gaq = _gaq || [];
_gaq.push(['_setAccount', 'UA-3420170-1']);
_gaq.push(['_setDomainName', '.fluke.com']);
_gaq.push(['_setAllowLinker', true]);
_gaq.push(['_setAllowHash', false]);
_gaq.push(['_trackPageview']);
(function() {
var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
})();
</script>

<%      
        response.write "</HEAD>" & vbCrLf
        response.write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
   	
        if Site_Enabled = True or not isblank(Site_Alias) then

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

          elseif not isblank(request("SubmitPassword")) then
          
            response.write "<FORM ACTION=""/register/register.asp"" METHOD=""POST"">" & vbCrLf
            response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
            response.write "<INPUT TYPE=""HIDDEN"" NAME=""Account_ID"" VALUE=""password"">" & vbCrLf
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
      response.redirect "/register/default.asp?ErrorString=" & "<LI>" & Translate("After you enter the name of the site, you must either click on [Logon] or [Register] to proceed.",login_Language,conn) & "</LI>" & "&backurl=" & strBackURL
    end if    
  end if
    
elseif instr(1,lcase(Site),"http://") > 0 or instr(1,lcase(Site),"https://") > 0 then

  response.redirect Site & "?&BackURL=" & strBackURL
  
else
  Site = ""
end if    

' --------------------------------------------------------------------------------------
' Default Main
' --------------------------------------------------------------------------------------

if isblank(Site) then

  Screen_Title =  Translate("Find it",Alt_Language,conn) & " @ " & Replace(Replace(ProperCase(Replace(request.ServerVariables("SERVER_NAME"),"."," "))," ","."),".Com",".com")
  Bar_Title = "<FIND_IT_HEADER><TABLE BGCOLOR=""Black"" BORDER=0><TR><TD VALIGN=TOP CLASS=Heading3Gold>" & Translate("Find it",Login_Language,conn) & "<BR>&nbsp;</TD><TD CLASS=Heading1White VALIGN=MIDDLE>@</TD><TD CLASS=Heading3Gold VALIGN=BOTTOM>&nbsp;<BR>" & Replace(Replace(ProperCase(Replace(request.ServerVariables("SERVER_NAME"),"."," "))," ","."),".Com",".com") & "</TD></TR></TABLE></FIND_IT_HEADER>"
  Navigation      = false
  Side_Navigation = false
  Content_Width = 95  ' Percent

  Site_ID = 100
  if Site_ID = 100 and instr(1,LCase(request.ServerVariables("SERVER_NAME")),"flukenetworks") > 0 then
    Screen_Title =  Translate("Find it",Alt_Language,conn) & " @ " & Replace(Replace(ProperCase(Replace(request.ServerVariables("SERVER_NAME"),"."," "))," ","."),".Com",".com")
    Bar_Title = "<FIND_IT_HEADER><TABLE BGCOLOR=""#013567"" BORDER=0><TR><TD VALIGN=TOP CLASS=Heading3Gold>" & Translate("Find it",Login_Language,conn) & "<BR>&nbsp;</TD><TD CLASS=Heading1White VALIGN=MIDDLE>@</TD><TD CLASS=Heading3Gold VALIGN=BOTTOM>&nbsp;<BR>" & Replace(Replace(ProperCase(Replace(request.ServerVariables("SERVER_NAME"),"."," "))," ","."),".Com",".com") & "</TD></TR></TABLE></FIND_IT_HEADER>"
  end if
  
%>
<!--#include virtual="/sw-common/sw-header.asp"-->
<!--#include virtual="/sw-common/sw-navigation.asp"-->
<!--#include virtual="/connections/connection_Search_Engine.asp"-->
<style>
  .ZeroValue  {font-size:8.5pt;font-weight:Bold;color:Black;background:#FFFF99;text-decoration:none;font-family:Arial,Verdana;}
  
  </style>
<%
  
  if not isblank(ErrorString) then
    response.write "<FONT CLASS=MediumRed>" & ErrorString & "</FONT><BR><BR>" & vbCrLf
    ErrorString = ""
    Session("ErrorString") = ""
  end if
  
  response.write "<FORM NAME=""Language"">" & vbCrLf    

  Call Table_Begin
 	response.write "      <TABLE WIDTH=""100%"" CELLPADDING=2 BORDER=0 BGCOLOR=""#666666"">" & vbCrLf
  response.write "        <TR>" & vbCrLf
  response.write "          <TD BGCOLOR=""Black"" VALIGN=MIDDLE Width=""2%"">&nbsp;"
  response.write "          </TD>" & vbCrLf
  response.write "         	<TD BGCOLOR=""#FFCC00"" NOWRAP VALIGN=MIDDLE CLASS=SmallBold WIDTH=""30%"">&nbsp;" & Translate("Preferred Language",Login_Language,conn) & ":</TD>"
  response.write "          <TD BGCOLOR=""#666666"" CLASS=Small WIDTH=""68%"">"
  response.write "            <SELECT NAME=""Language""  CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='default.asp?Language='+this.options[this.selectedIndex].value"">" & vbCrLf                      

  SQL = "SELECT Language.* FROM Language WHERE Language.Enable=" & CInt(True) & " ORDER BY Language.Sort"
  Set rsLanguage = Server.CreateObject("ADODB.Recordset")
  rsLanguage.Open SQL, conn, 3, 3
                        
  Do while not rsLanguage.EOF
    if LCase(rsLanguage("Code")) = LCase(Login_Language) then
   	  response.write "<OPTION SELECTED VALUE=""" & rsLanguage("Code") & """>" &  Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>" & vbCrLf             
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
  Call Table_End

  response.write "</FORM>"    
    
  response.write "<FORM NAME=""WhereToGo"" ACTION=""default.asp"" METHOD=""POST"">" & vbCrLf         
  response.write "<INPUT TYPE=""Hidden"" NAME=""Language"" VALUE=""" & Login_Language & """>"

  Call Table_Begin
 	response.write "      <TABLE WIDTH=""100%"" CELLPADDING=2 BORDER=0 BGCOLOR=""#666666"">" & vbCrLf
 	response.write "        <TR>" & vbCrLf
  response.write "          <TD BGCOLOR=""Black"" VALIGN=MIDDLE Width=""2%"">&nbsp;"
  response.write "          </TD>" & vbCrLf  
  response.write "          <TD BGCOLOR=""#FFCC00"" NOWRAP VALIGN=MIDDLE WIDTH=""30%"" CLASS=SmallBold>&nbsp;" & Translate("Name of the Site where you want to go",Login_Langugage,conn) & ":</TD>" & vbCrLf
  response.write "          <TD BGCOLOR=""#666666"" ALIGN=LEFT WIDTH=""68%"" VALIGN=MIDDLE CLASS=Small>" & vbCrLf


  if Site_Select <> "manually" then
    SQL = "SELECT Site_Code AS Site_Code, Site_Description " &_
          "FROM Site " &_
          "WHERE NoShow=" & CInt(False) & " " &_
          "ORDER BY Site_Description"

    Set rsSite = Server.CreateObject("ADODB.Recordset")
    rsSite.Open SQL, conn, 3, 3
    response.write "<SELECT CLASS=Small NAME=""Site_Select"" ONCHANGE=""CkManual();"">" & vbCrLf
    response.write "<OPTION Class=Small VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
    do while not rsSite.EOF

      response.write "<OPTION Class=Small VALUE=""" & rsSite("Site_Code") & """>" & Translate(rsSite("Site_Description"),Login_Language,conn) & "</OPTION>" & vbCrLf


      rsSite.MoveNext
    loop
    response.write "</OPTION>"
    response.write "<OPTION Class=ZeroValue VALUE=""manually"">" & Translate("Enter Site Name Manually",Login_Language,conn) & "" & "</OPTION>" & vbCrLf
  
    rsSite.close
    set rsSite = nothing
    
    response.write "</SELECT>" & vbCrLf
    
  else  
    response.write "<INPUT TYPE=""Text"" NAME=""Site"" SIZE=""15"" MAXLENGTH=""50"" VALUE="""" CLASS=Medium>"
  end if
      
  response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
  response.write "            <INPUT TYPE=""Submit"" ID=""SubmitLogin"" NAME=""SubmitLogin"" VALUE="" " & Translate("Logon",Login_Language,conn) & " "" CLASS=NavLeftHighlight1>"
  
  if Reg_Freeze = False then
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("or",Login_language,conn) & "</SPAN>&nbsp;&nbsp;&nbsp;"
    response.write "            <INPUT TYPE=""Submit"" ID=""SubmitRegister"" NAME=""SubmitRegister"" VALUE="" " & Translate("Register",Login_Language,conn) & " "" CLASS=NavLeftHighlight1>" & vbCrLf
  end if
  
  response.write "           </TD>" & vbCrLf
  response.write "        </TR>" & vbCrLf
  response.write "      </TABLE>" & vbCrLf
  Call Table_End

  response.write "</FORM>" & vbCrLf

 	response.write "        <TABLE WIDTH=""100%"" CELLPADDING=2 BORDER=0 BGCOLOR=""White"">" & vbCrLf
 	response.write "          <TR>" & vbCrLf
  response.write "<TR><TD COLSPAN=3 BGCOLOR=WHITE>&nbsp;</TD></TR>"
  response.write "           </TD>" & vbCrLf
  response.write "        </TR>" & vbCrLf
  response.write "      </TABLE>" & vbCrLf
  
  response.write "<script language=""Javascript"">"
  if Site_Select = "manually" then
  	response.write "document.WhereToGo.Site.focus();"
  else
    response.write "document.WhereToGo.Site_Select.focus();"
  	end if
  
  response.write "</script>"
 
  if Reg_Freeze = True then
    response.write "<BR><BR><FONT CLASS=NormalBold COLOR=Red>" & vbCrLf
    response.write "We are sorry, but we are currently performing a site data backup.&nbsp;&nbsp;No new registrations are being accepted at this time.&nbsp;&nbsp;Please visit this site again in a few hours to register for a new account.<BR><BR>This backup does not affect your ability to LOGON to the site if you already have an account." & vbCrLf
    response.write "</FONT>" & vbCrLf
  end if
    
%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

end if

response.flush

Call Disconnect_SiteWide

'--------------------------------------------------------------------------------------
' Subroutines and Functions
'--------------------------------------------------------------------------------------

sub Table_Begin()
    response.write "<TABLE BORDER=""" & Border_Toggle & """ CELLPADDING=""0"" CELLSPACING=""0"" VSPACE=""0"" HSPACE=""0"" BGCOLOR=#666666>" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif""            BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLr
    response.write "  <TR>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD VALIGN=""top"" WIDTH=""100%"">" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Table_End()
    response.write "    </TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_BL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_BR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>"
    response.write "</TABLE>" & vbCrLf
end sub
%>

<script language="JavaScript">
  function CkManual() {
    if (document.WhereToGo.Site_Select.value == "manually") {
      window.location.href="/register/default.asp?Site_Select=manually";
    }
	
	else if(document.WhereToGo.Site_Select.value.toLowerCase() == "met-support"){	        
            // document.getElementById('SubmitLogin').disabled = 'true';
            // document.getElementById('SubmitRegister').disabled = 'true';
            document.getElementById('SubmitLogin').setAttribute('disabled', 'disabled');
            document.getElementById('SubmitLogin').style.backgroundColor ='#dddddd';
            document.getElementById('SubmitRegister').setAttribute('disabled', 'disabled');
            document.getElementById('SubmitRegister').style.backgroundColor ='#dddddd';
            window.location.href="http://us.flukecal.com/support/my-met-support";
    }

  }
</script>

