<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author: Kelly Whitlock
' Date:   2/1/2000
' Title:  Register Thank-You
' --------------------------------------------------------------------------------------
%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%

Call Connect_SiteWide

Dim Site_ID

Dim Promotion_Region
Dim Promotion_Text
Dim Promotion_Image
Dim Promotion_Button
Dim Promotion_URL
Dim Reg_Complete_URL

if not isblank(request("Region")) then
  Promotion_Region = CInt(request("Region"))
else
  Promotion_Region = 0
end if
    
Promotion_Text   = ""
Promotion_Image  = ""
Promotion_Button = "Continue"
Promotion_URL    = ""
Promotion_Complete_URL = "/Default.asp"

SQL = "SELECT * FROM Site WHERE ID=" & CInt(Site_ID)
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code         = rsSite("Site_Code")
Site_Description  = rsSite("Site_Description")
Logo              = rsSite("Logo")
Footer_Disabled   = rsSite("Footer_Disabled")
Business          = CInt(rsSite("Business"))
Contrast          = rsSite("Contrast")
'  if CInt(rsSite("Business")) = -1 then Business = True else Business = False
Privacy_Statement = rsSite("Privacy_Statement_Link")
  
if not isblank(request("Site_ID")) then
  Site_ID = request("Site_ID")
else
  Site_ID = 0
  Site_Code = ""
end if

Navigation      = false
Side_Navigation = false
Content_Width = 95  ' Percent


if Site_ID > 0 then

  SQL = "SELECT * FROM Site WHERE ID=" & CInt(Site_ID)
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3

  Site_Code           = rsSite("Site_Code")
  Site_Description    = rsSite("Site_Description")
  Logo                = rsSite("Logo")
  Logo_Left         = rsSite("Logo_Left")  
' Site_Reg_Fields     = rsSite("Reg_Fields")
  Footer_Disabled     = rsSite("Footer_Disabled")
  
  rsSite.close
  set rsSite=nothing
  
  SQL       = "SELECT * FROM Promotions "
  SQL = SQL & "WHERE Site_ID=" & CInt(Site_ID) & " "
  SQL = SQL & "AND  Region=" & Promotion_Region & " "
  SQL = SQL & "AND (Promotion_Begin<='" & Date & "' "
  SQL = SQL & "AND  Promotion_End  >='" & Date & "')"
  Set rsPromotion = Server.CreateObject("ADODB.Recordset")
  rsPromotion.Open SQL, conn, 3, 3

  Promotion_Flag = False

  if not rsPromotion.EOF then

    Promotion_Flag = True

    if not isblank(rsPromotion("Reg_Fields")) then    
      Site_Reg_Fields  = rsPromotion("Reg_Fields")
    end if      
    if not isblank(rsPromotion("Promotion_Text")) then
      Promotion_Text   = rsPromotion("Promotion_Text")
    end if
    if not isblank(rsPromotion("Promotion_Image")) then
      Promotion_Image  = rsPromotion("Promotion_Image")
    end if    
    if not isblank(rsPromotion("Promotion_URL")) then
      Promotion_URL    = Replace(LCase(rsPromotion("Promotion_URL")), "support.fluke.com", LCase(Request("SERVER_NAME")))
'     Promotion_URL    = rsPromotion("Promotion_URL")
    end if
    if not isblank(rsPromotion("Promotion_Button")) then
      Promotion_Button = rsPromotion("Promotion_Button")
    end if        
    if not isblank(rsPromotion("Promotion_Begin")) then
      Promotion_Begin = rsPromotion("Promotion_Begin")
    end if   
    if not isblank(rsPromotion("Promotion_End")) then
      Promotion_End = rsPromotion("Promotion_End")
    end if    
    if not isblank(rsPromotion("Promotion_Complete_URL")) then
      Promotion_Complete_URL = Replace(LCase(rsPromotion("Promotion_Complete_URL")), "support.fluke.com", LCase(Request("SERVER_NAME")))
'     Promotion_Complete_URL = rsPromotion("Promotion_Complete_URL")
    else
      Promotion_Complete_URL = "/Default.asp"
    end if  
    
    rsPromotion.close
    set rsPromotion=nothing
  end if  

end if        
  
Screen_Title = Translate(Site_Description,Alt_Language,conn) &  " - " & Translate("Thank You",Alt_Language,conn)
Bar_Title = Translate(Site_Description,Login_Language,conn) &  "<BR><FONT CLASS=MediumBoldGold>" & Translate("Thank You",Login_Language,conn) & "</FONT>"
Navigation      = false
Top_Navigation  = false
Side_Navigation = false

Content_Width = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<%

with response
  .write "<TABLE BORDER=0 WIDTH=""100%"" CELLPADDING=0 CELLSPACING=0>" & vbCrLf
  .write "<TR>"
  .write "<TD WIDTH=""100%"" HEIGHT=6 CLASS=TopColorBar><IMG SRC=""/images/1x1trans.gif"" HEIGHT=6 BORDER=0 VSPACE=0></TD>" & vbCrLf
  .write "</TR>"
  .write "<TR>" & vbCrLf
  .write "<TD CLASS=SMALL WIDTH=""100%"">"    & vbCrLf
  .write "<DIV ALIGN=CENTER>" & vbCrLf
  .write "<TABLE BORDER=0 WIDTH=""" & Content_Width & "%"">" & vbCrLf
  .write "<TR>" & vbCrLf
  .write "<TD WIDTH=""100%"" CLASS=Medium>" & vbCrLf
  .write "<BR><BR>" 
  .write "<UL>" & vbCrLf
  .write "<LI>" & Translate("Your New Account Request has been submitted.",Login_Language,conn) & "</LI>" & vbCrLf
  .write "<LI>" & Translate("You will be notified by email, within 24 to 48-hours, if your request has been approved.",Login_Language,conn) & "</LI>" & vbCrLf
  .write "<LI>" & Translate("You will also receive complete instructions for accessing the site:",Login_Language,conn) & " " & "<B>" & Translate(Site_Description,Login_Language,conn) & "</B></LI>" & vbCrLf
end with
  
if Promotion_Flag = True _
  and not isblank(Promotion_URL)  _
  and not isblank(Promotion_Text) _
  and (isdate(Promotion_Begin) and CDate(Promotion_Begin) <= CDate(Date)) _
  and (isdate(Promotion_End)   and CDate(Promotion_End)   >= CDate(Date)) _
  then

  response.write "<BR><BR><LI>" & Translate(Promotion_Text,Login_Language,conn) & "</LI>" & vbCrLf
  
  with response
    .write "<BR><BR>"
    .write "<INDENT>"
  
    .write "<FORM ACTION=""/register/Register_Promo.asp"" METHOD=""Post"">" & VbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & VbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_Code"" VALUE=""" & Site_Code & """>" & VbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_Description"" VALUE=""" & Site_Description & """>" & VbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_Reg_Fields"" VALUE=""" & Site_Reg_Fields & """>" & VbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Promotion_URL"" VALUE=""" & Promotion_URL & """>" & VbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Promotion_Complete_URL"" VALUE=""" & Promotion_Complete_URL & """>" & VbCrLf  
    .write "<INPUT TYPE=""Submit"" VALUE=""" & Translate(Promotion_Button,Login_Language,conn) & """ CLASS=NavLeftHighlight1>" & VbCrLf    
  end with

  if not isblank(Promotion_Image) then
    with response
      .write "<BR><BR>" & vbCrLf
      .write "<INDENT>" & vbCrLf
      .write "<IMG SRC=""" & Promotion_Image & """ Border=0 WIDTH=220>" & vbCrLf
      .write "</INDENT>" & vbCrLf
    end with
  end if  
 
else
  with response
    .write "<FORM ACTION=""" & Promotion_Complete_URL & """ METHOD=""POST"">" & VbCrLf
    .write "<INPUT TYPE=""Submit"" VALUE=""" & Translate("Continue",Login_Language,conn) & """ CLASS=NavLeftHighlight1>" & VbCrLf
  end with
end if

with response
  .write "</FORM>" & VbCrLf
  .write "</INDENT>" & VbCrLf
  .write "</UL>" & VbCrLf
end with

%>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<% Call Disconnect_SiteWide %>
