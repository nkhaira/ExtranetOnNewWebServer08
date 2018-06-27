<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_EMail.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%

' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       6/5/2003
'             Dev
' --------------------------------------------------------------------------------------

response.buffer = true

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

%>
<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
<!-- #include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

' Modify Shopping Cart Accessibility based on user's region or country exclusion
if Shopping_Cart = CInt(True) then
   
  select case Login_Region
    case 1  ' US
      Shopping_Cart = Shopping_Cart_R1
    case 2  ' Europe
      Shopping_Cart = Shopping_Cart_R2
    case 3  ' Intercon
      Shopping_Cart = Shopping_Cart_R3
  end select

  if instr(1,LCase(Shopping_Cart_Country),LCase(Login_Country)) > 0 then
    Shopping_Cart = 0
  end if

end if
'Modified by Zensar for Nvision.
'---------------------------------------
if     Price_Delivery = false and _
       Order_Inquiry  = True  then
       response.redirect "/SW-common/SW-Order_Inquiry_Form.asp"
elseif Price_Delivery = true  and _
       Order_Inquiry  = false then
       response.redirect "/SW-common/SW-Avail_Form.asp"  

'if     Shopping_Cart  = false and _
'       Price_Delivery = false and _
'       Order_Inquiry  = True  then
'       response.redirect "/SW-common/SW-Order_Inquiry_Form.asp"

'elseif Shopping_Cart  = true  and _
'       Price_Delivery = false and _
'       Order_Inquiry  = false then
'       response.redirect "/SW-common/SW-Order_Inquiry_Literature.asp"  
       
'elseif Shopping_Cart  = false and _
'       Price_Delivery = true  and _
'       Order_Inquiry  = false then
'       response.redirect "/SW-common/SW-Avail_Form.asp"  
'---------------------------------------------------------------------------
else

  Dim Top_Navigation        ' True / False
  Dim Side_Navigation       ' True / False
  Dim Screen_Title          ' Window Title
  Dim Bar_Title             ' Black Bar Title
  
  Screen_Title    = Site_Description & " - " & Translate("Order Inquiry - Select",Login_Language,conn)
  Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & _
                    Translate("Order Inquiry - Select",Login_Language,conn) & "</FONT>"
  Top_Navigation  = False 
  Side_Navigation = True
  Content_Width   = 95  ' Percent
  BackURL = Replace(Session("BackURL"),"CID=9007","CID=9000")
  
  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
  <%
  
  response.write "<SPAN CLASS=Heading3>" & Translate("Order Inquiry - Select",Login_Language,conn) & "</SPAN><BR>"
  response.write "<BR>"
  
  response.write "<SPAN CLASS=""Small"">" & vbCrLf

  Call Nav_Border_Begin
  response.write "<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=4 BGCOLOR=""White"">"

  ' Product Order Inquiry
  
  if CInt(Order_Inquiry) = CInt(true) then
    response.write "<TR>"
    response.write "<TD CLASS=NavLeftHighlight1 VALIGN=TOP BGCOLOR=""WHITE"" NOWRAP>"
    response.write "<A HREF=""/SW-Common/SW-Order_Inquiry_Form.asp"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Product Order Status",Login_Language,conn) & "&nbsp;&nbsp;</A>"
    response.write "</TD>"
    response.write "<TD CLASS=Small BGCOLOR=""White"">" & Translate("Look up the status of your product orders.",Login_language,conn) & "</TD>"
    response.write "</TR>"
  end if  

  ' Price and Delivery

  if CInt(Price_Delivery) = CInt(true) then
    response.write "<TR>"
    response.write "<TD CLASS=NavLeftHighlight1 VALIGN=TOP BGCOLOR=""WHITE"" NOWRAP>"
    response.write "<A HREF=""/SW-Common/SW-Avail_Form.asp"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Product Price and Availibility",Login_Language,conn) & "&nbsp;&nbsp;</A>"
    response.write "</TD>"
    response.write "<TD CLASS=Small BGCOLOR=""White"">" & Translate("Look up the List and Net Prices and Availibility of Products.",Login_Language,conn) & "</TD>"
    response.write "</TR>"
  end if

  ' Literature Order Inquiry
  'Modified by zensar on 13-05-2009 as the literature order status will not be shown through portal anymore.
  'if CInt(Shopping_Cart) = CInt(true) then
  '  response.write "<TR>"
  '  response.write "<TD CLASS=NavLeftHighlight1 VALIGN=TOP BGCOLOR=""WHITE"" NOWRAP>"
  '  response.write "<A HREF=""/SW-Common/SW-Order_Inquiry_Literature.asp?Sync=True"" CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;" & Translate("Literature Order Status",Login_Language,conn) & "&nbsp;&nbsp;</A>"
  '  response.write "</TD>"
  '  response.write "<TD CLASS=Small BGCOLOR=""White"">" & Translate("Look up the status of your literature orders.",Login_Language,conn) & "</TD>"
  '  response.write "</TR>"
  'end if
  '>>>>>>>>>>>>>>>>>
    
  response.write "</TABLE>"
  Call Nav_Border_End  
  response.write "</SPAN>" & vbCrLf
  %>
  <!--#include virtual="/SW-Common/SW-Footer.asp"-->
  <%

end if

Call Disconnect_SiteWide

response.end

'--------------------------------------------------------------------------------------
' Functions and Subroutines
'--------------------------------------------------------------------------------------

sub Table_Begin()
    response.write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" CLASS=TableBorder>" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif""            BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLr
    response.write "  <TR>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD VALIGN=""top"">" & vbCrLf
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
    response.write "  </TR>" & vbCrLf
    response.write "</TABLE>" & vbCrLf
end sub  

'--------------------------------------------------------------------------------------
%>