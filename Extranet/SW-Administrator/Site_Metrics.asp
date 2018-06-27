<%
Session("LOGON_USER") = ""
Session("Password")   = ""
Session("Site_ID")    = 0
Session("Language")   = ""

Site_ID          = 101
Site_Code        = "metrics"
Screen_Title     = "Site Metrics - Visitor/Asset Tracking Report Selection"
Bar_Title        = "Site Metrics<BR><FONT CLASS=SmallBoldGold>Visitor/Asset Tracking Report Selection</FONT>"
Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
Call Connect_SiteWide
%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<%

select case request("Utility_ID")
  case 70, 71, 72
    Session("LOGON_USER") = "Fluke User"
    Session("Password")   = "metrics"
    Session("Site_ID")    = 101
    Session("Language")   = "eng"
    response.redirect "/sw-administrator/site_utility.asp?Utility_ID=" & request("Utility_ID")
  case else
    response.write "<FORM ACTION=""/sw-administrator/site_metrics.asp"" METHOD=POST>" & vbCrLf
    response.write "<SELECT NAME=""Utility_ID"" CLASS=Small>"
    response.write "<OPTION VALUE=""72"">WWW Asset Activity Detail</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""70"">Partner Portal - Asset Activity Detail</OPTION>" & vbCrLf
    response.write "<OPTION VALUE=""71"">Partner Partal - Site Activity Summary</OPTION>" & vbCrLf
    response.write "</SELECT>"
    response.write "<INPUT TYPE=""SUBMIT"" NAME=""Submit"" VALUE="" GO "" CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & vbCrLf
    response.write "<INPUT TYPE=""BUTTON"" NAME=""Home"" VALUE="" Home "" ONCLICK=""window.location.href='http://evtibg08.tc.fluke.com/webpros/default.asp';"" CLASS=NavLeftHighlight1>" & vbCrLf
    response.write "</FORM>"
end select

response.write "<DIV CLASS=Small>"
response.write "<UL>" & vbCrLf
response.write "<LI><B>WWW Asset Activity Detail</B> - Provides a summary of <B><U>document only</U></B> asset requests originating from www.Fluke.com utilizing the Electronic File Fulfillment services of Support.Fluke.com.</LI>"
response.write "<LI><B>Partner Portal - Asset Activity Detail</B> - Provides a summary of <B><U>all asset</U></B> requests from the various Partner Portal Extranet Sites utilizing the Electronic File Fulfillment services of Support.Fluke.com.  To view individual Partner Portal metrics by site, you must have Site or Content Administration privelages at Support.Fluke.com and logon to the <A HREF=""/sw-administrator"">Administrator's Took Kit</A> to view individual site metrics reports.</LI>"
response.write "<LI><B>Partner Portal - Site Activity Summary</B> - Provides a summary of <B><U>all activity</U></B> requests (Order Inquiry, EEF, Navigation Clicks, Asset Requests, Account Status, Visitor Status, etc.) from the various Partner Portal Extranet Sites.  To view individual Partner Portal metrics by site, you must have Site or Content Administration privelages at Support.Fluke.com and logon to the <A HREF=""/sw-administrator"">Administrator's Took Kit</A> to view individual site metrics reports.</LI>"
%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%
Call Disconnect_SiteWide
%> 