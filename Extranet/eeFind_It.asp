<!--#include virtual="/include/functions_string.asp"-->
<%
' --------------------------------------------------------------------------------------
' Author:     Kelly. Whitlock
' Date:       7/22/2005
' Name:       Asset File Find It
' Purpose:    Preformatter for /SW-Common/SW-Find_It.asp
' --------------------------------------------------------------------------------------
' 
' Find_It.asp?Locator=[#####0#####0#####0]
' Find_It.asp?Document=[7-Digit Oracle Item Number of User Viewable PDF document]
' Find_It.asp?Document=[5-Digit Information Store Item Number (Language code defaults to "eng") of User Viewable PDF document]
' Find_It.asp?Document=[5-Digit Information Store Item Number + "-" + 3-Digit Iso3 Language code of User Viewable PDF document]
'
' Method [Document] can use the following Key=Values
'   Docuemnt      7-Digit Oracle Item Number or Generic PP Item Number
'   Style         Path to Style Sheet
'   Verify        on/off
'   CMS_Site      usen
'   CMS_Path      /dmm/application_notes/...
'   SRC           Site Code
'   AID           Access Identifier (Match to Calendar.Subgroups Array)
'   LAN           2 or 3 Digit ISO Code
' --------------------------------------------------------------------------------------

if len(request("Document")) = 5 or len(request("Document")) = 9 then  ' Information Store

  select case len(request("Document"))
  
    case 5
      if isnumeric(mid(request("Document"),1,5)) then
        response.redirect "http://www.informationstore.net/fluke/efulfillment.asp?publication=" & request("Document")
      end if
    case 9
      if isnumeric(mid(request("Document"),1,5)) and mid(request("Document"),6,1) = "-" and not isnumeric(mid(request("Document"),7,3)) then
        response.redirect "http://www.informationstore.net/fluke/efulfillment.asp?publication=" & request("Document")
      end if
  end select

else
  
  response.write "<HTML>" & vbCrLf
  response.write "<HEAD>" & vbCrLf
  response.write "<TITLE>Find_It</TITLE>" & vbCrLf
  response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=iso-8859-1"">" & vbCrLf
  response.write "</HEAD>" & vbCrLf
  response.write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
  response.write "<FORM NAME=""FORM1"" ACTION=""/SW-Common/SW-eeFind_It.asp"" METHOD=""POST"">" & vbCrLf
  
  if request("Locator") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Locator"" VALUE=""" & request("Locator") & """>" & vbCrLf
  end if
  
  if request("SW-Locator") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""SW-Locator"" VALUE=""" & request("SW-Locator") & """>" & vbCrLf
  end if
  
  if request("Document") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Document"" VALUE=""" & request("Document") & """>" & vbCrLf
  end if
  
  if request("Style") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Style"" VALUE=""" & request("Style") & """>" & vbCrLf
  end if
  
  if request("Verify") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Verify"" VALUE=""" & request("Verify") & """>" & vbCrLf
  end if
  
  if request("CMS_Site") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""CMS_Site"" VALUE=""" & request("CMS_Site") & """>" & vbCrLf
  end if
  
  if request("CMS_Path") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""CMS_Path"" VALUE=""" & request("CMS_Path") & """>" & vbCrLf
  end if
  
  if request("SRC") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""SRC"" VALUE=""" & request("SRC") & """>" & vbCrLf
  end if
  
  if request("Debug") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Debug"" VALUE=""" & request("Debug") & """>" & vbCrLf
  end if
    
  ' Access Identifier
  if request("AID") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""AID"" VALUE=""" & request("AID") & """>" & vbCrLf
  end if
  
  ' Language 2 or 3 Digit ISO Code
  if request("LAN") <> "" then
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""LAN"" VALUE=""" & request("LAN") & """>" & vbCrLf
  end if

  LastSite = request.ServerVariables("HTTP_REFERER")
  
  if isblank(LastSite) then
    LastSite = "[Unknown HTTP_Referer]"
  end if
  
  if not isblank(request("Document")) then
    LastSite = "URL Link from: " & Server.URLEncode(LastSite & "?" & request.QueryString)
  elseif not isblank(request("Locator")) then
    LastSite = "Partner Portal Subscription URL Link: " & Server.URLEncode("http://Support.Fluke.com/eeFind_It.asp?Locator=" & request("Locator"))
  elseif not isblank(request("SW-Locator")) then
    LastSite = "Partner Portal Asset URL Link: " & Server.URLEncode("http://Support.Fluke.com/eeFind_It.asp?SW-Locator=" & request("SW-Locator"))
  else
    LastSite = "Unknown Method 1 by using URL Link: " & Server.URLEncode(LastSite & "?" & request.QueryString)
  end if
  
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Referer"" VALUE=""" & LastSite & """>" & vbCrLf
  
  response.write "</FORM>" & vbCrLf
  response.write "</BODY>" & vbCrLf
  response.write "</HTML>" & vbCrLf

end if
%>