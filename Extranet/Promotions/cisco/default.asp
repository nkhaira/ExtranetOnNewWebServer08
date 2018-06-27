<%
'(Angel: Sometimes I will pass mscssid for auto logon so this needs to be added to your code see --->> VVVVV)

' --------------------------------------------------------------------------------------
' Basic URL Construction for new site ----------------------------------------------------------------->> VVVVVVVVVVVVVVV
' http://buy.fluke.com/estore/dCatalogs.asp?catalog=CISCO_ALL&gowhere=catalog&dCatalog_Password=_mofassa_&mscssid=Q8WR36E
'
' Test URL (Includes extra parameters to make it compatible with old site
' Use this to test redirect to eStore
' https://support.fluke.com/promotions/cisco/default.asp?catalog=CISCO_ALL&gowhere=catalog&dCatalog_Password=_mofassa_&pgm=cisco_a&cid=9999&lvl=1&mscssid=Q8WR36E
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' New eStore Site
' --------------------------------------------------------------------------------------

if LCase(request("pgm")) = "cisco_a" and request("cid") <> "" and request("lvl") <> "" then

  Call Page_Header

  response.write "<FORM ACTION=""http://buy.fluke.com/estore/dCatalogs.asp"" METHOD=""POST"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""catalog"" VALUE=""CISCO_ALL"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""gowhere"" VALUE=""catalog"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""dCatalog_Password"" VALUE=""_mofassa_a"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""mscssid=""" & request("mscssid") & """>"

  Call Page_Footer


elseif LCase(request("pgm")) = "cisco_s" and request("cid") <> "" and request("lvl") <> "" then

  Call Page_Header

  response.write "<FORM ACTION=""http://buy.fluke.com/estore/dCatalogs.asp"" METHOD=""POST"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""catalog"" VALUE=""CISCO_S"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""gowhere"" VALUE=""catalog"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""dCatalog_Password"" VALUE=""_mofassa_s"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""mscssid=""" & request("mscssid") & """>"

  Call Page_Footer

elseif LCase(request("pgm")) = "cisco_f" and request("cid") <> "" and request("lvl") <> "" then

  Call Page_Header

  response.write "<FORM ACTION=""http://buy.fluke.com/estore/dCatalogs.asp"" METHOD=""POST"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""catalog"" VALUE=""CISCO_FVDC"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""gowhere"" VALUE=""catalog"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""dCatalog_Password"" VALUE=""_mofassa_f"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""mscssid=""" & request("mscssid") & """>"

  Call Page_Footer

elseif LCase(request("pgm")) = "cisco_r" and request("cid") <> "" and request("lvl") <> "" then

  Call Page_Header

  response.write "<FORM ACTION=""http://www.flukenetworks.com/us/_Promotions/CiscoNA.htm"" METHOD=""POST"">"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""cid=""" & request("cid") & """>"
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""pgm=""" & request("pgm") & """>"
  
  Call Page_Footer


else

  response.redirect "http://buy.fluke.com/eStore/"

end if

' --------------------------------------------------------------------------------------

sub Page_Header

  response.write "<HTML>"
  response.write "<HEAD>"
  response.write "<TITLE></TITLE>"
  response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=iso-8859-1"">"
  response.write "</HEAD>"
  response.write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>"

end sub

' --------------------------------------------------------------------------------------

sub Page_Footer

  response.write "</FORM>"
  response.write "</BODY>"
  response.write "</HTML>"

end sub

' --------------------------------------------------------------------------------------
%>
