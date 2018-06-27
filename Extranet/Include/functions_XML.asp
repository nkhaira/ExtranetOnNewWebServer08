<%

dim objXMLHTTP

URL = Request("URL")

if ( URL = "" ) then

  URL = "http://www.fluke.com/products/manuals.asp?AppGrp_ID=29"
  URL = "http://www.seattlefirstbaptist.org/default.asp?locator=site_map"
end if 

Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")

objXMLHTTP.Open "GET", URL, false

objXMLHTTP.Send

myPage = objXMLHTTP.responseText

'response.write Mid(myPage,instr(1,myPage,"<BODY"))

Response.Write myPage

Set objXMLHTTP = Nothing
%>
