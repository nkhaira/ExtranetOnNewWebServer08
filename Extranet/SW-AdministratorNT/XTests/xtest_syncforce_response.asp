<%
  Response.ContentType = "text/xml"
  Response.Charset     = "utf-16"
  Response.BinaryWrite "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
  response.Binarywrite "<html>"
  response.Binarywrite "<head>"
  response.Binarywrite "<title>Untitled</title>"
  response.Binarywrite "</head>"

  response.Binarywrite "<body>"

  response.Binarywrite "<P>"
  response.Binarywrite "<TABLE Border=1>"

  response.Binarywrite "<TR>"
  response.Binarywrite "<TD BGCOLOR=Silver><FONT FACE=Arial SIZE=2>Key1</FONT></TD>"
  response.Binarywrite "<TD BGCOLOR=Silver><FONT FACE=Arial SIZE=2>Value</FONT></TD>"
  response.Binarywrite "<TD BGCOLOR=Silver ALIGN=CENTER><FONT FACE=Arial SIZE=2>Array</FONT></TD>"  
  response.Binarywrite "</TR>"
  
'  for each item in request.querystring

  for each item in request.form
    response.Binarywrite "<TR><TD><FONT FACE=Arial SIZE=2>" & item & "</FONT></TD>"
    response.Binarywrite "<TD><FONT FACE=Arial SIZE=2>"
    if LCase(item) = "action" then
      response.Binarywrite "<FONT COLOR=""RED""><B>" & request(item) & "</B></FONT>"
    else  
      response.Binarywrite request(item)
    end if  
    response.Binarywrite "</FONT></TD>"
    
    response.Binarywrite "<TD><FONT FACE=Arial SIZE=2>"
    
    Check4Array = Split(request(item),",")
    if UBound(Check4Array) > 0 then
      response.Binarywrite "Yes"
    else
      response.Binarywrite "&nbsp;"
    end if

    response.Binarywrite "</FONT></TD>"
    response.Binarywrite "</TR>"
  next

  response.Binarywrite "</TABLE>"
  response.Binarywrite "</body>"
  response.Binarywrite "</html>"
%>


