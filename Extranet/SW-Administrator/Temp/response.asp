<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>

<P>
<%

  response.write "<TABLE Border=1>"

  response.write "<TR>"
  response.write "<TD BGCOLOR=Silver><FONT FACE=Arial SIZE=2>Key1</FONT></TD>"
  response.write "<TD BGCOLOR=Silver><FONT FACE=Arial SIZE=2>Value</FONT></TD>"
  response.write "<TD BGCOLOR=Silver ALIGN=CENTER><FONT FACE=Arial SIZE=2>Array</FONT></TD>"  
  response.write "</TR>"
  
'  for each item in request.querystring

  for each item in request.form
    response.write "<TR><TD><FONT FACE=Arial SIZE=2>" & item & "</FONT></TD>"
    response.write "<TD><FONT FACE=Arial SIZE=2>"
    if LCase(item) = "action" then
      response.write "<FONT COLOR=""RED""><B>" & request(item) & "</B></FONT>"
    else  
      response.write request(item)
    end if  
    response.write "</FONT></TD>"
    
    response.write "<TD><FONT FACE=Arial SIZE=2>"
    
    Check4Array = Split(request(item),",")
    if UBound(Check4Array) > 0 then
      response.write "Yes"
    else
      response.write "&nbsp;"
    end if
    response.write "</FONT></TD>"
    response.write "</TR>"
  next
  response.write "</TABLE>"

%>

</body>
</html>
