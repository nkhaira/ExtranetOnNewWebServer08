<%
' This code intercepts the response.querystring Key=Value pairs and reconstructs a redirect
' URL to be used to complete the Fluke TI20 training sequence at the conclusion of the off-site
' video.

BackURL = request.querystring("BackURL") & "?"

for each item in request.querystring
  if UCase(item) <> "BACKURL" then
    BackURL = BackURL & item & "=" & request.querystring(item) & "&"
  end if  
next

BackURL = Mid(BackURL, 1, Len(BackURL)-1)   ' Strip off last character "&"

response.write "<A HREF=""" & BackURL & """>End of Ti20 Video Training, Click to Continue</A>"
%>