<%
' This code intercepts the response.querystring Key=Value pairs and reconstructs a redirect
' URL to be used to complete the Fluke TI20 training sequence at the conclusion of the off-site
' video.


' for each item in request.querystring
'  response.write item & "|" & request.querystring(item) & "<BR>"
' next

if 1=1 then
  response.write "<HTML>" & vbCrLf
  response.write "<HEAD>" & vbCrLf
  response.write "<TITLE>Video Redirect</TITLE>" & vbCrLf
  response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8"">" & vbCrLf
  response.write "</HEAD>" & vbCrLf
  response.write "<BODY BGCOLOR=""White"" onLoad='document.FORM0.submit()'>" & vbCrLf
  
  response.write "<FORM NAME=""FORM0"" ACTION=""/sales-training/default.asp"" METHOD=POST>" & vbCrLf
  
  for each item in request.querystring
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""" & item & """ VALUE=""" & Replace(Replace(request.querystring(item),"http://",""),"HTTP://","") & """>" & vbCrLf
  next
  
  response.write "</FORM>" & vbCrLf
  response.write "</BODY>" & vbCrLf
  response.write "</HTML>" & vbCrLf
  
else

  BackURL = request.querystring("BackURL") & "?"
  
  for each item in request.querystring
    if UCase(item) <> "BACKURL" then
      BackURL = BackURL & item & "=" & request.querystring(item) & "&"
    end if  
  next
  
  BackURL = Mid(BackURL, 1, Len(BackURL)-1)   ' Strip off last character "&"

  response.write "<HTML>" & vbCrLf
  response.write "<HEAD>" & vbCrLf
  response.write "<TITLE>TI20 Video Redirect</TITLE>" & vbCrLf
  response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8"">" & vbCrLf
  response.write "</HEAD>" & vbCrLf
  
  response.write "<BODY BGCOLOR=""White"" onLoad='document.FORM0.submit()'>" & vbCrLf
  
  response.write "<FORM NAME=""FORM0"" ACTION=""http://fluke.stepframe.net/ti20vid/index.php"" METHOD=POST>" & vbCrLf
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""BackURL"" VALUE=""" & Replace(Replace(BackURL,"http://",""),"HTTP://","") & """>" & vbCrLf
  
  response.write "</FORM>" & vbCrLf
  response.write "</BODY>" & vbCrLf
  response.write "</HTML>" & vbCrLf

end if

%>