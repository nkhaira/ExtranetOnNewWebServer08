<%
' --------------------------------------------------------------------------------------
' Author:     K. D. Whitlock
' Date:       2/1/2000
'             SW-Footer.asp
' --------------------------------------------------------------------------------------      

if Content_Width <> "" then
  response.write "</TD></TR></TABLE></CENTER>" & vbCrLf
end if

response.write "<!-- End Contents -->" & vbCrLf

response.write "</FONT>" & vbCrLf
response.write "</TD>" & vbCrLf
response.write "</TR>" & vbCrLf

response.write "<!-- Begin Footer -->" & vbCrLf

response.write "<TR>" & vbCrLf
    
if Navigation = True then
  response.write "<TD>&nbsp;</TD>" & vbCrLf
end if
    
response.write "<TD COLSPAN=2 ALIGN=CENTER VALIGN=TOP>" & vbCrLf

response.write "<FONT CLASS=Small>"

response.write "&copy; 1995-" & DatePart("yyyy",Date) & " Fluke Corporation -  All rights reserved."

response.write "</FONT>" & vbCrLf
response.write "</TD>" & vbCrLf
response.write "</TR>" & vbCrLf

response.write "<!-- End Footer -->" & vbCrLf

response.write "</TABLE>" & vbCrLf

response.write "</BODY>" & vbCrLf
response.write "</HTML>" & vbCrLf

%>
