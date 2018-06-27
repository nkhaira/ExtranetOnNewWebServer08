<%
Renew_Days = 3
expiration = "10/6/2006"

if CDate(DateAdd("d",CInt(Renew_Days),Date())) > CDate(expiration) then
  response.write "Y " & CDate(DateAdd("d",CInt(Renew_Days),Date()))
else  
  response.write "N " & CDate(DateAdd("d",CInt(Renew_Days),Date()))
end if  
%>