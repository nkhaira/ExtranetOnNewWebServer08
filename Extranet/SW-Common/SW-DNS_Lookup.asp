<%@ Language="VBScript" %>
<%

if request("ipAddress") = "" then
  ' Set Initial Value to Local IP Address
  strIP = request.ServerVariables("REMOTE_ADDR")
else
  strIP = request("ipAddress") 
end if


strRemoteHostURL = "http://www.dnswatch.info/dns/dnslookup?la=en&host=" & strIP & "&submit=Resolve"
strPostString = ""
cMethod = "GET"

Set oHTTPComm = Server.CreateObject("Msxml2.SERVERXMLHTTP.6.0") 

' Open the connection
Call oHTTPComm.open(cMethod,strRemoteHostURL,0)   
Call oHTTPComm.setRequestHeader ("Content-Type", "application/x-www-form-urlencoded")
  
' Send the data over to the server.
Call oHTTPComm.send (strPostString)
  
' Get the response back (if any) from the other server. Check for errors.
strResponse = oHTTPComm.responseText

if oHTTPComm.status <> 200 then	
	strResponse = "<P><B>Errors:</B> " & oHTTPComm.status & "<P><B>Data Received:</B> " & strResponse & "<P>"
   set oHTTPComm = nothing
   HTTP_PostData = false
else
   strResponse = "200 OK<BR>" & strResponse
   set oHTTPComm = nothing
	HTTP_PostData = true
end if

response.write "<HTML>" & vbCrLf
response.write "<HEAD>" & vbCrLf
response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
response.write "<TITLE>Reverse IP Lookup</TITLE>" & vbCrLf
response.write "</HEAD>"
response.write "<BODY BGCOLOR=""White"" LINK =""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf

'response.write strResponse & "<P>"

if HTTP_PostData = true and instr(1,strResponse,"<table width=""500px"" class=""resultTable"" cellspacing=""10px"">") > 0 then
  
  sString = mid(strResponse,instr(1,strResponse,"<table width=""500px"" class=""resultTable"" cellspacing=""10px"">"))
  sString = mid(sString,1,instr(1,sString,"</table>") + 7)
  sString = replace(sString,"<table width=""500px"" class=""resultTable"" cellspacing=""10px"">","<table border=1 cellpadding=4>")
  sString = replace(sString,"resultTable","MediumBold")
  sString = replace(sString,"row1","Medium")
  sString = replace(sString,"<th","<td")
  sString = replace(sString,"</th>","</td>")
  sString = replace(sString,"Domain","Qualified IP Address")
  sString = replace(sString,"Answer","Domain Name")

  response.write "&nbsp;&nbsp;<SPAN CLASS=MediumBold>IP Address: " & strIP & "<P>"
  response.write sString
else
  response.write "Unable to resolve IP Address - No host found in addr.arpa"
end if

response.write "</BODY>"
response.write "</HTML>"

%>
