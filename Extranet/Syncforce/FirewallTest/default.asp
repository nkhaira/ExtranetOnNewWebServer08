<!-- #include virtual="/include/HTTP_DataTransfer.asp" -->
<%
Dim strErrorMessage
Dim strPost_QueryString

strPost_QueryString = "foo1:foo1value,foo2:foo2value,foo3:foo3value"

response.write("Posting data....<BR><BR>")

Result = HTTP_PostData(strPost_QueryString, ":", ",", HTTP_HTTP, HTTP_POST, "216.9.4.31", "82", "129.196.132.59", "syncforce", "/syncforce/default.asp", "/home.asp", strErrorMessage)
'Result = HTTP_PostData(strPost_QueryString, ":", ",", HTTP_HTTP, HTTP_POST, "216.9.4.31", "80", "129.196.225.89", "syncforce", "/syncforce/firewalltest/default.asp", "/SyncForce/response.asp", strErrorMessage)

response.write("Receiving data.....<BR><BR>")

Response.write Result & "<BR>"
response.write strErrorMessage


%>