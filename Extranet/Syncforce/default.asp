<!-- #include virtual="/include/HTTP_DataTransfer.asp" -->
<%
Dim strErrorMessage
Dim strPost_QueryString

Dim strProtocol
Dim strMethod
Dim iRemoteHostPort
Dim strRemoteHostIP
Dim strRemoteHostName
Dim strRemoteHostTargetFile
Dim strLocalHostIP
Dim strLocalHostName
Dim strLocalHostReferrerFile
Dim strResponse

strPost_QueryString = "footest1:foo1_value,footest2:foo2_value,footest3:foo3_value"
strProtocol = HTTP_HTTP
strMethod = HTTP_GET

iRemoteHostPort = 80
strRemoteHostIP = "194.109.100.86"
strRemoteHostName = "flukepartnerportal.com"
strRemoteHostTargetFile = "/FlowManager/Receive.asp"

strLocalHostIP = "129.196.231.94"
strLocalHostName = "support.fluke.com"
strLocalHostReferrerFile = "/syncforce/default.asp"


HTTP_PostData strPost_QueryString, ":", ",", strProtocol, strMethod, strRemoteHostIP, iRemoteHostPort, strRemoteHostName, strRemoteHostTargetFile, strLocalHostIP, strLocalHostName, strLocalHostReferrerFile, strResponse

response.write("strResponse on default.asp: " & strResponse & "<BR>")
%>