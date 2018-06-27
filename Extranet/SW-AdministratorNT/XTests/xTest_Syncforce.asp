<%

'strRemoteHostName       = "flukepartnerportal.com"
'strRemoteHostTargetFile = "/FlowManager/Receive.asp"
'strRemoteHostURL        = "http://" & strRemoteHostName & strRemoteHostTargetFile

'strRemoteHostName       = "www.thescripts.com"
'strRemoteHostTargetFile = "/forum/thread52484.html"
'strRemoteHostURL        = "http://" & strRemoteHostName & strRemoteHostTargetFile

strRemoteHostName       = "dtmevtvsdv15:9090"
strRemoteHostTargetFile = "/xTest_Server/xtest_syncforce_response.asp"
strRemoteHostURL        = "http://" & strRemoteHostName & strRemoteHostTargetFile

'response.Write strRemoteHostURL
'response.End
strLocalHostReferrerFile = request.ServerVariables("SCRIPT_Name")
strResponse              = ""
strPost_QueryString      = strPost_QueryString & "foo1=hello&foo2=Hello Ya'All"

Dim HTTPRequest

Set HTTPRequest = Server.CreateObject("Msxml2.SERVERXMLHTTP.6.0") 
  
Call HTTPRequest.open("POST",strRemoteHostURL,0,0,0)
Call HTTPRequest.setRequestHeader ("Content-Type", "application/x-www-form-urlencoded")
Call HTTPRequest.Send (strPost_QueryString)
  
if HTTPRequest.status <> 200 then
  strResponse = err.Description & "<BR>Post not received by remote server."
  bResponse   = err.Number
else  
  strResponse = HTTPRequest.responseText
  bResponse   = HTTPRequest.status
end if
  
response.write strResponse  

%>