<%
        strRemoteHostName = "dtmevtvsdv15"
        strRemoteHostTargetFile = "/syncforce/response.asp"

        strRemoteHostName = "support.fluke.com"
        strRemoteHostTargetFile = ""
        
        
        strRemoteHostURL = "http://" & strRemoteHostName & strRemoteHostTargetFile

        strLocalHostReferrerFile = request.ServerVariables("SCRIPT_Name")
        strResponse              = ""
  
        strPost_QueryString      = "foo1=hello&foo2=world"
       
        response.write strPost_QueryString & "<P>"
        Dim HTTPRequest

        Set HTTPRequest = Server.CreateObject("Msxml2.XMLHTTP.3.0") 
        
        HTTPRequest.open "POST", strRemoteHostURL, false 
        HTTPRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        on error resume next          
        HTTPRequest.send strPost_QueryString

        if HTTPRequest.status <> 200 then         
          strResponse = err.Description & "<BR>Order not received by remote server."
          bResponse   = err.Number
          response.write strResponse & "<P>"
          response.write bResponse
          response.flush
        else  
          strResponse = HTTPRequest.responseText
          bResponse   = HTTPRequest.status
          response.write strResponse & "<P>"
          response.write bResponse
          response.flush
        end if
       
%>
