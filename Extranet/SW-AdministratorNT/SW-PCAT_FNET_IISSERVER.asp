<%
strServerName = UCase(Request.ServerVariables("SERVER_NAME"))

if instr(strServerName,"DTMEVTVSDV15") > 0 then
   'striisserverpath = "http://dtmevtvsdv15:9090/ExtranetPcat/PcatHttpHandler.aspx"        
   'striisserverpath = "http://dtmevtvsdv17.danahertm.com:8080/extranetpcat/PcatHttpHandler.aspx"
   'striisserverpath = "http://dtmevtvsdv17:8080/extranetpcat/PcatHttpHandler.aspx"
   striisserverpath = "http://dtmevtsvdv06.danahertm.com/ExtranetPcat/PcatHttpHandler.aspx"
    'striisserverpath = "http://supportwebapps.test.ib.fluke.com/ExtranetPcat/PcatHttpHandler.aspx"
elseif  instr(strServerName,"DEV") > 0 then  
   striisserverpath = "http://DEV.SupportWebApps.FlukeNetworks.com/ExtranetPcat/PcatHttpHandler.aspx"
   'striisserverpath = "http://supportwebapps.test.ib.fluke.com/ExtranetPcat/PcatHttpHandler.aspx"     
elseif instr(strServerName,"TEST") > 0 then
   'Replace this path after fixing the Test server.
   striisserverpath = "http://supportwebapps.test.ib.fluke.com/ExtranetPcat/PcatHttpHandler.aspx"     
elseif instr(strServerName,"PRD") > 0 then
   'Replace this path after fixing the Production server.
   striisserverpath = "http://supportwebapps.flk.ib.fluke.com/ExtranetPcat/PcatHttpHandler.aspx"
else
	 striisserverpath = "http://supportwebapps.flk.ib.fluke.com/ExtranetPcat/PcatHttpHandler.aspx"
end if    
%>


