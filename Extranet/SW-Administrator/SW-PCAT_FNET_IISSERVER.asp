<%
strServerName = UCase(Request.ServerVariables("SERVER_NAME"))
if(Site_ID = "") then
    Site_ID = Request("Site_ID")
end if

if instr(strServerName,"DTMEVTVSDV15") > 0 then
   'striisserverpath = "http://dtmevtvsdv15:9090/ExtranetPcat/PcatHttpHandler.aspx"        
   'striisserverpath = "http://dtmevtvsdv17.danahertm.com:8080/extranetpcat/PcatHttpHandler.aspx"
   if (CInt(Site_ID) = 82) Then
		striisserverpath = "http://dtmevtvsdv17:8080/extranetpcat/PcatHttpHandler.aspx"
   else
		striisserverpath = "http://dtmevtsvdv08.danahertm.com/ExtranetPcat/PcatHttpHandler.aspx"
   end if
   'striisserverpath = "http://dtmevtsvdv06.danahertm.com/ExtranetPcat/PcatHttpHandler.aspx"
elseif  instr(strServerName,"DEV") > 0 then  
   striisserverpath = "http://DEV.SupportWebApps.FlukeNetworks.com/ExtranetPcat/PcatHttpHandler.aspx"     
elseif instr(strServerName,"TEST") > 0 then
   'Replace this path after fixing the Test server.
   if (CInt(Site_ID) = 82) Then
	 striisserverpath = "http://supportwebapps.test.ib.fluke.com/ExtranetPcat/PcatHttpHandler.aspx"
   else
	 striisserverpath = "http://supportwebapps.test.ib.fluke.com/IGExtranetPcat/pcathttphandler.aspx"
   end if
elseif instr(strServerName,"PRD") > 0 then
   'Replace this path after fixing the Production server.
   if (CInt(Site_ID) = 82) Then
	 striisserverpath = "http://supportwebapps.flk.ib.fluke.com/ExtranetPcat/PcatHttpHandler.aspx"
   else
	 striisserverpath = "http://supportwebapps.flk.ib.fluke.com/igextranetpcat/pcathttphandler.aspx"
   end if
   striisserverpath = "http://supportwebapps.flk.ib.fluke.com/ExtranetPcat/PcatHttpHandler.aspx"
else
	'Default to PRODUCTION
   if (CInt(Site_ID) = 82) Then
	 striisserverpath = "http://supportwebapps.flk.ib.fluke.com/ExtranetPcat/PcatHttpHandler.aspx"
   else
	 striisserverpath = "http://supportwebapps.flk.ib.fluke.com/igextranetpcat/pcathttphandler.aspx"
   end if
end if    
%>


