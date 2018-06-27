<%
strServerName = UCase(Request.ServerVariables("SERVER_NAME"))
cdnServerName = ""

if instr(strServerName,"DTMEVTVSDV15") > 0 then
	'cdnServerName = "http://content.fluke.com/dev/datapaq/"	
	cdnServerName = "http://download.fluke.com/datapaq/"
elseif  instr(strServerName,"DEV") > 0 then 
	cdnServerName = "http://content.fluke.com/dev/datapaq/"
elseif instr(strServerName,"TEST") > 0 then
	cdnServerName = "http://content.fluke.com/test/datapaq/"
	
elseif instr(strServerName,"PRD") > 0 then 
	cdnServerName = "http://download.fluke.com/datapaq/"
else
	cdnServerName = "http://download.fluke.com/datapaq/"
end if

%>




