<%
strServerName = UCase(Request.ServerVariables("SERVER_NAME"))
cdnServerName = ""

if instr(strServerName,"DTMEVTVSDV15") > 0 then
	cdnServerName = "http://download.fluke.com/FCal/"	
elseif  instr(strServerName,"DEV") > 0 then 
	cdnServerName = "http://content.fluke.com/FCal_Dev/"
elseif instr(strServerName,"TEST") > 0 then
	cdnServerName = "http://content.fluke.com/FCal_Test/"
	
elseif instr(strServerName,"PRD") > 0 then 
	cdnServerName = "http://download.fluke.com/FCal/"
else
	cdnServerName = "http://download.fluke.com/FCal/"
end if

%>


