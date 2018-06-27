<%
strServerName = UCase(Request.ServerVariables("SERVER_NAME"))
cdnServerName = ""

if instr(strServerName,"DTMEVTVSDV15") > 0 then
	cdnServerName = "http://download.fluke.com/Biomedical/"	
elseif  instr(strServerName,"DEV") > 0 then 
	cdnServerName = "http://content.fluke.com/Biomed_Dev/"
elseif instr(strServerName,"TEST") > 0 then
	cdnServerName = "http://content.fluke.com/Biomed_Test/"
	
elseif instr(strServerName,"PRD") > 0 then 
	cdnServerName = "http://download.fluke.com/Biomedical/"
else
	cdnServerName = "http://download.fluke.com/Biomedical/"
end if

%>


