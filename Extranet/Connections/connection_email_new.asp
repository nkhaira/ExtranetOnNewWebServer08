<%
Set msg = Server.CreateObject("CDO.Message")
Set conf= Server.CreateObject("CDO.Configuration")

confURL = "http://schemas.microsoft.com/cdo/configuration/"
with conf
    .Fields.Item(confURL & "sendusing") = 2
    '.Fields.Item(confURL & "smtpserverport") = 25
    .Fields.Item(confURL & "smtpserver") = GetEmailServerNew()
    .Fields.Update
end with

with msg
    .Fields.Item("urn:schemas:mailheader:X-Priority") = 2
    .Fields.Update
end with

function GetEmailServerNew()
	Dim sDomain,aNodes,myenv
	
	GetEmailServerNew = ""
	
	sDomain = UCase(request.ServerVariables("SERVER_NAME"))
    aNodes = split(sDomain,".")
    
    for each node in aNodes
		select case node
			case "DEV", "EVTIBG01"
				GetEmailServerNew = "mailhost.tc.fluke.com"
				exit for
            case "DTMEVTVSDV15", "DTMEVTVSDV18"
				GetEmailServerNew = "mailhost.tc.fluke.com"
				exit for
			case "TST", "TEST", "PRD"
				GetEmailServerNew = "mail.evt.danahertm.com"
				exit for
		end select
    Next
    
	if len(GetEmailServerNew) = 0 then GetEmailServerNew = "mail.evt.danahertm.com"
  
end function

%>