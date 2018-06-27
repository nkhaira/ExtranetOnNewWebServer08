<%
' ESCROW changes:
' line 24 to: if len(GetEmailServer) = 0 then GetEmailServer = "mail.fluke.com"

function GetEmailServer()
	Dim sDomain,aNodes,myenv
	
	GetEmailServer = ""
	
	sDomain = UCase(request.ServerVariables("SERVER_NAME"))
    aNodes = split(sDomain,".")
    
    for each node in aNodes
		select case node
			case "DEV", "EVTIBG01"
				GetEmailServer = "mailhost.tc.fluke.com"
				exit for
      case "DTMEVTVSDV15", "DTMEVTVSDV18"
				GetEmailServer = "mailhost.tc.fluke.com"
				exit for
			case "TST", "TEST", "PRD"
				GetEmailServer = "mail.evt.danahertm.com:25"
				exit for
		end select
    Next
    
	if len(GetEmailServer) = 0 then GetEmailServer = "mail.evt.danahertm.com:25"
  
end function
%>
