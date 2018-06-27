<%
' --------------------------------------------------------------------------------------
' Administrator Login
' --------------------------------------------------------------------------------------

' This functionality is no longer used.  The CreateObject and Sub call have been remarked
' to disable this functionality.  At some time, the SiteWide applications need to be reviewed
' to remove the <#!-- Include... reference to "Connection_Login_Admin.asp"
'
' --------------------------------------------------------------------------------------

'Set AU = Server.CreateObject("Persits.AspUser")
'Call LoginAdmin(AU)

sub LoginAdmin(AU)
	Dim sDomain,aNodes,myenv
	
	sDomain = UCase(request.ServerVariables("SERVER_NAME"))
	AU.Domain = sDomain
    
    
    aNodes = split(sDomain,".")
    myenv = ""
    for each node in aNodes
		select case node
			case "DEV", "TST", "TEST", "DTMEVTVSDV15", "DTMEVTVSDV15", "PRD"
				myenv = node
				exit for
		end select
    Next
    
    ' These will have to change with new FNET Asset Servers PORTWEB
    select case myenv
  		case "DTMEVTVSDV15"
  			AU.LogonUser "DTMEVTVSDV15", "ADMIN_ASPUser", "!Smithers2001#"
  		case "DTMEVTVSDV18"
  			AU.LogonUser "DTMEVTVSDV18", "ADMIN_ASPUser", "!Smithers2001#"
  		case "DEV"
  			AU.LogonUser "evtibg05",     "ADMIN_ASPUser", "!Smithers2001#"
      case "TST", "TEST"
  			AU.LogonUser "flkprd05",     "ADMIN_ASPUser", "!Smithers2001#"
      case else
  			AU.LogonUser "flkprd05",     "ADMIN_ASPUser", "!Smithers2001#"                
    end select

end sub
%>
