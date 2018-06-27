<%
Dim sDomain, sNodes, sMyEnvironment

	sDomain        = UCase(request.ServerVariables("SERVER_NAME"))
  sNodes         = split(sDomain,".")
  sMyEnvironment = ""

  for each node in sNodes
  	select case node
			case "DTMEVTVSDV15", "DEV", "TEST", "PRD"
				sMyEnvironment = node
				exit for
		end select
  next
  
  ' Special Case for Meterman - Go Directly to Logon
  if instr(1,sDomain,"METERMANTESTTOOLS") > 0 then
    select case sMyEnvironment
      ' Domains that do not support SSL
	  	case "DTMEVTVSDV15", "DEV", "PRD"
        response.redirect "http://"   & sDomain & "/register/login.asp?Site_ID=16"
  		case else
	      response.redirect "https://"  & sDomain & "/register/login.asp?Site_ID=16"
    end select    
  end if
  
  ' Select Site then Logon or Register
  select case sMyEnvironment
    ' Domains that do not support SSL
		case "DTMEVTVSDV15", "DEV", "PRD"
			response.redirect "http://"  & sDomain & "/register/default.asp"
		case else
			response.redirect "https://" & sDomain & "/register/default.asp"       
  end select
%>