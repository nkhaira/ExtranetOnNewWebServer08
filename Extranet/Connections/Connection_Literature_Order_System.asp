<%
' --------------------------------------------------------------------------------------
' Connection string to remote URL for Literature Order System
'
' Call this routine before you do the actual data transfer setup
'
' Kelly Whitlock
' 10/01/2003
' --------------------------------------------------------------------------------------

Dim strDomain, strNodes, strMyEnvironment

strDomain        = UCase(request.ServerVariables("SERVER_NAME"))
strNodes         = split(strDomain,".")
strMyEnvironment = ""

for each node in strNodes
 	select case UCase(node)
		case "DEV", "TEST", "PRD", "DTMEVTVSDV15"
			strMyEnvironment = node
			exit for
	end select
next
    
 
select case strMyEnvironment
	  case "DTMEVTVSDV15"
  	  'strRemoteHostName        = "dtmevtvsdv15.danahertm.com"
  	  'strRemoteHostTargetFile  = "/pod/dcg/response.asp"
	  strRemoteHostName        = "us.dev.fluke.com"
  	  strRemoteHostTargetFile  = "/testshoppingcart.asp"
          'strRemoteHostName = "www.nacorporation.com"
          'strRemoteHostTargetFile  = "/smart/PartnerPortal.asp"
'strRemoteHostName = "www.dev.fluke.com"
     'strRemoteHostTargetFile  = "/testshoppingcart.asp"
	  case "TEST"                    
     'Currently keeping the test url same. Since not sure if the NAC will provide the test URL.
  	  'strRemoteHostName        = "dev.dcgonline.net"
  	  'strRemoteHostTargetFile  = "/flukeportal/response.asp"

           strRemoteHostName = "us.test.fluke.com"
           strRemoteHostTargetFile  = "/testshoppingcart.asp"

           'strRemoteHostName = "www.nacorporation.com"
           'strRemoteHostTargetFile  = "/smart/PartnerPortal.asp"
	  case else
      'Commented out on 03/01/2009 as we are not going to use the DCG system.
      ' Support.Fluke.com (Debug on Live)
      
  	   'strRemoteHostName        = "support.fluke.com"
  	   'strRemoteHostTargetFile  = "/pod/dcg/response.asp"
      
	   'strRemoteHostName        = "www.hkmdm.com"
      '>>>>>>>>>>>>>>
      strRemoteHostName = "www.nacorporation.com"
      strRemoteHostTargetFile  = "/smart/PartnerPortal.asp"
    
      'Commented out on 03/01/2009 as we are not going to use the DCG system.
      'if LCase(Session("Logon_User")) = "dcgtest" then
  	    'strRemoteHostTargetFile  = "/fluke/responsedev.asp"
      'else
  	    'strRemoteHostTargetFile  = "/fluke/response.asp"
      'end if
      ''''>>>>>>>>>>>>>>>>>>>>>>>>>>>
end select

strRemoteHostURL = "http://" & strRemoteHostName & strRemoteHostTargetFile
%>