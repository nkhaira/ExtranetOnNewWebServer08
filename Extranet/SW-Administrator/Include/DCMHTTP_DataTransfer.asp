<%
'----------------------------------------------------------------------------------------------
' PostDataToServer()
'	Purpose: Establish a connection to a remote server target, send and receive data with this server.
' This version is specifically designed to interface with the European DCM System (Syncforce)
' 
' Original Author:  Jeff Patrick
' Author:  Kelly WHitlock
' Date: :  09/01/2006
'----------------------------------------------------------------------------------------------

function PostDataToServer(strPost_QueryString, strKeyValueDelimiter, strPairDelimiter, iSite_ID, iRegion_ID, strReferrerFile, strResponse)

	Dim bError
	Dim strProtocol
	Dim strMethod

	Dim strRemoteHostIP
	Dim iRemoteHostPort
	Dim strRemoteHostName
	Dim strRemoteHostTargetFile

	Dim strLocalHostIP
	Dim strLocalHostName

	' Retrieve connection information and initialize transfer object
	set rsConnInfo = GetConnectionData(iSite_ID, iRegion_ID)

	if err.number > 0 then
		PostDataToServer = false
		strResponse = "Error executing stored procedure: " & err.description
		exit function
	end if

	if not rsConnInfo.EOF then
		with rsConnInfo
			strProtocol = .Fields("Protocol")
			if strProtocol = "HTTP" then
				strProtocol = "HTTP"
			else
				strProtocol = "HTTPS"
			end if
			strMethod = uCase(.Fields("Method"))
			if strMethod = "GET" then
				strMethod = "GET"
			else
				strMethod = "POST"
			end if
			strRemoteHostName = .Fields("Remote_Name")
			strRemoteHostTargetFile = .Fields("TargetFile")
			set rsConnInfo = nothing
		end with
				
		bError = HTTP_PostData(strPost_QueryString, strKeyValueDelimiter, strPairDelimiter, strProtocol, strMethod, strRemoteHostIP, iRemoteHostPort, strRemoteHostName, strRemoteHostTargetFile, strLocalHostIP, strLocalHostName, strLocalHostReferrerFile, strResponse)
	end if
	
	PostDataToServer = bError

end function

'----------------------------------------------------------------------------------------------
' GetConnectionData()
'	Purpose: Get specific connection information necessary to establish a sockets connection
'	Input parameters: none
'	Output parameters:	Recordset containing the following information:
'		TransferMethod
'			0 - GET
'			1 - POST
'		RemotePort 		(typically 80)
'		RemoteHostIP	(ip address of server we're connecting to)
'		LocalHostIP		(ip address of local server)
'		HostName		(ip address or name of web server we're connecting to 
'							used to handle servers using host headers
'		TargetFile		(relative path to file being posted to)
'		RefererFile		(this file)
'----------------------------------------------------------------------------------------------

function GetConnectionData(iSite_ID, iRegion_ID)

	Dim cmd
	Dim tmpParameter
	Dim rsConnInfo
		
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "Get_CMS_System"
	Set prm = cmd.CreateParameter("@iSiteID", adInteger, adParamInput, , iSite_ID)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@iRegion_ID", adInteger, adParamInput, , iRegion_ID)
	cmd.Parameters.Append prm
	set rsGetConnInfo = cmd.execute
	set cmd = nothing

	set GetConnectionData = rsGetConnInfo

end function

'----------------------------------------------------------------------------------------------

Function HTTP_PostData(strPost_QueryString, strKeyValueDelimiter, strPairDelimiter, cProtocol, cMethod, strRemoteHostIP, iRemoteHostPort, strRemoteHostName, strRemoteHostTargetFile, strLocalHostIP, strLocalHostName, strLocalHostReferrerFile, strResponse)

	Dim oHTTPComm		' Object reference to ip com component
	Dim strRecData	' Variant that will be populated with received data
	Dim iRecDataLen	' Variant that will be populated with length (in bytes) of received data
	Dim aData			  ' Array to hold the query string once we split it
	Dim aDataItem		' Individual array item
	Dim iPos			  ' Used to help parse aDataItem
	Dim strKey			' Value from the left side of each data item. eg: foo from foo=test1
	Dim strValue		' Value from right side of each data item. eg: test1 from foo=test1
	Dim bResponse		' Boolean returned from oHTTPComm functions to indicate success/failure and any errors
	Dim strOutput		' Used to see the exact string sent over to the target ip
	
  strRemoteHostURL  = cProtocol & "://" & strRemoteHostName & strRemoteHostTargetFile

  Set oHTTPComm = Server.CreateObject("Msxml2.SERVERXMLHTTP.6.0") 
	' Load the query string into the HTTPComm COM component
	aData = Split(strPost_QueryString, strPairDelimiter)   
		
  strPostString = ""
  for each aDataItem in aData
'response.write aDataItem & " - " & "Test" & "<br>"
  	iPos = Instr(1, aDataItem, strKeyValueDelimiter)
  	strKey = left(aDataItem, iPos - 1)
  	strValue = mid(aDataItem, iPos + 1)
    if strPostString <> "" then strPostString = strPostString & "&"
  	strPostString = strPostString & strKey & "=" & strValue
  next
'response.write  "Succeded" & "<br>"    
  ' Open the connection

  Call oHTTPComm.open(cMethod,strRemoteHostURL,0,0,0)   
  Call oHTTPComm.setRequestHeader ("Content-Type", "application/x-www-form-urlencoded")
  
  ' Send the data over to the server.
  on error resume next
  Call oHTTPComm.send (strPostString)
  if Err <> 0 then
    response.write "<P>Post Error: " & Err.Number & " - " & Err.Description & "<P>"
    response.write "<P>" & strRemoteHostURL & "<P>"
    response.write "<P>" & strPostString & "<P>"
  end if
  on error goto 0
  
  ' Get the response back (if any) from the other server. Check for errors.
  on error resume next
  strResponse = oHTTPComm.responseText
  response.write  strResponse
  if Err <> 0 then
    response.write "<P>Response Error: " & Err.Number & " - " & Err.Description & "<P>"
    response.write "<P>" & strResponse & "<P>"
  end if
  
  on error goto 0

	if oHTTPComm.status <> 200 then	
		strResponse = "<P><B>Errors:</B> " & oHTTPComm.status & "<P><B>Data Received:</B> " & strResponse & "<P>"
    set oHTTPComm = nothing
    HTTP_PostData = false
	else
    strResponse = "200 OK<BR>" & strResponse
    set oHTTPComm = nothing
		HTTP_PostData = true
	end if

end function

'----------------------------------------------------------------------------------------------

Dim iTableStatus

iTableStatus = 0

Function HTTP_DebugOutput(strName, strValue, strAction)

	if HTTP_DEBUG=true then
		if uCase(strAction) = "BEGIN" then
			response.write "<TABLE BORDER=1>"
			iTableStatus = -1
		elseif uCase(strAction) = "END" then
			response.write "</TABLE>"
			iTableStatus = 0
		else
			if not trim(strName) = "" then
  '      if instr(1,strName,"<BR>") > 0 then
  '        strName = Replace(strName,"<BR>",": <BR>")
  '      else  
'  				strName = strName & ":"
  '      end if  
			end if
			if iTableStatus = -1 then
				response.write("<TR><TD VALIGN=TOP><FONT FACE=ARIAL SIZE=2>" & strName & "</FONT></TD><TD VALIGN=TOP><FONT FACE=ARIAL SIZE=2>" & strValue & "</FONT></TD></TR>")
			else
				response.write(strName & strValue & "<P>")
			end if
		end if
	end if

End Function

' -----------------------------------------------------------------------------------------------------

Function HTTP_ErrCheck(strName, strError)

	if not strError = "" then
		HTTP_DebugOutput strName, strError, ""
		g_strErrors = g_strErrors & strError & "<P>"
	end if

End Function

' -----------------------------------------------------------------------------------------------------
' Function: URLEncode(strPost_QueryString, strKeyDelimiter, strStringWrapper, strFieldDelimiter)
'	Purpose: URL encode a given string (usually a sql statement) for submission to a web server
'		 We can't just URL encode the whole string as we only want to encode the string values and not
'		 affect the post key value delimiter '='.
'	Example: foo = URLEncode("name='foo1',address='foo2',site_id=3)
'	Definition:
'		strPost_QueryString:	string to encode
'		strKeyDelimiter:	character separating field\value pair. "=" in example
'		strStringWrapper:	character wrapping string. "'" in example
'		strFieldDelimiter:	character separating field\value pairs. "," in example
' -----------------------------------------------------------------------------------------------------

Function URLEncode(strPost_QueryString, strKeyDelimiter, strStringWrapper, strFieldDelimiter)

	Dim strOutput
	Dim strKey
	Dim strValue
	Dim iLeftPos
	Dim iRightPos
	Dim iEndPos
	Dim strTarget

	strTarget = strPost_QueryString
	iLeftPos = 1
	iRightPos = InStr(iLeftPos, strTarget, strKeyDelimiter)
	

	' loop through the querystring until we get to the end or we can't find the next key delimiter

	do while (iRightPos < len(strTarget)) and (iRightPos >= 1)

		strKey = mid(strTarget, iLeftPos, iRightPos - iLeftPos)

		' if the value is a string then we need to parse

		if mid(strTarget, iRightPos + 1, 1) = strStringWrapper then
			iEndPos = InStr(iRightPos + 2, strTarget, strStringWrapper & strFieldDelimiter)
			if iEndPos < 1 then
				iEndPos = len(strTarget) + 1
			end if

			strValue = mid(strTarget, iRightPos + 2, iEndPos - (iRightPos + 2))
			strValue = Server.URLEncode(strValue)
			
			iEndPos = InStr(iEndPos, strTarget, strFieldDelimiter)
			if iEndPos < 1 then
				iEndPos = len(strTarget) + 1
			end if

		' else it's a numeric value and we can just take it
		else
			iEndPos = InStr(iRightPos + 1, strTarget, strFieldDelimiter)
			if iEndPos < 1 then
				iEndPos = len(strTarget) + 1
			end if

			strValue = mid(strTarget, iRightPos + 1, (iEndPos - 1) - iRightPos)
		end if

		strOutput = strOutput & strKey & "=" & strValue & "&"

		if (iEndPos >= len(strTarget)) then
			strTarget = ""
		else
			strTarget = right(strTarget, len(strTarget) - iEndPos)
		end if

		iLeftPos = 1
		iRightPos = InStr(iLeftPos, strTarget, strKeyDelimiter)

	loop

	if len(strOutput) > 0 then
		if right(strOutput, 1) = "&" then
			strOutput = left(strOutput, len(strOutput) - 1)
		end if
	end if
	
	URLEncode = strOutput

End Function

' -----------------------------------------------------------------------------------------------------
%>