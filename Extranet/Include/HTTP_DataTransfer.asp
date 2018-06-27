<%
Const HTTP_GET   = 0
Const HTTP_POST  = 1
Const HTTP_HTTP  = 3

Const HTTP_DEBUG = False

Dim g_strErrors				' String to hold any errors returned

'----------------------------------------------------------------------------------------------
' PostDataToServer()
'	Purpose: Establish a sockets connection to a given server, send and receive data with this server.
'	InputParameters:
'		
'	OutputParameters:
'----------------------------------------------------------------------------------------------

Function PostDataToServer(strPost_QueryString, strKeyValueDelimiter, strPairDelimiter, iSite_ID, iRegion_ID, strReferrerFile, strResponse)

'	on error resume next

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
				strProtocol = HTTP_HTTP
			else
				strProtocol = HTTP_HTTPS
			end if
			strMethod = uCase(.Fields("Method"))
			if strMethod = "GET" then
				strMethod = HTTP_GET
			else
				strMethod = HTTP_POST
			end if
			iRemoteHostPort = .Fields("Remote_Port")
			strRemoteHostIP = .Fields("Remote_IP")
			strRemoteHostName = .Fields("Remote_Name")
			strRemoteHostTargetFile = .Fields("TargetFile")

			strLocalHostIP = .Fields("Local_IP")
			strLocalHostName = .Fields("HostName")

			set rsConnInfo = nothing
		end with
				
		bError = HTTP_PostData(strPost_QueryString, strKeyValueDelimiter, strPairDelimiter, strProtocol, strMethod, strRemoteHostIP, iRemoteHostPort, strRemoteHostName, strRemoteHostTargetFile, strLocalHostIP, strLocalHostName, strLocalHostReferrerFile, strResponse)
	end if
	
	PostDataToServer = bError

End Function

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

Function GetConnectionData(iSite_ID, iRegion_ID)

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

End Function

'----------------------------------------------------------------------------------------------

Function HTTP_PostData(strPost_QueryString, strKeyValueDelimiter, strPairDelimiter, cProtocol, cMethod, strRemoteHostIP, iRemoteHostPort, strRemoteHostName, strRemoteHostTargetFile, strLocalHostIP, strLocalHostName, strLocalHostReferrerFile, strResponse)

	Dim oHTTPComm		' Object reference to ip com component
	Dim strRecData		' Variant that will be populated with received data
	Dim iRecDataLen		' Variant that will be populated with length (in bytes) of received data
	Dim aData			' Array to hold the query string once we split it
	Dim aDataItem		' Individual array item
	Dim iPos			' Used to help parse aDataItem
	Dim strKey			' Value from the left side of each data item. eg: foo from foo=test1
	Dim strValue		' Value from right side of each data item. eg: test1 from foo=test1
	Dim bResponse		' Boolean returned from oHTTPComm functions to indicate success/failure and any errors
	Dim strOutput		' Used to see the exact string sent over to the target ip
	
	HTTP_DebugOutput "<FONT FACE=Arial SIZE=2><P>*********************************** Begin Debug section *************************************<P>", "", ""
	HTTP_DebugOutput "<B>strPost_QueryString: </B><BR>", Replace(strPost_QueryString,"&","<BR><B>&</B>"), ""
	HTTP_DebugOutput "<B>strKeyValueDelimiter: </B>",strKeyValueDelimiter, ""
	HTTP_DebugOutput "<B>strPairDelimiter: </B>",strPairDelimiter, ""
	HTTP_DebugOutput "<B>cProtocol: </B>",cProtocol, ""
	HTTP_DebugOutput "<B>cMethod: </B>",cMethod, ""
	HTTP_DebugOutput "<B>strRemoteHostIP: </B>",strRemoteHostIP, ""
	HTTP_DebugOutput "<B>iRemoteHostPort: </B>",iRemoteHostPort, ""
	HTTP_DebugOutput "<B>strRemoteHostName: </B>",strRemoteHostName, ""
	HTTP_DebugOutput "<B>strRemoteHostTargetFile: </B>",strRemoteHostTargetFile, ""
	HTTP_DebugOutput "<B>strLocalHostIP: </B>",strLocalHostIP, ""
	HTTP_DebugOutput "<B>strLocalHostName: </B>",strLocalHostName, ""
	HTTP_DebugOutput "<B>strLocalHostReferrerFile: </B>",strLocalHostReferrerFile, ""
	HTTP_DebugOutput "<B>strResponse: </B>",strResponse, ""
	HTTP_DebugOutput "", "", "END"
	HTTP_DebugOutput "<P>*********************************** End Debug section *************************************<P>", "", ""

	Set oHTTPComm = Server.CreateObject("VBHTTPComm.cVBHTTPComm")
	with oHTTPComm
		.TransferMethod = cMethod
		.RemoteHostPort = iRemoteHostPort
		.RemoteHostIP = strRemoteHostIP
		.RemoteHostName = strRemoteHostName
		.RemoteHostTargetFile = strRemoteHostTargetFile
		.LocalHostIP = strLocalHostIP
		.LocalHostName = strLocalHostName
		.LocalHostRefererFile = strLocalHostReferrerFile

		' Load the query string into the HTTPComm COM component
		aData = Split(strPost_QueryString, strPairDelimiter)   
		
		for each aDataItem in aData
			iPos = Instr(1, aDataItem, strKeyValueDelimiter)
			strKey = left(aDataItem, iPos - 1)
			strValue = mid(aDataItem, iPos + 1)

			HTTP_DebugOutput "<B>Adding Key" & strKeyValueDelimiter & "Value Pair</B>", "", ""
 			HTTP_DebugOutput strKey & strKeyValueDelimiter, strValue, ""
			.AddData strKey, strValue
		next
    
		' Open the connection. Check for errors.
		bResponse = .OpenConnection
		Call HTTP_ErrCheck("<P><B>Error on OpenConnection()</B>", trim(.ErrorDescription))

		' Send the data over to the server. Check for errors.
		bResponse = .SendData(strOutput)
		HTTP_DebugOutput "<B>strOutput</B> ", strOutput, ""
		Call HTTP_ErrCheck("<P><B>Error on SendData()</B>", trim(.ErrorDescription))

		' Get the response back (if any) from the other server. Check for errors.
		bResponse = .ReceiveData(strRecData, iRecDataLen)
		Call HTTP_ErrCheck("<P><B>Error on ReceieveData()</B>", trim(.ErrorDescription))
	end with

	HTTP_DebugOutput "<P><B>Data Received</B> ", strRecData, ""
	
	if len(trim(g_strErrors)) > 0 then	
		strResponse = "<P><B>Errors:</B> " & g_strErrors & "<P><B>Data Received:</B> " & strRecData & "<P>"
		HTTP_PostData = false
	else
		strResponse = strRecData
		HTTP_PostData = true
	end if

End Function

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