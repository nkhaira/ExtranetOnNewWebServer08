<%
Const HTTP_GET   = 0
Const HTTP_POST  = 1
Const HTTP_HTTP  = 3
Const HTTP_HTTPS = 4

Const HTTP_DEBUG = true

'----------------------------------------------------------------------------------------------
' PostDataToServer()
'	Purpose: Establish a sockets connection to a given server, send and receive data with this server.
'	InputParameters:
'		
'	OutputParameters:
'----------------------------------------------------------------------------------------------

function test()

strString = "Site_ID=3,ExpirationDate=12/31/2001,Prefix=,FirstName=Whitlockaa,MiddleName=,LastName=Whitlockaa,Suffix=,"

PostDataToServer strString, "=", ",", 3, iRegion_ID, strReferrerFile, strResponse

end function

Function PostDataToServer(strPost_QueryString, strKeyValueDelimiter, strPairDelimiter, iSite_ID, iRegion_ID, strReferrerFile, strResponse)

	on error resume next
	
	Dim bError
	Dim strProtocol
	Dim strMethod
	Dim iRemote_Port
	Dim strRemoteHost_IP
	Dim strLocalHost_IP
	Dim strHostName
	Dim strTargetFile

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
			iRemote_Port = .Fields("Remote_Port")
			strRemoteHost_IP = .Fields("Remote_IP")
			strLocalHost_IP = .Fields("Local_IP")
			strHostName = .Fields("HostName")
			strTargetFile = .Fields("TargetFile")
			set rsConnInfo = nothing
		end with
				
		bError = HTTP_PostData(strPost_QueryString, strKeyValueDelimiter, strPairDelimiter, strProtocol, strMethod, strRemoteHost_IP, iRemote_Port, strLocalHost_IP, strHostName, strReferrerFile, strTargetFile, strResponse)
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

Function HTTP_PostData(strPost_QueryString, strKeyValueDelimiter, strPairDelimiter, cProtocol, cMethod, strRemoteHost_IP, iRemote_Port, strLocalHost_IP, strHostName, strReferrerFile, strTargetFile, strResponse)
	Dim strRecData
	Dim oHTTPComm
	Dim aData
	Dim aDataItem
	Dim strKey
	Dim strValue
	Dim iPost
	Dim strError
			

HTTP_DebugOutput "<br><br>'*********************************** Begin Debug section *************************************<br><BR>", ""
HTTP_DebugOutput "strPost_QueryString",strPost_QueryString
HTTP_DebugOutput "strKeyValueDelimiter",strKeyValueDelimiter
HTTP_DebugOutput "strPairDelimiter",strPairDelimiter
HTTP_DebugOutput "cProtocol",cProtocol
HTTP_DebugOutput "cMethod",cMethod
HTTP_DebugOutput "strRemoteHost_IP",strRemoteHost_IP
HTTP_DebugOutput "iRemote_Port",iRemote_Port
HTTP_DebugOutput "strLocalHost_IP",strLocalHost_IP
HTTP_DebugOutput "strHostName",strHostName
HTTP_DebugOutput "strReferrerFile",strReferrerFile
HTTP_DebugOutput "strTargetFile",strTargetFile
HTTP_DebugOutput "strResponse",strResponse
HTTP_DebugOutput "<br><br>", ""


		Set oHTTPComm = Server.CreateObject("VBHTTPComm.cVBHTTPComm")
		with oHTTPComm
			HTTP_DebugOutput "cMethod", cMethod
			.TransferMethod = cMethod
			HTTP_DebugOutput "iRemote_Port", iRemote_Port
			.RemotePort = iRemote_Port
			HTTP_DebugOutput "strRemoteHost_IP", strRemoteHost_IP
			.RemoteHostIP = strRemoteHost_IP
			HTTP_DebugOutput "strLocalHost_IP", strLocalHost_IP
			.LocalHostIP = strLocalHost_IP
			HTTP_DebugOutput "strHostName", strHostName
			.HostName = strHostName
			HTTP_DebugOutput "strTargetFile", strTargetFile
			.TargetFile = strTargetFile
			HTTP_DebugOutput "strRferrerFile", strReferrerFile & "<BR>"
			.RefererFile = strReferrerFile

			.Delimiter = "="

			aData = Split(strPost_QueryString, strPairDelimiter)   
			iDatasetCount = 0
			
			for each aDataItem in aData
				iPos = Instr(1, aDataItem, strKeyValueDelimiter)
				strKey = left(aDataItem, iPos - 1)
				strValue = mid(aDataItem, iPos + 1)

			        iDatasetCount = iDatasetCount + 1

				.AddData strKey, strValue
	  			HTTP_DebugOutput strKey, strValue
			next

			.OpenConnection
			strError = strError & .GetErrorDescription
			HTTP_DebugOutput "<BR>OpenConnectionErr", strError & "<BR>"

'			response.write(.SendData)
			.SendData

			strError = strError & .GetErrorDescription
			HTTP_DebugOutput "SendDataErr", strError

			bResponse = true

dim strOutput
dim bContinue
bContinue = true

'			do while .IsDataAvailable and bResponse and bContinue
			do while .IsDataAvailable and bResponse
				'strRecData = .ReceiveData(cStr(strRecData), cLng(iRecDataLen))
				bResponse = .ReceiveData(strRecData, iRecDataLen)
				strOutput = strOutput & strRecData
				HTTP_DebugOutput "<BR>bResponse", bResponse
				HTTP_DebugOutput "<BR>strRecData", strRecData & "<BR>"
				HTTP_DebugOutput "Length", iRecDataLen

				if trim(strRecData) = "" then
					bContinue = false
				end if
			loop
			strError = strError & .GetErrorDescription
			HTTP_DebugOutput "HTTP_Err", strError
			strError = strError & .GetErrorDescription
			HTTP_DebugOutput "Err", strError

		end with

'response.write("strOutput: " & strOutput & "<BR>")
'response.end

	if len(trim(strError)) > 0 then	
		strResponse = "Errors: " & strError & "<BR><BR>Data received: " & strRecData
		HTTP_PostData = false
	else
		strResponse = strRecData
		HTTP_PostData = true
	end if
End Function

'----------------------------------------------------------------------------------------------

Function HTTP_DebugOutput(strName, strValue)
	if HTTP_DEBUG = true then
		response.write(strName & ": " & strValue & "<BR>")
	end if
End Function
%>