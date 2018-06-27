<%@Language="VBScript"%>
<HTML>
<HEAD>
	<TITLE>SoftArtisans XXFileEE Resumable Download Sample</TITLE>
</HEAD>
<BODY>
<%
	'--- Declarations:
	Dim XFileEE
	Dim sMsg
    Set XFileEE = Server.CreateObject("SoftArtisans.FileUpEE")
	'--- Disable the download button.
	'document.all("GetButton").disabled = true
	XFileEE.ProcessRequest Request
	XFileEE.Resumable(1) = True
	XFileEE.TargetUrl = "http://dtmevtvsdv15/sw-administrator/TestResumeSupport.asp"
	'--- Set the request method to GET.
	'XFileEE.RequestMethod = "GET"
	
	'--- Set a ResumeInfo object.
	'Set XFileResumeInfo = XFileEE.ResumeInfo
	
	'--- Enable resumable downloading.
	'XFileResumeInfo.Resumable = True
	
	XFileEE.AddDownloadFile "http://dtmevtvsdv15/sw-administrator/9820185_ENG_A_X.EXE","c:\downloads\9820185_ENG_A_X.EXE"
	enumRetVal = XFileEE.SendRequest(Response)
	
	'--------- This line is for demonstation only -----------
	'--- Set the TimeOut to 2, so the download will fail 
	'--- (this will cause a timeout exception).
	'XFileEE.RequestTimeout =1
	XFileEE.ReceiveTimeout  = 2

	'------------------------------------------------------------
			
	'--- Turn default error handling off.
	On error resume next
	
	'--- Start the download.
	XFileEE.Start 
	
	MsgBox( "fdsgdfg")
	
	'--- Check the ResponseStatus.
	If instr(XFileEE.ResponseText, "200") > 0 then
		response.write("Download Completed")
	else
		'--- Display a message to the user with the reason for failure. Ask if 
		'--- the user wants to Resume the transfer.
		sMsg = "The download attempt failed. " & vbCrLf 
		sMsg = sMsg & "Reason: " & XFileEE.ResponseText & vbCrLf & vbCrLf
		sMsg = sMsg & "Would you like to resume the download?"
		%>
		
		<%
			'--------- This line is for demonstation only -----------
			'--- To force a download failure, XFile's TimeOut property was set to 2.
			'--- Before calling resume, reset TimeOut to a value that will allow the 
			'--- transfer to succeed.
			XFileEE.ReceiveTimeout = 90
			'-------------------------------------------------------------
			
			'--- Resume the transfer specified by the following JobId.
			XFileEE.Resume XFileEE.JobId
			
			'--- Check the response
			If instr(XFileEE.ResponseStatus, "200") > 0 OR instr(XFileEE.ResponseStatus, "206") > 0 then
				response.write("Resume Completed")
			Else
				response.write("Download Failed")
			End If			
		'Else
		'	response.write("Download Failed")
	End if
 %>
</BODY>
</HTML>

