<%@ Language=VBScript %>
<% Option Explicit %>

<!--METADATA TYPE="TypeLib" UUID="{6B16F98B-015D-417C-9753-74C0404EBC37}" -->

<HTML>
<HEAD>
<TITLE>SoftArtisans FileUpEE 3-Tier Progress Indicator Sample</TITLE>
</HEAD>
<BODY>
<p>
<img src="/fileupee/samples/images/fileupee.gif" alt="SoftArtisans FileUpEE">
</p>

<%

'-----------------------------------------------------------------------
'--- FileUpEE 3-Tier Progress Indicator Sample
'---
'--- Monitor the transfer progress of FileUpEE making a 3-Tier upload to a filserver
'--- with progress indication for the client>webserver and webserver>fileserver
'--- layers independently
'---
'--- webserver.asp -- this script goes on the web server
'---
'--- Copyright (c) 2003 SoftArtisans, Inc.
'--- Mail: info@softartisans.com   http://www.softartisans.com
'-----------------------------------------------------------------------

	'--- Declarations

  Dim Debug_Flag
  Debug_Flag = false

	Dim oFileUpEE
	Dim oFile
	Dim intSAResult

	'--- Instantiate the FileUpEE object
	Set oFileUpEE = Server.CreateObject("SoftArtisans.FileUpEE")

	'--- Let FileUpEE know what stage we're working on
	'--- Note: Be sure to set this immediately after instantiating the object
	oFileUpEE.TransferStage = saWebServer

  if Debug_Flag then
    oFileUpEE.DebugLevel = 3
    oFileUpEE.DebugLogFile = "D:\Inetpub\extranet\FileUpE_Log\WA-Debug_Log.txt"

    oFileUpEE.AuditLogDestination(saWebServer) = saLogFile
    oFileUpEE.AuditLogFile(saWebServer) = "D:\Inetpub\extranet\FileUpE_Log\WA_Audit_Log.txt"
  end if

	'--- Enable progress indication for both layers

	oFileUpEE.ProgressIndicator(saClient) = True
	oFileUpEE.ProgressIndicator(saWebServer) = True

	'--- Get the progress IDs from the querystring and assign them to the FileUpEE object
	oFileUpEE.ProgressID(saClient) = CInt(request.querystring("progressid"))
	oFileUpEE.ProgressID(saWebServer) = CInt(request.querystring("wprogressid"))

	'--- Call ProcessRequest to read the request from the form
	'--- Parameters for ProcessRequest:
	'---	1) Request - ASP Request object
	'---	2) SOAP Request? - False, the web server is not being sent SOAP from the form
	'---	3) Auto Process? - False, do not auto process on the web server
	'---	Note: False, False is typically what you will use in a web server script

	On Error Resume Next

  	oFileUpEE.ProcessRequest Request, False, False

  	If Err.Number <> 0 Then
			'--- Display error info if ProcessRequest fails
			Response.Write "<B>WebServer ProcessRequest error</B><BR>" & _
							Err.Description & " (" & Err.Source & ")"
			Response.End
		End If

  On Error Goto 0

	'--- Set the URL to the fileserver script
	'--- getTargetURL() is a function used in these samples
	'--- to get the virtual path to the current directory
	'--- Find the source for this function in the include file include.inc.asp

'  oFileUpEE.TargetURL = "http://support.dev.fluke.com/FileUpE/fileserver.asp"
  oFileUpEE.TargetURL = "http://support.dev.fluke.com/SW-FileUp_Upload/FileUpEE_FileServer.asp"

	'--- Set the directory on the fileserver where you want to save the files
	oFileUpEE.DestinationDirectory = "D:\inetpub\extranet\find-sales\download\temp\"

	'--- The Files collection contains uploaded files
	'--- access each individual element by ordinal number
	'--- or by string, referring to the HTML tag's NAME attribute
	Set oFile = oFileUpEE.Files("myFile")

	'--- Now that we've set all the necessary properties,
	'--- send the request up to the fileserver
	On Error Resume Next
		intSAResult = oFileUpEE.SendRequest()
		If Err.Number <> 0 Then
			Response.Write "<B>WebServer SendRequest error:</B><BR>" & _
							Err.Description & " (" & Err.Source & ")"
			Response.Write "<BR><B>FileServer returned:</B><BR>" & _
							oFileUpEE.HttpResponse.BodyText
			Response.End
		End If
	On Error Goto 0

	'--- If a file was submitted, show the results
	'--- FileEe.Size will be 0 if no file was submitted
	If oFile.Size <> 0 Then

		'--- Always check the return value from SendRequest to check for success
		'--- If saResult is saAllProcessed, then all of the files were processed correctly
		If intSAResult = saAllProcessed Then

			'--- Display some properties of the uploaded file
			Response.Write "<H3>FileUpEE Processed All Files Successfully</H3>"
			Response.Write "<DL>"
			'--- SASaveResult status code of the file
			Response.Write "<DT><B>File SASaveResult</B></DT><DD>&nbsp;" & oFile.Processed & "</DD>"
			'--- Name of the type="file" form field that submitted the file
			Response.Write "<DT><B>File field name</B></DT><DD>&nbsp;" & oFile.FormName & "</DD>"
			'--- Full path of the file on the client
			Response.Write "<DT><B>Path on client</B></DT><DD>&nbsp;" & oFile.ClientPath & "</DD>"
			'--- Full path of the file saved on the server
			Response.Write "<DT><B>Path of saved file on server</B></DT><DD>&nbsp;" & _
							oFileUpEE.DestinationDirectory & oFile.ClientFileName & "</DD>"
			'--- Byte size of the file
			Response.Write "<DT><B>Byte size</B></DT><DD>&nbsp;" & oFile.Size & " bytes</DD>"
			'--- Mime type of the file
			Response.Write "<DT><B>Content type</B></DT><DD>&nbsp;" & oFile.ContentType & "</DD>"
			Response.Write "</DL>"

		'--- If the SAResult isn't saAllProcessed, there was an error processing one or more files
		Else

			'--- Display some more detailed error information
			Response.Write "<H3>An error occurred during the processing of one or more files</H3>"

			Response.Write "<DL>"
			Response.Write "<DT><B>SAResult was</B><DD>&nbsp;" & intSAResult & "</DD>"

			Response.Write "<DT><B>FileUpEE.Error</B></DT><DD>&nbsp;" & oFileUpEE.Error & "</DD>"

			Response.Write "<DT><B>Status message for file field:</B> " & oFile.FormName & "</DT>"

			'--- Check the SASaveResult status code for the file to get the specific error
			'--- and display the appropriate error message
			Select Case oFile.Processed

				'--- This will be the saSaveResult of OverwriteFile is True
				Case saExists
					Response.Write "<DD>The file was not saved because OverwriteFile " & _
									"is set to False and a file with the same name exists " & _
									"in the destination directory.</DD>"

				'--- This will be the saSaveReult for a general error
				Case saError
					Response.Write "<DD>" & oFile.Error & "</DD>"

				'--- Unknown error
				Case Else
					Response.Write "An unknown error has occurred. SASaveResult: " & oFile.Processed & _
									". " & oFile.Error
			End Select

			Response.Write "</DL>"

		End If
	Else
		Response.Write "No file was submitted for upload."
	End If

	Set oFileUpEE = Nothing
%>


</BODY>
</HTML>