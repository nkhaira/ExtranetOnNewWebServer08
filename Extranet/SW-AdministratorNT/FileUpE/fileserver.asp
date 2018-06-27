<%@ Language=VBScript %>
<% Option Explicit %>

<!--METADATA TYPE="TypeLib" UUID="{6B16F98B-015D-417C-9753-74C0404EBC37}" -->

<%
'-----------------------------------------------------------------------
'--- FileUpEE Generic FileServer Auto-Process ASP Script
'---
'--- This file server script will auto process any request submitted
'--- to it by a FileUpEE web server script.
'---
'--- The file is uploaded to the web server and then forwarded
'--- to the FileServer via SOAP to be saved
'---
'--- fileserver.asp -- This script goes on the fileserver
'---
'--- Copyright (c) 2003 SoftArtisans, Inc.
'--- Mail: info@softartisans.com   http://www.softartisans.com
'-----------------------------------------------------------------------

	'--- Declarations
	Dim oFileUpEE

	'--- Instantiate the FileUp object
	Set oFileUpEE = Server.CreateObject("SoftArtisans.FileUpEE")

  '--- Set the TransferStage to the appropriate value
	oFileUpEE.TransferStage = saFileServer

oFileUpEE.AuditLogDestination(saFileServer) = saLogFile
oFileUpEE.AuditLogFile(saFileServer) = "C:\Inetpub\wwwroot\FileUpE\auditlog_FileServer.txt"

  '--- Debug logging
'oFileUpEE.DebugLevel = 1
oFileUpEE.DebugLogFile = "C:\Inetpub\wwwroot\FileUpE\fs-debuglogfile.txt"


	'--- Handle errors gracefully for the remainder of the script
	On Error Resume Next

	'--- Call ProcessRequest  to read the submitted upload
	'--- Parameters for ProcessRequest:
	'---	1) Request - ASP Request object
	'---	2) SOAP Request? - True, the web server sends SOAP requests to the fileserver
	'---	3) Auto Process? - True, we want to auto-process the request from the web server
	'---	Note: "SOAP Request" will typically be True in the file server script
	oFileUpEE.ProcessRequest Request, True, True
	If Err.Number <> 0 Then
		'--- Display error info if ProcessRequest fails
		Response.Write "<B>FileServer ProcessRequest error</B><BR>" & _
						Err.Number & ": " & Err.Description & " (" & Err.Source & ")"
		Repsonse.Status = 500
		Response.End
	End If

	'--- Send the SOAP response back to the web server

	oFileUpEE.SendResponse Response
	If Err.Number <> 0 Then
		'--- Display error info if SendResponse fails
		Response.Write "<B>FileServer SendResponse error</B><BR>" & _
						Err.Description & " (" & Err.Source & ")"
		Repsonse.Status = 500
		Response.End
	End If

	'--- Destroy objects on longer needed
	Set oFileUpEE = Nothing
%>
