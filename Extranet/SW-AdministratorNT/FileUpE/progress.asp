<%@ Language="VBScript" EnableSessionState="False" %>
<% Option Explicit %>

<!--METADATA TYPE="TypeLib" UUID="{6B16F98B-015D-417C-9753-74C0404EBC37}" -->

<%
'-----------------------------------------------------------------------
'--- FileUpEE 3-Tier Progress Indicator Sample
'--- 
'--- Monitor the transfer progress of FileUpEE making a 3-Tier upload to a filserver
'--- with progress indication for the client>webserver and webserver>fileserver
'--- layers independently
'---
'--- progress.asp -- this script goes on the web server
'---
'--- Copyright (c) 2003 SoftArtisans, Inc.
'--- Mail: info@softartisans.com   http://www.softartisans.com
'-----------------------------------------------------------------------

	'--- Declarations
	Dim oFileUpProgressClient
	Dim oFileUpProgressWS
	Dim strStatus

	'--- Instantiate the FileUpProgress object
	Set oFileUpProgressClient = Server.CreateObject("Softartisans.FileUpEEProgress")
	Set oFileUpProgressWS = Server.CreateObject("Softartisans.FileUpEEProgress")
	
	'--- Tell each of the progress objects which stage to monitor
	oFileUpProgressClient.Watch = saClient
	oFileUpProgressWS.Watch = saWebServer

	'--- Assign the progress IDs from the querystring to the objects
	oFileUpProgressClient.ProgressID = CInt(request.querystring("progressid"))
	oFileUpProgressWS.ProgressID = CInt(request.querystring("wprogressid"))

	'-------------------------------------------------------------------------------------------
	'--- Note:  The table below reads the properties of the two FileUpEEProgress objects
	'--- in order to display the transfer status.  See the HTML tables below to see how it works
	'-------------------------------------------------------------------------------------------


%>

<html>
<Head>
<%
	'--- If the transfer is not yet complete, continue to refresh the page
	If oFileUpProgressWS.percentage < 100 Then
		Response.Write("<Meta HTTP-EQUIV=""Refresh"" CONTENT=1>")
		strStatus = "Upload in progress..."
	Else
		strStatus = "Upload complete"
	End If
%>
<title>FileUpEE 3-Tier Upload Progress Indicator: <%=strStatus%></TITLE>

</head>
<Body>
<TABLE border=1 align=center>
<TR>
	<TD><B>Transfer Stage</B></TD>
	<TD><B>Progress ID</B></TD>
	<TD align=left width=300><B>Graphic Indicator</B></TD>
	<TD><B>Transferred Bytes</B></TD>
	<TD><B>Total Bytes</B></TD>
	<TD><B>Transferred Percentage</B></TD>
</TR>

<TR>
<TD><nobr>Client to Webserver</nobr></TD>
<TD><%=oFileUpProgressClient.progressid%></TD>
<TD align=left width=300 valign="center">
	
	<TABLE height=5 bordercolor="black" border=1 cellspacing=0  
	  WIDTH="<%=oFileUpProgressClient.Percentage%>%"> 
		<TR> 
			 <TD height=5 align=right BGCOLOR="blue"> 
			 </TD> 
		</TR> 
	</TABLE>
	
</TD>
	
	<TD align=left><%=oFileUpProgressClient.transferredbytes%></TD>
	
	<TD align=left><%=oFileUpProgressClient.totalbytes%></TD>
	<TD align=left><%=oFileUpProgressClient.percentage%>%</TD>

</TR>
<TR>
<TD><nobr>Webserver to Fileserver</nobr></TD>
<TD><%=oFileUpProgressWS.progressid%></TD>
<TD align=left width=300 valign="center">
	
	<TABLE height=5 bordercolor="black" border=1 cellspacing=0  
	  WIDTH="<%=oFileUpProgressWS.Percentage%>%"> 
		<TR> 
			 <TD height=5 align=right BGCOLOR="blue"> 
			 </TD> 
		</TR> 
	</TABLE>
	
</TD>
	
	<TD align=left><%=oFileUpProgressWS.transferredbytes%></TD>
	
	<TD align=left><%=oFileUpProgressWS.totalbytes%></TD>
	<TD align=left><%=oFileUpProgressWS.percentage%>%</TD>

</TR>
</Table>
</Body>
</Html>