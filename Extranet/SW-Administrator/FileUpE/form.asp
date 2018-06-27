<%@ Language=VBScript %>
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
'--- form.asp -- this script goes on the web server
'---
'--- Copyright (c) 2003 SoftArtisans, Inc.
'--- Mail: info@softartisans.com   http://www.softartisans.com
'-----------------------------------------------------------------------

	'--- Declarations
	Dim oFileUpEEProgressWS
	Dim oFileUpEEProgressClient
	Dim clProgID
	Dim wsProgID

	'--- On the form page, create two FileUpEEProgress instances.
	'--- one for each stage of the transfer we want to monitor
	Set oFileUpEEProgressWS = Server.CreateObject("SoftArtisans.FileUpEEProgress")
	Set oFileUpEEProgressClient = Server.CreateObject("SoftArtisans.FileUpEEProgress")

	'--- Tell each of the progress objects which stage to monitor
	oFileUpEEProgressClient.Watch = saClient
	oFileUpEEProgressWS.Watch = saWebServer

	'--- Get a new progress ID for the client>webserver layer
	clProgID = oFileUpEEProgressClient.NextProgressID

	'--- Get a new progress ID for the webserver>fileserver layer
	wsProgID = oFileUpEEProgressWS.NextProgressID

	'--- The client and webserver progress IDs (clProgID and wsProgID respectively)
	'--- are submitted to the progress indicator window and to the webserver script
	'--- as querystring parameters.  See the JavaScript startupload() function below.
	'--- Note:  The progress IDs MUST be submitted in the querystring, not in the Form
%>

<HTML>
<HEAD>
<TITLE>SoftArtisans FileUpEE 3-Tier Progress Indicator Sample</TITLE>

<SCRIPT Language="JavaScript">
/*
This small function makes sure that the progress indicator and server-side
processing page receive the new progress ID we just created.  Also, it pops up the
progress window.  This function fires with the form's onSubmit() event.
*/
function startUpload()
{
		//--- A string to specify the layout of the progress window
		winstyle="height=150,width=560,status=no,scrollbars=no,toolbar=no,menubar=no,location=no";
		//--- Launch the progress window.  Add the progressIDs to the querystring
		window.open("progress.asp?wprogressid=<%=wsProgID%>&progressid=<%=clProgID%>",null,winstyle);
		//--- Add the progressIDs to the querystring of the form action
		document.theForm.action="webserver.asp?wprogressid=<%=wsProgID%>&progressid=<%=clProgID%>";
}
</script>



</HEAD>

<BODY>
<p align=center>
<img src="/fileupee/samples/images/fileupee.gif" alt="SoftArtisans FileUpEE">
</p>

<H3 ALIGN=center>FileUpEE 3-Tier Progress Indicator Sample</H3>

<!--

Note: For any form uploading a file, the ENCTYPE="multipart/form-data" and
METHOD="POST" attributes MUST be present

-->

<TABLE ALIGN=center width="600" border=0>
<FORM onSubmit="startUpload();" name="theForm" ENCTYPE="MULTIPART/FORM-DATA" METHOD="POST">
<TR>
	<TD ALIGN="right">Current client progressID: </TD><TD><input type="text" size="3" value="<%=clProgID%>" readonly></TD>
</TR>
<TR>
	<TD ALIGN="right">Current webserver progressID: </TD><TD><input type="text" size="3" value="<%=wsProgID%>" readonly></TD>
</TR>

<TR>
	<TD ALIGN="RIGHT" VALIGN="TOP">Enter Filename:</TD>

<!--
Note: Notice this form element is of TYPE="FILE"
-->
	<TD ALIGN="LEFT"><INPUT TYPE="FILE" NAME="myFile"><BR>
	<I>Click "Browse" to select a file to upload</I>
	</TD>
</TR>
<TR>
	<TD ALIGN="RIGHT">&nbsp;</TD>
	<TD ALIGN="LEFT"><INPUT TYPE="submit" NAME="SUB1" VALUE="Upload File"></TD>
</TR>
<TR>
	<TD COLSPAN=2><HR NOSHADE><B>Note:</B> This sample will show you a progress indicator that lets your users know how much of the total upload has arrived at the server.  It's important to note that this is a <i>server-side</i> progress indicator.  As such, it can only indicate the progress of the <i>entire</i> upload, not individual files.  If you use a small file the progress indicator will complete very quickly.  Try a larger file to get a better observation of the progress indicator.</TD>
</TR>
</TABLE>
</FORM>


</BODY>
</HTML>
