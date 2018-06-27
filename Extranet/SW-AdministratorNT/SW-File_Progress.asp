<%@ Language="VBScript" EnableSessionState="False" %>
<% Option Explicit %>

<!--METADATA TYPE="TypeLib" UUID="{6B16F98B-015D-417C-9753-74C0404EBC37}" -->

<LINK REL=STYLESHEET HREF="/SW-Common/SW-Style.css">
<%
'-----------------------------------------------------------------------
'--- FileUpEE 3-Tier Progress Indicator Sample
'--- 
'--- Monitor the transfer progress of FileUpEE making a 3-Tier upload to a filserver
'--- with progress indication for the client>webserver and webserver>fileserver
'--- layers independently
'---
'--- progress.asp -- this script goes on the web server
'-----------------------------------------------------------------------

'--- Declarations

Dim oFileUpProgressClient
Dim oFileUpProgressWS
Dim strStatus, bStatus

'--- Instantiate the FileUpProgress object

Set oFileUpProgressClient = Server.CreateObject("Softartisans.FileUpEEProgress")
Set oFileUpProgressWS = Server.CreateObject("Softartisans.FileUpEEProgress")

'--- Tell each of the progress objects which stage to monitor

oFileUpProgressClient.Watch = saClient
oFileUpProgressWS.Watch     = saWebServer

'--- Assign the progress IDs from the querystring to the objects

oFileUpProgressClient.ProgressID = CInt(request.querystring("progressid"))
oFileUpProgressWS.ProgressID     = CInt(request.querystring("wprogressid"))

with response

  .write "<HTML>" & vbCrLf
  .write "<HEAD>" & vbCrLf

	'--- If the transfer is not yet complete, continue to refresh the page

	if oFileUpProgressWS.percentage < 100 then
    bStatus = false
		.write("<Meta HTTP-EQUIV=""Refresh"" CONTENT=1>")
		strStatus = "<SPAN CLASS=SMALLRED>File Upload in Progress...</SPAN>"
	else
    bStatus = true
		strStatus = "<SPAN CLASS=SMALL><FONT COLOR=""#006600"">File Upload Complete</FONT></SPAN>"
	end if

  .write "<TITLE>SiteWide Upload Progress Indicator</TITLE>" & vbCrLf
  .write "</HEAD>" & vbCrLf
  .write "<BODY>" & vbCrLf
  
  .write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 BGCOLOR=""Black"">" & vbCrLf
  .write "<TR>" & vbCrLf
  .write "<TD>" & vbCrLf
  
  .write "<TABLE BORDER=0 ALIGN=CENTER CELLPADDING=4 CELLSPACING=1 CLASS=SMALL BGCOLOR=""Black"">" & vbCrLf
  .write "<TR>" & vbCrLf
  .write "<TD BGCOLOR=""White"" CLASS=SMALL><B>Transfer Stage</B></TD>" & vbCrLf
  .write "<TD BGCOLOR=""White"" ALIGN=LEFT WIDTH=300 CLASS=SMALL><B>Progress</B></TD>" & vbCrLf
  .write "<TD ALIGN=CENTER BGCOLOR=""White"" CLASS=SMALL><B>Transferred KBytes</B></TD>" & vbCrLf
  .write "<TD ALIGN=CENTER BGCOLOR=""White"" CLASS=SMALL><B>Total KBytes</B></TD>" & vbCrLf
  .write "<TD ALIGN=CENTER BGCOLOR=""White"" CLASS=SMALL><B>Transferred Percentage</B></TD>" & vbCrLf
  .write "</TR>" & vbCrLf

  .write "<TR>" & vbCrLf
  .write "<TD BGCOLOR=""White"" CLASS=SMALL><NOBR>Client to Webserver</NOBR></TD>" & vbCrLf
  .write "<TD ALIGN=LEFT WIDTH=300 VALIGN=""center"" BGCOLOR=""White"" CLASS=SMALL>" & vbCrLf
  .write "<TABLE HEIGHT=5 BORDERCOLOR=""black"" BORDER=1 CELLSPACING=0 WIDTH=""" & oFileUpProgressClient.Percentage & "%"">" & vbCrLf
  .write "<TR>" & vbCrLf
  .write "<TD HEIGHT=5 ALIGN=RIGHT BGCOLOR=""blue"" CLASS=SMALL>" & vbCrLf
  .write "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf
  .write "</TABLE>" & vbCrLf
	
  .write "</TD>" & vbCrLf
	
  .write "<TD ALIGN=RIGHT BGCOLOR=""White"" CLASS=SMALL>" & FormatNumber((oFileUpProgressClient.transferredbytes / 1024),0) & "</TD>" & vbCrLf
	
  .write "<TD ALIGN=RIGHT BGCOLOR=""White"" CLASS=SMALL>" & FormatNumber((oFileUpProgressClient.totalbytes / 1024),0) & "</TD>" & vbCrLf
  .write "<TD ALIGN=RIGHT BGCOLOR=""White"" CLASS=SMALL>" & oFileUpProgressClient.percentage & "%</TD>" & vbCrLf
  .write "</TR>" & vbCrLf

  .write "<TR>" & vbCrLf
  .write "<TD BGCOLOR=""White"" CLASS=SMALL><NOBR>Webserver to Fileserver</NOBR></TD>" & vbCrLf
  .write "<TD ALIGN=LEFT WIDTH=300 VALIGN=""center"" BGCOLOR=""White"" CLASS=SMALL>" & vbCrLf
  
	.write "<TABLE HEIGHT=5 BORDERCOLOR=""black"" BORDER=1 CELLSPACING=0  WIDTH=""" & oFileUpProgressWS.Percentage & "%"">" & vbCrLf
  .write "<TR>" & vbCrLf
  .write "<TD HEIGHT=5 ALIGN=RIGHT BGCOLOR=""blue"" CLASS=SMALL>" & vbCrLf
  .write "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf
  .write "</TABLE>" & vbCrLf
  .write "</TD>" & vbCrLf
  
  .write "<TD ALIGN=RIGHT BGCOLOR=""White"" CLASS=SMALL>" & FormatNumber((oFileUpProgressWS.transferredbytes / 1024),0) & "</TD>" & vbCrLf
  .write "<TD ALIGN=RIGHT BGCOLOR=""White"" CLASS=SMALL>" & FormatNumber((oFileUpProgressWS.totalbytes / 1024),0) & "</TD>" & vbCrLf
  .write "<TD ALIGN=RIGHT BGCOLOR=""White"" CLASS=SMALL>" & oFileUpProgressWS.percentage & "%</TD>" & vbCrLf
  .write "</TR>" & vbCrLf
  .write "<TR>" & vbCrLf
  .write "<TD BGCOLOR=""White"" CLASS=SMALL><B>Status</B></TD>"
  .write "<TD COLSPAN=4 BGCOLOR=""White"">" & strStatus & "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf
  .write "</TABLE>" & vbCrLf

  .write "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf
  .write "</TABLE>" & vbCrLf

  .write "</BODY>" & vbCrLf
  .write "</HTML>" & vbCrLf
  
end with
%>  

<SCRIPT LANGUAGE="JavaScript">
if (<%=CInt(bStatus)%> == -1) {
  setTimeout("self.close()", 2000 )
}
</SCRIPT>