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

Dim Script_Debug
Script_Debug = false

if Script_Debug then
  response.write "<HTML><BODY><TABLE Border=1>"
  for each item in request.querystring
    response.write "<TR><TD>" & item & "</TD><TD>" & request.querystring(item) & "</TD></TR>"
  next
  response.write "</TABLE></BODY></HTML>"
  response.flush
  response.end
end if    

Dim oFileUpProgressClient
Dim strStatus, bStatus, BTime, ETime, Seconds, Item, FCnt

'--- Instantiate the FileUpProgress object

Set oFileUpProgressClient = Server.CreateObject("Softartisans.FileUpProgress")

'--- Assign the progress IDs from the querystring to the objects

oFileUpProgressClient.ProgressID = CInt(request.querystring("cProgressID"))
BTime = request.querystring("BTime")
FCnt  = request.querystring("FCnt")

with response

  .write "<HTML>" & vbCrLf
  .write "<HEAD>" & vbCrLf

	'--- If the transfer is not yet complete, continue to refresh the page

  'For Each Item In Request.ServerVariables 
  '  response.write item & ": " & Request.ServerVariables(item) & "<BR>"
  'next
  
	if oFileUpProgressClient.Percentage = 0 then
		.write("<Meta HTTP-EQUIV=""Refresh"" CONTENT=1>")
		strStatus = "<SPAN CLASS=SMALLRED>Waiting for File Upload to Begin...</SPAN>"
  elseif oFileUpProgressClient.Percentage < 100 then
		.write("<Meta HTTP-EQUIV=""Refresh"" CONTENT=1>")
		strStatus = "<SPAN CLASS=SMALLRED>File Upload in Progress... Total Files: (" & FCnt & ")</SPAN>"          
	else
    bStatus = true
		strStatus = "<SPAN CLASS=SMALL><FONT COLOR=""#006600"">File Upload Complete - Creating Archive File.</FONT></SPAN>"
	end if

  .write "<TITLE>SiteWide Upload Progress Indicator</TITLE>" & vbCrLf
  .write "</HEAD>" & vbCrLf
  .write "<BODY>" & vbCrLf
  
  .write "<FORM NAME=""Progress"" ID=""Progress"">" & vbCrLf
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
	
  ' Because refresh is every 1 second, use Total Bytes as opposed to Transferredbytes when Percentage = 100%
  ' So the user is not confused by they do not match.
  .write "<TD ALIGN=RIGHT BGCOLOR=""White"" CLASS=SMALL>"
  if oFileUpProgressClient.percentage = 100 then
    .write FormatNumber((oFileUpProgressClient.TotalBytes / 1024),0)
  else
    .write FormatNumber((oFileUpProgressClient.TransferredBytes / 1024),0)
  end if
  .write "</TD>" & vbCrLf
	
  .write "<TD ALIGN=RIGHT BGCOLOR=""White"" CLASS=SMALL>" & FormatNumber((oFileUpProgressClient.TotalBytes / 1024),0) & "</TD>" & vbCrLf
  .write "<TD ALIGN=RIGHT BGCOLOR="""
  select case oFileUpProgressClient.Percentage
    case 0
      .write "White"
    case 100
      .write "#33FF66"
    case else
      .write "#FFFF33"
  end select
  .write """ CLASS=SMALL>" & FormatNumber(oFileUpProgressClient.Percentage,1) & "%</TD>" & vbCrLf
  .write "</TR>" & vbCrLf
  
  .write "<TR>" & vbCrLf
  .write "<TD BGCOLOR=""White"" CLASS=SMALL><B>Status</B></TD>"
  .write "<TD COLSPAN=4 BGCOLOR=""White"">" & strStatus & "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf
  
  .write "<TR>" & vbCrLf
  .write "<TD BGCOLOR=""White"" CLASS=SMALL><B>Elapsed Time</B></TD>"
  .write "<TD COLSPAN=4 BGCOLOR=""White"" CLASS=Small>" & ConvertToTime(DateDiff("s", CDate(BTime), Now())) & "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf

  .write "</TABLE>" & vbCrLf

  .write "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf


  .write "</TABLE>" & vbCrLf
  .write "</FORM>" & vbCrLf

  .write "</BODY>" & vbCrLf
  .write "</HTML>" & vbCrLf
  
end with

function ConvertToTime(Seconds)

  Dim lHrs
  Dim lMinutes
  Dim lSeconds
  
  lSeconds = Seconds
  
  lHrs = Int(lSeconds / 3600)
  lMinutes = (Int(lSeconds / 60)) - (lHrs * 60)
  lSeconds = Int(lSeconds Mod 60)
  
  Dim sAns
  
  If lSeconds = 60 Then
      lMinutes = lMinutes + 1
      lSeconds = 0
  End If
  
  If lMinutes = 60 Then
      lMinutes = 0
      lHrs = lHrs + 1
  End If
  
  if Len(lHrs) = 1 then
    lHrs = "0" & lHrs
  end if

  if Len(lMinutes) = 1 then
    lMinutes = "0" & lMinutes
  end if
  
  if Len(lSeconds) = 1 then
    lSeconds = "0" & lSeconds
  end if  
  
  ConvertToTime = lHrs & ":" & lMinutes & ":" & lSeconds
  
end function

%>  
<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>
<!--
if (document.all) {
  self.resizeTo((screen.width / 2) ,screen.height / 4);
}
else if (document.layers || document.getElementById) {
  if (self.outerHeight < screen.height || self.outerWidth < screen.width){
    self.outerWidth = (screen.width / 2);
    self.outerHeight = screen.height / 4;        
  }
}
self.moveTo((screen.width / 4),100);

if (<%=CInt(bStatus)%> == -1) {
  setTimeout("self.close()", 4000 )
}
//-->
</SCRIPT>