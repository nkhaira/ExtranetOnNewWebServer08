<!-- #include virtual="/AdminTools/Connections/Connection_email.asp" -->
<html>
<head><title>ASP Mailer Test</title><head>
<body>
<H3>ASP Mailer Test</H3>
<%
if Request.Form("mailSubmit") <> "" then
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		
	Mailer.FromName = "Fluke Testing Inc."
	Mailer.FromAddress = "Webmaster@fluke.com"
	Mailer.RemoteHost = GetEmailServer()
	Mailer.AddRecipient "test user", Trim(Request.Form("email"))
	Mailer.ReturnReceipt = false
	Mailer.ConfirmRead = false
	
	
	strBody = "It looks like this worked." & vbcrlf & vbcrlf
	strBody = strBody & "AspMail version: " & Mailer.Version & vbcrlf & vbcrlf
	strBody = strBody & "Webserver: " & Request.ServerVariables("SERVER_NAME") & vbcrlf & vbcrlf
	strBody = strBody & "Mailer Remote Host: " & Mailer.RemoteHost & vbcrlf & vbcrlf
	strBody = strBody & "Happy ASP Programmer (" & Now & ")"
	Mailer.BodyText = strBody
		
	Mailer.ClearAttachments
	'Mailer.AddAttachment "c:\bentaa\unstdll.dll"
	rem Mailer.AddAttachment "c:\config.sys"
	
	
	if Request.Form("que") = "Q" then
		Mailer.QMessage = True
		smesg = "Mail sent... with queueing"
		Mailer.Subject = "Mail Queue Test"
	Else
		Mailer.Subject = "Mail Test"
		smesg = "Mail sent..."
	End if
	
	if Mailer.SendMail then
		Response.Write smesg
	else
		Response.Write "Mail failure. Check mail host server name and tcp/ip connection...<br>"
		Response.Write Mailer.Response
	end if
Else
	%>
<form id="mailForm" name="mailForm" METHOD="POST">
<input type="hidden" name="mailSubmit" value="">
<BR>Send To: <input type="text" name="email" size="40">
<P><input type="checkbox" name="que" value="Q"> Check to use Queueing
<P><input TYPE="button" name="btnSubmit" Value="Submit" onClick="processForm()">
</form>
<SCRIPT LANGUAGE="Javascript">
function processForm() {
	var dmf = document.mailForm;
	if (dmf.email.value != '') {
		dmf.mailSubmit.value = 'Doit';
		dmf.submit();
	}
	else {
		alert('Come on, give me an email address!');
	}
}
</script>
	<%
End if
%>

</body>
</html>
