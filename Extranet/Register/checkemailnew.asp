<%
Response.Write("Send Email using CDOSYS")
%>

<!--#include virtual="/include/functions_string.asp"-->

<%
Set msg = Server.CreateObject("CDO.Message")
Set conf= Server.CreateObject("CDO.Configuration")
%>

<%
confURL = "http://schemas.microsoft.com/cdo/configuration/"
with conf
.Fields.Item(confURL & "sendusing") = 2
.Fields.Item(confURL & "smtpserver") = "mail.evt.danahertm.com"
.Fields.Update
end with
%>

<%
with msg
.From = "ankur.jana@fluke.com"
.To = "ankur.jana@fluke.com"
'.Cc = "ankur.jana@fluke.com"
.Bcc = "jana.ankur@gmail.com"
end with
%>

<%
with msg
.Subject = "Send Email using CDOSYS: Testing ASP code on WIN2012"
.TextBody = "This is a plain text email"
.HTMLBody = "<em>Hi, This is a testing HTML email using CDOSYS</b>"
end with
%>

<%
msg.Configuration = conf
msg.Send
Set conf = Nothing
Set msg = Nothing
%>


<br />
<%

Response.Write("Email sent using CDOSYS")
%>
