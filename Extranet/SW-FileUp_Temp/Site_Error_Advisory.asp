  
<%
' If CONN Object is unavailable, send advisory email

set Mailer = Server.CreateObject("SMTPsvg.Mailer")

%>
<!--#include virtual="/connections/connection_email.asp"-->
<!--#include virtual="/connections/connection_email_timeout.asp"-->
<%
    
Mailer.QMessage       = False
Mailer.ReturnReceipt  = False
Mailer.Priority       = 3

Mailer.FromName       = UCase(request.ServerVariables("SERVER_NAME"))
Mailer.FromAddress    = "WebMail@Fluke.com"
'Mailer.ReplyTo        = "Kelly.Whitlock@Fluke.com"     
'Mailer.AddBCC           "Kelly Whitlock", "Kelly.Whitlock@Fluke.com"  ' Domain Administrator
Mailer.ReplyTo        = "ExtranetAlerts@Fluke.com"     
Mailer.AddBCC           "Extranet Group", "ExtranetAlerts@Fluke.com"  ' Domain Administrator

Mailer.Subject        = "Unable to connect to CONN object"
Mailer.BodyText       = "Automated Advisory from " & UCase(request.ServerVariables("SERVER_NAME"))

Mailer.SendMail

set Mailer = nothing

response.write "<DIV ALIGN=CENTER>We are sorry, but " & UCase(request.ServerVariables("SERVER_NAME")) & " is having technical difficulties.<BR>A notification has been sent to the Webmaster.<BR>Please try this site again in a few minutes.</DIV>"

%>