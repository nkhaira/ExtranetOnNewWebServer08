  
<%
' If CONN Object is unavailable, send advisory email

set Mailer = Server.CreateObject("SMTPsvg.Mailer")

%>
<!--#include virtual="/connections/connection_email.asp"-->
<!--#include virtual="/connections/connection_email_timeout.asp"-->
<%
    
Mailer.QMessage = False
Mailer.ReturnReceipt  = False
Mailer.Priority       = 3

Mailer.FromName       = "Support.Fluke.com"
Mailer.FromAddress  = "WebMail@Fluke.com"
Mailer.ReplyTo      = "Kelly.Whitlock@Fluke.com"     
Mailer.AddBCC         "Kelly Whitlock", "Kelly.Whitlock@Fluke.com"  ' Domain Administrator

Mailer.Subject  = "Unable to connect to CONN object"
Mailer.BodyText = "Automated Advisory from Support.fluke.com"
Mailer.SendMail

set Mailer = nothing
<%