  
<%
' If CONN Object is unavailable, send advisory email
'set Mailer = Server.CreateObject("SMTPsvg.Mailer") 
  'adding new email method
%>
<!--#include virtual="/connections/connection_email_new.asp"-->
<%
   
'Mailer.QMessage       = False
'Mailer.ReturnReceipt  = False
'Mailer.Priority       = 3

'Mailer.FromName       = UCase(request.ServerVariables("SERVER_NAME"))
'Mailer.FromAddress    = "WebMail@Fluke.com"
msg.From = """" & UCase(request.ServerVariables("SERVER_NAME")) & """" & "WebMail@Fluke.com"
'Mailer.ReplyTo        =      
msg.ReplyTo = "ibgassetsdev@Fluke.com"
'Mailer.AddBCC           "Assets Development Team", "ibgassetsdev@Fluke.com"  ' Domain Administrator
msg.Bcc = """" & "Assets Development Team" & """" & "ibgassetsdev@Fluke.com"
'Mailer.Subject        = "Unable to connect to CONN object"
msg.Subject = "Unable to connect to CONN object"
'Mailer.BodyText       = "Automated Advisory from " & UCase(request.ServerVariables("SERVER_NAME"))
msg.TextBody = "Automated Advisory from " & UCase(request.ServerVariables("SERVER_NAME"))

'Mailer.SendMail

msg.Configuration = conf

On Error Resume Next
msg.Send
If Err.Number = 0 then
'Success
Else
'ErrorMessage = ErrorMessage & vbCrLf & "<LI>" & Translate("Send Email Failure.",Login_Language,conn) & "<BR><BR>" & Translate("Error Description",Login_Language,conn) & ": " & Err.description & ". " & Translate("Report this error to the Site Administrator.",Login_Language,conn) & "</LI>"  
End If

'set Mailer = nothing
Set conf = Nothing
Set msg = Nothing

response.write "<DIV ALIGN=CENTER>We are sorry, but " & UCase(request.ServerVariables("SERVER_NAME")) & " is having technical difficulties.<BR>A notification has been sent to the Webmaster.<BR>Please try this site again in a few minutes.</DIV>"

%>