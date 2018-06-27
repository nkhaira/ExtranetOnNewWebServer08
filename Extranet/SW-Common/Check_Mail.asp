<%


%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/connections/connection_EMail.asp"--> 
<%

' --------------------------------------------------------------------------------------
' Connect to Email
' --------------------------------------------------------------------------------------

Set Mailer = Server.CreateObject("SMTPsvg.Mailer") 
%>
<!--#include virtual="/connections/connection_EMail_Timeout.asp"-->  
<%

Mailer.QMessage      = False
Mailer.ReturnReceipt  = False
Mailer.Priority       = 2


Mailer.FromName      = "ankur"
  Mailer.FromAddress   = "WebMail@Fluke.com"
  Mailer.ReplyTo       = "ankur jana"


Mailer.CustomCharSet = "utf-8"
Mailer.CharSet = 2
Mailer.Encoding = 1


Mailer.AddRecipient     "Ankur Jana", "ankur.jana@fluke.com"

Mailer.Subject = "Escolhendo o Fusível Correto para o seu Testador"

  
MailMessage = "As informações contidas neste documento lhe poderão ser úteis.Produto ou série de produtos:Título"

Mailer.BodyText = MailMessage 
Mailer.SendMail 


%>

