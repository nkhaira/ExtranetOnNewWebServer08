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

Mailer.CustomCharSet = "utf-8"
Mailer.CharSet = 2
Mailer.Encoding = 1

Mailer.FromName   = "Joe’s Widgets Corp."
Mailer.FromAddress= "webmail@fluke.com"
Mailer.AddRecipient "Ankur Jana", "ankur.jana@fluke.com"
Mailer.Subject    = "Escolhendo o Fusível Correto para o seu Testador"
Mailer.BodyText   = "As informações contidas neste documento lhe poderão ser úteis.Produto ou série de produtos:Título"
if Mailer.SendMail then
  Response.Write "Mail sent..."
else
  Response.Write "Mail send failure. Error was " & Mailer.Response
end if


%>

