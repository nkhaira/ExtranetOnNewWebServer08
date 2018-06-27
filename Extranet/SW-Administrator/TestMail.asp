<%
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")    

    Mailer.QMessage       = False
    Mailer.ReturnReceipt  = False
    Mailer.Priority       = 3
    Mailer.RemoteHost  = "mail.evt.danahertm.com:25"
    Mailer.TimeOut     = 120
    Mailer.WordWrap    = False
    Mailer.WordWrapLen = 150
    
    Mailer.FromName       = "amit.bendre@fluke.com"
    Mailer.FromAddress    = "amit.bendre@fluke.com"  
    Mailer.AddRecipient    "ankur.gautam@fluke.com", "santosh.tembhare@fluke.com"
    
    MailSubject = "Test Mail"
    
    MailMessage = "This is a Test Message" 
    
    
    Mailer.Subject  = "Test Subject"
    Mailer.BodyText = MailMessage
  
    err.clear
  
    ' --------------------------------------------------------------------------------------
    ' Attempt to Send Email Notification
    ' --------------------------------------------------------------------------------------
    
    if Mailer.SendMail then
      MailSent = True 
    else
      MailSent = False
      ErrorMessage = ErrorMessage  & "Error Description: " & MailMessage & vbCrLf & vbCrLf & Mailer.Response
    end if   
  

%>