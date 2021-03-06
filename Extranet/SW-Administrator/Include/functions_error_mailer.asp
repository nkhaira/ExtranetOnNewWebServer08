<%
' --------------------------------------------------------------------------------------
' Author: D. Whitlock
' Date:   02/01/2000
' Title:  Generic Error Handler with Webmaster Email Notification
' --------------------------------------------------------------------------------------

Function Error_Handler(err.number, err.line, err.description, path)
  
  Dim Error_Number
  Dim Error_Line
  Dim Error_Description
  Dim Error_Path
  
  Error_Number      = err.number
  Error_Line        = err.line
  Error_Description = err.description
  Error_Path        = Path
  
  if Error_Number then
    
    Dim MailMessage
    Dim MailSubject
    Dim ErrorMessage
    Dim MailSent
    
    ' --------------------------------------------------------------------------------------
    ' Open Connection to Mail Server
    ' --------------------------------------------------------------------------------------
    
    'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")    
    'adding new email method
    %>
    <!--#include virtual="/connections/connection_email_new.asp"-->
    <%
    
    ' --------------------------------------------------------------------------------------
    ' Compose Error Notification Message
    ' --------------------------------------------------------------------------------------
  
    'Mailer.QMessage       = False
    'Mailer.ReturnReceipt  = False
    'Mailer.Priority       = 3
    
    'Mailer.FromName       = "Support.Fluke.com"
    'Mailer.FromAddress    = "Webmail@Fluke.com"  
    'Mailer.AddRecipient     "Kellie Whitlock", "Kellie.Whitlock@Fluke.com"
    msg.From = """support.fluke.com""" & "webmail@fluke.com"
    msg.To = """Santosh Tembhare""" & "santosh.tembhare@fluke.com"
    
    MailSubject = "Priority 0: Support.Fluke.com - General Site Error Notification"
    
    MailMessage = "This is an automated Error Notification Message from SUPPORT.FLUKE.COM" & vbCrLf & vbCrLf
    MailMessage = MailMessage & "Date: "              & Now()             & vbCrLf
    MailMessage = MailMessage & "Script : "           & Error_Path        & vbCrLf
    MailMessage = MailMessage & "Error Number : "     & Error_Number      & vbCrLf
    MailMessage = MailMessage & "Error Line   : "     & Error_Line        & vbCrLf
    MailMessage = MailMessage & "Error Description: " & Error_Description & vbCrLf & vbCrLf
    
    'Mailer.Subject  = MailSubject
    'Mailer.BodyText = MailMessage

    msg.Subject = MailSubject
    msg.TextBody = MailMessage
  
    err.clear
  
    ' --------------------------------------------------------------------------------------
    ' Attempt to Send Email Notification
    ' --------------------------------------------------------------------------------------
    
    'if Mailer.SendMail then
    '  MailSent = True 
    'else
    '  MailSent = False
    '  ErrorMessage = ErrorMessage  & "Error Description: " & MailMessage & vbCrLf & vbCrLf & Mailer.Response
    'end if   

    msg.Configuration = conf
    On Error Resume Next
    msg.Send
    If Err.Number = 0 then
      MailSent = True
    Else
      MailSent = False
      ErrorMessage = ErrorMessage  & "Error Description: " & MailMessage & vbCrLf & vbCrLf & Err.Description
    End If

	Set conf = Nothing
	Set msg = Nothing
  
    ' --------------------------------------------------------------------------------------
    ' Display Error Screen
    ' --------------------------------------------------------------------------------------
    
    response.clear
  
    Screen_Title = "Support.Fluke.Com - General Site Error Report Screen"
    Navigation = false
    Content_Width = 95  ' Percent
  
    %>
    <!--#include virtual="/include/sw-header.asp"-->
    <%
    
    response.write Replace(MailMessage,vbCrLf,"<BR>")
  
    if MailSent then
      response.write "<B>The above error <FONT COLOR=""Red""><U>has been</U></FONT> reported to</B>:<BR><BR>"
      response.write "<A HREF=""mailto:David.Whitlock@fluke.com"">K. David Whitlock</A>, Fluke Technical Webmaster.<BR>"    
    else
      response.write "<FONT COLOR=""Red""><B>Please copy the above error message and send to</B></FONT>:<BR><BR>"
      response.write "<A HREF=""mailto:David.Whitlock@fluke.com"">K. David Whitlock</A>, Fluke Technical Webmaster.<BR>"
    end if    
  
    response.write"<BR><BR><BR><BR>"
    
    %>
    <!--#include virtual="/include/sw-footer.asp"-->
    <%
    
  end if 
    
end function    

' --------------------------------------------------------------------------------------
%>
