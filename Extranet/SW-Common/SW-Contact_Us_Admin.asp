<%

' --------------------------------------------------------------------------------------
' Site Wide Contact Us Form
' --------------------------------------------------------------------------------------

Site_ID = request("Site_ID")

SQL = "SELECT * FROM Site Where Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Groups = rsSite("Site_Code")

MailFromName    = rsSite("FromName")
MailFromAddress = rsSite("FromAddress")
MailReplyToName = Login_FirstName & " " & Login_LastName
MailReplyTo     = Login_EMail
MailCCName      = rsSite("MailCCName")
MailCC          = rsSite("MailCC")
MailBCCName     = rsSite("MailBCCName")
MailBCC         = rsSite("MailBCC")

rsSite.close
set rsSite = nothing

  'Mailer.ClearAllRecipients
  msg.To = ""
  
  'Mailer.QMessage       = False
  
  'Mailer.ReturnReceipt  = False
  'Mailer.Priority       = 1
  
  'Mailer.FromName       = MailFromName
  'Mailer.FromAddress    = MailFromAddress
  msg.From = """" & MailFromName & """" & MailFromAddress 
  
  'Mailer.ReplyTo        = MailReplyTo
  msg.ReplyTo = MailReplyTo
  
  ' Site Administrator
  if request("MessageTo") = "1" then
    'Mailer.AddRecipient   MailFromName, MailFromAddress

    if not isblank(msg.To) then
      msg.To = msg.To & ";" & """" & MailFromName & """" & MailFromAddress
    else
      msg.To = """" & MailFromName & """" & MailFromAddress
    end if
  end if

  ' Support Group
  if request("MessageTO") = "3" and not isblank(request("Fcm_Name")) and not isblank(request("Fcm_EMail")) then
    'Mailer.AddRecipient  request("Fcm_Name"), request("Fcm_EMail")

    if not isblank(msg.To) then
      msg.To = msg.To & ";" & """" & request("Fcm_Name") & """" & request("Fcm_EMail")
    else
      msg.To = """" & request("Fcm_Name") & """" & request("Fcm_EMail")
    end if
  end if  
  
  if not isblank(MailBCC) then
    'Mailer.AddBCC         MailBCCName, MailBCC
    msg.Bcc = """" & MailBCCName & """" & MailBCC
  end if
  
  ' Debug

  'Mailer.Subject =       "Contact Us - Information Request"
  msg.Subject = "Contact Us - Information Request"
 
  MailMessage = "This is an automated message from the" & vbCrLf
  MailMessage = MailMessage & MailFromName & " Extranet Server at:" & vbCrLf
  MailMessage = MailMessage & "http://" & Request("SERVER_NAME") & "/" & lcase(groups) & vbCrLf & vbCrLf
  MailMessage = MailMessage & vbCrLf & "Question(s) submitted to the Site's "
  select case request("MessageTO")
    case "1"
      MailMessage = MailMessage & "Administrator"
    case "3"
      MailMessage = MailMessage & "Support Group or Account Manager"
    case else
      MailMessage = MailMessage & "Webmaster"
  end select    
  MailMessage = MailMessage & ", by:" & vbCrLf & vbCrLf
  MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf
  MailMessage = MailMessage & "Name: " & Login_FirstName &  " " & Login_LastName & vbCrLf
  MailMessage = MailMessage & "Company Name: " & Login_Company & vbCrLf
  MailMessage = MailMessage & "City / Country: " & Login_City & ", " & Login_Country & vbCrLf
  MailMessage = MailMessage & "EMail Address: " & Login_EMail & vbCrLf
  MailMessage = MailMessage & "Phone Number: " & Login_Business_Phone & " " & Login_Business_Extension & vbCrLf & vbCrLf
  MailMessage = MailMessage & Login_FirstName & " writes:" & vbCrLf & vbCrLf  
  MailMessage = MailMessage & request("UserMessage") & vbCrLf & vbCrLf
  MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf  
  MailMessage = MailMessage & "Sincerely," & vbCrLf &vbCrLf & "The " & MailFromName & " Support Team"

  'Mailer.BodyText = MailMessage
  msg.TextBody = MailMessage

  'if Not Mailer.SendMail then
  '  response.write Translate("This site is currently having technical difficulties sending your &quot;Contact Us&quot; - Request.",Login_Language,conn) & "<BR><BR>"
  '  response.write Translate("For your convenience, here is a copy of the message you tried to send, so that you can copy and paste it into another email service.",Login_Language,conn) & "<BR><BR>"
  '  response.write Translate("Error Message",Login_Language,conn) & " " & Mailer.Response & "<BR><BR>"
  '  response.write replace(MailMessage,vbCrLf,"<BR>") & "<BR><BR>"
  '  if request("MessageTo") = "1" then
  '    response.write Translate("Re-send your EMail to:",Login_Language,conn) & " <A HREF=""MailTo:" & MailFromAddress & """>" & MailFromName & "</A>"
  '  elseif request("MessageTo") = "3" then
  '    response.write Translate("Re-send your EMail to:",Login_Language,conn) & " <A HREF=""MailTo:" & Fcm_EMail & """>" & Fcm_Name & "</A>"
  '  else
  '    response.write Translate("Re-send your EMail to:",Login_Language,conn) & " <A HREF=""MailTo:" & MailBCCAddress & """>" & MailBCCName & "</A>"
  '  end if        
  'else
  '  response.write Translate("Your &quot;Contact Us&quot; - Request was sent - Thank You!",Login_Language,conn)
  'end if

  msg.Configuration = conf
  On Error Resume Next
  msg.Send
  If Err.Number = 0 then
    'Success
    response.write Translate("Your &quot;Contact Us&quot; - Request was sent - Thank You!",Login_Language,conn)
  Else
    'Fail
    response.write Translate("This site is currently having technical difficulties sending your &quot;Contact Us&quot; - Request.",Login_Language,conn) & "<BR><BR>"
    response.write Translate("For your convenience, here is a copy of the message you tried to send, so that you can copy and paste it into another email service.",Login_Language,conn) & "<BR><BR>"
    response.write Translate("Error Message",Login_Language,conn) & " " & Err.description & "<BR><BR>"
    response.write replace(MailMessage,vbCrLf,"<BR>") & "<BR><BR>"
    if request("MessageTo") = "1" then
      response.write Translate("Re-send your EMail to:",Login_Language,conn) & " <A HREF=""MailTo:" & MailFromAddress & """>" & MailFromName & "</A>"
    elseif request("MessageTo") = "3" then
      response.write Translate("Re-send your EMail to:",Login_Language,conn) & " <A HREF=""MailTo:" & Fcm_EMail & """>" & Fcm_Name & "</A>"
    else
      response.write Translate("Re-send your EMail to:",Login_Language,conn) & " <A HREF=""MailTo:" & MailBCCAddress & """>" & MailBCCName & "</A>"
    end if
  End If

  err.clear
  Set MailMessage = nothing
  Set conf = Nothing
  Set msg = Nothing

' --------------------------------------------------------------------------------------

%>
