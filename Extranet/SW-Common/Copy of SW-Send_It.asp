<%
' --------------------------------------------------------------------------------------
' Author: D. Whitlock
' Date:   02/01/2000
' Script: SW-Send_IT.asp
'
' EMail - Sends file selected by user to the users email address as an attachment and
' also logs to Activity table.
' --------------------------------------------------------------------------------------

sub Send_It

  ' --------------------------------------------------------------------------------------
  ' Configure EMail Header Information
  ' --------------------------------------------------------------------------------------

  Mailer.ClearAllRecipients
  Mailer.ReturnReceipt  = True
  Mailer.Priority       = 3

  ' --------------------------------------------------------------------------------------
  ' Get Site Information
  ' --------------------------------------------------------------------------------------

  SQL = "SELECT * FROM Language WHERE Language.Code='" & Login_Language & "'"
  Set rsLanguage = Server.CreateObject("ADODB.Recordset")
  rsLanguage.Open SQL, conn, 3, 3

  Mailer.CustomCharSet = rsLanguage("Name_CharSet")

  rsLanguage.close
  set rsLanguage = nothing

  SQL = "SELECT * FROM Site Where Site.ID=" & Site_ID
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3

  if not rsSite.EOF then

    Mailer.FromName      = rsSite("Site_Description")
    Mailer.FromAddress   = "Fluke-Info@Fluke.com"           ' rsSite("FromAddress")
    Mailer.ReplyTo       = "Fluke-Info@Fluke.com"           ' rsSite("ReplyTo")

    if MailBCC <> "" then
      Mailer.AddBCC        rsSite("MailBCCName"), rsSite("MailBCC")
    end if
  end if
  
  rsSite.close
  set rsSite = nothing
  
  ' --------------------------------------------------------------------------------------
  ' Get User Information
  ' --------------------------------------------------------------------------------------

'  Mailer.AddRecipient     Login_FirstName & " " & Login_LastName, Login_Email
  Mailer.AddRecipient     "Kellie" & " " & "Whitlock", "whitlock@accessone.com"

  MailSubject     = Translate("File Requested",Login_Language,conn)
  
  MailMessage     = Translate("This is an automated message from the",Login_Language,conn) & " "
  MailMessage     = MailMessage & Translate(Site_Description,Login_Language,comm) & " - " & Translate("Extranet Support Site",Login_Language,conn) & vbCrLf
  MailMessage     = MailMessage & "http://support.fluke.com/" & lcase(Site_Code) & vbCrLf
  
  MailMessage = MailMessage & vbCrLf & "--------------------------------------------" & vbCrLf
  MailMessage = MailMessage & vbCrLf & Translate("Attached to this email message is the file that you have requested.",Login_Language,conn) & vbCrLf & vbCrLf
  
' Call Get_FileData
  
  MailMessage = MailMessage & Translate("Title",Alt_Language,conn) & " : " & File_Title & vbCrLf
  MailMessage = MailMessage & Translate("Description",Alt_Language,conn)  & " : " & File_Description & vbCrLf
  MailMessage = MailMessage & Translate("File Type",Alt_Language,conn)  & " : " & File_Type & vbCrLf & vbCrLf

  MailMessage = MailMessage & Translate("Special Instructions",Login_Language,conn) & ":" & vbCrLf
  if LCase(File_Extension) = "exe" then
    MailMessage = MailMessage & Translate("This file is either a self-extracting archive file or stand-alone application program.",Login_Language,conn) & vbCrLf
    MailMessage = MailMessage & Translate("Save this file to a temporary location on your local drive.",Login_Language,conn) & vbCrLf
    MailMessage = MailMessage & Translate("Then click on the filename.extension to un-archive or to execute.",Login_Language,conn) & vbCrLf
  elseif LCase(File_Extension) = "zip" then
    MailMessage = MailMessage & Translate("This file is an archive file, that must be un-archived before you can view the file&acute;s contents.",Login_Language,conn) & vbCrLf
    MailMessage = MailMessage & Translate("If you need un-archiving software, visit our site and look under the Library category Site Utilities for links to these software utility programs.",Login_Language,conn) & vbCrLf
  else
    MailMessage = MailMessage & Translate("Some files may require a special viewer and/or a application plug-ins to view this file on your computer.",Login_Language,conn) & vbCrLf
    MailMessage = MailMessage & Translate("If you need any of these special viewers and/or application plug-ins, visit our site and look under the [Library] category [Site Utilities] for links to these file viewers and/or special application plug-ins.",Login_Language,conn) & vbCrLf
  end if

  MailMessage = MailMessage & vbCrLf & "--------------------------------------------" & vbCrLf & vbCrLf
  
  MailMessage = MailMessage & Translate("Sincerely",Alt_Language,conn) & "," & vbCrLf & vbCrLf & Translate(Site_Description,Login_Language,conn) & " - " & Translate("Support Team",Alt_Language,conn)
    
  Mailer.Subject  = MailSubject
  Mailer.BodyText = MailMessage

  ' Clean up BackURL and BackURLSecure.  The SAID parameter should be last in the request.querystring
  
  BackURL       = Mid(BackURL,1,Instr(1,BackURL,"&SAID=")-1)
  BackURLSecure = Mid(BackURL,1,Instr(1,BackURLSecure,"&SAID=")-1)
  
  if Mailer.SendMail then
  ' Success
  else
    ErrorMessage = ErrorMessage & vbCrLf & "<LI>" & Translate("Sorry, but we seem to be having difficulty trying to the file that you requested.  Please try again at a later time, or download the file directly by clicking on the [Download] icon.",Login_Language,conn) & ".<BR><BR>" & Translate("Error Description",Login_Language,conn) & ": " & Mailer.Response & ". " & Translate("Send any errors noted to be reported to the Webmaster",Login_Language,conn) & ". </LI>"   
  end if   

end sub
  
%>

