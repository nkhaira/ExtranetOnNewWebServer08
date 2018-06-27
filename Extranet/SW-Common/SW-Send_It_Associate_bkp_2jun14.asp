<%
' --------------------------------------------------------------------------------------
' Author: Kelly. Whitlock
' Date:   02/01/2004
' Script: SW-Send_It_Associate.asp
'
' EMail - Sends file selected by email address specified by the user as an attachment and
' also logs to Activity table.
' --------------------------------------------------------------------------------------

Site_ID         = request.form("SendItSiteID")
Login_Language  = request.form("SendItLanguage")

select case LCase(Login_Language)
  case "chi", "zho", "thi", "jpn", "kor"
    Alt_Language = "eng"
  case else
    Alt_Language = LCase(Login_Language)
end select

SendItAccountID = request.form("SendItAccountID")
SendItSiteCode = request.form("SendItSiteCode")
SendItName      = request.form("SendItName")
SendItEmail     = request.form("SendItEmail")
SendItAssetID   = request.form("SendItAssetID")
SendItSubject   = request.form("SendItSubject")
SendItMessage   = request.form("SendItMessage")
SendItMethod    = request.form("SendItMethod")
SendItHow       = request.form("SendItHow")


Send_Flag = True ' Set by script to False if parameter is missing and checked before sending.

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

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------
%>
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_SiteWide

SQL = "SELECT FirstName, LastName, Company, Email FROM UserData WHERE ID=" & SendItAccountID
Set rsAccount = Server.CreateObject("ADODB.Recordset")
rsAccount.Open SQL, conn, 3, 3
     
if not rsAccount.EOF then
  Sender = rsAccount("FirstName") & " " & rsAccount("LastName")
  Company = rsAccount("Company")
  Mailer.FromName      = Sender
  Mailer.FromAddress   = "WebMail@Fluke.com"
  Mailer.ReplyTo       = rsAccount("Email")

else
  Send_Flag = False
end if

rsAccount.close
set rsAccount = nothing
  
' --------------------------------------------------------------------------------------
' Get Language Information
' --------------------------------------------------------------------------------------

SQL = "SELECT * FROM Language WHERE Language.Code='" & Login_Language & "'"
Set rsLanguage = Server.CreateObject("ADODB.Recordset")
rsLanguage.Open SQL, conn, 3, 3


Mailer.CustomCharSet = rsLanguage("Name_CharSet")
Mailer.CharSet = 2
Mailer.Encoding = 2

rsLanguage.close
set rsLanguage = nothing

' --------------------------------------------------------------------------------------
' Get User Information
' --------------------------------------------------------------------------------------

Mailer.AddRecipient     SendItName, SendItEmail

if isblank(SendItSubject) then
  MailSubject = Translate("Requested Information",Login_Language,conn)
else
  MailSubject = SendItSubject
end if

Mailer.Subject = SendItSubject

if not isblank(SendItMessage) then
  MailMessage = MailMessage & SendItMessage & vbCrLf & vbCrLf
else 
  MailMessage = MailMessage & vbCrLf & Translate("Attached to this email message is the file sent to you by",Login_Language,conn) & " " & Sender & vbCrLf & vbCrLf
end if  

SQLAsset = "SELECT * FROM Calendar WHERE ID=" & SendItAssetID
Set rsAsset = Server.CreateObject("ADODB.Recordset")
rsAsset.Open SQLAsset, conn, 3, 3

' Call Get_FileData
  
MailMessage = MailMessage & Translate("Product or Product Series",Login_Language,conn) & ":" & vbCrLf & Translate(rsAsset("Product"),Login_Language,conn) & vbCrLf & vbCrLf
MailMessage = MailMessage & Translate("Title",Alt_Language,conn) & " : " & vbCrLf & rsAsset("Title") & vbCrLf & vbCrLf

if not isblank(rsAsset("Description")) then
  MailMessage = MailMessage & Translate("Description",Alt_Language,conn)  & " : " & vbCrLf & rsAsset("Description") & vbCrLf & vbCrLf
end if  

' Attach File

File_Redirect = "http://" & request.ServerVariables("SERVER_NAME") & "/" & LCase(SendItSiteCode)

select case SendItMethod
  case 13 'Zip Version
    File_Redirect = File_Redirect & "/" & rsAsset("Archive_Name")
  case else
    File_Redirect = File_Redirect & "/" & rsAsset("File_Name")  
end select

select case SendItHow
  case 0    ' As an Attachment
    Attach_File = Server.MapPath(Replace(File_Redirect,"http://" & Request("SERVER_NAME"),""))
    on error resume next
    Mailer.AddAttachment Attach_File
    if err.number then
      MailMessage = MailMessage & Translate("Attachment File Not Found",Login_Language,conn) & vbCrLf & vbCrLf
    end if
    on error goto 0    
  case 1    ' As Link
    MailMessage = MailMessage & Translate("Link to Document",Login_Language,conn) & " : " & vbCrLf & File_Redirect & vbCrLf & vbCrLf
    MailMessage = MailMessage & Translate("Note",Login_Language,conn) & ": " & Translate("If the link above reports a 404 error, the link may have wrapped in the contents of this email.  Copy the entire link and paste it into your brower's address input box, then press [Enter].",Login_Language,conn) & vbCrLf & vbCrLf
end select    
  
rsAsset.close
set rsAsset = nothing
  
MailMessage = MailMessage & vbCrLf & Translate("Sincerely",Alt_Language,conn) & "," & vbCrLf & vbCrLf & Sender & vbCrLf & Company
    
Mailer.BodyText = MailMessage 

if Mailer.SendMail then
  ' Success
  %>
  <SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>
  <!--
    window.blur();
    window.close()
  // -->
  </SCRIPT>
  <%
else
  response.write Mailer.Response
  %>
  <SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>
  <!--
    window.focus();
  // -->
  </SCRIPT>
  <%

end if   

Call Disconnect_SiteWide

%>

