<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// EMAIL OBJECT
//
// ======================================================================

function SWEFEmail
(
)
{
  this.setToName ("");
  this.setToAddress ("");
  this.setFromName ("");
  this.setFromAddress ("");
  this.setSubject ("");
  this.setBody ("");

  return this;
}

// ======================================================================
//
// Interface to private member variables.
//
// ======================================================================

function getToName_eml_disc
(
)
{
  return String (this._toName);
}

function setToName_eml_disc
(
 strNewToName
)
{
  this._toName = String (strNewToName);
  return;
}

function getToAddress_eml_disc
(
)
{
  return String (this._toAddress);
}

function setToAddress_eml_disc
(
 strNewToAddress
)
{
  this._toAddress = String (strNewToAddress);
  return;
}

function getFromName_eml_disc
(
)
{
  return String (this._fromName);
}

function setFromName_eml_disc
(
 strNewFromName
)
{
  this._fromName = String (strNewFromName);
  return;
}

function getFromAddress_eml_disc
(
)
{
  return String (this._fromAddress);
}

function setFromAddress_eml_disc
(
 strNewFromAddress
)
{
  this._fromAddress = String (strNewFromAddress);
  return;
}

function getSubject_eml_disc
(
)
{
  return String (this._subject);
}

function setSubject_eml_disc
(
 strNewSubject
)
{
  this._subject = String (strNewSubject);
  return;
}

function getBody_eml_disc
(
)
{
  return String (this._body);
}

function setBody_eml_disc
(
 strNewBody
)
{
  this._body = String (strNewBody);
  return;
}

// ======================================================================
//
// Main object methods.
//
// ======================================================================

function getFromEmailLink_eml_disc
(
)
{
  var strHTMLout = "";

  if (config.ADMINSWITCH_ShowEmailAddresses)
    {
      var strLinkURL = "mailto:" + this.getFromAddress ();
      var strLinkText = "";
      strLinkText += config.USERTEXT_SHOW_PopupEmailPrefix;
      strLinkText += this.getFromName ();
      strLinkText += " (";
      strLinkText += this.getFromAddress ();
      strLinkText += ")";

      strHTMLout += A_open_html_disc (strLinkURL, strLinkText);
      strHTMLout += this.getFromName ();
      strHTMLout += SWEFHTML.A_close ();
    }
  else
    {
      strHTMLout = this.getFromName ();
    }

  return strHTMLout;
}

function getHTMLMessageBody_eml_disc
(
 msgCurrentMessage,
 strTitle,
 strBodyText,
 strSignature
)
{
  var strMessageBody = "";
  strMessageBody += SWEFHTML.DTD ();
  strMessageBody += SWEFHTML.HTML_open ();

  strMessageBody += SWEFHTML.BODY_open ("#FF00FF",
					undefined_disc,
					"#0000FF",
					"#FF0000",
					"#660066");

  strMessageBody += SWEFHTML.TABLE_open (config.ADMINSETTING_TableBorderSize,
					 undefined_disc,
					 undefined_disc,
					 undefined_disc,
					 0,
					 0);

  strMessageBody += SWEFHTML.TR_open ();
  strMessageBody += SWEFHTML.TD_open (120,
				      undefined_disc,
				      "Silver");
  strMessageBody += SWEFHTML.DIV_open (undefined_disc,
				       undefined_disc,
				       "center");
  strMessageBody += SWEFHTML.A_open ("http://" + Request("SERVER_NAME"));
  strMessageBody += SWEFHTML.IMG ("http://" + Request("SERVER_NAME") + "/images/FlukeLogo3.gif",
				  "Fluke Forums",
				  0,
				  120,
				  60);
  strMessageBody += SWEFHTML.A_close ();
  strMessageBody += SWEFHTML.DIV_close ();
  strMessageBody += SWEFHTML.TD_close ();
  strMessageBody += SWEFHTML.TD_open (468, undefined_disc, "Silver");
  strMessageBody += SWEFHTML.FONT_open (5, "Tahoma", "#FF00FF");
  strMessageBody += SWEFHTML.NBSP ();
  strMessageBody += strTitle;
  strMessageBody += SWEFHTML.FONT_close ();
  strMessageBody += SWEFHTML.TD_close ();
  strMessageBody += SWEFHTML.TR_close ();
  strMessageBody += SWEFHTML.TABLE_close ();

  strMessageBody += SWEFHTML.BLOCKQUOTE_open ();
  strMessageBody += SWEFHTML.FONT_open (-1, "Tahoma");
  strMessageBody += strBodyText;
  strMessageBody += SWEFHTML.FONT_close ();

  strMessageBody += SWEFHTML.P_open ();
  strMessageBody += SWEFHTML.NBSP ();
  strMessageBody += SWEFHTML.P_close ();
  strMessageBody += SWEFHTML.BLOCKQUOTE_close ();

  strMessageBody += strSignature;

  strMessageBody += SWEFHTML.P_open ();
  strMessageBody += SWEFHTML.NBSP ();
  strMessageBody += SWEFHTML.P_close ();

  return strMessageBody;
}

function sendHTML_eml_disc ()
{
  var strFrom = this.getFromName () + " <" + this.getFromAddress () + ">";
  var strTo = this.getToAddress ();

  sendHTMLMailMessage_eml_disc (strTo,
				strFrom,
				this.getSubject (),
				this.getBody (),
				config.getSiteBaseURL ());

  return;
}

function send_eml_disc ()
{
  var strFrom = this.getFromName () + " <" + this.getFromAddress () + ">";
  var strTo = this.getToAddress ();

  sendRegularMailMessage_eml_disc (strTo,
				   strFrom,
				   this.getSubject (),
				   this.getBody ());

  return;
}

SWEFEmail.prototype.getToName = getToName_eml_disc;
SWEFEmail.prototype.setToName = setToName_eml_disc;
SWEFEmail.prototype.getToAddress = getToAddress_eml_disc;
SWEFEmail.prototype.setToAddress = setToAddress_eml_disc;
SWEFEmail.prototype.getFromName = getFromName_eml_disc;
SWEFEmail.prototype.setFromName = setFromName_eml_disc;
SWEFEmail.prototype.getFromAddress = getFromAddress_eml_disc;
SWEFEmail.prototype.setFromAddress = setFromAddress_eml_disc;
SWEFEmail.prototype.getSubject = getSubject_eml_disc;
SWEFEmail.prototype.setSubject = setSubject_eml_disc;
SWEFEmail.prototype.getBody = getBody_eml_disc;
SWEFEmail.prototype.setBody = setBody_eml_disc;

SWEFEmail.prototype.getFromEmailLink = getFromEmailLink_eml_disc;
SWEFEmail.prototype.getHTMLMessageBody = getHTMLMessageBody_eml_disc;
SWEFEmail.prototype.sendHTML = sendHTML_eml_disc;
SWEFEmail.prototype.send = send_eml_disc;
</SCRIPT>

<%

Sub sendRegularMailMessage_eml_disc (messageTo, messageFrom, messageSubject, messageBody)

'response.write "<BR>"
'response.write "Regular mail<BR>"
'response.write MessageTO & "<BR>"
'response.write MessageFrom & "<BR>"
'response.write MessageSubject & "<BR>"
'response.write MessageBody & "<BR>"
'response.write "<BR>"

  On Error Resume Next

  Dim Mailer
  'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
 
  'adding new email method
  %>
  <!--#include virtual="/connections/connection_email_new.asp"-->
  <%

  'Mailer.QMessage       = False
  'Mailer.ClearAllRecipients 
	'Mailer.ReturnReceipt  = False
 	'Mailer.Priority     = 1

  'Mailer.FromName       = Session("Site_Description")
  'Mailer.FromAddress    = Session("Moderator_Email")
  'Mailer.AddRecipient MailTo, MailToFrom

  msg.From = """" & Session("Site_Description") & """" & Session("Moderator_Email")
  msg.To = MailTo & ";" & MailToFrom

  if CInt(Session("Forum_Moderated")) = CInt(True) then
    'Mailer.AddBCC Session("Moderator_Name"), Session("Moderator_EMail")
    msg.Bcc = """" & Session("Moderator_Name") & """" & Session("Moderator_EMail")
  end if
  
  'Mailer.Subject = "Forum" & " - " & messageSubject
  'Mailer.BodyText = messageBody

  msg.Subject = "Forum" & " - " & messageSubject
  msg.TextBody = messageBody

	'if Not Mailer.SendMail then
  '  	response.write "<BR><BR>This site is currently having technical difficulties sending your email response.<BR><BR>"
  '  	response.write Mailer.Response & "<BR><BR>"
	'end if

  msg.Configuration = conf
  On Error Resume Next
  msg.Send
  If Err.Number = 0 then
    'Success
  Else
    response.write "<BR><BR>This site is currently having technical difficulties sending your email response.<BR><BR>"
    response.write Err.Description & "<BR><BR>"
  End If	

'    Mailer.Send
  'Set Mailer = Nothing

End Sub

Sub sendHTMLMailMessage_eml_disc (messageTo, messageFrom, messageSubject, messageBody, baseURL)

'response.write "<BR>"
'response.write "HTML mail<BR>"
'response.write MessageTO & "<BR>"
'response.write MessageFrom & "<BR>"
'response.write MessageSubject & "<BR>"
'response.write MessageBody & "<BR>"
'response.write BaseURL & "<BR>"
'response.write "<BR>"


  On Error Resume Next

  Dim Mailer
  'Set Mailer = Server.CreateObject("SMTPsvg.Mailer") 

  'adding new email method
  %>
  <!--#include virtual="/connections/connection_email_new.asp"-->
  <%
  'Mailer.QMessage       = False
  'Mailer.ContentType    = "text/html;charset="""  & "iso-8859-1" & """"
    
	'Mailer.ClearAllRecipients 
	'Mailer.ReturnReceipt = False
  'Mailer.Priority       = 1
  'Mailer.FromName       = Session("Site_Description")
  'Mailer.FromAddress    = Session("Moderator_Email")

  msg.From = """" & Session("Site_Description") & """" & Session("Moderator_Email")

  'Mailer.AddRecipient MailTo, MailToFrom

  msg.To = """" & MailTo & """" & MailToFrom

  if CInt(Session("Forum_Moderated")) = CInt(True) then
    'Mailer.AddBCC Session("Moderator_Name"), Session("Moderator_EMail")
    msg.To = """" & Session("Moderator_Name") & """" & Session("Moderator_EMail")
  end if

  'Mailer.Subject = "Forum" & " - " & messageSubject
  'Mailer.BodyText = messageBody

  msg.Subject = "Forum" & " - " & messageSubject
  msg.TextBody = messageBody

	'if Not Mailer.SendMail then
  '  	response.write "<BR><BR>This site is currently having technical difficulties sending your email response.<BR><BR>"
  '  	response.write Mailer.Response & "<BR><BR>"
	'end if

  msg.Configuration = conf
  On Error Resume Next
  msg.Send
  If Err.Number = 0 then
    'Success
  Else
    response.write "<BR><BR>This site is currently having technical difficulties sending your email response.<BR><BR>"
    response.write Err.Description & "<BR><BR>"
  End If

  'Set Mailer = Nothing

End Sub

%>
<!--#include virtual="/connections/connection_EMail.asp"-->
