<!-- #INCLUDE file="discussionengine/adoconsts.asp" -->
<!-- #INCLUDE file="discussionengine/discussionengine.asp" -->

<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

config.ADMINSETTING_ForumName = Session("Forum_Title");
config.ADMINSETTING_DatabaseTable = "Forum_Data";

var i,myenv,ServerNodes;
var ServerName = Session("Server_Name");
// Request.ServerVariables('SERVER_NAME')
ServerName = ServerName.toUpperCase();
ServerNodes = ServerName.split(".");
//Array('SUPPORT','DEV','FLUKE','COM');

myenv = 'PRD';

for(i=0;i<ServerNodes.length;i++) {
	if ((ServerNodes[i] == 'DEV') ||
		(ServerNodes[i] == 'PRD') || 
		(ServerNodes[i] == 'STG') ) {
			myenv = ServerNodes[i];
			break;
	}
}

switch (myenv) {
	case 'DEV':
		ServerName = 'EVTIBG03';
		break;
	case 'STG':
		ServerName = 'FLKSTG03';
		break;
	case 'PRD':
	default:
		ServerName = 'FLKPRD18.DATA.IB.FLUKE.COM';
}

config.ADMINSETTING_DatabaseDSN = 'DRIVER={SQL Server}; SERVER=' + ServerName + '; '
config.ADMINSETTING_DatabaseDSN += 'UID=sitewide_web;DATABASE=fluke_sitewide;pwd=tuggy_boy';

config.ADMINSETTING_SiteID = "Site_ID";
config.ADMINSETTING_ForumID = "Forum_ID"; 
config.ADMINSETTING_EmailAdminName = Session("Site_description") + Session("Server_Name");
config.ADMINSETTING_EmailAdminAddress = Session("Moderator_Email");
config.ADMINSETTING_EmailAlertFromName = Session("Site_Description");
config.ADMINSETTING_EmailAlertFromAddress = Session("Moderator_Email");

config.ADMINSETTING_DaysMessagesActive = 90;
config.ADMINSETTING_ExpandFirstNThreads = 0;
config.DATABASE_MaxMessageSize = 10000;

config.ADMINSETTING_DefaultEmbeddedLinkTarget = "_top";

config.ADMINSETTING_ExpandImagePathname = "plus.gif";
config.ADMINSETTING_ExpandImageWidth = 9;
config.ADMINSETTING_ExpandImageHeight = 9;
config.ADMINSETTING_CollapseImagePathname = "minus.gif";
config.ADMINSETTING_CollapseImageWidth = 9;
config.ADMINSETTING_CollapseImageHeight = 9;
config.ADMINSETTING_NoExpandImagePathname = "blank.gif";
config.ADMINSETTING_NoExpandImageWidth = 9;
config.ADMINSETTING_NoExpandImageHeight = 9;

config.CACHE_Enabled = false;

config.ADMINSETTING_TableTitleColumnWidth = 100;
config.ADMINSETTING_TableFieldColumnWidth = 494;
config.ADMINSETTING_TableFullWidth = config.ADMINSETTING_TableTitleColumnWidth + config.ADMINSETTING_TableFieldColumnWidth;
config.ADMINSETTING_SubjectInputboxSize = 66;
config.ADMINSETTING_TextAreaCols = 66;
config.ADMINSETTING_TextAreaRows = 10;

config.ADMINSETTING_NoWrapMessageThreadViews = true;

config.ADMINSETTING_ToolbarImagePathname = "toolbar.gif";
config.ADMINSETTING_ToolbarButtonImagePathname = "button_normal.gif";
config.ADMINSETTING_ToolbarButtonImageUpPathname = "button_up.gif";
config.ADMINSETTING_ToolbarButtonImageDownPathname = "button_down.gif";
config.ADMINSETTING_ToolbarSpacerImagePathname = "toolbar_spacer.gif";

CONST_special_font_face = "Arial,Verdana";
CONST_background_colour = "#FFCC00";
CONST_header_text_colour = "#000000";
CONST_table_style = "";

function formatStrong ()
{
  return String ("<FONT CLASS=MediumBold>" + this + "</FONT>");
}

function formatStrongRed ()
{
  return String ("<FONT CLASS=MediumBoldRed>" + this + "</FONT>");
}

function formatStrongSmall ()
{
  return String ("<FONT CLASS=SmallBold>" + this + "</FONT>");
}

function formatWeak ()
{
  return String ("<FONT CLASS=Medium>" + this + "</FONT>");
}

function formatWeakSmall ()
{
  return String ("<FONT CLASS=Small>" + this + "</FONT>");
}

function formatMessageBody ()
{
  return "<BLOCKQUOTE CLASS=Medium>" + this + "</BLOCKQUOTE>";
}

function customShowForumLink ()
{
  this.showVersion ();

  var HTMLout = SWEFHTML.P_open ()

      HTMLout =  "<FORM NAME=\"Forum_Home\">";
      HTMLout += "<INPUT VALUE=\"Forum Home\" TYPE=\"BUTTON\" onclick=\"location.href='/SW-Forums/'\"";
      HTMLout += " CLASS=\"NavLeftHighlight1\">";
      HTMLout += "</FORM>";

  HTMLout += SWEFHTML.P_close ();
  HTMLout.show ();
    
  return;
}

function HTMLMailMessage (message, title, bodyText, signature)
{
  var strCurrentVirtualPath = config.ADMINSETTING_VirtualPath;
  config.ADMINSETTING_VirtualPath = "/SW-Forums/Default.asp" //+ config.ADMINSETTING_ForumName + "/";

  // Fluke Header Bar
  var messageMastead =  "<TABLE WIDTH=\"100%\" BORDER=\"0\" CELLSPACING=\"0\" CELLPADDING=\"0\" BGCOLOR=\"#000000\" FGCOLOR=\"FFFFFF\">\n";
      messageMastead += "<TR>\n";
      messageMastead += "<TD WIDTH=\"12\" HEIGHT=\"75\">&nbsp;</TD>\n";
      messageMastead += "<TD><FONT STYLE=\"font-weight:Bold; font-size:14pt; font-family: Arial,Verdana; color:#FFCC00; font-style: Normal;\">" + Session("Site_Description") + "</FONT>\n";
      messageMastead += "<BR><FONT STYLE=\"font-weight:Bold; font-size:8.5pt;font-family: Arial,Verdana; color:#FFCC00; font-style: Normal;\">Forum / Discussion Group</FONT>\n";
      messageMastead += "</TD>\n";
      messageMastead += "<TD ALIGN=\"RIGHT\">";
      messageMastead += "<IMG SRC=\"" + Session("Logo") + "\" WIDTH=134 HEIGHT=44 BORDER=0>";
      messageMastead += "</TD>\n";
      messageMastead += "</TR>\n";
      messageMastead += "<TR>\n";
      messageMastead += "<TD COLSPAN=\"10\" STYLE=\"background-color:#FFCC00;background:#FFCC00;color:#FFCC00;\" VSPACE=\"0\" HEIGHT=\"6\"></TD>\n";
      messageMastead += "</TR>\n";
      messageMastead += "</TABLE>\n\n";

  var messageHeadPreTitle   = "";  
  var messageHeadPostTitle  = SWEFHTML.BR();
      messageHeadPostTitle += "<FONT STYLE=\"font-weight:normal; font-size:8.5pt; font-family: Arial; color:#000000; font-style: Normal;\">";
      messageHeadPostTitle += "&nbsp;&nbsp;The following is an automated reply to your message that you posted to the";
      messageHeadPostTitle += SWEFHTML.BR();
      messageHeadPostTitle += "&nbsp;&nbsp;&quot;" + title + "&quot; Forum / Discussion Group at " + Session("Site_Description") + ".";
      messageHeadPostTitle += SWEFHTML.BR();
      messageHeadPostTitle += SWEFHTML.BR();      
      messageHeadPostTitle += "&nbsp;&nbsp;<B>Note:</B> This is a user supported Forum / Discussion Group. ";
      messageHeadPostTitle += SWEFHTML.BR();      
      messageHeadPostTitle += "&nbsp;&nbsp;Unless noted, "  + Session("Site_Description") + " is not responsible accuracy or appropiateness of user posting or replies to your questions.";
      messageHeadPostTitle += "</FONT>";
      messageHeadPostTitle += SWEFHTML.BR(1);

  var messageBody = "";
  messageBody += SWEFHTML.DTD ();
  messageBody += "\n";
  messageBody += SWEFHTML.BASE ("http://" + Request("SERVER_NAME"));
  messageBody += SWEFHTML.HTML_open ();

  messageBody += SWEFHTML.STYLE_open ();
  messageBody += " A:hover { color:\"#FF0000\"; } "
  messageBody += "Medium: { font-weight:Normal; font-size:10pt; font-family: Arial,Verdana; color:#000000; font-style: Normal; }"
  messageBody += SWEFHTML.STYLE_close () + "\n";

  messageBody += SWEFHTML.BODY_open ("#FFFFFF", undefined_disc, "#0000FF", "#FF0000", "#0000FF", 0, 0, 0, 0);
  messageBody += messageMastead;
  messageBody += messageHeadPreTitle;
  messageBody += SWEFHTML.BR(2);
  messageBody += "&nbsp;&nbsp;<FONT style=\"COLOR: #000000; FONT-FAMILY: Arial; FONT-SIZE: 10pt; FONT-STYLE: normal; FONT-WEIGHT: bold\">Forum: " + title; + "</FONT>"
  messageBody += SWEFHTML.BR(2);
  messageBody +=messageHeadPostTitle;
  
  messageBody += SWEFHTML.P_open ();
  messageBody += message.getRenderedHTML ();
  messageBody += SWEFHTML.P_close ();

//  if (message.getMessageID () != 0)
//    {
//      var messageThread = new SWEFThread (message.getMessageID ());
//      messageBody += SWEFHTML.P_open ();
//      messageBody += SWEFHTML.NBSP ();
//      messageBody += SWEFHTML.P_close ();
//      messageBody += messageThread.getFullThread (message.getThreadID ());
//      messageBody += SWEFHTML.P_open ();
//      messageBody += SWEFHTML.NBSP ();
//      messageBody += SWEFHTML.P_close ();

//      var newMessageForm = new SWEFForm (null);
//      newMessageForm.setMessageID (0);
//      newMessageForm.setSiteID (message.getSiteID ());
//      newMessageForm.setForumID (message.getForumID ());            
//      newMessageForm.setParentID (message.getMessageID ());
//      newMessageForm.setThreadID (message.getThreadID ());
//      newMessageForm.setSubject (message.getSubject ());
//      newMessageForm.setSortCode (0);
//      messageBody += newMessageForm.getMessageMailForm (config.getNewPostActionPagePath (),
//							config.USERTEXT_POST_SubmitButton);
//    }

  messageBody += signature;
  messageBody += SWEFHTML.P_open ();
  messageBody += SWEFHTML.NBSP ();
  messageBody += SWEFHTML.P_close ();
  messageBody += SWEFHTML.TD_close ();
  messageBody += SWEFHTML.TR_close ();
  messageBody += SWEFHTML.TABLE_close ();

  config.ADMINSETTING_VirtualPath = strCurrentVirtualPath;

  return messageBody;
}


// View Message
function getCustomRenderedHTML ()
{
  var strHTMLout = "";

  strHTMLout += SWEFHTML.TABLE_open (0, config.ADMINSETTING_TableFullWidth, undefined_disc, "border-style: groove; border-color: #FFFFFF; border-width: 2px 2px 2px 2px;", 0, 0);

  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth,
				  undefined_disc,
				  CONST_background_colour);
  strHTMLout += SWEFHTML.NBSP (1);
  strHTMLout += SWEFHTML.FONT_open (undefined_disc, undefined_disc, "#000000");
  strHTMLout += config.USERTEXT_POST_SubjectPrompt.weak ();
  strHTMLout += SWEFHTML.FONT_close ();
  strHTMLout += SWEFHTML.TD_close ();


  strHTMLout += SWEFHTML.TD_open (undefined_disc,
				  undefined_disc,
				  CONST_background_colour,
				  undefined_disc,
				  undefined_disc,
				  2);
  strHTMLout += SWEFHTML.FONT_open (3, CONST_special_font_face, "#000000");
  strHTMLout += (SWEFHTML.BR ()
             + SWEFHTML.NBSP (1)
             + SWEFHTML.NBSP () + this.getSubject ()).strong ();
  strHTMLout += SWEFHTML.FONT_close ();
  strHTMLout += SWEFHTML.P_open () + SWEFHTML.P_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();

  var emlEmail = new SWEFEmail();
  emlEmail.setFromName (this.getAuthorFullname ());
  emlEmail.setFromAddress (this.getAuthorEmail ());

  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth,
				  undefined_disc,
				  CONST_background_colour);
  strHTMLout += SWEFHTML.NBSP (2);
  strHTMLout += SWEFHTML.FONT_open (undefined_disc, undefined_disc, "#000000");
  strHTMLout += config.USERTEXT_SHOW_PostedByPrompt.weak ();
  strHTMLout += SWEFHTML.FONT_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth, undefined_disc, CONST_background_colour);
  strHTMLout += SWEFHTML.NBSP (2);
  strHTMLout += SWEFHTML.FONT_open (3, undefined_disc, "#000000");

  strHTMLout += (this.getAuthorFullname ()).weak ();
  strHTMLout += SWEFHTML.FONT_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();

  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth, undefined_disc, CONST_background_colour);
  strHTMLout += SWEFHTML.NBSP (2);
  strHTMLout += SWEFHTML.FONT_open (undefined_disc, undefined_disc, "#000000");
  strHTMLout += config.USERTEXT_SHOW_PostedOnPrompt.weak ();
  strHTMLout += SWEFHTML.FONT_close ();
  strHTMLout += SWEFHTML.BR ();
  strHTMLout += SWEFHTML.NBSP (1);
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth, undefined_disc, CONST_background_colour);
  strHTMLout += SWEFHTML.NBSP (2);
  strHTMLout += SWEFHTML.FONT_open (3, undefined_disc, "#000000");
  strHTMLout += this.getDateCreated ().getLongFormat ().weak ();
  strHTMLout += SWEFHTML.FONT_close ();
  strHTMLout += SWEFHTML.BR ();
  strHTMLout += SWEFHTML.NBSP (1);
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();

  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open ("100%", undefined_disc, "#F0F0F0", undefined_disc, undefined_disc, 2);
  strHTMLout += SWEFHTML.P_open () + SWEFHTML.NBSP () + SWEFHTML.P_close ();
  strHTMLout += this.getBody ().messageBody ();
  strHTMLout += SWEFHTML.P_open () + SWEFHTML.P_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();
  strHTMLout += SWEFHTML.TABLE_close ();

  return strHTMLout;
}

function getCustomFromEmailLink ()
{
  var HTMLout = "";
  HTMLout += SWEFHTML.A_open ("mailto:" + this.getFromAddress (), config.USERTEXT_SHOW_PopupEmailPrefix + this.getFromName () + " (" + this.getFromAddress () + ")");
  HTMLout += SWEFHTML.FONT_open (undefined_disc, undefined_disc, "#000000");
  HTMLout += this.getFromName ();
  HTMLout += SWEFHTML.FONT_close ();
  HTMLout += SWEFHTML.A_close ();

  return HTMLout;
}

function getCustomMessageForm (strActionPath, strButtonLabel, strURLTarget)
{
  return this.getSWEFMessageForm (strActionPath, strButtonLabel, strURLTarget, false);
}

function getSWEFMessageForm (strActionPath, strButtonLabel, strURLTarget, bShowUserPassFields)
{
  this.setFormName (config.FORM_MessageFormName);
  var strFormHTML = SWEFHTML.FORM_open (strActionPath, undefined_disc, this.getFormName (), strURLTarget);

  strFormHTML += SWEFHTML.INPUT_hidden (config.FORM_FieldMessageID, this.getMessageID ());
  strFormHTML += SWEFHTML.INPUT_hidden (config.FORM_FieldParentID, this.getParentID ());
//  strFormHTML += SWEFHTML.INPUT_hidden (config.FORM_FieldSiteID, this.getSiteID ());  
//  strFormHTML += SWEFHTML.INPUT_hidden (config.FORM_FieldForumID, this.getForumID ());
  strFormHTML += SWEFHTML.INPUT_hidden (config.FORM_FieldThreadID, this.getThreadID ());
  strFormHTML += SWEFHTML.INPUT_hidden (config.FORM_FieldSortCode, this.getSortCode ());
  strFormHTML += SWEFHTML.INPUT_hidden (config.FORM_FieldHiddenEmailOnResponse, this.getEmailParentOnResponse ());

  strFormHTML += SWEFHTML.TABLE_open (0,
				      config.ADMINSETTING_TableFullWidth,
				      undefined_disc,
				      "border-style: groove; border-color: Black; border-width: 1px 1px 1px 1px;",
				      0,
				      0);
  strFormHTML += SWEFHTML.TR_open ();
  strFormHTML += SWEFHTML.TD_open ("100%",
				   undefined_disc,
				   CONST_background_colour,
				   undefined_disc,
				   undefined_disc,
				   2);
  strFormHTML += SWEFHTML.TD_close ();
  strFormHTML += SWEFHTML.TR_close ();

  var strUserInfo;
  strUserInfo = (this.getAuthorFullname ()
		   + ", "
		   + this.getAuthorEmail ()).weak ();
  strFormHTML += SWEFHTML.TR_open ();
  strFormHTML += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth, undefined_disc, CONST_background_colour);
  strFormHTML += SWEFHTML.FONT_open (undefined_disc, undefined_disc, "#000000");
  strFormHTML += SWEFHTML.BR ();
  strFormHTML += SWEFHTML.NBSP (1);
  strFormHTML += config.USERTEXT_POST_SubjectPrompt.weak ();
  strFormHTML += SWEFHTML.BR ();
  strFormHTML += SWEFHTML.NBSP (1);
  strFormHTML += SWEFHTML.FONT_close ();
  strFormHTML += SWEFHTML.TD_close ();
  strFormHTML += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth, undefined_disc, CONST_background_colour);
  strFormHTML += SWEFHTML.FONT_open (undefined_disc, undefined_disc, "#000000");

  strFormHTML += SWEFHTML.BR ();
  strFormHTML += SWEFHTML.NBSP (1);
  strFormHTML += this.getSubjectInputField ();
  strFormHTML += SWEFHTML.BR ();
  strFormHTML += SWEFHTML.NBSP (1);

  strFormHTML += SWEFHTML.FONT_close ();
  strFormHTML += SWEFHTML.TD_close ();
  strFormHTML += SWEFHTML.TR_close ();

  if (bShowUserPassFields)
    {
      strFormHTML += SWEFHTML.TR_open ();
      strFormHTML += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth, undefined_disc, CONST_background_colour);
      strFormHTML += SWEFHTML.FONT_open (undefined_disc, undefined_disc, "#000000");
      strFormHTML += SWEFHTML.NBSP (1);
      strFormHTML += "&nbsp;Username:".weak ();
      strFormHTML += SWEFHTML.FONT_close ();
      strFormHTML += SWEFHTML.TD_close ();
      strFormHTML += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth, undefined_disc, CONST_background_colour);
      strFormHTML += SWEFHTML.INPUT_text ("EnteredUserName", "", 20);
      strFormHTML += SWEFHTML.TD_close ();
      strFormHTML += SWEFHTML.TR_close ();
      strFormHTML += SWEFHTML.TR_open ();
      strFormHTML += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth, undefined_disc, CONST_background_colour);
      strFormHTML += SWEFHTML.FONT_open (undefined_disc, undefined_disc, "#000000");
      strFormHTML += SWEFHTML.NBSP (1);
      strFormHTML += "Password:".weak ();
      strFormHTML += SWEFHTML.FONT_close ();
      strFormHTML += SWEFHTML.TD_close ();
      strFormHTML += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth, undefined_disc, CONST_background_colour);
      strFormHTML += SWEFHTML.INPUT_password ("EnteredPassword", "", 20);
      strFormHTML += SWEFHTML.TD_close ();
      strFormHTML += SWEFHTML.TR_close ();
    }
  else
    {
      strFormHTML += SWEFHTML.TR_open ();
      strFormHTML += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth, undefined_disc, CONST_background_colour);
      strFormHTML += SWEFHTML.FONT_open (undefined_disc, undefined_disc, "#000000");
      strFormHTML += SWEFHTML.NBSP (1);
      strFormHTML += config.USERTEXT_POST_PostedByPrompt.weak ();      
      strFormHTML += SWEFHTML.FONT_close ();
      strFormHTML += SWEFHTML.TD_close ();
      strFormHTML += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth, undefined_disc, CONST_background_colour);
      strFormHTML += SWEFHTML.FONT_open (undefined_disc, undefined_disc, "#000000");
      strFormHTML += strUserInfo.weak ();      
      strFormHTML += SWEFHTML.FONT_close ();
      strFormHTML += SWEFHTML.TD_close ();
      strFormHTML += SWEFHTML.TR_close ();
    }

  strFormHTML += SWEFHTML.TR_open ();
  strFormHTML += SWEFHTML.TD_open ("100%", undefined_disc, CONST_background_colour, undefined_disc, undefined_disc, 2);
  strFormHTML += SWEFHTML.NBSP ();
  strFormHTML += SWEFHTML.TD_close ();
  strFormHTML += SWEFHTML.TR_close ();

  strFormHTML += SWEFHTML.TR_open ();
  strFormHTML += SWEFHTML.TD_open ("100%", undefined_disc, CONST_background_colour, undefined_disc, undefined_disc, 2);

  if (bShowUserPassFields)
    {
      strFormHTML += SWEFHTML.TEXTAREA_open (config.FORM_FieldMessage,
					     config.ADMINSETTING_TextAreaCols,
					     config.ADMINSETTING_TextAreaRows);
      strFormHTML += this.getBody ().unformatFromStoring ();
      strFormHTML += SWEFHTML.TEXTAREA_close ();
    }
  else
    {
      strFormHTML += this.getBodyInputField ();
    }
  strFormHTML += SWEFHTML.TD_close ();
  strFormHTML += SWEFHTML.TR_close ();

  strFormHTML += this.getEmailResponsesSubform ();

  // Post Reply

  strFormHTML += SWEFHTML.TABLE_close ();
  strFormHTML += SWEFHTML.BR ();
  strFormHTML += SWEFHTML.INPUT_submit (strButtonLabel, strButtonLabel);
  strFormHTML += SWEFHTML.FORM_close ();

  return strFormHTML;
}

// Thread Marker and Display

function myGetViewEntry
(
 nActiveMessageID
)
{
  var strThreadMarker = SWEFHTML.IMG ("threadmarker.gif", "*", 0, "7", "7");

  var strHTMLout = SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassViewSubject);

  if (nActiveMessageID == this.getMessageID ())
    {
      strHTMLout += strThreadMarker + SWEFHTML.NBSP () + this.getSubject ().strong();
    }
  else
    {
      strHTMLout += SWEFHTML.NBSP (2) + this.getSubjectLink ().strong();
    }

  strHTMLout += SWEFHTML.SPAN_close ();

  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassViewAuthor);
  strHTMLout += config.USERTEXT_VIEW_SeparateSubjectAuthor;
  strHTMLout += SWEFHTML.QUOTE_open ();
  strHTMLout += this.getAuthorFullname ().strongSmall();
  strHTMLout += SWEFHTML.QUOTE_close ();  
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassViewDate);
  strHTMLout += config.USERTEXT_VIEW_SeparateAuthorDate;
  strHTMLout += this.getDateCreated ().getShortFormat ();
  strHTMLout += SWEFHTML.SPAN_close ();

  if (this.getNumChildren () == 0)
    {
      strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassViewNoChildren);
      strHTMLout += " (";
      strHTMLout += config.USERTEXT_VIEW_NoRepliesTag;
      strHTMLout += ")";
      strHTMLout += SWEFHTML.SPAN_close ();
    }
  else if (this.getNumChildren () == 1)
    {
      strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassViewOneChild);
      strHTMLout += " (1";
      strHTMLout += config.USERTEXT_VIEW_OneReplyTag;
      strHTMLout += ")";
      strHTMLout += SWEFHTML.SPAN_close ();
    }
  else
    {
      strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassViewManyChildren);
      strHTMLout += " (";
      strHTMLout += this.getNumChildren ();
      strHTMLout += config.USERTEXT_VIEW_ManyRepliesTag;
      strHTMLout += ")";
      strHTMLout += SWEFHTML.SPAN_close ();
    }

  strHTMLout = strHTMLout.weakSmall();

  return strHTMLout;
}

function getCustomMessageMailForm (actionPath, buttonLabel)
{
  return this.getSWEFMessageForm (actionPath, buttonLabel, "_new", true);
}

function customUserPrefersHTMLMail (usernameToCheck)
{
  var loginRecord = executePersonQuery ("SELECT * FROM [UserData] WHERE [NTLogin] = '" + usernameToCheck + "'");
  if (loginRecord.Fields ("Email_Method") == 1)
    {
      return true;
    }
  else
    {
      return false;
    }
}

function executePersonQuery (strSQL)
{
  var cnDBConnection;
  cnDBConnection = Server.CreateObject ("ADODB.Connection");
  cnDBConnection.Mode = adModeReadWrite;
  cnDBConnection.Open (config.getDatabaseDSN ());

  var rsPersonRecord;
  rsPersonRecord = Server.CreateObject ("ADODB.RecordSet");
  rsPersonRecord.Open (strSQL, cnDBConnection, adOpenKeyset, adLockOptimistic, adCmdText);

  return rsPersonRecord;
}

function customUserPrefersResponsesEmailed (usernameToCheck)
{
  if (Session ("EmailResponsesDefault"))
    {
      return true;
    }
  else
    {
      return false;
    }
}

SWEFMessage.prototype.getViewEntry = myGetViewEntry;
SWEFEmail.prototype.getFromEmailLink = getCustomFromEmailLink;
SWEFEmail.prototype.getHTMLMessageBody = HTMLMailMessage;
SWEFMessage.prototype.getRenderedHTML = getCustomRenderedHTML;
SWEFForm.prototype.getMessageForm = getCustomMessageForm;
SWEFForm.prototype.getSWEFMessageForm = getSWEFMessageForm;
SWEFForm.prototype.getMessageMailForm = getCustomMessageMailForm;
SWEFPageElement.showForumLink = customShowForumLink;
SWEFConfig.prototype.getUserHTMLMailPreference = customUserPrefersHTMLMail
SWEFConfig.prototype.getUserEmailResponsePreference = customUserPrefersResponsesEmailed
String.prototype.messageBody = formatMessageBody;
String.prototype.strong = formatStrong;
String.prototype.strongRed = formatStrongRed;
String.prototype.strongSmall = formatStrongSmall;
String.prototype.weak = formatWeak;
String.prototype.weakSmall = formatWeakSmall;
</SCRIPT>
<%

currentSite_ID_disc          = Session("Site_Id")
currentSite_Code_desc        = Session("Site_Code")
currentSite_Description_desc = Session("Site_Description")

currentForum_ID_disc         = Session("Forum_Id")
currentForum_Title_disc      = Session("Forum_Title")
currentForum_Moderated_disc  = Session("Forum_Moderated")

currentUsername_disc         = Session("UserName")
currentUserFullName_disc     = Session("FullName")
currentUserEmailAddress_disc = Session("EmailAddress")
currentLanguage_disc         = Session("Language")
currentUser_ID_disc          = Session("User_id")

isAdministrator_disc         = Session("IsAdministrator")

%>