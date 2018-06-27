<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// CONFIG OBJECT
//
// ======================================================================

function SWEFConfig
(
)
{
  this.ADMINSETTING_ForumName = currentForum_Title_disc;

  this.ADMINSETTING_DatabaseFilename = "";
  this.ADMINSETTING_DatabaseTable = "[Forum]";
  this.ADMINSETTING_DatabaseDSN = "";

  this.ADMINSETTING_DaysMessagesActive = 90;
  this.ADMINSETTING_CacheTimeoutMinutes = 120;
  this.ADMINSETTING_VirtualPath = undefined_disc;

  this.ADMINSETTING_ArchiveBeginYear = 2000;

  this.ADMINSETTING_EmailAdminName = "";
  this.ADMINSETTING_EmailAdminAddress = "";
  this.ADMINSETTING_EmailAlertFromName = "Forums Auto-Alert";
  this.ADMINSETTING_EmailAlertFromAddress = "Kelly.Whitlock@Fluke.com";
  this.ADMINSETTING_EmailAlertSignature = "\n\nCopyright 1995-2004 Fluke Corporation - All rights reserved\n";
  this.ADMINSETTING_EmailAlertSignatureHTML = "<BR><BR><BR><DIV ALIGN=CENTER><FONT STYLE=\"font-weight:Normal; font-size:8.5pt;font-family: Arial,Verdana; color:#000000; font-style: Normal;\">&copy; 1995-2001 Fluke Corporation - All rights reserved</FONT></DIV>";
  
  // The following govern the images used for expanding and collapsing threads.
  this.ADMINSETTING_ExpandImagePathname = "plus.gif";
  this.ADMINSETTING_ExpandImageWidth = 11;
  this.ADMINSETTING_ExpandImageHeight = 11;
  this.ADMINSETTING_CollapseImagePathname = "minus.gif";
  this.ADMINSETTING_CollapseImageWidth = 11;
  this.ADMINSETTING_CollapseImageHeight = 11;
  this.ADMINSETTING_NoExpandImagePathname = "blank.gif";
  this.ADMINSETTING_NoExpandImageWidth = 11;
  this.ADMINSETTING_NoExpandImageHeight = 11;
  this.ADMINSETTING_ExpandCollapseFontSize = 2;

  // The following govern the images used for the formatting toolbar on the IE5 text control.
  this.ADMINSETTING_ToolbarBackgroundColour = "#CCCCCC";
  this.ADMINSETTING_ToolbarImagePathname = "toolbar.gif";
  this.ADMINSETTING_ToolbarButtonImagePathname = "button_normal.gif";
  this.ADMINSETTING_ToolbarButtonImageUpPathname = "button_up.gif";
  this.ADMINSETTING_ToolbarButtonImageDownPathname = "button_down.gif";
  this.ADMINSETTING_ToolbarButtonImageWidth = 25;
  this.ADMINSETTING_ToolbarButtonImageHeight = 20;
  this.ADMINSETTING_ToolbarSpacerImagePathname = "toolbar_spacer.gif";
  this.ADMINSETTING_ToolbarSpacerImageWidth = 5;
  this.ADMINSETTING_ToolbarSpacerImageHeight = 20;
  this.ADMINSETTING_ToolbarTotalWidth = 300;

  // The following govern the appearance of some page elements.
  this.ADMINSETTING_CSSClass = "Medium";
  this.ADMINSETTING_CSSClassViewAuthor = "SW_Forum_ViewAuthor";
  this.ADMINSETTING_CSSClassViewNoChildren = "SW_Forum_ViewNoChildren";
  this.ADMINSETTING_CSSClassViewOneChild = "SW_Forum_ViewOneChild";
  this.ADMINSETTING_CSSClassViewManyChildren = "SW_Forum_ViewManyChildren";
  this.ADMINSETTING_CSSClassViewDate = "SW_Forum_ViewDate";
  this.ADMINSETTING_CSSClassViewSubject = "SW_Forum_ViewSubject";
  this.ADMINSETTING_CSSClassMsgPostedByLabel = "SW_Forum_MsgPostedByLabel";
  this.ADMINSETTING_CSSClassMsgPostedBy = "SW_Forum_MsgPostedBy";
  this.ADMINSETTING_CSSClassMsgPostedOnLabel = "SW_Forum_MsgPostedOnLabel";
  this.ADMINSETTING_CSSClassMsgPostedOn = "SW_Forum_MsgPostedOn";
  this.ADMINSETTING_CSSClassMsgBodyLabel = "SW_Forum_MsgBodyLabel";
  this.ADMINSETTING_CSSClassMsgBody = "SW_Forum_MsgBody";
  this.ADMINSETTING_CSSClassFormBody = "SW_Forum_FormBody";
  this.ADMINSETTING_CSSClassFormBodyLabel = "SW_Forum_FormBodyLabel";
  this.ADMINSETTING_CSSClassFormSubject = "SW_Forum_FormSubject";
  this.ADMINSETTING_CSSClassFormSubjectLabel = "SW_Forum_FormSubjectLabel";
  this.ADMINSETTING_CSSClassFormEmailResponses = "SW_Forum_FormEmailResponses";
  this.ADMINSETTING_CSSClassFormEmailResponsesLabel = "SW_Forum_FormEmailResponsesLabel";
  this.ADMINSETTING_CSSClassFormPostedBy = "SW_Forum_FormPostedBy";
  this.ADMINSETTING_CSSClassFormPostedByLabel = "SW_Forum_FormPostedByLabel";
  this.ADMINSETTING_CSSClassFormFullName = "SW_Forum_FormFullName";
  this.ADMINSETTING_CSSClassFormFullNameLabel = "SW_Forum_FormFullNameLabel";
  this.ADMINSETTING_CSSClassFormEmailAddress = "SW_Forum_FormEmailAddress";
  this.ADMINSETTING_CSSClassFormEmailAddressLabel = "SW_Forum_FormEmailAddressLabel";
  this.ADMINSETTING_CSSClassFormSearch = "SW_Forum_FormSearch";
  this.ADMINSETTING_CSSClassFormSearchSmall = "SW_Forum_FormSearchSmall";
  this.ADMINSETTING_CSSClassFormDateElement = "SW_Forum_FormDateElement";
  this.ADMINSETTING_CSSClassFormMonthElement = "SW_Forum_FormMonthElement";
  this.ADMINSETTING_CSSClassFormYearElement = "SW_Forum_FormYearElement";

  this.ADMINSETTING_DefaultEmbeddedLinkTarget = undefined_disc;
  this.ADMINSETTING_TopPageTarget = undefined_disc;
  this.ADMINSETTING_AdminPageTarget = undefined_disc;
  this.ADMINSETTING_ArchivePageTarget = undefined_disc;
  this.ADMINSETTING_EditPostActionPageTarget = undefined_disc;
  this.ADMINSETTING_EditPostPageTarget = undefined_disc;
  this.ADMINSETTING_MainPageTarget = undefined_disc;
  this.ADMINSETTING_NewPostActionPageTarget = undefined_disc;
  this.ADMINSETTING_NewPostPageTarget = undefined_disc;
  this.ADMINSETTING_SearchPageTarget = undefined_disc;
  this.ADMINSETTING_ShowMessagePageTarget = undefined_disc;

  this.ADMINSETTING_TableBorderSize = 0;
  this.ADMINSETTING_TableFullWidth = 468;
  this.ADMINSETTING_TableTitleColumnWidth = 100;
  this.ADMINSETTING_TableFieldColumnWidth = 368;
  this.ADMINSETTING_TextAreaRows = 10;
  this.ADMINSETTING_TextAreaCols = 52;
  this.ADMINSETTING_PrecisLength = 150;
  this.ADMINSETTING_ViewIndentResponseSpaces = 2;
  this.ADMINSETTING_InputboxMaxLength = 100;
  this.ADMINSETTING_SubjectInputboxSize = 50;
  this.ADMINSETTING_ExpandFirstNThreads = 0;
  this.ADMINSETTING_ShowAtMostNThreads = 0;
  this.ADMINSETTING_NoWrapMessageThreadViews = false;

  // The following switches turn on or off various ASP Forums features.
  this.ADMINSWITCH_AllowEmailResponses = true;
  this.ADMINSWITCH_AllowUserEditing = true;
  this.ADMINSWITCH_AllowRichFormatting = true;
  this.ADMINSWITCH_AllowRichFormattingImages = false;
  this.ADMINSWITCH_AllowHTMLEmail = true;
  this.ADMINSWITCH_ShowExpandCollapse = true;
  this.ADMINSWITCH_ShowEmailAddresses = true;
  this.ADMINSWITCH_ShowNewPostButtonOnMessage = true;
  this.ADMINSWITCH_ViewPostsAscending = false;
  this.ADMINSWITCH_ExpandAllThreads = false;

  // The following text strings are used on the default.asp page.
  this.USERTEXT_VIEW_NoMessages = "There are no messages to display.";
  this.USERTEXT_VIEW_NoRepliesTag = "<FONT COLOR=Gray>No Replies</FONT>";
  this.USERTEXT_VIEW_OneReplyTag = "&nbsp;Reply";
  this.USERTEXT_VIEW_ManyRepliesTag = "&nbsp;Replies";
  this.USERTEXT_VIEW_SeparateSubjectAuthor = " by ";
  this.USERTEXT_VIEW_SeparateAuthorDate = " on ";
  this.USERTEXT_VIEW_PopupExpandLink = "+ Expand this subject thread";
  this.USERTEXT_VIEW_PopupCollapseLink = "- Collapse this subject thread";

  // The following text strings are used on the showmessage.asp page.
  this.USERTEXT_SHOW_NewPostButton = "Post New Message";
  this.USERTEXT_SHOW_EditPostButton = "Edit Message";
  this.USERTEXT_SHOW_ReplyPostButton = "Post a Reply";
  this.USERTEXT_SHOW_PostedByPrompt = "Posted By:";
  this.USERTEXT_SHOW_PostedOnPrompt = "Posted On:";
  this.USERTEXT_SHOW_BodyPrompt = "Message:";
  this.USERTEXT_SHOW_PopupSubjectPrefix = "View Message: ";
  this.USERTEXT_SHOW_PopupEmailPrefix = "Email: ";
  this.USERTEXT_SHOW_PopupArchiveLinkForumPrefix = "View ";
  this.USERTEXT_SHOW_PopupArchiveLinkForumSuffix = " Archive.";
  this.USERTEXT_SHOW_ArchiveLinkForumSuffix = " Archive";
  this.USERTEXT_SHOW_AuthorEmailSeparator = " (";
  this.USERTEXT_SHOW_EmailSuffix = ")";

  // The following text strings are used on the newpost.asp, newpostaction.asp, editpost and editpostaction.asp pages.
  this.USERTEXT_POST_UpdateFailedPrefix = "Update message failed: ";
  this.USERTEXT_POST_UpdateSuccessful = "<LI>Your message has been updated.</LI>";
  this.USERTEXT_POST_PostFailedPrefix = "Post message failed: ";
  this.USERTEXT_POST_PostSuccessful = "<LI>Your message has been posted.</LI><LI>Please do not refresh this page, since that will repost your message!</LI>";
  this.USERTEXT_POST_ForumLinkPrefix = "Back to ";
  this.USERTEXT_POST_ReplySubjectPrefix = "Re: ";
  this.USERTEXT_POST_PreviousMessageLinkText = "Previous Message";
  this.USERTEXT_POST_UsernamePrompt = "&nbsp;Username:";
  this.USERTEXT_POST_FullnamePrompt = "&nbsp;Full Name:";
  this.USERTEXT_POST_EmailAddressPrompt = "&nbsp;Email Address:";
  this.USERTEXT_POST_PostedByPrompt = "&nbsp;Posted By:";
  this.USERTEXT_POST_SubjectPrompt = "&nbsp;Subject:";
  this.USERTEXT_POST_BodyPrompt = "Message:";
  this.USERTEXT_POST_EmailResponsesPrompt = "Send me an email when someone responds to this message.";
  this.USERTEXT_POST_SubmitButton = "Post the Message";
  this.USERTEXT_POST_SaveChangesButton = "Save Changes to Message";
  this.USERTEXT_POST_NotAuthorisedToEditMessage = "You are not authorized to edit this message.";
  this.USERTEXT_POST_ErrorNoSubject = "A Subject is required for each message.";
  this.USERTEXT_POST_ErrorNoBody = "A Message is required for each message.";
  this.USERTEXT_POST_ErrorNoUsername = "No Logon User Name was supplied for this message.";
  this.USERTEXT_POST_ErrorNoName = "No Name was supplied for message.";
  this.USERTEXT_POST_ErrorNoEmail = "No Email Address was supplied for message.";
  this.USERTEXT_POST_ErrorInvalidThreadID = "SW-Forum internal error (invalid Thread ID) has occurred.  Please capture this error message and email it to: David.Whitlock@Fluke.com";
  this.USERTEXT_POST_ErrorInvalidParentID = "SW-Forum internal error (invalid Parent ID) has occurred.  Please capture this error message and email it to: David.Whitlock@Fluke.com";
  this.USERTEXT_POST_ErrorInvalidSiteID = "SW-Forum internal error (invalid Site ID) has occurred.  Please capture this error message and email it to: David.Whitlock@Fluke.com";
  this.USERTEXT_POST_ErrorInvalidForumID = "SW-Forum internal error (invalid Forum ID) has occurred.  Please capture this error message and email it to: David.Whitlock@Fluke.com";      

  // The following strings are all used on the archive.asp page.
  this.USERTEXT_ARCHIVE_ShowResults = "View Messages for this Month / Year";

  // The following strings are used on the search.asp page.
  this.USERTEXT_SEARCH_SmallSearchPrefix = "Keyword:";
  this.USERTEXT_SEARCH_SubmitButton = "Search";
  this.USERTEXT_SEARCH_ResultsHeader1 = "Your keyword search for ";
  this.USERTEXT_SEARCH_ResultsHeader2 = " found ";
  this.USERTEXT_SEARCH_ResultsHeaderSuffix0Match = " matches.<BR><HR COLOR=\"black\" NOSHADE WIDTH=\"100%\">";
  this.USERTEXT_SEARCH_ResultsHeaderSuffix1Match = " match.<BR><HR COLOR=\"black\" NOSHADE WIDTH=\"100%\">";
  this.USERTEXT_SEARCH_ResultsHeaderSuffixManyMatches = " matches.<BR><HR COLOR=\"black\" NOSHADE WIDTH=\"100%\">";

  // The following strings are used on the admin.asp page.
  this.USERTEXT_SQL_EnterPrompt = "Enter your SQL queries or updates below:"
  this.USERTEXT_SQL_ExecutePrompt = "Execute SQL";
  this.USERTEXT_SQL_ResultsPrefix = "Results:";
  this.USERTEXT_SQL_StatementPrefix = "SQL statement: ";
  this.USERTEXT_SQL_DBErrorPrefix = "Database errors have occured:";
  this.USERTEXT_SQL_DBErrorNumberPrompt = "Error Number = ";
  this.USERTEXT_SQL_DBErrorDescriptionPrompt = "Error description = ";
  this.USERTEXT_SQL_DBErrorNoErrorPrompt = "No database errors have occured.";
  this.USERTEXT_SQL_DeleteSuccessfulMessage = "Deletion Successful";
  this.USERTEXT_SQL_DeleteUnsuccessfulMessage = "Deletion Failed";
  this.USERTEXT_SQL_DeleteHierarchyButton = "Delete Message";
  this.USERTEXT_SQL_PurgeCacheButton = "Purge Cache";

  // The following strings are all used in email messages.
  this.USERTEXT_MAIL_AdminNewPostSubjectPrefix = "New Message";
  this.USERTEXT_MAIL_AdminNewPostBodyPrefix = "A new message by ";
  this.USERTEXT_MAIL_AdminNewPostSeparateNameForum = " was posted in ";
  this.USERTEXT_MAIL_UserNewResponseSubjectPrefix = "New Response";
  this.USERTEXT_MAIL_UserNewPostSeparateNameMessage = " has responded to your message: ";
  this.USERTEXT_MAIL_UserNewPostNameSuffix = " has responded to your message. ";
  this.USERTEXT_MAIL_UserNewPostMessagePrefix = "The response message reads: ";

  // The following are used by our general string-handling routines.
  this.USERTEXT_STRING_StringTruncatedSuffix = "...";
  this.USERTEXT_STRING_WarningUnverifiedEmail = "WARNING: An Email Address has been included by author of this message. ";
  this.USERTEXT_STRING_WarningUnverifiedImage = "WARNING: An Embedded Image has been included by the author of this message. ";
  this.USERTEXT_STRING_WarningUnverifiedLink = "WARNING: A Hyperlink URL has been included by the author of this message. ";

  // You can rename the files involved, should you really want to.
  this.PAGE_AdminLocalPath = "admin.asp";
  this.PAGE_ArchiveLocalPath = "archive.asp";
  this.PAGE_EditPostActionLocalPath = "editpostaction.asp";
  this.PAGE_EditPostLocalPath = "editpost.asp";
  this.PAGE_MainLocalPath = "default.asp";
  this.PAGE_MainPreferredLocalPath = "./";
  this.PAGE_NewPostActionLocalPath = "newpostaction.asp";
  this.PAGE_NewPostLocalPath = "newpost.asp";
  this.PAGE_SearchLocalPath = "search.asp";
  this.PAGE_ShowMessageLocalPath = "showmessage.asp";

  this.CACHE_TimeStampKey = "Timestamp";
  this.CACHE_Enabled = false;
  this.CACHE_PurgeOnPageCleanup = false;
  this.CACHE_ItemSeparator = "|";
  this.CACHE_AllRootMessagesKey = "AllRoot";
  this.CACHE_CurrentRootMessagesKeyPrefix = "CurrentRoot";
  this.CACHE_AllCurrentMessagesKeyPrefix = "AllCurrent";
  this.CACHE_RootArchivedMessagesKeyPrefix = "ArchivedRoot";
  this.CACHE_AllArchivedMessagesKeyPrefix = "ArchivedAll";
  this.CACHE_AllThreadMessagesKey = "AllThread";
  this.CACHE_SubThreadMessagesKey = "SubThread";

  this.CONST_NoError = 0;
  this.CONST_Error = 1;

  this.FORM_MessageFormName = "SWEFMessage";
  this.FORM_SearchFormName = "SWEFSearch";
  this.FORM_ArchiveFormName = "SWEFArchive";
  this.FORM_AdminDeleteFormName = "SWEFAdminDelete";
  this.FORM_AdminPurgeFormName = "SWEFAdminPurge";
  this.FORM_AdminSQLFormName = "SWEFAdminSQL";
  this.FORM_CheckboxChecked = " CHECKED";
  this.FORM_CheckboxUnchecked = "";
  this.FORM_CheckboxTrue = "Yes";

  this.FORM_QueryStringMessageID = "messageID";
  this.FORM_QueryStringViewExpand = "expand";
  this.FORM_QueryStringViewCollapse = "collapse";
  this.FORM_QueryStringViewCentre = "centreOnMessage-";

  this.FORM_FieldMessageID = "messageID";
  this.FORM_FieldSubject = "subject";
  this.FORM_FieldMessage = "message";
  this.FORM_FieldSortCode = "sortCode";
  this.FORM_FieldSiteID = "Site_ID";
  this.FORM_FieldForumID = "Forum_ID";    
  this.FORM_FieldParentID = "parentID";
  this.FORM_FieldThreadID = "threadID";
  this.FORM_FieldEmailResponses = "emailResponses";
  this.FORM_FieldHiddenEmailOnResponse = "emailParentOnResponse";
  this.FORM_FieldSearchString = "SearchString";
  this.FORM_FieldUsername = "username";
  this.FORM_FieldEmailaddress = "emailaddress";
  this.FORM_FieldFullname = "fullname";
  this.FORM_FieldSQLStatement = "SQLStatement";
  this.FORM_FieldMessageIDToDelete = "messageIDToDelete";
  this.FORM_FieldPurgeCache = "purgeCache";
  this.FORM_FieldArchiveDate = "archiveDate";
  this.FORM_FieldDateDaySuffix = "Day";
  this.FORM_FieldDateMonthSuffix = "Month";
  this.FORM_FieldDateYearSuffix = "Year";

  this.DATABASE_MaxSortcodeSize = 25;
  this.DATABASE_MaxSubjectSize = 100;
  this.DATABASE_MaxMessageSize = 2500;
  this.DATABASE_MaxUsernameSize = 50;
  this.DATABASE_MaxFullnameSize = 50;
  this.DATABASE_MaxEmailAddressSize = 50;
  this.DATABASE_FieldSiteID = "Site_ID";  
  this.DATABASE_FieldForumID = "Forum_ID";
  this.DATABASE_FieldMessageID = "messageID";
  this.DATABASE_FieldParentID = "parent";
  this.DATABASE_FieldThreadID = "threadID";
  this.DATABASE_FieldSortCode = "sortCode";
  this.DATABASE_FieldNumChildren = "numChildren";
  this.DATABASE_FieldAuthorName = "author";
  this.DATABASE_FieldAuthorFullName = "authorFullName";
  this.DATABASE_FieldAuthorEmail = "authorEmail";
  this.DATABASE_FieldSubject = "subject";
  this.DATABASE_FieldBody = "body";
  this.DATABASE_FieldEmailParentOnResponse = "emailParentOnResponse";
  this.DATABASE_FieldDateCreated = "dateCreated";
  this.DATABASE_FieldDateModified = "dateModified";

  this.SYS_SafeDefaultDate = "01/01/2001";
  this.SYS_AllJavascriptEvents = "onAbort|onAfterUpdate|onBeforeUnload|onBeforeUpdate|onBlur|onBounce|onClick|onChange|onDataAvailable|onDataSetChanged|onDataSetComplete|onDblClick|onDragDrop|onError|onErrorUpdate|onFilterChange|onFocus|onHelp|onKeyDown|onKeyPress|onKeyUp|onLoad|onMouseDown|onMouseMove|onMouseOut|onMouseOver|onMouseUp|onMove|onReadyStateChange|onReset|onResize|onRowEnter|onRowExit|onScroll|onSelect|onSelectStart|onStart|onSubmit|onUnload";
  this.SYS_CurrentVersion = "2.1";
  this.SYS_CurrentVersionReference = "SiteWide-Forums " + this.SYS_CurrentVersion;

  var strCurrentPage = String (Request.ServerVariables ("PATH_INFO"));
  var nLastSeparator = strCurrentPage.lastIndexOf ("/");
  this.ADMINSETTING_VirtualPath = strCurrentPage.substr (0, nLastSeparator + 1);

  return this;
}

function getDHTMLFunctionName_cnf_disc
(
)
{
  return "expandCollapseForum" + this.getUniqueKey () + "Click_disc";
}

function getDHTMLEventHandler_cnf_disc
(
)
{
  return this.getDHTMLFunctionName () + " (event); return false;";
}

function getPagePath_cnf_disc
(
 strPageName
)
{
  var strPagePath;
  if (isUndefined_disc (this.ADMINSETTING_VirtualPath))
    {
      strPagePath = strPageName;
    }
  else
    {
      strPagePath = this.ADMINSETTING_VirtualPath + strPageName;
    }

  return strPagePath;
}

function getAdminPagePath_cnf_disc
(
)
{
  return this.getPagePath (this.PAGE_AdminLocalPath);
}

function getArchivePagePath_cnf_disc
(
)
{
  return this.getPagePath (this.PAGE_ArchiveLocalPath);
}

function getMainPagePath_cnf_disc
(
)
{
  return this.getPagePath (this.PAGE_MainLocalPath);
}

function getMainPagePreferredPath_cnf_disc
(
)
{
  return this.getPagePath (this.PAGE_MainPreferredLocalPath);
}

function getNewPostActionPagePath_cnf_disc
(
)
{
  return this.getPagePath (this.PAGE_NewPostActionLocalPath);
}

function getNewPostPagePath_cnf_disc
(
)
{
  return this.getPagePath (this.PAGE_NewPostLocalPath);
}

function getEditPostActionPagePath_cnf_disc
(
)
{
  return this.getPagePath (this.PAGE_EditPostActionLocalPath);
}

function getEditPostPagePath_cnf_disc
(
)
{
  return this.getPagePath (this.PAGE_EditPostLocalPath);
}

function getSearchPagePath_cnf_disc
(
)
{
  return this.getPagePath (this.PAGE_SearchLocalPath);
}

function getShowMessagePagePath_cnf_disc
(
)
{
  return this.getPagePath (this.PAGE_ShowMessageLocalPath);
}

function getDatabaseDSN_cnf_disc
(
)
{
  if (this.ADMINSETTING_DatabaseDSN == "")
    {
      var strDBPath = Server.MapPath (this.ADMINSETTING_DatabaseFilename);

      this.ADMINSETTING_DatabaseDSN = "DRIVER=Microsoft Access Driver (*.mdb);UID=admin;UserCommitSync=Yes;Threads=3;SafeTransactions=0;PageTimeout=5;MaxScanRows=8;MaxBufferSize=512;ImplicitCommitSync=Yes;FIL=MS Access;DriverId=25;DefaultDir=;DBQ=" + strDBPath
    }

  return this.ADMINSETTING_DatabaseDSN;
}

function getSiteBaseURL_cnf_disc
(
)
{
  return "http://" + String (Request.ServerVariables ("HTTP_HOST"));
}

function getEditableUserInfoSwitch_cnf_disc
(
)
{
  var bInfoEditable;
  if ((currentUsername_disc == "")
      || (currentUserFullName_disc == "")
      || (currentUserEmailAddress_disc == ""))
    {
      bInfoEditable = true;
    }
  else
    {
      bInfoEditable = false;
    }

  return bInfoEditable;
}

function getUserHTMLMailPreference_cnf_disc
(
 strUsernameToCheck
)
{
  return true;
}

function getUserEmailResponsePreference_cnf_disc
(
 strUsernameToCheck
)
{
  return false;
}

function getUniqueKey_cnf_disc
(
)
{
  var strKey;
  strKey = this.ADMINSETTING_ForumName + this.ADMINSETTING_DatabaseTable;
  strKey = strKey.replace (/ /gi, "");
  strKey = strKey.replace (/\[/gi, "");
  strKey = strKey.replace (/\]/gi, "");
  strKey = strKey.replace (/\(/gi, "");
  strKey = strKey.replace (/\)/gi, "");
  strKey = strKey.replace (/\-/gi, "");

  return strKey;
}

function forumContextSwitch_cnf_disc
(
 strNewForumName,
 strNewForumTableName,
 strNewForumDSN,
 strNewForumFilename,
 strNewForumVirtualPath
)
{
  if (isDefined_disc (strNewForumName))
    {
      this.ADMINSETTING_ForumName = strNewForumName;
    }

  if (isDefined_disc (strNewForumTableName))
    {
      this.ADMINSETTING_DatabaseTable = strNewForumTableName;
    }

  if (isDefined_disc (strNewForumDSN))
    {
      this.ADMINSETTING_DatabaseDSN = strNewForumDSN;
    }

  if (isDefined_disc (strNewForumFilename))
    {
      this.ADMINSETTING_DatabaseFilename = strNewForumFilename;
    }

  if (isDefined_disc (strNewForumVirtualPath))
    {
      this.ADMINSETTING_VirtualPath = strNewForumVirtualPath;
    }

  return;
}

SWEFConfig.prototype.getDHTMLFunctionName = getDHTMLFunctionName_cnf_disc;
SWEFConfig.prototype.getDHTMLEventHandler = getDHTMLEventHandler_cnf_disc;
SWEFConfig.prototype.getPagePath = getPagePath_cnf_disc;
SWEFConfig.prototype.getAdminPagePath = getAdminPagePath_cnf_disc;
SWEFConfig.prototype.getArchivePagePath = getArchivePagePath_cnf_disc;
SWEFConfig.prototype.getEditPostActionPagePath = getEditPostActionPagePath_cnf_disc;
SWEFConfig.prototype.getEditPostPagePath = getEditPostPagePath_cnf_disc;
SWEFConfig.prototype.getMainPagePath = getMainPagePath_cnf_disc;
SWEFConfig.prototype.getMainPagePreferredPath = getMainPagePreferredPath_cnf_disc;
SWEFConfig.prototype.getNewPostActionPagePath = getNewPostActionPagePath_cnf_disc;
SWEFConfig.prototype.getNewPostPagePath = getNewPostPagePath_cnf_disc;
SWEFConfig.prototype.getSearchPagePath = getSearchPagePath_cnf_disc;
SWEFConfig.prototype.getShowMessagePagePath = getShowMessagePagePath_cnf_disc;

SWEFConfig.prototype.getEditableUserInfoSwitch = getEditableUserInfoSwitch_cnf_disc;
SWEFConfig.prototype.getDatabaseDSN = getDatabaseDSN_cnf_disc;
SWEFConfig.prototype.getSiteBaseURL = getSiteBaseURL_cnf_disc;
SWEFConfig.prototype.getUserHTMLMailPreference = getUserHTMLMailPreference_cnf_disc;
SWEFConfig.prototype.getUserEmailResponsePreference = getUserEmailResponsePreference_cnf_disc;
SWEFConfig.prototype.getUniqueKey = getUniqueKey_cnf_disc;
SWEFConfig.prototype.forumContextSwitch = forumContextSwitch_cnf_disc;
</SCRIPT>

