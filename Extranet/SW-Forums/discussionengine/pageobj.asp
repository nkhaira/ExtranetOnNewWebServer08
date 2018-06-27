<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// PAGE OBJECT
//
// ======================================================================

function SWEFPageElement
(
)
{
  return this;
}

// ======================================================================
//
// Interface to private member variables.
//
// ======================================================================

// ======================================================================
//
// Main object methods.
//
// ======================================================================

function showVersion_page_disc
(
)
{
  if (!SWEFPageElement.versionRefOutput)
    {
      var strVersionRef = "";
      strVersionRef += "\n";
      strVersionRef += "\<!-- SiteWide-Forum -- Adapted SWEForum version: ";
      strVersionRef += config.SYS_CurrentVersionReference;
      strVersionRef += " -->\n";
      strVersionRef.show ();
      SWEFPageElement.versionRefOutput = true;
    }

  return;
}

function setCurrentMessage_page_disc
(
)
{
  var bMessageFound = false;
  if (isDefined_disc (currentMessage_disc))
    {
      bMessageFound = true;
    }
  else
    {
      if (!isNaN (Request.QueryString (config.FORM_QueryStringMessageID)))
	{
	  var dbDatabase = new SWEFDatabase ();
	  currentMessage_disc = dbDatabase.getMessageByID (Request.QueryString (config.FORM_QueryStringMessageID));
	  if (isDefined_disc (currentMessage_disc))
	    {
	      bMessageFound = true;
	    }

	  delete dbDatabase;
	}
    }

  return bMessageFound;
}

function showStandardMessage_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      this.showCurrentMessage ();

      var strHTMLout = "<BR>";
      strHTMLout += SWEFHTML.TABLE_open (config.ADMINSETTING_TableBorderSize);
      strHTMLout += SWEFHTML.TR_open ();
      strHTMLout += SWEFHTML.TD_open ();
      strHTMLout.show ();

//      this.showForumLink ();      
      
//      strHTMLout =  SWEFHTML.TD_close ();
//      strHTMLout += SWEFHTML.TD_open ();
//      strHTMLout.show ();

      this.showNewReplyButton ();

      strHTMLout = SWEFHTML.TD_close ();
      strHTMLout += SWEFHTML.TD_open ();
      strHTMLout.show ();

      if (config.ADMINSWITCH_ShowNewPostButtonOnMessage == true)
	{
	  this.showNewPostButton ();
	}

      strHTMLout = SWEFHTML.TD_close ();
      strHTMLout += SWEFHTML.TD_open ();
      strHTMLout.show ();

      this.showEditPostButton ();

      strHTMLout = SWEFHTML.TD_close ();
      strHTMLout += SWEFHTML.TD_open ();
      strHTMLout.show ();
      
      strHTMLout =  "<FORM NAME=\"DeleteMessage\" ACTION=\"Admin.asp\" METHOD=\"POST\">";
      strHTMLout += "<INPUT TYPE=\"Hidden\" NAME=\"MessageID\" VALUE=\"" + Request("MessageID") + "\">";
      strHTMLout += "<INPUT TYPE=\"SUBMIT\" VALUE=\"Delete Message\" CLASS=NavLeftHighlight1>";
      strHTMLout += "</FORM>";
      strHTMLout.show ();
            
      strHTMLout = SWEFHTML.TD_close ();      
      strHTMLout += SWEFHTML.TR_close ();
      strHTMLout += SWEFHTML.TABLE_close ();
      strHTMLout.show ();
    }

  return;
}

function showForumLink_page_disc
(
)
{
  this.showVersion ();

  var strHTMLout = "";
  strHTMLout += SWEFHTML.P_open ();
  strHTMLout += this.getForumLink (config.USERTEXT_POST_ForumLinkPrefix + config.ADMINSETTING_ForumName).weak ();
  strHTMLout += SWEFHTML.P_close ();
  strHTMLout.show ();

  return;
}

function showForumTitle_page_disc
(
)
{
  config.ADMINSETTING_ForumName.show ();
  return;
}

function showMessageSubject_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  currentMessage_disc.getSubject ().show ();
	}
    }

  return;
}

function showMessageBody_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  currentMessage_disc.getBody ().show ();
	}
    }

  return;
}

function showMessageAuthorName_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  currentMessage_disc.getAuthorName ().show ();
	}
    }

  return;
}

function showMessageAuthorEmail_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  currentMessage_disc.getAuthorEmail ().show ();
	}
    }

  return;
}

function showMessageAuthorFullname_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  currentMessage_disc.getAuthorFullname ().show ();
	}
    }

  return;
}

function showMessageDateCreated_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  String (currentMessage_disc.getDateCreated ()).show ();
	}
    }

  return;
}

function showMessageDateModified_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  String (currentMessage_disc.getDateModified ()).show ();
	}
    }

  return;
}

function showMessageSortCode_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  currentMessage_disc.getSortCode ().show ();
	}
    }

  return;
}

function showMessageNumChildren_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  String (currentMessage_disc.getNumChildren ()).show ();
	}
    }

  return;
}

function showMessageID_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  String (currentMessage_disc.getMessageID ()).show ();
	}
    }

  return;
}

function showMessageParentID_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  String (currentMessage_disc.getParentID ()).show ();
	}
    }

  return;
}

function showMessageSiteID_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  String (currentMessage_disc.getSiteID ()).show ();
	}
    }

  return;
}

function showMessageForumID_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  String (currentMessage_disc.getForumID ()).show ();
	}
    }

  return;
}

function showMessageThreadID_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  String (currentMessage_disc.getThreadID ()).show ();
	}
    }

  return;
}

function showParentMessageLink_page_disc
(
)
{
  this.showVersion ();

  var strLink = this.getParentMessageLink (config.USERTEXT_POST_PreviousMessageLinkText);
  strLink.show ();

  return;
}

function showSearchForm_page_disc
(
)
{
  this.showVersion ();

  var frmSearch = new SWEFForm (Request.Form, SWEF_FORM_OBJECT_TYPE_DISC);

  var strHTMLout = "";
  strHTMLout += frmSearch.getSearchForm (config.USERTEXT_SEARCH_SubmitButton);
  strHTMLout.show ();

  delete frmSearch;
  return;
}

function showSmallSearchForm_page_disc
(
)
{
  this.showVersion ();

  var frmSearch = new SWEFForm (Request.Form, SWEF_FORM_OBJECT_TYPE_DISC);

  var strHTMLout = "";
  strHTMLout += frmSearch.getSmallSearchForm (config.USERTEXT_SEARCH_SubmitButton);
  strHTMLout.show ();

  delete frmSearch;
  return;
}

function showNewPostButton_page_disc
(
)
{
  this.showVersion ();

  var strPostButton = "";
  strPostButton += SWEFHTML.FORM_open (config.getNewPostPagePath (),
				       undefined_disc,
				       undefined_disc,
				       config.ADMINSETTING_NewPostPageTarget);
  strPostButton += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormSearch);
  strPostButton += SWEFHTML.INPUT_submit (config.USERTEXT_SHOW_NewPostButton,
					   config.USERTEXT_SHOW_NewPostButton);
  strPostButton += SWEFHTML.FORM_close ();
  strPostButton += SWEFHTML.SPAN_close ();
  strPostButton.show ();

  return;
}

function showNewPostLink_page_disc
(
)
{
  this.showVersion ();

  var strLink = "";
  strLink += SWEFHTML.A_open (config.getNewPostPagePath (),
			      config.USERTEXT_SHOW_NewPostButton,
			      undefined_disc,
			      config.ADMINSETTING_NewPostPageTarget);
  strLink += config.USERTEXT_SHOW_NewPostButton;
  strLink += SWEFHTML.A_close ();
  strLink.show ();

  return;
}

function showCurrentMessage_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (isDefined_disc (currentMessage_disc))
	{
	  var strMessageHTML = currentMessage_disc.getRenderedHTML ();
	  strMessageHTML.show ();
	}
    }

  return;
}

function showNewReplyButton_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      var strHTMLout = this.getNewReplyButton ();
      strHTMLout.show ();
    }

  return;
}

// Edit Message Button

function showEditPostButton_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      if (((currentMessage_disc.getAuthorName () == currentUsername_disc)
	   && (config.ADMINSWITCH_AllowUserEditing == true))
	  || isAdministrator_disc)
	{
	  var strEditButton = "";
	  strEditButton += SWEFHTML.FORM_open (config.getEditPostPagePath (),
					       undefined_disc,
					       undefined_disc,
					       config.ADMINSETTING_EditPostPageTarget);
	  strEditButton += SWEFHTML.INPUT_hidden (config.FORM_FieldMessageID,
						  currentMessage_disc.getMessageID ());
	  strEditButton += SWEFHTML.INPUT_submit (config.USERTEXT_SHOW_EditPostButton,
						  config.USERTEXT_SHOW_EditPostButton);
	  strEditButton += SWEFHTML.FORM_close ();
	  strEditButton.show ();
	}
    }

  return;
}

function showAllThreads_page_disc
(
)
{
  this.showVersion ();
  this.setCurrentMessage ();

  var thdCurrentThread = new SWEFThread (0);
  var strHTMLout = thdCurrentThread.getAllSorted (config.ADMINSWITCH_ViewPostsAscending);
  strHTMLout.show ();

  delete thdCurrentThread;

  return;
}

function showCurrentThreadsStatic_page_disc
(
)
{
  this.showVersion ();
  this.setCurrentMessage ();

  var nCurrentMessageID = isUndefined_disc (currentMessage_disc) ? 0 : currentMessage_disc.getMessageID ();

  var thdCurrentThread = new SWEFThread (nCurrentMessageID);
  var strHTMLout = thdCurrentThread.getCurrentSorted (config.ADMINSWITCH_ViewPostsAscending);
  strHTMLout.show ();

  delete thdCurrentThread;

  return;
}

function showCurrentThreads_page_disc
(
)
{
  this.showVersion ();
  this.setCurrentMessage ();

  var bExpandTokensPresent = isUndefined_disc (String (Request.QueryString (config.FORM_QueryStringViewExpand))) ? false : true;
  var bCollapseTokensPresent = isUndefined_disc (String (Request.QueryString (config.FORM_QueryStringViewCollapse))) ? false : true;
  var nCurrentMessageID = isUndefined_disc (currentMessage_disc) ? 0 : currentMessage_disc.getMessageID ();

  var thdCurrentThread = new SWEFThread (nCurrentMessageID);
  var strHTMLout = "";
  if ((bExpandTokensPresent) || (bCollapseTokensPresent))
    {
      strHTMLout = thdCurrentThread.getCurrentSorted (config.ADMINSWITCH_ViewPostsAscending);
    }
  else
    {
      strHTMLout = thdCurrentThread.getCurrentSortedDHTML (config.ADMINSWITCH_ViewPostsAscending);
    }
  strHTMLout.show ();

  delete thdCurrentThread;

  return;
}

function showNewPostForm_page_disc
(
)
{
  this.showVersion ();

  var frmNewMessageForm = new SWEFForm (Request.Form, SWEF_FORM_OBJECT_TYPE_DISC);
  var strSubject = frmNewMessageForm.getSubject ().unformatFromStoring ();

  if ((strSubject != "")
      && (frmNewMessageForm.getThreadID () != "")
      && (strSubject.substring(0, config.USERTEXT_POST_ReplySubjectPrefix.length)
	  != config.USERTEXT_POST_ReplySubjectPrefix))
    {
      frmNewMessageForm.setSubject (config.USERTEXT_POST_ReplySubjectPrefix + strSubject);
    }
  else
    {
      frmNewMessageForm.setSubject (strSubject);
    }

  var strHTMLout = "";
  strHTMLout += frmNewMessageForm.getMessageForm (config.getNewPostActionPagePath (),
						  config.USERTEXT_POST_SubmitButton,
						  config.ADMINSETTING_NewPostActionPageTarget);
  strHTMLout.show ();

  delete frmNewMessageForm;
  return;
}

function showEditPostForm_page_disc
(
)
{
  this.showVersion ();

  var frmEditMessageForm = new SWEFForm (Request.Form, SWEF_FORM_OBJECT_TYPE_DISC);
  var dbDatabase = new SWEFDatabase ();

  var msgMessageToEdit = dbDatabase.getMessageByID (frmEditMessageForm.getMessageID ());
  frmEditMessageForm.setAllFields (msgMessageToEdit, SWEF_MESSAGE_OBJECT_TYPE_DISC);

  var strHTMLout = "";
  if ((frmEditMessageForm.getAuthorName () == currentUsername_disc) || isAdministrator_disc)
    {
      strHTMLout += frmEditMessageForm.getMessageForm (config.getEditPostActionPagePath (),
						       config.USERTEXT_POST_SaveChangesButton,
						       config.ADMINSETTING_EditPostActionPageTarget);
    }
  else
    {
      strHTMLout += config.USERTEXT_POST_NotAuthorisedToEditMessage.weak ();
    }
  strHTMLout.show ();

  delete frmEditMessageForm;
  delete dbDatabase;
  delete msgMessageToEdit;
  return;
}

function showCurrentThread_page_disc
(
)
{
  this.showVersion ();

  if (this.setCurrentMessage ())
    {
      var thdExpandedThread = new SWEFThread (currentMessage_disc.getMessageID ());
      var strHTMLout = thdExpandedThread.getFullThread (currentMessage_disc.getThreadID ());
      strHTMLout.show ();
      delete thdExpandedThread;
    }

  return;
}

function showArchiveLink_page_disc
(
)
{
  this.showVersion ();

  var strHTMLout = "";
  strHTMLout += SWEFHTML.A_open (config.getArchivePagePath (),
				 config.USERTEXT_SHOW_PopupArchiveLinkForumPrefix
				 + config.ADMINSETTING_ForumName
				 + config.USERTEXT_SHOW_PopupArchiveLinkForumSuffix,
				 undefined_disc,
				 config.ADMINSETTING_ArchivePageTarget);
//  strHTMLout += config.ADMINSETTING_ForumName;
  strHTMLout += config.USERTEXT_SHOW_ArchiveLinkForumSuffix;
  strHTMLout += SWEFHTML.A_close ();

  strHTMLout = strHTMLout.weak ().show ();

  return;
}

function showArchive_page_disc
(
)
{
  this.showVersion ();

  this.showArchiveSelector ();
  this.actionShowArchive ();

  return;
}

function showArchiveSelector_page_disc
(
)
{
  var frmArchiveForm = new SWEFForm (Request.QueryString, SWEF_FORM_OBJECT_TYPE_DISC);
  var strHTMLout = frmArchiveForm.getArchiveSelector ();
  strHTMLout.show ();

  return;
}

function showAdminSQLForm_page_disc
(
)
{
  this.showVersion ();

  var strHTMLout = "";
//  strHTMLout += SWEFHTML.P_open ();
//  strHTMLout += config.USERTEXT_SQL_EnterPrompt;
//  strHTMLout += SWEFHTML.P_close ();

  var frmAdminForm = new SWEFForm (Request.Form, SWEF_FORM_OBJECT_TYPE_DISC);
  strHTMLout += frmAdminForm.getAdminForm ();
  strHTMLout.show ();

  delete frmAdminForm;
  return;
}

function actionSaveNewMessage_page_disc
(
)
{
  var frmNewMessageForm = new SWEFForm (Request.Form, SWEF_FORM_OBJECT_TYPE_DISC);
  var msgNewMessage = new SWEFMessage (frmNewMessageForm, SWEF_FORM_OBJECT_TYPE_DISC);

  if (!msgNewMessage.saveAsNewMessage (config.USERTEXT_POST_PostSuccessful,
				       config.USERTEXT_POST_PostFailedPrefix))
    {
      var strHTMLout = frmNewMessageForm.getMessageForm (config.getNewPostActionPagePath (),
							 config.USERTEXT_POST_SubmitButton,
							 config.ADMINSETTING_NewPostActionPageTarget);
      strHTMLout.show ();
    }

  currentMessage_disc = msgNewMessage;

  delete frmNewMessageForm;
  return;
}

function actionSaveUpdatedMessage_page_disc
(
)
{
  var frmNewMessageForm = new SWEFForm (Request.Form, SWEF_FORM_OBJECT_TYPE_DISC);
  var msgNewMessage = new SWEFMessage (frmNewMessageForm, SWEF_FORM_OBJECT_TYPE_DISC);

  if (!msgNewMessage.saveAsUpdatedMessage (config.USERTEXT_POST_UpdateSuccessful,
					   config.USERTEXT_POST_UpdateFailedPrefix))
    {
      var strHTMLout = frmNewMessageForm.getMessageForm (config.getEditPostActionPagePath (),
							 config.USERTEXT_POST_SaveChangesButton,
							 config.ADMINSETTING_EditPostActionPageTarget);
      strHTMLout.show ();
    }

  currentMessage_disc = msgNewMessage;

  delete frmNewMessageForm;
  return;
}

function actionSearch_page_disc
(
)
{
  var frmSearchForm = new SWEFForm (Request.Form, SWEF_FORM_OBJECT_TYPE_DISC);
  var strSearchString = frmSearchForm.getSearchString ();

  if ("" != strSearchString)
    {
      var strHTMLout = "";
      var dbDatabase = new SWEFDatabase ();
      var rsFoundMessages = dbDatabase.searchForum (strSearchString);

      strHTMLout += config.USERTEXT_SEARCH_ResultsHeader1;
      strHTMLout += (SWEFHTML.QUOTE_open () + strSearchString + SWEFHTML.QUOTE_close ()).strong();
      strHTMLout += config.USERTEXT_SEARCH_ResultsHeader2;
      strHTMLout += String (rsFoundMessages.RecordCount).strong();

      if (rsFoundMessages.RecordCount == 0)
	{
	  strHTMLout += config.USERTEXT_SEARCH_ResultsHeaderSuffix0Match;
	  strHTMLout += SWEFHTML.BR ();
	}
      else
	{
	  if (rsFoundMessages.RecordCount == 1)
	    {
	      strHTMLout += config.USERTEXT_SEARCH_ResultsHeaderSuffix1Match;
	      strHTMLout += SWEFHTML.BR ();
	    }
	  else
	    {
	      strHTMLout += config.USERTEXT_SEARCH_ResultsHeaderSuffixManyMatches;
	      strHTMLout += SWEFHTML.BR ();
	    }
	}

      if (!rsFoundMessages.EOF)
	{
	  var msgCurrentMessage;

	  rsFoundMessages.MoveFirst ();
	  strHTMLout += SWEFHTML.OL_open ();

	  while (!rsFoundMessages.EOF)
	    {
	      msgCurrentMessage = new SWEFMessage (rsFoundMessages,
						   SWEF_RECORDSET_OBJECT_TYPE_DISC);

	      strHTMLout += msgCurrentMessage.getSummary ();
	      rsFoundMessages.MoveNext ();

	      delete msgCurrentMessage;
	    }
	  strHTMLout += SWEFHTML.OL_close ();
	}

      strHTMLout.show();
      rsFoundMessages.Close ();

      delete rsFoundMessages;
      delete dbDatabase;
    }

  delete frmSearchForm;
  return;
}

function actionExecuteAdminSQL_page_disc
(
)
{
  var strHTMLout = "";
  var frmSQLForm = new SWEFForm (Request.Form, SWEF_FORM_OBJECT_TYPE_DISC);
  var strSQL = frmSQLForm.getSQLStatement ();

  if (strSQL != "")
    {
      var dbDatabase = new SWEFDatabase ();
      var cnDBConnection = dbDatabase.getOpenDatabaseConnection ();
      var rsReturnedRecords;

      rsReturnedRecords = dbDatabase.getAdminSQLResults (cnDBConnection, strSQL);
      strHTMLout += SWEFHTML.TABLE_open (config.ADMINSETTING_TableBorderSize,
					 "100%",
					 "#FFFFFF");
      strHTMLout += SWEFHTML.TR_open ();
      strHTMLout += SWEFHTML.TD_open ();
      strHTMLout += (config.USERTEXT_SQL_ResultsPrefix + SWEFHTML.prototype.BR ()).strong ();
      strHTMLout += config.USERTEXT_SQL_StatementPrefix.weak ();
      strHTMLout += strSQL.strongSmall ();
      strHTMLout += SWEFHTML.BR ();

      strHTMLout += SWEFHTML.TD_close ();
      strHTMLout += SWEFHTML.TR_close ();
      strHTMLout += SWEFHTML.TR_open ();
      strHTMLout += SWEFHTML.TD_open ();

      strHTMLout += this.formatAdminSQLResults (strSQL, rsReturnedRecords);

      strHTMLout += dbDatabase.getConnectionErrors (cnDBConnection).formatForStoring ();

      strHTMLout += SWEFHTML.TD_close ();
      strHTMLout += SWEFHTML.TR_close ();
      strHTMLout += SWEFHTML.TABLE_close ();

      delete dbDatabase;
      delete rsReturnedRecords;
    }

  var nMessageIDToDelete = frmSQLForm.getMessageIDToDelete ();
  if (nMessageIDToDelete != 0)
    {
      if (0 != Number (nMessageIDToDelete))
	{
	  var dbDeleteDBase = new SWEFDatabase ();
	  dbDeleteDBase.deleteMessageHierarchy (nMessageIDToDelete).show ();
	  delete dbDeleteDBase;
	}
    }

  var bPurgeCache = frmSQLForm.getPurgeCache ();
  if (bPurgeCache)
    {
      var cchCurrentCache = new SWEFCache ();
      cchCurrentCache.purge ();
      delete cchCurrentCache;
    }

  strHTMLout.show ();

  delete frmSQLForm;
  return;
}

function actionShowArchive_page_disc
(
)
{
  var frmArchiveForm = new SWEFForm (Request.QueryString, SWEF_FORM_OBJECT_TYPE_DISC);
  var dtArchiveDate = frmArchiveForm.getArchiveDate ();

  if ((isDefined_disc (dtArchiveDate)) && !isNaN (dtArchiveDate))
    {
      var bExpandTokensPresent = isUndefined_disc (String (Request.QueryString (config.FORM_QueryStringViewExpand))) ? false : true;
      var bCollapseTokensPresent = isUndefined_disc (String (Request.QueryString (config.FORM_QueryStringViewCollapse))) ? false : true;

      var thdCurrentThread = new SWEFThread ();
      var strHTMLout = "";
      if ((bExpandTokensPresent) || (bCollapseTokensPresent))
	{
	  strHTMLout += thdCurrentThread.getArchiveSorted (dtArchiveDate,
							   config.ADMINSWITCH_ViewPostsAscending);
	}
      else
	{
	  strHTMLout += thdCurrentThread.getArchiveSortedDHTML (dtArchiveDate,
								config.ADMINSWITCH_ViewPostsAscending);
	}
      strHTMLout.show ();

      delete thdCurrentThread;
    }

  return;
}

// -
// Private 'get' routines
//

function getForumLink_page_disc
(
 strTextToUse
)
{
  var strHTMLout = "";
  strHTMLout += SWEFHTML.A_open (config.getMainPagePreferredPath (),
				 config.ADMINSETTING_ForumName,
				 undefined_disc,
				 config.ADMINSETTING_TopPageTarget);
  strHTMLout += strTextToUse;
  strHTMLout += SWEFHTML.A_close ();

  return strHTMLout;
}

function getParentMessageLink_page_disc
(
 strTextToUse
)
{
  var strHTMLout = "";
  var frmMessageForm = new SWEFForm (Request.Form, SWEF_FORM_OBJECT_TYPE_DISC);
  // Need Site ID Here ?
  var nParentID = frmMessageForm.getParentID ();
  if (nParentID != 0)
    {
      strHTMLout += SWEFHTML.P_open ();
      strHTMLout += SWEFHTML.A_open (config.getShowMessagePagePath ()
				     + "?"
				     + config.FORM_QueryStringMessageID
				     + "="
				     + nParentID,
				     strTextToUse,
				     undefined_disc,
				     config.ADMINSETTING_ShowMessagePageTarget);
      strHTMLout += strTextToUse.weak ();
      strHTMLout += SWEFHTML.A_close ();
      strHTMLout += SWEFHTML.P_close ();
    }

  delete frmMessageForm;
  return strHTMLout;
}

function formatAdminSQLResults_page_disc
(
 strSQL,
 rsRecordSet
)
{
  var strHTMLout = "";

  if (String (strSQL).search(/select/gi) > -1)
    {
      strHTMLout += SWEFHTML.TABLE_open (2, undefined_disc, "#FFFFFF");
      strHTMLout += SWEFHTML.TR_open ();

      var nCounter;
      var nNumRecordFields = rsRecordSet.fields.count - 1;
      for (nCounter = 0; nCounter <= nNumRecordFields; nCounter++)
	{
	  strHTMLout += SWEFHTML.TH_open (undefined_disc, "left", undefined_disc, true);
	  strHTMLout += String (rsRecordSet (nCounter).name).strong();
	  strHTMLout += SWEFHTML.TH_close ();
	}
      strHTMLout += SWEFHTML.TR_close ();

      while (!rsRecordSet.EOF)
	{
	  strHTMLout += SWEFHTML.TR_open ();

	  for (nCounter = 0; nCounter <= nNumRecordFields; nCounter++)
	    {
	      strHTMLout += SWEFHTML.TD_open (undefined_disc,
					      undefined_disc,
					      undefined_disc,
					      undefined_disc,
					      undefined_disc,
					      undefined_disc,
					      true);
	      strHTMLout += rsRecordSet.Fields (nCounter).value;
	      strHTMLout += SWEFHTML.TD_close ();
	    }

	  strHTMLout += SWEFHTML.TR_close ();
	  rsRecordSet.MoveNext ();
	}

      strHTMLout += SWEFHTML.TABLE_close ();
    }

  return strHTMLout;
}

function getNewReplyButton_page_disc
(
)
{
  this.setCurrentMessage ();

  return this.getNewReplyButtonWithTextAndMessage (config.USERTEXT_SHOW_ReplyPostButton,
						   currentMessage_disc);
}

function getNewReplyButtonWithTextAndMessage_page_disc
(
 strButtonText,
 msgMessage
)
{
  this.setCurrentMessage ();

  var strHTMLout = "";

  if (isDefined_disc (msgMessage))
    {
      strHTMLout += SWEFHTML.FORM_open (config.getNewPostPagePath (),
					undefined_disc,
					undefined_disc,
					config.ADMINSETTING_NewPostPageTarget);
//      strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldSiteID, msgMessage.getSiteID ());
//      strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldForumID, msgMessage.getForumID ());                          
      strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldParentID, msgMessage.getMessageID ());
      strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldThreadID, msgMessage.getThreadID ());
      strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldSortCode, msgMessage.getSortCode ());
      strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldSubject, msgMessage.getSubject ());
      strHTMLout += SWEFHTML.INPUT_submit (strButtonText, strButtonText);
      strHTMLout += SWEFHTML.FORM_close ();
    }

  return strHTMLout;
}

function doNothing_page_disc
(
)
{
  return;
}

SWEFPageElement.versionRefOutput = false;

SWEFPageElement.showVersion = showVersion_page_disc;
SWEFPageElement.setCurrentMessage = setCurrentMessage_page_disc;
SWEFPageElement.showStandardMessage = showStandardMessage_page_disc;
SWEFPageElement.showForumLink = showForumLink_page_disc;
SWEFPageElement.showForumTitle = showForumTitle_page_disc;

SWEFPageElement.showMessageSubject = showMessageSubject_page_disc;
SWEFPageElement.showMessageBody = showMessageBody_page_disc;
SWEFPageElement.showMessageAuthorName = showMessageAuthorName_page_disc;
SWEFPageElement.showMessageAuthorEmail = showMessageAuthorEmail_page_disc;
SWEFPageElement.showMessageAuthorFullname = showMessageAuthorFullname_page_disc;
SWEFPageElement.showMessageDateCreated = showMessageDateCreated_page_disc;
SWEFPageElement.showMessageDateModified = showMessageDateModified_page_disc;
SWEFPageElement.showMessageSortCode = showMessageSortCode_page_disc;
SWEFPageElement.showMessageNumChildren = showMessageNumChildren_page_disc;
SWEFPageElement.showMessageID = showMessageID_page_disc;
SWEFPageElement.showMessageSiteID = showMessageSiteID_page_disc;
SWEFPageElement.showMessageForumID = showMessageForumID_page_disc;
SWEFPageElement.showMessageParentID = showMessageParentID_page_disc;
SWEFPageElement.showMessageThreadID = showMessageThreadID_page_disc;

SWEFPageElement.showParentMessageLink = showParentMessageLink_page_disc;
SWEFPageElement.showSearchForm = showSearchForm_page_disc;
SWEFPageElement.showSmallSearchForm = showSmallSearchForm_page_disc;
SWEFPageElement.showNewPostButton = showNewPostButton_page_disc;
SWEFPageElement.showNewPostLink = showNewPostLink_page_disc;
SWEFPageElement.showCurrentMessage = showCurrentMessage_page_disc;
SWEFPageElement.showNewReplyButton = showNewReplyButton_page_disc;
SWEFPageElement.showEditPostButton = showEditPostButton_page_disc;
SWEFPageElement.showAllThreads = showAllThreads_page_disc;
SWEFPageElement.showCurrentThreadsStatic = showCurrentThreadsStatic_page_disc;
SWEFPageElement.showCurrentThreads = showCurrentThreads_page_disc;
SWEFPageElement.showNewPostForm = showNewPostForm_page_disc;
SWEFPageElement.showEditPostForm = showEditPostForm_page_disc;
SWEFPageElement.showCurrentThread = showCurrentThread_page_disc;
SWEFPageElement.showArchiveLink = showArchiveLink_page_disc;
SWEFPageElement.showArchive = showArchive_page_disc;
SWEFPageElement.showArchiveSelector = showArchiveSelector_page_disc;
SWEFPageElement.showAdminSQLForm = showAdminSQLForm_page_disc;

SWEFPageElement.actionSaveUpdatedMessage = actionSaveUpdatedMessage_page_disc;
SWEFPageElement.actionSaveNewMessage = actionSaveNewMessage_page_disc;
SWEFPageElement.actionSearch = actionSearch_page_disc;
SWEFPageElement.actionExecuteAdminSQL = actionExecuteAdminSQL_page_disc;
SWEFPageElement.actionShowArchive = actionShowArchive_page_disc;

SWEFPageElement.getForumLink = getForumLink_page_disc;
SWEFPageElement.getParentMessageLink = getParentMessageLink_page_disc;
SWEFPageElement.formatAdminSQLResults = formatAdminSQLResults_page_disc;
SWEFPageElement.getNewReplyButton = getNewReplyButton_page_disc;
SWEFPageElement.getNewReplyButtonWithTextAndMessage = getNewReplyButtonWithTextAndMessage_page_disc;
</SCRIPT>

