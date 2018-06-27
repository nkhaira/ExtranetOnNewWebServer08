<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// FORM OBJECT
//
// ======================================================================

function SWEFForm
(
 objFormData,
 nSourceType
)
{
  this.setAllFields (objFormData, nSourceType);

  return this;
}

// ======================================================================
//
// Interface to private member variables.
//
// ======================================================================

function getMessageID_frm_disc
(
)
{
  return Number (this._messageID);
}

function setMessageID_frm_disc
(
 nNewMessageID
)
{
  this._messageID = Number (nNewMessageID);
  return;
}

function getSubject_frm_disc
(
)
{
  return String (this._subject);
}

function setSubject_frm_disc
(
 strNewSubject
)
{
  this._subject = String (strNewSubject);
  return;
}

function getBody_frm_disc
(
)
{
  return String (this._body);
}

function setBody_frm_disc
(
 strNewBody
)
{
  this._body = String (strNewBody);
  return;
}

function getSortCode_frm_disc
(
)
{
  return String (this._sortCode);
}

function setSortCode_frm_disc
(
 strNewSortCode
)
{
  this._sortCode = String (strNewSortCode);
  return;
}

function getParentID_frm_disc
(
)
{
  return Number (this._parentID);
}

function getSiteID_frm_disc
(
)
{
  return Number (this._siteID);
}

function getForumID_frm_disc
(
)
{
  return Number (this._forumID);
}

function setParentID_frm_disc
(
 nNewParentID
)
{
  this._parentID = Number (nNewParentID);
  return;
}

function setSiteID_frm_disc
(
 nNewSiteID
)
{
  this._siteID = Number (nNewSiteID);
  return;
}

function setForumID_frm_disc
(
 nNewForumID
)
{
  this._forumID = Number (nNewForumID);
  return;
}


function getThreadID_frm_disc
(
)
{
  return Number (this._threadID);
}

function setThreadID_frm_disc
(
 nNewThreadID
)
{
  this._threadID = Number (nNewThreadID);
  return;
}

function getEmailParentOnResponse_frm_disc
(
)
{
  return Boolean (this._emailParentOnResponse);
}

function setEmailParentOnResponse_frm_disc
(
 bNewEmailParentOnResponse
)
{
  this._emailParentOnResponse = Boolean (bNewEmailParentOnResponse);
  return;
}

function getSearchString_frm_disc
(
)
{
  return String (this._searchString);
}

function setSearchString_frm_disc
(
 strNewSearchString
)
{
  this._searchString = String (strNewSearchString);
  return;
}

function getSQLStatement_frm_disc
(
)
{
  return String (this._SQLStatement).replace (/\"/gi, "'");
}

function setSQLStatement_frm_disc
(
 strNewSQLStatement
)
{
  this._SQLStatement = String (strNewSQLStatement);
  return;
}

function getFormName_frm_disc
(
)
{
  return String (this._formName);
}

function setFormName_frm_disc
(
 strNewFormName
)
{
  this._formName = String (strNewFormName);
  return;
}

function getMessageIDToDelete_frm_disc
(
)
{
  var nMessageID = Number (this._messageIDToDelete);
  if (isNaN (nMessageID))
    {
      nMessageID = 0;
    }

  return nMessageID;
}

function setMessageIDToDelete_frm_disc
(
 nNewMessageIDToDelete
)
{
  this._messageIDToDelete = Number (nNewMessageIDToDelete);
  return;
}

function getSwitchEditableUserInfo_frm_disc
(
)
{
  return Boolean (this._switchEditableUserInfo);
}

function setSwitchEditableUserInfo_frm_disc
(
 bNewSwitchEditableUserInfo
)
{
  this._switchEditableUserInfo = Boolean (bNewSwitchEditableUserInfo);
  return;
}

function getAuthorName_frm_disc
(
)
{
  return String (this._authorName);
}

function setAuthorName_frm_disc
(
 strNewAuthorName
)
{
  this._authorName = String (strNewAuthorName);
  return;
}

function getAuthorEmail_frm_disc
(
)
{
  return String (this._authorEmail);
}

function setAuthorEmail_frm_disc
(
 strNewAuthorEmail
)
{
  this._authorEmail = String (strNewAuthorEmail);
  return;
}

function getAuthorFullname_frm_disc
(
)
{
  return String (this._authorFullname);
}

function setAuthorFullname_frm_disc
(
 strNewAuthorFullname
)
{
  this._authorFullname = String (strNewAuthorFullname);
  return;
}

function getArchiveDate_frm_disc
(
)
{
  return new Date (this._archiveDate);
}

function setArchiveDate_frm_disc
(
 dtNewArchiveDate
)
{
  this._archiveDate = new Date (dtNewArchiveDate);
  return;
}

function getPurgeCache_frm_disc
(
)
{
  return Boolean (this._purgeCache);
}

function setPurgeCache_frm_disc
(
 bNewPurgeCache
)
{
  this._purgeCache = Boolean (bNewPurgeCache);
  return;
}

// ======================================================================
//
// Main object methods.
//
// ======================================================================

function setAllFields_frm_disc
(
 objNewFormData,
 nSourceType
)
{
  if (nSourceType == SWEF_MESSAGE_OBJECT_TYPE_DISC)
    {
      this.setAllFieldsFromMessage (objNewFormData);
    }
  else if (nSourceType == SWEF_FORM_OBJECT_TYPE_DISC)
    {
      this.setAllFieldsFromForm (objNewFormData);
    }
  else
    {
      this.setAllFieldsEmpty ();
    }

  return;
}

function setAllFieldsFromMessage_frm_disc
(
 msgMessageForForm
)
{
  this.setSubject (msgMessageForForm.getSubject ());
  this.setBody (msgMessageForForm.getBody ());
  this.setSortCode (msgMessageForForm.getSortCode ());
  this.setSiteID (msgMessageForForm.getSiteID ());
  this.setForumID (msgMessageForForm.getForumID ());
  this.setParentID (msgMessageForForm.getParentID ());
  this.setThreadID (msgMessageForForm.getThreadID ());
  this.setMessageID (msgMessageForForm.getMessageID ());
  this.setEmailParentOnResponse (msgMessageForForm.getEmailParentOnResponse ());
  this.setSearchString ("");
  this.setSQLStatement ("");
  this.setSwitchEditableUserInfo (config.getEditableUserInfoSwitch ());
  this.setAuthorName (msgMessageForForm.getAuthorName ());
  this.setAuthorEmail (msgMessageForForm.getAuthorEmail ());
  this.setAuthorFullname (msgMessageForForm.getAuthorFullname ());

  return;
}

function setAllFieldsFromForm_frm_disc
(
 objFormData
)
{
  this.setMessageID (safeStringDereference_disc (objFormData (config.FORM_FieldMessageID)));
  this.setSubject (safeStringDereference_disc (objFormData (config.FORM_FieldSubject)));
  this.setBody (safeStringDereference_disc (SWEFTextControl.getControlContents (objFormData, config.FORM_FieldMessage)));
  this.setSortCode (safeStringDereference_disc (objFormData (config.FORM_FieldSortCode)));  
  this.setSiteID (safeStringDereference_disc (objFormData (config.FORM_FieldSiteID)));  
  this.setForumID (safeStringDereference_disc (objFormData (config.FORM_FieldForumID)));  
  this.setParentID (safeStringDereference_disc (objFormData (config.FORM_FieldParentID)));
  this.setThreadID (safeStringDereference_disc (objFormData (config.FORM_FieldThreadID)));
  this.setEmailParentOnResponse ((isUndefined_disc (objFormData (config.FORM_FieldEmailResponses)) ? config.getUserEmailResponsePreference (currentUsername_disc) : (objFormData (config.FORM_FieldEmailResponses) == config.FORM_CheckboxTrue ? true : false)));
  this.setSearchString (safeStringDereference_disc (objFormData (config.FORM_FieldSearchString)));
  this.setSQLStatement (safeStringDereference_disc (objFormData (config.FORM_FieldSQLStatement)));
  this.setMessageIDToDelete (safeStringDereference_disc (objFormData (config.FORM_FieldMessageIDToDelete)));
  this.setSwitchEditableUserInfo (config.getEditableUserInfoSwitch ());
  this.setArchiveDate (this.interpretMonthYearSubform (objFormData, config.FORM_FieldArchiveDate));
  this.setPurgeCache (safeStringDereference_disc (objFormData (config.FORM_FieldPurgeCache)) == config.FORM_CheckboxTrue ? true : false);

  if (this.getSwitchEditableUserInfo ())
    {
      this.setAuthorName (safeStringDereference_disc (objFormData (config.FORM_FieldUsername)));
      this.setAuthorEmail (safeStringDereference_disc (objFormData (config.FORM_FieldEmailaddress)));
      this.setAuthorFullname (safeStringDereference_disc (objFormData (config.FORM_FieldFullname)));
    }
  else
    {
      this.setAuthorName (currentUsername_disc);
      this.setAuthorEmail (currentUserEmailAddress_disc);
      this.setAuthorFullname (currentUserFullName_disc);
    }

  return;
}

function setAllFieldsEmpty_frm_disc
(
)
{
  this.setSubject ("");
  this.setBody ("");
  this.setSortCode ("");
  this.setSiteID (Session("Site_ID"));
  this.setForumID (Session("Asset_ID"));
  this.setParentID ("");
  this.setThreadID ("");
  this.setMessageID ("");
  this.setEmailParentOnResponse (false);
  this.setSearchString ("");
  this.setMessageIDToDelete ("");
  this.setSQLStatement ("");
  this.setSwitchEditableUserInfo (config.getEditableUserInfoSwitch ());
  this.setArchiveDate (undefined_disc);
  this.setPurgeCache (false);

  if (this.getSwitchEditableUserInfo ())
    {
      this.setAuthorName (this (config.FORM_FieldUsername));
      this.setAuthorEmail (this (config.FORM_FieldEmailaddress));
      this.setAuthorFullname (this (config.FORM_FieldFullname));
    }
  else
    {
      this.setAuthorName (currentUsername_disc);
      this.setAuthorEmail (currentUserEmailAddress_disc);
      this.setAuthorFullname (currentUserFullName_disc);
    }

  return;
}

function getSubjectInputField_frm_disc
(
)
{
  var strSubject = this.getSubject ();

  return SWEFHTML.INPUT_text (config.FORM_FieldSubject,
			      strSubject,
			      config.ADMINSETTING_SubjectInputboxSize);
}

function getBodyInputField_frm_disc
(
)
{
  var txtBodyTextControl = new SWEFTextControl (config.FORM_FieldMessage,
						this.getBody (),
						this.getFormName ());
  return txtBodyTextControl.getControl ();
}

function getSubjectSubform_frm_disc
(
)
{
  var strHTMLout = "";

  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom");
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormSubjectLabel);
  strHTMLout += config.USERTEXT_POST_SubjectPrompt.weak ();
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();

  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom");
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormSubject);
  strHTMLout += this.getSubjectInputField ();
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();

  return strHTMLout;
}

function getBodySubform_frm_disc
(
)
{
  var strHTMLout = "";

  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  2);
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormBodyLabel);
  strHTMLout += (SWEFHTML.NBSP () + config.USERTEXT_POST_BodyPrompt).weak ();
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.BR ();
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormBody);
  strHTMLout += this.getBodyInputField ();
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();

  return strHTMLout;
}

// Send Email when someone responds to my Message

function getEmailResponsesSubform_frm_disc
(
)
{
  var strEmailResponsesChecked;
  if ((this.getEmailParentOnResponse () == true)
      || (this.getEmailParentOnResponse () == config.FORM_CheckboxTrue))
    {
      strEmailResponsesChecked = config.FORM_CheckboxChecked;
    }
  else
    {
      strEmailResponsesChecked = config.FORM_CheckboxUnchecked;
    }

  var strHTMLout = "";
  if (config.ADMINSWITCH_AllowEmailResponses)
    {
      strHTMLout += SWEFHTML.TR_open ();
      strHTMLout += SWEFHTML.TD_open (undefined_disc,
				      undefined_disc,
				      "Silver",
				      undefined_disc,
				      undefined_disc,
				      2);
      strHTMLout += SWEFHTML.P_open ();
      strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormEmailResponses);
      strHTMLout += SWEFHTML.INPUT_checkbox (config.FORM_FieldEmailResponses,
					     config.FORM_CheckboxTrue,
					     strEmailResponsesChecked);
      strHTMLout += SWEFHTML.SPAN_close ();
      strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormEmailResponsesLabel);
      strHTMLout += config.USERTEXT_POST_EmailResponsesPrompt.weak ();
      strHTMLout += SWEFHTML.SPAN_close ();
      strHTMLout += SWEFHTML.P_close ();
      strHTMLout += SWEFHTML.TD_close ();
      strHTMLout += SWEFHTML.TR_close ();
    }

  return strHTMLout;
}

function getNonEditableUserInfoSubform_frm_disc
(
)
{
  var strUsername = "";
  strUsername += this.getAuthorName ();
  strUsername += " (";
  strUsername += SWEFHTML.QUOTE_open ();
  strUsername += this.getAuthorFullname ();
  strUsername += SWEFHTML.QUOTE_close ();
  strUsername += ", ";
  strUsername += this.getAuthorEmail ();
  strUsername += ")";
  strUsername = strUsername.weak ();

  var strHTMLout = "";
  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom");
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormPostedByLabel);
  strHTMLout += config.USERTEXT_POST_PostedByPrompt.weak ();
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();

  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom");
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormPostedBy);
  strHTMLout += strUsername.strong ();
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();

  return strHTMLout;
}

function getEditableUserInfoSubform_frm_disc
(
)
{
  var strHTMLout = "";
  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom");
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormPostedByLabel);
  strHTMLout += config.USERTEXT_POST_PostedByPrompt.weak ();
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();

  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom");
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormPostedBy);
  strHTMLout += SWEFHTML.INPUT_text (config.FORM_FieldUsername,
				     this.getAuthorName (),
				     25,
				     50);
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();

  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom");
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormFullNameLabel);
  strHTMLout += config.USERTEXT_POST_FullnamePrompt.weak ();
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();

  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom");
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormFullName);
  strHTMLout += SWEFHTML.INPUT_text (config.FORM_FieldFullname,
				     this.getAuthorFullname (),
				     25,
				     50);
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();

  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableTitleColumnWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom");
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormEmailAddressLabel);
  strHTMLout += config.USERTEXT_POST_EmailAddressPrompt.weak ();
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();

  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableFieldColumnWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom");
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormEmailAddress);
  strHTMLout += SWEFHTML.INPUT_text (config.FORM_FieldEmailaddress,
				     this.getAuthorEmail (),
				     25,
				     50);
  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();

  return strHTMLout;
}

function getUserInfoSubform_frm_disc
(
)
{
  var strUserInfoSubform = "";

  if (this.getSwitchEditableUserInfo ())
    {
      strUserInfoSubform = this.getEditableUserInfoSubform ();
    }
  else
    {
      strUserInfoSubform = this.getNonEditableUserInfoSubform ();
    }

  return strUserInfoSubform;
}

function getDateSubform_frm_disc
(
 dtFieldValue,
 strFieldName,
 nStartYear,
 nEndYear
)
{
  if ((isUndefined_disc (dtFieldValue)) || isNaN (dtFieldValue))
    {
      dtFieldValue = new Date ();
    }
  else
    {
      dtFieldValue = new Date (dtFieldValue);
    }

  var strHTMLout = "";
  strHTMLout += this.getDateElement (strFieldName,
				     dtFieldValue.getDate ());
  strHTMLout += this.getMonthElement (strFieldName,
				      dtFieldValue.getMonth ());
  strHTMLout += this.getYearElement (strFieldName,
				     dtFieldValue.getFullYear (),
				     nStartYear, nEndYear);

  return strHTMLout;
}

function getMonthYearSubform_frm_disc
(
 dtFieldValue,
 strFieldName,
 nStartYear,
 nEndYear
)
{
  if ((isUndefined_disc (dtFieldValue)) || isNaN (dtFieldValue))
    {
      dtFieldValue = new Date ();
    }
  else
    {
      dtFieldValue = new Date (dtFieldValue);
    }

  var strHTMLout = "";
  strHTMLout += this.getMonthElement (strFieldName,
				      dtFieldValue.getMonth ());
  strHTMLout += this.getYearElement (strFieldName,
				     dtFieldValue.getFullYear (),
				     nStartYear,
				     nEndYear);

  return strHTMLout;
}

function getMessageForm_frm_disc
(
 strActionPath,
 strButtonLabel,
 strURLTarget
)
{
  this.setFormName (config.FORM_MessageFormName);
  var strHTMLout = "";
  strHTMLout += SWEFHTML.FORM_open (strActionPath, undefined_disc, this.getFormName (), strURLTarget);
  strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldMessageID, this.getMessageID ());
//  strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldSiteID, this.getSiteID ());
//  strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldForumID, this.getForumID ());   
//  strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldSiteID, this.getSiteID ());
//  strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldForumID, this.getForumID ());  
  strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldParentID, this.getParentID ());
  strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldThreadID, this.getThreadID ());
  strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldSortCode, this.getSortCode ());
  strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldHiddenEmailOnResponse,
				       this.getEmailParentOnResponse ());

  strHTMLout += SWEFHTML.TABLE_open (config.ADMINSETTING_TableBorderSize,
				     config.ADMINSETTING_TableFullWidth,
				     undefined_disc,
				     undefined_disc,
				     0,
				     0);

  strHTMLout += this.getUserInfoSubform ();

  strHTMLout += this.getSubjectSubform ();

  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  2);
  strHTMLout += SWEFHTML.NBSP ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();

  strHTMLout += this.getBodySubform ();

  strHTMLout += this.getEmailResponsesSubform ();

  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_TableFullWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "bottom",
				  2);
  strHTMLout += SWEFHTML.INPUT_submit (strButtonLabel,
				        strButtonLabel);
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();
  strHTMLout += SWEFHTML.TABLE_close ();
  
  strHTMLout += SWEFHTML.FORM_close ();

  return strHTMLout;
}

function getSearchForm_frm_disc
(
 strLabel
)
{
  this.setFormName (config.FORM_SearchFormName);
  var strHTMLout = "";
  strHTMLout = SWEFHTML.FORM_open (config.getSearchPagePath (),
				   undefined_disc,
				   this.getFormName (),
				   config.ADMINSETTING_SearchPageTarget);
//  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormSearch);
  strHTMLout += "<FONT CLASS=SmallBold>"
  strHTMLout += config.USERTEXT_SEARCH_SmallSearchPrefix;
  strHTMLout += "&nbsp;";
  strHTMLout += SWEFHTML.INPUT_text (config.FORM_FieldSearchString,
				     this.getSearchString (),
				     "25");
  strHTMLout += "</FONT>&nbsp;&nbsp;";
  strHTMLout += SWEFHTML.INPUT_submit (strLabel, strLabel);
//  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.FORM_close ();

  return strHTMLout;
}

function getSmallSearchForm_frm_disc
(
 strLabel
)
{
  this.setFormName (config.FORM_SearchFormName);

  var strHTMLout = "";
  strHTMLout += SWEFHTML.FORM_open (config.getSearchPagePath (),
				    undefined_disc,
				    this.getFormName (),
				    config.ADMINSETTING_SearchPageTarget);
//  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormSearchSmall);
  strHTMLout += config.USERTEXT_SEARCH_SmallSearchPrefix.weakSmall ();
//  strHTMLout += SWEFHTML.BR ();
  strHTMLout += SWEFHTML.INPUT_text (config.FORM_FieldSearchString,
				     this.getSearchString (),
				     "10");
//  strHTMLout += SWEFHTML.SPAN_close ();
  strHTMLout += SWEFHTML.FORM_close ();

  return strHTMLout;
}

function getArchiveSelector_frm_disc
(
)
{
  var dtCurrentDate = new Date ();
  var nCurrentYear = dtCurrentDate.getFullYear ();

  this.setFormName (config.FORM_ArchiveFormName);

  var strHTMLout = "";
  strHTMLout += SWEFHTML.FORM_open (config.getArchivePagePath (),
				    "get",
				    this.getFormName (),
				    config.ADMINSETTING_ArchivePageTarget);
  strHTMLout += this.getMonthYearSubform (this.getArchiveDate (),
					  config.FORM_FieldArchiveDate,
					  config.ADMINSETTING_ArchiveBeginYear,
					  nCurrentYear);
  strHTMLout += SWEFHTML.INPUT_submit (undefined_disc, config.USERTEXT_ARCHIVE_ShowResults);
  strHTMLout += SWEFHTML.FORM_close ();

  return strHTMLout;
}

function getAdminForm_frm_disc
(
)
{
  this.setFormName (config.FORM_AdminDeleteFormName);

  var strHTMLout = "";
  strHTMLout += SWEFHTML.FORM_open (config.getAdminPagePath (),
				    undefined_disc,
				    this.getFormName (),
				    config.ADMINSETTING_AdminPageTarget);

  strHTMLout += "<FONT CLASS=Medium><BR><LI></LI>Clicking on the [Delete Message] button will delete the current message that you had selected from the previous screen and its subthread messages.</FONT></LI><BR><BR><HR NOSHADE COLOR=BLACK><BR>";

  if (Request("MessageID") == "")
    {
    strHTMLout += "<FONT CLASS=Medium>Message ID: </FONT>";
    strHTMLout += SWEFHTML.INPUT_text (config.FORM_FieldMessageIDToDelete, Request("MessageID"), 10);
  }  
  else
    {
    strHTMLout += "<FONT CLASS=Medium>Message ID " + Request("MessageID") + ": </FONT>";
    strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldMessageIDToDelete, Request("MessageID"), 10);
  }
          
  strHTMLout += SWEFHTML.NBSP (4);
  strHTMLout += SWEFHTML.INPUT_submit (config.USERTEXT_SQL_DeleteHierarchyButton,
				       config.USERTEXT_SQL_DeleteHierarchyButton);
  strHTMLout += SWEFHTML.FORM_close ();

  if (Request("MessageID") == "")
    {
    this.setFormName (config.FORM_AdminPurgeFormName);

    strHTMLout += SWEFHTML.FORM_open (config.getAdminPagePath (),
		  		    undefined_disc,
  				    this.getFormName (),
  				    config.ADMINSETTING_AdminPageTarget);
    strHTMLout += SWEFHTML.INPUT_hidden (config.FORM_FieldPurgeCache,
  				       config.FORM_CheckboxTrue);
    strHTMLout += SWEFHTML.INPUT_submit (config.USERTEXT_SQL_PurgeCacheButton,
	  			       config.USERTEXT_SQL_PurgeCacheButton);
    strHTMLout += SWEFHTML.FORM_close ();

    this.setFormName (config.FORM_AdminSQLFormName);

    strHTMLout += SWEFHTML.FORM_open (config.getAdminPagePath (),
  				    undefined_disc,
  				    this.getFormName (),
  				    config.ADMINSETTING_AdminPageTarget);
    strHTMLout += SWEFHTML.TEXTAREA_open (config.FORM_FieldSQLStatement,
  					config.ADMINSETTING_TextAreaCols,
  					config.ADMINSETTING_TextAreaRows);
    strHTMLout += this.getSQLStatement ();
    strHTMLout += SWEFHTML.TEXTAREA_close ();
    strHTMLout += SWEFHTML.BR ();
    strHTMLout += SWEFHTML.INPUT_submit (config.USERTEXT_SQL_ExecutePrompt,
  				       config.USERTEXT_SQL_ExecutePrompt);
    strHTMLout += SWEFHTML.FORM_close ();
  }

  return strHTMLout;
}

function getDateElement_frm_disc
(
 strFieldName,
 strFieldValue
)
{
  var strHTMLout = "";
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormDateElement);
  strHTMLout += SWEFHTML.SELECT_open (strFieldName + config.FORM_FieldDateDaySuffix);
  for (var nCounter = 1; nCounter < 32; nCounter++)
    {
      strHTMLout += SWEFHTML.OPTION (nCounter, (nCounter == strFieldValue));
      strHTMLout += nCounter + Date.getDateSuffixByIndex (nCounter);
    }

  strHTMLout += SWEFHTML.SELECT_close ();
  strHTMLout += SWEFHTML.SPAN_close ();

  return strHTMLout;
}

function getMonthElement_frm_disc
(
 strFieldName,
 strFieldValue
)
{
  var strHTMLout = "";
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormMonthElement);
  strHTMLout += SWEFHTML.SELECT_open (strFieldName + config.FORM_FieldDateMonthSuffix);
  for (var nCounter = 0; nCounter < 12; nCounter++)
    {
      strHTMLout += SWEFHTML.OPTION (nCounter, (nCounter == strFieldValue));
      strHTMLout += Date.getMonthNameByIndex (nCounter);
    }

  strHTMLout += SWEFHTML.SELECT_close ();
  strHTMLout += SWEFHTML.SPAN_close ();

  return strHTMLout;
}

function getYearElement_frm_disc
(
 strFieldName,
 strFieldValue,
 nStartYear,
 nEndYear
)
{
  var strHTMLout = "";
  strHTMLout += SWEFHTML.SPAN_open (config.ADMINSETTING_CSSClassFormYearElement);
  strHTMLout += SWEFHTML.SELECT_open (strFieldName + config.FORM_FieldDateYearSuffix);
  for (var nCounter = nStartYear; nCounter < (nEndYear + 1); nCounter++)
    {
      strHTMLout += SWEFHTML.OPTION (nCounter, (nCounter == strFieldValue));
      strHTMLout += nCounter;
    }

  strHTMLout += SWEFHTML.SELECT_close ();
  strHTMLout += SWEFHTML.SPAN_close ();

  return strHTMLout;
}

function interpretDateSubform_frm_disc
(
 objForm,
 strFieldName
)
{
  var dtFieldDate;

  if (isDefined_disc (objForm (strFieldName + config.FORM_FieldDateMonthSuffix)))
    {
      var nDateValue = safeNumberDereference_disc (String (objForm (strFieldName + config.FORM_FieldDateDaySuffix)));
      var nMonthValue = safeNumberDereference_disc (String (objForm (strFieldName + config.FORM_FieldDateMonthSuffix)));
      var nYearValue = safeNumberDereference_disc (String (objForm (strFieldName + config.FORM_FieldDateYearSuffix)));
      dtFieldDate = new Date (config.SYS_SafeDefaultDate);
      dtFieldDate.setYear (nYearValue);
      dtFieldDate.setMonth (nMonthValue);
      dtFieldDate.setDate (nDateValue);
    }

  return dtFieldDate;
}

function interpretMonthYearSubform_frm_disc
(
 objForm,
 strFieldName
)
{
  var dtFieldDate;

  if (isDefined_disc (String (objForm (strFieldName + config.FORM_FieldDateMonthSuffix))))
    {
      var nMonthValue = safeNumberDereference_disc (String (objForm (strFieldName + config.FORM_FieldDateMonthSuffix)));
      var nYearValue = safeNumberDereference_disc (String (objForm (strFieldName + config.FORM_FieldDateYearSuffix)));
      dtFieldDate = new Date (config.SYS_SafeDefaultDate);
      dtFieldDate.setYear (nYearValue);
      dtFieldDate.setMonth (nMonthValue);
      dtFieldDate.setDate (1);
    }

  return dtFieldDate;
}

SWEFForm.prototype.getMessageID = getMessageID_frm_disc;
SWEFForm.prototype.setMessageID = setMessageID_frm_disc;
SWEFForm.prototype.getSubject = getSubject_frm_disc;
SWEFForm.prototype.setSubject = setSubject_frm_disc;
SWEFForm.prototype.getBody = getBody_frm_disc;
SWEFForm.prototype.setBody = setBody_frm_disc;
SWEFForm.prototype.getSortCode = getSortCode_frm_disc;
SWEFForm.prototype.setSortCode = setSortCode_frm_disc;
SWEFForm.prototype.getSiteID = getSiteID_frm_disc;
SWEFForm.prototype.setSiteID = setSiteID_frm_disc;
SWEFForm.prototype.getForumID = getForumID_frm_disc;
SWEFForm.prototype.setForumID = setForumID_frm_disc;
SWEFForm.prototype.getParentID = getParentID_frm_disc;
SWEFForm.prototype.setParentID = setParentID_frm_disc;
SWEFForm.prototype.getThreadID = getThreadID_frm_disc;
SWEFForm.prototype.setThreadID = setThreadID_frm_disc;
SWEFForm.prototype.getEmailParentOnResponse = getEmailParentOnResponse_frm_disc;
SWEFForm.prototype.setEmailParentOnResponse = setEmailParentOnResponse_frm_disc;
SWEFForm.prototype.getSearchString = getSearchString_frm_disc;
SWEFForm.prototype.setSearchString = setSearchString_frm_disc;
SWEFForm.prototype.getSQLStatement = getSQLStatement_frm_disc;
SWEFForm.prototype.setSQLStatement = setSQLStatement_frm_disc;
SWEFForm.prototype.getFormName = getFormName_frm_disc;
SWEFForm.prototype.setFormName = setFormName_frm_disc;
SWEFForm.prototype.getMessageIDToDelete = getMessageIDToDelete_frm_disc;
SWEFForm.prototype.setMessageIDToDelete = setMessageIDToDelete_frm_disc;
SWEFForm.prototype.getSwitchEditableUserInfo = getSwitchEditableUserInfo_frm_disc;
SWEFForm.prototype.setSwitchEditableUserInfo = setSwitchEditableUserInfo_frm_disc;
SWEFForm.prototype.getAuthorName = getAuthorName_frm_disc;
SWEFForm.prototype.setAuthorName = setAuthorName_frm_disc;
SWEFForm.prototype.getAuthorEmail = getAuthorEmail_frm_disc;
SWEFForm.prototype.setAuthorEmail = setAuthorEmail_frm_disc;
SWEFForm.prototype.getAuthorFullname = getAuthorFullname_frm_disc;
SWEFForm.prototype.setAuthorFullname = setAuthorFullname_frm_disc;
SWEFForm.prototype.getArchiveDate = getArchiveDate_frm_disc;
SWEFForm.prototype.setArchiveDate = setArchiveDate_frm_disc;
SWEFForm.prototype.getPurgeCache = getPurgeCache_frm_disc;
SWEFForm.prototype.setPurgeCache = setPurgeCache_frm_disc;

SWEFForm.prototype.setAllFields = setAllFields_frm_disc;
SWEFForm.prototype.setAllFieldsFromForm = setAllFieldsFromForm_frm_disc;
SWEFForm.prototype.setAllFieldsFromMessage = setAllFieldsFromMessage_frm_disc;
SWEFForm.prototype.setAllFieldsEmpty = setAllFieldsEmpty_frm_disc;
SWEFForm.prototype.getSubjectInputField = getSubjectInputField_frm_disc;
SWEFForm.prototype.getBodyInputField = getBodyInputField_frm_disc;
SWEFForm.prototype.getSubjectSubform = getSubjectSubform_frm_disc;
SWEFForm.prototype.getBodySubform = getBodySubform_frm_disc;
SWEFForm.prototype.getEmailResponsesSubform = getEmailResponsesSubform_frm_disc;
SWEFForm.prototype.getNonEditableUserInfoSubform = getNonEditableUserInfoSubform_frm_disc;
SWEFForm.prototype.getEditableUserInfoSubform = getEditableUserInfoSubform_frm_disc;
SWEFForm.prototype.getDateSubform = getDateSubform_frm_disc;
SWEFForm.prototype.getMonthYearSubform = getMonthYearSubform_frm_disc;
SWEFForm.prototype.getUserInfoSubform = getUserInfoSubform_frm_disc;
SWEFForm.prototype.getMessageForm = getMessageForm_frm_disc;
SWEFForm.prototype.getSearchForm = getSearchForm_frm_disc;
SWEFForm.prototype.getSmallSearchForm = getSmallSearchForm_frm_disc;
SWEFForm.prototype.getArchiveSelector = getArchiveSelector_frm_disc;
SWEFForm.prototype.getAdminForm = getAdminForm_frm_disc;
SWEFForm.prototype.getDateElement = getDateElement_frm_disc;
SWEFForm.prototype.getMonthElement = getMonthElement_frm_disc;
SWEFForm.prototype.getYearElement = getYearElement_frm_disc;
SWEFForm.prototype.interpretDateSubform = interpretDateSubform_frm_disc;
SWEFForm.prototype.interpretMonthYearSubform = interpretMonthYearSubform_frm_disc;
</SCRIPT>

