<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
// Global variables
// ======================================================================

var currentSite_ID_disc = Session("Site_ID");
var currentSite_Code_desc;
var currentSite_Description_desc;

var currentForum_ID_disc = Session("Asset_ID");
var currentForum_Title_disc = Session("Forum_Title");
var currentForum_Moderated_disc = Session("Forum_Moderated");

var currentMessage_disc;

var currentUsername_disc = "Unset User";
var currentUserFullName_disc = "Unset Name";
var currentUserEmailAddress_disc = "Unset@Email.com";
var currentUser_ID_disc = "Unset User ID";
var currentUserLanguage_disc = "eng";

var isAdministrator_disc = Session("IsaAdministrator");

var config = new SWEFConfig ("");

var undefined_disc;

// ======================================================================
// CONSTANTS - well, pseudo-constants really...
// ======================================================================
var SWEF_FORM_OBJECT_TYPE_DISC = 100;
var SWEF_MESSAGE_OBJECT_TYPE_DISC = 101;
var SWEF_RECORDSET_OBJECT_TYPE_DISC = 102;

// ======================================================================
//
// The TRIVIAL interface.
//
// These functions are the simple objects users add to their HTML docs.
// They're a bit simplistic, but their trivially easy to use.
//
// None of them take parameters, and they directly output the object -
// they don't return a string.
//
// ======================================================================

function FORUM_LINK_DISC
(
)
{
  SWEFPageElement.showForumLink ();
  return;
}

function PARENT_MESSAGE_LINK_DISC
(
)
{
  SWEFPageElement.showParentMessageLink ();
  return;
}

function SEARCH_FORM_DISC
(
)
{
  SWEFPageElement.showSearchForm ();
  return;
}

function SMALL_SEARCH_FORM_DISC
(
)
{
  SWEFPageElement.showSmallSearchForm ();
  return;
}

function NEW_POST_BUTTON_DISC
(
)
{
  SWEFPageElement.showNewPostButton ();
  return;
}

function NEW_POST_LINK_DISC
(
)
{
  SWEFPageElement.showNewPostLink ();
  return;
}

function NEW_REPLY_BUTTON_DISC
(
)
{
  SWEFPageElement.showNewReplyButton ();
  return;
}

function EDIT_POST_BUTTON_DISC
(
)
{
  SWEFPageElement.showEditPostButton ();
  return;
}

function ALL_ROOT_POSTS_DISC
(
)
{
  SWEFPageElement.showAllThreads ();
  return;
}

function ALL_ROOT_POSTS_THREADED_DISC
(
)
{
  SWEFPageElement.showAllThreads ();
  return;
}

function CURRENT_POSTS_THREADED_DHTML_DISC
(
)
{
  SWEFPageElement.showCurrentThreads ();
  return;
}

function CURRENT_ROOT_POSTS_THREADED_DISC
(
)
{
  SWEFPageElement.showCurrentThreads ();
  return;
}

function CURRENT_ROOT_POSTS_THREADED_STATIC_DISC
(
)
{
  SWEFPageElement.showCurrentThreadsStatic ();
  return;
}

function NEW_POST_FORM_DISC
(
)
{
  SWEFPageElement.showNewPostForm ();
  return;
}

function EDIT_POST_FORM_DISC
(
)
{
  SWEFPageElement.showEditPostForm ();
  return;
}

function ARCHIVE_LINK_DISC
(
)
{
  SWEFPageElement.showArchiveLink ();
  return;
}

function ARCHIVE_SELECTOR_DISC
(
)
{
  SWEFPageElement.showArchiveSelector ();
  return;
}

function ARCHIVE_RESULTS_DISC
(
)
{
  SWEFPageElement.actionShowArchive ();
  return;
}

function ARCHIVE_DISC
(
)
{
  SWEFPageElement.showArchive ();
  return;
}

function SEARCH_RESULTS_DISC
(
)
{
  SWEFPageElement.actionSearch ();
  return;
}

function POST_MESSAGE_DISC
(
)
{
  SWEFPageElement.actionSaveNewMessage ();
  return;
}

function SAVE_EDITED_MESSAGE_DISC
(
)
{
  SWEFPageElement.actionSaveUpdatedMessage ();
  return;
}

function MESSAGE_DISC
(
)
{
  SWEFPageElement.showCurrentMessage ();
  return;
}

function THREAD_DISC
(
)
{
  SWEFPageElement.showCurrentThread ();
  return;
}

function ADMIN_TOOLS_DISC
(
)
{
  SWEFPageElement.showAdminSQLForm ();
  SWEFPageElement.actionExecuteAdminSQL ();
  return;
}

function STD_MESSAGE_DISC
(
)
{
  SWEFPageElement.showStandardMessage ();
  return;
}

function FORUM_TITLE_DISC
(
)
{
  SWEFPageElement.showForumTitle ();
  return;
}

function SUBJECT_DISC
(
)
{
  SWEFPageElement.showMessageSubject ();
  return;
}

function MESSAGE_BODY_DISC
(
)
{
  SWEFPageElement.showMessageBody ();
  return;
}

function AUTHOR_DISC
(
)
{
  SWEFPageElement.showMessageAuthorName ();
  return;
}

function AUTHOR_EMAIL_DISC
(
)
{
  SWEFPageElement.showMessageAuthorEmail ();
  return;
}

function AUTHOR_FULL_NAME_DISC
(
)
{
  SWEFPageElement.showMessageAuthorFullname ();
  return;
}

function DATE_CREATED_DISC
(
)
{
  SWEFPageElement.showMessageDateCreated ();
  return;
}

function DATE_MODIFIED_DISC
(
)
{
  SWEFPageElement.showMessageDateModified ();
  return;
}

function SORT_CODE_DISC
(
)
{
  SWEFPageElement.showMessageSortCode ();
  return;
}

function NUM_CHILDREN_DISC
(
)
{
  SWEFPageElement.showMessageNumChildren ();
  return;
}

function MESSAGE_ID_DISC
(
)
{
  SWEFPageElement.showMessageID ();
  return;
}

function SITE_ID_DISC
(
)
{
  SWEFPageElement.showMessageSiteID ();
  return;
}

function FORUM_ID_DISC
(
)
{
  SWEFPageElement.showMessageForumID ();
  return;
}

function PARENT_ID_DISC
(
)
{
  SWEFPageElement.showMessageParentID ();
  return;
}

function THREAD_ID_DISC
(
)
{
  SWEFPageElement.showMessageThreadID ();
  return;
}

function CLEANUP_DISC
(
)
{
  var dbDatabase = new SWEFDatabase ();
  dbDatabase.cleanup ();

  if (config.CACHE_PurgeOnPageCleanup)
    {
      var cchRecordCache = new SWEFCache ();
      cchRecordCache.purge ();

      delete cchRecordCache;
    }

  delete dbDatabase;
  delete config;
  delete currentMessage_disc;
  return;
}

//=======================================================================
//
// Some global helper functions.  Some of these should have been built into
// the language if you ask me...
//
//=======================================================================

function isUndefined_disc
(
 JSvariable
)
{
  var bIsUndefined;

  if ("object" == typeof (JSvariable))
    {
      bIsUndefined = false;
    }
  else
    {
      bIsUndefined = (String (JSvariable).valueOf () == "undefined" ? true : false);
    }

  return bIsUndefined;
}

function isDefined_disc
(
 JSvariable
)
{
  return !isUndefined_disc (JSvariable);
}

function safeStringDereference_disc
(
 JSvariable
)
{
  return (isUndefined_disc (String (JSvariable)) ? "" : String (JSvariable));
}

function safeNumberDereference_disc
(
 JSvariable
)
{
  return (isUndefined_disc (JSvariable) ? "0" : JSvariable);
}

function isCurrentPage_disc
(
 strURLToCheck
)
{
  var bURLIsCurrentPage = false;
  var strCurrentPage = String (Request.ServerVariables ("PATH_INFO"));

  var strURLToCheckUC = strURLToCheck.toUpperCase ();
  var strCurrentPageUC = strCurrentPage.toUpperCase ();

  if ((strURLToCheckUC == strCurrentPageUC) || ((strURLToCheckUC + "DEFAULT.ASP") == strCurrentPageUC))
    {
      bURLIsCurrentPage = true;
    }

  return bURLIsCurrentPage;
}

</SCRIPT>

<!-- #INCLUDE file="cacheobj.asp"-->
<!-- #INCLUDE file="configobj.asp"-->
<!-- #INCLUDE file="databaseobj.asp"-->
<!-- #INCLUDE file="dateobj.asp"-->
<!-- #INCLUDE file="errorobj.asp"-->
<!-- #INCLUDE file="emailobj.asp"-->
<!-- #INCLUDE file="expandcollapseobj.asp"-->
<!-- #INCLUDE file="formobj.asp"-->
<!-- #INCLUDE file="htmlobj.asp"-->
<!-- #INCLUDE file="messageobj.asp"-->
<!-- #INCLUDE file="pageobj.asp"-->
<!-- #INCLUDE file="stringobj.asp"-->
<!-- #INCLUDE file="textcontrolobj.asp"-->
<!-- #INCLUDE file="threadobj.asp"-->
<!-- #INCLUDE file="viewobj.asp"-->