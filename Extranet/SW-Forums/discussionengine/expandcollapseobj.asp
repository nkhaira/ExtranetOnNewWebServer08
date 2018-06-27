<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// EXPANDCOLLAPSE OBJECT
//
// ======================================================================

function SWEFExpandCollapse
(
 msgMessageToUse,
 nViewIndex,
 nCurrentThreadID,
 nActiveThreadID,
 strPageURLToUse
)
{
  this.setMessage (msgMessageToUse);
  this.setViewIndex (nViewIndex);
  this.setCurrentThreadID (nCurrentThreadID);
  this.setActiveThreadID (nActiveThreadID);
  this.setQueryString ("");
  if (isUndefined_disc (strPageURLToUse))
    {
      this.setPageURLToUse (config.getMainPagePath ());
    }
  else
    {
      this.setPageURLToUse (strPageURLToUse);
    }

  return this;
}

// ======================================================================
//
// Interface to private member variables.
//
// ======================================================================

function getMessage_exp_disc
(
)
{
  return this._message;
}

function setMessage_exp_disc
(
 msgNewMessage
)
{
  this._message = msgNewMessage;
  return;
}

function getViewIndex_exp_disc
(
)
{
  return Number (this._viewIndex);
}

function setViewIndex_exp_disc
(
 nNewViewIndex
)
{
  this._viewIndex = Number (nNewViewIndex);
  return;
}

function getCurrentThreadID_exp_disc
(
)
{
  return Number (this._currentThreadID);
}

function setCurrentThreadID_exp_disc
(
 nNewCurrentThreadID
)
{
  this._currentThreadID = Number (nNewCurrentThreadID);
  return;
}

function getActiveThreadID_exp_disc
(
)
{
  return Number (this._activeThreadID);
}

function setActiveThreadID_exp_disc
(
 nNewActiveThreadID
)
{
  this._activeThreadID = Number (nNewActiveThreadID);
  return;
}

function getQueryString_exp_disc
(
)
{
  return String (this._QueryString);
}

function setQueryString_exp_disc
(
 strNewQueryString
)
{
  this._QueryString = String (strNewQueryString);
  return;
}

function getPageURLToUse_exp_disc
(
)
{
  return String (this._pageURLToUse);
}

function setPageURLToUse_exp_disc
(
 strNewPageURLToUse
)
{
  this._pageURLToUse = String (strNewPageURLToUse);
  return;
}

// ======================================================================
//
// Main object methods.
//
// ======================================================================

function getIcon_exp_disc
(
)
{
  var strToggle = "";
  if (config.ADMINSWITCH_ShowExpandCollapse)
    {
      if (0 == this.getMessage ().getNumChildren ())
	{
	  strToggle = this.getNoExpandToggle();
	}
      else
	{
	  if (((this.getCurrentThreadID () == this.getActiveThreadID ())
	       || (this.getViewIndex () <= config.ADMINSETTING_ExpandFirstNThreads))
	      && !(this.doWeCollapse()))
	    {
	      strToggle = this.getCollapseToggle();
	    }
	  else
	    {
	      strToggle = this.getExpandToggle();
	    }
	}
    }

  return strToggle;
}

function getDHTMLIcon_exp_disc
(
 nCurrentIndentLevel,
 nThreadCounter
)
{
  var strToggle = "";;

  if (0 == this.getMessage ().getNumChildren ())
    {
      strToggle += SWEFHTML.IMG (config.ADMINSETTING_NoExpandImagePathname,
				" ",
				0,
				config.ADMINSETTING_NoExpandImageWidth,
				config.ADMINSETTING_NoExpandImageHeight,
				config.getUniqueKey () + this.getMessage ().getMessageID (),
				this.getMessage ().getMessageID (),
				config.getDHTMLEventHandler ());
      strToggle += SWEFHTML.NBSP ();
    }
  else
    {
      //if ((nCurrentIndentLevel == 1) && !config.ADMINSWITCH_ExpandAllThreads);
      if (config.ADMINSWITCH_ExpandAllThreads
	  || (nThreadCounter <= config.ADMINSETTING_ExpandFirstNThreads)
	  || (nCurrentIndentLevel != 1))

	{
	  strToggle += this.getCollapseLinkTag ();
	  strToggle += SWEFHTML.IMG (config.ADMINSETTING_CollapseImagePathname,
				     config.USERTEXT_VIEW_PopupCollapseLink,
				     0,
				     config.ADMINSETTING_CollapseImageWidth,
				     config.ADMINSETTING_CollapseImageHeight,
				     config.getUniqueKey () + this.getMessage ().getMessageID (),
				     this.getMessage ().getMessageID (),
				     config.getDHTMLEventHandler ());
	  strToggle += SWEFHTML.A_close ();
	  strToggle += SWEFHTML.NBSP ();
	}
      else
 	{
	  strToggle += this.getExpandLinkTag ();
	  strToggle += SWEFHTML.IMG (config.ADMINSETTING_ExpandImagePathname,
				     config.USERTEXT_VIEW_PopupExpandLink,
				     0,
				     config.ADMINSETTING_ExpandImageWidth,
				     config.ADMINSETTING_ExpandImageHeight,
				     config.getUniqueKey () + this.getMessage ().getMessageID (),
				     this.getMessage ().getMessageID (),
				     config.getDHTMLEventHandler ());
	  strToggle += SWEFHTML.A_close ();
	  strToggle += SWEFHTML.NBSP ();
	}
   }

  return strToggle;
}

function getNoExpandToggle_exp_disc
(
)
{
  var strToggle = "";
  strToggle += SWEFHTML.IMG (config.ADMINSETTING_NoExpandImagePathname,
			     " ",
			     0,
			     config.ADMINSETTING_NoExpandImageWidth,
			     config.ADMINSETTING_NoExpandImageHeight,
			     config.getUniqueKey () + this.getMessage ().getMessageID (),
			     this.getMessage ().getMessageID (),
			     "");
  strToggle += SWEFHTML.NBSP ();

  return strToggle;
}

function getEmptyNoExpandToggle_exp_disc
(
)
{
  var strToggle = "";
  strToggle += SWEFHTML.IMG (config.ADMINSETTING_NoExpandImagePathname,
			     " ",
			     0,
			     config.ADMINSETTING_NoExpandImageWidth,
			     config.ADMINSETTING_NoExpandImageHeight,
			     undefined_disc,
			     undefined_disc,
			     "");
  strToggle += SWEFHTML.NBSP ();

  return strToggle;
}

function getExpandLinkTag_exp_disc
(
)
{
  var strLinkURL = "";
  strLinkURL += this.getExpandURL ();
  strLinkURL += "#";
  strLinkURL += config.FORM_QueryStringViewCentre;
  strLinkURL += "-";
  strLinkURL += this.getMessage ().getMessageID ();

  var strLinkName = "";
  strLinkName += config.FORM_QueryStringViewCentre;
  strLinkName += "-";
  strLinkName += this.getMessage ().getMessageID ();

  var strLinkID = "";
  strLinkID += config.FORM_QueryStringViewCentre;
  strLinkID += "-";
  strLinkID += this.getMessage ().getMessageID ();

  var strToggle = "";
  strToggle += SWEFHTML.A_open (strLinkURL,
				config.USERTEXT_VIEW_PopupExpandLink,
				strLinkName,
				config.ADMINSETTING_MainPageTarget,
				strLinkID);
  return strToggle;
}

function getExpandToggle_exp_disc
(
)
{
  var strImage = "";
  strImage += SWEFHTML.IMG (config.ADMINSETTING_ExpandImagePathname,
			    config.USERTEXT_VIEW_PopupExpandLink,
			    0,
			    config.ADMINSETTING_ExpandImageWidth,
			    config.ADMINSETTING_ExpandImageHeight,
			    this.getMessage ().getMessageID (),
			    this.getMessage ().getMessageID (),
			    "");

  return this.getExpandLinkTag () + strImage + SWEFHTML.A_close () + SWEFHTML.NBSP ();
}

function getCollapseLinkTag_exp_disc
(
)
{
  var strLinkURL = "";
  strLinkURL += this.getCollapseURL ();
  strLinkURL += "#";
  strLinkURL += config.FORM_QueryStringViewCentre;
  strLinkURL += "-";
  strLinkURL += this.getMessage ().getMessageID ();

  var strLinkName = "";
  strLinkName += config.FORM_QueryStringViewCentre;
  strLinkName += "-";
  strLinkName += this.getMessage ().getMessageID ();

  var strLinkID = "";
  strLinkID += config.FORM_QueryStringViewCentre;
  strLinkID += "-";
  strLinkID += this.getMessage ().getMessageID ();

  return SWEFHTML.A_open (strLinkURL,
			  config.USERTEXT_VIEW_PopupCollapseLink,
			  strLinkName,
			  config.ADMINSETTING_MainPageTarget,
			  strLinkID);
}

function getCollapseToggle_exp_disc
(
)
{
  var strImage = "";
  strImage += SWEFHTML.IMG (config.ADMINSETTING_CollapseImagePathname,
			    config.USERTEXT_VIEW_PopupCollapseLink,
			    0,
			    config.ADMINSETTING_CollapseImageWidth,
			    config.ADMINSETTING_CollapseImageHeight,
			    this.getMessage ().getMessageID (),
			    this.getMessage ().getMessageID (),
			    "");

  return this.getCollapseLinkTag () + strImage + SWEFHTML.A_close () + SWEFHTML.NBSP ();
}

function getExpandURL_exp_disc
(
)
{
  var strURL = this.getPageURLToUse ();
  var strURLToUse = "";

  if (this.stripThisIDFromQString ())
    {
      strURLToUse += (this.getQueryString () == ""
		      ? config.getMainPagePreferredPath ()
		      : strURL + "?" + this.getQueryString ());
    }
  else
    {
      strURLToUse += strURL;
      strURLToUse += "?";
      strURLToUse += config.FORM_QueryStringViewExpand;
      strURLToUse += "=";
      strURLToUse += this.getMessage ().getMessageID ();
      strURLToUse += (this.getQueryString () == "" ? "" : "&" + this.getQueryString ());
    }

  return strURLToUse;
}

function getCollapseURL_exp_disc
(
)
{
  var strURL = this.getPageURLToUse ();
  var strURLToUse = "";

  if (this.stripThisIDFromQString())
    {
      strURLToUse += (this.getQueryString () == ""
		      ? config.getMainPagePreferredPath ()
		      : strURL + "?" + this.getQueryString ());
    }
  else
    {
      strURLToUse += strURL;
      strURLToUse += "?";
      strURLToUse += config.FORM_QueryStringViewCollapse;
      strURLToUse += "=";
      strURLToUse += this.getMessage ().getMessageID ();
      strURLToUse += (this.getQueryString () == "" ? "" : "&" + this.getQueryString ());
    }

  return strURLToUse;
}

function stripThisIDFromQString_exp_disc
(
)
{
  var bIDRemoved = false;
  var strQueryString = "";
  var nItemCount;
  var nNumItems;

  nNumItems = Request.QueryString (config.FORM_QueryStringViewExpand).count;
  for (nItemCount = 1; nItemCount <= nNumItems; nItemCount++)
    {
      if (Request.QueryString (config.FORM_QueryStringViewExpand)(nItemCount)
	  == this.getMessage ().getMessageID ())
	{
	  bIDRemoved = true;
	}
      else
	{
	  strQueryString += config.FORM_QueryStringViewExpand;
	  strQueryString += "=";
	  strQueryString += Server.URLEncode (Request.QueryString (config.FORM_QueryStringViewExpand)(nItemCount));
	  strQueryString += "&";
	}
    }

  nNumItems = Request.QueryString (config.FORM_QueryStringViewCollapse).count;
  for (nItemCount = 1; nItemCount <= nNumItems; nItemCount++)
    {
      if (Request.QueryString (config.FORM_QueryStringViewCollapse)(nItemCount)
	  == this.getMessage ().getMessageID ())
	{
	  bIDRemoved = true;
	}
      else
	{
	  strQueryString += config.FORM_QueryStringViewCollapse;
	  strQueryString += "=";
	  strQueryString += Server.URLEncode (Request.QueryString (config.FORM_QueryStringViewCollapse)(nItemCount));
	  strQueryString += "&";
	}
    }

  nNumItems = Request.QueryString.count;
  for (nItemCount = 1; nItemCount <= nNumItems; nItemCount++)
    {
      if ((Request.QueryString.Key (nItemCount) != config.FORM_QueryStringViewCollapse)
	  && (Request.QueryString.Key (nItemCount) != config.FORM_QueryStringViewExpand)
	  && (Request.QueryString.Key (nItemCount) != config.FORM_QueryStringViewCentre))
	{
	  strQueryString += Request.QueryString.Key (nItemCount);
	  strQueryString += "=";
	  strQueryString += Request.QueryString (nItemCount);
	  strQueryString += "&";
	}
    }

  this.setQueryString (strQueryString.substr(0, strQueryString.length - 1));

  return bIDRemoved;
}

function doWeExpand_exp_disc
(
)
{
  var bDoExpand = false;

  if (config.ADMINSWITCH_ExpandAllThreads)
    {
      bDoExpand = true;
    }
  else
    {
      if (this.getViewIndex () <= config.ADMINSETTING_ExpandFirstNThreads)
	{
	  bDoExpand = true;
	}
      else
	{
	  var nNumItems;
	  nNumItems = Request.QueryString (config.FORM_QueryStringViewExpand).count;
	  var bExpand = false;
	  for (var nItemCount = 1; nItemCount <= nNumItems; nItemCount++)
	    {
	      if (Request.QueryString (config.FORM_QueryStringViewExpand)(nItemCount)
		  == this.getMessage ().getMessageID ())
		{
		  bExpand = true;
		}
	    }

	  bDoExpand = bExpand;
	}
    }

  return bDoExpand;
}

function doWeCollapse_exp_disc
(
)
{
  var bCollapse = false;
  var nNumItems = Request.QueryString (config.FORM_QueryStringViewCollapse).count;
  for (var nItemCount = 1; nItemCount <= nNumItems; nItemCount++)
    {
      if (Request.QueryString (config.FORM_QueryStringViewCollapse)(nItemCount)
	  == this.getMessage ().getMessageID ())
	{
	  bCollapse = true;
	}
    }

  return bCollapse;
}

SWEFExpandCollapse.prototype.getMessage = getMessage_exp_disc;
SWEFExpandCollapse.prototype.setMessage = setMessage_exp_disc;
SWEFExpandCollapse.prototype.getViewIndex = getViewIndex_exp_disc;
SWEFExpandCollapse.prototype.setViewIndex = setViewIndex_exp_disc;
SWEFExpandCollapse.prototype.getCurrentThreadID = getCurrentThreadID_exp_disc;
SWEFExpandCollapse.prototype.setCurrentThreadID = setCurrentThreadID_exp_disc;
SWEFExpandCollapse.prototype.getActiveThreadID = getActiveThreadID_exp_disc;
SWEFExpandCollapse.prototype.setActiveThreadID = setActiveThreadID_exp_disc;
SWEFExpandCollapse.prototype.getQueryString = getQueryString_exp_disc;
SWEFExpandCollapse.prototype.setQueryString = setQueryString_exp_disc;
SWEFExpandCollapse.prototype.getPageURLToUse = getPageURLToUse_exp_disc;
SWEFExpandCollapse.prototype.setPageURLToUse = setPageURLToUse_exp_disc;

SWEFExpandCollapse.prototype.getIcon = getIcon_exp_disc;
SWEFExpandCollapse.prototype.getDHTMLIcon = getDHTMLIcon_exp_disc;
SWEFExpandCollapse.prototype.getNoExpandToggle = getNoExpandToggle_exp_disc;
SWEFExpandCollapse.prototype.getEmptyNoExpandToggle = getEmptyNoExpandToggle_exp_disc;
SWEFExpandCollapse.prototype.getExpandLinkTag = getExpandLinkTag_exp_disc;
SWEFExpandCollapse.prototype.getExpandToggle = getExpandToggle_exp_disc;
SWEFExpandCollapse.prototype.getCollapseLinkTag = getCollapseLinkTag_exp_disc;
SWEFExpandCollapse.prototype.getCollapseToggle = getCollapseToggle_exp_disc;
SWEFExpandCollapse.prototype.getExpandURL = getExpandURL_exp_disc;
SWEFExpandCollapse.prototype.getCollapseURL = getCollapseURL_exp_disc;
SWEFExpandCollapse.prototype.stripThisIDFromQString = stripThisIDFromQString_exp_disc;
SWEFExpandCollapse.prototype.doWeExpand = doWeExpand_exp_disc;
SWEFExpandCollapse.prototype.doWeCollapse = doWeCollapse_exp_disc;

SWEFExpandCollapse.getEmptyNoExpandToggle = getEmptyNoExpandToggle_exp_disc;
</SCRIPT>

