<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// VIEW OBJECT
//
// ======================================================================

function SWEFView
(
 rsMessages,
 nCurrentMessageID,
 strPageURLToUse
)
{
  this.setMessageList (rsMessages);
  this.setCurrentMessageID (nCurrentMessageID);
  this.setThreadCounter (0);
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

function getMessageList_vw_disc
(
)
{
  return this._messages;
}

function setMessageList_vw_disc
(
 rsNewMessageList
)
{
  this._messages = rsNewMessageList;
  return;
}

function getCurrentMessageID_vw_disc
(
)
{
  return Number (this._currentMessageID);
}

function setCurrentMessageID_vw_disc
(
 nNewCurrentMessageID
)
{
  this._currentMessageID = Number (nNewCurrentMessageID);
  return;
}

function getPageURLToUse_vw_disc
(
)
{
  return String (this._pageURLToUse);
}

function setPageURLToUse_vw_disc
(
 strNewPageURLToUse
)
{
  this._pageURLToUse = String (strNewPageURLToUse);
  return;
}

function getThreadCounter_vw_disc
(
)
{
  return Number (this._threadCounter);
}

function setThreadCounter_vw_disc
(
 nThreadCounter
)
{
  this._threadCounter = Number (nThreadCounter);
  return;
}

// ======================================================================
//
// Main object methods.
//
// ======================================================================

function getView_vw_disc
(
)
{
  var strView = "";

  if (!this.getMessageList ().EOF)
    {
      var msgCurrentMessage;
      var strViewLine;

      this.getMessageList ().MoveFirst();

      var nShowMaxThreads = -1;
      if (config.ADMINSETTING_ShowAtMostNThreads > 0)
	{
	  nShowMaxThreads = config.ADMINSETTING_ShowAtMostNThreads;
	}

      while ((!this.getMessageList ().EOF)
	     && ((nShowMaxThreads == -1) || (this.getThreadCounter () < nShowMaxThreads)))
	{
	  this.setThreadCounter (this.getThreadCounter () + 1);

	  msgCurrentMessage = new SWEFMessage (this.getMessageList (),
					       SWEF_RECORDSET_OBJECT_TYPE_DISC);
	  var exExpandCollapse = new SWEFExpandCollapse (msgCurrentMessage,
							 this.getThreadCounter (),
							 msgCurrentMessage.getMessageID (),
							 0,
							 this.getPageURLToUse ());
	  if (exExpandCollapse.doWeExpand())
	    {
	      var thdExpandedThread = new SWEFThread (this.getCurrentMessageID (),
						      this.getPageURLToUse ());
	      strView += thdExpandedThread.getExpandedThread (msgCurrentMessage.getThreadID (),
							      msgCurrentMessage.getSortCode ());
	      delete thdExpandedThread;
	    }
	  else
	    {
	      strViewLine = msgCurrentMessage.getExpandCollapse (this.getThreadCounter (),
							      1,
							      this.getPageURLToUse ());
	      strViewLine += msgCurrentMessage.getViewEntry (this.getCurrentMessageID ());
	      strView += this.getIndented (1, strViewLine);
	    }

	  delete msgCurrentMessage;
	  delete exExpandCollapse;
	  this.getMessageList ().MoveNext();
	}
    }
  else
    {
      strView = this.getIndented (1,
				  SWEFExpandCollapse.getEmptyNoExpandToggle ()
				  + config.USERTEXT_VIEW_NoMessages.weak());
    }

  return strView;
}

function getDHTMLView_vw_disc
(
)
{
  var strView = "";
  if (!this.getMessageList ().EOF)
    {
      this.getMessageList ().MoveFirst();

      strView += this.getDHTMLExpandCollapseCode ();

      var nShowMaxThreads = -1;
      if (config.ADMINSETTING_ShowAtMostNThreads > 0)
	{
	  nShowMaxThreads = config.ADMINSETTING_ShowAtMostNThreads;
	}

      while ((!this.getMessageList ().EOF)
	     && ((nShowMaxThreads == -1) || (this.getThreadCounter () < nShowMaxThreads)))
	{
	  this.setThreadCounter (this.getThreadCounter () + 1);
	  strView += this.getDHTMLThread (this.getMessageList ());
	}
    }
  else
    {
      strView += this.getIndented (1,
				   SWEFExpandCollapse.getEmptyNoExpandToggle ()
				   + config.USERTEXT_VIEW_NoMessages.weak());
    }

  return strView;
}

function getDHTMLThread_vw_disc
(
 rsMessages
)
{
  var nRelativeIndentLevel;
  var nAbsoluteIndentLevel;
  var strSortCode;
  var msgCurrentMessage;
  var strSubjectLine;

  nRelativeIndentLevel = 1;
  msgCurrentMessage = new SWEFMessage (rsMessages, SWEF_RECORDSET_OBJECT_TYPE_DISC);
  nAbsoluteIndentLevel = this.getSortcodeIndentLevel (msgCurrentMessage.getSortCode ());
  strSortCode = msgCurrentMessage.getSortCode ();
  strSubjectLine = msgCurrentMessage.getDHTMLExpandCollapse (nRelativeIndentLevel,
							     this.getPageURLToUse (),
							     this.getThreadCounter ());
  strSubjectLine += msgCurrentMessage.getViewEntry (0);

  var nCurrentID = String (msgCurrentMessage.getMessageID ());
  var strCurrentSortCode = String (msgCurrentMessage.getSortCode ());

  delete msgCurrentMessage;

  var strDHTMLThread = "";
  strDHTMLThread += this.getIndented (nRelativeIndentLevel, strSubjectLine);
  strDHTMLThread += this.getDHTMLChildren (nCurrentID,
					   strCurrentSortCode,
					   nAbsoluteIndentLevel,
					   nRelativeIndentLevel,
					   strSortCode,
					   rsMessages);
  return strDHTMLThread;
}

function getDHTMLChildren_vw_disc
(
 nPreviousID,
 strPreviousSortCode,
 nAbsoluteIndentLevel,
 nRelativeIndentLevel,
 strWorkingSortCode,
 rsMessages
)
{
  var strViewLine = "";
  var strDisplaySubthread = "";

  rsMessages.MoveNext ();
  var msgNextMessage = new SWEFMessage (rsMessages, SWEF_RECORDSET_OBJECT_TYPE_DISC);
  if (this.isChild (strPreviousSortCode, msgNextMessage.getSortCode ()))
    {
      nAbsoluteIndentLevel++;
      nRelativeIndentLevel++;

      if (config.ADMINSWITCH_ExpandAllThreads
	  || (this.getThreadCounter () <= config.ADMINSETTING_ExpandFirstNThreads)
	  || (nRelativeIndentLevel != 2))
	{
	  strDisplaySubthread = "";
	}
      else
	{
	  strDisplaySubthread = "none";
	}
      strViewLine += SWEFHTML.DIV_open ("Subthread-" + config.getUniqueKey () + nPreviousID,
					"margin-left:0;display:" + strDisplaySubthread + ";");

      var nMessageOutputCounter = 0;
      while (!rsMessages.EOF && (this.isChild (strPreviousSortCode, msgNextMessage.getSortCode ())))
	{
	  var strSubjectLine = msgNextMessage.getDHTMLExpandCollapse (nRelativeIndentLevel,
								   this.getPageURLToUse (),
								   this.getThreadCounter ());
	  strSubjectLine += msgNextMessage.getViewEntry (0);
	  strViewLine += this.getIndented (nRelativeIndentLevel, strSubjectLine);
	  strViewLine += this.getDHTMLChildren (String (msgNextMessage.getMessageID ()),
						msgNextMessage.getSortCode (),
						nAbsoluteIndentLevel,
						nRelativeIndentLevel,
						msgNextMessage.getSortCode (),
						rsMessages);
	  delete msgNextMessage;
	  msgNextMessage = new SWEFMessage (rsMessages, SWEF_RECORDSET_OBJECT_TYPE_DISC);
	  nMessageOutputCounter++;
	}

      strViewLine += SWEFHTML.DIV_close ();
    }

  delete msgNextMessage;
  return strViewLine;
}

function getViewThread_vw_disc
(
)
{
  var strThread = "";

  if (!this.getMessageList ().EOF)
    {
      var nLastIndentLevel = 0;
      var nCurrentIndentLevel;
      var strViewLine;
      var strFirstMessageSortCode;

      this.getMessageList ().MoveFirst();

      var dbDatabase = new SWEFDatabase ();
      strFirstMessageSortCode = dbDatabase.getSortCodeField (this.getMessageList ());

      while (!this.getMessageList ().EOF)
	{
	  this.setThreadCounter (this.getThreadCounter () + 1);

	  var msgCurrentMessage = new SWEFMessage (this.getMessageList (),
						   SWEF_RECORDSET_OBJECT_TYPE_DISC);
	  var exExpandCollapseThread = new SWEFExpandCollapse (msgCurrentMessage,
							       this.getThreadCounter (),
							       0,
							       0,
							       this.getPageURLToUse ());

	  nCurrentIndentLevel = 1 + this.getIndentLevel () - this.getSortcodeIndentLevel (strFirstMessageSortCode);
	  if (nCurrentIndentLevel < 1)
	    {
	      nCurrentIndentLevel = 1;
	      strFirstMessageSortCode = dbDatabase.getSortCodeField (this.getMessageList ());
	}

	  strViewLine = msgCurrentMessage.getExpandCollapse (this.getThreadCounter (),
							     0,
							     this.getPageURLToUse ());
	  strViewLine += msgCurrentMessage.getViewEntry (this.getCurrentMessageID ());
	  strThread += this.getIndented (nCurrentIndentLevel, strViewLine);

	  if (exExpandCollapseThread.doWeCollapse())
	    {
	      this.skipToNextSibling ();
	      if (!this.getMessageList ().EOF)
		{
		  strFirstMessageSortCode = this.getParentSortCode (String (dbDatabase.getSortCodeField (this.getMessageList ())));
		}
	    }
	  else
	    {
	      if (!this.getMessageList ().EOF)
		{
		  this.getMessageList ().MoveNext();
		}
	    }

	  delete msgCurrentMessage;
	  delete exExpandCollapseThread;
	}
    }
  else
    {
      strThread = this.getIndented (1, SWEFExpandCollapse.getEmptyNoExpandToggle () + config.USERTEXT_VIEW_NoMessages.weak());
    }

  delete dbDatabase;
  return strThread;
}

function getParentSortCode_vw_disc
(
 strSortcode
)
{
  var strParentSortCode = "";
  var nLastSeparatorIndex = strSortcode.lastIndexOf (".");
  if (-1 != nLastSeparatorIndex)
    {
      strParentSortCode = strSortcode.substring (0, nLastSeparatorIndex);
    }

  return strParentSortCode;
}

function getViewFullThread_vw_disc
(
)
{
  var strFullThread = "";

  if (!this.getMessageList ().EOF)
    {
      var nLastIndentLevel;
      var nCurrentIndentLevel;
      var strViewLine;
      var strFirstMessageSortCode;

      this.getMessageList ().MoveFirst();

      nLastIndentLevel = 0;
      var dbDatabase = new SWEFDatabase ();
      strFirstMessageSortCode = String (dbDatabase.getSortCodeField (this.getMessageList ()));

      while (!this.getMessageList ().EOF)
	{
	  var msgCurrentMessage = new SWEFMessage (this.getMessageList (),
						   SWEF_RECORDSET_OBJECT_TYPE_DISC);

	  nCurrentIndentLevel = this.getIndentLevel ();
	  strViewLine = msgCurrentMessage.getViewEntry (this.getCurrentMessageID ());

	  strFullThread += this.getIndented (nCurrentIndentLevel, strViewLine);

	  if (!this.getMessageList ().EOF)
	    {
	      this.getMessageList ().MoveNext();
	    }

	  delete msgCurrentMessage;
	}
    }
  else
    {
      strFullThread = this.getIndented (1, SWEFExpandCollapse.getEmptyNoExpandToggle () + config.USERTEXT_VIEW_NoMessages.weak());
    }

  delete dbDatabase;
  return strFullThread;
}

function getIndented_vw_disc
(
 nCurrentIndentLevel,
 strViewText
)
{
  var strHTMLout = "";

  strHTMLout += SWEFHTML.TABLE_open (config.ADMINSETTING_TableBorderSize,
				     undefined_disc,
				     undefined_disc,
				     undefined_disc,
				     0,
				     0);
  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "top");
  strHTMLout += SWEFHTML.NBSP ();

  for (var nCounter = 0; nCounter < nCurrentIndentLevel; nCounter++)
    {
      strHTMLout += SWEFHTML.NBSP (config.ADMINSETTING_ViewIndentResponseSpaces);
    }
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  "top");
  strHTMLout += strViewText;
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();
  strHTMLout += SWEFHTML.TABLE_close ();

  return strHTMLout;
}

function getIndentLevel_vw_disc
(
)
{
  var msgCurrentMessage = new SWEFMessage (this.getMessageList (),
					   SWEF_RECORDSET_OBJECT_TYPE_DISC);
  var strParsedCode = msgCurrentMessage.getSortCode ();
  var nFoundAt = 1;
  var nIndentCounter = 0;

  while (nFoundAt > 0)
    {
      nFoundAt = String (strParsedCode).search ("[\\.]");
      strParsedCode = String (strParsedCode).substr (nFoundAt + 1);
      nIndentCounter++;
    }

  delete msgCurrentMessage;
  return nIndentCounter;
}

function getSortcodeIndentLevel_vw_disc
(
 strSortCode
)
{
  var strParsedCode = strSortCode;
  var nFoundAt = 1;
  var nIndentCounter = 0;

  while (nFoundAt > 0)
    {
      nFoundAt = String (strParsedCode).search ("[^0-9]");
      strParsedCode = String (strParsedCode).substr (nFoundAt + 1);
      nIndentCounter++;
    }
  return nIndentCounter;
}

function isChild_vw_disc
(
 strSortCode1,
 strSortCode2
)
{
  var bIsChild = false;
  var strPossibleParentSortCode = strSortCode1 + ".";
  var strPossibleChildSortCode = strSortCode2;
  if (strPossibleParentSortCode == strPossibleChildSortCode.substring (0, strPossibleParentSortCode.length))
    {
      bIsChild = true;
    }

  return bIsChild;
}

function skipToNextSibling_vw_disc
(
)
{
  var msgCurrentMessage = new SWEFMessage (this.getMessageList (),
					   SWEF_RECORDSET_OBJECT_TYPE_DISC);
  var strCurrentSortCode = msgCurrentMessage.getSortCode ();

  this.getMessageList ().MoveNext();
  while ((!this.getMessageList ().EOF) && (String (strCurrentSortCode).length < String (new SWEFMessage(this.getMessageList (), SWEF_RECORDSET_OBJECT_TYPE_DISC).getSortCode ()).length))
    {
      this.getMessageList ().MoveNext();
    }

  delete msgCurrentMessage;
  return;
}

function getDHTMLExpandCollapseCode_vw_disc ()
{
  var strDHTMLCode = "";

  strDHTMLCode += "\n<SCR" + "IPT LANGUAGE=\"Javascript\" RUNAT=\"client\">";
  strDHTMLCode += "\nfunction " + config.getDHTMLFunctionName () + " (DOMEvent)";
  strDHTMLCode += "\n{";
  strDHTMLCode += "\n\tvar customGetByID;";
  strDHTMLCode += "\n\tvar parentID;";
  strDHTMLCode += "\n\tvar child;";
  strDHTMLCode += "\n\tif (document.all)";
  strDHTMLCode += "\n\t{";
  strDHTMLCode += "\n\t\tdocument.getElementById = document.all;";
  strDHTMLCode += "\n\t\tparentID = window.event.srcElement.id;";
  strDHTMLCode += "\n\t\tchild = document.getElementById (\"Subthread-\" + parentID);";
  strDHTMLCode += "\n\t}";
  strDHTMLCode += "\n\telse";
  strDHTMLCode += "\n\t{";
  strDHTMLCode += "\n\t\tparentID = DOMEvent.target.id;";
  strDHTMLCode += "\n\t\tchild = document.getElementById (\"Subthread-\" + parentID);";
  strDHTMLCode += "\n\t}";
  strDHTMLCode += "\n\tif (child != null)";
  strDHTMLCode += "\n\t{";
  strDHTMLCode += "\n\t\tvar parentImage = document.getElementById (parentID);";
  strDHTMLCode += "\n\t\tif (child.style.display == \"none\")";
  strDHTMLCode += "\n\t\t{";
  strDHTMLCode += "\n\t\t\tchild.style.display = \"block\";";
  strDHTMLCode += "\n\t\t\tparentImage.src = \"" + config.ADMINSETTING_CollapseImagePathname + "\";";
  strDHTMLCode += "\n\t\t\tparentImage.alt = \"" + config.USERTEXT_VIEW_PopupCollapseLink + "\";";
  strDHTMLCode += "\n\t\t}";
  strDHTMLCode += "\n\t\telse";
  strDHTMLCode += "\n\t\t{";
  strDHTMLCode += "\n\t\t\tchild.style.display = \"none\";";
  strDHTMLCode += "\n\t\t\tparentImage.src = \"" + config.ADMINSETTING_ExpandImagePathname + "\";";
  strDHTMLCode += "\n\t\t\tparentImage.alt = \"" + config.USERTEXT_VIEW_PopupExpandLink + "\";";
  strDHTMLCode += "\n\t\t}";
  strDHTMLCode += "\n\t}";
  strDHTMLCode += "\n\treturn true;";
  strDHTMLCode += "\n}";
  strDHTMLCode += "\n</SCR" + "IPT>";

  return strDHTMLCode;
}

SWEFView.prototype.getMessageList = getMessageList_vw_disc;
SWEFView.prototype.setMessageList = setMessageList_vw_disc;
SWEFView.prototype.getCurrentMessageID = getCurrentMessageID_vw_disc;
SWEFView.prototype.setCurrentMessageID = setCurrentMessageID_vw_disc;
SWEFView.prototype.getPageURLToUse = getPageURLToUse_vw_disc;
SWEFView.prototype.setPageURLToUse = setPageURLToUse_vw_disc;
SWEFView.prototype.getThreadCounter = getThreadCounter_vw_disc;
SWEFView.prototype.setThreadCounter = setThreadCounter_vw_disc;

SWEFView.prototype.getView = getView_vw_disc;
SWEFView.prototype.getDHTMLView = getDHTMLView_vw_disc;
SWEFView.prototype.getDHTMLThread = getDHTMLThread_vw_disc;
SWEFView.prototype.getDHTMLChildren = getDHTMLChildren_vw_disc;
SWEFView.prototype.getViewThread = getViewThread_vw_disc;
SWEFView.prototype.getParentSortCode = getParentSortCode_vw_disc;
SWEFView.prototype.getViewFullThread = getViewFullThread_vw_disc;
SWEFView.prototype.getIndented = getIndented_vw_disc;
SWEFView.prototype.getIndentLevel = getIndentLevel_vw_disc;
SWEFView.prototype.getSortcodeIndentLevel = getSortcodeIndentLevel_vw_disc;
SWEFView.prototype.skipToNextSibling = skipToNextSibling_vw_disc;
SWEFView.prototype.isChild = isChild_vw_disc;
SWEFView.prototype.getDHTMLExpandCollapseCode = getDHTMLExpandCollapseCode_vw_disc;
</SCRIPT>