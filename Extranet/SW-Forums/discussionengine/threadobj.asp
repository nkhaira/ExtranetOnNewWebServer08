<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// THREAD OBJECT
//
// ======================================================================

function SWEFThread
(
 nCurrentMessageID,
 strPageURLToUse
)
{
  this.setCurrentMessageID (nCurrentMessageID);
  if (isUndefined_disc (strPageURLToUse))
    {
      this.setPageURLToUse = config.getMainPagePath ();
    }
  else
    {
      this.setPageURLToUse = strPageURLToUse;
    }

  return this;
}

// ======================================================================
//
// Interface to private member variables.
//
// ======================================================================

function getCurrentMessageID_thd_disc
(
)
{
  return Number (this._currentMessageID);
}

function setCurrentMessageID_thd_disc
(
 nNewCurrentMessageID
)
{
  this._currentMessageID = Number (nNewCurrentMessageID);
  return;
}

// ======================================================================
//
// Main object methods.
//
// ======================================================================

function getAllSorted_thd_disc
(
 bSortAscending
)
{
  var dbDatabase = new SWEFDatabase ();
  var rsMessages = dbDatabase.getAllRootMessages (bSortAscending);
  var vwMessageView = new SWEFView (rsMessages, this.getCurrentMessageID ());
  var thdThreadInfo = vwMessageView.getView ();

  delete dbDatabase;
  delete rsMessages;
  delete vwMessageView;
  return thdThreadInfo;
}

function getCurrentSorted_thd_disc
(
 bSortAscending
)
{
  var dbDatabase = new SWEFDatabase ();
  var rsMessages = dbDatabase.getCurrentRootMessages (bSortAscending);
  var vwMessageView = new SWEFView (rsMessages, this.getCurrentMessageID ());
  var thdThreadInfo = vwMessageView.getView ();

  delete dbDatabase;
  delete rsMessages;
  delete vwMessageView;
  return thdThreadInfo;
}

function getCurrentSortedDHTML_thd_disc
(
 bSortAscending
)
{
  var dbDatabase = new SWEFDatabase ();
  var rsMessages = dbDatabase.getAllCurrentMessages (bSortAscending);
  var vwMessageView = new SWEFView (rsMessages, this.getCurrentMessageID ());
  var thdThreadInfo = vwMessageView.getDHTMLView ();

  delete dbDatabase;
  delete rsMessages;
  delete vwMessageView;
  return thdThreadInfo;
}

function getArchiveSorted_thd_disc
(
 dtArchiveDate,
 bSortAscending
)
{
  var dbDatabase = new SWEFDatabase ();
  var rsMessages = dbDatabase.getArchiveRootMessages (dtArchiveDate, bSortAscending);
  var vwMessageView = new SWEFView (rsMessages,
				    this.getCurrentMessageID (),
				    config.getArchivePagePath ());
  var thdThreadInfo = vwMessageView.getView ();

  delete dbDatabase;
  delete rsMessages;
  delete vwMessageView;
  return thdThreadInfo;
}

function getArchiveSortedDHTML_thd_disc
(
 dtArchiveDate,
 bSortAscending
)
{
  var dbDatabase = new SWEFDatabase ();
  var rsMessages = dbDatabase.getAllArchiveMessages (dtArchiveDate, bSortAscending);
  var vwMessageView = new SWEFView (rsMessages,
				    this.getCurrentMessageID (),
				    config.getArchivePagePath ());
  var thdThreadInfo = vwMessageView.getDHTMLView ();

  delete dbDatabase;
  delete rsMessages;
  delete vwMessageView;
  return thdThreadInfo;
}

function getExpandedThread_thd_disc
(
 nThreadID,
 nStartAt
)
{
  var dbDatabase = new SWEFDatabase ();
  var rsMessages = dbDatabase.getSubThreadMessages (nThreadID, nStartAt);
  var vwMessageView = new SWEFView (rsMessages, this.getCurrentMessageID (), this.pageURLToUse);
  var thdExpandedThread = vwMessageView.getViewThread ();

  delete dbDatabase;
  delete rsMessages;
  delete vwMessageView;
  return thdExpandedThread;
}

function getSubThread_thd_disc
(
 nThreadID,
 nStartAt
)
{
  var dbDatabase = new SWEFDatabase ();
  var rsMessages = dbDatabase.getSubThreadMessages (nThreadID, nStartAt);
  var vwMessageView = new SWEFView (rsMessages, this.getCurrentMessageID ());
  var thdSubThread = vwMessageView.getViewFullThread ();

  delete dbDatabase;
  delete rsMessages;
  delete vwMessageView;
  return thdSubThread;
}

function getFullThread_thd_disc
(
 nThreadID
)
{
  var dbDatabase = new SWEFDatabase ();
  var rsMessages = dbDatabase.getAllThreadMessages (nThreadID);
  var vwMessageView = new SWEFView (rsMessages, this.getCurrentMessageID ());
  var thdFullThread = vwMessageView.getViewFullThread ();

  delete dbDatabase;
  delete rsMessages;
  delete vwMessageView;
  return thdFullThread;
}

SWEFThread.prototype.getCurrentMessageID = getCurrentMessageID_thd_disc;
SWEFThread.prototype.setCurrentMessageID = setCurrentMessageID_thd_disc;

SWEFThread.prototype.getAllSorted = getAllSorted_thd_disc;
SWEFThread.prototype.getCurrentSorted = getCurrentSorted_thd_disc;
SWEFThread.prototype.getCurrentSortedDHTML = getCurrentSortedDHTML_thd_disc;
SWEFThread.prototype.getArchiveSorted = getArchiveSorted_thd_disc;
SWEFThread.prototype.getArchiveSortedDHTML = getArchiveSortedDHTML_thd_disc;
SWEFThread.prototype.getExpandedThread = getExpandedThread_thd_disc;
SWEFThread.prototype.getSubThread = getSubThread_thd_disc;
SWEFThread.prototype.getFullThread = getFullThread_thd_disc;
</SCRIPT>

