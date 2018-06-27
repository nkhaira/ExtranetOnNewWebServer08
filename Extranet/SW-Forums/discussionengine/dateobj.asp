<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// DATE OBJECT ADDITIONS
//
// ======================================================================

function getODBCNormalisedDate_date_disc
(
)
{
  var strYear;
  var strMonth;
  var strDay;

  strYear = this.getFullYear ();

  strMonth = this.getMonth () + 1;
  if (strMonth < 10)
    {
      strMonth = "0" + strMonth;
    }

  strDay = this.getDate ();
  if (strDay < 10)
    {
      strDay = "0" + strDay;
    }

  return strYear + "-" + strMonth + "-" + strDay;
}

function getODBCNormalisedTime_date_disc
(
)
{
  var strNormalisedTime = "";

  if (this.getHours () >= 10)
    {
      strNormalisedTime += this.getHours ();
    }
  else
    {
      strNormalisedTime += "0" + this.getHours ();
    }

  strNormalisedTime += ":";

  if (this.getMinutes () >= 10)
    {
      strNormalisedTime += this.getMinutes ();
    }
  else
    {
      strNormalisedTime += "0" + this.getMinutes ();
    }

  strNormalisedTime += ":";

  if (this.getSeconds () >= 10)
    {
      strNormalisedTime += this.getSeconds ();
    }
  else
    {
      strNormalisedTime += "0" + this.getSeconds ();
    }

  return strNormalisedTime;
}

function getODBCNormalisedTimeStamp_date_disc
(
)
{
  return this.getODBCNormalisedDate () + " " + this.getODBCNormalisedTime ();
}

function getMonthName_date_disc
(
)
{
  return this.getMonthNameByIndex (this.getMonth ());
}

function getDayName_date_disc
(
)
{
  return this.getDayNameByIndex (this.getDay ());
}

function getDateSuffix_date_disc
(
)
{
  return this.getDateSuffixByIndex (this.getDate ());
}

function getMonthNameByIndex_date_disc
(
 nIndex
)
{
  var strMonthName;
  switch (nIndex)
    {
    case 0:
      strMonthName = "January";
      break;

    case 1:
      strMonthName = "February";
      break;

    case 2:
      strMonthName = "March";
      break;

    case 3:
      strMonthName = "April";
      break;

    case 4:
      strMonthName = "May";
      break;

    case 5:
      strMonthName = "June";
      break;

    case 6:
      strMonthName = "July";
      break;

    case 7:
      strMonthName = "August";
      break;

    case 8:
      strMonthName = "September";
      break;

    case 9:
      strMonthName = "October";
      break;

    case 10:
      strMonthName = "November";
      break;

    case 11:
      strMonthName = "December";
      break;

    default:
      strMonthName = "Error";
      break;
    }

  return strMonthName;
}

function getDayNameByIndex_date_disc
(
 nIndex
)
{
  var strDayName;
  switch (nIndex)
    {
    case 0:
      strDayName = "Sunday";
      break;

    case 1:
      strDayName = "Monday";
      break;

    case 2:
      strDayName = "Tuesday";
      break;

    case 3:
      strDayName = "Wednesday";
      break;

    case 4:
      strDayName = "Thursday";
      break;

    case 5:
      strDayName = "Friday";
      break;

    case 6:
      strDayName = "Saturday";
      break;

    default:
      strDayName = "Error";
      break;
    }

  return strDayName;
}

function getDateSuffixByIndex_date_disc
(
 nIndex
)
{
  var strDateSuffix;

  switch (nIndex)
    {
    case 1:
    case 21:
    case 31:
      strDateSuffix = "st";
      break;

    case 2:
    case 22:
      strDateSuffix = "nd";
      break;

    case 3:
    case 23:
      strDateSuffix = "rd";
      break;

    case 4:
    case 5:
    case 6:
    case 7:
    case 8:
    case 9:
    case 10:
    case 11:
    case 12:
    case 13:
    case 14:
    case 15:
    case 16:
    case 17:
    case 18:
    case 19:
    case 20:
    case 21:
    case 22:
    case 23:
    case 24:
    case 25:
    case 26:
    case 27:
    case 28:
    case 29:
    case 30:
      strDateSuffix = "th";
      break;

    case 0:
    default:
      strDateSuffix = "Error";
      break;
    }

  return strDateSuffix;
}

function getShortFormat_date_disc
(
)
{
  var strHTMLout = "";
  strHTMLout += this.getDayName ();
  strHTMLout += ", ";
  strHTMLout += this.getDate ();
  strHTMLout += this.getDateSuffix ();
  strHTMLout += " ";
  strHTMLout += this.getMonthName ();
  strHTMLout += " ";
  strHTMLout += this.getFullYear();

  return strHTMLout;
}

function getLongFormat_date_disc ()
{
  var strHTMLout = "";
  strHTMLout += this.getDayName ();
  strHTMLout += ", ";
  strHTMLout += this.getDate ();
  strHTMLout += this.getDateSuffix ();
  strHTMLout += " ";
  strHTMLout += this.getMonthName ();
  strHTMLout += " ";
  strHTMLout += this.getFullYear();

  return strHTMLout;
}

Date.getMonthNameByIndex = getMonthNameByIndex_date_disc;
Date.getDayNameByIndex = getDayNameByIndex_date_disc;
Date.getDateSuffixByIndex = getDateSuffixByIndex_date_disc;

Date.prototype.getODBCNormalisedDate = getODBCNormalisedDate_date_disc;
Date.prototype.getODBCNormalisedTime = getODBCNormalisedTime_date_disc;
Date.prototype.getODBCNormalisedTimeStamp = getODBCNormalisedTimeStamp_date_disc;
Date.prototype.getMonthName = getMonthName_date_disc;
Date.prototype.getDayName = getDayName_date_disc;
Date.prototype.getDateSuffix = getDateSuffix_date_disc;
Date.prototype.getMonthNameByIndex = getMonthNameByIndex_date_disc;
Date.prototype.getDayNameByIndex = getDayNameByIndex_date_disc;
Date.prototype.getDateSuffixByIndex = getDateSuffixByIndex_date_disc;
Date.prototype.getShortFormat = getShortFormat_date_disc;
Date.prototype.getLongFormat = getLongFormat_date_disc;
</SCRIPT>

