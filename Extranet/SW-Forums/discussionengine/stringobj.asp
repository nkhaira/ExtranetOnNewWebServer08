<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// STRING OBJECT ADDITIONS
//
// ======================================================================

// ======================================================================
//
// Main object additions.
//
// ======================================================================

function strong_str_disc
(
)
{
  return SWEFHTML.FONT_open () + SWEFHTML.STRONG_open () + this + SWEFHTML.STRONG_close () + SWEFHTML.FONT_close ();
}

function strongBig_str_disc
(
)
{
  return SWEFHTML.FONT_open ("+1") + SWEFHTML.STRONG_open () + this + SWEFHTML.STRONG_close () + SWEFHTML.FONT_close ();
}

function strongSmall_str_disc
(
)
{
  return SWEFHTML.FONT_open ("-1") + SWEFHTML.STRONG_open () + this + SWEFHTML.STRONG_close () + SWEFHTML.FONT_close ();
}

function weak_str_disc
(
)
{
  return SWEFHTML.FONT_open () + this + SWEFHTML.FONT_close ();
}

function weakBig_str_disc
(
)
{
  return SWEFHTML.FONT_open ("+1") + this + SWEFHTML.FONT_close ();
}

function weakSmall_str_disc
(
)
{
  return SWEFHTML.FONT_open ("-1") + this + SWEFHTML.FONT_close ();
}

function messageBody_str_disc
(
)
{
  return SWEFHTML.BLOCKQUOTE_open () + this + SWEFHTML.BLOCKQUOTE_close ();
}

function paragraph_str_disc
(
)
{
  return SWEFHTML.P_open () + this + SWEFHTML.P_close ();
}

function show_str_disc
(
)
{
  Response.Write (this);
}

function formatForURL_str_disc
(
)
{
  var strWorkingText = this;
  strWorkingText = strWorkingText.replace (/\ /gi, "%20");
  strWorkingText = strWorkingText.replace (/\\/gi, "/");

  return strWorkingText;
}

function safeFormat_str_disc
(
)
{
  var strWorkingText = this;

  strWorkingText = Server.HTMLEncode (strWorkingText);
  strWorkingText = strWorkingText.replace (/\'/gi, "&#146;");
  strWorkingText = strWorkingText.replace (/\"/gi, "&quot;");
  strWorkingText = strWorkingText.replace (/\n/gi, " ");
  strWorkingText = strWorkingText.replace (/\r/gi, " ");

  return strWorkingText;
}

function javascriptSafeFormat_str_disc
(
)
{
  var strWorkingText = this;

  strWorkingText = strWorkingText.replace (/\"/gi, "\\\"");
  strWorkingText = strWorkingText.replace (/\n/gi, " ");
  strWorkingText = strWorkingText.replace (/\r/gi, " ");

  return strWorkingText;
}

function HTMLiseLinefeeds_str_disc
(
)
{
  var strWorkingText = this;

  strWorkingText = strWorkingText.replace (/\r\n\r\n/gi, "</P><P>");
  strWorkingText = strWorkingText.replace (/\n/gi, "<BR>");
  strWorkingText = strWorkingText.replace (/\r/gi, "");

  return strWorkingText;
}

function purify_str_disc
(
)
{
  var strWorkingText = this;

  //To filter words you would change "replace-me" to be a string you wanted to filter.
  //strWorkingText = strWorkingText.replace (/replace-me/gi, "**********");
  //strWorkingText = strWorkingText.replace (/replace-me-too/gi, "****");

  return strWorkingText;
}

function formatForStoring_str_disc
(
)
{
  var strWorkingText = this.purify();

  if (!config.ADMINSWITCH_AllowRichFormatting)
    {
      strWorkingText = strWorkingText.stripAllTags ();
    }

  strWorkingText = strWorkingText.replace (/\r\n\r\n/gi, SWEFHTML.P_open () + SWEFHTML.P_close ());
  strWorkingText = strWorkingText.removeMaliciousTags ();
  strWorkingText = strWorkingText.convertComplexTags ();
  strWorkingText = strWorkingText.fixBrokenTags ();

  return strWorkingText;
}

function unformatFromStoring_str_disc
(
)
{
  var strWorkingText = this;

  strWorkingText = strWorkingText.replace (/\<P[^\>]*\>/gi, "\n");
  strWorkingText = strWorkingText.replace (/\<\/P[^\>]\>/gi, "\n");
  strWorkingText = strWorkingText.replace (/\<BR[^\>]*\>/gi, "\n");
  strWorkingText = strWorkingText.replace (/\&amp\;/gi, "&");
  strWorkingText = strWorkingText.replace (/\"/gi, "&quot;");

  return strWorkingText;
}

function removeMaliciousTags_str_disc
(
)
{
  var strWorkingText = this;
  strWorkingText = strWorkingText.replace (/\<\s*SCRIPT[^\\>]*\>/gi, " ");
  strWorkingText = strWorkingText.replace (/\<\s*META[^\\>]*\>/gi, " ");
  strWorkingText = strWorkingText.replace (/\<\s*\/?\s*BODY[^\\>]*\>/gi, " ");
  strWorkingText = strWorkingText.replace (/\<\s*\/?\s*HTML[^\\>]*\>/gi, " ");

  var reJSEvents = new RegExp ("\\<[^\\>]*("
			       + config.SYS_AllJavascriptEvents
			       + ")[^\\>]*\\>", "gi");
  strWorkingText = strWorkingText.replace (reJSEvents, " ");

  return strWorkingText;
}

function fixBrokenTags_str_disc
(
)
{
  var strWorkingText = this;

  strWorkingText = strWorkingText.fixTag ("b");
  strWorkingText = strWorkingText.fixTag ("i");
  strWorkingText = strWorkingText.fixTag ("u");
  strWorkingText = strWorkingText.fixTag ("ul");
  strWorkingText = strWorkingText.fixTag ("ol");
  strWorkingText = strWorkingText.fixTag ("font");
  strWorkingText = strWorkingText.fixTag ("h1");
  strWorkingText = strWorkingText.fixTag ("h2");
  strWorkingText = strWorkingText.fixTag ("h3");
  strWorkingText = strWorkingText.fixTag ("h4");
  strWorkingText = strWorkingText.fixTag ("h5");
  strWorkingText = strWorkingText.fixTag ("h6");
  strWorkingText = strWorkingText.fixTag ("table");
  strWorkingText = strWorkingText.fixTag ("tr");
  strWorkingText = strWorkingText.fixTag ("th");
  strWorkingText = strWorkingText.fixTag ("td");

  return strWorkingText;
}

function fixTag_str_disc
(
 strTag
)
{
  var reOpenTag = new RegExp ("\<" + strTag + "\>", "gi");
  var reCloseTag = new RegExp ("\\<\\s*\\/\\s*" + strTag + "[^a-zA-Z0-9][^\\>]*\\>", "gi");
  var nTagDifference = this.countTags (reOpenTag) - this.countTags (reCloseTag);
  var strAdditionalTags = "";

  if (nTagDifference > 0)
    {
      for (var nCounter = 0; nCounter < nTagDifference; nCounter++)
	{
	  strAdditionalTags += "</" + strTag + " HTMLFixup>";
	}
    }

  delete reOpenTag;
  delete reCloseTag;
  return this + strAdditionalTags;
}

function countTags_str_disc
(
 strTag
)
{
  var nCounter = 0;
  var strWorkingText = this;
  var nFoundAt = strWorkingText.search (strTag);

  while (nFoundAt != -1)
    {
      nCounter += 1;
      strWorkingText = strWorkingText.substr (nFoundAt + 1, strWorkingText.length - nFoundAt + 1);

      nFoundAt = strWorkingText.search (strTag);
    }

  return nCounter;
}

function convertComplexTags_str_disc
(
)
{
  var strWorkingText = this;
  var strNewText = "";
  var nStartAt = strWorkingText.search ("<");
  var nEndAt = strWorkingText.search (">") + 1;
  var nLinkAt = strWorkingText.search (this.PROTOCOLS_RX_DISC);
  this.inLink = false;

  while ((nStartAt != -1) || (nLinkAt != -1))
    {
      if (((nLinkAt < nStartAt) || (nStartAt == -1)) && (nLinkAt >= 0) && (!this.inLink))
	{
	  strNewText += strWorkingText.substring (0, nLinkAt);
	  var strURL = strWorkingText.substr (nLinkAt);
	  var strURLEnd =  new RegExp (this.END_OF_URL_RX_DISC, "gi");
	  nEndAt = strURL.search (strURLEnd);
	  nEndAt = (nEndAt == -1 ? strURL.length : nEndAt);
	  var strLinkText = strWorkingText.substring (nLinkAt, nLinkAt + nEndAt);
	  strLinkText = strLinkText.stripAllTags ();
	  strLinkText = strLinkText.protocolise ();
	  strNewText += SWEFHTML.A_open (strLinkText,
					 config.USERTEXT_STRING_WarningUnverifiedLink,
					 undefined_disc,
					 config.ADMINSETTING_DefaultEmbeddedLinkTarget);
	  strNewText += strLinkText;
	  strNewText += SWEFHTML.A_close ();
	  strWorkingText = strURL.substr (nEndAt);
	}
      else
	{
	  strNewText += strWorkingText.substring (0, nStartAt);
	  strNewText += strWorkingText.substring (nStartAt, nEndAt).sanitiseTag ();
	  strWorkingText = strWorkingText.substr (nEndAt);
	}

      nStartAt = strWorkingText.search ("<");
      nEndAt = strWorkingText.search (">") + 1;
      nLinkAt = strWorkingText.search (this.PROTOCOLS_RX_DISC);
    }

  return strNewText + strWorkingText;
}

function sanitiseTag_str_disc
(
)
{
  var strTag = new SWEFHTML (this);

  if (strTag.isLinkOpen ())
    {
      this.inLink = true;
    }
  else if (strTag.isLinkClose ())
    {
      this.inLink = false;
    }

  return strTag.sanitise ();
}

function getAttribute_str_disc
(
 strAttributeName
)
{
  var strAttributeValue;
  var reAttrib = new RegExp (strAttributeName + "\s*=\s*", "i");
  var nFoundAt = this.search (reAttrib);

  if (nFoundAt > -1)
    {
      var strSubtag = this.substr (nFoundAt);
      var nTagStart = strSubtag.search ("=");
      strSubtag = strSubtag.substr (nTagStart + 1);

      var nCharCount = 0;
      while ((nCharCount < strSubtag.length) && (strSubtag.charAt (nCharCount) == " "))
	{
	  nCharCount++;
	}

      var bQuoted = false;
      var strQuoteUsed = "";
      if ((strSubtag.charAt (nCharCount) == "\"") || (strSubtag.charAt (nCharCount) == "'"))
	{
	  bQuoted = true;
	  strQuoteUsed = strSubtag.charAt (nCharCount);
	  nCharCount += 1
	}

      strSubtag = strSubtag.substr (nCharCount);

      var strAttribEndMarker = "\\s";
      if (bQuoted == true)
	{
	  strAttribEndMarker = strQuoteUsed;
	}

      var nAttribEnd = strSubtag.search (strAttribEndMarker);

      strAttributeValue = strSubtag.substr (0, nAttribEnd);

      if (strAttributeValue == "")
	{
	  strAttributeValue = undefined_disc;
	}
    }

  return strAttributeValue;
}

function protocolise_str_disc
(
)
{
  var strProtocolisedVersion = "";

  var strProtocol = this.substr (0, this.indexOf ("//"));

  switch (strProtocol)
    {
    case "file:":
    case "ftp:":
    case "gopher:":
    case "http:":
    case "https:":
    case "mailto:":
    case "news:":
    case "telnet:":
    case "wais:":
      strProtocolisedVersion = this;
      break;

    default:
      strProtocolisedVersion = "http://" + this;
      break;
    }

  return strProtocolisedVersion;
}

function precis_str_disc
(
)
{
  var strPrecisString = this;

  if (strPrecisString.length > config.ADMINSETTING_PrecisLength)
    {
      strPrecisString = strPrecisString.substr (0, config.ADMINSETTING_PrecisLength)
	+ config.USERTEXT_STRING_StringTruncatedSuffix;
    }

  strPrecisString = strPrecisString.replace (/\<[^\\>]*\>/gi, " ");

  if (strPrecisString.lastIndexOf (">") < strPrecisString.lastIndexOf ("<"))
    {
      strPrecisString = strPrecisString.substr (0, strPrecisString.lastIndexOf ("<"));
    }

  return strPrecisString;
}

function stripAllTags_str_disc
(
)
{
  var strTextWithoutHTML = this;
  strTextWithoutHTML = strTextWithoutHTML.replace (/<[^>]*>/gi, " ");
  strTextWithoutHTML = strTextWithoutHTML.replace (/&nbsp;/gi, " ");
  strTextWithoutHTML = strTextWithoutHTML.replace (/&amp;/gi, "&");
  strTextWithoutHTML = strTextWithoutHTML.replace (/&#145;/gi, "'");
  strTextWithoutHTML = strTextWithoutHTML.replace (/&#146;/gi, "'");

  return strTextWithoutHTML;
}

String.prototype.PROTOCOLS_RX_DISC = "http:|www\\.|ftp:|https:|gopher:|file:|mailto:|news:|telnet:|wais:";
String.prototype.END_OF_URL_RX_DISC = "[^A-Za-z0-9\.:;,=~@_&\\/\\?\\!\\+\\$\\{\\}\\-]";

String.prototype.strong = strong_str_disc;
String.prototype.strongBig = strongBig_str_disc;
String.prototype.strongSmall = strongSmall_str_disc;
String.prototype.weak = weak_str_disc;
String.prototype.weakBig = weakBig_str_disc;
String.prototype.weakSmall = weakSmall_str_disc;
String.prototype.messageBody = messageBody_str_disc;
String.prototype.paragraph = paragraph_str_disc;
String.prototype.show = show_str_disc;
String.prototype.formatForURL = formatForURL_str_disc;
String.prototype.safeFormat = safeFormat_str_disc;
String.prototype.javascriptSafeFormat = javascriptSafeFormat_str_disc;
String.prototype.HTMLiseLinefeeds = HTMLiseLinefeeds_str_disc;
String.prototype.purify = purify_str_disc;
String.prototype.formatForStoring = formatForStoring_str_disc;
String.prototype.unformatFromStoring = unformatFromStoring_str_disc;
String.prototype.removeMaliciousTags = removeMaliciousTags_str_disc;
String.prototype.fixBrokenTags = fixBrokenTags_str_disc;
String.prototype.fixTag = fixTag_str_disc;
String.prototype.countTags = countTags_str_disc;
String.prototype.convertComplexTags = convertComplexTags_str_disc;
String.prototype.sanitiseTag = sanitiseTag_str_disc;
String.prototype.getAttribute = getAttribute_str_disc;
String.prototype.protocolise = protocolise_str_disc;
String.prototype.precis = precis_str_disc;
String.prototype.stripAllTags = stripAllTags_str_disc;
</SCRIPT>

