<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// HTML OBJECT
//
// ======================================================================

function SWEFHTML
(
 strTagString
)
{
  this.setTagSource (strTagString);

  return this;
}

// ======================================================================
//
// Interface to private member variables.
//
// ======================================================================

function getTagSource_html_disc
(
)
{
  return String (this._tagSource);
}

function setTagSource_html_disc
(
 strNewTagSource
)
{
  this._tagSource = String (strNewTagSource);

  return;
}

// ======================================================================
//
// Main object methods.
//
// ======================================================================

function A_open_html_disc
(
 strLinkURL_mand,
 strLinkTitle,
 strLinkName,
 strLinkTarget,
 strLinkID,
 strOnClick
)
{
  var strHTMLout = "";

  strHTMLout += "<A HREF=\"";
  strHTMLout += strLinkURL_mand;
  strHTMLout += "\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strLinkTitle))
    {
      strHTMLout += " TITLE=\"";
      strHTMLout += strLinkTitle;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strLinkName))
    {
      strHTMLout += " NAME=\"";
      strHTMLout += strLinkName;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strLinkTarget))
    {
      strHTMLout += " TARGET=\"";
      strHTMLout += strLinkTarget;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strLinkID))
    {
      strHTMLout += " ID=\"";
      strHTMLout += strLinkID;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strOnClick))
    {
      strHTMLout += " onclick=\"";
      strHTMLout += strOnClick;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function A_close_html_disc
(
)
{
  return "</A>";
}

function BASE_html_disc
(
 strHREF_mand
)
{
  return "<BASE HREF=\"" + strHREF_mand + "\">";
}

function BLOCKQUOTE_open_html_disc
(
)
{
  var strHTMLout = "";

  strHTMLout += "<BLOCKQUOTE";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function BLOCKQUOTE_close_html_disc
(
)
{
  return "</BLOCKQUOTE>\n";
}

function BODY_open_html_disc
(
 strBackgroundColour,
 strBackgroundImage,
 strNormalLinkColour,
 strActiveLinkColour,
 strVisitedLinkColour,
 nTopMargin,
 nLeftMargin,
 nMarginHeight,
 nMarginWidth
)
{
  var strHTMLout = "";
  strHTMLout += "<BODY";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strBackgroundColour))
    {
      strHTMLout += " BGCOLOR=\"";
      strHTMLout += strBackgroundColour;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strBackgroundImage))
    {
      strHTMLout += " BACKGROUND=\"";
      strHTMLout += strBackgroundImage;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strNormalLinkColour))
    {
      strHTMLout += " LINK=\"";
      strHTMLout += strNormalLinkColour;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strActiveLinkColour))
    {
      strHTMLout += " ALINK=\"";
      strHTMLout += strActiveLinkColour;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strVisitedLinkColour))
    {
      strHTMLout += " VLINK=\"";
      strHTMLout += strVisitedLinkColour;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nTopMargin))
    {
      strHTMLout += " TOPMARGIN=\"";
      strHTMLout += nTopMargin;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nLeftMargin))
    {
      strHTMLout += " LEFTMARGIN=\"";
      strHTMLout += nLeftMargin;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nMarginHeight))
    {
      strHTMLout += " MARGINHEIGHT=\"";
      strHTMLout += nMarginHeight;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nMarginWidth))
    {
      strHTMLout += " MARGINWIDTH=\"";
      strHTMLout += nMarginWidth;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function BODY_close_html_disc
(
)
{
  return "</BODY>\n\n";
}

function BR_html_disc
(
)
{
  var strHTMLout = "";

  strHTMLout += "<BR";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function DIV_open_html_disc
(
 strID,
 strStyle,
 strAlignment
)
{
  var strHTMLout = "";
  strHTMLout += "<DIV";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strID))
    {
      strHTMLout += " ID=\"";
      strHTMLout += strID;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strStyle))
    {
      strHTMLout += " STYLE=\"";
      strHTMLout += strStyle;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strAlignment))
    {
      strHTMLout += " ALIGN=\"";
      strHTMLout += strAlignment;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function DIV_close_html_disc
(
)
{
  return "</DIV>\n";
}

function DOUBLE_QUOTES_html_disc
(
)
{
  return "&quot;";
}

function DTD_html_disc
(
)
{
  return "<!DOCTYPE HTML PUBLIC \"-//IETF//DTD HTML//EN\">\n";
}

function FONT_open_html_disc
(
 strSize,
 strFace,
 strColour,
 strStyle
)
{
  var strHTMLout = "";
  strHTMLout += "<FONT";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strSize))
    {
      strHTMLout += " SIZE=\"";
      strHTMLout += strSize;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strFace))
    {
      strHTMLout += " FACE=\"";
      strHTMLout += strFace;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strColour))
    {
      strHTMLout += " COLOR=\"";
      strHTMLout += strColour;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strStyle))
    {
      strHTMLout += " STYLE=\"";
      strHTMLout += strStyle;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function FONT_close_html_disc
(
)
{
  return "</FONT>";
}

function FORM_open_html_disc
(
 strActionURL_mand,
 strMethod,
 strName,
 strTarget
)
{
  var strHTMLout = "";
  strHTMLout += "<FORM ACTION=\"";
  strHTMLout += strActionURL_mand;
  strHTMLout += "\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strMethod))
    {
      strHTMLout += " METHOD=\"";
      strHTMLout += strMethod;
      strHTMLout += "\"";
    }
  else
    {
      strHTMLout += " METHOD=\"post\"";
    }

  if (isDefined_disc (strName))
    {
      strHTMLout += " NAME=\"";
      strHTMLout += strName;
      strHTMLout += "\"";
    }
  else
    {
      strHTMLout += " NAME=\"SWEF\"";
    }

  if (isDefined_disc (strTarget))
    {
      strHTMLout += " TARGET=\"";
      strHTMLout += strTarget;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function FORM_close_html_disc
(
)
{
  return "</FORM>\n\n";
}

function HR_html_disc
(
 strWidth
)
{
  var strHTMLout = "";

  strHTMLout += "<HR ";

  if (isDefined_disc (strWidth))
    {
      strHTMLout += strWidth;
    }
  else
    {
      strHTMLout += "80%";
    }
  strHTMLout += ">";

  return strHTMLout;
}

function HTML_open_html_disc
(
)
{
  var strHTMLout = "";

  strHTMLout += "<HTML";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function HTML_close_html_disc
(
)
{
  return "</HTML>\n";
}

function IFRAME_open_html_disc
(
 strName_mand,
 strSrc_mand,
 strID,
 nWidth,
 nHeight
)
{
  var strHTMLout = "";
  strHTMLout += "<IFRAME NAME=\"";
  strHTMLout += strName_mand;
  strHTMLout += "\" SRC=\"";
  strHTMLout += strSrc_mand;
  strHTMLout += "\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strID))
    {
      strHTMLout += " ID=\"";
      strHTMLout += strID;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nWidth))
    {
      strHTMLout += " WIDTH=\"";
      strHTMLout += nWidth;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nHeight))
    {
      strHTMLout += " HEIGHT=\"";
      strHTMLout += nHeight;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function IFRAME_close_html_disc
(
)
{
  return "</IFRAME>\n";
}

function IMG_html_disc
(
 strSourceURL_mand,
 strAltText,
 nBorder,
 nWidth,
 nHeight,
 strID,
 strName,
 strOnClick,
 strOnMouseOver,
 strOnMouseOut,
 strOnMouseDown,
 strOnMouseUp,
 strAlign,
 nHspace,
 nVspace
)
{
  var strHTMLout = "";
  strHTMLout += "<IMG SRC=\"";
  strHTMLout += strSourceURL_mand;
  strHTMLout += "\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strAltText))
    {
      strHTMLout += " ALT=\"";
      strHTMLout += strAltText;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nBorder))
    {
      strHTMLout += " BORDER=\"";
      strHTMLout += nBorder;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nWidth))
    {
      strHTMLout += " WIDTH=\"";
      strHTMLout += nWidth;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nHeight))
    {
      strHTMLout += " HEIGHT=\"";
      strHTMLout += nHeight;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strID))
    {
      strHTMLout += " ID=\"";
      strHTMLout += strID;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strName))
    {
      strHTMLout += " NAME=\"";
      strHTMLout += strName;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strOnClick))
    {
      strHTMLout += " onclick=\"";
      strHTMLout += strOnClick;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strOnMouseOver))
    {
      strHTMLout += " onmouseover=\"";
      strHTMLout += strOnMouseOver;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strOnMouseOut))
    {
      strHTMLout += " onmouseout=\"";
      strHTMLout += strOnMouseOut;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strOnMouseDown))
    {
      strHTMLout += " onmousedown=\"";
      strHTMLout += strOnMouseDown;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strOnMouseUp))
    {
      strHTMLout += " onmouseup=\"";
      strHTMLout += strOnMouseUp;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strAlign))
    {
      strHTMLout += " ALIGN=\"";
      strHTMLout += strAlign;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nHspace))
    {
      strHTMLout += " HSPACE=\"";
      strHTMLout += nHspace;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nVspace))
    {
      strHTMLout += " VSPACE=\"";
      strHTMLout += nVspace;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function INPUT_checkbox_html_disc
(
 strFieldName_mand,
 strFieldValue_mand,
 strCheckedOrNot
)
{
  var strHTMLout = "";
  strHTMLout += "<INPUT NAME=\"";
  strHTMLout += strFieldName_mand;
  strHTMLout += "\" VALUE=\"";
  strHTMLout += strFieldValue_mand;
  strHTMLout += "\" TYPE=\"checkbox\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strCheckedOrNot))
    {
      strHTMLout += strCheckedOrNot;
    }

  strHTMLout += ">";

  return strHTMLout;
}

function INPUT_hidden_html_disc
(
 strFieldName_mand,
 strFieldValue_mand
)
{
  var strHTMLout = "";
  strHTMLout += "<INPUT NAME=\"";
  strHTMLout += strFieldName_mand;
  strHTMLout += "\" VALUE=\"";
  strHTMLout += strFieldValue_mand;
  strHTMLout += "\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  strHTMLout += "\" TYPE=\"hidden\">";

  return strHTMLout;
}

function INPUT_password_html_disc
(
 strFieldName_mand,
 strFieldValue_mand,
 strFieldSize,
 strFieldMaxLength
)
{
  var strHTMLout = "";
  strHTMLout += "<INPUT NAME=\"";
  strHTMLout += strFieldName_mand;
  strHTMLout += "\" VALUE=\"";
  strHTMLout += strFieldValue_mand;
  strHTMLout += "\" TYPE=\"password\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isUndefined_disc (strFieldSize))
    {
      strFieldSize = config.ADMINSETTING_TextAreaRows;
    }

  strHTMLout += " SIZE=\"";
  strHTMLout += strFieldSize;
  strHTMLout += "\"";

  if (isUndefined_disc (strFieldMaxLength))
    {
      strFieldMaxLength = config.ADMINSETTING_InputboxMaxLength;
    }
  strHTMLout += " MAXLENGTH=\"";
  strHTMLout += strFieldMaxLength;
  strHTMLout += "\">";

  return strHTMLout;
}

function INPUT_submit_html_disc
(
 strFieldName,
 strFieldValue_mand
)
{
  var strHTMLout = "";

  if (SWEFTextControl.isUsed ())
    {
      strHTMLout += "<INPUT VALUE=\"";
      strHTMLout += strFieldValue_mand;
      strHTMLout += "\" TYPE=\"button\" onclick=\"submitRecord (); return false;\"";

      if (isDefined_disc (strFieldName))
	{
	  strHTMLout += " NAME=\"";
	  strHTMLout += strFieldName;
	  strHTMLout += "\"";
	}

      if (isDefined_disc (config.ADMINSETTING_CSSClass))
	{
	  strHTMLout += " CLASS=\"";
	  strHTMLout += "NavLeftHighlight1";	//config.ADMINSETTING_CSSClass;
	  strHTMLout += "\"";
	}

      strHTMLout += ">";
    }
  else
    {
      strHTMLout += "<INPUT VALUE=\"";
      strHTMLout += strFieldValue_mand;
      strHTMLout += "\" TYPE=\"submit\"";

      if (isDefined_disc (config.ADMINSETTING_CSSClass))
	{
	  strHTMLout += " CLASS=\"";
	  strHTMLout += "NavLeftHighlight1";   // config.ADMINSETTING_CSSClass;
	  strHTMLout += "\"";
	}

      if (isDefined_disc (strFieldName))
	{
	  strHTMLout += " NAME=\"";
	  strHTMLout += strFieldName;
	  strHTMLout += "\"";
	}
      strHTMLout += ">";
    }

  return strHTMLout;
}

function INPUT_text_html_disc
(
 strFieldName_mand,
 strFieldValue_mand,
 strFieldSize,
 strFieldMaxLength
)
{
  var strHTMLout = "";
  strHTMLout += "<INPUT NAME=\"";
  strHTMLout += strFieldName_mand;
  strHTMLout += "\" VALUE=\"";
  strHTMLout += strFieldValue_mand;
  strHTMLout += "\" TYPE=\"text\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isUndefined_disc (strFieldSize))
    {
      strFieldSize = config.ADMINSETTING_TextAreaRows;
    }
  strHTMLout += " SIZE=\"";
  strHTMLout += strFieldSize;
  strHTMLout += "\"";

  if (isUndefined_disc (strFieldMaxLength))
    {
      strFieldMaxLength = config.ADMINSETTING_InputboxMaxLength;
    }
  strHTMLout += " MAXLENGTH=\"";
  strHTMLout += strFieldMaxLength;
  strHTMLout += "\">";

  return strHTMLout;
}

function LI_html_disc
(
)
{
  var strHTMLout = "";

  strHTMLout += "<LI";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function JS_open_html_disc
(
)
{
  return "\n<SCR" + "IPT LANGUAGE=\"JavaScript\">\n<" + "!--\n";
}

function JS_close_html_disc
(
)
{
  return "\n/" + "/ -->\n</SCR" + "IPT>\n";
}

function MAILTO_html_disc
(
 strEmailAddress,
 strName_mand,
 strEmailLinkText
)
{
  var strHTMLout = "";
  if (config.ADMINSWITCH_ShowEmailAddresses)
    {
      var strLinkURL = "mailto:" + strEmailAddress;
      var strLinkText = "";
      strLinkText += config.USERTEXT_SHOW_PopupEmailPrefix;
      strLinkText += strName_mand;
      strLinkText += " (";
      strLinkText += strEmailAddress;
      strLinkText += ")";

      strHTMLout += SWEFHTML.A_open (strLinkURL, strLinkText);
      strHTMLout += strEmailLinkText;
      strHTMLout += SWEFHTML.A_close ();
    }
  else
    {
      strHTMLout += strName_mand;
    }

  return strHTMLout;
}

function NBSP_html_disc
(
 nNumRequired
)
{
  var strHTMLout = "&nbsp;";

  if (isDefined_disc (nNumRequired))
    {
      for (var nCounter = 1; nCounter < nNumRequired; nCounter++)
	{
	  strHTMLout += "&nbsp;";
	}
    }

  return strHTMLout;
}

function OL_open_html_disc
(
)
{
  var strHTMLout = "";

  strHTMLout += "<OL";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function OL_close_html_disc
(
)
{
  return "</OL>\n";
}

function OPTION_html_disc
(
 strValue_mand,
 bSelected
)
{
  var strSelected = "";
  if (isDefined_disc (bSelected) && bSelected)
    {
      strSelected = " SELECTED";
    }

  var strHTMLout = "";
  strHTMLout += "<OPTION VALUE=\"";
  strHTMLout += strValue_mand;
  strHTMLout += "\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  strHTMLout += strSelected;
  strHTMLout += ">";

  return strHTMLout;
}

function P_open_html_disc
(
)
{
  var strHTMLout = "";

  strHTMLout += "<P";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function P_close_html_disc
(
)
{
  return "</P>";
}

function QUOTE_open_html_disc
(
)
{
  return "&#145;";
}

function QUOTE_close_html_disc
(
)
{
  return "&#146;";
}

function SCRIPT_open_html_disc
(
 strSource,
 strID,
 strLanguage
)
{
  var strHTMLout = "";

  strHTMLout += "<" + "SCRIPT";

  if (isDefined_disc (strSource))
    {
      strHTMLout += " SRC=\"";
      strHTMLout += strSource;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strID))
    {
      strHTMLout += " ID=\"";
      strHTMLout += strID;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strLanguage))
    {
      strHTMLout += " LANGUAGE=\"";
      strHTMLout += strLanguage;
      strHTMLout += "\"";
    }
  else
    {
      strHTMLout += " LANGUAGE=\"Javascript\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function SCRIPT_close_html_disc
(
)
{
  return "</" + "SCRIPT>\n\n";
}

function SELECT_open_html_disc
(
 strName_mand,
 strOnChange
)
{
  var strHTMLout = "";
  strHTMLout += "<SELECT NAME=\"";
  strHTMLout += strName_mand;
  strHTMLout += "\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strOnChange))
    {
      strHTMLout += " onchange=\"";
      strHTMLout += strOnChange;
      strHTMLout += "\"";
    }
  strHTMLout += ">";

  return strHTMLout;
}

function SELECT_close_html_disc
(
)
{
  return "</SELECT>\n";
}

function SPAN_open_html_disc
(
 strClass_mand
)
{
  var strHTMLout = "";
  strHTMLout += "<SPAN CLASS=\"";
  strHTMLout += strClass_mand;
  strHTMLout += "\">";

  return strHTMLout;
}

function SPAN_close_html_disc
(
)
{
  return "</SPAN>\n";
}

function STRONG_open_html_disc
(
)
{
  var strHTMLout = "";

  strHTMLout += "<STRONG";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function STRONG_close_html_disc
(
)
{
  return "</STRONG>";
}

function STYLE_open_html_disc
(
 strType
)
{
  if (isUndefined_disc (strType))
    {
      strType = "text/css";
    }

  return "<STYLE TYPE=\"" + strType + "\">";
}

function STYLE_close_html_disc
(
)
{
  return "</STYLE>\n";
}

function TABLE_open_html_disc
(
 nBorder,
 nWidth,
 strBackgroundColour,
 strStyle,
 nCellpadding,
 nCellspacing,
 strBackgroundImage
)
{
  var strHTMLout = "";
  strHTMLout += "<TABLE";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nBorder))
    {
      strHTMLout += " BORDER=\"";
      strHTMLout += nBorder;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nWidth))
    {
      strHTMLout += " WIDTH=\"";
      strHTMLout += nWidth;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strBackgroundColour))
    {
      strHTMLout += " BGCOLOR=\"";
      strHTMLout += strBackgroundColour;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strStyle))
    {
      strHTMLout += " STYLE=\"";
      strHTMLout += strStyle;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nCellpadding))
    {
      strHTMLout += " CELLPADDING=\"";
      strHTMLout += nCellpadding;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nCellspacing))
    {
      strHTMLout += " CELLSPACING=\"";
      strHTMLout += nCellspacing;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strBackgroundImage))
    {
      strHTMLout += " BACKGROUND=\"";
      strHTMLout += strBackgroundImage;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function TABLE_close_html_disc
(
)
{
  return "</TABLE>\n\n";
}

function TD_open_html_disc
(
 nWidth,
 strStyle,
 strBackgroundColour,
 strAlignment,
 strVAlignment,
 nColspan,
 bNowrap,
 strBackgroundImage
)
{
  var strHTMLout = "";
  strHTMLout += "<TD";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nWidth))
    {
      strHTMLout += " WIDTH=\"";
      strHTMLout += nWidth;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strStyle))
    {
      strHTMLout += " STYLE=\"";
      strHTMLout += strStyle;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strBackgroundColour))
    {
      strHTMLout += " BGCOLOR=\"";
      strHTMLout += strBackgroundColour;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strAlignment))
    {
      strHTMLout += " ALIGN=\"";
      strHTMLout += strAlignment;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strVAlignment))
    {
      strHTMLout += " VALIGN=\"";
      strHTMLout += strVAlignment;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nColspan))
    {
      strHTMLout += " COLSPAN=\"";
      strHTMLout += nColspan;
      strHTMLout += "\"";
    }

  if (isDefined_disc (bNowrap))
    {
      if (bNowrap)
	{
	  strHTMLout += " NOWRAP";
	}
    }

  if (isDefined_disc (strBackgroundImage))
    {
      strHTMLout += " BACKGROUND=\"";
      strHTMLout += strBackgroundImage;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function TD_close_html_disc
(
)
{
  return "</TD>\n";
}

function TEXTAREA_open_html_disc
(
 strFieldName_mand,
 nColumns,
 nRows
)
{
  var strHTMLout = "";
  strHTMLout += "<TEXTAREA NAME=\"";
  strHTMLout += strFieldName_mand;
  strHTMLout += "\" WRAP=\"PHYSICAL\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nColumns))
    {
      strHTMLout += " COLS=\"";
      strHTMLout += nColumns;
      strHTMLout += "\"";
    }

  if (isDefined_disc (nRows))
    {
      strHTMLout += " ROWS=\"";
      strHTMLout += nRows;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function TEXTAREA_close_html_disc
(
)
{
  return "</TEXTAREA>\n";
}

function TH_open_html_disc
(
 strStyle,
 strAlignment,
 strVAlignment,
 bNowrap
)
{
  var strHTMLout = "";
  strHTMLout += "<TH";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strStyle))
    {
      strHTMLout += " STYLE=\"";
      strHTMLout += strStyle;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strAlignment))
    {
      strHTMLout += " ALIGN=\"";
      strHTMLout += strAlignment;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strVAlignment))
    {
      strHTMLout += " VALIGN=\"";
      strHTMLout += strVAlignment;
      strHTMLout += "\"";
    }

  if (isDefined_disc (bNowrap))
    {
      if (bNowrap)
	{
	  strHTMLout += " NOWRAP";
	}
    }

  strHTMLout += ">";

  return strHTMLout;
}

function TH_close_html_disc
(
)
{
  return "</TH>\n";
}

function TR_open_html_disc
(
 strStyle
)
{
  var strHTMLout = "";
  strHTMLout += "<TR";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  if (isDefined_disc (strStyle))
    {
      strHTMLout += " STYLE=\"" + strStyle + "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function TR_close_html_disc
(
)
{
  return "</TR>\n";
}

function UL_open_html_disc
(
)
{
  var strHTMLout = "";

  strHTMLout += "<UL TYPE=\"square\"";

  if (isDefined_disc (config.ADMINSETTING_CSSClass))
    {
      strHTMLout += " CLASS=\"";
      strHTMLout += config.ADMINSETTING_CSSClass;
      strHTMLout += "\"";
    }

  strHTMLout += ">";

  return strHTMLout;
}

function UL_close_html_disc
(
)
{
  return "</UL>\n";
}

function sanitise_html_disc
(
)
{
  var strSanitisedTag = "";

  if (this.isImage ())
    {
      strSanitisedTag = this.sanitiseImage ();
    }
  else if (this.isLinkOpen ())
    {
      strSanitisedTag = this.sanitiseLink ();
    }
  else
    {
      strSanitisedTag = this.getTagSource ();
    }

  return strSanitisedTag;
}

function isImage_html_disc
(
)
{
  var bIsImageTag = false;
  if (this.getTagSource ().search ("^<\\s*IMG\\s") > -1)
    {
      bIsImageTag = true;
    }

  return bIsImageTag;
}

function isLinkOpen_html_disc
(
)
{
  var bIsLinkTag = false;
  if (this.getTagSource ().search ("^<\\s*A\\s") > -1)
    {
      bIsLinkTag = true;
    }

  return bIsLinkTag;
}

function isLinkClose_html_disc ()
{
  var bIsLinkTag = false;
  if (this.getTagSource ().search ("^<\\s*\\/\\s*A\\s") > -1)
    {
      bIsLinkTag = true;
    }

  return bIsLinkTag;
}

// ======================================================================
//
// Private member methods.
//
// ======================================================================

function sanitiseImage_html_disc
(
)
{
  var strTag = this.getTagSource ();

  var strSrc = strTag.getAttribute ("src");
  var strAlt = strTag.getAttribute ("alt");
  if ((isDefined_disc (strAlt)) && (strAlt != config.USERTEXT_STRING_WarningUnverifiedImage))
    {
      strAlt = config.USERTEXT_STRING_WarningUnverifiedImage + this.QUOTE_open () + strAlt + this.QUOTE_close ();
    }
  else
    {
      strAlt = config.USERTEXT_STRING_WarningUnverifiedImage;
    }

  var nBorder = strTag.getAttribute ("border");
  var strAlign = strTag.getAttribute ("align");
  var nHspace = strTag.getAttribute ("hspace");
  var nVspace = strTag.getAttribute ("vspace");

  return this.IMG (strSrc,
		   strAlt,
		   nBorder,
		   undefined_disc,
		   undefined_disc,
		   undefined_disc,
		   undefined_disc,
		   undefined_disc,
		   undefined_disc,
		   undefined_disc,
		   undefined_disc,
		   undefined_disc,
		   strAlign,
		   nHspace,
		   nVspace);
}

function sanitiseLink_html_disc
(
)
{
  var strTag = this.getTagSource ();
  var strHref = strTag.getAttribute ("href");
  var strTitle = strTag.getAttribute ("title");
  if ((isDefined_disc (strTitle)) && (strTitle != config.USERTEXT_STRING_WarningUnverifiedLink))
    {
      strTitle = config.USERTEXT_STRING_WarningUnverifiedLink + this.QUOTE_open () + strTitle + this.QUOTE_close ();
    }
  else
    {
      strTitle = config.USERTEXT_STRING_WarningUnverifiedLink;
    }
  var strName = strTag.getAttribute ("name");
  var strID = strTag.getAttribute ("id");

  return this.A_open (strHref, strTitle, strName, undefined_disc, strID);
}

SWEFHTML.A_open = A_open_html_disc;
SWEFHTML.A_close = A_close_html_disc;
SWEFHTML.BASE = BASE_html_disc;
SWEFHTML.BLOCKQUOTE_open = BLOCKQUOTE_open_html_disc;
SWEFHTML.BLOCKQUOTE_close = BLOCKQUOTE_close_html_disc;
SWEFHTML.BODY_open = BODY_open_html_disc;
SWEFHTML.BODY_close = BODY_close_html_disc;
SWEFHTML.BR = BR_html_disc;
SWEFHTML.DIV_open = DIV_open_html_disc;
SWEFHTML.DIV_close = DIV_close_html_disc;
SWEFHTML.DOUBLE_QUOTES = DOUBLE_QUOTES_html_disc;
SWEFHTML.DTD = DTD_html_disc;
SWEFHTML.FONT_open = FONT_open_html_disc;
SWEFHTML.FONT_close = FONT_close_html_disc;
SWEFHTML.FORM_open = FORM_open_html_disc;
SWEFHTML.FORM_close = FORM_close_html_disc;
SWEFHTML.HR = HR_html_disc;
SWEFHTML.HTML_open = HTML_open_html_disc;
SWEFHTML.HTML_close = HTML_close_html_disc;
SWEFHTML.IFRAME_open = IFRAME_open_html_disc;
SWEFHTML.IFRAME_close = IFRAME_close_html_disc;
SWEFHTML.IMG = IMG_html_disc;
SWEFHTML.INPUT_checkbox = INPUT_checkbox_html_disc;
SWEFHTML.INPUT_hidden = INPUT_hidden_html_disc;
SWEFHTML.INPUT_password = INPUT_password_html_disc;
SWEFHTML.INPUT_submit = INPUT_submit_html_disc;
SWEFHTML.INPUT_text = INPUT_text_html_disc;
SWEFHTML.JS_open = JS_open_html_disc;
SWEFHTML.JS_close = JS_close_html_disc;
SWEFHTML.LI = LI_html_disc;
SWEFHTML.MAILTO = MAILTO_html_disc;
SWEFHTML.NBSP = NBSP_html_disc;
SWEFHTML.NBSP_multi = NBSP_html_disc;
SWEFHTML.OL_open = OL_open_html_disc;
SWEFHTML.OL_close = OL_close_html_disc;
SWEFHTML.OPTION = OPTION_html_disc;
SWEFHTML.P_open = P_open_html_disc;
SWEFHTML.P_close = P_close_html_disc;
SWEFHTML.QUOTE_open = QUOTE_open_html_disc;
SWEFHTML.QUOTE_close = QUOTE_close_html_disc;
SWEFHTML.SCRIPT_open = SCRIPT_open_html_disc;
SWEFHTML.SCRIPT_close = SCRIPT_close_html_disc;
SWEFHTML.SELECT_open = SELECT_open_html_disc;
SWEFHTML.SELECT_close = SELECT_close_html_disc;
SWEFHTML.SPAN_open = SPAN_open_html_disc;
SWEFHTML.SPAN_close = SPAN_close_html_disc;
SWEFHTML.STRONG_open = STRONG_open_html_disc;
SWEFHTML.STYLE_open = STYLE_open_html_disc;
SWEFHTML.STYLE_close = STYLE_close_html_disc;
SWEFHTML.STRONG_close = STRONG_close_html_disc;
SWEFHTML.TABLE_open = TABLE_open_html_disc;
SWEFHTML.TABLE_close = TABLE_close_html_disc;
SWEFHTML.TD_open = TD_open_html_disc;
SWEFHTML.TD_close = TD_close_html_disc;
SWEFHTML.TEXTAREA_open = TEXTAREA_open_html_disc;
SWEFHTML.TEXTAREA_close = TEXTAREA_close_html_disc;
SWEFHTML.TH_open = TH_open_html_disc;
SWEFHTML.TH_close = TH_close_html_disc;
SWEFHTML.TR_open = TR_open_html_disc;
SWEFHTML.TR_close = TR_close_html_disc;
SWEFHTML.UL_open = UL_open_html_disc;
SWEFHTML.UL_close = UL_close_html_disc;


SWEFHTML.prototype.getTagSource = getTagSource_html_disc;
SWEFHTML.prototype.setTagSource = setTagSource_html_disc;

SWEFHTML.prototype.A_open = A_open_html_disc;
SWEFHTML.prototype.A_close = A_close_html_disc;
SWEFHTML.prototype.BASE = BASE_html_disc;
SWEFHTML.prototype.BLOCKQUOTE_open = BLOCKQUOTE_open_html_disc;
SWEFHTML.prototype.BLOCKQUOTE_close = BLOCKQUOTE_close_html_disc;
SWEFHTML.prototype.BODY_open = BODY_open_html_disc;
SWEFHTML.prototype.BODY_close = BODY_close_html_disc;
SWEFHTML.prototype.BR = BR_html_disc;
SWEFHTML.prototype.DIV_open = DIV_open_html_disc;
SWEFHTML.prototype.DIV_close = DIV_close_html_disc;
SWEFHTML.prototype.DOUBLE_QUOTES = DOUBLE_QUOTES_html_disc;
SWEFHTML.prototype.DTD = DTD_html_disc;
SWEFHTML.prototype.FONT_open = FONT_open_html_disc;
SWEFHTML.prototype.FONT_close = FONT_close_html_disc;
SWEFHTML.prototype.FORM_open = FORM_open_html_disc;
SWEFHTML.prototype.FORM_close = FORM_close_html_disc;
SWEFHTML.prototype.HR = HR_html_disc;
SWEFHTML.prototype.HTML_open = HTML_open_html_disc;
SWEFHTML.prototype.HTML_close = HTML_close_html_disc;
SWEFHTML.prototype.IFRAME_open = IFRAME_open_html_disc;
SWEFHTML.prototype.IFRAME_close = IFRAME_close_html_disc;
SWEFHTML.prototype.IMG = IMG_html_disc;
SWEFHTML.prototype.INPUT_checkbox = INPUT_checkbox_html_disc;
SWEFHTML.prototype.INPUT_hidden = INPUT_hidden_html_disc;
SWEFHTML.prototype.INPUT_password = INPUT_password_html_disc;
SWEFHTML.prototype.INPUT_submit = INPUT_submit_html_disc;
SWEFHTML.prototype.INPUT_text = INPUT_text_html_disc;
SWEFHTML.prototype.JS_open = JS_open_html_disc;
SWEFHTML.prototype.JS_close = JS_close_html_disc;
SWEFHTML.prototype.LI = LI_html_disc;
SWEFHTML.prototype.MAILTO = MAILTO_html_disc;
SWEFHTML.prototype.NBSP = NBSP_html_disc;
SWEFHTML.prototype.NBSP_multi = NBSP_html_disc;
SWEFHTML.prototype.OPTION = OPTION_html_disc;
SWEFHTML.prototype.OL_open = OL_open_html_disc;
SWEFHTML.prototype.OL_close = OL_close_html_disc;
SWEFHTML.prototype.P_open = P_open_html_disc;
SWEFHTML.prototype.P_close = P_close_html_disc;
SWEFHTML.prototype.QUOTE_open = QUOTE_open_html_disc;
SWEFHTML.prototype.QUOTE_close = QUOTE_close_html_disc;
SWEFHTML.prototype.SCRIPT_open = SCRIPT_open_html_disc;
SWEFHTML.prototype.SCRIPT_close = SCRIPT_close_html_disc;
SWEFHTML.prototype.SELECT_open = SELECT_open_html_disc;
SWEFHTML.prototype.SELECT_close = SELECT_close_html_disc;
SWEFHTML.prototype.STRONG_open = STRONG_open_html_disc;
SWEFHTML.prototype.SPAN_open = SPAN_open_html_disc;
SWEFHTML.prototype.SPAN_close = SPAN_close_html_disc;
SWEFHTML.prototype.STRONG_close = STRONG_close_html_disc;
SWEFHTML.prototype.STYLE_open = STYLE_open_html_disc;
SWEFHTML.prototype.STYLE_close = STYLE_close_html_disc;
SWEFHTML.prototype.TABLE_open = TABLE_open_html_disc;
SWEFHTML.prototype.TABLE_close = TABLE_close_html_disc;
SWEFHTML.prototype.TD_open = TD_open_html_disc;
SWEFHTML.prototype.TD_close = TD_close_html_disc;
SWEFHTML.prototype.TEXTAREA_open = TEXTAREA_open_html_disc;
SWEFHTML.prototype.TEXTAREA_close = TEXTAREA_close_html_disc;
SWEFHTML.prototype.TH_open = TH_open_html_disc;
SWEFHTML.prototype.TH_close = TH_close_html_disc;
SWEFHTML.prototype.TR_open = TR_open_html_disc;
SWEFHTML.prototype.TR_close = TR_close_html_disc;
SWEFHTML.prototype.UL_open = UL_open_html_disc;
SWEFHTML.prototype.UL_close = UL_close_html_disc;

SWEFHTML.prototype.sanitise = sanitise_html_disc;
SWEFHTML.prototype.isImage = isImage_html_disc;
SWEFHTML.prototype.isLinkOpen = isLinkOpen_html_disc;
SWEFHTML.prototype.isLinkClose = isLinkClose_html_disc;
SWEFHTML.prototype.sanitiseImage = sanitiseImage_html_disc;
SWEFHTML.prototype.sanitiseLink = sanitiseLink_html_disc;

</SCRIPT>