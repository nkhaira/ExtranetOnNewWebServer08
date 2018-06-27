<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// TEXTCONTROL OBJECT
//
// ======================================================================

function SWEFTextControl
(
 strName,
 strContents,
 strFormName
)
{
  this.setName (strName);

  if (isDefined_disc (strContents))
    {
      this.setContents (strContents);
    }
  else
    {
      this.setContents ("");
    }

  if (isDefined_disc (strFormName))
    {
      this.setFormName (strFormName);
    }
  else
    {
      this.setFormName ("SWEF");
    }

  return this;
}

// ======================================================================
//
// Interface to private member variables.
//
// ======================================================================

function getName_txt_disc
(
)
{
  return String (this._controlName);
}

function setName_txt_disc
(
 strNewName
)
{
  this._controlName = String (strNewName);
  return;
}

function getContents_txt_disc
(
)
{
  return String (this._contents);
}

function setContents_txt_disc
(
 strNewContents
)
{
  this._contents = String (strNewContents);
  return;
}

function getFormName_txt_disc
(
)
{
  return String (this._formName);
}

function setFormName_txt_disc
(
 strNewFormName
)
{
  this._formName = String (strNewFormName);
  return;
}

function getInitialised_txt_disc
(
)
{
  return Boolean (SWEFTextControl._initialised);
}

function setInitialised_txt_disc
(
 bNewInitialisationFlag
)
{
  SWEFTextControl._initialised = Boolean (bNewInitialisationFlag);
  return;
}

// ======================================================================
//
// Private and public methods.
//
// ======================================================================

function isUsed_txt_disc
(
)
{
  return SWEFTextControl.getInitialised ();
}

function getImageLoader_txt_disc
(
)
{
  var strJSout = SWEFHTML.JS_open ();

  strJSout += "var buttonImage_normal = new Image ();\n";
  strJSout += "buttonImage_normal.src = \"";
  strJSout += config.ADMINSETTING_ToolbarButtonImagePathname;
  strJSout += "\";\n";
  strJSout += "var buttonImage_up = new Image ();\n";
  strJSout += "buttonImage_up.src = \"";
  strJSout += config.ADMINSETTING_ToolbarButtonImageUpPathname;
  strJSout += "\";\n";
  strJSout += "var buttonImage_down = new Image ();\n";
  strJSout += "buttonImage_down.src = \"";
  strJSout += config.ADMINSETTING_ToolbarButtonImageDownPathname;
  strJSout += "\";\n\n";

  strJSout += SWEFHTML.JS_close ();

  return strJSout;
}

function getGenericRoutines_txt_disc
(
)
{
  var strJSout = SWEFHTML.JS_open ();
  strJSout += "function flipImage_disc (imageTagName, newImageObject)\n";
  strJSout += "{\n";
  strJSout += "\tif (document.images)\n";
  strJSout += "\t{\n";
  strJSout += "\t\tdocument [imageTagName].src = newImageObject.src;\n";
  strJSout += "\t}\n\n";
  strJSout += "\treturn;\n";
  strJSout += "}\n\n";
  strJSout += SWEFHTML.JS_close ();

  return strJSout;
}

function getControlSpecificRoutines_txt_disc
(
 strControlName,
 strFormName
)
{
  var strJSout = SWEFHTML.JS_open ();
  strJSout += "function submitRecord ()\n";
  strJSout += "{\n";
  strJSout += "\tvar text = frames." + strControlName + ".document.body.innerText;\n";
  strJSout += "\tvar html = frames." + strControlName + ".document.body.innerHTML;\n";
  strJSout += "\tif ((html == \"\") || (text == \"\"))\n";
  strJSout += "\t{\n";
  strJSout += "\t\tframes." + strControlName + ".focus();\n";
  strJSout += "\t\talert (\"The body of this message cannot be empty.\");\n";
  strJSout += "\t\treturn;\n";
  strJSout += "\t}\n";
  strJSout += "\tdocument." + strFormName + "." + strControlName + ".value = html;\n";
  strJSout += "\tdocument." + strFormName + ".submit();\n";
  strJSout += "\treturn;\n";
  strJSout += "}\n\n";

  strJSout += "function getRange_" + strControlName + " ()\n";
  strJSout += "{\n";
  strJSout += "\tframes." + strControlName + ".focus();\n";
  strJSout += "\tvar selection = frames.";
  strJSout += strControlName;
  strJSout += ".document.selection.createRange ();\n";
  strJSout += "\tif(selection.text == \"\");\n";
  strJSout += "\t{\n";
  strJSout += "\t\tselection = frames ['";
  strJSout += strControlName;
  strJSout += "'].document;\n"; 
  strJSout += "\t}\n";
  strJSout += "\treturn selection;\n";
  strJSout += "}\n\n";

  strJSout += "function setStyle_" + strControlName + " (newStyle)\n";
  strJSout += "{\n";
  strJSout += "\tvar selection = getRange_" + strControlName + " ()\n";
  strJSout += "\tselection.execCommand (newStyle);\n";
  strJSout += "\tframes." + strControlName + ".focus();\n";
  strJSout += "\treturn;\n";
  strJSout += "}\n\n";

  strJSout += "function setFont_" + strControlName + " ()\n";
  strJSout += "{\n";
  strJSout += "\tvar selectedRange = getRange_" + strControlName + " ()\n";
  strJSout += "\tselectedRange.execCommand (\"FontName\", true, document.";
  strJSout += strFormName;
  strJSout += ".FontName.options [document.";
  strJSout += strFormName;
  strJSout += ".FontName.selectedIndex].value);\n";
  strJSout += "\tselectedRange.execCommand (\"FontSize\", true, document.";
  strJSout += strFormName;
  strJSout += ".FontSize.options [document.";
  strJSout += strFormName;
  strJSout += ".FontSize.selectedIndex].value);\n";
  strJSout += "\tframes." + strControlName + ".focus();\n";
  strJSout += "\treturn;\n";
  strJSout += "}\n\n";

  strJSout += SWEFHTML.JS_close ();

  return strJSout;
}

function getImageControlHTML_txt_disc
(
 strControl,
 strName,
 strCommand,
 strTitle,
 strPathname,
 nWidth,
 nHeight
)
{
  var strHTMLout = "";
  strHTMLout += "<IMG ALT=\"";
  strHTMLout += strName;
  strHTMLout += "\" HEIGHT=\"";
  strHTMLout += nHeight;
  strHTMLout += "\" WIDTH=\"";
  strHTMLout += nWidth;
  strHTMLout += "\" NAME=\"";
  strHTMLout += strTitle;
  strHTMLout += "\" SRC=\"";
  strHTMLout += strPathname;
  strHTMLout += "\" onclick=\"flipImage_disc ('";
  strHTMLout += strTitle;
  strHTMLout += "', '";
  strHTMLout += strTitle;
  strHTMLout += "Image_down'); setStyle_";
  strHTMLout += strControl;
  strHTMLout += " ('";
  strHTMLout += strCommand;
  strHTMLout += "'); return false;\" onmouseover=\"flipImage_disc ('";
  strHTMLout += strTitle;
  strHTMLout += "', '";
  strHTMLout += strTitle;
  strHTMLout += "Image_up');\" onmouseout=\"flipImage_disc ('";
  strHTMLout += strTitle;
  strHTMLout += "', '";
  strHTMLout += strTitle;
  strHTMLout += "Image_normal');\">";

  return strHTMLout;
}

function getControlActivator_txt_disc
(
 strControlName
)
{
  var strHTMLout = "";
  strHTMLout += SWEFHTML.JS_open ();
  strHTMLout += "var editControl = frames." + strControlName + ".document;\n";
  strHTMLout += "editControl.designMode = \"On\";\n";
  if (isDefined_disc (this.getContents ()) && (this.getContents () != ""))
    {
      strHTMLout += "editControl.open (\"text/html\", \"replace\");\n";
      strHTMLout += "editControl.write (\"";
      strHTMLout += this.getContents ().javascriptSafeFormat ();
      strHTMLout += "\");\n";
      strHTMLout += "editControl.close ();\n";
    }
  strHTMLout += SWEFHTML.JS_close ();

  return strHTMLout;
}

function getToolbarButtonControl_txt_disc
(
 strControlName,
 strButtonName,
 strButtonTitle,
 strButtonAction
)
{
  var strOnMouseOver = "";
  strOnMouseOver += "flipImage_disc ('";
  strOnMouseOver += strButtonName;
  strOnMouseOver += "', buttonImage_up)";

  var strOnMouseOut = "";
  strOnMouseOut += "flipImage_disc ('";
  strOnMouseOut += strButtonName;
  strOnMouseOut += "', buttonImage_normal)";

  var strOnMouseDown = "";
  strOnMouseDown += "flipImage_disc ('";
  strOnMouseDown += strButtonName;
  strOnMouseDown += "', buttonImage_down); setStyle_";
  strOnMouseDown += strControlName;
  strOnMouseDown += " ('" + strButtonAction + "'); return false;";

  var strOnMouseUp = "";
  strOnMouseUp += "flipImage_disc ('";
  strOnMouseUp += strButtonName;
  strOnMouseUp += "', buttonImage_up); return false;";

  return SWEFHTML.IMG (config.ADMINSETTING_ToolbarButtonImagePathname,
		       strButtonTitle,
		       0,
		       config.ADMINSETTING_ToolbarButtonImageWidth,
		       config.ADMINSETTING_ToolbarButtonImageHeight,
		       strButtonName,
		       strButtonName,
		       undefined_disc,
		       strOnMouseOver,
		       strOnMouseOut,
		       strOnMouseDown,
		       strOnMouseUp);
}

function getToolbarSpacer_txt_disc
(
)
{
  return SWEFHTML.IMG (config.ADMINSETTING_ToolbarSpacerImagePathname,
		       " ",
		       0,
		       config.ADMINSETTING_ToolbarSpacerImageWidth,
		       config.ADMINSETTING_ToolbarSpacerImageHeight);
}

function getToolbarButtons_txt_disc
(
 strControlName
)
{
  var nButtonsSize = 12 * config.ADMINSETTING_ToolbarButtonImageWidth;
  var nSpacersSize = 6 * config.ADMINSETTING_ToolbarSpacerImageWidth;

  var strHTMLout = "";
  strHTMLout += SWEFHTML.TABLE_open (0,
				     nButtonsSize + nSpacersSize,
				     undefined_disc,
				     undefined_disc,
				     0,
				     0);
  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarSpacerImageWidth);
  strHTMLout += this.getToolbarSpacer ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "bold",
					      "Bold",
					      "bold");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "italic",
					      "Italic",
					      "italic");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "underline",
					      "Underline",
					      "underline");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarSpacerImageWidth);
  strHTMLout += this.getToolbarSpacer ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "justifyleft",
					      "Left",
					      "JustifyLeft");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "justifycenter",
					      "Center",
					      "JustifyCenter");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "justifyright",
					      "Right",
					      "JustifyRight");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "justifyfull",
					      "Justify",
					      "JustifyFull");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarSpacerImageWidth);
  strHTMLout += this.getToolbarSpacer ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "bulleted",
					      "Bulleted List",
					      "InsertUnorderedList");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "numbered",
					      "Numbered List",
					      "InsertOrderedList");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarSpacerImageWidth);
  strHTMLout += this.getToolbarSpacer ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "indent",
					      "Indent",
					      "Indent");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "outdent",
					      "Outdent",
					      "Outdent");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarSpacerImageWidth);
  strHTMLout += this.getToolbarSpacer ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarButtonImageWidth);
  strHTMLout += this.getToolbarButtonControl (strControlName,
					      "link",
					      "Create/Edit Hyperlink",
					      "CreateLink");
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarSpacerImageWidth);
  strHTMLout += this.getToolbarSpacer ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();
  strHTMLout += SWEFHTML.TABLE_close ();

  return strHTMLout;
}

function getControlHTML_txt_disc
(
 strControlName
)
{
  var nControlWidth = 9 * config.ADMINSETTING_TextAreaCols;
  var nControlHeight = 20 * config.ADMINSETTING_TextAreaRows;
  var strHTMLout = "";
  strHTMLout += SWEFHTML.TABLE_open (0,
				     nControlWidth,
				     config.ADMINSETTING_ToolbarBackgroundColour,
				     undefined_disc,
				     0,
				     0);
  strHTMLout += SWEFHTML.TR_open ();
  strHTMLout += SWEFHTML.TD_open (config.ADMINSETTING_ToolbarTotalWidth,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  undefined_disc,
				  config.ADMINSETTING_ToolbarImagePathname);
  strHTMLout += this.getToolbarButtons (strControlName);
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TD_open (nControlWidth - config.ADMINSETTING_ToolbarTotalWidth);
  strHTMLout += SWEFHTML.SELECT_open ("FontName", "setFont_" + strControlName + " (); return false;");
  strHTMLout += SWEFHTML.OPTION ("Arial, Sans Serif", true) + "Arial";
  strHTMLout += SWEFHTML.OPTION ("Times New Roman, Arial, Sans Serif") + "Times";
  strHTMLout += SWEFHTML.OPTION ("Courier, Century") + "Courier";
  strHTMLout += SWEFHTML.SELECT_close ();
  strHTMLout += SWEFHTML.SELECT_open ("FontSize", "setFont_" + strControlName + " (); return false;");
  strHTMLout += SWEFHTML.OPTION ("1") + "8pt";
  strHTMLout += SWEFHTML.OPTION ("2", true) + "10pt";
  strHTMLout += SWEFHTML.OPTION ("3") + "12pt";
  strHTMLout += SWEFHTML.OPTION ("4") + "14pt";
  strHTMLout += SWEFHTML.OPTION ("5") + "18pt";
  strHTMLout += SWEFHTML.OPTION ("6") + "24pt";
  strHTMLout += SWEFHTML.OPTION ("7") + "36pt";
  strHTMLout += SWEFHTML.SELECT_close ();
  strHTMLout += SWEFHTML.TD_close ();
  strHTMLout += SWEFHTML.TR_close ();
  strHTMLout += SWEFHTML.TABLE_close ();
  strHTMLout += SWEFHTML.IFRAME_open (strControlName,
				      "about:blank",
				      strControlName,
				      nControlWidth,
				      nControlHeight);
  strHTMLout += SWEFHTML.IFRAME_close ();
  strHTMLout += SWEFHTML.BR ();
  strHTMLout += SWEFHTML.INPUT_hidden (strControlName, "");

  return strHTMLout;
}

function isRichTextCapable_txt_disc
(
)
{
  var bClientCanHandleIt;
  if (((String (Request.ServerVariables ("HTTP_USER_AGENT")).indexOf ("MSIE 4") > 0)
       || (String (Request.ServerVariables ("HTTP_USER_AGENT")).indexOf ("MSIE 5") > 0))
      && (String (Request.ServerVariables ("HTTP_USER_AGENT")).indexOf ("Windows") > 0))
    {
      bClientCanHandleIt = true;
    }
  else
    {
      bClientCanHandleIt = false;
    }

  return bClientCanHandleIt;
}

function getControl_txt_disc
(
)
{
  var strHTMLout = "";
  if ((this.isRichTextCapable ()) && (config.ADMINSWITCH_AllowRichFormatting))
    {
      SWEFTextControl.setInitialised (true);
      strHTMLout += this.getImageLoader ();
      strHTMLout += this.getGenericRoutines ();
      strHTMLout += this.getControlSpecificRoutines (this.getName (), this.getFormName ());
      strHTMLout += this.getControlHTML (this.getName ());
      strHTMLout += this.getControlActivator (this.getName ());
    }
  else
    {
      SWEFTextControl.setInitialised (false);
      strHTMLout += SWEFHTML.TEXTAREA_open (this.getName (),
					    config.ADMINSETTING_TextAreaCols,
					    config.ADMINSETTING_TextAreaRows);
      strHTMLout += this.getContents ().unformatFromStoring ();
      strHTMLout += SWEFHTML.TEXTAREA_close ();
    }

  return strHTMLout;
}

function getControlContents_txt_disc
(
 objFormData,
 strControlName
)
{
  return objFormData (strControlName);
}

SWEFTextControl._initialised = false;
SWEFTextControl.isUsed = isUsed_txt_disc;
SWEFTextControl.getInitialised = getInitialised_txt_disc;
SWEFTextControl.setInitialised = setInitialised_txt_disc;
SWEFTextControl.isRichTextCapable = isRichTextCapable_txt_disc;
SWEFTextControl.getControlContents = getControlContents_txt_disc;

SWEFTextControl.prototype.getName = getName_txt_disc;
SWEFTextControl.prototype.setName = setName_txt_disc;
SWEFTextControl.prototype.getContents = getContents_txt_disc;
SWEFTextControl.prototype.setContents = setContents_txt_disc;
SWEFTextControl.prototype.getFormName = getFormName_txt_disc;
SWEFTextControl.prototype.setFormName = setFormName_txt_disc;

SWEFTextControl.prototype.isUsed = isUsed_txt_disc;
SWEFTextControl.prototype.getInitialised = getInitialised_txt_disc;
SWEFTextControl.prototype.setInitialised = setInitialised_txt_disc;
SWEFTextControl.prototype.getImageLoader = getImageLoader_txt_disc;
SWEFTextControl.prototype.getGenericRoutines = getGenericRoutines_txt_disc;
SWEFTextControl.prototype.getControlSpecificRoutines = getControlSpecificRoutines_txt_disc;
SWEFTextControl.prototype.getImageControlHTML = getImageControlHTML_txt_disc;
SWEFTextControl.prototype.getControlActivator = getControlActivator_txt_disc;
SWEFTextControl.prototype.getToolbarButtonControl = getToolbarButtonControl_txt_disc;
SWEFTextControl.prototype.getToolbarSpacer = getToolbarSpacer_txt_disc;
SWEFTextControl.prototype.getToolbarButtons = getToolbarButtons_txt_disc;
SWEFTextControl.prototype.getControlHTML = getControlHTML_txt_disc;
SWEFTextControl.prototype.getControl = getControl_txt_disc;
SWEFTextControl.prototype.isRichTextCapable = isRichTextCapable_txt_disc;
SWEFTextControl.prototype.getControlContents = getControlContents_txt_disc;
</SCRIPT>