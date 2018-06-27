<html>
<head>
	<title>SiteWide Rich Text Editor</title>
</head>
<body bgcolor="#ffffff">

<% 'Include the DevEdit class file %>
<!--#INCLUDE virtual="/include/RTEditor/class.devedit.asp" -->

<form action="detest_multi.asp" method="post">
<%
	
	' This sample page shows DevEdit with 3 controls on
	' the one page
	
	' Create DevEdit class #1
	dim myDE1
	set myDE1 = new DevEdit
	myDE1.SetName "myDevEditControl1"
	myDE1.SetValue myDE1.GetValue(false)
  
  myDE1.SetLanguage(de_AMERICAN)
  myDE1.SetFontList("Arial,Verdana")
  myDE1.SetFontSizeList("1,2,3")
  myDE1.SetImageDisplayType de_IMAGE_TYPE_THUMBNAIL
  myDE1.SetDocumentType de_DOC_TYPE_SNIPPET
  'myDE1.SetPathType de_PATH_ABSOLUTE
  
  'myDE1.DisableXHTMLFormatting
  'myDE1.DisableSingleLineReturn

  'myDE1.DisableSourceMode
  'myDE1.DisablePreviewMode

  'myDE1.HideFullScreenButton
  'myDE1.HideSpellingButton
  'myDE1.HideRemoveTextFormattingButton
  'myDE1.HideBoldButton
  'myDE1.HideUnderlineButton
  'myDE1.HideItalicButton
  'myDE1.HideStrikethroughButton
  
  'myDE1.HideNumberListButton
  'myDE1.HideBulletListButton
  
  'myDE1.HideDecreaseIndentButton
  'myDE1.HideIncreaseIndentButton
  
  'myDE1.HideLeftAlignButton
  'myDE1.HideCenterAlignButton
  'myDE1.HideRightAlignButton
  'myDE1.HideJustifyButton
  
  'myDE1.HideSuperScriptButton
  'myDE1.HideSubScriptButton
  
  'myDE1.HideTextBoxButton
  
  'myDE1.HideHorizontalRuleButton
  
  'myDE1.HideLinkButton
  'myDE1.HideMailLinkButton
  'myDE1.HideAnchorButton
  
  'myDE1.HideHelpButton
  
  'myDE1.HideFontList
  'myDE1.HideSizeList
  'myDE1.HideFormatList
  'myDE1.HideStyleList
  
  'myDE1.HideForeColorButton
  'myDE1.HideBackColorButton

  'myDE1.HideTableButton
  'myDE1.HideFormButton

  'myDE1.HideImageButton
  'myDE1.DisableImageDeleting
  'myDE1.DisableImageUploading

  'myDE1.HideSymbolButton
  'myDE1.HidePropertiesButton
  'myDE1.HideCleanHTMLButton
  'myDE1.HideAbsolutePositionButton  
  'myDE1.HideGuidelinesButton
  'myDE1.EnableGuidelines
  
  'myDE1.HideCopyButton
  'myDE1.HideCutButton
  'myDE1.HidePasteButton
  'myDE1.HideFindButton
  'myDE1.HideUndoButton
  'myDE1.HideRedoButton
  
  
  myDE1.SetTextAreaDimensions 30,30
  
	myDE1.ShowControl "50%", "50%", "/find-sales/download/thumbnail"

  
if 1=2 then  
	' Create DevEdit class #2
	dim myDE2
	set myDE2 = new DevEdit
	myDE2.SetName "myDevEditControl2"
	myDE2.SetValue myDE2.GetValue(false)
	myDE2.ShowControl "500", "400", "/images"

	' Create DevEdit class #3
	dim myDE3
	set myDE3 = new DevEdit
	myDE3.SetName "myDevEditControl3"
	myDE3.SetValue myDE3.GetValue(false)
	myDE3.ShowControl "100%", "100%", "/images"
end if
	
	'Display the rest of the form
	%>
		<br><br>
		<input type="submit" value="Get HTML >>"><br><br>
		Value from DevEdit control #1:<br>
		<textarea cols="100" rows="10"><%=myDE1.GetValue(false) %></textarea>
<%if 1=2 then%>    
		<hr>
		Value from DevEdit control #2:<br>
		<textarea cols="100" rows="10"><%=myDE2.GetValue(false) %></textarea>
		<hr>
		Value from DevEdit control #3:<br>
		<textarea cols="100" rows="10"><%=myDE3.GetValue(false) %></textarea>
		<hr>
<%end if%>
    <input type=submit name=submit value=Submit>    
	</form>
</body>
</html>