<%

'/***********************************************************\
'|                                                            |
'|  DevEdit v4.0.4 Copyright Interspire Pty Ltd 2003		  |
'|  All rights reserved. Do NOT modify this file. If		  |
'|  you attempt to do so then we canot provide support. This  |
'|  or any other DevEdit files may NOT be shared or			  |
'|  distributed in any way. To purchase more licences, please |
'|  visit www.devedit.com                                     |
'|                                                            |
'\************************************************************/

	' Check if HTTPS is enabled
	Dim HTTPStr

	if UCase(Request.ServerVariables("HTTPS")) = "ON" then
		HTTPStr = "https"
	else
		HTTPStr = "http"
	end if
%>

<!--#include file="de_lang/language.asp"-->
<!-- #INCLUDE file="class.fileupload.asp" -->
<%

	'Constant variables to make function calling more logical
	const de_PATH_TYPE_FULL = 0
	const de_PATH_TYPE_ABSOLUTE = 1
	const de_DOC_TYPE_SNIPPET = 0
	const de_DOC_TYPE_HTML_PAGE = 1
	const de_IMAGE_TYPE_ROW = 0
	const de_IMAGE_TYPE_THUMBNAIL = 1

	'Language variations
	const de_AMERICAN = 1
	const de_BRITISH = 2
	const de_CANADIAN = 3

	private sub DisplayIncludes (file, errorMsg)
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8 
		dim fso, f, ts, fileContent, includeFile
		dim URL, scriptName, serverName, scriptDir, slashPos, oMatch
		set fso = server.CreateObject("Scripting.FileSystemObject") 

		if file = "enc_functions.js" then
			includeFile = Server.mapPath("de/de_includes/" & file)
		else
			includeFile = Server.mapPath("de_includes/" & file)
		End if

		if (fso.FileExists(includeFile)=true) Then
				
			set f = fso.GetFile(includeFile)
			set ts = f.OpenAsTextStream(ForReading, -2) 
				
			Do While not ts.AtEndOfStream
				fileContent = fileContent & ts.ReadLine & vbCrLf
			Loop

			URL = Request.ServerVariables("http_host")
			scriptName = "de/class.devedit.asp"
			

			'Workout the location of class.devedit.asp
			scriptDir = strreverse(Request.ServerVariables("path_info"))
			slashPos = instr(1, scriptDir, "/")
			scriptDir = strreverse(mid(scriptDir, slashPos, len(scriptDir)))

			scriptName = scriptDir & scriptName
				
			fileContent = replace(fileContent,"$URL",URL)
			fileContent = replace(fileContent,"$SCRIPTNAME",ScriptName)
			fileContent = replace(fileContent,"$HTTPStr",HTTPStr)
				
			Dim re
			Set re = New RegExp
			re.global = true

			re.Pattern = "\[sTxt(\w*)\]"

			For Each oMatch in re.Execute(fileContent)
			 	fileContent = replace(fileContent,oMatch,eval("sTxt" & oMatch.SubMatches(0)))
			Next

			response.write(fileContent)

		else
			response.write("file not found:" & file)
		End if
		
	End Sub

	' Examine the value of the ToDo argument and proceed to correct sub
	dim ToDo
	ToDo = Request("ToDo")

	if Request.QueryString("ToDo") = "" then
	%>
	<link rel="stylesheet" href="de/de_includes/de_styles.css" type="text/css">
	<% else %>
	<link rel="stylesheet" href="de_includes/de_styles.css" type="text/css">
	<% end if

	if ToDo = "InsertImage" Then
		' pass to insert image screen
		%><!-- #INCLUDE file="de_includes/insert_image.asp" --><%
	elseif ToDo = "DeleteImage" Then
		%><!-- #INCLUDE file="de_includes/insert_image.asp" --><%
	elseif ToDo = "UploadImage" Then
		%><!-- #INCLUDE file="de_includes/insert_image.asp" --><%
	elseif ToDo = "FindReplace" Then
		DisplayIncludes "find_replace.inc","Find and Replace"
	elseif ToDo = "SpellCheck" Then
		DisplayIncludes "spell_check.inc","Spell Check"
	elseif ToDo = "DoSpell" Then
		DisplayIncludes "do_spell.inc","Spell Check"
	elseif ToDo = "InsertTable" Then
		DisplayIncludes "insert_table.inc","Insert Table"
	elseif ToDo = "ModifyTable" Then
		DisplayIncludes "modify_table.inc","Modify Table"
	elseif ToDo = "ModifyCell" Then
		DisplayIncludes "modify_cell.inc","Modify Cell"
	elseif ToDo = "ModifyImage" Then
		DisplayIncludes "modify_image.inc","Modify Image"
	elseif ToDo = "InsertForm" Then
		DisplayIncludes "insert_form.inc","Insert Form"
	elseif ToDo = "ModifyForm" Then
		DisplayIncludes "modify_form.inc","Modify Form"
	elseif ToDo = "InsertTextField" Then
		DisplayIncludes "insert_textfield.inc","Insert Text Field"
	elseif ToDo = "ModifyTextField" Then
		DisplayIncludes "modify_textfield.inc","Modify Text Field"
	elseif ToDo = "InsertTextArea" Then
		DisplayIncludes "insert_textarea.inc","Insert Text Area"
	elseif ToDo = "ModifyTextArea" Then
		DisplayIncludes "modify_textarea.inc","Modify Text Area"
	elseif ToDo = "InsertHidden" Then
		DisplayIncludes "insert_hidden.inc","Insert Hidden Field"
	elseif ToDo = "ModifyHidden" Then
		DisplayIncludes "modify_hidden.inc","Modify Hidden Field"
	elseif ToDo = "InsertButton" Then
		DisplayIncludes "insert_button.inc","Insert Button"
	elseif ToDo = "ModifyButton" Then
		DisplayIncludes "modify_button.inc","Modify Button"
	elseif ToDo = "InsertCheckbox" Then
		DisplayIncludes "insert_checkbox.inc","Insert Checkbox"
	elseif ToDo = "ModifyCheckbox" Then
		DisplayIncludes "modify_checkbox.inc","Modify CheckBox"
	elseif ToDo = "InsertRadio" Then
		DisplayIncludes "insert_radio.inc","Insert Radio"
	elseif ToDo = "ModifyRadio" Then
		DisplayIncludes "modify_radio.inc","Modify Radio"
	elseif ToDo = "PageProperties" Then
		DisplayIncludes "page_properties.inc","Page Properties"
	elseif ToDo = "InsertLink" Then
		DisplayIncludes "insert_link.inc","Insert HyperLink"
	elseif ToDo = "InsertEmail" Then
		DisplayIncludes "insert_email.inc","Insert Email Link"
	elseif ToDo = "InsertAnchor" Then
		DisplayIncludes "insert_anchor.inc","Insert Email Link"
	elseif ToDo = "ModifyAnchor" Then
		DisplayIncludes "modify_anchor.inc","Insert Email Link"
	elseif ToDo = "CustomInsert" Then
		DisplayIncludes "custom_insert.inc","Insert Custom HTML"
	elseif ToDo = "ShowHelp" Then
		DisplayIncludes "help.inc","Help"
	End if

	class devedit
	
		private e__controlName
		private e__controlWidth
		private e__controlHeight
		private e__initialValue
		private e__langPack
		private e__hideSpelling
		private e__hideRemoveTextFormatting
		private e__hideFullScreen
		private e__hideBold
		private e__hideUnderline
		private e__hideItalic
		private e__hideStrikethrough
		private e__hideNumberList
		private e__hideBulletList
		private e__hideDecreaseIndent
		private e__hideIncreaseIndent
		private e__hideSuperScript
		private e__hideSubScript
		private e__hideLeftAlign
		private e__hideCenterAlign
		private e__hideRightAlign
		private e__hideJustify
		private e__hideHorizontalRule
		private e__hideLink
		private e__hideAnchor
		private e__hideMailLink
		private e__hideHelp
		private e__hideFont
		private e__hideSize
		private e__hideFormat
		private e__hideStyle
		private e__hideForeColor
		private e__hideBackColor
		private e__hideTable
		private e__hideForm
		private e__hideImage
		private e__hideTextBox
		private e__hideSymbols
		private e__hideProps
		private e__hideClean
		private e__hideWord
		private e__hideAbsolute
		private e__hideGuidelines
		private e__disableSourceMode
		private e__disablePreviewMode
		private e__guidelinesOnByDefault
		private e__imagePathType
		private e__docType
		private e__imageDisplayType
		private e__disableImageUploading
		private e__disableImageDeleting
		private e__enableXHTMLSupport
		private e__useSingleLineReturn
		private e__customInsertArray
		private e__hasCustomInserts
		private e__snippetCSS
		private e__textareaCols
		private e__textareaRows
		private e__fontNameList
		private e__fontSizeList
		private e__hideWebImage
		private e__language
    
    private e__hideCopy
    private e__hideCut
    private e__hidePaste
    private e__hideFind
    private e__hideUndo
    private e__hideRedo
		
		'Keep track of how many buttons are hidden in the top row.
		'If they are all hidden, then we dont show that row of the menu.
		private e__numTopHidden
		private e__numBottomHidden

		public sub Class_Initialize()

			'Set the default value of all private variables for the class
			 e__controlName = ""
			 e__controlWidth = 0
			 e__controlHeight = 0
			 e__initialValue = ""
			 e__langPack = 0
			 e__hideSpelling = 0
			 e__hideRemoveTextFormatting = 0
			 e__hideFullScreen = 0
			 e__hideBold = 0
			 e__hideUnderline = 0
			 e__hideItalic = 0
			 e__hideStrikethrough = 0
			 e__hideNumberList = 0
			 e__hideBulletList = 0
			 e__hideDecreaseIndent = 0
			 e__hideIncreaseIndent = 0
			 e__hideSuperScript = 0
			 e__hideSubScript = 0
			 e__hideLeftAlign = 0
			 e__hideCenterAlign = 0
			 e__hideRightAlign = 0
			 e__hideJustify = 0
			 e__hideHorizontalRule = 0
			 e__hideLink = 0
			 e__hideAnchor = 0
			 e__hideMailLink = 0
			 e__hideHelp = 0
			 e__hideFont = 0
			 e__hideSize = 0
			 e__hideFormat = 0
			 e__hideStyle = 0
			 e__hideForeColor = 0
			 e__hideBackColor = 0
			 e__hideTable = 0
			 e__hideForm = 0
			 e__hideImage = 0
			 e__hideTextBox = 0
			 e__hideSymbols = 0
			 e__hideProps = 0
			 e__hideWord = 0
			 e__hideClean = 0
			 e__hideAbsolute = 0
			 e__hideGuidelines = 0
			 e__disableSourceMode = 0
			 e__disablePreviewMode = 0
			 e__guidelinesOnByDefault = 0
 			 e__numTopHidden = 0
			 e__numBottomHidden = 0
			 e__imagePathType = 0
			 e__docType = 0
			 e__imageDisplayType = 0
			 e__disableImageUploading = 0
			 e__disableImageDeleting = 0
			 e__enableXHTMLSupport = 1
			 e__useSingleLineReturn = 1
			 set e__customInsertArray = Server.CreateObject("Scripting.Dictionary")
			 e__hasCustomInserts = false
			 e__snippetCSS = ""
			 e__textareaCols = 30
			 e__textareaRows = 10
			 e__fontNameList = array()
			 e__fontSizeList = array()
			 e__hideWebImage = 0
			 e__language = de_AMERICAN
       
       e__hideCopy = 0
       e__hideCut = 0
       e__hidePaste = 0
       e__hideFind = 0
       e__hideUndo = 0
       e__hideRedo = 0

		end sub

		public sub SetName(CtrlName)

			e__controlName = CtrlName
		
		end sub

		public sub SetWidth(Width)
			e__controlWidth = Width
		end sub
		
		public sub SetHeight(Height)
			e__controlHeight = Height
		end sub
		
		public sub SetValue(HTMLValue)

			if e__docType = de_DOC_TYPE_SNIPPET and e__snippetCSS <> "" then
				HTMLValue = "<link rel='stylesheet' type='text/css' href='" & e__snippetCSS & "'>" & HTMLValue
			end if
			
			'Format the initial text so that we can set the content of the iFrame to its value
			e__initialValue = HTMLValue

			if e__initialValue <> "" then

				if isIE55OrAbove = true then
					e__initialValue = HTMLValue
					e__initialValue = replace(e__initialValue, "\", "\\")
       				e__initialValue = replace(e__initialValue, "'", "\'")
       				e__initialValue = replace(e__initialValue, chr(13), "")
       				e__initialValue = replace(e__initialValue, chr(10), "")
				else
					e__initialValue = HTMLValue
				end if

			end if

		end sub

		public function GetValue(ConvertQuotes)
		
			dim tmpVal
			
			tmpVal = Request.Form(e__controlName & "_html")

			if ConvertQuotes = true then
				tmpVal = Replace(tmpVal, "'", "''")
				tmpVal = Replace(tmpVal, """", """""")
			end if
			
			GetValue = tmpVal
		
		end function

		public sub HideCopyButton()

			' Hide the copy button
			e__hideCopy = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideCutButton()

			' Hide the cut button
			e__hideCut = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HidePasteButton()

			' Hide the paste button
			e__hidePaste = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideFindButton()

			' Hide the find button
			e__hideFind = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideUndoButton()

			' Hide the undo button
			e__hideUndo = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
    
		public sub HideRedoButton()

			' Hide the redo button
			e__hideRedo = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub


    
    
		public sub HideSpellingButton()

			' Hide the spelling button
			e__hideSpelling = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideRemoveTextFormattingButton()

			' Hide the remove text formatting button
			e__hideRemoveTextFormatting = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideFullScreenButton()
		
			'Hide the full screen button
			e__hideFullScreen = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideBoldButton()
		
			'Hide the bold button
			e__hideBold = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideUnderlineButton()
		
			'Hide the underline button
			e__hideUnderline = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideItalicButton()
		
			'Hide the italic button
			e__hideItalic = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideStrikethroughButton()

			'Hide the strikethrough button
			e__hideStrikethrough = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideNumberListButton()
		
			'Hide the number list button
			e__hideNumberList = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideBulletListButton()
		
			'Hide the bullet list button
			e__hideBulletList = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideDecreaseIndentButton()
		
			'Hide the decrease indent button
			e__hideDecreaseIndent = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideIncreaseIndentButton()
		
			'Hide the increase indent button
			e__hideIncreaseIndent = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideSuperScriptButton()
		
			'Hide the superscript button
			e__hideSuperScript = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideSubScriptButton()
		
			'Hide the subscript button
			e__hideSubScript = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideLeftAlignButton()
		
			'Hide the left align button
			e__hideLeftAlign = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideCenterAlignButton()
		
			'Hide the center align button
			e__hideCenterAlign = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideRightAlignButton()
		
			'Hide the right align button
			e__hideRightAlign = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideJustifyButton()
		
			'Hide the left align button
			e__hideJustify = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideHorizontalRuleButton()
		
			'Hide the horizontal rule button
			e__hideHorizontalRule = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideLinkButton()
		
			'Hide the link button
			e__hideLink = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideAnchorButton()
		
			'Hide the anchor button
			e__hideAnchor = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideMailLinkButton()
		
			'Hide the mail link button
			e__hideMailLink = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideHelpButton()
		
			'Hide the help button
			e__hideHelp = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideFontList()
		
			'Hide the font list
			e__hideFont = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideSizeList()
		
			'Hide the size list
			e__hideSize = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub
		
		public sub HideFormatList()
		
			'Hide the format list
			e__hideFormat = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideStyleList()
		
			'Hide the style list
			e__hideStyle = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideForeColorButton()
		
			'Hide the forecolor button
			e__hideForeColor = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub
		
		public sub HideBackColorButton()
		
			'Hide the backcolor button
			e__hideBackColor = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub
		
		public sub HideTableButton()
		
			'Hide the table button
			e__hideTable = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub
		
		public sub HideFormButton()
		
			'Hide the form button
			e__hideForm = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideImageButton()
		
			'Hide the image button
			e__hideImage = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideTextBoxButton()

			'Hide the textbox button
			e__hideTextBox = true
			e__numBottomHidden = e__numBottomHidden + 1

		end sub

		public sub HideSymbolButton()
		
			'Hide the symbol button
			e__hideSymbols = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HidePropertiesButton()
		
			'Hide the properties button
			e__hideProps = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideCleanHTMLButton()
		
			'Hide the clean HTML button
			e__hideClean = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideAbsolutePositionButton()
		
			'Hide the absolute position button
			e__hideAbsolute = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideGuidelinesButton()
		
			'Hide the guidelines button
			e__hideGuidelines = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub
		
		public sub DisableSourceMode()
		
			'Hide the source mode button
			e__disableSourceMode = true
		
		end sub
		
		public sub DisablePreviewMode()
		
			'Hide the preview mode button
			e__disablePreviewMode = true
		
		end sub
		
		public sub EnableGuidelines()

			'Set the table guidelines on by default
			e__guidelinesOnByDefault = 1

		end sub

		public sub SetPathType(PathType)

			'How do we want to include the path to the images? 0 = Full, 1 = Absolute
			e__imagePathType = PathType

		end sub

		public sub SetDocumentType(DocType)

			'Is the user editing a full HTML document
			e__docType = DocType

		end sub

		public sub SetImageDisplayType(DisplayType)

			'How should the images be displayed in the image manager? 0 = Line / 1 = Thumbnails
			e__imageDisplayType = DisplayType
		
		end sub

		public sub DisableImageUploading()

			'Do we need to stop images being uploaded?
			e__disableImageUploading = 1

		end sub

		public sub DisableImageDeleting()

			'Do we need to stop images from being delete?
			e__disableImageDeleting = 1
		
		end sub

		public function isIE55OrAbove()

			' Is it MSIE?
			dim browserCheck1, browserCheck2, browserCheck3
			browserCheck1 = instr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE")

			if browserCheck1 > 0 then
				browserCheck1 = true
			else
				browserCheck1 = false
			end if

			' Is it NOT Opera?
			browserCheck2 = instr(1, Request.ServerVariables("HTTP_USER_AGENT"), "Opera")

			if browserCheck2 = 0 then
				browserCheck2 = true
			else
				browserCheck2 = false
			end if

			if browserCheck1 = true AND browserCheck2 = true then
				isIE55OrAbove = true
			else
				isIE55OrAbove = false
			end if

		end function

		' -------------------------
		' Version 3.0 new functions

		public function DisableXHTMLFormatting()

			' Disable XHTML formatting of inline code
			e__enableXHTMLSupport = 0
		
		end function
		
		public function DisableSingleLineReturn()
		
			' Instead of adding a <p> tag for a new line, add <br> instead
			e__useSingleLineReturn = 0
		
		end function

		public function LoadHTMLFromAccessQuery(ByVal DatabaseFile, ByVal DatabaseQuery, ByRef ErrorDesc)
		
			' Grabs a value from an Access database based on a SELECT query.
			' It will return a text value from the field on success, or false on failure
			
			Err.Clear
			On Error Resume Next
			
			dim aConn
			dim aRS
			
			Set aConn =  Server.CreateObject("ADODB.Connection") 
			aConn.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & DatabaseFile & ";"
			
			if Err.number <> 0 then
				ErrorDesc = Err.Description
				LoadHTMLFromAccessQuery = false
				exit function
			else
				Set aRS = aConn.Execute(DatabaseQuery)
				
				if Err.number <> 0 then
					ErrorDesc = Err.Description
					LoadHTMLFromAccessQuery = false
					exit function
				else
					if aRS.EOF then
						ErrorDesc = Err.Description
						LoadHTMLFromAccessQuery = false
						exit function
					else
						'We have a field value, set it and return
						SetValue aRS.Fields(0).value
						LoadHTMLFromAccessQuery = true
						exit function
					end if
				end if
			end if
			
		end function

		public function LoadHTMLFromSQLServerQuery(ByVal DatabaseServer, ByVal DatabaseName, ByVal DatabaseUser, ByVal DatabasePassword, ByVal DatabaseQuery, ByRef ErrorDesc)
		
			' Grabs a value from an SQL Server database based on a SELECT query.
			' It will return a text value from the field on success, or false on failure
			
			'Err.Clear
			'On Error Resume Next
			
			dim aConn
			dim aRS
			
			Set aConn =  Server.CreateObject("ADODB.Connection") 
			aConn.Open "Provider=SQLOLEDB; Data Source = " & DatabaseServer & "; Initial Catalog = " & DatabaseName & "; User Id = " & DatabaseUser & "; Password=" & DatabasePassword
			
			if Err.number <> 0 then
				ErrorDesc = Err.Description
				LoadHTMLFromSQLServerQuery = false
				exit function
			else
				Set aRS = aConn.Execute(DatabaseQuery)
				
				if Err.number <> 0 then
					ErrorDesc = Err.Description
					LoadHTMLFromSQLServerQuery = false
					exit function
				else
					if aRS.EOF then
						ErrorDesc = Err.Description
						LoadHTMLFromSQLServerQuery = false
						exit function
					else
						'We have a field value, set it and return
						SetValue aRS.Fields(0).value
						LoadHTMLFromSQLServerQuery = true
						exit function
					end if
				end if
			end if
			
		end function
		
		public function LoadFromFile(ByVal FilePath, ByRef ErrorDesc)
		
			' This function will load the entire text from a file and
			' set the value of the DevEdit control
			
			Err.Clear
			On Error Resume Next
			
			dim fso
			dim file
			dim fileContents
			
			set fso = Server.CreateObject("Scripting.FileSystemObject")
			FilePath = Server.mapPath(FilePath)
			
			if fso.FileExists(FilePath) then
				'File exists, open it and read it in
				set file = fso.OpenTextFile(FilePath, 1, false, -2)
				
				if err.number <> 0 then
					'An error occured while opening the file
					ErrorDesc = Err.Description
					LoadFromFile = false
					exit function
				else
					'The file was loaded OK, read it into a variable
					fileContents = file.ReadAll
					SetValue fileContents
					LoadFromFile = true
					exit function
				end if
			else
				'File doesnt exist
				ErrorDesc = "File " & FilePath & " doesn't exist"
				LoadFromFile = false
				exit function
			end if

		end function

		public function SaveToFile(ByVal FilePath, ByRef ErrorDesc)
		
			' Writes the contents of the DevEdit control to a file
			Err.Clear
			On Error Resume Next
			
			dim fso
			dim file
			dim stream
			dim fileContents
			
			set fso = Server.CreateObject("Scripting.FileSystemObject")
			
			if len(GetValue(false)) = 0 then
				' No data to write to the file
				ErrorDesc = "Cannot save an empty value to " & FilePath
				SaveToFile = false
				exit function
			else
				if NOT fso.FileExists(FilePath) then
					'Attempt to create the file
					set file = fso.CreateTextFile(FilePath)
					
					if err.number <> 0 then
						ErrorDesc = err.Description
						SaveToFile = false
						exit function
					end if
				end if
			
				' Now that the file definetly exists, we can attempt to open it for writing
				set file = fso.GetFile(FilePath)
			
				if err.number <> 0 then
					'Failed to grab the file
					ErrorDesc = err.Description
					SaveToFile = false
					exit function
				else
					'The file was opened OK, write to it
					set stream = file.OpenAsTextStream(2)
					
					if err.number <> 0 then
						'Failed to open the file for writing
						ErrorDesc = err.Description
						SaveToFile = false
						exit function
					else
						'Opened the file for writing OK
						stream.Write(GetValue(false))
						stream.Close
						SaveToFile = true
						exit function
					end if
				end if
			end if
		
		end function

		public function AddCustomInsert(InsertName, InsertHTMLCode)
		
			e__hasCustomInserts = true
			
			if e__customInsertArray.Exists(InsertName) = false then
				e__customInsertArray.Add InsertName, InsertHTMLCode
			end if

		end function
		
		private function FormatCustomInsertText()
		
			' Private Function - This function will return all of the custom inserts as JavaScript arrays
			dim i, name, html, ciText, keys, vals
			
			if e__hasCustomInserts = true then
				
				ciText = "["
				keys = e__customInsertArray.Keys
				vals = e__customInsertArray.Items
			
				for i = 0 to e__customInsertArray.Count-1
					name = replace(replace(keys(i), """", "\"""), chr(13) & chr(10), "\r\n")
					html = replace(replace(vals(i), """", "\"""), chr(13) & chr(10), "\r\n")
					ciText = ciText & "[""" & name & """, """ & html & """],"
				next
				
				ciText = left(ciText, len(ciText)-1)
				ciText = ciText & "]"
				
			else
				ciText = "[]"
			end if
			
			FormatCustomInsertText = ciText
		
		end function

		public function SetSnippetStyleSheet(StyleSheetURL)

			' Sets the location of the stylesheet for a code snippet
			e__docType = de_DOC_TYPE_SNIPPET
			e__snippetCSS = StyleSheetURL

		end function
		
		public function SetTextAreaDimensions(Cols, Rows)
		
			' Sets the rows and cols attributes of the <textarea> tag that will appear
			' if the client isnt using Internet explorer
			e__textareaCols = Cols
			e__textareaRows = Rows
		
		end function

		' End Version 3.0 new functions
		' Version 4.0 new functions

		public sub SetLanguage(Lang)

			select case(CStr(Lang))
				case "1"
					e__language = "american"
				case "2"
					e__language = "british"
				case "3"
					e__language = "canadian"
				case else
					e__language = "american"
			end select

		end sub

		public function DisableInsertImageFromWeb

			e__hideWebImage = 1

		end function

		public function BuildSizeList()

			%><option selected><%=sTxtSize%></option><%

			if uBound(e__fontSizeList) >= 0 then

				' Build the list of font sizes from the list that the user has specified
				for i = 0 to uBound(e__fontSizeList)
					%><option value="<%=trim(e__fontSizeList(i)) %>"><%=trim(e__fontSizeList(i)) %></option><%
				next

			else

				' Build the list of font sizes manually
				%>
					<option value="1">1</option>
			  		<option value="2">2</option>
			  		<option value="3">3</option>
			  		<option value="4">4</option>
			  		<option value="5">5</option>
			  		<option value="6">6</option>
			  		<option value="7">7</option>
				<%
			end if

		end function

		public function BuildFontList()

			%><option selected><%=sTxtFont%></option><%

			if uBound(e__fontNameList) >= 0 then

				' Build the list of fonts from the list that the user has specified
				for i = 0 to uBound(e__fontNameList)
					%><option value="<%=trim(e__fontNameList(i)) %>"><%=trim(e__fontNameList(i)) %></option><%
				next

			else

				' Build the list of fonts manually
				%>
					<option value="Times New Roman">Default</option>
					<option value="Arial">Arial</option>
					<option value="Verdana">Verdana</option>
					<option value="Tahoma">Tahoma</option>
					<option value="Courier New">Courier New</option>
					<option value="Georgia">Georgia</option>
				<%
			end if

		end function

		public function SetFontList(FontList)

			dim tmpFontList
			tmpFontList = split(FontList, ",")

			if isArray(tmpFontList) then
				e__fontNameList = tmpFontList
			end if

		end function

		public function SetFontSizeList(SizeList)

			dim tmpSizeList
			tmpSizeList = split(SizeList, ",")

			if isArray(tmpSizeList) then
				e__fontSizeList = tmpSizeList
			end if

		end function

		' End Version 4.0 new functions
		' -------------------------

		public sub ShowControl(Width, Height, ImagePath)
		
			SetWidth(Width)
			SetHeight(Height)

			if e__controlName = "" then
				Response.Write "<b>ERROR: Must set an DevEdit control name using the SetName() function</b>"
				Response.End
			end if

			' If the browser isn't IE5.5 or above, show a <textarea> tag and die
			if isIE55OrAbove = false then
			%>
				<span style="background-color: lightyellow"><font face="verdana" size="1" color="red"><b><%=sTxtTextAreaError%></b></font></span><br>
				<textarea style="width:<%=e__controlWidth %>; height:<%=e__controlHeight%>" rows="<%=e__textareaRows%>" cols="<%=e__textareaCols%>" name="<%=e__controlName %>_html"><%=e__initialValue%></textarea>
			<%
			else

					'Output the hidden textarea buffer tag which will contain the iFrame source
					Response.Write "<textarea style=display:none id='" & e__controlName & "_src'>"

					if Request.QueryString("ToDo") = "" then
					%>
						<link rel="stylesheet" href="de/de_includes/de_styles.css" type="text/css">
					<% else %>
						<link rel="stylesheet" href="de_includes/de_styles.css" type="text/css">
					<% end if

        			'Do we need to hide the page properties button?
        			if e__hideProps <> 0 or e__docType = 0 then
        				HidePropertiesButton
        			end if
        			
 '       			Dim re
 '       			Set re = New RegExp
 '       			re.global = true
        
 '       			re.Pattern = "\[sTxt(\w*)\]"
        
 '       			Dim oMatches, oMatch
        
 '       			Const ForReading = 1, ForWriting = 2, ForAppending = 8 
 '       			dim fso, f, ts, fileContent, includeFile
        			dim URL, scriptName, scriptDir, slashPos
        
        			' Print JSFunctions
 '       			set fso = server.CreateObject("Scripting.FileSystemObject") 
        
 '       			includeFile = Server.mapPath("de/de_includes/jsfunctions.inc")
        
 '       			if (fso.FileExists(includeFile)=true) Then
 '       				set f = fso.GetFile(includeFile)
 '       				set ts = f.OpenAsTextStream(ForReading, -2) 
 '       				Do While not ts.AtEndOfStream
 '       				 		fileContent = fileContent & ts.ReadLine & vbCrLf
 '       				Loop
        				
        				URL = Request.ServerVariables("http_host")
        				scriptName = "de/class.devedit.asp"
        				
        				'Workout the location of class.devedit.asp
        				scriptDir = strreverse(Request.ServerVariables("path_info"))
        				slashPos = instr(1, scriptDir, "/")
        				scriptDir = strreverse(mid(scriptDir, slashPos, len(scriptDir)))
        				
        				scriptName = scriptDir & scriptName

        
 '       				fileContent = replace(fileContent,"$URL", URL)
 '       				fileContent = replace(fileContent,"$SCRIPTNAME", scriptName)
 '       				fileContent = replace(fileContent,"$IMAGEDIR", Server.URLEncode(ImagePath))
 '       				fileContent = replace(fileContent,"$SHOWTHUMBNAILS", e__imageDisplayType)
 '       				fileContent = replace(fileContent,"$EDITINGHTMLDOC", e__docType)
 '       				fileContent = replace(fileContent,"$PATHTYPE", e__imagePathType)
 '       				fileContent = replace(fileContent,"$GUIDELINESDEFAULT", e__guidelinesOnByDefault)
 '       				fileContent = replace(fileContent,"$DISABLEIMAGEUPLOADING", e__disableImageUploading)
 '       				fileContent = replace(fileContent,"$DISABLEIMAGEDELETING", e__disableImageDeleting)
 '       				fileContent = replace(fileContent,"$XHTML", e__enableXHTMLSupport)
 '       				fileContent = replace(fileContent,"$USEBR", e__useSingleLineReturn)
 '       				fileContent = replace(fileContent,"$CUSTOMINSERTS", FormatCustomInsertText)

 '       				For Each oMatch in re.Execute(fileContent)
 '       				 	fileContent = replace(fileContent,oMatch,eval("sTxt" & oMatch.SubMatches(0)))
 '       				Next
        
 '       				response.write(fileContent)
 '       			else
 '       				response.write("jsfunctions.inc: file not found")
 '       				response.end
 '       			End if

					'if e__enableXHTMLSupport = 1 then
					%>
						<script language="JavaScript" src="de/de_includes/ro_attributes.js" type="text/javascript"></script>
						<script language="JavaScript" src="de/de_includes/ro_xml.js" type="text/javascript"></script>
						<script language="JavaScript" src="de/de_includes/ro_stringbuilder.js" type="text/javascript"></script>
					<%
					'end if
					%>
					<script>
							var customInserts = <%=FormatCustomInsertText%>
							var tableDefault = <%=e__guidelinesOnByDefault%>
							var useBR = <%=e__useSingleLineReturn%>
							var useXHTML = "<%=e__enableXHTMLSupport%>"
							var ContextMenuWidth = <%=sTxtContextMenuWidth%>
							var URL = "<%=URL%>"
							var ScriptName = "<%=scriptName%>"
							var sTxtGuidelines = "<%=sTxtGuidelines%>"
							var sTxtOn = "<%=sTxtOn%>"
							var sTxtOff = "<%=sTxtOff%>"
							var sTxtClean = "<%=sTxtClean%>"
							// var re2 = /href="<%=HTTPStr%>:\/\/<%=URL%>/g
							var re3 = /src="<%=HTTPStr%>:\/\/<%=URL%>/g
							var re4 = /src="<%=HTTPStr%>:\/\/<%=URL%>/g
							var re5 = /src="http:\/\/<%=URL%>/g
							var isEditingHTMLPage = <%=e__docType%>;
							var pathType = <%=e__imagePathType%>;
							var imageDir = "<%=Server.URLEncode(ImagePath)%>"
							var showThumbnails = <%=e__imageDisplayType%>;
							var disableImageUploading = <%=e__disableImageUploading%>;
							var disableImageDeleting = <%=e__disableImageDeleting%>;
							var HideWebImage = <%=e__hideWebImage%>;
							var HTTPStr = "<%=HTTPStr%>";
							var spellLang = "<%=e__language%>";
							var controlName = "<%=e__controlName%>_frame";
					</script>

					<script>
					<% DisplayIncludes "enc_functions.js", "Javascript Functions" %>
					</script>
					<script language="JavaScript" src="de/de_includes/de_functions.js" type="text/javascript"></script>

        			<table id="fooContainer" width="100%" height="100%" border="1" cellspacing="0" cellpadding="0">
        					<tr>
        						<td height=1>
        			<%
        
        			'Include the toolbar
        			%><!-- #INCLUDE file="de_includes/toolbar.asp" -->
        			
        							</td></tr>
        							<tr><td>
        							<table class=iframe height=100% width=100%>
        								<tr height=100%>
        									<td>
        										<iFrame onBlur="updateValue()" SECURITY="restricted" contenteditable HEIGHT=100% id="foo" style="width:100%;" src=''></iFrame>
        										<iframe onBlur="updateValue()" id=previewFrame height=100% style="width=100%; display:none"></iframe>
        									</td>
        								</tr>
        							</table>
        							</td></tr>
        							<tr><td height=1>
        							<table cellpadding=0 cellspacing=0 width=100% style="background-color: threedface" class=status>
        								<tr>
        									<td background=de/de_images/status_border.gif height=22><img style="cursor:hand;" id=editTab src=de/de_images/status_edit_up.gif width=98 height=22 border=0 onClick=editMe()><img style="cursor:hand; <% if e__disableSourceMode = true then %>display:none<% end if %>" id=sourceTab src=de/de_images/status_source.gif width=98 height=22 border=0 onClick=sourceMe()><img style="cursor:hand; <% if e__disablePreviewMode = true then %>display:none<% end if %>" id=previewTab src=de/de_images/status_preview.gif width=98 height=22 border=0 onClick=previewMe()></td>
        									<td background=de/de_images/status_border.gif id=statusbar align=right valign=bottom><img src=de/de_images/button_zoom.gif width=42 height=17 valign=bottom onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" class=toolbutton onClick="showMenu('zoomMenu',65,178)"></td>
        								</tr>
        							</table>
        						</td>
        					</tr>
        				</table>
        
        				<script language="JavaScript">
        							
        					var fooWidth = "<%=e__controlWidth %>";
        					var fooHeight = "<%=e__controlHeight %>";
        
        					function setValue()
        					{
        							foo.document.write('<%=Replace(Replace(e__initialValue, "</script>", "<\/script>", 1, -1, 1), "<", "&lt;") %>');
        							foo.document.close()
        					}
        
        					function updateValue()
        					{
									if (document.activeElement) {
										if (document.activeElement.parentElement.id == "de") {
											return false;
										} else {
											if (parent.document.all.<%=e__controlName%>_html != null) {
												parent.document.all.<%=e__controlName%>_html.value = SaveHTMLPage();
											}
										}
									}
        					}
        							
        				</script>
        
        			<%

					'End the iFrame source text area buffer
					%>
						</textarea>

						<iframe id="<%=e__controlName%>_frame" width="<%=e__controlWidth%>" height="<%=e__controlHeight%>" frameborder=0 scrolling=auto style="position:relative"></iframe>

						<input type="hidden" name="<%=e__controlName%>_html">

						<script language="JavaScript">
							<%=e__controlName%>_frame.document.write(document.getElementById("<%=e__controlName%>_src").value)
							<%=e__controlName%>_frame.document.close()
							<%=e__controlName%>_frame.document.body.style.margin = "0px";
						</script>

					<%
			end if
			
		end sub
		
	end class
%>

