	var imageWin
	var propWin
	var inserttableWin
	var previewWin
	var modifytableWin
	var insertFormWin
	var textFieldWin
	var hiddenWin
	var buttonWin
	var checkboxWin
	var radioWin
	var linkWin
	var emailWin
	var anchorWin
	var showHelpWin
	var customInsertWin

	var selectedTD
	var selectedTR
	var selectedTBODY
	var selectedTable
	var selectedImage
	var selectedForm
	var selectedTextField
	var selectedTextArea
	var selectedHidden
	var selectedbutton
	var selectedCheckbox
	var selectedRadio

	var controlName

	var doSave = 0
	var zoomSize = 100

	// URL of StyleSheet used when adding StyleSheet with CodeSnippet
	var myStyleSheet = ""

	var fileCache
	fileCache = 0

	var statusMode = ""
	var statusBorders = ""
	var toggle = "off"
	var borderShown = "no"
	var fooURL
	var reloaded
	var justSwitched = false
	reloaded = 0
	var colorType = 0

	window.onload = doLoad
	window.onerror = stopError

	var loaded = false
  
	function stopError() {
		return true;
	}

	function doLoad() {
		foo.document.designMode = 'On';
		setValue()

		if (tableDefault != 0) {
			toggleBorders();
		}
		fooURL = foo.location.href
		saveHistory(false)
		initFoo()
		updateValue()
		doZoom(zoomSize)
	}
	
	var stylesDisplayed = 0
	function doStyles() {
		if (foo.document.styleSheets.length > 0) {
			
			if (stylesDisplayed != 1)
			{
				displayUserStyles()
				stylesDisplayed = 1
			}

		}
	}

	function initFoo() {

		var iframes = document.all.tags("IFRAME");
		el = iframes[0];

		el.frameWindow = document.frames[el.id];

		el.frameWindow.document.oncontextmenu = function () {
			if (!el.frameWindow.event.ctrlKey){
				showContextMenu(el.frameWindow.event)
				return false;
			}
		}

		el.frameWindow.document.onerror = function () {
			return true;
		}

		el.frameWindow.document.onselectionchange = function () {
			if (doSave == 0)
			{
				doToolbar();
			}
		}

		el.frameWindow.document.onkeydown = function ()
		{
			if (el.frameWindow.event.keyCode == 13)
			{
				if (useBR)
				{

					var sel = el.frameWindow.document.selection;
					if (sel.type == "Control")
						return;

					if (!el.frameWindow.event.shiftKey)
					{

						var r = sel.createRange();	
						r.pasteHTML("<BR>");
						el.frameWindow.event.cancelBubble = true; 
						el.frameWindow.event.returnValue = false; 

						r.select();
						r.moveEnd("character", 1);
						r.moveStart("character", 1);
						r.collapse(false);
						
						return false;

					} else
					{
					
						if (isCursorInList())
						{
							var r = sel.createRange();	
							r.pasteHTML("<li>&nbsp;</li>");
							el.frameWindow.event.cancelBubble = true; 
							el.frameWindow.event.returnValue = false; 
							
							r.moveStart("character", -1);
							r.collapse(true);
							r.select();

							return false;
						} else
						{
							var r = sel.createRange();	
							r.pasteHTML("<br><br>");
							el.frameWindow.event.cancelBubble = true; 
							el.frameWindow.event.returnValue = false; 

							r.collapse(true);
							r.select();
						
							return false;
						}
					}
				}
			}
			if(el.frameWindow.event.ctrlKey) {
				if(el.frameWindow.event.keyCode == 90) {
					if (editModeOn) {
				      goHistory(-1);
					  return false;
					}
				} else if(el.frameWindow.event.keyCode == 89) {
					if (editModeOn) {
				      goHistory(1);
					  return false;
					}
				} else if(el.frameWindow.event.keyCode == 68) {
			      pasteWord();
				  return false;
				} else if(el.frameWindow.event.keyCode == 66) {
			      doCommand("bold");
				  return false;
				} else if(el.frameWindow.event.keyCode == 85) {
			      doCommand("underline");
				  return false;
				} else if(el.frameWindow.event.keyCode == 73) {
			      doCommand("italic");
				  return false;
				} else if(el.frameWindow.event.keyCode == 75) {
					if (document.getElementById("toolbarLink_on") != null) {
				      doLink();
					  return false;
					} else {
					  return false;
					}
				}
			}

			if(!el.frameWindow.event.ctrlKey && el.frameWindow.event.keyCode != 90 && el.frameWindow.event.keyCode != 89) {
				if (el.frameWindow.event.keyCode == 32 || el.frameWindow.event.keyCode == 13)
				{
					saveHistory()
				}
			}

			if (el.frameWindow.event.keyCode == 118) {
				if (document.getElementById("toolbarSpell") != null)
				{
					spellCheck();
					return false;
				}
			}
		}

		el.frameWindow.document.onkeyup = function() {
			showCutCopyPaste()
			showPosition()
			showLink()
			showUndoRedo()
		}

		foo.document.execCommand("2D-Position",false, true)
	}

	function showColor(oBox,oColor) {
		oBox.innerHTML = oColor.style.backgroundColor.toUpperCase();
		oBox.style.backgroundColor = oColor.style.backgroundColor
	}

	function doColor(oColor) {
		if (colorType == 2) {
			myCommand = 'BackColor'
		} else {
			myCommand = 'ForeColor'
		}

		foo.document.execCommand(myCommand,false,oColor.innerHTML);
		oPopup.hide()
	}

	var oPopup = window.createPopup();
	function showMenu(menu, width, height)
	{
    
	var lefter = event.clientX;
	var leftoff = event.offsetX
	var topper = event.clientY;
	var topoff = event.offsetY;
	var oPopBody = oPopup.document.body;
	moveMe = 0

	if (menu == "pasteMenu")
	{
		moveMe = 22
	}

	if (menu == "zoomMenu")
	{
		lefter = lefter-18
		topper = topper - 203
	}

	if (menu == "colorMenu") {
		colorType = "0"
	}

	if (menu == "colorMenu2") {
		colorType = "2"
		menu = "colorMenu"
	}

	if (menu == "formMenu")
	{
		if (isCursorInForm()) {
			document.getElementById("modifyForm1").disabled = false
		} else {
			document.getElementById("modifyForm1").disabled = true
		}
	}

	if (menu == "tableMenu")
	{
	
		if (isCursorInTableCell() || isTableSelected()) {
			document.getElementById("modifyTable").disabled = false
		} else {
			document.getElementById("modifyTable").disabled = true
		}

		if (isCursorInTableCell())
		{
			document.getElementById("modifyCell").disabled = false
			document.getElementById("rowAbove").disabled = false
			document.getElementById("rowBelow").disabled = false
			document.getElementById("deleteRow").disabled = false
			document.getElementById("colAfter").disabled = false
			document.getElementById("colBefore").disabled = false
			document.getElementById("deleteCol").disabled = false
			document.getElementById("increaseSpan").disabled = false
			document.getElementById("decreaseSpan").disabled = false

		} else {
			document.getElementById("modifyCell").disabled = true
			document.getElementById("rowAbove").disabled = true
			document.getElementById("rowBelow").disabled = true
			document.getElementById("deleteRow").disabled = true
			document.getElementById("colAfter").disabled = true
			document.getElementById("colBefore").disabled = true
			document.getElementById("deleteCol").disabled = true
			document.getElementById("increaseSpan").disabled = true
			document.getElementById("decreaseSpan").disabled = true

		}
	}

	var HTMLContent = eval(menu).innerHTML
	oPopBody.innerHTML = HTMLContent
	oPopup.show(lefter - leftoff - 2 - moveMe, topper - topoff + 22, width, height, document.body);

	return false;
	}

	var oPopup2 = window.createPopup();
	function showContextMenu(event)
	{
    
		menu = "contextMenu"
		width = ContextMenuWidth
		height = "67"

		var lefter = event.clientX;
		var topper = event.clientY;
	
		var oPopBody = oPopup2.document.body;

		height = parseInt(height)

		if (foo.document.queryCommandEnabled("cut"))
		{
			document.getElementById("cmCut").disabled = false
		} else {
			document.getElementById("cmCut").disabled = true
		}

		if (foo.document.queryCommandEnabled("paste"))
		{
			document.getElementById("cmPaste").disabled = false
		} else {
			document.getElementById("cmPaste").disabled = true
		}

		if (foo.document.queryCommandEnabled("copy"))
		{
			document.getElementById("cmCopy").disabled = false
		} else {
			document.getElementById("cmCopy").disabled = true
		}

		var HTMLContent = "<table style='BORDER-LEFT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-RIGHT: #404040 1px solid; BORDER-BOTTOM: #404040 1px solid;' cellpadding=0 cellspacing=0><tr><td>"
		HTMLContent = HTMLContent + eval(menu).innerHTML

		if (isImageSelected())
		{
			HTMLContent = HTMLContent + eval("cmImageMenu").innerHTML
			height = height + 23

			if (document.getElementById("toolbarLink_on") != null)
			{
			HTMLContent = HTMLContent + eval("cmLinkMenu").innerHTML
			height = height + 23
			}
			
		}

		if (isTextSelected() && (sourceModeView != true))
		{
			if (document.getElementById("toolbarLink_on") != null)
			{
			HTMLContent = HTMLContent + eval("cmLinkMenu").innerHTML
			height = height + 23
			}
		}

		if (isTableSelected() || isCursorInTableCell())
		{
			if (document.getElementById("toolbarTables") != null)
			{
				HTMLContent = HTMLContent + eval("cmTableMenu").innerHTML
				height = height + 23
			}
		}
		
		if (isCursorInTableCell())
		{
			if (document.getElementById("toolbarTables") != null)
			{

			HTMLContent = HTMLContent + eval("cmTableFunctions").innerHTML
			height = height + 199

			}
		}

		if (document.getElementById("toolbarSpell") != null)
		{
			HTMLContent = HTMLContent + eval("cmSpellMenu").innerHTML
			height = height + 23
		}

		HTMLContent = HTMLContent + "</td></tr></table>"
		oPopBody.innerHTML = HTMLContent
		
		oPopup2.show(lefter + 2,topper + 2, width, height, foo.document.body)
	}

	function doCommand(cmd) {

		if (isAllowed())
		{
			document.execCommand(cmd)
		}
		oPopup.hide()
		doToolbar()

		if (cmd == "AbsolutePosition")
		{
			foo.document.execCommand("2D-Position",false, true)
		}
	}

	function doFont(oFont) {
		if (isAllowed() && isTextSelected())
		{
			foo.document.execCommand('FontName',false,oFont)
		}
		foo.focus()
		doToolbar()
	}

	function doSize(oSize) {
		if (isAllowed() && isTextSelected())
		{
			foo.document.execCommand('FontSize',false,oSize)
		}
		foo.focus()
		doToolbar()
	}
	
	function doFormat(oFormat) {
		if (isAllowed() && isTextSelected())
		{
			foo.document.execCommand('formatBlock',false,oFormat)
		}
		foo.focus()
		doToolbar()
	}

	function doZoom(size) {

		foo.document.body.runtimeStyle.zoom = size + "%"

		document.getElementById("zoom500_").innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;500%&nbsp";
		document.getElementById("zoom200_").innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;200%&nbsp";
		document.getElementById("zoom150_").innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;150%&nbsp";
		document.getElementById("zoom100_").innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;100%&nbsp";
		document.getElementById("zoom75_").innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;75%&nbsp";
		document.getElementById("zoom50_").innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;50%&nbsp";
		document.getElementById("zoom25_").innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;25%&nbsp";
		document.getElementById("zoom10_").innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10%&nbsp";

		document.getElementById("zoom" + size + "_").innerHTML = "&nbsp;&nbsp;&nbsp;&#149;&nbsp;" + size + "%&nbsp";
		zoomSize = size
		oPopup.hide()
	}

	function doTextbox() {
		foo.focus()
		var oSel = foo.document.selection.createRange();
		oSel.pasteHTML("<table id=de_textBox style='position:absolute;'><tr><td>Text box</td></tr></table>");

		textBox = foo.document.getElementById("de_textBox")

		if (borderShown == "yes")
		{
			textBox.runtimeStyle.border = "1px dotted #BFBFBF"
			allRows = textBox.rows
				for (y=0; y < allRows.length; y++) {
			 	allCellsInRow = allRows[y].cells
					for (x=0; x < allCellsInRow.length; x++) {
							allCellsInRow[x].runtimeStyle.border = "1px dotted #BFBFBF"
					}
				}
		}

		textBox.removeAttribute("id")
	}

	function isTextSelected() {
			if (foo.document.selection.type == "Text") {
				return true;
			} else {
				return false;
			}
	}

	function insertChar(oChar) {
		var sel = foo.document.selection;
		var rng = sel.createRange();
		rng.pasteHTML(oChar.innerHTML)
		oPopup.hide()
	}

	function doCleanCode(code) {
		// removes all Class attributes on a tag eg. '<p class=asdasd>xxx</p>' returns '<p>xxx</p>'
		code = code.replace(/<([\w]+) class=([^ |>]*)([^>]*)/gi, "<$1$3")
		// removes all style attributes eg. '<tag style="asd asdfa aasdfasdf" something else>' returns '<tag something else>'
		code = code.replace(/<([\w]+) style="([^"]*)"([^>]*)/gi, "<$1$3")
		// gets rid of all xml stuff... <xml>,<\xml>,<?xml> or <\?xml>
		code = code.replace(/<\\?\??xml[^>]>/gi, "")
        // get rid of ugly colon tags <a:b> or </a:b>
		code = code.replace(/<\/?\w+:[^>]*>/gi, "")
		// removes all empty <p> tags
		code = code.replace(/<p([^>])*>(&nbsp;)*\s*<\/p>/gi,"")
		// removes all empty span tags
		code = code.replace(/<span([^>])*>(&nbsp;)*\s*<\/span>/gi,"")
		return code
	}

	function pasteWord() {
		foo.focus()
		oPopup.hide()

		var oSel = foo.document.selection.createRange()
		if(oSel.parentElement)
		{
			TempArea = document.getElementById("myTempArea")
			TempArea.focus()
			TempArea.document.execCommand("SelectAll")
			TempArea.document.execCommand("Paste")
			code = doCleanCode(TempArea.innerHTML)
			oSel.pasteHTML(doCleanCode(myTempArea.innerHTML))
			oSel.select()
		}
	}

	function isAllowed() {
		var sel
		var obj
			
			sel = foo.document.selection
			
			if (sel.type != "Control")
			{
				obj = sel.createRange().parentElement()
			} else {
				obj = sel.createRange()(0)
			}
			
			if (obj.isContentEditable) {
				foo.focus()
				return true
			} else {
				return false
			}
	}
			
	function scrollUp() {
		foo.scrollBy(0,0);
	}
	
	var editModeOn = true;
	var editModeView = true;

	function editMe() {
		if (editModeView == false)
		{
			toolbar_full.className = "bevel3";
			toolbar_code.className = "hide";
			toolbar_preview.className = "hide"

			document.all.foo.style.display = "";
			document.all.previewFrame.style.display = "none";

			if (sourceModeOn)
			{
				SwitchMode()
			}

			editModeView = true;
			sourceModeView = false;
			previewModeView = false;

			sourceModeOn = false;
			editModeOn = true;

			document.getElementById("editTab").src = "/include/RTEditor/Images/status_edit_up.gif"
			document.getElementById("sourceTab").src = "/include/RTEditor/Images/status_source.gif"
			document.getElementById("previewTab").src = "/include/RTEditor/Images/status_preview.gif"
			initFoo()
			foo.focus()
		}
	}

	var sourceModeOn = false;
	var sourceModeView = false;
	function sourceMe() {
		if (sourceModeView == false)
		{

			if (isEditingHTMLPage == 0)
			{
				if (foo.document.styleSheets.length > 0)
				{
					myStyleSheet = foo.document.styleSheets(0).href
				}
			}

			toolbar_full.className = "hide";
			toolbar_code.className = "bevel3";
			toolbar_preview.className = "hide"

			document.all.foo.style.display = "";
			document.all.previewFrame.style.display = "none";

			if (editModeOn)
			{
				SwitchMode()
			}

			sourceModeView = true;
			editModeView = false;
			previewModeView = false;

			editModeOn = false;
			sourceModeOn = true;

			document.getElementById("editTab").src = "/include/RTEditor/Images/status_edit.gif"
			document.getElementById("sourceTab").src = "/include/RTEditor/Images/status_source_up.gif"
			document.getElementById("previewTab").src = "/include/RTEditor/Images/status_preview.gif"

			// update value breaks redo / undo buffer
			updateValue()

			foo.focus()
		}
	}

	var previewModeView = false;
	function previewMe() {
		if (previewModeView == false)
		{
			toolbar_full.className = "hide";
			toolbar_code.className = "hide";
			toolbar_preview.className = "bevel3"

			document.all.foo.style.display = "none";
			document.all.previewFrame.style.display = "";

			sourceModeView = false;
			editModeView = false;
			previewModeView = true;

			if (sourceModeOn) {
				ShowPreview(1)
			} else {
				ShowPreview(0)
			}

			document.getElementById("editTab").src = "/include/RTEditor/Images/status_edit.gif"
			document.getElementById("sourceTab").src = "/include/RTEditor/Images/status_source.gif"
			document.getElementById("previewTab").src = "/include/RTEditor/Images/status_preview_up.gif"
			previewFrame.focus()
		}
	}

		var Mode = "1";
	var toggleWasOn

	function SwitchMode () {

		 if (Mode == "1") {
			if (borderShown == "yes") {
				toggleBorders()
				toggleWasOn = "yes"
			} else {
				toggleWasOn = "no"
			}
			

			toolbar_full.className = "hide";
			toolbar_code.className = "bevel3";

			// Put HTML in editor
			if (isEditingHTMLPage == "0") {
				if (useXHTML == "1") {
					code = getXHTML(document.frames('foo').document.body)
				} else {
					code = foo.document.body.innerHTML
				}

			} else {

				if (useXHTML == "1") {
					code = getXHTML(document.frames('foo').document)
				} else {
					code = foo.document.documentElement.outerHTML
				}
			}

			re = /&amp;/g
			code = code.replace(re,'&')

			if (pathType == "1") {
	
				// replaceHref = 'href="'
				replaceImage = 'src="'

				// code = code.replace(re2,replaceHref)
				code = code.replace(re3,replaceImage)
			}

			code = ConvertSSLImages(code)

			foo.document.body.innerText = code
			foo.document.body.innerHTML = colourCode(foo.document.body.innerHTML);

			// nice looking source editor
			foo.document.body.runtimeStyle.fontFamily = "Courier"
			foo.document.body.runtimeStyle.fontSize = "10px"
			foo.document.body.runtimeStyle.bgColor = '#FFFFFF';
			foo.document.body.runtimeStyle.text = '#000000';
			foo.document.body.runtimeStyle.background = '';
			foo.document.body.runtimeStyle.marginTop = '10px';
			foo.document.body.runtimeStyle.marginLeft = '10px';
			
			Mode = "2";
		} else {
			code = foo.document.body.innerText
			if (myStyleSheet != "")
			{
				code = "<link rel='stylesheet' href='" + myStyleSheet + "' type='text/css'>" + code
			}

			code = RevertSSLImages(code)

			foo.document.write(code);
			foo.document.close()

			foo.document.body.runtimeStyle.cssText = ""

			toolbar_full.className = "bevel3";
			toolbar_code.className = "hide";

			Mode = "1";

			if (toggleWasOn == "yes") {
				toggleBorders()
				toggleWasOn = "no"
			}
		}
	}

	function SaveHTMLPage() {
		var code = ""

		if (previewModeView) {
			if (isEditingHTMLPage) {
				if (useXHTML == "1") {
					code = getXHTML(document.frames('previewFrame').document)
				} else {
					code = previewFrame.document.documentElement.outerHTML;
				}
			} else {
				if (useXHTML == "1") {
					code = getXHTML(document.frames('previewFrame').document.body)
				} else {
					code = previewFrame.document.body.innerHTML;
				}
			}
		}

		if (sourceModeView) {
			code = foo.document.body.innerText;
		}

		if (editModeView)
		{
			if (isEditingHTMLPage == "0") {
				if (useXHTML == "1") {
					code = getXHTML(document.frames('foo').document.body)
				} else {
					code = foo.document.body.innerHTML
				}

			} else {

				if (useXHTML == "1") {
					code = getXHTML(document.frames('foo').document)
				} else {
					code = foo.document.documentElement.outerHTML
				}
			}
		}

		re = /&amp;/g
		code = code.replace(re,'&')
		
		if (pathType == "1")
		{
			// replaceHref = 'href="'
			replaceImage = 'src="'

			// code = code.replace(re2,replaceHref)
			code = code.replace(re3,replaceImage)
		}

		code = ConvertSSLImages(code)

		return code;
	}

	// convert src=https to just src=http
	function ConvertSSLImages(code) {
		replaceImage = 'src=\"http://' + URL
		code = code.replace(re4,replaceImage)
		return code;
	}

	function RevertSSLImages(code) {
		replaceImage = 'src=\"' + HTTPStr + '://' + URL
		code = code.replace(re5,replaceImage)
		return code;
	}

	function button_over(eButton){
		if (eButton.style.borderBottom != "#ffffff 1px solid")
		{
		eButton.style.borderBottom = "#808080 solid 1px";
		eButton.style.borderLeft = "#FFFFFF solid 1px";
		eButton.style.borderRight = "#808080 solid 1px";
		eButton.style.borderTop = "#FFFFFF solid 1px";
		}
	}
			
	function button_out2(eButton){
		if (eButton.style.borderBottom != "#ffffff 1px solid")
		{
		eButton.style.borderColor = "#d4d0c8";
		}
	}
				
	function button_out(eButton){
		eButton.style.borderColor = "#d4d0c8";
	}

	function char_out(eButton){
		eButton.style.borderColor = "#666666";
	}

	function button_down(eButton){
		eButton.style.borderBottom = "#FFFFFF solid 1px";
		eButton.style.borderLeft = "#808080 solid 1px";
		eButton.style.borderRight = "#FFFFFF solid 1px";
		eButton.style.borderTop = "#808080 solid 1px";
	}

	function button_up(eButton){
		eButton.style.borderBottom = "#808080 solid 1px";
		eButton.style.borderLeft = "#FFFFFF solid 1px";
		eButton.style.borderRight = "#808080 solid 1px";
		eButton.style.borderTop = "#FFFFFF solid 1px";
		eButton = null; 
	}

	function contextHilite(menu){
	    menu.runtimeStyle.backgroundColor = "Highlight";
	    if (menu.state){
	        menu.runtimeStyle.color = "GrayText";
	    } else {
	        menu.runtimeStyle.color = "HighlightText";
	    }
	}

	function contextDelite(menu){
	    menu.runtimeStyle.backgroundColor = "";
	    menu.runtimeStyle.color = "";
	}

	function toggleTick(tick, state) {

		if(tick.id.indexOf("zoom" + zoomSize + "_") > -1)
		{
			if(state == 1)
			{
				// We are over the selected zoom
				tick.src = 'RTEdit/Images/button_tick_inverted.gif'
			}
			else
			{
				// We are over the selected zoom
				tick.src = 'RTEdit/Images/button_tick.gif'
			}
		}
	}

	function closePopups() {
		if (imageWin) imageWin.close()
		if (propWin) propWin.close()
		if (inserttableWin) inserttableWin.close()
		if (previewWin) previewWin.close()
		if (modifytableWin) modifytableWin.close()
		if (insertFormWin) insertFormWin.close()
		if (textFieldWin) textFieldWin.close()
		if (hiddenWin) hiddenWin.close()
		if (buttonWin) buttonWin.close()
		if (checkboxWin) checkboxWin.close()
		if (radioWin) radioWin.close()
		if (linkWin) linkWin.close()
		if (emailWin) emailWin.close()
		if (anchorWin) anchorWin.close()
		if (showHelpWin) showHelpWin.close()
	}

	function isSelection() {
			if ((foo.document.selection.type == "Text") || (foo.document.selection.type == "Control")) {
				return true;
			} else {
				return false;
			}
	}

	function isTextSelected() {
			if (foo.document.selection.type == "Text") {
				return true;
			} else {
				return false;
			}
	}

	function selectImage(image) {
			document.execCommand("InsertImage",false,image);
	}

	function setBackgd(image) {
			foo.document.body.background = image
	}

	function ShowPreview(source) {

		var previewHTML
		if (source == 1)
		{
			previewHTML = foo.document.body.innerText
		} else {
			previewHTML = foo.document.documentElement.outerHTML
		}

		if (myStyleSheet != "")
		{
			previewHTML = "<link rel='stylesheet' href='" + myStyleSheet + "' type='text/css'>" + previewHTML
		}

		re = /<!DOCTYPE([^>])*>/
		previewHTML = previewHTML.replace(re,"")

		previewHTML = RevertSSLImages(previewHTML)

		previewFrame.document.write(previewHTML)

		previewFrame.document.close()
	}

	function doLink() {
		if (isAllowed())
		{
			if (isSelection()) { 
				var leftPos = (screen.availWidth-500) / 2
				var topPos = (screen.availHeight-300) / 2 
		 		linkWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertLink','','width=500,height=300,scrollbars=yes,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
			} else
				return
		}
	}

	function doImage() {

		if (isAllowed())
		{

		if (isImageSelected()) {	 
			var leftPos = (screen.availWidth-500) / 2
			var topPos = (screen.availHeight-320) / 2 
			imageWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyImage&imgDir=' + imageDir + '&tn=' + showThumbnails + '&du=' + disableImageUploading + '&dd=' + disableImageDeleting,'','width=500,height=320,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		} else {
			var leftPos = (screen.availWidth-700) / 2
			var topPos = (screen.availHeight-500) / 2 
			imageWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertImage&imgDir=' + imageDir + '&wi=' + HideWebImage + '&tn=' + showThumbnails + '&du=' + disableImageUploading + '&dd=' + disableImageDeleting + '&dt=' + isEditingHTMLPage,'','width=700,height=500,scrollbars=yes,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}

		}
	}

	function ModifyProperties() {
		if (isAllowed())
		{

		var leftPos = (screen.availWidth-500) / 2
		var topPos = (screen.availHeight-450) / 2 
	 	propWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=PageProperties','','width=500,height=450,scrollbars=no,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
		
		}
	}

	function ShowInsertTable() {
		if (isAllowed())
		{

		var leftPos = (screen.availWidth-500) / 2
		var topPos = (screen.availHeight-300) / 2 
 		inserttableWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertTable','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);

		}
	}

	function ModifyTable() {
		if (isAllowed())
		{

		if (isTableSelected() || isCursorInTableCell()) {
			var leftPos = (screen.availWidth-500) / 2
			var topPos = (screen.availHeight-300) / 2 
	 		modifytableWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyTable','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}

		}
	}

	function ModifyCell() {
		if (isAllowed())
		{

		if (isCursorInTableCell()) {
			var leftPos = (screen.availWidth-500) / 2
			var topPos = (screen.availHeight-300) / 2 
	 		modifytableWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyCell','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}

		}
	}

	function modifyForm() {
		if (isAllowed)
		{

		if (isCursorInForm()) {
			var leftPos = (screen.availWidth-500) / 2
			var topPos = (screen.availHeight-300) / 2 
	 		modifyFormWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyForm','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}

		}
	}

	function insertForm() {
		if (isAllowed())
		{
			var leftPos = (screen.availWidth-500) / 2
			var topPos = (screen.availHeight-300) / 2 
	 		insertFormWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertForm','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}
	}

	function doCustomInserts() {
		if (isAllowed())
		{
			var leftPos = (screen.availWidth-500) / 2
			var topPos = (screen.availHeight-300) / 2 
	 		customInsertWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=CustomInsert','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}
	}

	function doAnchor() {
			if (isAllowed())
			{

			var leftPos = (screen.availWidth-500) / 2
			var topPos = (screen.availHeight-250) / 2 
		
			if ((foo.document.selection.type == "Control") && (foo.document.selection.createRange()(0).tagName == "A") && (foo.document.selection.createRange()(0).href == ""))
			{
				anchorWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyAnchor','','width=500,height=250,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
			} else {
	 			anchorWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertAnchor','','width=500,height=250,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
			}

			}
	}

	function doEmail() {
		if (isAllowed())
		{
			if (isSelection()) { 
				var leftPos = (screen.availWidth-500) / 2
				var topPos = (screen.availHeight-300) / 2 
	 			emailWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertEmail','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
			} else
				return
		}
	}

	function ShowFindDialog() {
		if (isAllowed())
		{
		showModelessDialog(HTTPStr + "://" + URL + ScriptName + "?ToDo=FindReplace", foo, "dialogWidth:385px; dialogHeight:165px; scroll:no; status:no; help:no;" );
		}
	}
	
	function spellCheck(){
		var leftPos = (screen.availWidth-300) / 2
		var topPos = (screen.availHeight-220) / 2 
		arr = getWords();
	    rng = getRange();
	    spellcheckWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=SpellCheck', "spellwin", "width=300,height=220,scrollbars=no, top=" + topPos + ",left=" + leftPos);
	}

	function doHelp() {
		var leftPos = (screen.availWidth-500) / 2
		var topPos = (screen.availHeight-400) / 2 
	 	showHelpWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ShowHelp','','width=500,height=400,scrollbars=yes,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
	}

	function doTextField() {
		if (isAllowed())
		{

		var leftPos = (screen.availWidth-500) / 2
		var topPos = (screen.availHeight-300) / 2 

		if (foo.document.selection.type == "Control") {
			var oControlRange = foo.document.selection.createRange();
			if (oControlRange(0).tagName.toUpperCase() == "INPUT") {
				if ((oControlRange(0).type.toUpperCase() == "TEXT") || (oControlRange(0).type.toUpperCase() == "PASSWORD")) {
					selectedTextField = foo.document.selection.createRange()(0);
					textFieldWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyTextField','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
				}
				return true;
			}	
		} else {
			textFieldWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertTextField','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}

		}
	}

	function doHidden() {
		if (isAllowed())
		{

		var leftPos = (screen.availWidth-500) / 2
		var topPos = (screen.availHeight-300) / 2 

		if (foo.document.selection.type == "Control") {
			var oControlRange = foo.document.selection.createRange();
			if (oControlRange(0).tagName.toUpperCase() == "INPUT") {
				if (oControlRange(0).type.toUpperCase() == "HIDDEN") {
					selectedHidden = foo.document.selection.createRange()(0);
					hiddenWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyHidden','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
				}
				return true;
			}	
		} else {
			hiddenWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertHidden','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}

		}
	}

	function doTextArea() {
		if (isAllowed())
		{

		var leftPos = (screen.availWidth-500) / 2
		var topPos = (screen.availHeight-300) / 2 

		if (foo.document.selection.type == "Control") {
			var oControlRange = foo.document.selection.createRange();
			if (oControlRange(0).tagName.toUpperCase() == "TEXTAREA") {
					selectedTextArea = foo.document.selection.createRange()(0);
					textFieldWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyTextArea','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
				return true;
			}	
		} else {
			textFieldWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertTextArea','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}

		}
	}

	function doButton() {
		if (isAllowed())
		{

		var leftPos = (screen.availWidth-500) / 2
		var topPos = (screen.availHeight-300) / 2 

		if (foo.document.selection.type == "Control") {
			var oControlRange = foo.document.selection.createRange();
			if (oControlRange(0).tagName.toUpperCase() == "INPUT") {
				if ((oControlRange(0).type.toUpperCase() == "RESET") || (oControlRange(0).type.toUpperCase() == "SUBMIT") || (oControlRange(0).type.toUpperCase() == "BUTTON")) {
					selectedButton = foo.document.selection.createRange()(0);
					buttonWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyButton','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
				}
				return true;
			}	
		} else {
			buttonWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertButton','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}

		}
	}

	function doCheckbox() {
		if (isAllowed())
		{

		var leftPos = (screen.availWidth-500) / 2
		var topPos = (screen.availHeight-300) / 2 

		if (foo.document.selection.type == "Control") {
			var oControlRange = foo.document.selection.createRange();
			if (oControlRange(0).tagName.toUpperCase() == "INPUT") {
				if (oControlRange(0).type.toUpperCase() == "CHECKBOX") {
					selectedCheckbox = foo.document.selection.createRange()(0);
					checkboxWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyCheckbox','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
				}
				return true;
			}	
		} else {
			checkboxWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertCheckbox','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}

		}
	}

	function doRadio() {
		if (isAllowed()) {

		var leftPos = (screen.availWidth-500) / 2
		var topPos = (screen.availHeight-300) / 2 

		if (foo.document.selection.type == "Control") {
			var oControlRange = foo.document.selection.createRange();
			if (oControlRange(0).tagName.toUpperCase() == "INPUT") {
				if (oControlRange(0).type.toUpperCase() == "RADIO") {
					selectedRadio = foo.document.selection.createRange()(0);
					radioWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=ModifyRadio','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
				}
				return true;
			}	
		} else {
			radioWin = window.open(HTTPStr + '://' + URL + ScriptName + '?ToDo=InsertRadio','','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
		}

		}
	}

	function cleanCode() {
		if (confirm(sTxtClean)){

			var borderwason

			if (borderShown == "yes") {
			 	toggleBorders()
			 	borderwason = true
			}

			foo.document.body.innerHTML = doCleanCode(foo.document.body.innerHTML)

			if (borderwason) {
			 	toggleBorders()
			}
		}
	}

// Colorize Code in Source Mode
function colourCode(code) {

	htmlTag = /(&lt;([\s\S]*?)&gt;)/gi
	tableTag = /(&lt;(table|tbody|th|tr|td|\/table|\/tbody|\/th|\/tr|\/td)([\s\S]*?)&gt;)/gi
	commentTag = /(&lt;!--([\s\S]*?)&gt;)/gi
	imageTag = /(&lt;img([\s\S]*?)&gt;)/gi
	linkTag = /(&lt;(a|\/a)([\s\S]*?)&gt;)/gi
	scriptTag = /(&lt;(script|\/script)([\s\S]*?)&gt;)/gi

	code = code.replace(htmlTag,"<font color=#000080>$1</font>")
	code = code.replace(tableTag,"<font color=#008080>$1</font>")
	code = code.replace(commentTag,"<font color=#808080>$1</font>")
	code = code.replace(imageTag,"<font color=#800080>$1</font>")
	code = code.replace(linkTag,"<font color=#008000>$1</font>")
	code = code.replace(scriptTag,"<font color=#800000>$1</font>")

	return code;
}
// End colorize