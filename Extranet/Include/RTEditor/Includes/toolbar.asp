<script>

var tickImg = new Image
tickImg.src = "/include/RTEditor/Images/button_tick.gif"

var tickImg2 = new Image
tickImg.src = "/include/RTEditor/Images/button_tick_inverted.gif"

</script>

<table width="100%" cellspacing="0" cellpadding="0" class=toolbar>
	<tr>
	<td class="body" height="22">
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class="hide" align="center" id="toolbar_preview">
		<tr>
		  <td class="body" height="52">
		  &nbsp;&nbsp;&nbsp;<b>Preview Mode</b>
		  </td>
		 </tr>
	</table>
	 <table width="100%" border="0" cellspacing="0" cellpadding="0" class="hide" align="center" id="toolbar_code">
		<tr>
		  <td class="body" height="22">
		  <table border="0" cellspacing="0" cellpadding="1">
			  <tr id=de>
				  <% if e__hideFullScreen <> true then %>
					<td>
						<img id=fullscreen2 border="0" src="/include/RTEditor/Images/button_fullscreen.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='toggleSize();foo.focus();' title="<%=sTxtFullscreen%>" class=toolbutton>
					</td>
					<% end if %>
          <% if e__hideCut <> true then %>
  				<TD>
  				  <IMG BORDER="0" DISABLED ID="toolbarCut2_off" SRC="/Include/RTEditor/Images/button_cut_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtCut%> (Ctrl+X)" class=toolbutton><IMG BORDER="0" ID="toolbarCut2_on" SRC="/Include/RTEditor/Images/button_cut.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Cut");FOO.FOCUS();' TITLE="<%=sTxtCut%> (Ctrl+X)" class=toolbutton style="display:none">
  				</TD>
          <% end if %>
  
          <% if e__hideCopy <> true then %>
  				<TD>
  				  <IMG BORDER="0" DISABLED ID="toolbarCopy2_off" SRC="/Include/RTEditor/Images/button_copy_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton><IMG BORDER="0" ID="toolbarCopy2_on" SRC="/Include/RTEditor/Images/button_copy.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Copy");FOO.FOCUS();' TITLE="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton style="display:none">
  				</TD>
          <% end if %>
  
          <% if e__hidePaste <> true then %>
          <TD>
  				  <IMG BORDER="0" DISABLED ID="toolbarPasteButton2_off" SRC="/Include/RTEditor/Images/button_paste_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtPaste%> (Ctrl+V)" class=toolbutton><IMG BORDER="0" ID="toolbarPasteButton2_on" SRC="/Include/RTEditor/Images/button_paste.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Paste");FOO.FOCUS();' TITLE="<%=sTxtPaste%> (Ctrl+V)" class=toolbutton style="display:none">
  				</TD>
          <% end if %>
  
          <% if e__hideFind <> true then %>
          <TD>
	  			  <IMG BORDER="0" SRC="/Include/RTEditor/Images/button_find.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='SHOWFINDDIALOG();FOO.FOCUS();' TITLE="<%=sTxtFindReplace%>" class=toolbutton>
			  	</TD>
          <% end if %>
    
          <% if e__hideFullScreen <> true or _
                e__hideCut <> true or _
                e__hideCopy <> true or _
                e__hidePast <> true or _
                e__hideFind <> true then %>
                
  				<TD><IMG SRC="/Include/RTEditor/Images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
          <% end if %>
  
          <% if e__hideUndo <> true then %>
  				<TD>
            <IMG BORDER="0" DISABLED ID="undo2_off" SRC="/Include/RTEditor/Images/button_undo_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton><IMG BORDER="0" ID="undo2_on" SRC="/Include/RTEditor/Images/button_undo.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Undo");' TITLE="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton style="display:none">
  				</TD>
          <% end if %>
  
          <% if e__hideRedo <> true then %>
  				<TD>
  				  <IMG BORDER="0" DISABLED ID="redo2_off" SRC="/Include/RTEditor/Images/button_redo_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton><IMG BORDER="0" ID="redo2_on" SRC="/Include/RTEditor/Images/button_redo.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Redo");' TITLE="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton style="display:none">
  				</TD>
          <% end if %>
				</tr>
			</table>
		  </td>
		 </tr>
		<tr>
		  <td class="body" bgcolor="#808080"><img src="/Include/RTEditor/Images/1x1.gif" width="1" height="1"></td>
		</tr>
		<tr>
		  <td class="body" bgcolor="#FFFFFF"><img src="/Include/RTEditor/Images/1x1.gif" width="1" height="1"></td>
		</tr>
		 <tr><td height=28>&nbsp;</td></tr>
	</table>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="bevel3" align="center" id="toolbar_full">
		<tr>
		  <td class="body" height="22">
			<table border="0" cellspacing="0" cellpadding="1">
			  <tr id=de>

				  <% if e__hideFullScreen <> true then %>
					<td>
						<img id=fullscreen border="0" src="/include/RTEditor/Images/button_fullscreen.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='toggleSize();foo.focus();' title="<%=sTxtFullscreen%>" class=toolbutton>
					</td>
				  <% end if %>

          <% if e__hideCut <> true then %>
					<TD>
   			    <IMG BORDER="0" DISABLED ID="toolbarCut_off" SRC="/Include/RTEditor/Images/button_cut_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtCut%> (Ctrl+X)" class=toolbutton><IMG BORDER="0" ID="toolbarCut_on" SRC="/Include/RTEditor/Images/button_cut.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Cut");FOO.FOCUS();' TITLE="<%=sTxtCut%> (Ctrl+X)" class=toolbutton style="display:none"><DIV CLASS="pasteArea" ID="myTempArea" CONTENTEDITABLE></DIV>
					</TD>
          <% end if %>
    
          <% if e__hideCopy <> true then %>
					<TD>
						<IMG BORDER="0" DISABLED ID="toolbarCopy_off" SRC="/Include/RTEditor/Images/button_copy_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton><IMG BORDER="0" ID="toolbarCopy_on" SRC="/Include/RTEditor/Images/button_copy.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Copy");FOO.FOCUS();' TITLE="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton style="display:none">
					</TD>
          <% end if %>
    
          <% if e__hidePaste <> true then %>
					<TD>
						<IMG ID=TOOLBARPASTEBUTTON_OFF DISABLED CLASS=TOOLBUTTON WIDTH="21" HEIGHT="20" SRC="/Include/RTEditor/Images/button_paste_disabled.gif" BORDER=0 UNSELECTABLE="on" TITLE="<%=sTxtPaste%> (Ctrl+V)"><IMG ID=TOOLBARPASTEBUTTON_ON CLASS=TOOLBUTTON onMouseDown="button_down(this);" onMouseOver="button_over(this); button_over(toolbarPasteDrop_on)" onClick="doCommand('Paste'); foo.focus()" onMouseOut="button_out(this); button_out(toolbarPasteDrop_on);" WIDTH="21" HEIGHT="20" SRC="/Include/RTEditor/Images/button_paste.gif" BORDER=0 UNSELECTABLE="on" TITLE="<%=sTxtPaste%> (Ctrl+V)" style="display:none"><IMG ID=TOOLBARPASTEDROP_OFF DISABLED CLASS=TOOLBUTTON WIDTH="7" HEIGHT="20" SRC="/Include/RTEditor/Images/button_drop_menu_disabled.gif" BORDER=0 UNSELECTABLE="on"><IMG ID=TOOLBARPASTEDROP_ON CLASS=TOOLBUTTON onMouseDown="button_down(this);" onMouseOver="button_over(this); button_over(toolbarPasteButton_on)" onClick="showMenu('pasteMenu',180,42)" onMouseOut="button_out(this); button_out(toolbarPasteButton_on);" WIDTH="7" HEIGHT="20" SRC="/Include/RTEditor/Images/button_drop_menu.gif" BORDER=0 UNSELECTABLE="on" STYLE="display:none">
					</TD>
          <% end if %>
    
          <% if e__hideFind <> true then %>
					<TD>
					  <IMG BORDER="0" SRC="/Include/RTEditor/Images/button_find.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='SHOWFINDDIALOG();FOO.FOCUS();' TITLE="<%=sTxtFindReplace%>" class=toolbutton>
					</TD>
          <% end if %>
    
          <% if e__hideFullScreen <> true or _
                e__hideCut <> true or _
                e__hideCopy <> true or _
                e__hidePaste <> true or _
                e__hideFind <> true then
          %>
					<TD>
						<IMG SRC="/Include/RTEditor/Images/seperator.gif" WIDTH="2" HEIGHT="20">
					</TD>
          <% end if %>

          <% if e__hideUndo <> true then %>
					<TD>
						<IMG ID="undo_off" DISABLED UNSELECTABLE="on" BORDER="0" SRC="/Include/RTEditor/Images/button_undo_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton>
						<IMG ID="undo_on" UNSELECTABLE="on" BORDER="0" SRC="/Include/RTEditor/Images/button_undo.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='GOHISTORY(-1);' TITLE="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton style="display:none">
					</TD>
          <% end if %>
    
          <% if e__hideRedo <> true then %>
					<TD>
						<IMG ID="redo_off" DISABLED UNSELECTABLE="on" BORDER="0" SRC="/Include/RTEditor/Images/button_redo_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton>
						<IMG ID="redo_on" UNSELECTABLE="on" BORDER="0" SRC="/Include/RTEditor/Images/button_redo.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='GOHISTORY(1);' TITLE="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton style="display:none">
					</TD>
          <% end if %>
    
          <% if e__hideUndo <> true or e__hideRedo <> true then %>
  			  <TD><IMG SRC="/Include/RTEditor/Images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
          <% end if %>

				  <% if e__hideSpelling <> true then %>
					<td>
						<img id="toolbarSpell" border="0" src="/include/RTEditor/Images/button_spellcheck.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='spellCheck();' title="Check Spelling (F7)" class=toolbutton>
					</td>
				  <td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				  <% end if %>

				  <% if e__hideRemoveTextFormatting <> true then %>
					<td>
						<img border="0" src="/include/RTEditor/Images/button_remove_format.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("RemoveFormat");' title="<%=sTxtRemoveFormatting%>" class=toolbutton>
					</td>
					<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				  <% end if %>
				  <% if e__hideBold <> true then %>
					<td>
						<img id="fontBold" border="0" src="/include/RTEditor/Images/button_bold.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("Bold");foo.focus();' title="<%=sTxtBold%> (Ctrl+B)" class=toolbutton>
					</td>
				  <% end if %>
				  <% if e__hideUnderline <> true then %>
					<td>
						<img id="fontUnderline" border="0" src="/include/RTEditor/Images/button_underline.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("Underline");foo.focus();' title="<%=sTxtUnderline%> (Ctrl+U)" class=toolbutton>
					</td>
				  <% end if %>
				  <% if e__hideItalic <> true then %>
					<td>
						<img id="fontItalic" border="0" src="/include/RTEditor/Images/button_italic.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("Italic");foo.focus();' title="<%=sTxtItalic%> (Ctrl+I)" class=toolbutton>
					</td>
				  <% end if %>

  				<% if e__hideStrikethrough <> true then %>
  					<td>
  						<img id="fontStrikethrough" border="0" src="/include/RTEditor/Images/button_strikethrough.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("Strikethrough");foo.focus();' title="<%=sTxtStrikethrough%>" class=toolbutton>
  					</td>
  				<% end if %>
        
          <% if e__hideBold <> true or _
                e__hideUnderline <> true or _
                e__hideItalic <> true or _
                e__hideStrikethrough <> true then%>
              
  			  <TD><IMG SRC="/Include/RTEditor/Images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
          <% end if %>
  				<!-- End -->

				  <% if e__hideNumberList <> true then %>
					<td>
						<img id="fontInsertOrderedList" border="0" src="/include/RTEditor/Images/button_numbers.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("InsertOrderedList");foo.focus();' title="<%=sTxtNumList%>" class=toolbutton>
					</td>
				  <% end if %>
				  <% if e__hideBulletList <> true then %>
					<td>
						<img id="fontInsertUnorderedList" border="0" src="/include/RTEditor/Images/button_bullets.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("InsertUnorderedList");foo.focus();' title="<%=sTxtBulletList%>" class=toolbutton>
					</td>
				  <% end if %>
				  <% if e__hideDecreaseIndent <> true then %>
					<td>
					<img border="0" src="/include/RTEditor/Images/button_decrease_indent.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Outdent");foo.focus();' title="<%=sTxtDecreaseIndent%>" class=toolbutton>
					</td>
				  <% end if %>
				  <% if e__hideIncreaseIndent <> true then %>
					<td>
						<img border="0" src="/include/RTEditor/Images/button_increase_indent.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Indent");foo.focus();' title="<%=sTxtIncreaseIndent%>" class=toolbutton>
					</td>
					<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				  <% end if %>
				  <% if e__hideSuperScript <> true then %>
					<td>
						<img id="fontSuperScript" border="0" src="/include/RTEditor/Images/button_superscript.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("superscript");foo.focus();' title="<%=sTxtSuperscript%>" class=toolbutton>
					</td>
				  <% end if %>
				  <% if e__hideSubScript <> true then %>
					<td>
						<img id="fontSubScript" border="0" src="/include/RTEditor/Images/button_subscript.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("subscript");foo.focus();' title="<%=sTxtSubscript%>" class=toolbutton>
					</td>
					<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				  <% end if %>
				  <% if e__hideLeftAlign <> true then %>
					<td>
						<img id="fontJustifyLeft" border="0" src="/include/RTEditor/Images/button_align_left.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("JustifyLeft");foo.focus();' title="<%=sTxtAlignLeft%>" class=toolbutton>
					</td>
				  <% end if %>
				  <% if e__hideCenterAlign <> true then %>
					<td>
						<img id="fontJustifyCenter" border="0" src="/include/RTEditor/Images/button_align_center.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("JustifyCenter");foo.focus();' title="<%=sTxtAlignCenter%>" class=toolbutton>
					</td>
				  <% end if %>
				  <% if e__hideRightAlign <> true then %>
					<td>
						<img id="fontJustifyRight" border="0" src="/include/RTEditor/Images/button_align_right.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("JustifyRight");foo.focus();' title="<%=sTxtAlignRight%>" class=toolbutton>
					</td>
				  <% end if %>
				  <% if e__hideJustify <> true then %>
					<td>
						<img id="fontJustifyFull" border="0" src="/include/RTEditor/Images/button_align_justify.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("JustifyFull");foo.focus();' title="<%=sTxtAlignJustify%>" class=toolbutton>
					</td>
					<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				  <% end if %>
				  <% if e__hideLink <> true then %>
					<td>
						<img disabled id="toolbarLink_off" border="0" src="/include/RTEditor/Images/button_link_disabled.gif" width="21" height="20" title="<%=sTxtHyperLink%>" class=toolbutton><img id="toolbarLink_on" border="0" src="/include/RTEditor/Images/button_link.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doLink()' title="<%=sTxtHyperLink%>" class=toolbutton style="display:none">
					</td>
				  <% end if %>
				  <% if e__hideMailLink <> true then %>
					<td>
						<img border="0" id="toolbarEmail_off" disabled src="/include/RTEditor/Images/button_email_disabled.gif" width="21" height="20" title="<%=sTxtEmail%>" class=toolbutton><img border="0" id="toolbarEmail_on" src="/include/RTEditor/Images/button_email.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doEmail()' title="<%=sTxtEmail%>" class=toolbutton style="display:none">
					</td>
				  <% end if %>
				  <% if e__hideAnchor <> true then %>
					<td>
						<img border="0" src="/include/RTEditor/Images/button_anchor.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doAnchor()' title="<%=sTxtAnchor%>" class=toolbutton>
					</td>
					<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				  <% end if %>
				  <% if e__hideHelp <> true then %>
					<td>
						<img border="0" src="/include/RTEditor/Images/button_help.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doHelp()' title="<%=sTxtHelp%>" class=toolbutton>
					</td>
				  <% end if %>
			  </tr>
	  		</table>
		  </td>
		</tr>
			<tr>
		  <td class="body" bgcolor="#808080"><img src="/Include/RTEditor/Images/1x1.gif" width="1" height="1"></td>
		</tr>
		<tr>
		  <td class="body" bgcolor="#FFFFFF"><img src="/Include/RTEditor/Images/1x1.gif" width="1" height="1"></td>
		</tr>
		<tr>
		  <td class="body">
			<table border="0" cellspacing="1" cellpadding="1">
			  <tr id=de>
				<% if e__hideStyle <> true then %>
				<td>
				  <select id="sStyles" onChange="applyStyle(this[this.selectedIndex].value);foo.focus();this.selectedIndex=0;foo.focus();" class="Text90" unselectable="on" onmouseenter="doStyles()">
				    <option selected><%=sTxtStyle%></option>
				    <option value="">None</option>
				  </select>
				</td>
				<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				<% end if %>
				<% if e__hideFont <> true then %>
				<td>
				  <select id="fontDrop" onChange="doFont(this[this.selectedIndex].value)" class="Text120" unselectable="on">
					<%= BuildFontList() %>
				  </select>
				</td>
				<% end if %>
				<% if e__hideSize <> true then %>
				<td>
				  <select id="sizeDrop" onChange="doSize(this[this.selectedIndex].value)" class=Text50 unselectable="on">
					<%= BuildSizeList() %>
	  			  </select>
				</td>
				<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				<% end if %>
				<% if e__hideFormat <> true then %>
				<td>
				  <select id="formatDrop" onChange="doFormat(this[this.selectedIndex].value)" class="Text70" unselectable="on">
				    <option selected><%=sTxtFormat%>
				    <option value="<P>">Normal
				    <option value="<H1>">Heading 1
				    <option value="<H2>">Heading 2
				    <option value="<H3>">Heading 3
				    <option value="<H4>">Heading 4
				    <option value="<H5>">Heading 5
				    <option value="<H6>">Heading 6
				  </select>
				</td>
				<% end if %>
				<% if e__hideForeColor <> true then %>
				<td>
				  <img border="0" src="/include/RTEditor/Images/button_font_color.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="(isAllowed()) ? showMenu('colorMenu',180,291) : foo.focus()" class=toolbutton title="<%=sTxtColour%>">
				</td>
				<% end if %>
				<% if e__hideBackColor <> true then %>
				<td>
				  <img border="0" src="/include/RTEditor/Images/button_highlight.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="(isAllowed()) ? showMenu('colorMenu2',180,291) : foo.focus()" class=toolbutton title="<%=sTxtBackColour%>">
				</td>
				<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				<% end if %>
				<% if e__hideTable <> true then %>
				<td id=toolbarTables>
				  <img border="0" src="/include/RTEditor/Images/button_table_down.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="(isAllowed()) ? showMenu('tableMenu',160,262) : foo.focus()" class=toolbutton title="<%=sTxtTableFunctions%>">
				</td>
				<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				<% end if %>
				<% if e__hideForm <> true then %>
				<td>
				  <img class=toolbutton onMouseDown=button_down(this); onMouseOver=button_over(this); onClick="(isAllowed()) ? showMenu('formMenu',180,189) : foo.focus()" onMouseOut=button_out(this); type=image width="21" height="20" src="/include/RTEditor/Images/button_form_down.gif" border=0 title="<%=sTxtFormFunctions%>">
				</td>
				<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				<% end if %>
				<% if e__hideImage <> true then %>
				<td>
				  <img border="0" src="/include/RTEditor/Images/button_image.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="doImage()" class=toolbutton title="<%=sTxtImage%>">
				</td>
				<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				<% end if %>
				<% if e__hideTextBox <> true then %>
				<td>
				  <img border="0" src="/include/RTEditor/Images/button_textbox.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="doTextbox()" class=toolbutton title="<%=sTxtTextbox%>">
				</td>
				<% end if %>
				 <% if e__hideHorizontalRule <> true then %>
					<td>
						<img border="0" src="/include/RTEditor/Images/button_hr.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("InsertHorizontalRule");foo.focus();' title="<%=sTxtInsertHR%>" class=toolbutton>
					</td>
				  <% end if %>
				<% if e__hideSymbols <> true then %>
				<td>
				  <img border="0" src="/include/RTEditor/Images/button_chars.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="(isAllowed()) ? showMenu('charMenu',104,111) : foo.focus()" class=toolbutton title="<%=sTxtChars%>">
				</td>
				<% end if %>
				<% if e__hideProps <> true then %>
				<td>
				  <img border="0" src="/include/RTEditor/Images/button_properties.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="ModifyProperties()" class=toolbutton title="<%=sTxtPageProperties%>">
				</td>
				<% end if %>

				<% if e__hasCustomInserts = true then %>
				<td>
					<img class=toolbutton onmousedown="button_down(this);" onmouseover="button_over(this);" onClick="doCustomInserts()" onmouseout="button_out(this);" type=image width="21" height="20" src="/include/RTEditor/Images/button_custom_inserts.gif" border=0 title="<%=sTxtCustomInserts%>">
				</td>
				<% end if %>

				<% if e__hideAbsolute <> true then %>
				<td>
					<img id="fontAbsolutePosition_off" disabled class=toolbutton onmousedown="button_down(this);" onmouseover="button_over(this);" width="21" height="20" src="/include/RTEditor/Images/button_absolute_disabled.gif" border=0 title="<%=sTxtTogglePosition%>"><img id="fontAbsolutePosition" class=toolbutton onmousedown="button_down(this);" onmouseover="button_over(this);" onClick="doCommand('AbsolutePosition')" onmouseout="button_out2(this);" type=image width="21" height="20" src="/include/RTEditor/Images/button_absolute.gif" border=0 title="<%=sTxtTogglePosition%>" style="display:none">
				</td>
				<% end if %>

				<% if e__hideGuidelines <> true then %>
				<td>
				  <img class=toolbutton onMouseDown="button_down(this);" onMouseOver="button_over(this);" onClick="toggleBorders()" onMouseOut="button_out2(this);" type=image width="21" height="20" src="/include/RTEditor/Images/button_show_borders.gif" border=0 title="<%=sTxtToggleGuidelines%>" id=guidelines>
				</td>
				<% end if %>

				<% if e__hideClean <> true then %>
				<td><img src="/include/RTEditor/Images/seperator.gif" width="2" height="20"></td>
				<td>				
				  <img class=toolbutton onmousedown="button_down(this);" onmouseover="button_over(this);" onClick="cleanCode()" onmouseout="button_out(this);" type=image width="21" height="20" src="/include/RTEditor/Images/button_clean_code.gif" border=0 title="<%=sTxtCleanCode%>">
				</td>
				<% end if %>
        
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr> 
</table>
<!-- table menu -->
<DIV ID="tableMenu" STYLE="display:none">
<table border="0" cellspacing="0" cellpadding="0" width=160 style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
  <tr onClick="parent.ShowInsertTable()" title="<%=sTxtTable%>" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=20> 
      &nbsp;&nbsp;<%=sTxtTable%>...&nbsp; </td>
  </tr>
  <tr onClick=parent.ModifyTable(); title="<%=sTxtTableModify%>" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=20 id=modifyTable> 
	  &nbsp;&nbsp;<%=sTxtTableModify%>...&nbsp;</td>
  </tr>
  <tr title="<%=sTxtCellModify%>" onClick=parent.ModifyCell() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=20 id=modifyCell> 
	&nbsp;&nbsp;<%=sTxtCellModify%>...&nbsp; </td>
  </tr>
  <tr height=10> 
    <td align=center><img src="/include/RTEditor/Images/vertical_spacer.gif" width="140" height="2"></td>
  </tr>
  <tr title="<%=sTxtInsertColA%>" onClick=parent.InsertColAfter() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <td style="cursor: hand; font:8pt tahoma;" height=20 id=colAfter> 
      &nbsp;&nbsp;<%=sTxtInsertColA%>&nbsp;
    </td>
  </tr>
  <tr title="<%=sTxtInsertColB%>" onClick=parent.InsertColBefore() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <td style="cursor: hand; font:8pt tahoma;" height=20 id=colBefore> 
      &nbsp;&nbsp;<%=sTxtInsertColB%>&nbsp;
    </td>
  </tr>
  <tr height=10> 
    <td align=center><img src="/include/RTEditor/Images/vertical_spacer.gif" width="140" height="2"></td>
  </tr>
  <tr title="<%=sTxtInsertRowA%>" onClick=parent.InsertRowAbove() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
	<td style="cursor: hand; font:8pt tahoma;" height=20 id=rowAbove> 
      &nbsp;&nbsp;<%=sTxtInsertRowA%>&nbsp;
    </td>
  </tr>
  <tr title="<%=sTxtInsertRowB%>" onClick=parent.InsertRowBelow() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
	<td style="cursor: hand; font:8pt tahoma;" height=20 id=rowBelow> 
      &nbsp;&nbsp;<%=sTxtInsertRowB%>&nbsp;
    </td>
  </tr>
  <tr height=10> 
    <td align=center><img src="/include/RTEditor/Images/vertical_spacer.gif" width="140" height="2"></td>
  </tr>
  <tr title="<%=sTxtDeleteRow%>" onClick=parent.DeleteRow() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <td style="cursor: hand; font:8pt tahoma;" height=20 id=deleteRow>
      &nbsp;&nbsp;<%=sTxtDeleteRow%>&nbsp;
    </td>
  </tr>
  <tr title="<%=sTxtDeleteCol%>" onClick=parent.DeleteCol() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <td style="cursor: hand; font:8pt tahoma;" height=20 id=deleteCol>
      &nbsp;&nbsp;<%=sTxtDeleteCol%>&nbsp;
    </td>
  </tr>
  <tr height=10> 
    <td align=center><img src="/include/RTEditor/Images/vertical_spacer.gif" width="140" height="2" tabindex=1 HIDEFOCUS></td>
  </tr>
  <tr title="<%=sTxtIncreaseColSpan%>" onClick=parent.IncreaseColspan() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <td style="cursor: hand; font:8pt tahoma;" height=20 id=increaseSpan>
      &nbsp;&nbsp;<%=sTxtIncreaseColSpan%>&nbsp;
    </td>
  </tr>
  <tr title="<%=sTxtDecreaseColSpan%>" onClick=parent.DecreaseColspan() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <td style="cursor: hand; font:8pt tahoma;" height=20 id=decreaseSpan>
      &nbsp;&nbsp;<%=sTxtDecreaseColSpan%>&nbsp;
    </td>
  </tr>
</table>
</div>
<!-- end table menu -->

<!-- form menu -->
<DIV ID="formMenu" STYLE="display:none;">
<table border="0" cellspacing="0" cellpadding="0" width=180 style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
  <tr title="<%=sTxtForm%>" onClick=parent.insertForm() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22>
      <img width="21" height="20" src="/include/RTEditor/Images/button_form.gif" border=0 align="absmiddle">&nbsp;<%=sTxtForm%>...&nbsp;</td>
  </tr>
  <tr title="<%=sTxtFormModify%>" onClick=parent.modifyForm() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <td style="cursor: hand; font:8pt tahoma;" id="modifyForm1" height=22 class=dropDown>
      <img id="modifyForm2" width="21" height="20" src="/include/RTEditor/Images/button_modify_form.gif" border=0 align="absmiddle">&nbsp;<%=sTxtFormModify%>...&nbsp;</td>
  </tr>
  <tr height=10> 
    <td align=center><img src="/include/RTEditor/Images/vertical_spacer.gif" width="140" height="2" tabindex=1 HIDEFOCUS></td>
  </tr>
  <tr title="<%=sTxtTextField%>" onClick=parent.doTextField() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22>
      <img width="21" height="20" src="/include/RTEditor/Images/button_textfield.gif" border=0 align="absmiddle">&nbsp;<%=sTxtTextField%>...&nbsp;</td>
  </tr>
  <tr title="<%=sTxtTextArea%>" onClick=parent.doTextArea() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <td style="cursor: hand; font:8pt tahoma;" height=22>
      <img type=image width="21" height="20" src="/include/RTEditor/Images/button_textarea.gif" border=0 align="absmiddle">&nbsp;<%=sTxtTextArea%>...&nbsp;</td>
  </tr>
  <tr title="<%=sTxtHidden%>" onClick=parent.doHidden(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <td style="cursor: hand; font:8pt tahoma;" height=22>
      <img width="21" height="20" src="/include/RTEditor/Images/button_hidden.gif" border=0 align="absmiddle">&nbsp;<%=sTxtHidden%>...&nbsp;</td>
  </tr>
  <tr title="<%=sTxtButton%>" onClick=parent.doButton(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22>
      <img width="21" height="20" src="/include/RTEditor/Images/button_button.gif" border=0 align="absmiddle">&nbsp;<%=sTxtButton%>...&nbsp;</td>
  </tr>
  <tr title="<%=sTxtCheckbox%>" onClick=parent.doCheckbox(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22>
      <img width="21" height="20" src="/include/RTEditor/Images/button_checkbox.gif" border=0 align="absmiddle">&nbsp;<%=sTxtCheckbox%>...&nbsp;</td>
  </tr>
  <tr title="<%=sTxtRadioButton%>" onClick=parent.doRadio(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22>
      <img width="21" height="20" src="/include/RTEditor/Images/button_radio.gif" border=0 align="absmiddle">&nbsp;<%=sTxtRadioButton%>...&nbsp;</td>
  </tr>
</table>
</div>
<!-- formMenu -->

<!-- zoom menu -->
<DIV ID="zoomMenu" STYLE="display:none;">
<table border="0" cellspacing="0" cellpadding="0" width=65 style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
  <tr onClick=parent.doZoom(500) onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom500_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom500_,0);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom500_">
     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;500%&nbsp;</td>
  </tr>
  <tr onClick=parent.doZoom(200) onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom200_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom200_,0);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom200_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;200%&nbsp;</td>
  </tr>
  <tr onClick=parent.doZoom(150) onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom150_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom150_,0);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom150_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;150%&nbsp;</td>
  </tr>
  <tr onClick="parent.doZoom(100)" onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom100_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom100_,0)";">
    <td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom100_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;100%&nbsp;</td>
  </tr>
  <tr onClick=parent.doZoom(75); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom75_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom75_,0);">
    <td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom75_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;75%&nbsp;</td>
  </tr>
  <tr onClick=parent.doZoom(50); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom50_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom50_,0);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom50_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;50%&nbsp;</td>
  </tr>
  <tr onClick=parent.doZoom(25); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom25_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom25_,0);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom25_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;25%&nbsp;</td>
  </tr>
  <tr onClick=parent.doZoom(10); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom10_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom10_,0);"> 
    <td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom10_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10%&nbsp;</td>
  </tr>
</table>
</div>
<!-- zoomMenu -->

<DIV ID="colorMenu" STYLE="display:none;">
<table cellpadding="1" cellspacing="5" border="1" bordercolor="#666666" style="cursor: hand;font-family: Verdana; font-size: 7px; BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
  <tr>
	<td colspan="10" id=color style="height=20px;font-family: verdana; font-size:12px;">&nbsp;</td>
  </tr>
  <tr>
    <td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FF0000;width=12px">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFFF00;width=12px">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#00FF00;width=12px">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#00FFFF;width=12px">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#0000FF;width=12px">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FF00FF;width=12px">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFFFFF;width=12px">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#F5F5F5;width=12px">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#DCDCDC;width=12px">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFFAFA;width=12px">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#D3D3D3">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#C0C0C0">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#A9A9A9">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#808080">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#696969">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#000000">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#2F4F4F">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#708090">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#778899">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#4682B4">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#4169E1">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#6495ED">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#B0C4DE">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#7B68EE">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#6A5ACD">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#483D8B">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#191970">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#000080">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#00008B">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#0000CD">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#1E90FF">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#00BFFF">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#87CEFA">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#87CEEB">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#ADD8E6">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#B0E0E6">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#F0FFFF">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#E0FFFF">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#AFEEEE">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#00CED1">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#5F9EA0">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#48D1CC">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#00FFFF">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#40E0D0">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#20B2AA">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#008B8B">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#008080">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#7FFFD4">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#66CDAA">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#8FBC8F">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#3CB371">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#2E8B57">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#006400">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#008000">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#228B22">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#32CD32">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#00FF00">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#7FFF00">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#7CFC00">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#ADFF2F">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#98FB98">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#90EE90">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#00FF7F">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#00FA9A">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#556B2F">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#6B8E23">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#808000">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#BDB76B">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#B8860B">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#DAA520">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFD700">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#F0E68C">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#EEE8AA">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFEBCD">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFE4B5">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#F5DEB3">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFDEAD">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#DEB887">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#D2B48C">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#BC8F8F">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#A0522D">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#8B4513">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#D2691E">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#CD853F">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#F4A460">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#8B0000">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#800000">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#A52A2A">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#B22222">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#CD5C5C">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#F08080">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FA8072">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#E9967A">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFA07A">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FF7F50">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FF6347">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FF8C00">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFA500">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FF4500">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#DC143C">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FF0000">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FF1493">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FF00FF">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FF69B4">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFB6C1">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFC0CB">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#DB7093">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#C71585">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#800080">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#8B008B">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#9370DB">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#8A2BE2">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#4B0082">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#9400D3">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#9932CC">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#BA55D3">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#DA70D6">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#EE82EE">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#DDA0DD">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#D8BFD8">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#E6E6FA">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#F8F8FF">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#F0F8FF">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#F5FFFA">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#F0FFF0">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FAFAD2">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFFACD">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFF8DC">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFFFE0">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFFFF0">&nbsp;</td>
  </tr>
  <tr>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFFAF0">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FAF0E6">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FDF5E6">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FAEBD7">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFE4C4">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFDAB9">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFEFD5">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFF5EE">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFF0F5">&nbsp;</td>
	<td onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" style="background-color:#FFE4E1">&nbsp;</td>
  </tr>
  <tr>
	<td colspan="10" style="height=15px;font-family: verdana; font-size:10px;" onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)">&nbsp;None</td>
  </tr>
</table>
</DIV>
<!-- end color menu -->
<!-- Special Char Menu -->
<DIV ID="charMenu" STYLE="display:none;">
<table cellpadding="1" cellspacing="5" border="1" bordercolor="#666666" style="cursor: hand;font-family: Verdana; font-size: 14px; font-weight: bold; BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
  <tr> 
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&copy;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&reg;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#153;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&pound;</td>
  </tr>
  <tr> 
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#151;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#133;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&divide;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&aacute;</td>
  </tr>
  <tr> 
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&yen;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&euro;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#147;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#148;</td>
  </tr>
  <tr> 
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#149;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&para;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&eacute;</td>
    <td style="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&uacute;</td>
  </tr>
</table>
</DIV>
<!-- end char menu -->
<DIV ID="contextMenu" style="display:none;">
<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <tr id=cmCut onClick ='parent.document.execCommand("Cut");parent.oPopup2.hide()'>
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtCut%>&nbsp;</td>
  </tr>
  <tr id=cmCopy onClick ='parent.document.execCommand("Copy");parent.oPopup2.hide()'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtCopy%>&nbsp;</td>
  </tr>
  <tr id=cmPaste onClick ='parent.document.execCommand("Paste");parent.oPopup2.hide()'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtPaste%>&nbsp;</td>
  </tr>
</table>
</div>

<DIV ID="cmTableMenu" style="display:none">
<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <tr onClick ='parent.ModifyTable();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtTableModify%>...&nbsp;</td>
  </tr>
</table>
</DIV>

<DIV ID="cmTableFunctions" style="display:none">
<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <tr onClick ='parent.ModifyCell();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtCellModify%>...&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <tr onClick ='parent.InsertColBefore(); parent.oPopup2.hide();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertColB%>&nbsp;</td>
  </tr>
  <tr onClick ='parent.InsertColAfter(); parent.oPopup2.hide();'> 
   <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertColA%>&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <tr onClick ='parent.InsertRowAbove(); parent.oPopup2.hide();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertRowA%>&nbsp;</td>
  </tr>
  <tr onClick ='parent.InsertRowBelow(); parent.oPopup2.hide();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertRowB%>&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <tr onClick ='parent.DeleteRow(); parent.oPopup2.hide();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtDeleteRow%>&nbsp;</td>
  </tr>
  <tr onClick ='parent.DeleteCol(); parent.oPopup2.hide();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtDeleteCol%>&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <tr onClick ='parent.IncreaseColspan(); parent.oPopup2.hide();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtIncreaseColSpan%>&nbsp;</td>
  </tr>
  <tr onClick ='parent.DecreaseColspan(); parent.oPopup2.hide();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp<%=sTxtDecreaseColSpan%>&nbsp;</td>
  </tr>
</table>
</DIV>

<DIV ID="cmImageMenu" style="display:none">
<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <tr onClick ='parent.doImage();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtModifyImage%>...&nbsp;</td>
  </tr>
</table>
</DIV>

<DIV ID="cmLinkMenu" style="display:none">
<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <tr onClick ='parent.doLink();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtHyperLink%>...&nbsp;</td>
  </tr>
</table>
</DIV>

<DIV ID="cmSpellMenu" style="display:none">
<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <tr onClick ='parent.spellCheck();'> 
    <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" class="parent.toolbutton" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;Check Spelling...&nbsp;</td>
  </tr>
</table>
</DIV>

<!-- Start Paste Menu -->
<DIV ID="pasteMenu" STYLE="display:none">
<table border="0" cellspacing="0" cellpadding="0" width=180 style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
  <tr onClick="parent.doCommand('Paste');"> 
    <td height=20 style="cursor: hand; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
&nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtPaste%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ctrl+V </td>
  </tr>
  <tr onClick="parent.pasteWord();"> 
    <td height=20 style="cursor: hand; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtPasteWord%>&nbsp;&nbsp;&nbsp;&nbsp;Ctrl+D </td>
  </tr>
</table>
</div>
<!-- End Paste Menu -->