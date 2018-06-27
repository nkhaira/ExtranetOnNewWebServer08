<SCRIPT>

var tickImg = new Image
tickImg.src = "de/de_images/button_tick.gif"

var tickImg2 = new Image
tickImg.src = "de/de_images/button_tick_inverted.gif"

</SCRIPT>

<TABLE WIDTH="100%" CELLSPACING="0" CELLPADDING="0" CLASS=TOOLBAR>
	<TR>
  	<TD CLASS="body" HEIGHT="22">
    
      <!-- Preview View -->
      
    	<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0" CLASS="hide" ALIGN="center" ID="toolbar_preview">
    		<TR>
          <%
          ' Total Top Buttons = 28.  If all turned off, do not display Top Menu Bar Table
          if e__numTopHidden < 28  and e__numBottomHidden < 16 then
          %>
    		  <TD CLASS="body" HEIGHT="57">
          <% else %>
    		  <TD CLASS="body" HEIGHT="28">
          <% end if %>          
      		  &nbsp;&nbsp;&nbsp;<B>Preview Mode</B>
    		  </TD>
        </TR>
    	</TABLE>
      
      <!-- Source View -->
      
      <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0" CLASS="hide" ALIGN="center" ID="toolbar_code">
        <% 'if e__hideFullScreen <> true or _
           '   e__hideCut <> true or _
           '   e__hideCopy <> true or _
           '   e__hidePaste <> true or _
           '   e__hideFind <> true or _
           '   e__hideUndo <> true or _
           '   e__hideRedo <> true then
        %>
    		<TR>
		      <TD CLASS="body" HEIGHT="22">
      		  <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="1">
      			  <TR ID=DE>
      				  <% if e__hideFullScreen <> true then %>
      					<TD>
      						<IMG ID=FULLSCREEN2 BORDER="0" SRC="de/de_images/button_fullscreen.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='TOGGLESIZE();FOO.FOCUS();' TITLE="<%=sTxtFullscreen%>" class=toolbutton>
      					</TD>
      					<% end if %>
          
                <% if e__hideCut <> true then %>
        				<TD>
        				  <IMG BORDER="0" DISABLED ID="toolbarCut2_off" SRC="de/de_images/button_cut_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtCut%> (Ctrl+X)" class=toolbutton><IMG BORDER="0" ID="toolbarCut2_on" SRC="de/de_images/button_cut.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Cut");FOO.FOCUS();' TITLE="<%=sTxtCut%> (Ctrl+X)" class=toolbutton style="display:none">
        				</TD>
                <% end if %>
        
                <% if e__hideCopy <> true then %>
        				<TD>
        				  <IMG BORDER="0" DISABLED ID="toolbarCopy2_off" SRC="de/de_images/button_copy_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton><IMG BORDER="0" ID="toolbarCopy2_on" SRC="de/de_images/button_copy.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Copy");FOO.FOCUS();' TITLE="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton style="display:none">
        				</TD>
                <% end if %>
        
                <% if e__hidePaste <> true then %>
				        <TD>
        				  <IMG BORDER="0" DISABLED ID="toolbarPasteButton2_off" SRC="de/de_images/button_paste_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtPaste%> (Ctrl+V)" class=toolbutton><IMG BORDER="0" ID="toolbarPasteButton2_on" SRC="de/de_images/button_paste.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Paste");FOO.FOCUS();' TITLE="<%=sTxtPaste%> (Ctrl+V)" class=toolbutton style="display:none">
        				</TD>
                <% end if %>
        
				        <% if e__hideFind <> true then %>
                <TD>
      	  			  <IMG BORDER="0" SRC="de/de_images/button_find.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='SHOWFINDDIALOG();FOO.FOCUS();' TITLE="<%=sTxtFindReplace%>" class=toolbutton>
      			  	</TD>
                <% end if %>
          
                <% if e__hideFullScreen <> true or _
                      e__hideCut <> true or _
                      e__hideCopy <> true or _
                      e__hidePast <> true or _
                      e__hideFind <> true then %>
                      
        				<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
                <% end if %>
        
                <% if e__hideUndo <> true then %>
        				<TD>
				          <IMG BORDER="0" DISABLED ID="undo2_off" SRC="de/de_images/button_undo_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton><IMG BORDER="0" ID="undo2_on" SRC="de/de_images/button_undo.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Undo");' TITLE="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton style="display:none">
        				</TD>
                <% end if %>
        
                <% if e__hideRedo <> true then %>
        				<TD>
        				  <IMG BORDER="0" DISABLED ID="redo2_off" SRC="de/de_images/button_redo_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton><IMG BORDER="0" ID="redo2_on" SRC="de/de_images/button_redo.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Redo");' TITLE="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton style="display:none">
        				</TD>
                <% end if %>
              </TR>
			      </TABLE>
    		  </TD>
		    </TR>

    		<TR>
    		  <TD CLASS="body" BGCOLOR="#808080"><IMG SRC="de/de_images/1x1.gif" WIDTH="1" HEIGHT="1"></TD>
    		</TR>
    		<TR>
    		  <TD CLASS="body" BGCOLOR="#FFFFFF"><IMG SRC="de/de_images/1x1.gif" WIDTH="1" HEIGHT="1"></TD>
    		</TR>
        <% 'end if %>
  		  <TR>
          <TD HEIGHT=28>&nbsp;</TD>
        </TR>
      </TABLE>
      
      <!-- Edit View -->
  
    	<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0" CLASS="bevel3" ALIGN="center" ID="toolbar_full">

        <%
        ' Total Top Buttons = 28.  If all turned off, do not display Top Menu Bar Table
'        if e__numTopHidden < 28 then
        %>

		    <TR>
    		  <TD CLASS="body" HEIGHT="22">
          
      			<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="1">
      			  <TR ID=DE>

      				  <% if e__hideFullScreen <> true then %>
      					<TD>
      						<IMG ID=FULLSCREEN BORDER="0" SRC="de/de_images/button_fullscreen.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='TOGGLESIZE();FOO.FOCUS();' TITLE="<%=sTxtFullscreen%>" class=toolbutton>
      					</TD>
      				  <% end if %>
          
                <% if e__hideCut <> true then %>
      					<TD>
			      			<IMG BORDER="0" DISABLED ID="toolbarCut_off" SRC="de/de_images/button_cut_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtCut%> (Ctrl+X)" class=toolbutton><IMG BORDER="0" ID="toolbarCut_on" SRC="de/de_images/button_cut.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Cut");FOO.FOCUS();' TITLE="<%=sTxtCut%> (Ctrl+X)" class=toolbutton style="display:none"><DIV CLASS="pasteArea" ID="myTempArea" CONTENTEDITABLE></DIV>
      					</TD>
                <% end if %>
          
                <% if e__hideCopy <> true then %>
      					<TD>
      						<IMG BORDER="0" DISABLED ID="toolbarCopy_off" SRC="de/de_images/button_copy_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton><IMG BORDER="0" ID="toolbarCopy_on" SRC="de/de_images/button_copy.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Copy");FOO.FOCUS();' TITLE="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton style="display:none">
      					</TD>
                <% end if %>
          
                <% if e__hidePaste <> true then %>
      					<TD>
      						<IMG ID=TOOLBARPASTEBUTTON_OFF DISABLED CLASS=TOOLBUTTON WIDTH="21" HEIGHT="20" SRC="de/de_images/button_paste_disabled.gif" BORDER=0 UNSELECTABLE="on" TITLE="<%=sTxtPaste%> (Ctrl+V)"><IMG ID=TOOLBARPASTEBUTTON_ON CLASS=TOOLBUTTON onMouseDown="button_down(this);" onMouseOver="button_over(this); button_over(toolbarPasteDrop_on)" onClick="doCommand('Paste'); foo.focus()" onMouseOut="button_out(this); button_out(toolbarPasteDrop_on);" WIDTH="21" HEIGHT="20" SRC="de/de_images/button_paste.gif" BORDER=0 UNSELECTABLE="on" TITLE="<%=sTxtPaste%> (Ctrl+V)" style="display:none"><IMG ID=TOOLBARPASTEDROP_OFF DISABLED CLASS=TOOLBUTTON WIDTH="7" HEIGHT="20" SRC="de/de_images/button_drop_menu_disabled.gif" BORDER=0 UNSELECTABLE="on"><IMG ID=TOOLBARPASTEDROP_ON CLASS=TOOLBUTTON onMouseDown="button_down(this);" onMouseOver="button_over(this); button_over(toolbarPasteButton_on)" onClick="showMenu('pasteMenu',180,42)" onMouseOut="button_out(this); button_out(toolbarPasteButton_on);" WIDTH="7" HEIGHT="20" SRC="de/de_images/button_drop_menu.gif" BORDER=0 UNSELECTABLE="on" STYLE="display:none">
      					</TD>
                <% end if %>
          
                <% if e__hideFind <> true then %>
      					<TD>
      					  <IMG BORDER="0" SRC="de/de_images/button_find.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='SHOWFINDDIALOG();FOO.FOCUS();' TITLE="<%=sTxtFindReplace%>" class=toolbutton>
      					</TD>
                <% end if %>
          
                <% if e__hideFullScreen <> true or _
                      e__hideCut <> true or _
                      e__hideCopy <> true or _
                      e__hidePaste <> true or _
                      e__hideFind <> true then
                %>
      					<TD>
      						<IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20">
      					</TD>
                <% end if %>

                <% if e__hideUndo <> true then %>
      					<TD>
      						<IMG ID="undo_off" DISABLED UNSELECTABLE="on" BORDER="0" SRC="de/de_images/button_undo_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton>
      						<IMG ID="undo_on" UNSELECTABLE="on" BORDER="0" SRC="de/de_images/button_undo.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='GOHISTORY(-1);' TITLE="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton style="display:none">
      					</TD>
                <% end if %>
          
                <% if e__hideRedo <> true then %>
      					<TD>
      						<IMG ID="redo_off" DISABLED UNSELECTABLE="on" BORDER="0" SRC="de/de_images/button_redo_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton>
      						<IMG ID="redo_on" UNSELECTABLE="on" BORDER="0" SRC="de/de_images/button_redo.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='GOHISTORY(1);' TITLE="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton style="display:none">
      					</TD>
                <% end if %>
          
                <% if e__hideUndo <> true or e__hideRedo <> true then %>
  		    			<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
                <% end if %>

          		  <% if e__hideSpelling <> true then %>
      					<TD>
      						<IMG ID="toolbarSpell" BORDER="0" SRC="de/de_images/button_spellcheck.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='SPELLCHECK();' TITLE="Check Spelling (F7)" CLASS=TOOLBUTTON>
      					</TD>
      				  <TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
      				  <% end if %>

      				  <% if e__hideRemoveTextFormatting <> true then %>
      					<TD>
      						<IMG BORDER="0" SRC="de/de_images/button_remove_format.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("RemoveFormat");' TITLE="<%=sTxtRemoveFormatting%>" class=toolbutton>
      					</TD>
      					<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
      				  <% end if %>
          
      				  <% if e__hideBold <> true then %>
      					<TD>
      						<IMG ID="fontBold" BORDER="0" SRC="de/de_images/button_bold.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Bold");FOO.FOCUS();' TITLE="<%=sTxtBold%> (Ctrl+B)" class=toolbutton>
      					</TD>
      				  <% end if %>
          
      				  <% if e__hideUnderline <> true then %>
      					<TD>
      						<IMG ID="fontUnderline" BORDER="0" SRC="de/de_images/button_underline.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Underline");FOO.FOCUS();' TITLE="<%=sTxtUnderline%> (Ctrl+U)" class=toolbutton>
      					</TD>
      				  <% end if %>
          
      				  <% if e__hideItalic <> true then %>
      					<TD>
      						<IMG ID="fontItalic" BORDER="0" SRC="de/de_images/button_italic.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Italic");FOO.FOCUS();' TITLE="<%=sTxtItalic%> (Ctrl+I)" class=toolbutton>
      					</TD>
      				  <% end if %>

      				  <% if e__hideStrikethrough <> true then %>
      					<TD>
      						<IMG ID="fontStrikethrough" BORDER="0" SRC="de/de_images/button_strikethrough.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Strikethrough");FOO.FOCUS();' TITLE="<%=sTxtStrikethrough%>" class=toolbutton>
      					</TD>
        				<% end if %>

                <% if e__hideBold <> true or _
                      e__hideUnderline <> true or _
                      e__hideItalic <> true or _
                      e__hideStrikethrough <> true then%>
                      
      					<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
                <% end if %>


      				  <% if e__hideNumberList <> true then %>
      					<TD>
      						<IMG ID="fontInsertOrderedList" BORDER="0" SRC="de/de_images/button_numbers.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("InsertOrderedList");FOO.FOCUS();' TITLE="<%=sTxtNumList%>" class=toolbutton>
      					</TD>
      				  <% end if %>
          
      				  <% if e__hideBulletList <> true then %>
      					<TD>
      						<IMG ID="fontInsertUnorderedList" BORDER="0" SRC="de/de_images/button_bullets.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("InsertUnorderedList");FOO.FOCUS();' TITLE="<%=sTxtBulletList%>" class=toolbutton>
      					</TD>
      				  <% end if %>
          
      				  <% if e__hideDecreaseIndent <> true then %>
      					<TD>
        					<IMG BORDER="0" SRC="de/de_images/button_decrease_indent.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Outdent");FOO.FOCUS();' TITLE="<%=sTxtDecreaseIndent%>" class=toolbutton>
      					</TD>
      				  <% end if %>
          
      				  <% if e__hideIncreaseIndent <> true then %>
      					<TD>
      						<IMG BORDER="0" SRC="de/de_images/button_increase_indent.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("Indent");FOO.FOCUS();' TITLE="<%=sTxtIncreaseIndent%>" class=toolbutton>
      					</TD>
      					<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
      				  <% end if %>
          
      				  <% if e__hideSuperScript <> true then %>
      					<TD>
      						<IMG ID="fontSuperScript" BORDER="0" SRC="de/de_images/button_superscript.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("superscript");FOO.FOCUS();' TITLE="<%=sTxtSuperscript%>" class=toolbutton>
      					</TD>
      				  <% end if %>
          
      				  <% if e__hideSubScript <> true then %>
      					<TD>
      						<IMG ID="fontSubScript" BORDER="0" SRC="de/de_images/button_subscript.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("subscript");FOO.FOCUS();' TITLE="<%=sTxtSubscript%>" class=toolbutton>
      					</TD>
      					<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
      				  <% end if %>
          
      				  <% if e__hideLeftAlign <> true then %>
      					<TD>
      						<IMG ID="fontJustifyLeft" BORDER="0" SRC="de/de_images/button_align_left.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("JustifyLeft");FOO.FOCUS();' TITLE="<%=sTxtAlignLeft%>" class=toolbutton>
      					</TD>
      				  <% end if %>
          
      				  <% if e__hideCenterAlign <> true then %>
      					<TD>
      						<IMG ID="fontJustifyCenter" BORDER="0" SRC="de/de_images/button_align_center.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("JustifyCenter");FOO.FOCUS();' TITLE="<%=sTxtAlignCenter%>" class=toolbutton>
      					</TD>
      				  <% end if %>
          
      				  <% if e__hideRightAlign <> true then %>
      					<TD>
      						<IMG ID="fontJustifyRight" BORDER="0" SRC="de/de_images/button_align_right.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("JustifyRight");FOO.FOCUS();' TITLE="<%=sTxtAlignRight%>" class=toolbutton>
      					</TD>
      				  <% end if %>
                
      				  <% if e__hideJustify <> true then %>
      					<TD>
      						<IMG ID="fontJustifyFull" BORDER="0" SRC="de/de_images/button_align_justify.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("JustifyFull");FOO.FOCUS();' TITLE="<%=sTxtAlignJustify%>" class=toolbutton>
      					</TD>
      					<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
      				  <% end if %>
          
      				  <% if e__hideLink <> true then %>
      					<TD>
      						<IMG DISABLED ID="toolbarLink_off" BORDER="0" SRC="de/de_images/button_link_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtHyperLink%>" class=toolbutton><IMG ID="toolbarLink_on" BORDER="0" SRC="de/de_images/button_link.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOLINK()' TITLE="<%=sTxtHyperLink%>" class=toolbutton style="display:none">
      					</TD>
      				  <% end if %>
          
      				  <% if e__hideMailLink <> true then %>
      					<TD>
      						<IMG BORDER="0" ID="toolbarEmail_off" DISABLED SRC="de/de_images/button_email_disabled.gif" WIDTH="21" HEIGHT="20" TITLE="<%=sTxtEmail%>" class=toolbutton><IMG BORDER="0" ID="toolbarEmail_on" SRC="de/de_images/button_email.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOEMAIL()' TITLE="<%=sTxtEmail%>" class=toolbutton style="display:none">
      					</TD>
      				  <% end if %>
          
      				  <% if e__hideAnchor <> true then %>
      					<TD>
      						<IMG BORDER="0" SRC="de/de_images/button_anchor.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOANCHOR()' TITLE="<%=sTxtAnchor%>" class=toolbutton>
      					</TD>
      					<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
      				  <% end if %>
          
      				  <% if e__hideHelp <> true then %>
      					<TD>
      						<IMG BORDER="0" SRC="de/de_images/button_help.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOHELP()' TITLE="<%=sTxtHelp%>" class=toolbutton>
      					</TD>
      				  <% end if %>
          
      			  </TR>
	  	    	</TABLE>
    		  </TD>
		    </TR>

			  <!--TR>
    		  <TD CLASS="body" BGCOLOR="#808080">1<IMG SRC="de/de_images/1x1.gif" WIDTH="1" HEIGHT="1"></TD>
		    </TR>
        
    		<TR>
		      <TD CLASS="body" BGCOLOR="#FFFFFF">2<IMG SRC="de/de_images/1x1.gif" WIDTH="1" HEIGHT="1"></TD>
    		</TR-->
        
        <% 'end if %>
                
    		<TR>
		      <TD CLASS="body">

      			<TABLE BORDER="0" CELLSPACING="1" CELLPADDING="1">
      			  <TR ID=DE>
        				<% if e__hideFont <> true then %>
        				<TD>
        				  <SELECT ID="fontDrop" onChange="doFont(this[this.selectedIndex].value)" CLASS="Text120" UNSELECTABLE="on">
        					<%= BuildFontList() %>
        				  </SELECT>
        				</TD>
        				<% end if %>
            
        				<% if e__hideSize <> true then %>
        				<TD>
        				  <SELECT ID="sizeDrop" onChange="doSize(this[this.selectedIndex].value)" CLASS=TEXT50 UNSELECTABLE="on">
        					<%= BuildSizeList() %>
      	  			  </SELECT>
        				</TD>
        				<% end if %>
            
        				<% if e__hideFormat <> true then %>
        				<TD>
        				  <SELECT ID="formatDrop" onChange="doFormat(this[this.selectedIndex].value)" CLASS="Text70" UNSELECTABLE="on">
        				    <OPTION SELECTED><%=sTxtFormat%>
        				    <OPTION VALUE="<P>">Normal
        				    <OPTION VALUE="<H1>">Heading 1
        				    <OPTION VALUE="<H2>">Heading 2
        				    <OPTION VALUE="<H3>">Heading 3
        				    <OPTION VALUE="<H4>">Heading 4
        				    <OPTION VALUE="<H5>">Heading 5
        				    <OPTION VALUE="<H6>">Heading 6
    		    		  </SELECT>
        				</TD>
    		    		<% end if %>
    				
                <% if e__hideStyle <> true then %>
        				<TD>
        				  <SELECT ID="sStyles" onChange="applyStyle(this[this.selectedIndex].value);foo.focus();this.selectedIndex=0;foo.focus();" CLASS="Text90" UNSELECTABLE="on" onmouseenter="doStyles()">
        				    <OPTION SELECTED><%=sTxtStyle%></OPTION>
        				    <OPTION VALUE="">None</OPTION>
        				  </SELECT>
        				</TD>
    		    		<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
        				<% end if %>
    				
                <% if e__hideForeColor <> true then %>
        				<TD>
        				  <IMG BORDER="0" SRC="de/de_images/button_font_color.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="(isAllowed()) ? showMenu('colorMenu',180,291) : foo.focus()" CLASS=TOOLBUTTON TITLE="<%=sTxtColour%>">
        				</TD>
        				<% end if %>
    				
                <% if e__hideBackColor <> true then %>
        				<TD>
        				  <IMG BORDER="0" SRC="de/de_images/button_highlight.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="(isAllowed()) ? showMenu('colorMenu2',180,291) : foo.focus()" CLASS=TOOLBUTTON TITLE="<%=sTxtBackColour%>">
        				</TD>
        				<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
        				<% end if %>
    				
                <% if e__hideTable <> true then %>
        				<TD ID=TOOLBARTABLES>
        				  <IMG BORDER="0" SRC="de/de_images/button_table_down.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="(isAllowed()) ? showMenu('tableMenu',160,262) : foo.focus()" CLASS=TOOLBUTTON TITLE="<%=sTxtTableFunctions%>">
        				</TD>
        				<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
    		    		<% end if %>
    				
                <% if e__hideForm <> true then %>
        				<TD>
        				  <IMG CLASS=TOOLBUTTON onMouseDown=BUTTON_DOWN(THIS); onMouseOver=BUTTON_OVER(THIS); onClick="(isAllowed()) ? showMenu('formMenu',180,189) : foo.focus()" onMouseOut=BUTTON_OUT(THIS); TYPE=IMAGE WIDTH="21" HEIGHT="20" SRC="de/de_images/button_form_down.gif" BORDER=0 TITLE="<%=sTxtFormFunctions%>">
        				</TD>
        				<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
        				<% end if %>
    				
                <% if e__hideImage <> true then %>
        				<TD>
        				  <IMG BORDER="0" SRC="de/de_images/button_image.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="doImage()" CLASS=TOOLBUTTON TITLE="<%=sTxtImage%>">
        				</TD>
        				<TD><IMG SRC="de/de_images/seperator.gif" WIDTH="2" HEIGHT="20"></TD>
        				<% end if %>
    				
                <% if e__hideTextBox <> true then %>
        				<TD>
        				  <IMG BORDER="0" SRC="de/de_images/button_textbox.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="doTextbox()" CLASS=TOOLBUTTON TITLE="<%=sTxtTextbox%>">
        				</TD>
        				<% end if %>
    				
                <% if e__hideHorizontalRule <> true then %>
        				<TD>
        				  <IMG BORDER="0" SRC="de/de_images/button_hr.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='DOCOMMAND("InsertHorizontalRule");FOO.FOCUS();' TITLE="<%=sTxtInsertHR%>" class=toolbutton>
        				</TD>
        				<% end if %>
    				
                <% if e__hideSymbols <> true then %>
        				<TD>
        				  <IMG BORDER="0" SRC="de/de_images/button_chars.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="(isAllowed()) ? showMenu('charMenu',104,111) : foo.focus()" CLASS=TOOLBUTTON TITLE="<%=sTxtChars%>">
        				</TD>
        				<% end if %>
    				
                <% if e__hideProps <> true then %>
        				<TD>
        				  <IMG BORDER="0" SRC="de/de_images/button_properties.gif" WIDTH="21" HEIGHT="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="ModifyProperties()" CLASS=TOOLBUTTON TITLE="<%=sTxtPageProperties%>">
        				</TD>
        				<% end if %>
    
        				<% if e__hideClean <> true then %>
        				<TD>				
        				  <IMG CLASS=TOOLBUTTON onmousedown="button_down(this);" onmouseover="button_over(this);" onClick="cleanCode()" onmouseout="button_out(this);" TYPE=IMAGE WIDTH="21" HEIGHT="20" SRC="de/de_images/button_clean_code.gif" BORDER=0 TITLE="<%=sTxtCleanCode%>">
        				</TD>
        				<% end if %>
    
        				<% if e__hasCustomInserts = true then %>
        				<TD>
        					<IMG CLASS=TOOLBUTTON onmousedown="button_down(this);" onmouseover="button_over(this);" onClick="doCustomInserts()" onmouseout="button_out(this);" TYPE=IMAGE WIDTH="21" HEIGHT="20" SRC="de/de_images/button_custom_inserts.gif" BORDER=0 TITLE="<%=sTxtCustomInserts%>">
        				</TD>
        				<% end if %>
    
        				<% if e__hideAbsolute <> true then %>
        				<TD>
        					<IMG ID="fontAbsolutePosition_off" DISABLED CLASS=TOOLBUTTON onmousedown="button_down(this);" onmouseover="button_over(this);" WIDTH="21" HEIGHT="20" SRC="de/de_images/button_absolute_disabled.gif" BORDER=0 TITLE="<%=sTxtTogglePosition%>"><IMG ID="fontAbsolutePosition" CLASS=TOOLBUTTON onmousedown="button_down(this);" onmouseover="button_over(this);" onClick="doCommand('AbsolutePosition')" onmouseout="button_out2(this);" TYPE=IMAGE WIDTH="21" HEIGHT="20" SRC="de/de_images/button_absolute.gif" BORDER=0 TITLE="<%=sTxtTogglePosition%>" style="display:none">
        				</TD>
        				<% end if %>
    
        				<% if e__hideGuidelines <> true then %>
        				<TD>
        				  <IMG CLASS=TOOLBUTTON onMouseDown="button_down(this);" onMouseOver="button_over(this);" onClick="toggleBorders()" onMouseOut="button_out2(this);" TYPE=IMAGE WIDTH="21" HEIGHT="20" SRC="de/de_images/button_show_borders.gif" BORDER=0 TITLE="<%=sTxtToggleGuidelines%>" id=guidelines>
        				</TD>
        				<% end if %>
                
    		  	  </TR>
      			</TABLE>
    		  </TD>
		    </TR>
    	</TABLE> 
    </TD>
  </TR> 
</TABLE>

<!-- Table Menu -->

<DIV ID="tableMenu" STYLE="display:none">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH=160 STYLE="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" BGCOLOR="threedface">
  <TR onClick="parent.ShowInsertTable()" TITLE="<%=sTxtTable%>" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20> 
      &nbsp;&nbsp;<%=sTxtTable%>...&nbsp; </TD>
  </TR>
  <TR onClick=PARENT.MODIFYTABLE(); TITLE="<%=sTxtTableModify%>" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20 ID=MODIFYTABLE> 
	  &nbsp;&nbsp;<%=sTxtTableModify%>...&nbsp;</TD>
  </TR>
  <TR TITLE="<%=sTxtCellModify%>" onClick=parent.ModifyCell() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20 ID=MODIFYCELL> 
	&nbsp;&nbsp;<%=sTxtCellModify%>...&nbsp; </TD>
  </TR>
  <TR HEIGHT=10> 
    <TD ALIGN=CENTER><IMG SRC="de/de_images/vertical_spacer.gif" WIDTH="140" HEIGHT="2"></TD>
  </TR>
  <TR TITLE="<%=sTxtInsertColA%>" onClick=parent.InsertColAfter() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20 ID=COLAFTER> 
      &nbsp;&nbsp;<%=sTxtInsertColA%>&nbsp;
    </TD>
  </TR>
  <TR TITLE="<%=sTxtInsertColB%>" onClick=parent.InsertColBefore() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20 ID=COLBEFORE> 
      &nbsp;&nbsp;<%=sTxtInsertColB%>&nbsp;
    </TD>
  </TR>
  <TR HEIGHT=10> 
    <TD ALIGN=CENTER><IMG SRC="de/de_images/vertical_spacer.gif" WIDTH="140" HEIGHT="2"></TD>
  </TR>
  <TR TITLE="<%=sTxtInsertRowA%>" onClick=parent.InsertRowAbove() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
	<TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20 ID=ROWABOVE> 
      &nbsp;&nbsp;<%=sTxtInsertRowA%>&nbsp;
    </TD>
  </TR>
  <TR TITLE="<%=sTxtInsertRowB%>" onClick=parent.InsertRowBelow() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
	<TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20 ID=ROWBELOW> 
      &nbsp;&nbsp;<%=sTxtInsertRowB%>&nbsp;
    </TD>
  </TR>
  <TR HEIGHT=10> 
    <TD ALIGN=CENTER><IMG SRC="de/de_images/vertical_spacer.gif" WIDTH="140" HEIGHT="2"></TD>
  </TR>
  <TR TITLE="<%=sTxtDeleteRow%>" onClick=parent.DeleteRow() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20 ID=DELETEROW>
      &nbsp;&nbsp;<%=sTxtDeleteRow%>&nbsp;
    </TD>
  </TR>
  <TR TITLE="<%=sTxtDeleteCol%>" onClick=parent.DeleteCol() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20 ID=DELETECOL>
      &nbsp;&nbsp;<%=sTxtDeleteCol%>&nbsp;
    </TD>
  </TR>
  <TR HEIGHT=10> 
    <TD ALIGN=CENTER><IMG SRC="de/de_images/vertical_spacer.gif" WIDTH="140" HEIGHT="2" TABINDEX=1 HIDEFOCUS></TD>
  </TR>
  <TR TITLE="<%=sTxtIncreaseColSpan%>" onClick=parent.IncreaseColspan() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20 ID=INCREASESPAN>
      &nbsp;&nbsp;<%=sTxtIncreaseColSpan%>&nbsp;
    </TD>
  </TR>
  <TR TITLE="<%=sTxtDecreaseColSpan%>" onClick=parent.DecreaseColspan() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=20 ID=DECREASESPAN>
      &nbsp;&nbsp;<%=sTxtDecreaseColSpan%>&nbsp;
    </TD>
  </TR>
</TABLE>
</DIV>

<!-- End Table Menu -->

<!-- Form Menu -->

<DIV ID="formMenu" STYLE="display:none;">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH=180 STYLE="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" BGCOLOR="threedface">
  <TR TITLE="<%=sTxtForm%>" onClick=parent.insertForm() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22>
      <IMG WIDTH="21" HEIGHT="20" SRC="de/de_images/button_form.gif" BORDER=0 ALIGN="absmiddle">&nbsp;<%=sTxtForm%>...&nbsp;</TD>
  </TR>
  <TR TITLE="<%=sTxtFormModify%>" onClick=parent.modifyForm() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" ID="modifyForm1" HEIGHT=22 CLASS=DROPDOWN>
      <IMG ID="modifyForm2" WIDTH="21" HEIGHT="20" SRC="de/de_images/button_modify_form.gif" BORDER=0 ALIGN="absmiddle">&nbsp;<%=sTxtFormModify%>...&nbsp;</TD>
  </TR>
  <TR HEIGHT=10> 
    <TD ALIGN=CENTER><IMG SRC="de/de_images/vertical_spacer.gif" WIDTH="140" HEIGHT="2" TABINDEX=1 HIDEFOCUS></TD>
  </TR>
  <TR TITLE="<%=sTxtTextField%>" onClick=parent.doTextField() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22>
      <IMG WIDTH="21" HEIGHT="20" SRC="de/de_images/button_textfield.gif" BORDER=0 ALIGN="absmiddle">&nbsp;<%=sTxtTextField%>...&nbsp;</TD>
  </TR>
  <TR TITLE="<%=sTxtTextArea%>" onClick=parent.doTextArea() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22>
      <IMG TYPE=IMAGE WIDTH="21" HEIGHT="20" SRC="de/de_images/button_textarea.gif" BORDER=0 ALIGN="absmiddle">&nbsp;<%=sTxtTextArea%>...&nbsp;</TD>
  </TR>
  <TR TITLE="<%=sTxtHidden%>" onClick=parent.doHidden(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22>
      <IMG WIDTH="21" HEIGHT="20" SRC="de/de_images/button_hidden.gif" BORDER=0 ALIGN="absmiddle">&nbsp;<%=sTxtHidden%>...&nbsp;</TD>
  </TR>
  <TR TITLE="<%=sTxtButton%>" onClick=parent.doButton(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22>
      <IMG WIDTH="21" HEIGHT="20" SRC="de/de_images/button_button.gif" BORDER=0 ALIGN="absmiddle">&nbsp;<%=sTxtButton%>...&nbsp;</TD>
  </TR>
  <TR TITLE="<%=sTxtCheckbox%>" onClick=parent.doCheckbox(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22>
      <IMG WIDTH="21" HEIGHT="20" SRC="de/de_images/button_checkbox.gif" BORDER=0 ALIGN="absmiddle">&nbsp;<%=sTxtCheckbox%>...&nbsp;</TD>
  </TR>
  <TR TITLE="<%=sTxtRadioButton%>" onClick=parent.doRadio(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22>
      <IMG WIDTH="21" HEIGHT="20" SRC="de/de_images/button_radio.gif" BORDER=0 ALIGN="absmiddle">&nbsp;<%=sTxtRadioButton%>...&nbsp;</TD>
  </TR>
</TABLE>
</DIV>

<!-- End Form Menu -->

<!-- Zoom Menu -->

<DIV ID="zoomMenu" STYLE="display:none;">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH=65 STYLE="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" BGCOLOR="threedface">
  <TR onClick=PARENT.DOZOOM(500) onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom500_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom500_,0);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22 ID="zoom500_">
     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;500%&nbsp;</TD>
  </TR>
  <TR onClick=PARENT.DOZOOM(200) onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom200_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom200_,0);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22 ID="zoom200_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;200%&nbsp;</TD>
  </TR>
  <TR onClick=PARENT.DOZOOM(150) onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom150_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom150_,0);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22 ID="zoom150_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;150%&nbsp;</TD>
  </TR>
  <TR onClick="parent.doZoom(100)" onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom100_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom100_,0)";">
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22 ID="zoom100_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;100%&nbsp;</TD>
  </TR>
  <TR onClick=PARENT.DOZOOM(75); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom75_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom75_,0);">
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22 ID="zoom75_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;75%&nbsp;</TD>
  </TR>
  <TR onClick=PARENT.DOZOOM(50); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom50_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom50_,0);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22 ID="zoom50_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;50%&nbsp;</TD>
  </TR>
  <TR onClick=PARENT.DOZOOM(25); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom25_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom25_,0);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22 ID="zoom25_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;25%&nbsp;</TD>
  </TR>
  <TR onClick=PARENT.DOZOOM(10); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom10_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom10_,0);"> 
    <TD STYLE="cursor: hand; font:8pt tahoma;" HEIGHT=22 ID="zoom10_">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10%&nbsp;</TD>
  </TR>
</TABLE>
</DIV>

<!-- End Zoom Menu -->

<!-- Color Menu -->

<DIV ID="colorMenu" STYLE="display:none;">
<TABLE CELLPADDING="1" CELLSPACING="5" BORDER="1" BORDERCOLOR="#666666" STYLE="cursor: hand;font-family: Verdana; font-size: 7px; BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" BGCOLOR="threedface">
  <TR>
	<TD COLSPAN="10" ID=COLOR STYLE="height=20px;font-family: verdana; font-size:12px;">&nbsp;</TD>
  </TR>
  <TR>
    <TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FF0000;width=12px">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFFF00;width=12px">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#00FF00;width=12px">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#00FFFF;width=12px">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#0000FF;width=12px">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FF00FF;width=12px">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFFFFF;width=12px">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#F5F5F5;width=12px">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#DCDCDC;width=12px">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFFAFA;width=12px">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#D3D3D3">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#C0C0C0">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#A9A9A9">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#808080">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#696969">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#000000">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#2F4F4F">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#708090">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#778899">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#4682B4">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#4169E1">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#6495ED">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#B0C4DE">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#7B68EE">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#6A5ACD">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#483D8B">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#191970">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#000080">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#00008B">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#0000CD">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#1E90FF">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#00BFFF">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#87CEFA">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#87CEEB">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#ADD8E6">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#B0E0E6">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#F0FFFF">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#E0FFFF">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#AFEEEE">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#00CED1">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#5F9EA0">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#48D1CC">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#00FFFF">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#40E0D0">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#20B2AA">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#008B8B">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#008080">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#7FFFD4">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#66CDAA">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#8FBC8F">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#3CB371">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#2E8B57">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#006400">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#008000">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#228B22">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#32CD32">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#00FF00">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#7FFF00">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#7CFC00">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#ADFF2F">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#98FB98">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#90EE90">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#00FF7F">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#00FA9A">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#556B2F">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#6B8E23">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#808000">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#BDB76B">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#B8860B">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#DAA520">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFD700">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#F0E68C">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#EEE8AA">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFEBCD">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFE4B5">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#F5DEB3">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFDEAD">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#DEB887">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#D2B48C">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#BC8F8F">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#A0522D">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#8B4513">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#D2691E">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#CD853F">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#F4A460">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#8B0000">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#800000">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#A52A2A">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#B22222">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#CD5C5C">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#F08080">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FA8072">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#E9967A">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFA07A">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FF7F50">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FF6347">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FF8C00">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFA500">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FF4500">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#DC143C">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FF0000">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FF1493">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FF00FF">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FF69B4">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFB6C1">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFC0CB">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#DB7093">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#C71585">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#800080">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#8B008B">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#9370DB">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#8A2BE2">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#4B0082">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#9400D3">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#9932CC">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#BA55D3">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#DA70D6">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#EE82EE">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#DDA0DD">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#D8BFD8">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#E6E6FA">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#F8F8FF">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#F0F8FF">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#F5FFFA">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#F0FFF0">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FAFAD2">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFFACD">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFF8DC">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFFFE0">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFFFF0">&nbsp;</TD>
  </TR>
  <TR>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFFAF0">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FAF0E6">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FDF5E6">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FAEBD7">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFE4C4">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFDAB9">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFEFD5">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFF5EE">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFF0F5">&nbsp;</TD>
	<TD onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)" STYLE="background-color:#FFE4E1">&nbsp;</TD>
  </TR>
  <TR>
	<TD COLSPAN="10" STYLE="height=15px;font-family: verdana; font-size:10px;" onMouseOver="parent.showColor(color,this)" onClick="parent.doColor(color)">&nbsp;None</TD>
  </TR>
</TABLE>
</DIV>

<!-- End Color Menu -->

<!-- Special Character Menu -->

<DIV ID="charMenu" STYLE="display:none;">
<TABLE CELLPADDING="1" CELLSPACING="5" BORDER="1" BORDERCOLOR="#666666" STYLE="cursor: hand;font-family: Verdana; font-size: 14px; font-weight: bold; BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" BGCOLOR="threedface">
  <TR> 
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&copy;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&reg;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#153;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&pound;</TD>
  </TR>
  <TR> 
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#151;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#133;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&divide;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&aacute;</TD>
  </TR>
  <TR> 
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&yen;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&euro;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#147;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#148;</TD>
  </TR>
  <TR> 
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&#149;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&para;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&eacute;</TD>
    <TD STYLE="width=15px; cursor: hand;" onClick="parent.insertChar(this)" onMouseOver="parent.button_over(this);" onMouseOut="parent.char_out(this);" onMouseDown="parent.button_down(this);">&uacute;</TD>
  </TR>
</TABLE>
</DIV>

<!-- End Special Character Menu -->

<!-- Context Menu -->

<DIV ID="contextMenu" STYLE="display:none;">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <TR ID=CMCUT onClick ='PARENT.DOCUMENT.EXECCOMMAND("Cut");PARENT.OPOPUP2.HIDE()'>
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtCut%>&nbsp;</TD>
  </TR>
  <TR ID=CMCOPY onClick ='PARENT.DOCUMENT.EXECCOMMAND("Copy");PARENT.OPOPUP2.HIDE()'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtCopy%>&nbsp;</TD>
  </TR>
  <TR ID=CMPASTE onClick ='PARENT.DOCUMENT.EXECCOMMAND("Paste");PARENT.OPOPUP2.HIDE()'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtPaste%>&nbsp;</TD>
  </TR>
</TABLE>
</DIV>

<!-- End Context Menu -->

<!-- Table Menu -->

<DIV ID="cmTableMenu" STYLE="display:none">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <TR ONCLICK ='PARENT.MODIFYTABLE();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtTableModify%>...&nbsp;</TD>
  </TR>
</TABLE>
</DIV>

<DIV ID="cmTableFunctions" STYLE="display:none">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <TR ONCLICK ='PARENT.MODIFYCELL();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtCellModify%>...&nbsp;</TD>
  </TR>
</TABLE>
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <TR ONCLICK ='PARENT.INSERTCOLBEFORE(); PARENT.OPOPUP2.HIDE();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertColB%>&nbsp;</TD>
  </TR>
  <TR ONCLICK ='PARENT.INSERTCOLAFTER(); PARENT.OPOPUP2.HIDE();'> 
   <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertColA%>&nbsp;</TD>
  </TR>
</TABLE>
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <TR ONCLICK ='PARENT.INSERTROWABOVE(); PARENT.OPOPUP2.HIDE();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertRowA%>&nbsp;</TD>
  </TR>
  <TR ONCLICK ='PARENT.INSERTROWBELOW(); PARENT.OPOPUP2.HIDE();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertRowB%>&nbsp;</TD>
  </TR>
</TABLE>
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <TR ONCLICK ='PARENT.DELETEROW(); PARENT.OPOPUP2.HIDE();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtDeleteRow%>&nbsp;</TD>
  </TR>
  <TR ONCLICK ='PARENT.DELETECOL(); PARENT.OPOPUP2.HIDE();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtDeleteCol%>&nbsp;</TD>
  </TR>
</TABLE>
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <TR ONCLICK ='PARENT.INCREASECOLSPAN(); PARENT.OPOPUP2.HIDE();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtIncreaseColSpan%>&nbsp;</TD>
  </TR>
  <TR ONCLICK ='PARENT.DECREASECOLSPAN(); PARENT.OPOPUP2.HIDE();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp<%=sTxtDecreaseColSpan%>&nbsp;</TD>
  </TR>
</TABLE>
</DIV>

<!-- End Table Menu -->

<!-- Image Menu -->

<DIV ID="cmImageMenu" STYLE="display:none">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <TR ONCLICK ='PARENT.DOIMAGE();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtModifyImage%>...&nbsp;</TD>
  </TR>
</TABLE>
</DIV>

<!-- End Image Menu -->

<!-- Link Menu -->

<DIV ID="cmLinkMenu" STYLE="display:none">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <TR ONCLICK ='PARENT.DOLINK();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtHyperLink%>...&nbsp;</TD>
  </TR>
</TABLE>
</DIV>

<!-- End Link Menu -->

<!-- Spell Menu -->

<DIV ID="cmSpellMenu" STYLE="display:none">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
  <TR ONCLICK ='PARENT.SPELLCHECK();'> 
    <TD STYLE="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" CLASS="parent.toolbutton" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;Check Spelling...&nbsp;</TD>
  </TR>
</TABLE>
</DIV>

<!-- End Spell Menu -->

<!-- Start Paste Menu -->

<DIV ID="pasteMenu" STYLE="display:none">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH=180 STYLE="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" BGCOLOR="threedface">
  <TR onClick="parent.doCommand('Paste');"> 
    <TD HEIGHT=20 STYLE="cursor: hand; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
&nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtPaste%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ctrl+V </TD>
  </TR>
  <TR onClick="parent.pasteWord();"> 
    <TD HEIGHT=20 STYLE="cursor: hand; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
      &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtPasteWord%>&nbsp;&nbsp;&nbsp;&nbsp;Ctrl+D </TD>
  </TR>
</TABLE>
</DIV>

<!-- End Paste Menu -->