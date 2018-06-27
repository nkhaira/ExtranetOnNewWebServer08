<script language=JavaScript>
window.onload = this.focus

function getLink() {
		if (window.opener.foo.document.selection.type == "Control") {
			var oControlRange = window.opener.foo.document.selection.createRange();
			if (oControlRange(0).tagName.toUpperCase() == "IMG") {
				var oSel = oControlRange(0).parentNode;
			} else {
				alert("Link can only be created on an Image or Text")
			}
		} else {
			oSel = window.opener.foo.document.selection.createRange().parentElement();
		}

		if (oSel.tagName.toUpperCase() == "A")
		{
			document.linkForm.targetWindow.value = oSel.target
			document.linkForm.link.value = oSel.getAttribute("href",2)
		}
	}

	function InsertLink() {
		targetWindow = document.linkForm.targetWindow.value;
		var linkSource = document.linkForm.link.value

		if (linkSource != "")
		{
			var oNewLink = window.opener.foo.document.createElement("<A>");
			oNewSelection = window.opener.foo.document.selection.createRange()

			if (window.opener.foo.document.selection.type == "Control")
			{
				selectedImage = window.opener.foo.document.selection.createRange()(0);
				selectedImage.width = selectedImage.width
				selectedImage.height = selectedImage.height
			}

			oNewSelection.execCommand("CreateLink",false,linkSource);

			if (window.opener.foo.document.selection.type == "Control")
			{
				oLink = oNewSelection(0).parentNode;
			} else
				oLink = oNewSelection.parentElement()

			if (targetWindow != "")
			{
				oLink.target = targetWindow;
			} else
				oLink.removeAttribute("target")

			window.opener.foo.focus();
			window.opener.showLink()
			self.close();
		} else {
			alert("[sTxtLinkErr]")
			document.linkForm.link.focus()
		}
	}

	function CreateLink(LinkSource) {
		document.linkForm.link.value = LinkSource;
		document.linkForm.link.focus()
	}

	function RemoveLink() {
		if (window.opener.foo.document.selection.type == "Control")
		{
			selectedImage = window.opener.foo.document.selection.createRange()(0);
			selectedImage.width = selectedImage.width
			selectedImage.height = selectedImage.height
		}

		window.opener.foo.document.execCommand("Unlink");
		window.opener.foo.focus();
		window.opener.showLink()
		self.close();
	}

	function getAnchors() {
		var allLinks = window.opener.foo.document.body.getElementsByTagName("A");
		for (a=0; a < allLinks.length; a++) {
				if (allLinks[a].href.toUpperCase() == "") {
					document.write("<option value=#" + allLinks[a].name + ">" + allLinks[a].name + "</option>")
				}
		}
	}
</script>

<link rel="stylesheet" href="style.css" type="text/css">
<FORM METHOD=POST name=linkForm>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
	  <td width="15"><img src="de_images/1x1.gif" width="15" height="1"></td>
	  <td class="heading1">[sTxtLinkManager]</td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
	  <td class="body">[sTxtLinkManagerInst]<br>
		[sTxtCloseWin]</td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
	  <td class="body">&nbsp;</td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
	  <td class="body">
		<table width="98%" border="0" cellspacing="0" cellpadding="0" class="bevel1">
  		  <tr>
		    <td>&nbsp;&nbsp;[sTxtLinkInfo]</td>
		  </tr>
	    </table>
	  </td>
	</tr>
	<tr>
	  <td colspan="2" height="10"><img src="de_images/1x1.gif" width="1" height="10"></td>
	</tr>
    <tr>
	  <td>&nbsp;</td>
	  <td class="body">
		<table border="0" cellspacing="0" cellpadding="5" width="98%" class="bevel2">
	  	  
		  <tr>
			<td class=body width="100">[sTxtUrl]:</td>
			<td class=body>
			  <input type="text" name="link" value="" class="Text220">
			</td>
		  </tr>
		  <tr>
			<td class=body>[sTxtTargetWin]:</td>
			<td class=body>
			  <input type="text" name="targetWindow" value="" class="Text90">
			  <select name="targetText" class="Text90" onChange="targetWindow.value = targetText[targetText.selectedIndex].value; targetText.value = ''; targetWindow.focus()">
			  <option value=""></option>
	  		  <option value="">None</option>
			  <option value=_blank>_blank</option>
			  <option value=_parent>_parent</option>
  			  <option value=_self>_self</option>
   			  <option value=_top>_top</option>
			  </select></td>
			</td>
		  </tr>
		  <tr>
		  <td class=body>[sTxtAnch]:</td>
		  <td class=body>
			  <select name="targetAnchor" class="Text90" onChange="link.value = targetAnchor[targetAnchor.selectedIndex].value; targetAnchor.value = ''; link.focus()">
				<option value=""></option>
				<script>getAnchors()</script>
			  </select></td>
		  </tr>
	    </table>
	  </td>
	</tr>
  <tr>
	<td colspan="2"><img src="de_images/1x1.gif" width="1" height="10"></td>
  </tr>
	<tr>
	  <td>&nbsp;</td>
	  <td>	
		<input type="button" name="insertLink" value="[sTxtInsertLink]" class="Text90" onClick="javascript:InsertLink();">
		<input type="button" name="removeLink" value="[sTxtRemoveLink]" class="Text90" onClick="javascript:RemoveLink();">
		<input type=button name="Cancel" value="[sTxtCancel]" class="Text70" onClick="javascript:window.close()">
		<input type=hidden name=newdir value="<%=newdir%>">
		</td>
	</tr>
	
  </table>
	</form>
	<script>getLink()</script>