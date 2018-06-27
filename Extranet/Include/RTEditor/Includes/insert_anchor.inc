<script language=JavaScript>
window.onload = this.focus

window.onload = doLoad;

function doLoad() {
		document.anchorForm.anchor_name.focus()
}

function InsertAnchor() {
	error = 0
	var sel = window.opener.foo.document.selection;
	if (sel!=null) {
		var rng = sel.createRange();
	   	if (rng!=null) {

			name = document.anchorForm.anchor_name.value

        	if (error != 1) {
				if (name == "") {
					alert("[sTxtInsertAnchorErr]")
					document.anchorForm.anchor_name.focus
					error = 1
				} else {
					
					rng.pasteHTML("<a id=ewp_element_to_style name=" + anchorForm.anchor_name.value + ">")
					oAnchor = window.opener.foo.document.getElementById("ewp_element_to_style")
					
					if (window.opener.borderShown == "yes") {
						oAnchor.runtimeStyle.border = "1px dashed #000000"
						oAnchor.runtimeStyle.width = "20px"
						oAnchor.runtimeStyle.height = "16px"
						oAnchor.runtimeStyle.backgroundColor = "#FFFFCC"
						oAnchor.runtimeStyle.color = "#FFFFCC"	
					}

					oAnchor.removeAttribute("id")
				}
			}
		}
	}
	
	if (error != 1) {
		window.opener.foo.focus()
		self.close();
	}
}

document.onkeydown = function () { 
			if (event.keyCode == 13) {	// ENTER
				InsertAnchor()
			}
};

document.onkeypress = onkeyup = function () {
	if (event.keyCode == 13) {	// ENTER
	event.cancelBubble = true;
	event.returnValue = false;
	return false;			
	}
};

</script>

<form name=anchorForm>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td width="15"><img src="_images/1x1.gif" width="15" height="1"></td>
	  <td class="heading1">[sTxtInsertAnchor]</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">[sTxtInsertAnchorInst]<br>
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
		    <td>&nbsp;&nbsp;[sTxtInsertAnchor]</td>
		</tr>
	  </table>
	</td>
  </tr>
  <tr>
	<td colspan="2"><img src="de_images/1x1.gif" width="1" height="10"></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td class="body">
	    <table border="0" cellspacing="0" cellpadding="5" width="98%" class="bevel2">
		  <tr>
		    <td class="body" width="90">[sTxtInsertAnchorName]:</td>
			<td class="body">
			  <input type="text" name="anchor_name" size="10" class="Text150" maxlength="150">
		  </td>
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
	    <input type="button" name="insertAnchor" value="[sTxtInsertAnchor]" class="Text90" onClick="javascript:InsertAnchor();">
	<input type="button" name="Submit" value="[sTxtCancel]" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>