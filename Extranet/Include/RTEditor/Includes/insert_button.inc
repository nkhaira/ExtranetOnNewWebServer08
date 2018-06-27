<script language=JavaScript>
window.onload = this.focus

function InsertButton() {
	error = 0
	var sel = window.opener.foo.document.selection;
	if (sel!=null) {
		var rng = sel.createRange();
	   	if (rng!=null) {

			name = document.buttonForm.button_name.value
			value = document.buttonForm.button_value.value
			type = document.buttonForm.button_type[buttonForm.button_type.selectedIndex].text

			if (value != "") {
				value = ' value="' + value + '"'
			} else {
				value = ""
			}

			if (name != "") {
				name = ' name="' + name + '"'
			} else {
				name = ""
			}

			HTMLTextField = '<input type="' + type + '"' + name + value + '>'
			rng.pasteHTML(HTMLTextField)
		} // End if
	} // End If

	if (error != 1) {
		window.opener.foo.focus();
		self.close();
	}
} // End function

document.onkeydown = function () { 
			if (event.keyCode == 13) {	// ENTER
				InsertButton()
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

<form name=buttonForm>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td width="15"><img src="_images/1x1.gif" width="15" height="1"></td>
	  <td class="heading1">[sTxtInsertButton]</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">[sTxtInsertButtonInst]<br>
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
		    <td>&nbsp;&nbsp;[sTxtInsertButton]</td>
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
		    <td class="body" width="80">[sTxtName]:</td>
			<td class="body" width="200">
			  <input type="text" name="button_name" size="10" class="Text90" maxlength="50">
		  </td>
			<td class="body" width="80">[sTxtInitialValue]:</td>
			<td class="body">
			  <input type="text" name="button_value" size="10" class="Text150">
			</td>
		  </tr>
		  
		  <tr>
			<td class="body" width="80">[sTxtType]:</td>
			<td class="body">
			  <select name="button_type" class=text70>
			    <option selected>Submit
			    <option>Reset
			    <option>Button</option>
			  </select>
			</td>
			<td class="body" width="80">&nbsp;</td>
			<td class="body">&nbsp;</td>
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
	    <input type="button" name="insertButton" value="[sTxtInsertButton]" class="Text90" onClick="javascript:InsertButton();">
	<input type="button" name="Submit" value="[sTxtCancel]" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>