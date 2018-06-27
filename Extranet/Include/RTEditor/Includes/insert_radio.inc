<script language=JavaScript>
window.onload = this.focus
var error
function InsertRadio() {
	var sel = window.opener.document.selection;
	if (sel!=null) {
		var rng = sel.createRange();
	   	if (rng!=null) {
			name = document.radioForm.radio_name.value
			value = document.radioForm.radio_value.value
			checked = document.radioForm.radio_type[radioForm.radio_type.selectedIndex].text

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

		if (checked == "Unchecked"){
			checked = ""
		}

		HTMLTextField = '<input type=radio ' + checked + name + value + '>'
		window.opener.foo.focus();
		rng.pasteHTML(HTMLTextField)
		
		} // End if
	} // End If

	if (error != 1) {
		self.close();
	}
} // End function

document.onkeydown = function () { 
			if (event.keyCode == 13) {	// ENTER
				InsertRadio()
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


<form name=radioForm>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td width="15"><img src="_images/1x1.gif" width="15" height="1"></td>
	  <td class="heading1">Insert Radio Button</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">[sTxtInsertRadioInst]<br>
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
		    <td>&nbsp;&nbsp;[sTxtInsertRadio]</td>
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
			  <input type="text" name="radio_name" size="10" class="Text90" maxlength="50">
		  </td>
			<td class="body" width="80">[sTxtInitialValue]:</td>
			<td class="body">
			  <input type="text" name="radio_value" size="10" class="Text150">
			</td>
		  </tr>
		  
		  <tr>
			<td class="body" width="80">[sTxtInitialState]:</td>
			<td class="body">
			  <select name="radio_type" class=text90>
				<option>Checked</option>
				<option selected>Unchecked</option>
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
	    <input type="button" name="insertRadio" value="[sTxtInsertRadio]" class="Text120" onClick="javascript:InsertRadio();">
	<input type="button" name="Submit" value="[sTxtCancel]" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>