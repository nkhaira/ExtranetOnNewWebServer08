<script language=JavaScript>
var myPage = window.opener;
window.onload = setValues;

var hiddenName = myPage.selectedHidden.name;
var hiddenValue = myPage.selectedHidden.value;

function setValues() {

	hiddenForm.hidden_value.value = hiddenValue;
	hiddenForm.hidden_name.value = hiddenName;
	this.focus();
}

function doModify() {
	var sel = window.opener.document.selection;
		if (sel!=null) {
			var rng = sel.createRange();
		}

		name = document.hiddenForm.hidden_name.value
		value = document.hiddenForm.hidden_value.value

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

		HTMLTextField = '<input type=hidden' + name + value + '>'
		myPage.selectedHidden.outerHTML = HTMLTextField
    
    window.close()
}

document.onkeydown = function () { 
			if (event.keyCode == 13) {	// ENTER
				doModify()
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

<form name=hiddenForm>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td width="15"><img src="_images/1x1.gif" width="15" height="1"></td>
	  <td class="heading1">[sTxtModifyField]</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">[sTxtModifyFieldInst]<br>
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
		    <td>&nbsp;&nbsp;[sTxtModifyField]</td>
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
			  <input type="text" name="hidden_name" size="10" class="Text90" maxlength="50">
		  </td>
			<td class="body" width="80">[sTxtInitialValue]:</td>
			<td class="body">
			  <input type="text" name="hidden_value" size="10" class="Text150">
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
	    <input type="button" name="modifyHiddenField" value="[sTxtModifyField]" class="Text120" onClick="javascript:doModify();">
	<input type="button" name="Submit" value="[sTxtCancel]" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>