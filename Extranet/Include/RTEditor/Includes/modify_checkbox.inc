<script language=JavaScript>
var myPage = window.opener;
window.onload = setValues;

var checkboxName = myPage.selectedCheckbox.name;
var checkboxValue = myPage.selectedCheckbox.value;
var checkboxType = myPage.selectedCheckbox.checked; // true or false
var checkboxClass = myPage.selectedCheckbox.className;

function setValues() {

	if (checkboxClass != "") {
		checkboxClass = " class=" + checkboxClass
	}

	checkboxForm.checkbox_value.value = checkboxValue;
	checkboxForm.checkbox_name.value = checkboxName;
	this.focus();
}

function doModify() {
	var sel = window.opener.document.selection;
		if (sel!=null) {
			var rng = sel.createRange();
		}

		name = document.checkboxForm.checkbox_name.value
		value = document.checkboxForm.checkbox_value.value
		type = document.checkboxForm.checkbox_type[checkboxForm.checkbox_type.selectedIndex].text

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

   		HTMLTextField = '<input type=checkbox' + name + value + type + checkboxClass + '>'
   		myPage.selectedCheckbox.outerHTML = HTMLTextField
    
    window.close()
}

function printType() {
	if ((checkboxType != undefined) || (checkboxType != "")) {
		if (checkboxType == false) { 
			checkboxType = "Unchecked"
		}

		if (checkboxType == true) {
			checkboxType = "Checked"
		}

		document.write('<option selected>' + checkboxType)
		document.write('<option>Checked')
		document.write('<option>Unchecked</option>')
	}
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

<form name=checkboxForm>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td width="15"><img src="_images/1x1.gif" width="15" height="1"></td>
	  <td class="heading1">[sTxtModifyCheckbox]</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">[sTxtModifyCheckboxInst]<br>
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
		    <td>&nbsp;&nbsp;[sTxtModifyCheckbox]</td>
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
			  <input type="text" name="checkbox_name" size="10" class="Text90" maxlength="50">
		  </td>
			<td class="body" width="80">[sTxtInitialValue]:</td>
			<td class="body">
			  <input type="text" name="checkbox_value" size="10" class="Text150">
			</td>
		  </tr>
		  
		  <tr>
			<td class="body" width="80">[sTxtInitialState]:</td>
			<td class="body">
			  <select name="checkbox_type" class=text90>
				<script>printType()</script>
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
	    <input type="button" name="modifyCheckbox" value="[sTxtModifyCheckbox]" class="Text90" onClick="javascript:doModify();">
	<input type="button" name="Submit" value="[sTxtCancel]" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>