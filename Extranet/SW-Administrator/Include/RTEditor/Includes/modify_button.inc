<script language=JavaScript>
var myPage = window.opener;
window.onload = setValues;

var buttonName = myPage.selectedButton.name;
var buttonValue = myPage.selectedButton.value;
var buttonType = myPage.selectedButton.type;
var buttonClass = myPage.selectedButton.className;

function setValues() {

	if (buttonClass != "") {
		buttonClass = " class=" + buttonClass
	}

	buttonForm.button_value.value = buttonValue;
	buttonForm.button_name.value = buttonName;
	this.focus();
}

function doModify() {
	var sel = window.opener.document.selection;
		if (sel!=null) {
			var rng = sel.createRange();
		}

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

   		HTMLTextField = '<input type="' + type + '"' + name + value + buttonClass + '>'
   		myPage.selectedButton.outerHTML = HTMLTextField
    
    window.close()
}

function printType() {
	if ((buttonType != undefined) && (buttonType != "")) {
		document.write('<option selected>' + buttonType)
		document.write('<option>Submit')
		document.write('<option>Reset')
		document.write('<option>Button')
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

<form name=buttonForm>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td width="15"><img src="_images/1x1.gif" width="15" height="1"></td>
	  <td class="heading1">[sTxtModifyButton]</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">[sTxtModifyButtonInst]<br>
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
		    <td>&nbsp;&nbsp;[sTxtModifyButton]</td>
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
	    <input type="button" name="modifyButton" value="[sTxtModifyButton]" class="Text90" onClick="javascript:doModify();">
	<input type="button" name="Submit" value="[sTxtCancel]" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>