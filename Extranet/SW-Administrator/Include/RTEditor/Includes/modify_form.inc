<script language=javascript>
var myPage = window.opener;

window.onload = setValues;

var formName = myPage.selectedForm.name;
var formAction = myPage.selectedForm.action;
var formMethod = myPage.selectedForm.method;

function setValues() {

	formForm.form_name.value = formName;
	formForm.form_action.value = formAction;
	this.focus();
}

function doModify() {

	if (formForm.form_name.value != "") {
		myPage.selectedForm.name = formForm.form_name.value
	} else {
		myPage.selectedForm.removeAttribute('name',0)
	}

	if (formForm.form_action.value != "") {
		myPage.selectedForm.action = formForm.form_action.value
	} else {
		myPage.selectedForm.removeAttribute('action',0)
	}

	if (formForm.method[formForm.method.selectedIndex].text != "None") {
    	myPage.selectedForm.method = formForm.method[formForm.method.selectedIndex].text
    } else {
		myPage.selectedForm.removeAttribute('method',0)
    }
        
    window.close()
}

function printMethod() {
	if ((formMethod != undefined) && (formMethod != "")) {
		document.write('<option selected>' + formMethod)
		document.write('<option>None')
	} else {
		document.write('<option selected>None')
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

<form name=formForm>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td width="15"><img src="_images/1x1.gif" width="15" height="1"></td>
	  <td class="heading1">[sTxtModifyForm]</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">Enter the required information and click &quot;Modify Form&quot; to modify the form properties of your form.<br>
		Click the &quot;Cancel&quot; Button to  close this window.</td>
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
		    <td>&nbsp;&nbsp;[sTxtModifyForm]</td>
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
			<td class="body">
			  <input type="text" name="form_name" size="10" class="Text90" maxlength="50">
		  </td>
		  </tr>
		  <tr>
		    <td class="body" width="80">[sTxtAction]:</td>
			<td class="body">
			  <input type="text" name="form_action" size="50" class="Text250">
		  </td>
		  </tr>
		  <tr>
			<td class="body" width="80">[sTxtMethod]:</td>
			<td class="body">
			  <SELECT class=text70 name=method>
			    <script>printMethod()</script>
			    <option>Post
			    <option>Get</option>
			  </select>
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
	    <input type="button" name="modifyForm" value="[sTxtModifyForm]" class="Text120" onClick="javascript:doModify();">
	<input type="button" name="Submit" value="Cancel" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>
