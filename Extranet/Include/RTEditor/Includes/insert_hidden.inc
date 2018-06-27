<script language=JavaScript>
window.onload = this.focus
var error
function InsertHiddenField() {
	var sel = window.opener.document.selection;
	if (sel!=null) {
		var rng = sel.createRange();
	   	if (rng!=null) {
			name = document.hiddenForm.hidden_name.value
			value = document.hiddenForm.hidden_value.value

		error = 0

		if (error != 1) {
				if (value != "None") {
					value = ' value="' + value + '"'
				} else {
					value = ""
				}

				if (name != "") {
					name = ' name="' + name + '"'
				} else {
					name = ""
				}

        		HTMLTextField = '<input id=ewp_element_to_style type=hidden' + name + value + '>'
         		rng.pasteHTML(HTMLTextField)

				oHidden = window.opener.foo.document.getElementById("ewp_element_to_style")
					
				if (window.opener.borderShown == "yes") {
					oHidden.runtimeStyle.border = "1px dashed #000000"
					oHidden.runtimeStyle.width = "15px"
					oHidden.runtimeStyle.height = "15px"
					oHidden.runtimeStyle.backgroundColor = "#FDADAD"
					oHidden.runtimeStyle.color = "#FDADAD"
				}

				oHidden.removeAttribute("id")


		} // End if
		} // End if
	} // End If

	if (error != 1) {
		window.opener.foo.focus();
		self.close();
	}
} // End function

document.onkeydown = function () { 
			if (event.keyCode == 13) {	// ENTER
				InsertHiddenField()
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
	  <td class="heading1">[sTxtInsertHidden]</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">[sTxtInsertHiddenInst]<br>
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
		    <td>&nbsp;&nbsp;[sTxtInsertHidden]</td>
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
	    <input type="button" name="insertHiddenField" value="[sTxtInsertHidden]" class="Text120" onClick="javascript:InsertHiddenField();">
	<input type="button" name="Submit" value="[sTxtCancel]" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>