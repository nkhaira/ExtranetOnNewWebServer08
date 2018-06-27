<script language=JavaScript>
window.onload = this.focus

function InsertForm() {
	error = 0
	var sel = window.opener.document.selection;
	if (sel!=null) {
		var rng = sel.createRange();
	   	if (rng!=null) {
			name = document.formForm.form_name.value
			action = document.formForm.form_action.value
			method = document.formForm.form_method[formForm.form_method.selectedIndex].text

        		if (error != 1) {

				if (method != "None") {
					method = ' method="' + method + '"'
				} else {
					method = ""
				}

				if (name != "") {
					name = ' name="' + name + '"'
				} else {
					name = ""
				}

				if (action != "") {
					action = ' action="' + action + '"'
				} else {
					action = ""
				}

        			HTMLForm = "<form id=ewp_element_to_style" + name + action + method +">&nbsp;</form>"
         			rng.pasteHTML(HTMLForm)

					oForm = window.opener.foo.document.getElementById("ewp_element_to_style")
					
					if (window.opener.borderShown == "yes") {
						oForm.runtimeStyle.border = "1px dotted #FF0000"
					}

					oForm.removeAttribute("id")


        		}
		}
	
	}
	
	if (error != 1) {
		window.opener.foo.focus();
		self.close();
	}
}

document.onkeydown = function () { 
			if (event.keyCode == 13) {	// ENTER
				InsertForm()
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
	  <td class="heading1">[sTxtInsertForm]</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">[sTxtInsertFormInst]<br>
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
		    <td>&nbsp;&nbsp;[sTxtInsertForm]</td>
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
			  <select class=text70 name=form_method>
				<option selected>None
				<option>Post
				<option>Get</option>
				</select>
			</td>
		  </tr>
	    </table>
	</td>
  </tr>
  <tr>
	<td colspan="2"><img src="ew_images/1x1.gif" width="1" height="10"></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td>
	    <input type="button" name="insertForm" value="[sTxtInsertForm]" class="Text90" onClick="javascript:InsertForm();">
	<input type="button" name="Submit" value="[sTxtCancel]" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>