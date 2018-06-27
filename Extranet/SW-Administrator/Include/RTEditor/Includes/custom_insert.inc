<script language=JavaScript>
window.onload = doLoad

function doLoad() {
	ci = window.opener.customInserts
	// alert(document.getElementById("myCi").outerHTML)
	
	for (i = 0; i < ci.length; i++) {
		newOption = document.createElement("option");
		newOption.value = ci[i][1]
		newOption.text = ci[i][0]
		document.getElementById("myCi").add(newOption)
	}
}

re = /&lt;/g
re2 = /&gt;/g
function doPreview(myForm) {
	customHTML = myForm.customInserts.value
	previewPane = document.getElementById("p1")
	if (customHTML == "") {
		foo.document.body.innerHTML = ""
	} else {
		customHTML = customHTML.replace(re,"<")
		customHTML = customHTML.replace(re2,">")
		foo.document.body.innerHTML = customHTML
	}
}

function InsertHTML() {
	error = 0
	var sel = window.opener.foo.document.selection;
	if (sel!=null) {
		var rng = sel.createRange();
	   	if (rng!=null) {

			customHTML = foo.document.body.innerHTML

        	if (error != 1) {
				if (customHTML == "") {
					alert("[sTxtCustomInsertErr]")
					error = 1
				} else {
					rng.pasteHTML(customHTML)
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
				InsertHTML()
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

<form name=customForm>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td width="15"><img src="_images/1x1.gif" width="15" height="1"></td>
	  <td class="heading1">[sTxtCustomInserts]</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">[sTxtCustomInsertInst]<br>
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
		    <td>&nbsp;&nbsp;[sTxtCustomInsert]</td>
		</tr>
	  </table>
	</td>
  </tr>
  <tr>
	<td colspan="2"><img src="ew_images/1x1.gif" width="1" height="10"></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td class="body">
	    <table border="0" cellspacing="0" cellpadding="5" width="98%" class="bevel2">
		  <tr> 
			<td class="body" width="90" valign=top>[sTxtCustomInsert]:</td>
			<td class="body"> 
			  <select name="customInserts" id=myCi size="4" class=text220 onChange="doPreview(this.form)">
			  </select>
			</td>
		  </tr>
		  <tr>
			<td class="body" width="90" valign=top>[sTxtPreview]:</td>
			<td class="body"><iframe id="foo" src="" border="1" style="width:220px;height:60px;"></iframe></td>
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
	    <input type="button" name="insertHTML" value="[sTxtInsertHTML]" class="Text90" onClick="javascript:InsertHTML();">
	<input type="button" name="Submit" value="[sTxtCancel]" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>