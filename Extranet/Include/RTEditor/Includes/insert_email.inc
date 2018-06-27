<script language=JavaScript>
window.onload = this.focus

function getLink() {
		if (window.opener.foo.document.selection.type == "Control") {
			var oControlRange = window.opener.foo.document.selection.createRange();
			if (oControlRange(0).tagName.toUpperCase() == "IMG") {
				var oSel = oControlRange(0).parentNode;
			} else {
				alert("Link can only be created on an Image or Text")
			}
		} else {
			oSel = window.opener.foo.document.selection.createRange().parentElement();
		}

		if (oSel.tagName.toUpperCase() == "A")
		{
			myHref = oSel.getAttribute("href",2)
			if (myHref.indexOf("?") >-1 )
			{
				myHrefEmail = myHref.substring(7, myHref.indexOf("?"))
				myHrefSubject = myHref.substring(myHref.indexOf(myHrefEmail)+myHrefEmail.length+9, myHref.length)
			} else {
				myHrefEmail = myHref.substring(7, myHref.length)
				myHrefSubject = ""
			}

			document.emailForm.email.value = myHrefEmail
			document.emailForm.subject.value = myHrefSubject
		}
}

function InsertEmail() {
	error = 0
	var sel = window.opener.foo.document.selection;
	if (sel!=null) {
		var rng = sel.createRange();
	   	if (rng!=null) {

			if (window.opener.foo.document.selection.type == "Control")
				{
					selectedImage = window.opener.foo.document.selection.createRange()(0);
					selectedImage.width = selectedImage.width
					selectedImage.height = selectedImage.height
				}

			email = document.emailForm.email.value
			subject = document.emailForm.subject.value

        	if (error != 1) {

				if (email == "") {
					alert("[sTxtEmailErr]")
					document.emailForm.email.focus
					error = 1
				} else {
					mailto = "mailto:" + email
					if (subject != "")
					{
						mailto = mailto + "?subject=" + subject
					}
					rng.execCommand("CreateLink",false,mailto)
				}
			}
		}
	}
	
	if (error != 1) {
		window.opener.foo.focus()
		window.opener.showLink()
		self.close();
	}
}

function RemoveLink() {
	if (window.opener.foo.document.selection.type == "Control")
	{
		selectedImage = window.opener.foo.document.selection.createRange()(0);
		selectedImage.width = selectedImage.width
		selectedImage.height = selectedImage.height
	}

	window.opener.foo.document.execCommand("Unlink");
	window.opener.foo.focus();
	window.opener.showLink()
	self.close();
}

document.onkeydown = function () { 
			if (event.keyCode == 13) {	// ENTER
				InsertEmail()
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

<form name=emailForm>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td width="15"><img src="_images/1x1.gif" width="15" height="1"></td>
	  <td class="heading1">[sTxtInsertEmail]</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	  <td class="body">[sTxtInsertEmailInst]<br>
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
		    <td>&nbsp;&nbsp;[sTxtInsertEmail]</td>
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
		    <td class="body" width="90">[sTxtEmailAddress]:</td>
			<td class="body">
			  <input type="text" name="email" size="10" class="Text150" maxlength="150">
		  </td>
		  </tr>
		  <tr>
		    <td class="body" width="80">
			[sTxtSubject]:</td>
			<td class="body">
			  <input type="text" name="subject" size="10" class="Text150">
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
	    <input type="button" name="insertLink" value="[sTxtInsertEmailLink]" class="Text90" onClick="javascript:InsertEmail();">
		<input type="button" name="removeLink" value="[sTxtRemoveEmailLink]" class="Text120" onClick="javascript:RemoveLink();">
	<input type="button" name="Submit" value="[sTxtCancel]" class="Text50" onClick="javascript:window.close()">
	</td>
  </tr>
</table>
</form>
<script>getLink()</script>