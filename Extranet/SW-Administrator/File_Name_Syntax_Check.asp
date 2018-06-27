
<!-- TWO STEPS TO INSTALL VALIDATION (NUM. OR CHARS):

  1.  Copy the coding into the HEAD of your HTML document
  2.  Add the last code into the BODY of your HTML document  -->

<!-- STEP ONE: Paste this code into the HEAD of your HTML document  -->

<HEAD>

<SCRIPT LANGUAGE="JavaScript">

<!-- This script and many more are available free online at -->
<!-- The JavaScript Source!! http://javascript.internet.com -->

<!-- Begin
function validate(field) {
var valid = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-"
var ok = "yes";
var temp;
for (var i=0; i<field.value.length; i++) {
temp = "" + field.value.substring(i, i+1);
if (valid.indexOf(temp) == "-1") ok = "no";
}
if (ok == "no") {
alert("Invalid entry!  Only characters and numbers are accepted!");
field.focus();
field.select();
   }
}
//  End -->
</script>
</HEAD>

<!-- STEP TWO: Copy this code into the BODY of your HTML document  -->

<BODY>

<center>
<form>
<input type=text name="entry" onBlur="validate(this)">
<br>
<input type=submit value="Submit Form">
</form>
</center>

<p><center>
<font face="arial, helvetica" size="-2">Free JavaScripts provided<br>
by <a href="http://javascriptsource.com">The JavaScript Source</a></font>
</center><p>

<!-- Script Size:  1.13 KB -->