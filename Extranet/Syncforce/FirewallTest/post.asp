<html>
<body>

<body onload="alert('here');return document.foo_form.submit();">
<form action="http://216.9.4.31:90/response.asp" method="POST" name="foo_form">
	<input type="textbox" name="foo_1" value="foo_1_test">
	<input type="textbox" name="foo_2" value="foo_2_test">
	<input type="textbox" name="foo_3" value="foo_3_test">
	<input type="textbox" name="foo_4" value="foo_4_test">
</form>

</body>
</html>