<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<%
vooty = now()
response.write vooty
response.write "<P>"

vooty = CLng(DateAdd("d",30,Now()))
vooty = 38987
response.write vooty
response.write "<P>"

vooty = CDate(vooty)
response.write vooty
response.write "<P>"

Locator="46O19488O20529O3O37208O38987O0"

%>


</body>
</html>
