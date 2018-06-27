<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>

<%
response.write("Post data output<BR><BR>")

for each foo in request.form
	response.write("Post data: " & foo & "<BR>")
next
%>

</body>
</html>
