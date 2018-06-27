<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<%
Dim Oracle_ID_1, Oracle_ID_2
Oracle_ID_1="1272029"
Oracle_ID_2="9820000"
%>

<body>
<A HREF="Javascript:void();" LANGUAGE="JavaScript" onclick="openit('http://dtmevtvsdv15/find_it.asp?Document=<%=Oracle_ID_1%>&Style=82','Vertical');return false;">Pop-Up Link</A> Small File
<BR>
<A HREF="Javascript:void();" LANGUAGE="JavaScript" onclick="openit('http://dtmevtvsdv15/find_it.asp?Document=<%=Oracle_ID_2%>&Style=82','Vertical');return false;">Pop-Up Link</A> Big File
</body>
</html>
<SCRIPT LANGUAGE="JavaScript">
function openit(DaURL, orient) {
  var File_Load;
  File_Load = window.open(DaURL,"File_Load","status=yes,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=no,links=no");
//  File_Load.close();
}
</SCRIPT>
