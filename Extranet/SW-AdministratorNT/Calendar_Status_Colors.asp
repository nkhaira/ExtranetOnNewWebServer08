<HTML>
<HEAD>
  <TITLE>Status Colors</TITLE>
  <LINK REL=STYLESHEET HREF="/sw-common/SW-Style.css">
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso8859-1">
  <META NAME="LANG" CONTENT="ENGLISH">
  <META AUTHOR="David Whitlock - David.Whitlock@fluke.com">
</HEAD>

<BODY BGCOLOR="White" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" LINK="#000000" VLINK="#000000" ALINK="#000000">
<%
response.write "<DIV ALIGN=CENTER>"
response.write "<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0>"
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">&nbsp;&nbsp;ID (Status)&nbsp;&nbsp;</FONT></TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""Yellow"" ALIGN=CENTER>&nbsp;&nbsp;Pending Review&nbsp;&nbsp;</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#00CC00"" ALIGN=CENTER>&nbsp;&nbsp;Live&nbsp;&nbsp;</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#AAAAFF"" ALIGN=CENTER>&nbsp;&nbsp;Archive&nbsp;&nbsp;</TD></TR>" & vbCrLf

response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""White"" ALIGN=CENTER>&nbsp;</TD></TR>" & vbCrLf

response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">&nbsp;&nbsp;Asset Status&nbsp;&nbsp;</FONT></TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""Yellow"" ALIGN=CENTER>&nbsp;&nbsp;Missing&nbsp;&nbsp;</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""Red"" ALIGN=CENTER>&nbsp;&nbsp;Required&nbsp;&nbsp;</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#FF9966"" ALIGN=CENTER>&nbsp;&nbsp;Literature Fulfillment&nbsp;&nbsp;</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#90EE90"" ALIGN=CENTER>&nbsp;&nbsp;Shopping Cart&nbsp;&nbsp;</TD></TR>" & vbCrLf

response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""White"" ALIGN=CENTER>&nbsp;</TD></TR>" & vbCrLf

response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">&nbsp;&nbsp;Date Status&nbsp;&nbsp;</FONT></TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""Yellow"" ALIGN=CENTER>&nbsp;&nbsp;Date not Reached&nbsp;&nbsp;</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#00CC00"" ALIGN=CENTER>&nbsp;&nbsp;Live&nbsp;&nbsp;</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#AAAAFF"" ALIGN=CENTER>&nbsp;&nbsp;Archive&nbsp;&nbsp;</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""Orange"" ALIGN=CENTER>&nbsp;&nbsp;Date Missed&nbsp;&nbsp;</TD></TR>" & vbCrLf

response.write "</TABLE>"
response.write "</DIV>"
%>
</BODY>
</HTML>
