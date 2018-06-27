<%@ Language="VBScript" CODEPAGE="65001" %>
<HTML>
<HEAD>
  <TITLE>Asset Container Status Codes and Colors</TITLE>
  <LINK REL=STYLESHEET HREF="/sw-common/SW-Style.css">
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
  <META NAME="LANG" CONTENT="ENGLISH">
  <META AUTHOR="Kelly Whitlock - Kelly.Whitlock@fluke.com">
</HEAD>

<BODY BGCOLOR="White" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" LINK="#000000" VLINK="#000000" ALINK="#000000">
<%
response.write "<DIV ALIGN=CENTER>"
response.write "<TABLE CELLPADDING=2 CELLSPACING=0 BORDER=0>"
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">L</FONT></TD><TD CLASS=SmallBold HEIGHT=8>Link to Web Page</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">A</FONT></TD><TD CLASS=SmallBold HEIGHT=8>Asset File (Low Resolution)</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">Z</FONT></TD><TD CLASS=SmallBold HEIGHT=8>Asset File (Low Resolution) - Archive (ZIP)</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">P</FONT></TD><TD CLASS=SmallBold HEIGHT=8>Asset File (POD Resolution)</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">T</FONT></TD><TD CLASS=SmallBold HEIGHT=8>Thumbnail Image File</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">U</FONT></TD><TD CLASS=SmallBold HEIGHT=8>End-User Viewable via Email Fulfilment</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">D</FONT></TD><TD CLASS=SmallBold HEIGHT=8>End-User Viewable via Digital Library</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER VALIGN=TOP><FONT COLOR=""#FFCC00"">S</FONT></TD><TD CLASS=SmallBold HEIGHT=8>Shopping Cart<BR>"
response.write "<TABLE CELLPADDING=2 CELLSPACING=0 BORDER=0>"
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#90EE90"" ALIGN=CENTER>Y</TD><TD CLASS=Small HEIGHT=8>Orderable</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#0099FF"" ALIGN=CENTER><U>N</U></TD><TD CLASS=Small HEIGHT=8>Not Orderable</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#FFFF00"" ALIGN=CENTER><U>?</U></TD><TD CLASS=Small HEIGHT=8>Item Not in Oracle Literature DB</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#F5DEB3"" ALIGN=CENTER><U>P</U></TD><TD CLASS=Small HEIGHT=8>PDF Download Only</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#FF8000"" ALIGN=CENTER><U>E</U></TD><TD CLASS=Small HEIGHT=8>Excluded via Asset Container Checkbox</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#AAAAFF"" ALIGN=CENTER><U>R</U></TD><TD CLASS=Small HEIGHT=8>Retired</TD></TR>" & vbCrLf
response.write "</TABLE>"
response.write "</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER VALIGN=TOP><FONT COLOR=""#FFCC00"">O</FONT></TD><TD CLASS=SmallBold HEIGHT=8>Oracle Deliverable Status<BR>"
response.write "<TABLE CELLPADDING=2 CELLSPACING=0 BORDER=0>"
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#FF0000"" ALIGN=CENTER>?</TD><TD CLASS=Small HEIGHT=8>Asset Status <> Oracle Status</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small HEIGHT=8 BGCOLOR=""#FFCCFF"" ALIGN=CENTER>?</TD><TD CLASS=Small HEIGHT=8>Asset Expiration Date has been reached</TD></TR>" & vbCrLf
response.write "</TABLE>"

response.write "</TABLE>"
response.write "</DIV>"
%>
</BODY>
</HTML>
