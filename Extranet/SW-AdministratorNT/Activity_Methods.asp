<HTML>
<HEAD>
  <TITLE>Asset Activity Methods</TITLE>
  <LINK REL=STYLESHEET HREF="/sw-common/SW-Style.css">
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso8859-1">
  <META NAME="LANG" CONTENT="ENGLISH">
  <META AUTHOR="Kelly Whitlock - Kelly.Whitlock@fluke.com">
</HEAD>

<BODY BGCOLOR="White" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" LINK="#000000" VLINK="#000000" ALINK="#000000">
<%
response.write "<DIV ALIGN=CENTER>"
response.write "<P>"
response.write "<TABLE CELLPADDING=4 CELLSPACING=0 BORDER=0>"
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">OLV</FONT></TD><TD CLASS=Small HEIGHT=8>On-Line View</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">OVD</FONT></TD><TD CLASS=Small HEIGHT=8>On-Line Download</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">OVS</FONT></TD><TD CLASS=Small HEIGHT=8>On-Line Send via Email (Attachment)</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">SSV</FONT></TD><TD CLASS=Small HEIGHT=8>Subscription Service View</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">SSD</FONT></TD><TD CLASS=Small HEIGHT=8>Subscription Service Download</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">SSS</FONT></TD><TD CLASS=Small HEIGHT=8>Subscription Service Send via Email (Attachment)</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">OLL</FONT></TD><TD CLASS=Small HEIGHT=8>On-Line Web Page Link</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">OLG</FONT></TD><TD CLASS=Small HEIGHT=8>On-Line Gateway to Auxiliary Application</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">EEF</FONT></TD><TD CLASS=Small HEIGHT=8>Electronic Email Fulfillment (Oracle Fulfillment)</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">EDL</FONT></TD><TD CLASS=Small HEIGHT=8>Electronic Digital Library (www.fluke.com)</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=SmallBold HEIGHT=8 BGCOLOR=""Black"" ALIGN=CENTER><FONT COLOR=""#FFCC00"">LOS</FONT></TD><TD CLASS=Small HEIGHT=8>Literature Order System (DCG)</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small COLSPAN=2><B>Unique Users</B> - Count of Unique User Accounts for the Reporting Period.</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small COLSPAN=2><B>Navigation Clicks</B> - Count of all Navigation Clicks for the Reporting Period.</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small COLSPAN=2><B>Asset Activity</B> - Count of Assets where one of the above Methods was used to acquire the asset.</TD></TR>" & vbCrLf
response.write "<TR><TD CLASS=Small COLSPAN=2><B>Extranet vs. Total</B> - Total minus EEF and EDL count of assets equals count for extrnet activity.</TD></TR>" & vbCrLf
response.write "</TABLE>"
response.write "</DIV>"
%>
</BODY>
</HTML>
