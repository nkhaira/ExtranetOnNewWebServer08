
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<HTML>
<HEAD>
	<TITLE>Fluke Service</TITLE>
</HEAD>

<!--#include virtual="/service/center/euro_lock_out.asp"-->

<% if Session("view") <> "2" then %>
<BODY BACKGROUND="/service/center/images/bg.gif" BGCOLOR="White" ALINK="#008400" LINK="#008400" VLINK="#008400">
<% else %>
<BODY BGCOLOR="White" ALINK="#008400" LINK="#008400" VLINK="#008400">
<% end if %>

<% if Session("view") <> "2" then %>

<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH=100%>
<TR><TD ALIGN="RIGHT"><IMG SRC="/service/center/images/flukelogo.gif" WIDTH=143 HEIGHT=50 ALT="" BORDER="0"></TD></TR></TABLE>
<BR>
<TABLE WIDTH=100% BORDER="0" CELLSPACING="0" CELLPADDING="0" ALIGN="LEFT">
<TR>
<TD WIDTH=125 ALIGN="LEFT" VALIGN="TOP">
   <TABLE WIDTH=125 BORDER=0 CELLPADDING=0 CELLSPACING=0>
<!-- START LEFT NAVIGATION ROWS -->
<TR>
<TD WIDTH=125 COLSPAN=3 VALIGN=TOP>
       <IMG SRC="/service/center/images/spacer.gif" WIDTH=125 HEIGHT=1 ALT="" BORDER="0">
</TD>
</TR>
	<TR>
   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
       <A HREF="http://www.fluke.com"><IMG SRC="/service/center/buttons/fluke.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A><BR>
</TD>
	</TR>
	<TR>
   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
       <A HREF="http://www.fluke.com/service/default.asp"><IMG SRC="/service/center/buttons/service.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A><BR>
</TD>
	</TR>
	<TR>
   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
       <A HREF="/service/center/default.asp"><IMG SRC="/service/center/buttons/srvcsuph.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A><BR>
	<IMG SRC="/service/center/images/spacer.gif" WIDTH=1 HEIGHT=15 ALT="" BORDER="0">
</TD>
	</TR>
   <TR>
   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
       <A HREF="/service/center/documents/whatsnew.asp"><IMG SRC="/service/center/buttons/whatsnew.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
   </TD>
	</TR>
   <TR>
   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
       <A HREF="/service/center/default.asp"><IMG SRC="/service/center/buttons/subservice.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
   </TD>
   </TR>
   <TR>
   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
      <A HREF="/service/center/downloads.asp"><IMG SRC="/service/center/buttons/downloads.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
   </TD>
   </TR>
   	   
   <TR>
   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
       <A HREF="/service/center/databases.asp"><IMG SRC="/service/center/buttons/databasesh.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
   </TD>
   </TR>
   <TR>
	   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	      <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
        <A HREF="/service/center/documents/svcindex.asp?region=<%=Session("region")%>&view=<%=Session("view")%>"><IMG SRC="/service/center/buttons/documentation.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
	   </TD>
	   </TR>
	    <TR>
	   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	      <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
        <A HREF="/service/center/products/svcmodel.asp?region=<%=Session("region")%>&view=<%=Session("view")%>"><IMG SRC="/service/center/buttons/support.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
	   </TD>
	   </TR>

     <% if Euro_Lock_Out = false or Euro_Lock_Out = 2 then %>
	   <TR>
	     <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	       <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
         <A HREF="/service/center/parts/prtqueryform.asp?region=<%=request("region")%>&view=<%=request("view")%>"><IMG SRC="/service/center/buttons/parts.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
  	   </TD>
	   </TR>
	   <TR>
	     <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	       <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
         <A HREF="/service/center/parts/RnCqueryform.asp?region=<%=request("region")%>&view=<%=request("view")%>"><IMG SRC="/service/center/buttons/usRnCh.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
  	   </TD>
	   </TR>
      <TR>
	   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	      <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
        <A HREF="/service/center/download/metcal/metcal.asp"><IMG SRC="/service/center/buttons/metcal.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
	   </TD>
	   </TR>
     <% end if %>
     
     <% if Euro_Lock_Out = true or Euro_Lock_Out = 2 then %>
      <TR>
	   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	      <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
        <A HREF="/service/center/download/calnet/calnet.asp"><IMG SRC="/service/center/buttons/calnet.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
	   </TD>
	   </TR>
     <% end if %>

   <TR>
   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
      <A HREF="/service/center/region/default.asp"><IMG SRC="/service/center/buttons/regsup.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
   </TD>
   </TR>
	<TR>
   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
      <A HREF="/service/center/help/helphome.asp"><IMG SRC="/service/center/buttons/help.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
   </TD>
   </TR>
<TR>
   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
      <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=20 HEIGHT=1 ALT="" BORDER="0"><A HREF="mailto:service@fluke.com"><IMG SRC="/service/center/images/buttons/envelope.gif" BORDER=0 ALT="Send E-Mail to Service Engineering Webmaster"></A></NOBR>
   </TD>
   </TR>
   </TABLE>
  </TD>
  <TD WIDTH=100% ALIGN="left" VALIGN="TOP">

<% end if %>

<!--Content-->

<IMG SRC="/service/center/images/headlines/sv-usRnC-Pricing.gif" WIDTH=97 HEIGHT=25 ALT="" BORDER="0"><BR>
<IMG SRC="/service/center/images/headlines/sv-query.gif" WIDTH=255 HEIGHT=30 ALT="" BORDER="0">
<BR><BR>

<FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica">

<UL><FONT COLOR="#FF0000"><%=Session("ErrorString")%></FONT></UL>
<UL><LI>Click on [New Search] to enter new search criteria.</LI></UL>
<UL><A HREF="rncqueryform.asp?region=<%=Session("region")%>&view=<%=Session("view")%>&Discount=<%=Session("strDiscount")%>&rate=<%=Session("strRate")%>&Returned=<%=Session("PerPage")%>"><IMG SRC="/service/center/images/buttons/new-search-button.gif" WIDTH=90 HEIGHT=21 BORDER="0"></A></UL>

</FONT>
<BR><BR>

<!--End Content -->
 
<!--#include virtual="/service/center/footer.asp"-->

<% if Session("view") <> "2" then %>

    </TD>
  </TR>
</TABLE>

<% end if %>

</BODY>
</HTML>

