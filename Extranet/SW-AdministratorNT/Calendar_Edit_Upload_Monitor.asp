<%@ENABLESESSIONSTATE =False%>

<%
'--- Instanciate SA-Progress

Set uplp = Server.CreateObject("Softartisans.FileUpProgress")
uplp.progressid = CInt(request.querystring("ProgressID"))

Dim Pause
Pause = False

response.write "<HTML>" & vbCrLF
response.write "<HEAD>" & vbCrLF
response.write "<TITLE>Upload Stream Monitor</TITLE>" & vbCrLf
response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf

' Slow down update for inactivity
response.write "<META HTTP-EQUIV=""Refresh"" CONTENT="""
if uplp.items.count <= 0  then
  response.write "3"
else
  response.write "1"
end if
response.write """>" & vbCrLF

response.write "</HEAD>" & vbCrLF
response.write "<BODY>" & vbCrLF

if uplp.items.count <= 0 then
  response.write "<IMG SRC=""/images/sniffing_dog.gif"" Border=0 ALIGN=RIGHT>" & vbCrLf
  response.write "<SPAN CLASS=NORMALBOLD>Upload Stream Monitor - Sniffing...</B>" & vbCrLf
  response.write "<SCRIPT Language=""Javascript"">" & vbCrLf
  response.write "window.blur();" & vbCrLf
  response.write "</SCRIPT>" & vbCrLf
else
  response.write "<SPAN CLASS=NORMALBOLD>Upload Stream Monitor - Uploading...</SPAN>" & vbCrLf
  response.write "<SCRIPT Language=""Javascript"">" & vbCrLf
  response.write "window.focus();" & vbCrLf
  response.write "</SCRIPT>" & vbCrLf
end if  
%>

<TABLE WIDTH=""100%"" BGCOLOR="Black" Class=Small BORDER=0>
  <TR>
    <TD WIDTH=""100%"">
      <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=1 WIDTH="100%">
        <TR>
          <TD BGCOLOR="White" CLASS=SmallBold>ID</TD>
          <TD BGCOLOR="White" CLASS=SmallBold ALIGN=LEFT WIDTH=200>Percentage Completed</TD>
          <TD BGCOLOR="White" CLASS=SmallBold ALIGN=CENTER WIDTH=30>%</TD>
          <TD BGCOLOR="White" CLASS=SmallBold ALIGN=CENTER>Transferred Bytes</TD>
          <TD BGCOLOR="White" CLASS=SmallBold ALIGN=CENTER>Total Bytes</TD>
        </TR>

        <%
        for each item in uplp.items
  	      response.write "<TR>"
          response.write "<TD BGCOLOR=""White"" CLASS=Small>" & item.progressid & "</TD>"
          response.write "<TD BGCOLOR=""Gainsboro"" align=left width=200 CLASS=Small NOWRAP>"
        	response.write "  <HR style='color : green' size='10' width='"   & Int(item.percentage * 2) & "'>"
          response.write "</TD>"
          response.write "<TD BGCOLOR=""White"" align=RIGHT CLASS=Small WIDTH=30>" & item.percentage & "%</TD>"
        	response.write "<TD BGCOLOR=""White"" align=RIGHT CLASS=Small>" & item.transferredbytes & "</TD>"
        	response.write "<TD BGCOLOR=""White"" align=RIGHT CLASS=Small>" & item.totalbytes & "</TD>"
    	    response.write "</TR>"
        next
        response.write "</TABLE>"
        %>
    </TD>
  </TR>
</TABLE>
<BR>
<SPAN CLASS=Small>
<B>Total Bytes</B> represents the total size of <U>all files</U> uploaded to the server. 
<B>Transferred Bytes</B> and <B>Percentage Completed</B> represent the amount of successful packet transfers.
This monitor will automatically disapear when the transfer has successfully completed, 
or the error will be reported.
<P>
<SPAN CLASS=SMALL>Last Update:<%=now()%>
</BODY>
</HTML>