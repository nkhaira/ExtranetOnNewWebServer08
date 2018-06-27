

<!-- Side Navigation Rows and Container -->

<TABLE WIDTH="100%" CELLPADDING=0 CELLSPACING=0 BORDER=0> 
  <TR>
    <TD CLASS=TopColorBar><IMG SRC="/images/1x1trans.gif" HEIGHT=6 BORDER=0 VSPACE=0></TD>
  </TR>
</TABLE>


<TABLE WIDTH="100%" HEIGHT="100%" CELLPADDING=0 CELLSPACING=0 BORDER="0">
  <TR VALIGN=TOP>  
  	<TD>&nbsp;</td>
    <TD><IMG SRC="/Images/Spacer.gif" WIDTH=4></TD>	
    
    <!-- END LEFT NAVIGATION ROWS -->  

    <!-- BEGIN CONTENT CONTAINER-->

    <TD VALIGN="TOP" WIDTH="100%">
    
    <% if isnumeric(Content_Width) then            
         response.write "<DIV ALIGN=CENTER>" & vbCrLf
         response.write "<TABLE BORDER=0 WIDTH=""" & Content_Width & "%"">" & vbCrLf
         response.write "  <TR>" & vbCrLf
         response.write "    <TD CLASS=NORMAL VALIGN=""TOP"" WIDTH=""100%"">" & vbCrLf
       end if
    %>     

      <BR CLEAR=ALL>
    
    
