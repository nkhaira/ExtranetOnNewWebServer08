

<!-- Side Navigation Rows and Container -->

<TABLE WIDTH="100%" CELLPADDING=0 CELLSPACING=0 BORDER=0> 
  <TR>
    <TD CLASS=TopColorBar><IMG SRC="/images/1x1trans.gif" HEIGHT=6 BORDER=0 VSPACE=0></TD>
  </TR>
</TABLE>

<TABLE WIDTH="100%" HEIGHT="100%" CELLPADDING=0 CELLSPACING=0 BORDER="0">
  <TR VALIGN=TOP>  

    <!-- Side NAVIGATION ROWS -->

    <TD WIDTH=4><IMG SRC="/Images/Spacer.gif" WIDTH=4></TD>
    <TD WIDTH=128 VALIGN=TOP BGCOLOR=White CLASS=Small>
    <IMG SRC="/Images/Spacer.gif" WIDTH=128 HEIGHT=4><BR>
    
      <% Call Nav_Border_Begin %>
      
      <TABLE WIDTH=128 BORDER=0 CELLPADDING=2 CELLSPACING=0>

        <TR>
          <!--TD WIDTH=8></TD-->
          
            <% if isblank(BackURL) then
                 response.write "<TD CLASS=NavLeft1 BGCOLOR=""#000000"">"
            %>     
                 <INPUT TYPE=BUTTON VALUE="<%Response.write Translate("Close Window",Login_Language,conn)%>" LANGUAGE="JavaScript" ONCLICK="window.close();" NAME="close" CLASS=NavLeftHighlight1>
            <% else
                 response.write "<TD CLASS=NavLeftHighlight1 BGCOLOR=""" & Contrast & """>"
            %>     
               <IMG SRC="/images/home.gif" WIDTH=21 HEIGHT=10 BORDER=0 VSPACE=0 ALT="Home" ALIGN=RIGHT><A HREF="<%=BackURL%>" CLASS=NavLeftHighlight1 TITLE="Home"><%response.write Translate("Home",Login_Language,conn)%></A>
            <% end if %>
          </TD>
        </TR>
     
      </TABLE>
      
      <% Call Nav_Border_End %>

    </TD>
    
    <!-- END LEFT NAVIGATION ROWS -->  

    <!-- BEGIN CONTENT CONTAINER-->

    <TD VALIGN="top" CLASS=Normal WIDTH="100%">

    <% if isnumeric(Content_Width) then            
         response.write "<DIV ALIGN=CENTER>" & vbCrLf
         response.write "<TABLE BORDER=0 WIDTH=""" & Content_Width & "%"">" & vbCrLf
         response.write "  <TR>" & vbCrLf
         response.write "    <TD CLASS=NORMAL VALIGN=""TOP"" WIDTH=""100%"">" & vbCrLf
       end if
    %>     

      <BR CLEAR=ALL>

<%      
'--------------------------------------------------------------------------------------
' Subroutines and Functions
'--------------------------------------------------------------------------------------

sub Nav_Border_Begin()
    response.write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" CLASS=NavBorder>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLr
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "        <TD VALIGN=""top"" CLASS=NavBorder>" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Nav_Border_End()
    response.write "        </TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_BL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_BR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "    </TABLE>" & vbCrLf
end sub  
%>      
    
    
