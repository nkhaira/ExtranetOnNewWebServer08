<HTML>
<HEAD>
  <TITLE>
  <% if Screen_Title <> "" then
      if instr(1,lcase(Screen_Title),"<img") > 0 then
        if instr(1,lcase(Screen_Title),"alt=") > 0 then
          TempTitle = mid(Screen_Title,instr(1,lcase(Screen_Title),"alt=")+4)
          TempTitle = mid(TempTitle,1,instr(1,temptitle,chr(34))-1)          
          response.write TempTitle
        else
          response.write "Fluke Extranet Support Site"
        end if            
      else  
        response.write Screen_Title
      end if  
     else
      response.write "Fluke Extranet Support Site"
     end if
  %>    
  </TITLE>

  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso8859-1">
  <META NAME="LANG" CONTENT="ENGLISH">
  <META AUTHOR="David Whitlock - David.Whitlock@fluke.com">
</HEAD>

<% if Navigation = True then %>
  <BODY BGCOLOR="white" BACKGROUND="/images/bg.gif" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" LINK="000000" VLINK="000000">
<% else %>
  <BODY BGCOLOR="white" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" LINK="000000" VLINK="000000">
<% end if %>  

<A NAME="VERY_TOP"></A>

<!-- Top Header -->

<TABLE WIDTH="100%" CELLPADDING=0 CELLSPACING=0 BORDER="0">
  <TR>
    <TD HEIGHT=56 VALIGN=TOP ALIGN=LEFT>

    <!-- Top Navigation -->

  	  <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0" BGCOLOR="#000000">
    		<TR>
    			<TD WIDTH="84" ALIGN="LEFT" VALIGN="TOP" HEIGHT="16">
    				<TABLE BORDER="0" CELLPADDING="2" CELLSPACING="0" WIDTH="36">
    						<TR HEIGHT="16">
    							<TD>&nbsp;</TD>
    						</TR>
    						<TR HEIGHT="15">
    							<TD>&nbsp;</TD>
    						</TR>
    						<TR HEIGHT="15">
    							<TD>&nbsp;</TD>
    						</TR>
    						<TR HEIGHT="15">
    							<TD>&nbsp;</TD>
    						</TR>
    				</TABLE>
    			</TD>
          		
          <!-- Spacer -->
    			<TD ALIGN="LEFT" VALIGN="MIDDLE">
            &nbsp;
    			</TD>
          
          <!-- Title -->
    			<TD WIDTH="100%">
    				<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
    					<TR>
    						<TD>
                  <%
                  if Screen_Title <> "" then
                    response.write "<FONT FACE=""Verdana"" SIZE=5 COLOR=""#FFCC00""><B>" & Screen_Title & "</B></FONT>"
                  else
                    response.write "&nbsp;"
                  end if
                  %>
                </TD>
    					</TR>
    				</TABLE>
    			</TD>
          
          <!-- Logo -->
    			<TD WIDTH="146" ALIGN="RIGHT" VALIGN="TOP">
    				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0">
              <TR>
                <TD>
    					    <IMG SRC="/images/spacer.gif" BORDER=0 WIDTH=1 HEIGHT=10><BR>
        					<A HREF="/default.asp" TARGET="_top"><IMG SRC="/images/FlukeLogo3.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
        				</TD>
              </TR>
            </TABLE>
    			</TD>         
    		</TR>
    	</TABLE>
    </TD>
  </TR>
</TABLE>

<!-- Side Navigation Rows and Container -->

<TABLE WIDTH=100% CELLPADDING=4 CELLSPACING=0 BORDER="0">
  <TR>  
    <!-- SIDE NAVIGATION ROWS -->
    <% if Navigation = True then %>
    <TD VALIGN="top">
      <TABLE WIDTH=75 BORDER="0" CELLPADDING=0 CELLSPACING=0>
  	    <TR>
        	<TD COLSPAN=3 VALIGN="middle" HEIGHT=25 NOWRAP ALIGN="left"><BR><BR>
  	        &nbsp;
        	</TD>
    	  </TR>
    	</TABLE>
    </TD>
    <% end if %>
    
    <!-- END SIDE NAVIGATION ROWS -->  

    <!-- BEGIN CONTENT -->

    <TD VALIGN="top">
      <% if Content_Width <> "" then            
           response.write "<CENTER><TABLE WIDTH=""" & Content_Width & "%""><TR><TD>"
         end if
      %>     
          
      <BR CLEAR=ALL>
      <FONT FACE="Verdana, Arial, Helvetica" SIZE=2>

      <!-- START OF CONTENT -->

      