<HTML>
<HEAD>
  <TITLE>
  <% if not isblank(Screen_Title) then
      response.write Screen_Title
     end if
  %>    
  </TITLE>
  <LINK REL=STYLESHEET HREF="/sw-common/SW-Style.css">
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso8859-1">
  <META NAME="LANG" CONTENT="ENGLISH">
  <META AUTHOR="David Whitlock - David.Whitlock@fluke.com">
</HEAD>

<BODY BGCOLOR="White" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" LINK="#000000" VLINK="#000000" ALINK="#000000">

<A NAME="VERY_TOP"></A>

<!-- Top Header -->

<TABLE WIDTH="100%" CELLPADDING=0 CELLSPACING=0 BORDER="0">
  <TR>
    <TD HEIGHT=56 VALIGN=TOP ALIGN=LEFT>

    <!-- Top Navigation -->

  	  <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0" BGCOLOR="#000000">
    		<TR>
    			<TD WIDTH="128" ALIGN="LEFT" VALIGN="TOP" HEIGHT="16">
     				<TABLE BORDER="0" CELLPADDING="2" CELLSPACING="0" WIDTH="8">
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
    						<TD BGCOLOR="#000000" CLASS=Heading3Fluke>
                  <%
                  if not isblank(Bar_Title) then
                    response.write Bar_Title
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
                  <BR>
        					<A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FlukeLogo3.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
        				</TD>
              </TR>
            </TABLE>
    			</TD>         
    		</TR>
    	</TABLE>
    </TD>
  </TR>
</TABLE>

<!-- END HEADER -->

