<!-- Whitlock's SiteWide Content Server SWCS 2.0 12/8/00 10:39:29 PM PST -->

<%
'   Query String Variables

'   AID               = User Account ID Number for polling SiteWide DB
'   Site_Description  = Text String
'   Site_ID           = SiteWide Site ID Number
'   BackURL           = Redirection URL back to SiteWide
'   Language          = 3-Character ISO Language Code

%>

<HTML>
<HEAD>
<TITLE>
<%=Request("Site_Description")%> - Extranet Support Site
</TITLE>
<LINK REL=STYLESHEET HREF="/SW-Common/SW-Style.css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso8859-1">
<META HTTP-EQUIV="Content-language" CONTENT="eng">
<META NAME="LANGUAGE" CONTENT="English">
<META NAME="AUTHOR" CONTENT="K. David Whitlock - David.Whitlock@fluke.com">
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
    						<TD BGCOLOR="#000000" CLASS=HEADING3FLUKE>
                  <%=Request("Site_Description")%>
                  <BR>
                  <FONT CLASS=SMALLBOLDGOLD>Image Store</FONT>
                </TD>
    					</TR>
    				</TABLE>
    			</TD>
          
          <!-- Logo -->
    			<TD WIDTH="146" ALIGN="RIGHT" VALIGN="TOP">
    				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0">
              <TR>
                <TD VALIGN=MIDDLE>
                  <BR>
                
                    <A HREF="<%=Request("BackURL")%>" TARGET="VERY_TOP"><IMG SRC="/images/FlukeLogo3.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
                    
                  <BR>                  
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


<!-- Top Navigation -->

<TABLE WIDTH="100%" CELLPADDING=0 CELLSPACING=0 BORDER=0 BGCOLOR="#CCCCCC" CLASS=TABLEBACKGROUND>
  
  
  <TR>
    <TD  CLASS=LINEBACKGROUND><IMG SRC="/images/1x1trans.gif" HEIGHT=6 BORDER=0 VSPACE=0></TD>
  </TR>
</TABLE>

<!-- Side Navigation Rows and Container -->

<TABLE WIDTH="100%" HEIGHT="100%" CELLPADDING=0 CELLSPACING=0 BORDER="0">
  <TR>  

    <!-- Side NAVIGATION ROWS -->

    <TD WIDTH=128 VALIGN=TOP BGCOLOR=#000000 CLASS=NAVLEFTTABLEBACKGROUND>
      <TABLE WIDTH=128 BORDER=0 CELLPADDING=2 CELLSPACING=0>                     

        <!-- Level 1 Menu Item -->
        <TR>
          <TD></TD>
          <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
        </TR>
            
        <TR>           
          <TD WIDTH=8></TD>
          <TD CLASS=NAVLEFT1>
            <A HREF="<%=Request("BACKURL")%>" CLASS=NavLeft1 TITLE="Home">Home</A>&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="/images/home.gif" WIDTH=21 HEIGHT=10 BORDER=0 VSPACE=0 ALT="Home" ALIGN=ABSMIDDLE>
          </TD>
        </TR>
                        
        <!-- Level 1 Menu Item -->
        <TR>
          <TD></TD>
          <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
        </TR>

        <TR>
         <TD WIDTH=8></TD>
         <TD CLASS=NAVLEFTHIGHLIGHT1>
           <A HREF="" CLASS=NAVLEFTHIGHLIGHT1 TITLE="Search">Search</A>
         </TD>
       </TR>
                                          
        <!-- Level 1 Menu Item -->
        <TR>
          <TD></TD>
          <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
        </TR>
        
        <TR>
          <TD WIDTH=8></TD>
          <TD CLASS=NAVLEFT1>
            <A HREF="" CLASS=NAVLEFT1 TITLE="Tree">Tree</A>
          </TD>
        </TR>
                          
        <!-- Level 1 Menu Item -->
        <TR>
          <TD></TD>
          <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
        </TR>
        
        <TR>          
          <TD WIDTH=8></TD>
          <TD CLASS=NAVLEFT1>
            <A HREF="" CLASS=NAVLEFT1 TITLE="Order Preview">Order Preview</A>
          </TD>
        </TR>

        <!-- Level 1 Menu Item -->
        <TR>
          <TD></TD>
          <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
        </TR>
      
        <TR>         
          <TD WIDTH=8></TD>
          <TD CLASS=NAVLEFT1>
            <A HREF="" CLASS=NAVLEFT1 TITLE="Help">Help</A>
          </TD>
        </TR>
                                
        <!-- Level 1 Menu Item -->
        <TR>
          <TD></TD>
          <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
        </TR>
        
        <TR>         
          <TD WIDTH=8></TD>
          <TD CLASS=NAVLEFT1>
            <A HREF="" CLASS=NAVLEFT1 TITLE="Back">&lt;&lt;&nbsp;Back</A>
          </TD>
        </TR>

        <TR>
          <TD></TD>
          <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
        </TR>
                          
      </TABLE>
    </TD>
        
    <!-- END LEFT NAVIGATION ROWS -->  

    <!-- BEGIN CONTENT CONTAINER-->

    <TD VALIGN="top" CLASS=NORMAL WIDTH="100%">

    <DIV ALIGN=CENTER>
      <TABLE WIDTH="95%">
        <TR>
          <TD CLASS=NORMAL VALIGN="TOP" WIDTH="100%">
     
            <BR CLEAR=ALL>

            <!-- BEGIN CONTENT -->

            <FONT CLASS=HEADING3>Image Store</FONT>
            <BR><BR>
                 
            <!-- Begin Content Container -->

          	<TABLE BORDER="0" WIDTH="400" CELLSPACING="10" CELLPADDING="2">
          	<FORM ACTION="ssearch.asp?begin=1" METHOD="POST" NAME="searchform">
          	<!--<form ACTION="sresult.asp" METHOD="POST" NAME="searchform">-->
          	<TR>
            <TD BGCOLOR=#FFCC00 Class=Medium>
          	<FONT CLASS=MediumBOLD>Search by Image Type:</FONT><BR>
            
            <SELECT CLASS=NORMAL NAME="objecttype" ONCHANGE="searchform.submit()">
          	<!--	<option value="">Choose Image Type</option>-->
          	<OPTION VALUE></OPTION>
        		<OPTION CLASS=MEDIUM VALUE="1">Digital Image</OPTION>
        		<OPTION CLASS=MEDIUM VALUE="2">Logo</OPTION>	
        		<OPTION CLASS=MEDIUM VALUE="3">Clipart</OPTION>	
          	</SELECT>
          	</TD></TR><TR>
          	</FORM>
            
          	<FORM ACTION="sresult.asp?Stype=id" METHOD="POST" NAME="searchform2">
          	<TD BGCOLOR="#FFCC00" CLASS=Medium><FONT CLASS=MediumBOLD>Search by Image ID:
            <BR>
          	<INPUT CLASS=MEDIUM TYPE="text" SIZE="6" NAME="object_id"><BR>Use this when you know the image ID</FONT>
          	</TD>
          	</TR>
          
            <!--
          	<tr>
          	<td bgcolor="#fbe997"><font face="Verdana" siZe="2"><b>Search by distributor:</b><br><a href="http://islay/20000136/" target="_self">Search</a><br>&nbsp;<br></font>
          	</td>
          	</tr>
            -->
          
          	<TR><TD><INPUT TYPE="hidden" NAME="begin" VALUE="2"></TD></TR>
          	</FORM>
          	</TABLE>
      <!-- END CONTENT -->      
      
          </TD>
        </TR>
      </TABLE>
    </DIV>                

      <!-- END CONTENT CONTAINER-->

    </TD>
  </TR>

  <!-- Begin Footer -->

  <TR>   
    <TD BGCOLOR="#000000">&nbsp;</TD> 
    
    <TD ALIGN="CENTER" VALIGN="TOP" CLASS=SMALL>
      <%response.write "&copy; 1995-" & DatePart("yyyy",Date) & " Fluke Corporation - " & "All rights reserved" & "."%>
    </TD>
  </TR>

  <!-- End Footer -->

</TABLE>

</BODY>
</HTML>

<!-- Whitlock's SiteWide Content Server SWCS 2.0 12/8/00 10:39:30 PM PST -->
