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
<LINK REL=STYLESHEET HREF="http://Support.Fluke.com/SW-Common/SW-Style.css">
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
                  <FONT CLASS=SMALLBOLDGOLD>Publication Store</FONT>
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
                
                    <A HREF="<%=Request("BACKURL")%>" TARGET="VERY_TOP"><IMG SRC="/images/FlukeLogo3.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
                    
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
          <TD CLASS=NAVLEFT1>
            <A HREF="" CLASS=NAVLEFT1 TITLE="Search">Search</A>
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
           <A HREF="" CLASS=NAVLEFTHIGHLIGHT1 TITLE="Tree">Tree</A>
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

            <FONT CLASS=HEADING3>Publication Store</FONT>
            <BR><BR>
                
            <!-- Begin Content Container -->

            <FORM ACTION="filter.asp" METHOD="POST" NAME="searchform">
            <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
              <TR>
                <TD ALIGN="right" VALIGN="middle">&nbsp &nbsp<FONT CLASS=MEDIUM>Pubtype:</FONT>
          
                  <SELECT CLASS=MEDIUM SIZE="1" NAME="imagetype" TABINDEX="3">
                		<OPTION CLASS=MEDIUM SELECTED VALUE="0">All</OPTION>
        	    			<OPTION CLASS=MEDIUM VALUE="1" >Document</OPTION>
            				<OPTION CLASS=MEDIUM VALUE="2" >Video</OPTION>		
            				<OPTION CLASS=MEDIUM VALUE="3" >CD/DVD/3H</OPTION>
                  </SELECT>
            	  </TD>
              </TR>
            </TABLE>
          </FORM>

          <HR  COLOR=<%=Contrast%>>            
            
        	<TABLE BORDER="0" ALIGN="center" WIDTH="800">
          	<TR>	
        			<TD VALIGN="top">   				
              	<UL><LI><B><FONT FACE="verdana" SIZE="2">Industrial/Electrical </FONT></B></LI><UL>
	
            		<UL>
		
          			<LI><FONT FACE="verdana" SIZE="2">
        				<A HREF="sresult.asp?Stype=tree&child_id=100001">
        				DMM's&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
              	<FONT SIZE="1" FACE="arial">(34)</FONT>
          			</B>
          			</LI>
			
          			<LI><FONT FACE="verdana" SIZE="2">
        				<A HREF="sresult.asp?Stype=tree&child_id=100002">
        				Industrial General&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
              	<FONT SIZE="1" FACE="arial">(24)</FONT>
          			</B>
	          		</LI>
			
          			<LI><FONT FACE="verdana" SIZE="2">
        				<A HREF="sresult.asp?Stype=tree&child_id=100003">
        				Automotive&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
              	<FONT SIZE="1" FACE="arial">(7)</FONT>
          			</B>
        			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100004">
				Scopemeters&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(12)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100005">
				Clamp Meters&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(3)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100006">
				Electrical Testers&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(6)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100007">
				Thermometers&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(6)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100008">
				Gas Measurement&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(1)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100009">
				Process Tools&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(26)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100010">
				Accessories&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(6)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100011">
				Power Quality Tools&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(16)</FONT>

			</B>
			</LI>
			
		</UL>
	
			</TD><TD VALIGN="top">
				
	<UL><LI><B><FONT FACE="verdana" SIZE="2">Networks </FONT></B></LI><UL>
	
		<UL>
		
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100012">
				Networks General&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(6)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100013">
				Cable Testers&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(21)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100014">
				Communication Test Tools&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(16)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100015">
				Network Troubleshooting Tools&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(59)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100024">
				Support Services&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(5)</FONT>

			</B>
			</LI>
			
		</UL>
	
			</TD></TR>
			<TR><TD VALIGN="top">
				
	<UL><LI><B><FONT FACE="verdana" SIZE="2">ITI </FONT></B></LI><UL>
	
		<UL>
		
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100025">
				General&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(3)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100018">
				Oscilloscopes&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(11)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100019">
				Power Supplies&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(1)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100020">
				RCL Meters&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(4)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100021">
				Signal Sources&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(2)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100022">
				Timers/Counters&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(22)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100023">
				TV Generators&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(9)</FONT>

			</B>
			</LI>
			
		</UL>
	
			</TD><TD VALIGN="top">
				
	<UL><LI><B><FONT FACE="verdana" SIZE="2">Other Brands</FONT></B></LI><UL>
	
		<UL>
		
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100028">
				Agilent&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(1)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100029">
				Pomona&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(0)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100026">
				Robin&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(2)</FONT>

			</B>
			</LI>
			
			<LI><FONT FACE="verdana" SIZE="2">
				<A HREF="sresult.asp?Stype=tree&child_id=100027">
				Wavetek/Meterman&nbsp;<FONT SIZE="1" FACE="verdana"></FONT></A>&nbsp;<FONT FACE="verdana" SIZE="2">
	<FONT SIZE="1" FACE="arial">(5)</FONT>

			</B>
			</LI>
			
		</UL>
	
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
