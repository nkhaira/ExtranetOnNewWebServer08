<% if 1=2 then %>

<DIV ALIGN=CENTER>
<!--TABLE border=0 BORDERCOLOR="#8D8961">
<TR>
<TD BGCOLOR="#C2BFA5"-->
<FORM NAME="Login" METHOD="POST">
<DIV STYLE="position:relative;left:0;top:0;width:402px;height:254px;background-color:#C2BFA5;border:solid 1px #000000">
<DIV STYLE="position:absolute;left:0;top:0;width:402px;height:16px;background-color:Maroon">
<SPAN STYLE="position:absolute;left:2;font-family:Arial;font-size:10pt;color:#E1E0D2;font-weight:bold">Enter Network Password</SPAN></DIV>
<IMG SRC="/images/IE-Logon-Key.jpg" BORDER=0 STYLE="position:absolute;top:38px;left:22px">
<SPAN STYLE="position:absolute;left: 68px;top:42px;font-family:Arial;font-size:8pt;color:Black">Please type your user name and password:</SPAN>

<SPAN STYLE="position:absolute;left: 68px;top: 70px;font-family:Arial;font-size:8pt;color:Black">Site:</SPAN>
<SPAN STYLE="position:absolute;left:142px;top: 70px;font-family:Arial;font-size:8pt;color:Black">Support.Fluke.com</SPAN>

<SPAN STYLE="position:absolute;left: 68px;top:102px;font-family:Arial;font-size:8pt;color:Black"><U>U</U>ser Name</SPAN>
<INPUT TYPE="TEXT" NAME="User_Name"    STYLE="width:200px;position:absolute;left:138px;top: 96px;font-family:Arial;font-size:9pt;color:Black">

<SPAN STYLE="position:absolute;left: 68px;top:132px;font-family:Arial;font-size:8pt;color:Black"><U>P</U>assword</SPAN>
<INPUT TYPE="PASSWORD" NAME="Password" STYLE="width:200px;position:absolute;left:138px;top:126px;font-family:Arial;font-size:9pt;color:Black">

<SPAN STYLE="position:absolute;left: 68px;top:162px;font-family:Arial;font-size:8pt;color:Black"><U>D</U>omain:</SPAN>
<INPUT DISABLED TYPE="TEXT" NAME="Domain"       STYLE="width:200px;position:absolute;left:138px;top:156px;font-family:Arial;font-size:9pt;color:Black;background-color:#C2BFA5">

<INPUT DISABLED TYPE="CHECKBOX" NAME="Save_Password" STYLE="position:absolute;left:64px;top:186px;font-family:Arial;font-size:9pt;color:Black;background-color:#C2BFA5">
<SPAN STYLE="position:absolute;left:86px;top:191px;font-family:Arial;font-size:8pt;color:Black"><U>S</U>ave this password in your password list</SPAN>

<INPUT TYPE="button" VALUE="OK"     STYLE="width:76;height:24;position:absolute;left:190px;top:216px;background-color:#C2BFA5;color:#000000;font-family:Arial;font-size:9pt">
<INPUT TYPE="button" VALUE="Cancel" STYLE="width:76;height:24;position:absolute;left:312px;top:216px;background-color:#C2BFA5;color:#000000;font-family:Arial;font-size:9pt">

</DIV>
</FORM>
<!--/TD>
</TR>
</TABLE-->
</DIV>

<% end if %>

<BR><BR><BR><BR>

<DIV ALIGN=CENTER>
<FORM Name="Login" Action "Login_Admin.asp" Method="POST">
<TABLE WIDTH=402 HEIGHT=254 BORDER=2 CELLSPACING=0 CELLPADDING=0 BORDER=0>
  <TR>
    <TD WIDTH="100%" BGCOLOR="#C2BFAF">
      <TABLE WIDTH=402 HEIGHT=254 BGCOLOR="#C2BFA5" CELLSPACING=0 CELLPADDING=0>
        <TR>
          <TD COLSPAN=3 BGCOLOR=Maroon HEIGHT=16>
          <IMG SRC="/images/Button-Close.gif" ALIGN=RIGHT BORDER=0 ALT="Cancels Network Login Request."> <IMG SRC="/images/Button-Help.gif" ALIGN=RIGHT BORDER=0 ALT="Provides a space for you to type the User ID and Password for you to gain access to the selectd web site.">
          <SPAN STYLE="font-family:Arial;font-size:10pt;color:#E1E0D2;font-weight:bold">Enter Network Password</SPAN></TD>
        </TR>

        <TR>
          <TD ROWSPAN=8 ALIGN=CENTER VALIGN=TOP WIDTH=52><BR><IMG SRC="/images/IE-Logon-Key.jpg" BORDER=0></TD>
          <TD WIDTH=350 COLSPAN=2 HEIGHT=8><IMG SRC="/images/1x1trans.gif" HEIGHT=8 BORDER=0 VSPACE=0></TD>
        </TR>

        <TR>
          <TD COLSPAN=2 VALIGN=Bottom><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black">Please type your user name and password:</SPAN></TD>
        </TR>

        <TR>
          <TD WIDTH= 75 VALIGN=Bottom><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black">Site:</SPAN></TD>
          <TD WIDTH=275 VALIGN=Bottom><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black">Support.Fluke.com</SPAN></TD>          
        </TR>

        <TR>
          <TD VALIGN=Bottom><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><U>U</U>ser Name</SPAN><BR><IMG SRC="/images/1x1trans.gif" HEIGHT=4 BORDER=0 VSPACE=0></TD>
          <TD VALIGN=Bottom><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><INPUT TYPE="TEXT" NAME="User_Name"></SPAN></TD>
        </TR>

        <TR>
          <TD VALIGN=Bottom><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><U>P</U>assword</SPAN><BR><IMG SRC="/images/1x1trans.gif" HEIGHT=4 BORDER=0 VSPACE=0></TD>
          <TD VALIGN=Bottom><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><INPUT TYPE="PASSWORD" NAME="Password"></SPAN></TD>
        </TR>

        <TR>
          <TD VALIGN=Bottom><SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><U>D</U>omain:</SPAN><BR><IMG SRC="/images/1x1trans.gif" HEIGHT=4 BORDER=0 VSPACE=0></TD>
          <TD VALIGN=Bottom><SPAN STYLE="background-color:#C2BFA5;font-family:Arial;font-size:8.5pt;color:Black"><INPUT DISABLED TYPE="TEXT" NAME="Domain"></SPAN></TD>
        </TR>

        <TR>
          <TD COLSPAN=2>
            <INPUT DISABLED TYPE="CHECKBOX" NAME="Save_Password">
            <SPAN STYLE="font-family:Arial;font-size:8.5pt;color:Black"><U>S</U>ave this password in your password list</SPAN>
          </TD>  
          
        </TR>

        <TR>
          <TD COLSPAN=2>
            <TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0>
              <TR>
                <TD ALIGN=RIGHT HEIGHT=24 WIDTH="75%">
                  <INPUT TYPE="button" NAME="Action" VALUE="    OK    " STYLE="width:76;height:24;background-color:#C2BFA5;color:#000000;font-family:Arial;font-size:9pt" TITLE="Submits your User ID and Password for verification and selected site access.">&nbsp;&nbsp;&nbsp;&nbsp;
                </TD>
                <TD ALIGN=RIGHT HEIGHT=24 WIDTH="25%">
                  <INPUT TYPE="button" NAME="Action" VALUE="Cancel"     STYLE="width:76;height:24;background-color:#C2BFA5;color:#000000;font-family:Arial;font-size:9pt">&nbsp;&nbsp;<FONT COLOR="#C2BFA5">.</FONT>
                </TD>
              </TR>
            </TABLE>  
          </TD>  
        </TR>

      </TABLE>
    </TD>
  </TR>
</TABLE>
</FORM>
</DIV>
