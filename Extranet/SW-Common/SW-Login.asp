<HTML>
<HEAD>
<TITLE>Enter Network Password ................................................................................................</TITLE>
<LINK REL=STYLESHEET HREF="/SW-Common/SW-Style.css">
</HEAD>

<BODY BGCOLOR="white" TEXT="#000000" STYLE="font-size:10pt;font-family:Arial">

<!-- Left Vertical Container -->

<TABLE STYLE="width:402px;height:254px;position:center">
<TR>
<TD>
<DIV STYLE="position:absolute;left:16;top:16;background-color:#C2BFA5;text-align:center;width:50;height:176">
<IMG SRC="/images/IE-Logon-Key.jpg" BORDER="0" STYLE="position:relative;top:0">
</DIV>


<DIV STYLE="position:absolute;left:63;top:16;background-color:#C2BFA5;width:316;height:176">

<FORM NAME="Login" METHOD="POST">

<SPAN STYLE="position:absolute;left:3;top:4px;font-family:Arial;font-size:8pt;color:Black">Please type your user name and password:</SPAN>
<SPAN STYLE="position:absolute;left:3;top:34px;font-family:Arial;font-size:8pt;color:Black">Site:</SPAN>
<SPAN STYLE="position:absolute;left:76px;top:34px;font-family:Arial;font-size:8pt;color:Black">Support.Fluke.com</SPAN>

<SPAN STYLE="position:absolute;left:3;top:62px;font-family:Arial;font-size:8pt;color:Black"><U>U</U>ser Name:</SPAN>
<INPUT TYPE="TEXT" NAME="User_Name" STYLE="width:200px;position:absolute;left:76px;top:58px;font-family:Arial;font-size:9pt;color:Black">
<SPAN STYLE="position:absolute;left:3;top:92px;font-family:Arial;font-size:8pt;color:Black"><U>P</U>assword:</SPAN>
<INPUT TYPE="PASSWORD" NAME="Password" STYLE="width:200px;position:absolute;left:76px;top:88px;font-family:Arial;font-size:9pt;color:Black">
<SPAN STYLE="position:absolute;left:3;top:122px;font-family:Arial;font-size:8pt;color:Black"><U>D</U>omain:</SPAN>
<INPUT DISABLED TYPE="TEXT" NAME="Domain" STYLE="width:200px;position:absolute;left:76px;top:118px;font-family:Arial;font-size:9pt;color:Black;background-color:#C2BFA5">
<INPUT TYPE="CHECKBOX" NAME="Save_Password" STYLE="position:absolute;left:0px;top:144px;font-family:Arial;font-size:9pt;color:Black">
<SPAN STYLE="position:absolute;left:24;top:148px;font-family:Arial;font-size:8pt;color:Black" onclick="chgstatus()"><U>S</U>ave this password in your password list</SPAN>

<INPUT TYPE="button" VALUE="OK"     STYLE="width:76;height:24;position:absolute;left:150;top:176;background-color:#C2BFA5;color:#000000;font-family:"Arial";font-size:7.5pt;border:2px outset #FFFFFF" onclick="ntip()">
<INPUT TYPE="button" VALUE="Cancel" STYLE="width:76;height:24;position:absolute;left:248;top:176;background-color:#C2BFA5;color:#000000;font-family:"Arial";font-size:7.5pt;border:2px outset #FFFFFF" onclick="endf()">

</FORM>

</DIV>
</TD>
</TR>
</TABLE>

</BODY>
