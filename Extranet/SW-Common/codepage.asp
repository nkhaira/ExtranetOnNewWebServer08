<!-- #include virtual="/connections/connections_parts.asp" -->

<% 

  call connect_parts
  
  SQL = "SELECT * FROM vcturbo_c3 WHERE c3_orderable=0 ORDER BY c3_id"
  Set Session("rsCodeList") = Server.CreateObject("ADODB.Recordset")
  Session("rsCodeList").open SQL,DBConn,3,1,1

%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<HTML>
<HEAD>
<TITLE>Fluke Replacement Parts - Restriction Code List</TITLE>
<LINK REL=STYLESHEET HREF="SW-Common-Style.css">
</HEAD>
<BODY BGCOLOR="White">

	<TABLE BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
    <TR>
      <TD>
        <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0>
          <TR>
          	<TD WIDTH="5%" ALIGN="left" BGCOLOR="#000000" CLASS=SMALLBOLDGOLD>Code #</TD>
          	<TD WIDTH="95%" BGCOLOR="#000000" CLASS=SMALLBOLDGOLD>Definition</TD>
          </TR>
          
          <% DO WHILE NOT Session("rsCodeList").EOF 
				if Not isNull (Session("rsCodeList")("c3_value")) then %>		
					<TR>
					  <TD ALIGN="CENTER" BGCOLOR="#EEEEEE" CLASS=SMALL><%= Session("rsCodeList")("c3_id") %></TD>
					  <TD BGCOLOR="#FFFFFF" CLASS=SMALL><%= Session("rsCodeList")("c3_value") %></TD>            
					</TR>    
			<%	end if 
				Session("rsCodeList").MoveNext
        	loop
          %>

          <TR>
            <TD ALIGN="center" COLSPAN="2" BGCOLOR="Black" CLASS=NORMAL>
              <INPUT LANGUAGE="VBScript" TYPE=SUBMIT VALUE="Close Window" ONCLICK="window.close" NAME="close" CLASS=NavLeftHighlight1>
            </TD>
          </TR>
        </TABLE>
      </TD>
    </TR>
  </TABLE>      

<%

Session("rsCodeList").close
call disconnect_parts

Set Session("rsCodeList") = Nothing

%>

</BODY>
</HTML>