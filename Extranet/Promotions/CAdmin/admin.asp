<%
response.buffer = true
Server.ScriptTimeOut = 1800

if Request.ServerVariables("LOGON_USER") = "" then
  response.redirect "http://www.fluke.com/default.asp"
end if

%>

<!--#include virtual="/connections/connection_cisco.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<HTML>
<HEAD>
	<TITLE>Cisco Administration Screen</TITLE>
</HEAD>

<BODY BGCOLOR="#FFFFFF">
<FONT FACE="Verdana" SIZE=2>

<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
  <TR>
		<TD BGCOLOR="Black"><FONT FACE="Verdana, helvetica" SIZE="5" COLOR="#FFCC00">&nbsp;<B>Cisco - Administration Screen</B></FONT></TD>   	    
		<TD WIDTH="3" BGCOLOR="White"></TD>
	    <TD ALIGN="RIGHT" VALIGN="MIDDLE" BGCOLOR="#FFCC00" WIDTH="148"><IMG SRC="/images/flukelogo.gif" ALIGN="RIGHT" WIDTH="136" HEIGHT="44">
		</TD>
	</TR>
</TABLE>
<BR><BR>

<FORM ACTION="orderForm.asp" METHOD="post">

<%

if request("view") = "closed" then 
  response.write "<INPUT TYPE=""hidden"" NAME=""view"" VALUE=""closed"">"
  response.write "<A HREF=""admin.asp?view=open"">View Open CISCO Orders</A><BR><BR>"
else
  response.write "<INPUT TYPE=""hidden"" NAME=""view"" VALUE=""open"">"
  response.write "<A HREF=""admin.asp?view=closed"">View Closed CISCO Orders</A><BR><BR>"
end if

Call Connect_Cisco

if request("delete") = "true" then
  sql = "DELETE FROM cisco WHERE (((uon)='1' Or (uon)='2'))"
  set rs = conn.execute(sql)
end if  

if len(request("oNum")) > 0  and isnumeric(request("oNum")) then
  sql = "UPDATE cisco SET cisco.status='closed', cisco.cc_num='' WHERE (((cisco.oNum)=" & request("oNum") & "))"
  set rs = conn.execute(sql)
end if

if request("view") = "closed" then
  sql = "SELECT DISTINCT ddate, oNum FROM cisco WHERE (((status) ='closed')) ORDER BY ddate desc"
else
  sql = "SELECT DISTINCT ddate, oNum FROM cisco WHERE (((status)<>'closed' OR (status) Is Null)) ORDER BY ddate desc"
end if

'response.write sql & "<BR><BR>"

set rs = conn.execute(sql)

if rs.eof or rs.bof then
%>

  <B>There are no CISCO
  
  <% if request("view") = "closed" then %>
    &nbsp;Closed&nbsp;
  <% else %>
    &nbsp;Open&nbsp;
  <% end if %>  
  
  Orders at this time.</B>

<% else %>

<TABLE BGCOLOR="#000000" CELLPADDING="0" CELLSPACING="0" BORDER="0" WIDTH="100%">
	<TR>
  	<TD WIDTH="100%">
	    <TABLE BORDERCOLOR="#000000" BORDER="0" CELLPADDING="2" WIDTH="100%">
        <TR>
          <TD BGCOLOR="#000000" WIDTH="6%"><FONT FACE="verdana" COLOR="#FFCC00" SIZE="1"><B>Status</B></FONT></TD>
          <TD BGCOLOR="#000000" WIDTH="8%"><FONT FACE="verdana" COLOR="#FFCC00" SIZE="1"><B>Date and Time</B></FONT></TD>
          <TD BGCOLOR="#000000" WIDTH="4%"><FONT FACE="verdana" COLOR="#FFCC00" SIZE="1"><B>Cisco</B></FONT></TD>
          <TD BGCOLOR="#000000" WIDTH="22%"><FONT FACE="verdana" COLOR="#FFCC00" SIZE="1"><B>Item #/Model #/(Qty)</B></FONT></TD>
          <TD BGCOLOR="#000000" WIDTH="14%"><FONT FACE="verdana" COLOR="#FFCC00" SIZE="1"><B>Name</B></FONT></TD>
          <TD BGCOLOR="#000000" WIDTH="18%"><FONT FACE="verdana" COLOR="#FFCC00" SIZE="1"><B>Address</B></FONT></TD>
          <TD BGCOLOR="#000000" WIDTH="18%"><FONT FACE="verdana" COLOR="#FFCC00" SIZE="1"><B>Shipping Address</B></FONT></TD>
          <TD BGCOLOR="#000000" WIDTH="10%"><FONT FACE="verdana" COLOR="#FFCC00" SIZE="1"><B>Payment Method</B></FONT></TD>
        </TR>

<% do while not rs.eof 
payOther = false

	bgcolor = "#FFFFFF"
	tcolor = "#000000"

SQL = "select * from cisco where oNum = " & rs("oNum")

'response.write(SQL)
set rs2 = conn.execute(SQL)

select case rs2("cc_type")
	case "AE"
		cc_type = "American Express"
	case "MC"
		cc_type = "MasterCard"
	case "Visa"
		cc_type = "Visa"
	case "PO"
		cc_type = "Purchase Order"
	case else
		payOther = true
		cc_type = rs2("cc_type")
end select

do while not rs2.eof
'	itemStr = itemStr & "<FONT COLOR='BLUE'>" & mid(rs2("item"),1,instr(1,rs2("item")," ")) & "</FONT>" & mid(rs2("item"),instr(1,rs2("item")," ")) & " <FONT COLOR='Red'>(" & rs2("quantity") & ")</FONT><br>"
	itemStr = itemStr & rs2("item") & " <FONT COLOR='Red'>(" & rs2("quantity") & ")</FONT><br>"
  response.flush
	rs2.movenext
loop

rs2.movefirst

%>

        <TR>
          <TD BGCOLOR="#EEEEEE" valign="middle" ALIGN="center">
            <FONT COLOR="<%= tcolor %>" face="arial" size="2">
            <% if request("view") = "closed" then %>
              Closed
            <% else %>
              <INPUT TYPE="button" onClick="location.href=&quot;admin.asp?oNum=<%=rs2("oNum")%>&quot;" VALUE="CLOSE">
            <% end if %>
            </FONT>
          </TD>
          <TD BGCOLOR="<%= bgcolor %>" valign="top">
            <FONT COLOR="<%= tcolor %>" face="arial" size="2">
            <%= DateValue(rs2("dDate")) %><BR><%=  TimeValue(rs2("dDate")) %>
            </FONT>
          </TD>
          <TD BGCOLOR="<%= bgcolor %>" valign="top">
            <FONT COLOR="<%= tcolor %>" face="arial" size="2">
            <%= rs2("uon") %>
            </FONT>
          </TD>
          <TD BGCOLOR="<%= bgcolor %>" valign="top">
            <FONT COLOR="<%= tcolor %>" face="arial" size="1">
            <%= itemStr %>
            </FONT>
          </TD>
          <TD BGCOLOR="<%= bgcolor %>" valign="top">
            <FONT COLOR="<%= tcolor %>" face="arial" size="2">
            <% if rs2("Name")  <> "" then response.write(rs2("Name")  & "<BR>") %>
            <% if rs2("Title") <> "" then response.write(rs2("Title") & "<BR>") %>
            <% if rs2("Email") <> "" then response.write("<BR>E: " & rs2("EMail") & "<BR>") %>            
            <% if rs2("Phone") <> "" then response.write("P: " & rs2("Phone") & "<BR>") %>
            <% if rs2("Fax")   <> "" then response.write("F: " & rs2("Fax") & "<BR>") %>
            </FONT>
          </TD>
          <TD BGCOLOR="<%= bgcolor %>" valign="top">
            <FONT COLOR="<%= tcolor %>" face="arial" size="2">
            <% if rs2("Company")  <> "" then response.write(rs2("Company") & "<BR>") %>            
            <% if rs2("MailStop") <> "" then response.write(rs2("MailStop") & "<BR>") %>
            <% if rs2("Address")  <> "" then response.write(rs2("Address") & "<BR>") %>
            <% if rs2("Address2") <> "" then response.write(rs2("Address2") & "<BR>") %>
            <% response.write(rs2("city") & ", " & rs2("st") & "&nbsp;&nbsp;" & rs2("zip") & "<BR>") %>
            <% if rs2("Country") <> "" then response.write(rs2("Country") & "<BR>") %>                        
            </FONT>
          </TD>
          <TD BGCOLOR="<%= bgcolor %>" valign="top">
            <FONT COLOR="<%= tcolor %>" face="arial" size="2">
            <% if rs2("ship_Company")  <> "" then response.write(rs2("ship_Company") & "<BR>") %>                        
            <% if rs2("ship_Address")  <> "" AND rs2("ship_address") <> ", " then response.write(rs2("ship_Address") & "<BR>") %>
            <% if rs2("ship_Address2") <> "" then response.write(rs2("ship_Address2") & "<BR>") %>
            <% response.write(rs2("ship_city") & ", " & rs2("ship_state") & "&nbsp;&nbsp;" & rs2("ship_zip") & "<BR>") %>
            <% if instr(1,rs2("ship_country"),",") > 0 then
                response.write(mid(rs2("ship_country"),1,instr(1,rs2("ship_country"),",")-1) & "<BR>")
                response.write("<BR>U: " & mid(rs2("ship_country"),instr(1,rs2("ship_country"),",")+1))
               else 
                response.write ("<BR>U: " & rs2("ultimate_country"))
               end if
            %>    
            </FONT>
          </TD>
          <TD BGCOLOR="<%= bgcolor %>" valign="top">
            <FONT COLOR="<%= tcolor %>" face="arial" size="2">
            <% if payOther then response.write("Other:<br>")%>
            <%= cc_type %><BR>
            <%= rs2("poNum") %><BR>
            <%= rs2("cc_num") %><BR>
            <%= rs2("exDate") %><BR>
            <%= rs2("cName") %><BR>
            <% if rs2("taxable") = "y" then
                 response.write("<BR>Taxable<BR>")
               elseif rs2("taxable") = "n" then
                 response.write("<BR>Not Taxable<BR>")
               else
                 response.write("<BR>Tax Status Unknown<BR>")
               end if
            %>
            <%= rs2("tax_ID") %><BR> 
            </FONT>
          </TD>
        </TR>

<% 
itemStr = blank
response.flush
rs.movenext 
response.flush
loop %>

      </TABLE>
    </TD>
	</TR>
</TABLE>

</FORM>  
<% end if %>

<%


'conn.close
'rs.close
'rs2.close
'set conn = Nothing
set rs = Nothing
set rs2 = Nothing

Call Disconnect_Cisco()
%>

</FONT>
</BODY>
</HTML>

