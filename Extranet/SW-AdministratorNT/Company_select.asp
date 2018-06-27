<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%
' --------------------------------------------------------------------------------------
' Author:     Peter Barbee
' Date:       01/09/2002
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim HomeURL
Dim BackURL

Dim Site_ID

Dim SQL
Dim SQL_Extract

Dim Email_Delimited
' --------------------------------------------------------------------------------------
' Decode QueryString Parameters
' --------------------------------------------------------------------------------------
site_id = Request("Site_ID")
logon_user = Request("logon_user")
Language = Request("Language")
business_system = Request("bus_sys")

' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_Login_Admin.asp"-->
<%

Bar_Tag = Translate("Company Selection",Login_Language,conn)

' --------------------------------------------------------------------------------------
' Determine Site Code and Description based on Site_ID Number
' --------------------------------------------------------------------------------------

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
'Response.write SQL & "<P>"
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code        = rsSite("Site_Code")     
Site_Description = rsSite("Site_Description")
Screen_Title     = rsSite("Site_Description") & " - " & Translate("Account Administrator",Alt_Language,conn)
Bar_Title        = rsSite("Site_Description") & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Account Administrator",Login_Language,conn) & "</FONT>"
Bar_Title        = Bar_Title & "<BR><FONT CLASS=SmallBoldGold>" & Translate("List",Login_Language,conn) & " / " & Translate("Edit",Login_Language,conn) & " " & Translate("Group",Login_Language,conn) & ": " & Bar_Tag & "</FONT>"

Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<%

rsSite.close
set rsSite=nothing

sql = "select max(len(Company)) as mymax" & vbcrlf &_
	"FROM  UserData" & vbcrlf &_
	"WHERE (Fluke_ID IS NOT NULL)" & vbcrlf &_
	"AND (Business_System IS NOT NULL)" & vbcrlf &_
	"AND (Site_ID = " & site_id & ")"
	
set dbRS = conn.Execute(sql)
    
maxlen = Cint(dbRS("mymax"))
max_cust = maxlen
dbRS.Close
	
sql = "SELECT Company, Fluke_ID, Business_System" & vbcrlf &_
	"FROM  UserData" & vbcrlf &_
	"WHERE (Fluke_ID IS NOT NULL)" & vbcrlf &_
	"AND (Business_System IS NOT NULL)" & vbcrlf &_
	"AND (Site_ID = " & site_id & ")" & vbcrlf &_
	"order by Company, Fluke_ID"
	
set dbRS = conn.Execute(sql)

%>
<DIV ALIGN=CENTER>

<FORM NAME="comp_input" METHOD="POST">
<INPUT TYPE="hidden" NAME="srcsystem">
<INPUT TYPE="hidden" NAME="customer_number">
<%Call Table_Begin%>
<TABLE BORDER=1 WIDTH="90%" BORDERCOLOR="#666666" BGCOLOR=<%=Contrast%> CELLPADDING=0 CELLSPACING=2>
  <TR>
    <TD ALIGN="center">
      <%
      Head_Num = Translate("Customer",Login_Language,conn)
      Head_Cus = Translate("Customer Name",Login_Language,conn)
      %>
			<SELECT NAME="order_list" MULTIPLE SIZE="20" ondblclick="Chk_list();" style="font-weight:Normal; font-size:10pt; font-family: Courier; font-style: Normal;">
      <%

      ' format the header line
      response.write "<option style=""background:Silver;"" value="""">"
      Response.write "&nbsp;&nbsp;" & Head_Cus
      for i = 1 to (max_cust-len(head_cus))
    	  Response.write "&nbsp;"
      next
      Response.write Head_Num & " # &nbsp;"
      response.write "</option>"
      response.write "<option value=""""></option>"

pnum = ""
psrc = ""
do until dbRS.EOF
	name = dbRS("Company")
	num = dbRS("Fluke_ID")
	src = dbRS("Business_system")
	
	if (num <> pnum) OR (src <> psrc) then
		Response.write "<option value=""" & src & ":" & num & """>" & Trim(UCase(name))
		'Response.write name & "&nbsp;&nbsp;&nbsp;"
    for i=1 to (max_cust - len(Trim(name)) + 2)    
			Response.write "&nbsp;"
		next
		Response.write num & "</option>" & vbcrlf
	end if
	
	pnum = num
	psrc = src
	dbRS.MoveNext
loop
set dbRS = nothing
%>
					</SELECT>
				</TD>
			<TR>
				<TD ALIGN="center"BGCOLOR="#666666"><INPUT CLASS=NAVLEFTHIGHLIGHT1 TYPE="button" VALUE=" Change " onclick="Chk_list();"></TD>
			</TR>
		<!--/table>
    </TD>
  </TR-->
</TABLE>
<%Call Table_End%>
</FORM>

</DIV>
<P><FONT CLASS=MEDIUM>
<DIV ALIGN=CENTER>
<LI><%=Translate("Single Click to select the Customer Name from the above list, then click on [Change] or Double Click on the Customer Name.",Login_Language,conn)%></LI>
</DIV>

<SCRIPT LANGUAGE="Javascript">
	function Chk_list() {
		var pf = opener.document.branch;
		var dfl = document.comp_input.order_list;
		
		for (var i=0;i<dfl.length;i++) {
			if (dfl.options[i].selected) {
				var vals = dfl.options[i].value.split(":");
				pf.bus_sys.value = vals[0];
				pf.fluke_id.value = vals[1];
				self.close();
				parent.focus();
				pf.submit();
			}
		}
	}
</SCRIPT>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

sub Table_Begin()
    response.write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" CLASS=TableBorder>" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif""            BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLr
    response.write "  <TR>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD VALIGN=""top"">" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Table_End()
    response.write "    </TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_BL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_BR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLf
    response.write "</TABLE>" & vbCrLf
end sub  

Call Disconnect_SiteWide
Response.end
%>
  
