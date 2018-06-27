<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Author:     P. Barbee
' Date:       11/15/2000
'             Dev
' --------------------------------------------------------------------------------------

response.buffer = true

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

if Session("Logon_user") <> "" then
	%>
	<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->
	<%
else
  response.redirect "/register/default.asp"
	site_id = 3
end if

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Site_Description & " - " & "Order Inquiry - Change Company"
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & "Order Inquiry - Change Company" & "</FONT>"
Top_Navigation  = False 
Side_Navigation = True
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%
response.write "<SPAN CLASS=Heading3>" & Translate("Order Inquiry - Change Company",Login_Language,conn) & "</SPAN><BR>"
response.write "<BR>"
	
'BackURL (took this out 2/1/02 PFB because we have a Home button on left nav
'Response.write "<SPAN CLASS=SmallBold>"
'Response.write "<A HREF=""" & BackURL & """>" & Translate("New Search",Login_Language,conn) & "</A>"
'response.write "</SPAN><P>"
	
sql = "select max(len(SalesmanName)) as mymax from order_salesmen"
set dbRS = conn.Execute(sql)
max_cust = dbRS("mymax")
set dbRS = nothing
	
sql = "select SalesmanName,SalesmanNumber,Source_system" & vbcrlf &_
	"from order_salesmen" & vbcrlf &_
	"order by SalesmanName"
	
set dbRS = conn.Execute(sql)

strResponse = request.QueryString("caller")

select case strResponse
  case "avail"
    strAction = "SW-Avail_Form.asp"
  case else
    strAction = "SW-Order_Inquiry_Form.asp"
end select

%>
<form name="oi_comp_input" action="<%=strAction%>" method="POST">
<input type="hidden" name="srcsystem">
<input type="hidden" name="rep_number">

<table border=1 width="90%" bordercolor="#666666" bgcolor="<%=Contrast%>" cellpadding=0 cellspacing=0>
  <tr>
    <td align="center">
		<table cellpadding=4 cellspacing=2 border=0 bgcolor="<%=Contrast%>" width="80%">
			<tr>
				<td align="center">
<%
Head_Num = Translate("Rep",Login_Language,conn)
Head_Cus = Translate("Rep Name",Login_Language,conn)
%>
					<select name="order_list" multiple size="20" ondblclick="Chk_list();" style="font-weight:Normal; font-size:10pt; font-family: Courier; font-style: Normal;">
          				<option style="background:Silver;" value="">
<%
' format the header line
Response.write "&nbsp;&nbsp;" & Head_Cus
for i = 1 to (max_cust-len(head_cus))
	Response.write "&nbsp;"
next
Response.write Head_Num & " # &nbsp;"
%></option>
          				<option value=""></option>
<%
'font-weight:Normal; font-size:10pt;  font-family: Courier; font-style: Normal;
do until dbRS.EOF
	name = dbRS("SalesmanName")
	num  = dbRS("SalesmanNumber")
	src  = dbRS("Source_system")
	
	Response.write "						<option value=""" & src & ":" & num & """>" & name
	for i=1 to (max_cust - len(name) + 2)
		Response.write "&nbsp;"
	next
	Response.write num
	Response.write "&nbsp;</option>" & vbcrlf
  
	dbRS.MoveNext
  
loop
set dbRS = nothing
%>
					</select>
				</td>
			<tr>
				<td align="center" BGCOLOR="#000000"><input CLASS=NavLeftHighlight1 type="button" value=" <%=Translate("Change",Login_Language,conn)%> " onclick="Chk_list();"></td>
			</tr>
		</table>
    </td>
  </tr>
</table>

</form>

</center>
<LI><%=Translate("Single Click to select the Customer Name from the above list, then click on [Change] or Double Click on the Customer Name.",Login_Language,conn)%></LI>

<script language="Javascript">
	function Chk_list() {
		var df = document.oi_comp_input;
		var dfl = document.oi_comp_input.order_list;
		
		for (var i=0;i<dfl.length;i++) {
			if (dfl.options[i].selected) {
				var vals = dfl.options[i].value.split(":");
				df.srcsystem.value = vals[0];
				df.rep_number.value = vals[1];
				df.submit();
				//alert(df.srcsystem.value + ' - ' + df.customer_number.value);
			}
		}
	}
</script>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
response.flush
%>