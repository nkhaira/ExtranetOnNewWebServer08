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

strResponse = request.QueryString("caller")

select case LCase(strResponse)
  case "avail"
    strTitle = "Availability - Change Company"
  case else  
    strTitle = "Order Inquiry - Change Company"
end select
    
Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate(strTitle,Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) &_
                  "<BR><FONT CLASS=SmallBoldGold>" & _
                  Translate(strTitle,Login_Language,conn) & "</FONT>"
Top_Navigation  = False 
Side_Navigation = True
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->
<%
response.write "<SPAN CLASS=Heading3>" & Translate(strTitle,Login_Language,conn) & "</SPAN><BR>"
response.write "<BR>"

Call Table_Begin
response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Home"" VALUE=""" & " " & Translate("Home",Login_Language,conn) & """ "
response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & BackURL & "'; return false;"" TITLE=""Return to Order Inquiry Select"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
Call Table_End
	
sql = "select max(len(Customer_Name)) as mymax from orders"
set dbRS = conn.Execute(sql)
max_cust = dbRS("mymax")
set dbRS = nothing
	
sql = "select Customer_Name,Customer_Number,Source_system" & vbcrlf &_
	"from orders" & vbcrlf &_
	"Group by Customer_Name,Customer_Number,Source_system" & vbcrlf &_
	"order by Customer_Name"
	
set dbRS = conn.Execute(sql)

select case strResponse
  case "avail"
    strAction = "SW-Avail_Form.asp"
  case else
    strAction = "SW-Order_Inquiry_Form.asp"
end select

response.write "<DIV ALIGN=CENTER>"

%>
<form name="oi_comp_input" action="<%=strAction%>" method="POST">
<input type="hidden" name="srcsystem">
<input type="hidden" name="customer_number">

<% Call Table_Begin %>

<table border=1 width="90%" CLASS=tablebackground cellpadding=0 cellspacing=2>
  <tr>
    <td align="center">
<%
Head_Num = Translate("Customer",Login_Language,conn)
Head_Cus = Translate("Customer Name",Login_Language,conn)
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
	name = dbRS("Customer_Name")
	num  = dbRS("Customer_Number")
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
				<td align="center"><input CLASS=NavLeftHighlight1 type="button" value=" <%=Translate("Change",Login_Language,conn)%> " onclick="Chk_list();" onmouseover="this.className='NavLeftButtonHover'" onmouseout="this.className='Navlefthighlight1'"></td>
			</tr>
		</table>

<% Call Table_End %>

</form>

</DIV>

<LI><%=Translate("Single Click to select the Customer Name from the above list, then click on [Change] or Double Click on the Customer Name.",Login_Language,conn)%></LI>

<script language="Javascript">
	function Chk_list() {
		var df = document.oi_comp_input;
		var dfl = document.oi_comp_input.order_list;
		
		for (var i=0;i<dfl.length;i++) {
			if (dfl.options[i].selected) {
				var vals = dfl.options[i].value.split(":");
				df.srcsystem.value = vals[0];
				df.customer_number.value = vals[1];
				df.submit();
				//alert(df.srcsystem.value + ' - ' + df.customer_number.value);
			}
		}
	}
</script>

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
response.flush
%>