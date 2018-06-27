<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/Include/Adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Author:     Peter Barbee
' Date:       01/09/2002
'             K. Whitlock 7/5/2003 Updated to Site Look and Feel Standard and fixed alignments
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim HomeURL
Dim BackURL

Dim Site_ID

Dim SQL
Dim bShow_Main

Dim asites(100),aAlpha
asites(1) = "Calibration-Sales"
asites(3) = "FInd-Sales"
asites(14) = "Pomona-Sales"
aAlpha = Array("Show All","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R",_
    "S","T","U","V","W","X","Y","Z")

%>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_Login_Admin.asp"-->
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%
site_id = Request("Site_ID")
'logon_user = Request("logon_user")
'fluke_id = Request("fluke_id")
'business_system = Request("bus_sys")

Bar_Tag = Translate("Associate Price List Access Code with Oracle Customer Number",Login_Language,conn)

' --------------------------------------------------------------------------------------
' Determine Site Code and Description based on Site_ID Number
' --------------------------------------------------------------------------------------

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code        = rsSite("Site_Code")     
Site_Description = rsSite("Site_Description")
Screen_Title     = Translate(rsSite("Site_Description"),Alt_Language,conn) & " - " & Translate("Account Administrator",Alt_Language,conn)
Bar_Title        = Translate(rsSite("Site_Description"),Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Account Administrator",Login_Language,conn) & "</FONT>"
Bar_Title        = Bar_Title & "<BR><FONT CLASS=SmallBoldGold>" & Translate("List",Login_Language,conn) & " / " & Translate("Edit",Login_Language,conn) & " " & Translate("Price List Access Code",Login_Language,conn) & ": " & Bar_Tag & "</FONT>"

Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<%

rsSite.close
set rsSite=nothing

response.write "<FONT CLASS=NormalBoldRed>"
select case Admin_Access
  case 2
    response.write Translate("Content Submitter",Login_Language,conn)
  case 4
    response.write Translate("Content Administrator",Login_Language,conn)
  case 6
    response.write Translate("Account Administrator",Login_Language,conn)
  case 8
    response.write Translate("Site Administrator",Login_Language,conn)
  case 9
    response.write Translate("Domain Administrator",Login_Language,conn)
end select

with response
	.write "</FONT><BR><FONT CLASS=MediumBold>" & Admin_FirstName & " " & Admin_LastName & "<BR>"
	.write Admin_Company & "</FONT><BR><BR>"

' put a tiny menu in
  .write vbcrlf & "<form name=""discount"" method=""POST"">" & vbcrlf
  Call Nav_Border_Begin
  .write "<input type=""hidden"" name=""dofilter"">" & vbcrlf
	'.write "<A HREF=""/" & site_code & "/default.asp?Site_ID=" & site_id & "&Language=" & Login_Language
	'.write "&NS=true&CID=9999&SCID=0&PCID=0&CIN=0&CINN=0"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;"
	'.write Translate("Logoff",Login_Language,conn) & "&nbsp;&nbsp;</A>&nbsp;&nbsp;"
  .write "<A HREF=""default.asp?Site_ID=" & site_id
  .write """ CLASS=NavLeftHighlight1>&nbsp;&nbsp;Main Menu&nbsp;&nbsp;</A>&nbsp;&nbsp;"
  .write "<A href=""#instruct""  CLASS=""NavLeftHighlight1""> Instructions </a>&nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf
'  .write "<SPAN  class=""SmallBoldWhite"">Companies starting with: </span>" & vbcrlf
 ' .write "<select name=""filterby"" onchange=""document.discount.dofilter.value=1;"
  '.write "document.discount.submit();"" class=""Small"">" & vbcrlf


Call Nav_Border_End

Custid1=request.form("Custid")
if Request.Form("PostFlag") = "1" then
	Update_table
end if
if Request.Form("PostFlag") = "2" then
	Delete_table
end if

'response.write "Access Level: " & admin_access & "<BR>"
' for now only admin_access > 5 (subgroup => account) accesses this so:
' this may need a change when we introduce distributor-based branch editing
if admin_access > 5 then
	Create_form
else
	.write Translate("You do not have access to this function",Login_Language,conn)
end if

end with
%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide
Response.end

' ------------------------ end of main -------- subroutines below -------------

sub Update_table
	
	dim strSQL
	dim strRS
	set strRS=server.createobject("Adodb.Recordset")
	dim strPrice
	strPrice=request.form("PriceList")

	strSql="select *from Customer_PriceLists where CustomerNumber= '" & CustID1 & "' and Site_ID=" & Site_ID & ""	
	strRs.open strSQL, conn
	if strRS.eof then
		conn.execute "insert into  Customer_PriceLists(CustomerNumber,Site_ID,BusinessSystem,PriceListCode) values('" & CustID1 & "'," & Site_ID & ",'ORA','" & strPrice & "')"
Response.write "<center><font face=Arial size=2 color=red>Price List Access Code Added successfully...</font></center>"

else 

	
        
    sql = "update Customer_PriceLists" & vbcrlf &_
       
        "set PriceListCode='" & replace(trim(request.form("PriceList")),"'","''") & "'" & vbcrlf &_
		"WHERE CustomerNumber = '" & trim(request.form("CustID")) & "'" & vbcrlf &_
        "and Site_ID = " & Site_ID
    Response.write "<center><font face=Arial size=2 color=red>Price List Access Code Updated successfully...</font></center>"

    'response.write replace(sql,vbcrlf,"<BR>"&vbcrlf)  & "<BR>" & vbcrlf
'response.write sql
    conn.Execute(sql)
end if

'Response.write "<center><font face=Arial size=2 color=red>Price Code Updated Successfully</font></center>"
    'Response.write "<BR><SPAN CLASS=SmallBoldRed>" & Translate("Modified",Login_Language,conn) & ": " & request.form("eCName") & " - "
    'Response.write asites(Cint(trim(request.form("eSite_ID")))) & "</SPAN><BR>" & vbcrlf
  'end if

end sub
sub Delete_table
	
	dim strSQL
	dim strRS
	set strRS=server.createobject("Adodb.Recordset")
	dim strPrice
	strPrice=request.form("PriceList")
	
	
	strSql="select *from Customer_PriceLists where CustomerNumber= '" & CustID1 & "' and Site_ID=" & Site_ID & ""	
	strRs.open strSQL, conn
	if not strRS.eof then
		conn.execute "delete from Customer_PriceLists where CustomerNumber= '" & CustID1 & "' and Site_ID= " & Site_ID & ""	
	     
		Response.write "<center><font face=Arial size=2 color=red>Price List Access Code Deleted successfully...</font></center>"
	else
		Response.write "<center><font face=Arial size=2 color=red>Price List Access Code Deleted successfully...</font></center>"
	end if

end sub

sub Create_form
	' make a simple array of site names
  Dim eSiteID,eCustNum
  eSiteID = 0
  eCustNum = 0
    
	' now we need the basics of the form
	with Response
		.write "<input type=""hidden"" name=""PostFlag"" value=""1"">" & vbcrlf

	.write "<input type=""hidden"" name=""CustID1"">" & vbcrlf
    .write "<input type=""Hidden"" name=""writenew"">" & vbcrlf
    .write "<input type=""Hidden"" name=""saveedit"" value=""0"">" & vbcrlf

    ' write the new table
    .write "<BR>" & vbcrlf
    .write "<DIV ALIGN=CENTER>"

           
    Call Table_Begin

    .write "<TABLE BORDER=0 WIDTH=""100%"" BORDERCOLOR=""#666666"" BGCOLOR=""#FFCC00"""
    .write " CELLPADDING=0 CELLSPACING=0>" & vbcrlf
    .write "  <TR><TD>"

    .write "<TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR=""#FFCC00"""
    .write " WIDTH=""100%"">" & vbcrlf
    .write "  <TR>" & vbcrlf
		.write "    <TD class=""SmallBold"" Align=""LEFT"" width=""30%"">Business System: </td>" & vbcrlf

	.write "    <TD class=""SmallBold"" Align=""LEFT"">"
		.write "ORACLE</td>" & vbcrlf

		.write "  </TR>" & vbcrlf


.write "  <TR>" & vbcrlf
		.write "    <TD width=""25"">&nbsp;</td>" & vbcrlf
	
		.write "  </TR>" & vbcrlf


    .write "  <TR>" & vbcrlf
    .write "    <TD class=""SmallBold"" Align=""LEFT"" >Select Customer Number:</td>" & vbcrlf
	''''''''''''''''''''''''
dim SQL1

SQL1 = "Select distinct Fluke_ID from Userdata where Business_System='ORA' and Fluke_ID is not null and Fluke_ID <> '' and Fluke_ID <> '-' and Site_ID=" & Site_ID & ""
Set rsCustNum = Server.CreateObject("ADODB.Recordset")
rsCustNum.Open SQL1, conn, 3, 3
''''''''''''''''''''''''
		
		.write "    <TD><select class=""Small"" name=""CustID"" onchange=""return cust_onchange(discount)"">" & vbcrlf
    

 .write "      <option >"
          .write " --- Select Customer Number --- "& "</option>" & vbcrlf        
          .write "</option>" & vbcrlf







do while not rsCustNum.eof 

	if Request.form("CustID") = rsCustNum("Fluke_ID") then
		
    

          .write "      <option value=""" & rsCustNum("Fluke_ID") & """ SELECTED>"
          .write rsCustNum("Fluke_ID") & "</option>" & vbcrlf        
          .write "</option>" & vbcrlf
	else
          .write "      <option value=""" & rsCustNum("Fluke_ID") & """>"
          .write rsCustNum("Fluke_ID") & "</option>" & vbcrlf        
          .write "</option>" & vbcrlf

	end if  
     
 rsCustNum.movenext
loop
   rsCustNum.close()
		.write "     </select></td></tr>" & vbcrlf

	


.write "  <TR>" & vbcrlf
		.write "    <TD width=""25"">&nbsp;</td>" & vbcrlf
		
		.write "  </TR>" & vbcrlf
    .write "  <TR>" & vbcrlf
    .write "    <TD class=""SmallBold"" Align=""LEFT"">Enter Price List Access Code:</td>" & vbcrlf
		'.write "    <TD>"
		

    

''''''''''''''''''''''''
dim SQL2
dim Price

if not CustID1 = "" then
SQL2 = "Select PriceListCode from Customer_PriceLists where CustomerNumber= '" & CustID1 & "' and Site_ID=" & Site_ID & ""
Set rsPriceNum = Server.CreateObject("ADODB.Recordset")
rsPriceNum.Open SQL2 , conn, 3, 3
if not rsPriceNum.eof then
	Price=rsPriceNum("PriceListCode")
else
	Price=""
end if
''''''''''''''''''''''''
else
Price=""
end if
       
		.write "    <TD><input "
		.write "class=""Small"" type=""text"" name=""PriceList"" value=""" & Price & """ size=30 maxlength=30></td>" & vbcrlf
.write "     </td></tr>" & vbcrlf




	
    .write "<tr><td>&nbsp;</td></tr>"
    ' write the edit table
    
    .write "			<TR>" & vbcrlf
	.write "    <TD width=""25"">&nbsp;</td>" & vbcrlf
    .write "			  <TD colspan=50 align=""LEFT"" >" & vbcrlf
    .write "			  	<input type=""submit"" onclick=""return Update();"" value="""
    .write Translate("Save",Login_Language,conn) & """ CLASS=""NavLeftHighlight1"">" & vbcrlf
    .write "                &nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf
	.write "			  	<input type=""submit"" onclick=""return onDelete(this.form);"" value="""
    .write Translate("Delete",Login_Language,conn) & """ CLASS=""NavLeftHighlight1"">" & vbcrlf
    .write "                &nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf

    .write "			  	<input type=""button"" onclick=""return ClearFields();"" value=""" 
    .write Translate("Reset",Login_Language,conn) & """ CLASS=""NavLeftHighlight1"">" & vbcrlf
    .write "			  </td>" & vbcrlf
    .write "			</tr>" & vbcrlf
    .write "		</TABLE>" & vbcrlf
    .write "    </TD>" & vbcrlf
    .write "  </TR>" & vbcrlf
    .write "</TABLE>" & vbcrlf
    Call Table_End
    .write "</DIV>"
    .write "</FORM>" & vbcrlf
    .write "<P>" & vbcrlf
    .write "<A NAME=""Instruct""></A><span class=""SmallBold"">" & Translate("Instructions",Login_Language,conn) & "</span>" & vbcrlf
    .write "<ul class=""small"">" & vbcrlf
    .write "  <li>" & Translate("Use the Update button to execute any and all actions.",Login_Language,conn) & "</li>" & vbcrlf
    .write "  <li>" & Translate("Select the row, from the Customer Number drop down box, you want to edit and enter the price list access code in the textbox.",Login_Language,conn) & "</li>" & vbcrlf
    .write "  <li>" & Translate("For deleting the Price List Access Code associated with Customer number, select the Customer Number and click on Delete button.",Login_Language,conn) & "</li>" & vbcrlf
    .write "  <li>" & Translate("Click on the Reset button to clear the current data.",Login_Language,conn) & "</li>" & vbcrlf
  '  .write "  <li>" & Translate("If you change the <B>Companies starting with</b> it will do so immediately with no ther action.",Login_Language,conn) & "</li>" & vbcrlf
    .write "</ul>" & vbcrlf
	end with

' the javascript
%>
<SCRIPT language="JavaScript1.2">
    function SEdit() {
        document.discount.saveedit.value = 1;
    }
    
    function Update() {
        var df = document.discount;
        var i;
        var cnt = 0;
        var chkval = '';
        var radval = '';
        var badnew = false;

	if(df.CustID.selectedIndex==0)
{
		alert("Please select Customer Number.");
		return false;
}
if (df.PriceList.value == false)
{
	alert("Please enter Price List Access Code.")
	df.PriceList.focus();
	return false;
}
        
        
       
       
    }
</script>

<SCRIPT Language="Javascript">

  var headTop = -1;
  var FloatHead1;

  function processScroll() {
    if (headTop < 0) {
      saveHeadPos();
    }
    if (headTop > 0) {
      if (document.documentElement && document.documentElement.scrollTop)
        theTop = document.documentElement.scrollTop;
      else if (document.body)
        theTop = document.body.scrollTop;

    if (theTop > headTop)
      FloatHead1.style.top = (theTop-headTop) + 'px';
    else
      FloatHead1.style.top = '0px';
  }
}

function saveHeadPos() {
  parTable = document.getElementById("ContentTableStart");
  if (parTable != null) {
    headTop = parTable.offsetTop + 3;
    FloatHead1 = document.getElementById("ContentHeader1");
    FloatHead1.style.position = "relative";
  }
}

window.onscroll = processScroll;

function cust_onchange(theform)
{
	theform.PostFlag.value=0;
	theform.submit();
}
function onDelete(theform)
{
	
	if(theform.CustID.selectedIndex!=0)
	{

		var where_to=confirm("Are you sure you want to delete the Customer Number - Price List Association?")
		if(where_to==true)
		{
			theform.PostFlag.value=2;
			theform.submit();
		}
	}
	else
	{
		
		alert("Please select Customer Number.");
		return false;
	}

}
function ClearFields()
{
document.discount.PriceList.value = "";
//document.discount.CustID.options[document.discount.CustID.selectedIndex].value=document.discount.CustID.options[0].value;
document.discount.CustID.selectedIndex = 0;
document.discount.CustID.options[0].selected = true; 


}
</SCRIPT>
<% end sub %>