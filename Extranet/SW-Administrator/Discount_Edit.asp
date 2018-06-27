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
'site_id = Request("Site_ID")
'logon_user = Request("logon_user")
'fluke_id = Request("fluke_id")
'business_system = Request("bus_sys")

Bar_Tag = Translate("Discounts Edit",Login_Language,conn)

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
  .write "<SPAN  class=""SmallBoldWhite"">Companies starting with: </span>" & vbcrlf
  .write "<select name=""filterby"" onchange=""document.discount.dofilter.value=1;"
  .write "document.discount.submit();"" class=""Small"">" & vbcrlf


for each var in aAlpha
  if Request.Form("filterby") = var then
    .write "  <option value=""" & var & """ SELECTED>" & var & "</option>" & vbcrlf
  else
    .write "  <option value=""" & var & """>" & var & "</option>" & vbcrlf
  end if
next
.write "</select>"
Call Nav_Border_End

if Request.Form("PostFlag") = "1" and Request.Form("dofilter") <> "1" then
	Update_table
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
	
	if request.form("ldelete") <> "" then
    aTempD = split(request.form("ldelete"),", ")
    for each var in aTempD
      aLtempD = split(trim(var),"-")
      
    	sql = "DELETE avail_Discounts" & vbcrlf &_
    		"WHERE CustomerNumber = " & trim(aLtempD(0)) & vbcrlf &_
        "and Site_ID = " & trim(aLtempD(1))
        
      conn.Execute(sql)
      Response.write "<BR><SPAN CLASS=SmallBoldRed>" & Translate("Deleted",Login_Language,conn) & ": " & trim(aLtempD(0)) & " - " & trim(aLtempD(1)) & "<BR>" & vbcrlf
    next
  end if
	
	if request.form("writenew") = "1" then
        
    sql = "INSERT avail_Discounts" & vbcrlf &_
        "(CustomerName,CustomerNumber,Site_ID,BusinessSystem,PriceListName,DiscountRate)" & vbcrlf &_
        "VALUES('" & replace(trim(request.form("nCName")),"'","''") & "'" & vbcrlf &_
        "," & trim(request.form("nCNum")) & vbcrlf &_
        "," & trim(request.form("nSite_ID")) & vbcrlf &_
        ",'" & trim(request.form("nBusSys")) & "'" & vbcrlf &_
        ",'" & replace(trim(request.form("nPriceList")),"'","''") & "'" & vbcrlf &_
        "," & trim(request.form("nDRate")) & ")"
    
    'response.write replace(sql,vbcrlf,"<BR>"&vbcrlf)  & "<BR>" & vbcrlf
    conn.Execute(sql)
    Response.write "<BR><SPAN CLASS=SmallBoldRed>" & Translate("Inserted",Login_Language,conn) & ": " & request.form("nCName") & " - "
    Response.write asites(Cint(trim(request.form("nSite_ID")))) & "</SPAN><BR>" & vbcrlf
  end if
	
	if request.form("saveedit") = "1" then
        
    sql = "update avail_Discounts" & vbcrlf &_
        "set CustomerName = '" & replace(trim(request.form("eCName")),"'","''") & "'" & vbcrlf &_
        ",BusinessSystem = '" & trim(request.form("eBusSys")) & "'" & vbcrlf &_
        ",PriceListName = '" & replace(trim(request.form("ePriceList")),"'","''") & "'" & vbcrlf &_
        ",DiscountRate = " & trim(request.form("eDRate")) & vbcrlf &_
		"WHERE CustomerNumber = " & trim(request.form("eCNum")) & vbcrlf &_
        "and Site_ID = " & trim(request.form("eSite_ID"))
    
    'response.write replace(sql,vbcrlf,"<BR>"&vbcrlf)  & "<BR>" & vbcrlf
    conn.Execute(sql)
    Response.write "<BR><SPAN CLASS=SmallBoldRed>" & Translate("Modified",Login_Language,conn) & ": " & request.form("eCName") & " - "
    Response.write asites(Cint(trim(request.form("eSite_ID")))) & "</SPAN><BR>" & vbcrlf
  end if
end sub

sub Create_form
	' make a simple array of site names
  Dim eSiteID,eCustNum
  eSiteID = 0
  eCustNum = 0
    
	' now we need the basics of the form
	with Response
		.write "<input type=""hidden"" name=""PostFlag"">" & vbcrlf
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
		.write "    <TD width=""25"">&nbsp;</td>" & vbcrlf
		.write "    <TD class=""SmallBold"">"
		.write Translate("Customer Name",Login_Language,conn) & "</td>" & vbcrlf
		.write "    <TD class=""SmallBold"">"
		.write Translate("Customer Number",Login_Language,conn) & "</td>" & vbcrlf
		.write "    <TD class=""SmallBold"" NOWRAP>"
		.write Translate("Site",Login_Language,conn) & "</td>" & vbcrlf
		.write "    <TD class=""SmallBold"">"
		.write Translate("Business System",Login_Language,conn) & "</td>" & vbcrlf
		.write "    <TD class=""SmallBold"">"
		.write Translate("Price List Name",Login_Language,conn) & "</td>" & vbcrlf
		.write "    <TD class=""SmallBold"">"
		.write Translate("Discount Rate",Login_Language,conn) & "</td>" & vbcrlf
		.write "    <TD width=""25""></td>" & vbcrlf
		.write "  </TR>" & vbcrlf
    .write "  <TR>" & vbcrlf
    .write "    <TD class=""SmallBold"">New</td>" & vbcrlf
		.write "    <TD>"
		.write "<input class=""Small"" type=""text"" name=""nCName"" size=25></td>" & vbcrlf
		.write "    <TD>"
		.write "<input class=""Small"" type=""text"" name=""nCNum"" size=8></td>" & vbcrlf
		.write "    <TD><select class=""Small"" name=""nSite_ID"">" & vbcrlf
    
    for i = 0 to ubound(asites)
      if asites(i) <> "" then
        if i = 14 then
          .write "      <option value=""" & i & """ SELECTED>"
          .write asites(i) & "</option>" & vbcrlf
        else
          .write "      <option value=""" & i & """>" & asites(i)
          .write "</option>" & vbcrlf
        end if
      end if
    next
    
		.write "     </select>" & vbcrlf
		.write "    <TD><input "
		.write "class=""Small"" type=""text"" name=""nBusSys"" value=""ORA"" size=4></td>" & vbcrlf
		.write "    <TD>"
		.write "<input class=""Small"" type=""text"" name=""nPriceList"" size=25></td>" & vbcrlf
		.write "    <TD>"
		.write "<input class=""Small"" type=""text"" name=""nDRate"" size=6></td>" & vbcrlf
		.write "  </TR>" & vbcrlf
    
    ' write the edit table
    if Request.Form("ledit") <> "" then
      ' get the row to be edited
      aTemp1 = split(Request.Form("ledit"),"-")
      eCustNum = Trim(aTemp1(0))
      eSiteID = Trim(aTemp1(1))
    	sql = "SELECT CustomerName" & vbcrlf &_
            ",CustomerNumber" & vbcrlf &_
            ",Site_ID" & vbcrlf &_
            ",BusinessSystem" & vbcrlf &_
            ",PriceListName" & vbcrlf &_
            ",DiscountRate" & vbcrlf &_
            "FROM avail_Discounts" & vbcrlf &_
    		"WHERE CustomerNumber = " & eCustNum & vbcrlf &_
            "and Site_ID = " & eSiteID

    	set dbRS = conn.Execute(sql)
    	
    	if dbRS.EOF then
    		.write "Odd, Could not find row to edit???<BR>" & vbcrlf
      else
        .write "  <TR>" & vbcrlf
        .write "    <TD class=""MediumBold"">Edit</td>" & vbcrlf
    		.write "    <TD><input onchange=""SEdit();"" "
    		.write "class=""Small"" type=""text"" name=""eCName"" value=""" & dbRS("CustomerName")
        .write """ size=25></td>" & vbcrlf
    		.write "    <TD class=""Small"" ALIGN=""RIGHT"">" & dbRS("CustomerNumber") & "</td>" & vbcrlf
    		.write "    <TD class=""Small"" NOWRAP>" & asites(Cint(dbRS("Site_ID"))) & "</td>" & vbcrlf
    		.write "    <TD><input onchange=""SEdit();"" "
    		.write "class=""Small"" type=""text"" name=""eBusSys"" value=""" & dbRS("BusinessSystem")
        .write """ size=4></td>" & vbcrlf
    		.write "    <TD class=""small"" NOWRAP><input onchange=""SEdit();"" class=""Small"" "
    		.write "type=""text"" name=""ePriceList"" value=""" & dbRS("PriceListName")
        .write """ size=25></td>" & vbcrlf
    		.write "    <TD Class=""Small"" ALIGN=""RIGHT""><input onchange=""SEdit();"" "
    		.write "class=""Small"" type=""text"" name=""eDRate"" value=""" & dbRS("DiscountRate")
        .write """ size=6></td>" & vbcrlf
    		.write "    </TR>" & vbcrlf
        .write "</TABLE>"
        .write "</td></tr></TABLE>" & vbcrlf
        Call Table_End
        .write "</DIV>"
        .write "<input type=""Hidden"" name=""eCNum"" value=""" & dbRS("CustomerNumber")
        .write """>" & vbcrlf
        .write "<input type=""Hidden"" name=""eSite_ID"" value=""" & dbRS("Site_ID")
        .write """>" & vbcrlf
    	end if
    	set dbRS = Nothing
    else
      .write "</TABLE>"
      .write "</td></tr></TABLE>" & vbcrlf
      Call Table_End
      .write "</DIV>"
      .write "<input type=""Hidden"" name=""eCName"">" & vbcrlf
      .write "<input type=""Hidden"" name=""eBusSys"">" & vbcrlf
      .write "<input type=""Hidden"" name=""ePriceList"">" & vbcrlf
      .write "<input type=""Hidden"" name=""eDRate"">" & vbcrlf
    end if
	  
  	' get the current SQL table
  	sql = "SELECT CustomerName" & vbcrlf &_
          ",CustomerNumber" & vbcrlf &_
          ",Site_ID" & vbcrlf &_
          ",BusinessSystem" & vbcrlf &_
          ",PriceListName" & vbcrlf &_
          ",DiscountRate" & vbcrlf &_
          "FROM avail_Discounts" & vbcrlf
    
    if len(Request.Form("filterby")) = 1 then
    	sql = sql & "WHERE CustomerName like '" & Request.Form("filterby") & "%'" & vbcrlf
    end if
    
  	sql = sql & "ORDER by CustomerName, Site_ID"
  	
  	set dbRS = conn.Execute(sql)
  	
  	if dbRS.EOF then
  		.write "</form><BR>" & Translate("Table is empty.",Login_Language,conn) & "<BR>"
  		set dbRS = Nothing
  		exit sub
  	end if
	  
    ' start the main table
    .write "<BR>" & vbcrlf
    .write "<DIV ALIGN=CENTER>"
    Call Table_Begin
    .write "<TABLE BORDER=1 WIDTH=""100%"" BORDERCOLOR=""#666666"" BGCOLOR=""#FFCC00"""
    .write " CELLPADDING=0 CELLSPACING=0>" & vbcrlf
    .write "  <TR>" & vbcrlf
    .write "    <TD align=""center"">" & vbcrlf
    
    .write "<DIV ID=""ContentTableStart"" STYLE=""position: absolute;"">" & vbCrLf
    .write "</DIV>" & vbCrLf
    
    .write "		<TABLE CELLPADDING=2 CELLSPACING=0 BORDER=0 BGCOLOR=""black"" WIDTH=""100%"" >" & vbcrlf
    .write "			<TR ID=""ContentHeader1"" BGCOLOR=""BLACK"">" & vbcrlf
		.write "                <TD class=""SmallBoldGold"" BGCOLOR=""Black"">"
		.write Translate("Edit",Login_Language,conn) & "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"">"
		.write Translate("Customer Name",Login_Language,conn) & "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"" ALIGN=RIGHT>"
		.write Translate("Customer Number",Login_Language,conn) & "&nbsp;&nbsp;&nbsp;&nbsp;</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"">"
		.write Translate("Site",Login_Language,conn) & "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"">"
		.write Translate("Business System",Login_Language,conn) & "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"">"
		.write Translate("Price List Name",Login_Language,conn) & "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"" ALIGN=RIGHT>"
		.write Translate("Discount Rate",Login_Language,conn) & "</td>" & vbcrlf
		.write "                <TD class=""SmallBoldGold"" ALIGN=CENTER>"
		.write Translate("Delete",Login_Language,conn) & "</td>" & vbcrlf
		.write "              </TR>" & vbcrlf  
	  
	  do until dbRS.EOF
      sLineCustNum = dbRS("CustomerNumber")
      sLine_SiteID = dbRS("Site_ID")
      sLineID = sLineCustNum & "-" & sLine_SiteID
      
      if (Clng(eCustNum) = Clng(sLineCustNum)) AND (Clng(eSiteID) = Clng(sLine_SiteID)) then
        ShowEdit = "SmallWhite"
      else
        ShowEdit = "Small"
      end if

       if toggle = "#EAEAEA" then toggle = "#DBDBDB" else toggle = "#EAEAEA"

			.write "              <TR>" & vbcrlf
			.write "                <TD class=""" & ShowEdit & """ BGCOLOR=""" & Toggle & """>"
			.write "<input type=""radio"" name=""ledit"" value=""" & sLineID & """></td>" & vbcrlf
			.write "                <TD class=""Small"" NOWRAP BGCOLOR=""" & Toggle & """>"
			.write dbRS("CustomerName") & "</td>" & vbcrlf
			.write "                <TD class=""Small"" ALIGN=""RIGHT"" BGCOLOR=""" & Toggle & """>"
			.write sLineCustNum & "&nbsp;&nbsp;&nbsp;&nbsp;</td>" & vbcrlf
			.write "                <TD class=""Small"" NOWRAP BGCOLOR=""" & Toggle & """>"
			.write asites(Cint(sLine_SiteID)) & "</td>" & vbcrlf
			.write "                <TD class=""Small"" ALIGN=CENTER BGCOLOR=""" & Toggle & """>"
			.write dbRS("BusinessSystem") & "</td>" & vbcrlf
			.write "                <TD class=""Small"" NOWRAP BGCOLOR=""" & Toggle & """>"
			.write dbRS("PriceListName") & "</td>" & vbcrlf
			.write "                <TD class=""Small"" ALIGN=RIGHT BGCOLOR=""" & Toggle & """>"
			.write dbRS("DiscountRate") & "</td>" & vbcrlf
			.write "                <TD class=""Small"" ALIGN=CENTER BGCOLOR=""" & Toggle & """>"
			.write "<input type=""checkbox"" name=""ldelete"" value=""" & sLineID & """>"
			.write "</td>" & vbcrlf
      .write "              </TR>" & vbcrlf
		
		  dbRS.MoveNext
	  loop
	  set dbRS = Nothing
    
    .write "			<TR>" & vbcrlf
    .write "			  <TD colspan=26 align=""center"" bgcolor=""#666666"">" & vbcrlf
    .write "			  	<input type=""button"" onclick=""Update();"" value="""
    .write Translate("Update",Login_Language,conn) & """ CLASS=""NavLeftHighlight1"">" & vbcrlf
    .write "                &nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf
    .write "			  	<input type=""reset"" value="""
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
    .write "  <li>" & Translate("Select the row you want to edit with the left hand column of radios.",Login_Language,conn) & "</li>" & vbcrlf
    .write "  <li>" & Translate("Select the rows you want to delete with the right hand column of checkboxes, you may delete more than one row at a time.",Login_Language,conn) & "</li>" & vbcrlf
    .write "  <li>" & Translate("You may add a new row, edit a row, and choose the next row to edit (or delete) all in the same screen.",Login_Language,conn) & "</li>" & vbcrlf
    .write "  <li>" & Translate("If you change the <B>Companies starting with</b> it will do so immediately with no ther action.",Login_Language,conn) & "</li>" & vbcrlf
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
        
        if ((df.nCName.value.length) || (df.nCNum.value.length)) {
            if ((df.nCName.value.length) && (df.nCNum.value.length) &&
                (df.nBusSys.value.length) && (df.nPriceList.value.length) && (df.nDRate.value.length)) {
                df.writenew.value=1;
                cnt = 1;
            }
            else {
                alert('<%=Translate("Either all, or none, of the textboxes in the New row must be used.",Alt_Language,conn)%>');
                return;
            }
        }
        
        if (df.eCName.value.length) {
            if (!((df.eCName.value.length) && (df.eBusSys.value.length) &&
                (df.ePriceList.value.length) && (df.eDRate.value.length))) {
                alert('<%=Translate("Either all, or none, of the textboxes in the Edit row must be used.",Alt_Language,conn)%>');
                df.saveedit.value=0;
                return;
            }
        }
        
        if (df.saveedit.value == 1) { cnt = 1; }
        
        for (i = 0;i<df.ledit.length;i++) {
            if (df.ledit[i].checked) {
                radval = df.ledit[i].value;
            }
        }
        for (i = 0;i<df.ldelete.length;i++) {
            if (df.ldelete[i].checked) {
                chkval += df.ldelete[i].value + '|';
            }
        }
        cnt += radval.length + chkval.length;
        
        if (cnt < 1) {
            alert('<%=Translate("No action by user, nothing to do ??",Alt_Language,conn)%>');
        }
        else {
            //alert('Submit form:\n\tedit: ' + radval + '\n\tdelete: ' + chkval);
            df.PostFlag.value = 1;
            df.submit();
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

</SCRIPT>

<%
end sub
%>
