<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
'             Sandbox
' --------------------------------------------------------------------------------------

'response.buffer = true

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

if request("Site_ID") <> "" and isnumeric(request("Site_ID")) then
  Site_ID            = request("Site_ID")
  Session("Site_ID") = request("Site_ID")  
elseif session("Site_ID") <> "" and isnumeric(session("Site_ID")) then
  Site_ID = session("Site_ID")
else
  response.redirect "/register/default.asp"
end if

Dim BackURL
BackURL = Session("BackURL")    

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code = rsSite("Site Code")
Site_Description = rsSite("Site Description")
  
rsSite.close
set rsSite=nothing

Call Disconnect_SiteWide

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Dim TypeOfSearch
Dim strRate
Dim strDiscount
Dim strInvalid
Dim strStrip
Dim pagenum
Dim limit
Dim view
Dim rt

strRate = Session("StrRate")
strDiscount = Session("StrDiscount")
limit = Session("Perpage")
view = Session("view")
rt = Request("rt")

Screen_Title    = Site_Description & " - " & "US Replacement Parts Database"
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & "US Replacement Parts Database" & "</FONT>"
 
Side_Navigation = True
Content_Width   = 95  ' Percent
%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include file="SW-Common-Header.asp"-->
<!--#include file="SW-Common-Navigation.asp"-->

<%
response.write "<FONT CLASS=Heading3>US Replacement Parts Database</FONT><BR>"
response.write "<FONT CLASS=Heading4>Equivalent or Replaced By - Search Results</FONT><BR><BR>"
%>

<UL>
<LI>Clicking on either a <FONT COLOR="Red"><B>Replaced By</B></FONT> or a <FONT COLOR="Red"><B>Equivalent</B></FONT> value will requiry the database for updated information.
<LI>Clicking on a <FONT COLOR="Red"><B> <A HREF="codepage.asp" TARGET="codes" onclick="openit('codepage.asp','Vertical');return false;">Code</A></B></FONT> value will display the <FONT COLOR="Red"><B>Parts Code table</B></FONT> in a separate browser window.  When you are done viewing the parts code table, you can close that window.
</UL>

<% Sub WriteDiscountInfo %>

  <% if strDiscount <> 100 or strRate <> 1 then %>
    <UL>
      <LI>All <FONT COLOR="Red"><B>Local Prices</B></FONT> shown are estimates for quoting purposes only and are subject to verification at time of order.</LI>
    <% if strDiscount <> 100 then %>
  	  <LI>A <FONT COLOR="Red"><B>Discount</B></FONT> of <B><%=strDiscount%></B>% from US List Price is reflected in the <FONT COLOR="Red"><B>Local Price</B></FONT> column below.</LI>
    <% end if%>
  	
    <% if strRate <> 1 then %>
  	  <LI>A <FONT COLOR="Red"><B>Local Currency Exchange Rate</B></FONT> of <B><%=strRate%></B> to US Dollars is <%if strDiscount <> 100 then response.write " also "%>relected in the <FONT COLOR="Red"><B>Local Price</B></FONT> column below.</LI>
    <% end if %>
    </UL>
  <% end if %>

<% end sub %>

<% Sub WriteTableHeaders %>
		
	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="Black" BGCOLOR="#666666">
    <TR>
      <TD>
        <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
          <TR>
            <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD>Part<BR>Number</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD>Replaced<BR>By</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD>Equivalent</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD>Code</TD>
            <% if strDiscount <> 100 or strRate <> 1 then %>            
            <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD>Local<BR>Price</TD>
            <% end if %>
            <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD>US List<BR>Price</TD>
            <TD BGCOLOR="#000000" CLASS=SMALLBOLDGOLD>Model Noun</TD>            
            <TD BGCOLOR="#000000" CLASS=SMALLBOLDGOLD>Description</TD>
      		</TR>
    	
<% End Sub %>

<% Sub WriteRecordsToPage %>

		<TR>
			<TD BGCOLOR="#FFFFFF" ALIGN="RIGHT" NOWRAP CLASS=SMALL><%=Session("equivRS")("pfid")%></TD>
			
			<TD ALIGN="RIGHT" BGCOLOR="#FFFFFF" NOWRAP CLASS=SMALL>
				<% IF Session("equivRS")("pt2") <> 0 THEN
            response.write "<A HREF=""equivquery.asp?rt=" & pagenum & "&view=" & view & "&returned=" & limit & "&part=" & Session("equivRS")("PT2") & "&Rate=" & strRate & "&discount=" & strDiscount & """>" & Session("equivRS")("PT2") & "</A>"
  				 ELSE 
			  		response.write "&nbsp;"
				   END IF
        %>
			</TD>
			
			<TD ALIGN="RIGHT" BGCOLOR="#FFFFFF" NOWRAP CLASS=SMALL>
				<% IF Session("equivRS")("ppnp") <> 0 THEN
            response.write "<A HREF=""equivquery.asp?rt=" & pagenum & "&view=" & view & "&returned=" & limit & "&part=" & Session("equivRS")("PPNP") & "&Rate=" & strRate & "&discount=" & strDiscount & """>" & Session("equivRS")("ppnp") & "</A>"
  				 ELSE 
			  		response.write "&nbsp;"
				   END IF
        %>
			</TD>
			
			<TD BGCOLOR="#FFFFFF" NOWRAP CLASS=SMALL>
				<%IF Session("equivRS")("C3") <> "" THEN%>
					&nbsp;&nbsp;&nbsp;&nbsp;<A HREF="codepage.asp" TARGET="codes" onclick="openit('codepage.asp','Vertical');return false;"><%=Session("equivRS")("C3")%></A>&nbsp;&nbsp;<A HREF="codepage.asp" TARGET="codes" onclick="openit('codepage.asp','Vertical');return false;"><%=Session("equivRS")("c4")%></A>						
				<%ELSE%>
					&nbsp;
				<%END IF%>
			</TD>

      <% if strDiscount <> 100 or strRate <> 1 then %>            
			<TD ALIGN="right" BGCOLOR="#FFFFFF" NOWRAP CLASS=SMALL>
				<%IF Session("equivRS")("list_price") <> 0 THEN%>
					<%=FormatNumber(((Session("equivRS")("list_price") / 100) * strRate) * ((100 - strDiscount) / 100)) %>
				<%ELSE%>
					&nbsp;
				<%END IF%>
			</TD>
      <% end if %>
						
			<TD ALIGN="right" BGCOLOR="#FFFFFF" NOWRAP CLASS=SMALL>
				<%IF Session("equivRS")("list_price") <> 0 THEN%>
					<%=FormatNumber(Session("equivRS")("list_price") / 100) %>
				<%ELSE%>
					&nbsp;
				<%END IF%>
			</TD>			
						
			<TD BGCOLOR="#FFFFFF" CLASS=SMALL>
				<%                  
        IF Session("equivRS")("name") <> "" THEN
  			 response.write Session("equivRS")("name")
        ELSE
				  response.write "&nbsp;"
				END IF
        %>
			</TD>

			<TD BGCOLOR="#FFFFFF" CLASS=SMALL>
				<%
        IF Session("equivRS")("short_description") <> "" THEN
          response.write Session("equivRS")("short_description")
				ELSE
				  response.write "&nbsp;"
				END IF
        %>
			</TD>
		</TR>

<% End Sub %>	

<%
Sub writepages
		response.write "<td CLASS=NormalBold>"
		for i = 1 to Session("epartpages")
'			if i=26 then
'				response.write "</td>"
'				exit sub
'			end if
			if i = cint(pagenum) then
				response.write "<a href=""equivresults.asp?view=" & view & "&whatpage=" & i & """><FONT COLOR=""Red"">[" & cstr(i) & "]</FONT></a>  "
			else
				response.write "<a href=""equivresults.asp?view=" & view & "&whatpage=" & i & """>" & cstr(i) & "</a>  "
			end if
		next
		response.write "</td>"
	end sub
%>

<%

'--------------------------subs

'start

Call WriteDiscountInfo

if Session("epartpages") > 1 then
	if request("Whatpage") <> "" then
		pagenum=request("Whatpage")
	else
		pagenum=1
	end if
		
	response.write "<UL>"
  if pagenum=1 then
		response.write "<table><tr>"
		writepages
		response.write "<td CLASS=Normal>"
    response.write "<INPUT TYPE=""Button"" VALUE=""Next Page >>"" CLASS=NavLeftHighLight1 onClick=""location.href='equivresults.asp?view=" & view & "&whatpage=" & pagenum+1 & "'"">"
    response.write "&nbsp;&nbsp;"
    response.write "<INPUT TYPE=""Button"" VALUE=""New Search"" CLASS=NavLeftHighLight1 onClick=""location.href='prtqueryform.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
		response.write "</td></tr></table>"
		Session("equivRS").movefirst
	else
		if Cint(pagenum) = Cint(Session("epartpages")) then
			response.write "<table><tr><td CLASS=Normal>"
      response.write "<INPUT TYPE=""Button"" VALUE=""<< Previous Page"" CLASS=NavLeftHighLight1 onClick=""location.href='equivresults.asp?view=" & view & "&whatpage=" & pagenum-1 & "'"">"
			response.write "</td>"
			writepages
      response.write "<td CLASS=Normal>"
      response.write "<INPUT TYPE=""Button"" VALUE=""New Search"" CLASS=NavLeftHighLight1 onClick=""location.href='prtqueryform.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
			response.write "</td></tr></table>"
		else
			response.write "<table><tr><td CLASS=Normal>"
      response.write "<INPUT TYPE=""Button"" VALUE=""<< Previous Page"" CLASS=NavLeftHighLight1 onClick=""location.href='equivresults.asp?view=" & view & "&whatpage=" & pagenum-1 & "'"">"
			response.write "</td>"
			writepages
			response.write "<td CLASS=Normal>"
      response.write "<INPUT TYPE=""Button"" VALUE=""Next Page >>"" CLASS=NavLeftHighLight1 onClick=""location.href='equivresults.asp?view=" & view & "&whatpage=" & pagenum+1 & "'"">"
      response.write "&nbsp;&nbsp;"
      response.write "<INPUT TYPE=""Button"" VALUE=""New Search"" CLASS=NavLeftHighLight1 onClick=""location.href='prtqueryform.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
			response.write "</td></tr></table>"
		end if
		Session("equivRS").movefirst
		Session("equivRS").move limit*(pagenum-1)
	end if
	response.write "</UL>"
else
	if not Session("equivRS").eof then
		Session("equivRS").movefirst
	end if
  response.write "<UL>"
	response.write "<table><tr><td CLASS=Normal>"  
  response.write "<INPUT TYPE=""Button"" VALUE=""New Search"" CLASS=NavLeftHighLight1 onClick=""location.href='prtqueryform.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
	response.write "</td>"
	response.write "</tr></table>" 
	response.write "</UL>"          
end if

IF Session("equivRS").eof THEN %>

	<UL><LI><FONT COLOR="#FF0000">Sorry, No records match the search criteria you have entered.</FONT></LI></UL>
	<UL>Click on [New Search] to enter new search criteria.</UL>	

<% else
	
	Call WriteTableHeaders	    
	while ((not Session("equivRS").EOF) AND (x <= limit))
		Call WriteRecordsToPage
		Session("equivRS").MoveNext
		x=x+1
	wend
	
	response.write "</TABLE></TD></TR></TABLE>"

end if %>

<%if Session("epartpages") > 1 then

	response.write "<BR>"

	if request("Whatpage") <> "" then
		pagenum=request("Whatpage")
	else
		pagenum=1
	end if
		
	response.write "<UL>"
	if pagenum=1 then
		response.write "<table><tr>"
		writepages
		response.write "<td CLASS=Normal>"
    response.write "<INPUT TYPE=""Button"" VALUE=""Next Page >>"" CLASS=NavLeftHighLight1 onClick=""location.href='equivresults.asp?view=" & view & "&whatpage=" & pagenum+1 & "'"">"
    response.write "&nbsp;&nbsp;"
    response.write "<INPUT TYPE=""Button"" VALUE=""New Search"" CLASS=NavLeftHighLight1 onClick=""location.href='prtqueryform.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
		response.write "</td></tr></table>"
	else
		if Cint(pagenum) = Cint(Session("epartpages")) then
			response.write "<table><tr><td CLASS=Normal>"
      response.write "<INPUT TYPE=""Button"" VALUE=""<< Previous Page"" CLASS=NavLeftHighLight1 onClick=""location.href='equivresults.asp?view=" & view & "&whatpage=" & pagenum-1 & "'"">"
			response.write "</td>"
			writepages
      response.write "<td CLASS=Normal>"
      response.write "<INPUT TYPE=""Button"" VALUE=""New Search"" CLASS=NavLeftHighLight1 onClick=""location.href='prtqueryform.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
			response.write "</td></tr></table>"
		else
			response.write "<table><tr><td CLASS=Normal>"
      response.write "<INPUT TYPE=""Button"" VALUE=""<< Previous Page"" CLASS=NavLeftHighLight1 onClick=""location.href='equivresults.asp?view=" & view & "&whatpage=" & pagenum-1 & "'"">"
			response.write "</td>"
			writepages
			response.write "<td CLASS=Normal>"
      response.write "<INPUT TYPE=""Button"" VALUE=""Next Page >>"" CLASS=NavLeftHighLight1 onClick=""location.href='equivresults.asp?view=" & view & "&whatpage=" & pagenum+1 & "'"">"
      response.write "&nbsp;&nbsp;"
      response.write "<INPUT TYPE=""Button"" VALUE=""New Search"" CLASS=NavLeftHighLight1 onClick=""location.href='prtqueryform.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
			response.write "</td></tr></table>"
		end if
	end if
	response.write "</UL>"
else
  response.write "<UL>"
	response.write "<table><tr><td CLASS=Normal>"  
  response.write "<INPUT TYPE=""Button"" VALUE=""New Search"" CLASS=NavLeftHighLight1 onClick=""location.href='prtqueryform.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
	response.write "</td>"
	response.write "</tr></table>" 
	response.write "</UL>"        
end if %>

 <!--End Content -->

</FONT>
<BR><BR>

<!--#include file="SW-Common-Footer.asp"-->

<SCRIPT LANGUAGE=JAVASCRIPT>

<!--

function Checkitout(){

//      Gets Browser and Version

        var appver = "null";
        var browser = navigator.appName;
        var version = navigator.appVersion;
        if ((browser == "Netscape")) version = navigator.appVersion.substring(0, 3);
        if ((browser == "Microsoft Internet Explorer")) version = navigator.appVersion.substring(22, 25);

//      Gives AppVersion (appver) for Detect Strings

        if ((browser == "Microsoft Internet Explorer") && (version >= 3)) appver = "ie3+";
        if ((browser == "Netscape") && (version >= 3)) appver = "ns3+";
        if ((browser == "Netscape") && (version < 3)) appver = "ns2";


       if ((appver == "ie3+")) {
                return 0;
        }  else {
                return 1;
                }
}

function PopoffWindow(DaURL, orient) {

	var ItsTheWindow;
	if (Checkitout())  {
		if (orient == "Horizontal")  {
			ItsTheWindow = window.open(DaURL,"himom","status,height=400,width=400,scrollbars=yes,resizable=no,toolbar=0");
		} else if (orient == "Vertical")  {
		    ItsTheWindow = window.open(DaURL,"himom","status,height=400,width=400,scrollbars=yes,resizable=no,toolbar=0");
		}


	} else {
		if (orient == "Horizontal")  {
	        ItsTheWindow = window.open(DaURL,"himom","scrollbars=yes,menubar=no,toolbar=no,links=no,status=no,height=400,width=400,resizable=no");
		} else if (orient == "Vertical")  {
	        ItsTheWindow = window.open(DaURL,"himom","scrollbars=yes,menubar=no,toolbar=no,links=no,status=no,height=400,width=400,resizable=no");
		}
			if (parseInt(navigator.appVersion) >= 3){
       		ItsTheWindow.focus();
        }
	}

}
function openit(DaURL, orient) {
		var ItsTheWindow;
        if (Checkitout())  {
                if (orient == "Horizontal")  {
                        ItsTheWindow = window.open(DaURL,"codes","status,height=400,width=600,scrollbars=1,resizable=1,toolbar=0");
                } else if (orient == "Vertical")  {
                    ItsTheWindow = window.open(DaURL,"codes","status,height=580,width=545,scrollbars=1,resizable=1,toolbar=0");
                }


        } else {
                if (orient == "Horizontal")  {
                ItsTheWindow = window.open(DaURL,"codes","scrollbars=1,menubar=0,toolbar=0,links=0,status=1,height=400,width=600,resizable=1");
                } else if (orient == "Vertical")  {
                ItsTheWindow = window.open(DaURL,"codes","scrollbars=1,menubar=0,toolbar=0,links=0,status=1,height=580,width=545,resizable=1");
                }
                        if (parseInt(navigator.appVersion) >= 3){
                ItsTheWindow.focus();
        }
        }
}

//-->

</SCRIPT>
