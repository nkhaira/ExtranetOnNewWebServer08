<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<HTML>
<HEAD>
	<TITLE>Fluke Service</TITLE>
</HEAD>

<!--#include virtual="/service/center/euro_lock_out.asp"-->

<%

Dim TypeOfSearch
Dim strRate
Dim strDiscount
Dim strInvalid
Dim strStrip
Dim pagenum
Dim limit
Dim view
Dim rt
Dim Repair
Dim PartNumber(30)
Dim ServiceOption(30)
Dim PartNumberMax
Dim TempPartNumber
Dim xOption

PartNumber(0) ="1000180"
ServiceOption(0) ="Repair, Standard Price"
PartNumber(1) ="1231580"
ServiceOption(1) ="Performance Test"
PartNumber(2) ="1007020"
ServiceOption(2) ="Calibration, Traceable, no Data"
PartNumber(3) ="1015530"
ServiceOption(3) ="Calibration, Z540 Traceable, no Data"
PartNumber(4) ="1259210"
ServiceOption(4) ="Calibration, Z540 Traceable, no Data"
PartNumber(5) ="1025160"
ServiceOption(5) ="Calibration, Stds Lab, Certified, no Data"
PartNumber(6) ="1230640"
ServiceOption(6) ="Calibration, Artifact"
PartNumber(7) ="1024360"
ServiceOption(7) ="Calibration, Accredited"
PartNumber(8) ="1259800"
ServiceOption(8) ="Calibration, Traceable, no Data"
PartNumber(9) ="1256480"
ServiceOption(9) ="Calibration, Z540 Traceable, no Data"
PartNumber(10)="1258910"
ServiceOption(10)="Calibration, Z540 Traceable, no Data"
PartNumber(11)="1256990"
ServiceOption(11)="Calibration, Accredited"
PartNumber(12)="1024830"
ServiceOption(12)="Agreement, Extended Warranty"
PartNumber(13)="1028820"
ServiceOption(13)="Agreement, Calibration, Traceable, no Data"
PartNumber(14)="1259170"
ServiceOption(14)="Agreement, Calibration, Z540 Traceable, no Data"
PartNumber(15)="1258730"
ServiceOption(15)="Agreement, Calibration, Z540 Traceable, no Data"
PartNumber(16)="1259340"
ServiceOption(16)="Agreement, Calibration, Accredited" 
PartNumber(17)="1257890"
ServiceOption(17)="Agreement, Gold Priority Support"
PartNumber(18)="1540600"
ServiceOption(18)="Agreement, Calibration, Artifact"

' PartNumberMax = last array number
PartNumberMax = 18


'set RS = Session("RS")
strRate = Session("StrRate")
strDiscount = Session("StrDiscount")
limit = Session("Perpage")
view = Session("view")

%>

<% if view <> "1" and view <> "2" then %>
  <BODY BACKGROUND="/service/center/images/bg.gif" BGCOLOR="White" ALINK="#008400" LINK="#008400" VLINK="#008400">
<% else %>
  <BODY BGCOLOR="White" ALINK="#008400" LINK="#008400" VLINK="#008400">
<% end if %>

<% if view <> "1" and view <> "2" then %>

<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH=100%>
  <TR>
    <TD ALIGN="RIGHT"><IMG SRC="/service/center/images/flukelogo.gif" WIDTH=143 HEIGHT=50 ALT="" BORDER="0"></TD></TR></TABLE>
      <BR>
      <TABLE WIDTH=100% BORDER="0" CELLSPACING="0" CELLPADDING="0" ALIGN="LEFT">
        <TR>
          <TD WIDTH=125 ALIGN="LEFT" VALIGN="TOP">
            <TABLE WIDTH=125 BORDER=0 CELLPADDING=0 CELLSPACING=0>

<!-- START LEFT NAVIGATION ROWS -->
<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP>
    <IMG SRC="/service/center/images/spacer.gif" WIDTH=125 HEIGHT=1 ALT="" BORDER="0">
  </TD>
</TR>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
     <A HREF="http://www.fluke.com"><IMG SRC="/service/center/buttons/fluke.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A><BR>
  </TD>
</TR>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
     <A HREF="http://www.fluke.com/service/default.asp"><IMG SRC="/service/center/buttons/service.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A><BR>
  </TD>
</TR>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
    <A HREF="/service/center/default.asp"><IMG SRC="/service/center/buttons/srvcsuph.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A><BR>
  	<IMG SRC="/service/center/images/spacer.gif" WIDTH=1 HEIGHT=15 ALT="" BORDER="0">
  </TD>
</TR>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
    <A HREF="/service/center/documents/whatsnew.asp"><IMG SRC="/service/center/buttons/whatsnew.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
  </TD>
</TR>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
    <A HREF="/service/center/default.asp"><IMG SRC="/service/center/buttons/subservice.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
  </TD>
</TR>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
    <A HREF="/service/center/downloads.asp"><IMG SRC="/service/center/buttons/downloads.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
  </TD>
</TR>
   	   
<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
    <A HREF="/service/center/databases.asp"><IMG SRC="/service/center/buttons/databasesh.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
  </TD>
</TR>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	  <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
    <A HREF="/service/center/documents/svcindex.asp?region=<%=request("region")%>&view=<%=request("view")%>"><IMG SRC="/service/center/buttons/documentation.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
	</TD>
</TR>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
  	<NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
    <A HREF="/service/center/products/svcmodel.asp?region=<%=request("region")%>&view=<%=request("view")%>"><IMG SRC="/service/center/buttons/support.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
	</TD>
</TR>

     <% if Euro_Lock_Out = false or Euro_Lock_Out = 2 then %>
	   <TR>
	     <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	       <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
         <A HREF="/service/center/parts/prtqueryform.asp?region=<%=request("region")%>&view=<%=request("view")%>"><IMG SRC="/service/center/buttons/parts.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
  	   </TD>
	   </TR>
	   <TR>
	     <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	       <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
         <A HREF="/service/center/parts/RnCqueryform.asp?region=<%=request("region")%>&view=<%=request("view")%>"><IMG SRC="/service/center/buttons/usRnCh.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
  	   </TD>
	   </TR>
      <TR>
	   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	      <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
        <A HREF="/service/center/download/metcal/metcal.asp"><IMG SRC="/service/center/buttons/metcal.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
	   </TD>
	   </TR>
     <% end if %>
     
     <% if Euro_Lock_Out = true or Euro_Lock_Out = 2 then %>
      <TR>
	   <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
	      <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=10 HEIGHT=1 ALT="" BORDER="0">
        <A HREF="/service/center/download/calnet/calnet.asp"><IMG SRC="/service/center/buttons/calnet.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A></NOBR>
	   </TD>
	   </TR>
     <% end if %>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
    <A HREF="/service/center/region/default.asp"><IMG SRC="/service/center/buttons/regsup.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
  </TD>
</TR>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
    <A HREF="/service/center/help/helphome.asp"><IMG SRC="/service/center/buttons/help.gif" WIDTH=100 HEIGHT=21 ALT="" BORDER="0"></A>
  </TD>
</TR>

<TR>
  <TD WIDTH=125 COLSPAN=3 VALIGN=TOP HEIGHT=25>
    <NOBR><IMG SRC="/service/center/images/spacer.gif" WIDTH=20 HEIGHT=1 ALT="" BORDER="0"><A HREF="mailto:service@fluke.com"><IMG SRC="/service/center/images/buttons/envelope.gif" BORDER=0 ALT="Send E-Mail to Service Engineering Webmaster"></A></NOBR>
  </TD>
</TR>

</TABLE>
  </TD>
  
  <TD WIDTH=100% ALIGN="left" VALIGN="TOP">

<% end if %>

<!--Content-->

<IMG SRC="/service/center/images/headlines/sv-usRnC-Pricing.gif" ALT="" BORDER="0"><BR>
<IMG SRC="/service/center/images/headlines/sv-query.gif" WIDTH=255 HEIGHT=30 ALT="" BORDER="0">
<BR><BR>

<FONT SIZE=2 FACE="ARIAL, Verdana, Helvetica">

<% Sub WriteDiscountInfo %>

  <% if strDiscount <> 100 or strRate <> 1 then %>
    <UL>
      <LI>All <FONT COLOR="#008400"><B>Local Prices</B></FONT> shown are estimates for quoting purposes only and are subject to verification at time of order.</LI>
    <% if strDiscount <> 100 then %>
  	  <LI>A <FONT COLOR="#008400"><B>Discount</B></FONT> of <B><%=strDiscount%></B>% from US List Price is reflected in the <FONT COLOR="#008400"><B>Local Price</B></FONT> column below.</LI>
    <% end if%>
  	
    <% if strRate <> 1 then %>
  	  <LI>A <FONT COLOR="#008400"><B>Local Currency Exchange Rate</B></FONT> of <B><%=strRate%></B> to US Dollars is <%if strDiscount <> 100 then response.write " also "%>relected in the <FONT COLOR="#008400"><B>Local Price</B></FONT> column below.</LI>
    <% end if %>
    </UL>
  <% end if %>

<% end sub %>

<% Sub WriteTableHeaders %>
		
	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
    <TR>
      <TD>
        <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
          <TR>
            <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50><NOBR><FONT SIZE="1" FACE="Verdana" COLOR="<%=Contrast%>"><B>Model Number<BR>Service Option</B></FONT></TD>
            <TD BGCOLOR="#000000"><FONT SIZE="1" FACE="Verdana" COLOR="#FFCC00"><B>Service Option Description</B></FONT></TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50><FONT SIZE="1" FACE="Verdana" COLOR="<%=Contrast%>"><B>US List<BR>Price</B></FONT></TD>            
      		</TR>
    	
<% End Sub %>

<% Sub WriteServiceOptions

   tempPartNumber = mid(Session("RS")("pfid"),instr(1,Session("RS")("pfid"),"-")+1)
   xOption=0
   While xOption <= PartNumberMax
    if cStr(tempPartNumber) = PartNumber(xOption) then
      response.write ServiceOption(xOption)
      xOption=PartNumberMax
    end if
    xOption=xOption + 1
   Wend
   
  End Sub %>    

<% Sub WriteRecordsToPage_Bundled %>

		<TR>
			<TD BGCOLOR="#FFFFFF" ALIGN="RIGHT"><FONT SIZE=1 FACE="ARIAL, Verdana, Helvetica">
        <% response.write mid(Session("RS")("pfid"), 1, instr(1,Session("RS")("pfid"),"-")-1) & "-" & PartNumber(0) & "<BR>" & Session("RS")("pfid") %>
        </FONT>
      </TD>
			<TD BGCOLOR="#FFFFFF"><FONT SIZE=1 FACE="ARIAL, Verdana, Helvetica">
				<% if Session("RS")("short_description") <> "" then
            response.write "<FONT COLOR=""#FF0000"">Standard Price Repair</FONT> and <BR>"
            call WriteServiceOptions
				   else
					  response.write "&nbsp;"
				   end if
        %>   
        </FONT>
			</TD>
			<TD ALIGN="right" BGCOLOR="#FFFFFF" VALIGN="MIDDLE"><FONT SIZE=1 FACE="ARIAL, Verdana, Helvetica">
        <% ' N/A
           if isnull(Repair) or Repair = "" then
              response.write "N / A"
           ' T & M   
           elseif CDbl(Repair) = 0 then
              if isnull(Session("RS")("list_price")) or Session("RS")("list_price") = "" then
                response.write "N / A"
              elseif CDbl(Session("RS")("list_price")) = 0 then
                response.write "T &amp; M"
              elseif CDbl(Session("RS")("list_price")) <> 0 then                
                response.write "T &amp; M"
              end if
           ' Bundled                   
           elseif CDbl(Repair) <> 0 then              
              if isnull(Session("RS")("list_price")) or Session("RS")("list_price") = "" then
                response.write "N / A"
              elseif CDbl(Session("RS")("list_price")) = 0 then
                response.write "T &amp; M"
              elseif CDbl(Session("RS")("list_price")) <> 0 then                
                response.write FormatNumber((Session("RS")("list_price") / 100) + (Repair / 100))
              end if
           else
              response.write "Error"
           end if   
        %>
        </FONT>
			</TD>
		</TR>

<% End Sub %>	

<% Sub WriteRecordsToPage_AlaCart %>

		<TR>
      <% if instr(1,Session("RS")("short_description"),"Repair") = 1 or Session("RS")("C3") = "99" then
	  		  response.write "<TD BGCOLOR=""Silver"" ALIGN=""RIGHT"" NOWRAP>"
         else
	  		  response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"" NOWRAP>"
         end if
        
         response.write "<FONT SIZE=1 FACE=""ARIAL, Verdana, Helvetica"">" & Session("RS")("pfid")
      %>
        </FONT>
      </TD>

      <% if instr(1,Session("RS")("short_description"),"Repair") = 1 or Session("RS")("C3") = "99" then
	  		  response.write "<TD BGCOLOR=""Silver"">"
         else
	  		  response.write "<TD BGCOLOR=""#FFFFFF"">"
         end if

         response.write "<FONT SIZE=1 FACE=""ARIAL, Verdana, Helvetica"">"

			   if Session("RS")("short_description") <> "" then
           call WriteServiceOptions
           if not isnull(Session("RS")("C3")) or Session("RS")("C3") = "99" then
             response.write " <FONT COLOR=""#FF0000"">(One-Time Calibration)</FONT>"
            end if
          else
				   response.write "&nbsp;"
				  end if
      %>
        </FONT>
			</TD>

      <% if instr(1,Session("RS")("short_description"),"Repair") = 1 or Session("RS")("C3") = "99" then
	  		  response.write "<TD BGCOLOR=""Silver"" ALIGN=""RIGHT"" NOWRAP>"
         else
	  		  response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"" NOWRAP>"
         end if

         response.write "<FONT SIZE=1 FACE=""ARIAL, Verdana, Helvetica"">"

         if isnull(Session("RS")("list_price")) or (Session("RS")("list_price")) = "" then
            response.write "N / A"
				 elseif cdbl(Session("RS")("list_price")) = 0 Then
              response.write "T &amp; M"
          ElseIf cdbl(Session("RS")("list_price")) <> 0 THEN
  					  response.write FormatNumber(Session("RS")("list_price") / 100)
				  Else
            response.write "&nbsp;"
				  END IF
        %>
        </FONT>
			</TD>			
		</TR>

<% End Sub %>	

<%
Sub writepages
		response.write "<td><font face=""ARIAL, HELVETICA, TIMES ROMAN"" size=""2""><b>"
		for i = 1 to Session("partpages")
'			if i=26 then
'				response.write "</b></td>"
'				exit sub
'			end if
			if i = cint(pagenum) then
				response.write "<a href=""rncResults.asp?view=" & view & "&whatpage=" & i & """><FONT COLOR=""black"">[" & cstr(i) & "]</FONT></a>  "
			else
				response.write "<a href=""rncResults.asp?view=" & view & "&whatpage=" & i & """>" & cstr(i) & "</a>  "
			end if
		next
		response.write "</b></td>"
	end sub
%>

<%

'--------------------------subs

'start

'Call WriteDiscountInfo

if Session("partpages") > 1 then
	if request("Whatpage") <> "" then
		pagenum=request("Whatpage")
	else
		pagenum=1
	end if
		
	response.write "<UL>"
  if pagenum=1 then
		response.write "<table><tr>"
		writepages
		response.write "<td>"
		response.write "<a href=""rncResults.asp?view=" & view & "&whatpage=" & pagenum+1 & """><img src=""/service/center/images/buttons/next-button.gif"" width=90 height=21 border=""0""></a>&nbsp;&nbsp;<a href=""/service/center/parts/rncQueryForm.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit &  """><img src=""/service/center/images/buttons/new-search-button.gif"" width=90 height=21 border=""0""></a>"
		response.write "</td></tr></table>"
		Session("RS").movefirst
	else
		if Cint(pagenum) = Cint(Session("partpages")) then
			response.write "<table><tr><td>"
			response.write "<a href=""rncResults.asp?view=" & view & "&whatpage=" & pagenum-1 & """><img src=""/service/center/images/buttons/previous-button.gif"" width=90 height=21 border=""0""></a>"
			response.write "</td>"
			writepages
      response.write "<td>"
      response.write "<a href=""/service/center/parts/rncQueryForm.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit &  """><img src=""/service/center/images/buttons/new-search-button.gif"" width=90 height=21 border=""0""></a>"
			response.write "</td></tr></table>"
		else
			response.write "<table><tr><td>"
			response.write "<a href=""rncResults.asp?view=" & view & "&whatpage=" & pagenum-1 & """><img src=""/service/center/images/buttons/previous-button.gif"" width=90 height=21 border=""0""></a>" 
			response.write "</td>"
			writepages
			response.write "<td>"
			response.write "<a href=""rncResults.asp?view=" & view & "&whatpage=" & pagenum+1 & """><img src=""/service/center/images/buttons/next-button.gif"" width=90 height=21 border=""0""></a>&nbsp;&nbsp;<a href=""/service/center/parts/rncQueryForm.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit &  """><img src=""/service/center/images/buttons/new-search-button.gif"" width=90 height=21 border=""0""></a>"
			response.write "</td></tr></table>"
		end if
		Session("RS").movefirst
		Session("RS").move limit*(pagenum-1)
	end if
	response.write "</UL>"
else
	if not session("RS").eof then
		session("RS").movefirst
	end if
  response.write "<UL>"
	response.write "<table><tr><td>"  
  response.write "<a href=""rncQueryForm.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit &  """><img src=""/service/center/images/buttons/new-search-button.gif"" width=90 height=21 border=""0""></a>"

	response.write "</td>"
	response.write "</tr></table>" 
	response.write "</UL>"          
end if

IF Session("RS").eof THEN %>

	<UL><LI><FONT COLOR="#FF0000">Sorry, No Service Options match the search criteria you have entered.</FONT></LI></UL>
	<UL>Click on [New Search] to enter new search criteria.</UL>	

<% else

  Model_Noun = Replace(Session("RS")("name")," ,",", ")
  
  response.write "<UL>"
  response.write "<LI>Click on the <B>[Back]</B> button of your browser to return to the previous model list or click on <B>[New Search]</B> to enter new search criteria.</LI><BR><BR>"
  response.write "<LI><B>Fluke Model Name&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: <FONT COLOR=""#FF0000"">" &  Model_Noun & "</FONT></B></LI>"
  response.write "<LI><B>Oracle Item Number : <FONT COLOR=""#FF0000"">" &  mid(Session("RS")("pfid"),1,instr(1,Session("RS")("pfid"),"-")-1) & "</FONT></B></LI>"
  response.write "</UL>"
	
  Response.write "<B>Standard Price Repair bundled with a Performance Check or Calibration Service Option</B></BR><BR>"

  ' Determine Standard Repair Price for Bundling
  
	Session("RS").MoveFirst
    
  While not Session("RS").EOF
    if instr(1,Session("RS")("short_description"),"Repair") = 1 then     
      Repair = Session("RS")("list_price")
      Session("RS").MoveLast
    end if
		Session("RS").MoveNext    
  wend       

  ' Bundled Service Options
     
	Session("RS").MoveFirst
  x=0  
	Call WriteTableHeaders	    
	while ((not Session("RS").EOF) AND (x <= limit))
    if instr(1,Session("RS")("short_description"),"Performance") = 1 or instr(1,Session("RS")("short_description"),"Calibration") = 1 then
      if isnull( Session("RS")("C3"))  or Session("RS")("C3") = "" then
    		Call WriteRecordsToPage_Bundled
      end if
    end if      
	  	Session("RS").MoveNext
		  x=x+1
	wend
	
	response.write "</TABLE></TD></TR></TABLE>"

  ' Ala Carte Service Options
  
  Response.write "<BR><B>Al&aacute; Carte Service Options</B> (<FONT COLOR=""Gray"">Gray</FONT> Service Options are not to be quoted Al&aacute; Carte, but provided <U>only</U> as reference information.)<BR><BR>"

  Session("RS").MoveFirst  
  x=0
	Call WriteTableHeaders	    
	while ((not Session("RS").EOF) AND (x <= limit))
		Call WriteRecordsToPage_AlaCart
		Session("RS").MoveNext
		x=x+1
	wend
	
	response.write "</TABLE></TD></TR></TABLE>"

end if %>

<%if Session("partpages") > 1 then

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
		response.write "<td>"
		response.write "<a href=""rncResults.asp?view=" & view & "&whatpage=" & pagenum+1 & """><img src=""/service/center/images/buttons/next-button.gif"" width=90 height=21 border=""0""></a>&nbsp;&nbsp;<a href=""/service/center/parts/rncQueryForm.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit &  """><img src=""/service/center/images/buttons/new-search-button.gif"" width=90 height=21 border=""0""></a>"
		response.write "</td></tr></table>"
	else
		if Cint(pagenum) = Cint(Session("partpages")) then
			response.write "<table><tr><td>"
			response.write "<a href=""rncResults.asp?view=" & view & "&whatpage=" & pagenum-1 & """><img src=""/service/center/images/buttons/previous-button.gif"" width=90 height=21 border=""0""></a>"
			response.write "</td>"
			writepages
      response.write "<td>"
      response.write "<a href=""/service/center/parts/rncQueryForm.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit &  """><img src=""/service/center/images/buttons/new-search-button.gif"" width=90 height=21 border=""0""></a>"
			response.write "</td></tr></table>"
		else
			response.write "<table><tr><td>"
			response.write "<a href=""rncResults.asp?view=" & view & "&whatpage=" & pagenum-1 & """><img src=""/service/center/images/buttons/previous-button.gif"" width=90 height=21 border=""0""></a>" 
			response.write "</td>"
			writepages
			response.write "<td>"
			response.write "<a href=""rncResults.asp?view=" & view & "&whatpage=" & pagenum+1 & """><img src=""/service/center/images/buttons/next-button.gif"" width=90 height=21 border=""0""></a>&nbsp;&nbsp;<a href=""/service/center/parts/rncQueryForm.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit &  """><img src=""/service/center/images/buttons/new-search-button.gif"" width=90 height=21 border=""0""></a>"
			response.write "</td></tr></table>"
		end if
	end if
	response.write "</UL>"
else
  response.write "<UL>"
	response.write "<table><tr><td>"  
  response.write "<a href=""rncQueryForm.asp?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit &  """><img src=""/service/center/images/buttons/new-search-button.gif"" width=90 height=21 border=""0""></a>"
	response.write "</td>"
	response.write "</tr></table>" 
	response.write "</UL>"        
end if %>

 <!--End Content -->

</FONT>
<BR><BR>
<!--#include virtual="/service/center/footer.asp"-->

<% if view <> "1" and view <> "2" then %>

    </TD>
  </TR>
</TABLE>

<% end if %>

</BODY>
</HTML>

