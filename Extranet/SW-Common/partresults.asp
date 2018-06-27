<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
' --------------------------------------------------------------------------------------
%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connections_parts.asp" -->
<%

Call Connect_SiteWide

response.buffer = true

if request("Site_ID") <> "" and isnumeric(request("Site_ID")) then
  Site_ID            = request("Site_ID")
  Session("Site_ID") = request("Site_ID")  
elseif session("Site_ID") <> "" and isnumeric(session("Site_ID")) then
  Site_ID = session("Site_ID")
else
  Session.Abandon
  Session("ErrorString") = "<LI>Your session has expired.  For your protection, you have been automatically logged off of your extranet site account.</LI><LI>To establish another session, please type in the site's URL in your browser's address line, then re-enter your User Name and Password, or</LI><LI>Use the Site Search feature below.</LI>"
  response.redirect "http://" & Request("SERVER_NAME") & "/register/default.asp?ErrorString=" & Session("ErrorString")
end if

Dim BackURL
BackURL = Session("BackURL")
   

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code = rsSite("Site_Code")
Site_Description = rsSite("Site_Description")
  
rsSite.close
set rsSite=nothing

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

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
Dim dept_id

strRate     = Session("StrRate")
strDiscount = Session("StrDiscount")
limit       = Session("Perpage")
view        = Session("view")
part        = Session("part")
dept_id     = CInt(Session("dept_id"))

if dept_id = 1 then
  Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("US Replacement Parts Database",Alt_Language,conn)
  Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("US Replacement Parts Database",Login_Language,conn) & "</FONT>"
else  
  Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("US Mainframe / Option / Accessory Database",Alt_Language,conn)
  Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("US Mainframe / Option / Accessory Database",Login_Language,conn) & "</FONT>"
end if 

Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

if dept_id = 1 then
  response.write "<FONT CLASS=Heading3>" & Translate("US Replacement Parts Database",Login_Language,conn) & "</FONT><BR>"
else
  response.write "<FONT CLASS=Heading3>" & Translate("US Mainframe / Option / Accessory Database",Login_Language,conn) & "</FONT><BR>"
end if    

response.write "<FONT CLASS=Heading4>" & Translate("Search Results",Login_Language,conn) & "</FONT><BR><BR>"

response.write "<FONT CLASS=Medium>"

if dept_id = 1 then
  %>
  <UL>
  <LI><%response.write Translate("Clicking on either a <B>Replaced By</B> or a <B>Equivalent</B> value will requery the database for updated information.",Login_Language,conn)%></LI>
  <LI><%response.write Translate("Clicking on a",Login_Language,conn)%><B> <A HREF="codepage.asp" TARGET="codes" onclick="openit('codepage.asp','Vertical');return false;">Code</A></B> <%response.write Translate("value will display the <B>Parts Code Table</B> in a separate browser window.  When you are done viewing the Parts Code Table, you can close that window.",Login_Language,conn)%>
  </UL>
<%
end if

Call WriteDiscountInfo

if Session("partpages") > 1 then
	if request("Whatpage") <> "" then
		pagenum=request("Whatpage")
	else
		pagenum=1
	end if
		
	response.write "<UL>"
  if pagenum=1 then
		response.write "<TABLE><TR>"
		Call WritePages
		response.write "<TD CLASS=Normal>"
    response.write "&nbsp;<A HREF=""PartResults.asp?view=" & view & "&whatpage=" & pagenum+1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>"
'    response.write "<INPUT TYPE=""Button"" VALUE="" >> "" CLASS=NavLeftHighLight1 onClick=""location.href='partresults.asp?view=" & view & "&whatpage=" & pagenum+1 & "'"">"
    response.write "&nbsp;&nbsp;"
    response.write "<A HREF="""
'    response.write "<INPUT TYPE=""Button"" VALUE=""" & Translate("New Search",Login_Language,conn) & """ CLASS=NavLeftHighLight1 onClick=""location.href='"
    if dept_id = 20 then
      response.write "MdlQueryForm.asp"
    else
      response.write "PrtQueryForm.asp"
    end if
    response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
		response.write "</TD></TR></TABLE>"
		Session("rs").movefirst
	else
		if Cint(pagenum) = Cint(Session("partpages")) then
			response.write "<TABLE><TR><TD CLASS=NORMAL>"
      response.write "<A HREF=""PartResults.asp?view=" & view & "&whatpage=" & pagenum-1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>"
      response.write "&nbsp;&nbsp;"
			response.write "</TD>"
			Call WritePages
      response.write "<TD CLASS=Normal>"
      response.write "<A HREF="""
'      response.write "<INPUT TYPE=""Button"" VALUE=""" & Translate("New Search",Login_Language,conn) & """ CLASS=NavLeftHighLight1 onClick=""location.href='"
      if dept_id = 20 then
        response.write "MdlQueryForm.asp"
      else
        response.write "PrtQueryForm.asp"
      end if
      response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"      
'      response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
			response.write "</TD></TR></TABLE>"
		else
			response.write "<TABLE><TR><TD CLASS=Normal>"
      response.write "<A HREF=""PartResults.asp?view=" & view & "&whatpage=" & pagenum-1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>"      
'      response.write "<INPUT TYPE=""Button"" VALUE="" << "" CLASS=NavLeftHighLight1 onClick=""location.href='partresults.asp?view=" & view & "&whatpage=" & pagenum-1 & "'"">"
      response.write "&nbsp;&nbsp;"
			response.write "</TD>"
			Call WritePages
			response.write "<TD CLASS=Normal>"
      response.write "&nbsp;<A HREF=""PartResults.asp?view=" & view & "&whatpage=" & pagenum+1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>"      
'      response.write "<INPUT TYPE=""Button"" VALUE="" >> "" CLASS=NavLeftHighLight1 onClick=""location.href='partresults.asp?view=" & view & "&whatpage=" & pagenum+1 & "'"">"
      response.write "&nbsp;&nbsp;"
      response.write "<A HREF="""
'      response.write "<INPUT TYPE=""Button"" VALUE=""" & Translate("New Search",Login_Language,conn) & """ CLASS=NavLeftHighLight1 onClick=""location.href='"
      if dept_id = 20 then
        response.write "MdlQueryForm.asp"
      else
        response.write "PrtQueryForm.asp"
      end if
      response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
'      response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
			response.write "</TD></TR></TABLE>"
		end if
		Session("rs").movefirst
		Session("rs").move limit*(pagenum-1)
	end if
	response.write "</UL>"
else
	if not session("rs").eof then
		session("rs").movefirst
	end if
  response.write "<UL>"
	response.write "<TABLE><TR><TD CLASS=Normal>"
  response.write "<A HREF="""  
'  response.write "<INPUT TYPE=""Button"" VALUE=""" & Translate("New Search",Login_Language,conn) & """ CLASS=NavLeftHighLight1 onClick=""location.href='"
  if dept_id = 20 then
    response.write "MdlQueryForm.asp"
  else
    response.write "PrtQueryForm.asp"
  end if
  response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
'  response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
	response.write "</TD>"
	response.write "</TR></TABLE>" 
	response.write "</UL>"          
end if

if Session("rs").eof then %>

	<UL><LI><FONT COLOR="#FF0000"><%response.write Translate("Sorry, no records match the search criteria you have entered.",Login_Language,conn)%></FONT></LI></UL>
	<UL><%response.write Translate("Click on [New Search] to enter new search criteria.",Login_Language,conn)%></UL>	

<% else
	
	Call WriteTableHeaders	    
	while ((not Session("rs").EOF) AND (x <= limit))
		Call WriteRecordsToPage
		Session("rs").MoveNext
		x=x+1
	wend
	
	response.write "</TABLE></TD></TR></TABLE>"

end if

if Session("partpages") > 1 then

	response.write "<BR>"

	if request("Whatpage") <> "" then
		pagenum=request("Whatpage")
	else
		pagenum=1
	end if
		
	response.write "<UL>"
	if pagenum=1 then
		response.write "<TABLE><TR>"
		writepages
		response.write "<TD CLASS=Normal>"
    response.write "&nbsp;<A HREF=""PartResults.asp?view=" & view & "&whatpage=" & pagenum+1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>"    
'    response.write "<INPUT TYPE=""Button"" VALUE="" >> "" CLASS=NavLeftHighLight1 onClick=""location.href='partresults.asp?view=" & view & "&whatpage=" & pagenum+1 & "'"">"
    response.write "&nbsp;&nbsp;"
    response.write "<A HREF="""
'    response.write "<INPUT TYPE=""Button"" VALUE=""" & Translate("New Search",Login_Language,conn) & """ CLASS=NavLeftHighLight1 onClick=""location.href='"
    if dept_id = 20 then
      response.write "MdlQueryForm.asp"
    else
      response.write "PrtQueryForm.asp"
    end if
    response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
'    response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"

		response.write "</TD></TR></TABLE>"
	else
		if Cint(pagenum) = Cint(Session("partpages")) then
			response.write "<TABLE><TR><TD CLASS=Normal>"
      response.write "<A HREF=""PartResults.asp?view=" & view & "&whatpage=" & pagenum-1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>"      
'      response.write "<INPUT TYPE=""Button"" VALUE="" << "" CLASS=NavLeftHighLight1 onClick=""location.href='partresults.asp?view=" & view & "&whatpage=" & pagenum-1 & "'"">"
			response.write "</TD>"
			writepages
      response.write "<TD CLASS=Normal>"
      response.write "<A HREF="""
'      response.write "<INPUT TYPE=""Button"" VALUE=""" & Translate("New Search",Login_Language,conn) & """ CLASS=NavLeftHighLight1 onClick=""location.href='"
      if dept_id = 20 then
        response.write "MdlQueryForm.asp"
      else
        response.write "PrtQueryForm.asp"
      end if
      response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
'      response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
			response.write "</TD></TR></TABLE>"
		else
			response.write "<TABLE><TR><TD CLASS=Normal>"
      response.write "<A HREF=""PartResults.asp?view=" & view & "&whatpage=" & pagenum-1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>"      
'      response.write "<INPUT TYPE=""Button"" VALUE="" << "" CLASS=NavLeftHighLight1 onClick=""location.href='partresults.asp?view=" & view & "&whatpage=" & pagenum-1 & "'"">"
      response.write "&nbsp;&nbsp;"
			response.write "</TD>"
			writepages
			response.write "<TD CLASS=Normal>"
      response.write "&nbsp;<A HREF=""PartResults.asp?view=" & view & "&whatpage=" & pagenum+1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>"      
'      response.write "<INPUT TYPE=""Button"" VALUE="" >> "" CLASS=NavLeftHighLight1 onClick=""location.href='partresults.asp?view=" & view & "&whatpage=" & pagenum+1 & "'"">"
      response.write "&nbsp;&nbsp;"
      response.write "<A HREF="""
'      response.write "<INPUT TYPE=""Button"" VALUE=""" & Translate("New Search",Login_Language,conn) & """ CLASS=NavLeftHighLight1 onClick=""location.href='"
      if dept_id = 20 then
        response.write "MdlQueryForm.asp"
      else
        response.write "PrtQueryForm.asp"
      end if
      response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
'      response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"
			response.write "</TD></TR></TABLE>"
		end if
	end if
	response.write "</UL>"
else
  response.write "<UL>"
	response.write "<TABLE><TR><TD CLASS=Normal>"
  response.write "<A HREF="""  
'  response.write "<INPUT TYPE=""Button"" VALUE=""" & Translate("New Search",Login_Language,conn) & """ CLASS=NavLeftHighLight1 onClick=""location.href='"
  if dept_id = 20 then
    response.write "MdlQueryForm.asp"
  else
    response.write "PrtQueryForm.asp"
  end if
  response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
'  response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & "'"">"

	response.write "</TD>"
	response.write "</TR></TABLE>" 
	response.write "</UL>"        
end if %>

<!--#include virtual="SW-Common/SW-Footer.asp"-->

<%

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' Subroutines
' --------------------------------------------------------------------------------------

Sub WriteDiscountInfo

  if strDiscount <> 100 or strRate <> 1 then

    response.write "<UL>"
    response.write "<LI>" & Translate("All Local Prices shown are estimates based on your discount and exchange rate criteria that you entered. Actual price is subject to verification at time of order placement with Fluke.",Login_Language,conn) & "</LI>"
    if strDiscount <> 100 then
  	  response.write "<LI>" & Translate("A Discount of",Login_Language,conn) & " " & strDiscount & " " & Translate("from US List Price is reflected in the Local Price column below.",Login_Language,conn) & "</LI>"
    end if
  	
    if strRate <> 1 and strDiscount = 100 then
    
  	  response.write "<LI>" & Translate("A Local Currency Exchange Rate of",Login_Language,conn) & " " & strRate & " " & Translate("to US Dollars is",Login_Language,conn) & " "
      response.write Translate("reflected in the Local Price column below.",Login_Language,conn) & "</LI>"
    elseif strRate <> 1 and strDiscount <> 100 then  
  	  response.write "<LI>" & Translate("A Local Currency Exchange Rate of",Login_Language,conn) & " " & strRate & " " & Translate("to US Dollars is also reflected in the Local Price column below.",Login_Language,conn) & " "
    end if
    response.wrie "</UL>"
    
  end if

  response.write "<UL><LI>" & Translate("Search for",Login_Language,conn) & ": <FONT CLASS=SmallRed>" & UCASE(Session("Part")) & "</FONT></LI></UL>"

end sub

' --------------------------------------------------------------------------------------

Sub WriteTableHeaders

  %>
		
	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="Black" BGCOLOR="#666666">
    <TR>
      <TD>
        <TABLE CELLPADDING=2 CELLSPACING=1 BORDER=0  WIDTH="100%">
          <TR>
            <TD BGCOLOR="#000000" CLASS=SMALLBOLDGOLD><%response.write Translate("Model Noun",Login_Language,conn)%></TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD><%response.write Translate("Part Number",Login_Language,conn)%></TD>

            <% if dept_id = 1 then %>
              <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD><%response.write Translate("Replaced By",Login_Language,conn)%></TD>
              <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD><%response.write Translate("Equivalent",Login_Language,conn)%></TD>
              <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD><%response.write Translate("Code",Login_Language,conn)%></TD>
            <% end if %>
            
            <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD><%response.write Translate("US List Price",Login_Language,conn)%></TD>
            <% if strDiscount <> 100 or strRate <> 1 then %>           
              <TD BGCOLOR="#000000" ALIGN="CENTER" WIDTH=50 CLASS=SMALLBOLDGOLD><% response.write Translate("Local Price",Login_Language,conn)%></TD>
            <% end if %>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SMALLBOLDGOLD><%response.write Translate("Unit",Login_Language,conn)%></TD>
            <TD BGCOLOR="#000000" CLASS=SMALLBOLDGOLD><%response.write Translate("Description",Login_Language,conn)%></TD>
            <% if dept_id = 20 then %>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SMALLBOLDGOLD><%response.write Translate("UPC Code",Login_Language,conn)%></TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SMALLBOLDGOLD><%response.write Translate("Weight",Login_Language,conn)%></TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SMALLBOLDGOLD>CE</TD>                                          
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SMALLBOLDGOLD><%Response.write Translate("0rigin",Login_Language,conn)%></TD>
            <% end if %>            
            
      		</TR>
    	
  <%
End Sub

' --------------------------------------------------------------------------------------

Sub WriteRecordsToPage

		response.write "<TR>"

    ' Name
    
  	response.write "<TD BGCOLOR=""#FFFFFF"" CLASS=SMALL NOWRAP>"
    if Session("rs")("name") <> "" then
		  response.write Session("rs")("name")
  	else
		  response.write "&nbsp;"
	  end if
    response.write "</TD>"

    ' PFID
    
    response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=RIGHT NOWRAP CLASS=SMALL>" & Session("rs")("pfid") & "</TD>"
	
    if dept_id = 1 then

  		' Replaced By
    
    	response.write "<TD ALIGN=RIGHT BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
 	  	if Session("RS")("pt2") <> 0 then
        response.write "<A HREF=""equivquery.asp?rt=" & pagenum & "&view=" & view & "&returned=" & limit & "&part=" & Session("rs")("PT2") & "&Rate=" & strRate & "&discount=" & strDiscount & """>" & Session("rs")("PT2") & "</A>"
  		else
			 	response.write "&nbsp;"
			end if
      response.write "</TD>"
      
      ' Equivalent
	
			response.write "<TD ALIGN=RIGHT BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
			if Session("rs")("ppnp") <> 0 then
        response.write "<A HREF=""equivquery.asp?rt=" & pagenum & "&view=" & view & "&returned=" & limit & "&part=" & Session("rs")("PPNP") & "&Rate=" & strRate & "&discount=" & strDiscount & """>" & Session("rs")("ppnp") & "</A>"
  		else
			  response.write "&nbsp;"
			end if
      response.write "</TD>"
      
      ' Part Restriction Code
	
   		response.write "<TD BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
			if Session("rs")("C3") <> "" then
        %>
			  &nbsp;&nbsp;&nbsp;&nbsp;<A HREF="codepage.asp" TARGET="codes" onclick="openit('codepage.asp','Vertical');return false;"><%=Session("rs")("C3")%></A>&nbsp;&nbsp;<A HREF="codepage.asp" TARGET="codes" onclick="openit('codepage.asp','Vertical');return false;"><%=session("rs")("c4")%></A>						
				<%
      else
				response.write "&nbsp;"
      end if
      response.write "</TD>"
      
    end if
        
    ' US List Price
    			
  	response.write "<TD ALIGN=RIGHT BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
	  if Session("rs")("list_price") <> 0 then
	    response.write FormatNumber(Session("rs")("list_price") / 100)
  	else
	    response.write "&nbsp;"
  	end if
	  response.write "</TD>"

    ' Local Price

    if strDiscount <> 100 or strRate <> 1 then
      response.write "<TD ALIGN=RIGHT BGCOLOR=""#EEEEEE"" NOWRAP CLASS=SMALL>"
		  if Session("rs")("list_price") <> 0 then
			  response.write FormatNumber(((Session("rs")("list_price") / 100) * strRate) * ((100 - strDiscount) / 100))
  		else
	  	  response.write "&nbsp;"
  		end if
      response.write "</TD>"
	  end if
        
    ' Unit of Measure
    
  	response.write "<TD ALIGN=""CENTER"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
		if Session("rs")("unit_of_measure") <> "" THEN
      response.write Session("rs")("unit_of_measure")
 		else
			response.write "EA"
		end if
    response.write "</TD>"

    ' Short Description
    
  	response.write "<TD BGCOLOR=""#FFFFFF"" CLASS=SMALL>"
    if Session("rs")("short_description") <> "" then
      tempStr = UCase(Replace(Session("rs")("short_description")," ,",", "))
      response.write Replace(tempStr, UCase(Session("Part")), "<FONT COLOR=""Red"">" & UCase(Session("Part")) & "</FONT>")
      tempStr = null
    else
	  response.write "&nbsp;"
  	end if
    response.write "</TD>"
    
    if dept_id = 20 then
    
			' UPC Code
      
      response.write "<TD ALIGN=""CENTER"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
  		if Session("rs")("upc_code") <> "" THEN
        response.write Session("rs")("upc_code")
    	else
  			response.write "&nbsp;"
  		end if
      response.write "</TD>"
    
      ' Weight
           
			response.write "<TD ALIGN=""RIGHT"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
 			if Session("rs")("weight") <> "" and (isnumeric(Session("rs")("weight")) and Session("rs")("weight") <> 0) THEN
        response.write FormatNumberFloat(CStr(Session("rs")("weight")),3)
        if UCase(Session("rs")("weight_code")) = "KILOGRAM" then
          response.write " " & Replace(Session("rs")("weight_code"),"KILOGRAM","kg")
          response.write "<BR>" & FormatNumberFloat(CStr((CDbl(Session("rs")("weight")) / 0.4535924)),3) & "&nbsp;&nbsp;lb"
        elseif UCase(Session("rs")("weight_code")) = "GRAM" then
          response.write " " & Replace(Session("rs")("weight_code"),"GRAM"," g")        
          response.write "<BR>" & FormatNumberFloat(CStr((CDbl(Session("rs")("weight")) * 0.035)),3) & "&nbsp;&nbsp;oz"
        elseif Session("rs")("weight_code") <> "" then
          response.write " " & Session("rs")("weight_code")
        end if  
  		else
				response.write "&nbsp;<BR>&nbsp;"
 			end if
      response.write "</TD>"

      ' CE Status
      
			response.write "<TD ALIGN=""CENTER"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
 			if Session("rs")("ce_status") <> "" THEN
        response.write Session("rs")("ce_status")
   		else
				response.write "&nbsp;"
 			end if
      response.write "</TD>"
      
			' Country of Origin
      
      response.write "<TD ALIGN=""CENTER"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
 			if Session("rs")("origin") <> "" THEN
        response.write Session("rs")("origin")
   		else
			response.write "&nbsp;"
 			end if
      response.write "</TD>"
          
    end if   

  response.write "</TR>"

End Sub

' --------------------------------------------------------------------------------------

Sub writepages

		response.write "<TD CLASS=NORMAL>"

		for i = 1 to Session("partpages")
'			if i=26 then
'				response.write "</TD>"
'				exit sub
'			end if
			if i = cint(pagenum) then
				response.write "<A HREF=""partresults.asp?view=" & view & "&whatpage=" & i & """><FONT CLASS=NavLeftHighLight1>&nbsp;" & cstr(i) & "&nbsp;</FONT></A>&nbsp;"
			else
				response.write "<A HREF=""partresults.asp?view=" & view & "&whatpage=" & i & """><FONT CLASS=NavTopHighLight>&nbsp;" & cstr(i) & "&nbsp;</FONT></A>&nbsp;"
			end if
		next
		response.write "</TD>"
	end sub

' --------------------------------------------------------------------------------------  
%>

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
