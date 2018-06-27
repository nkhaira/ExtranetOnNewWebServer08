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
<%

Call Connect_SiteWide

%>
<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->
<%

response.buffer = true

Dim BackURL
BackURL = Session("BackURL")

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

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
limit       = Session("ePerpage")
view        = Session("view")
part        = Session("epart")
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

response.write "<FONT CLASS=Heading4>" & Translate("Equivalent or Replaced By - Search Results",Login_Language,conn) & "</FONT><BR><BR>"

response.write "<FONT CLASS=Medium>"

if Session("EquivRS").eof then

  response.write "<UL>"
  response.write "<LI><FONT COLOR=""#FF0000"">" & Translate("Sorry, no records match the search criteria you have entered.",Login_Language,conn) & "</FONT></LI>"
  response.write "</UL>"
	response.write "<UL>"
  response.write Translate("Click on [New Search] to enter new search criteria.",Login_Language,conn)
  response.write "</UL>"

else

  if dept_id = 1 then
    response.write "<UL>"
    response.write "<LI>" & Translate("Clicking on either a <B>Replaced By</B> or a <B>Equivalent</B> value will requery the database for updated information.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Clicking on a",Login_Language,conn) & " <B><A HREF=""Part_CodePage.asp"" TARGET=""codes"" onclick=""openit('Part_CodePage.asp','Vertical');return false;"">" & Translate("Code",Login_language,conn) & "</A></B> " & Translate("value will display the <B>Parts Code Table</B> in a separate browser window.  When you are done viewing the Parts Code Table, you can close that window.",Login_Language,conn) & "</LI>"
    response.write "<LI>" & Translate("Use the <FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;[Back]&nbsp;</FONT> button of your browser to return to the previous query results.",Login_Language,conn) & "</LI>"
    response.write "</UL>"
  end if

  Call WriteDiscountInfo
  Call WriteNavigation
	Call WriteTableHeaders
  
	while ((not Session("EquivRS").EOF) AND (x <= limit))
		Call WriteRecordsToPage
		Session("EquivRS").MoveNext
		x = x + 1
	wend
	
	'response.write "</TABLE></TD></TR></TABLE>"
	response.write "</TABLE>"  
  Call Nav_Border_End
  

end if

Call WriteNavigation

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<!--#include virtual="/include/Pop-UP.asp" -->
<%

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' Subroutines
' --------------------------------------------------------------------------------------

Sub WriteNavigation

  if Session("epartpages") > 1 then
  	if request("Whatpage") <> "" then
  		pagenum=request("Whatpage")
  	else
  		pagenum=1
  	end if
  		
    if pagenum=1 then
      response.write "<BR>"    
      Call Nav_Border_Begin
  		response.write "<TABLE><TR>"
  		Call WritePages
  		response.write "<TD CLASS=Normal>"
      response.write "&nbsp;<A HREF=""Part2_Results.asp?view=" & view & "&whatpage=" & pagenum+1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>"
      response.write "&nbsp;&nbsp;"
      response.write "<A HREF="""
      if dept_id = 20 then
        response.write "Model_QueryForm.asp"
      else
        response.write "Part_QueryForm.asp"
      end if
      response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
  		response.write "</TD></TR></TABLE>"
      Call Nav_Border_End
      response.write "<BR>"    

  		Session("EquivRS").movefirst

  	else
  		if Cint(pagenum) = Cint(Session("partpages")) then
        response.write "<BR>"
        Call Nav_Border_Begin
  			response.write "<TABLE><TR><TD CLASS=NORMAL>"
        response.write "<A HREF=""Part2_Results.asp?view=" & view & "&whatpage=" & pagenum-1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>"
        response.write "&nbsp;&nbsp;"
  			response.write "</TD>"
  			Call WritePages
        response.write "<TD CLASS=Normal>"
        response.write "<A HREF="""
        if dept_id = 20 then
          response.write "Model_QueryForm.asp"
        else
          response.write "Part_QueryForm.asp"
        end if
        response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"      
  			response.write "</TD></TR></TABLE>"
        Call Nav_Border_End
        response.write "<BR>"        
  		else
        response.write "<BR>"
        Call Nav_Border_Begin
  			response.write "<TABLE><TR><TD CLASS=Normal>"
        response.write "<A HREF=""Part2_Results.asp?view=" & view & "&whatpage=" & pagenum-1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>"      
        response.write "&nbsp;&nbsp;"
  			response.write "</TD>"
  			Call WritePages
  			response.write "<TD CLASS=Normal>"
        response.write "&nbsp;<A HREF=""Part2_Results.asp?view=" & view & "&whatpage=" & pagenum+1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>"      
        response.write "&nbsp;&nbsp;"
        response.write "<A HREF="""
        if dept_id = 20 then
          response.write "Model_QueryForm.asp"
        else
          response.write "Part_QueryForm.asp"
        end if
        response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
  			response.write "</TD></TR></TABLE>"
        Call Nav_Border_End
        response.write "<BR>"        
  		end if
  		Session("EquivRS").movefirst
  		Session("EquivRS").move limit*(pagenum-1)
  	end if

  else
  	if not session("EquivRS").eof then
  		session("EquivRS").movefirst
  	end if

    response.write "<BR>"        
    Call Nav_Border_Begin
  	response.write "<TABLE><TR><TD CLASS=Normal>"
    response.write "<A HREF="""  
    if dept_id = 20 then
      response.write "Model_QueryForm.asp"
    else
      response.write "Part_QueryForm.asp"
    end if
    response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
  	response.write "</TD>"
  	response.write "</TR></TABLE>" 
    Call Nav_Border_End
    response.write "<BR>"
  end if

end sub

Sub WriteDiscountInfo

  if strDiscount <> 100 or strRate <> 1 then

    response.write "<UL>"
    response.write "<LI>" & Translate("All Local Prices shown are estimates based on your discount and exchange rate criteria that you entered. Actual price is subject to verification at time of order placement with Fluke.",Login_Language,conn) & "</LI>"
    if strDiscount <> 100 then
  	  response.write "<LI>" & Translate("A Discount of",Login_Language,conn) & " " & strDiscount & "% " & Translate("from US List Price is reflected in the Local Price column below.",Login_Language,conn) & "</LI>"
    end if
  	
    if strRate <> 1 and strDiscount = 100 then
    
  	  response.write "<LI>" & Translate("A Local Currency Exchange Rate of",Login_Language,conn) & " " & strRate & " " & Translate("to US Dollars is",Login_Language,conn) & " "
      response.write Translate("reflected in the Local Price column below.",Login_Language,conn) & "</LI>"
    elseif strRate <> 1 and strDiscount <> 100 then  
  	  response.write "<LI>" & Translate("A Local Currency Exchange Rate of",Login_Language,conn) & " " & strRate & " " & Translate("to US Dollars is also reflected in the Local Price column below.",Login_Language,conn) & " "
    end if
    response.write "</UL>"
    
  end if

  response.write "<UL><LI>" & Translate("Search for",Login_Language,conn) & ": <FONT CLASS=SmallRed>" & UCASE(Session("ePart")) & "</FONT></LI></UL>"

end sub

' --------------------------------------------------------------------------------------

Sub WriteTableHeaders

		
'	<TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="Black" BGCOLOR="#666666">
'    <TR>
'      <TD>
Call Nav_Border_Begin
%>
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
    if Session("EquivRS")("name") <> "" then
		  response.write Session("EquivRS")("name")
  	else
		  response.write "&nbsp;"
	  end if
    response.write "</TD>"

    ' PFID
    
    response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=RIGHT NOWRAP CLASS=SMALL>" & Session("EquivRS")("pfid") & "</TD>"
	
    if dept_id = 1 then

  		' Replaced By
    
    	response.write "<TD ALIGN=RIGHT BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
 	  	if Session("EquivRS")("pt2") <> 0 then
        response.write "<A HREF=""Part2_Query.asp?rt=" & pagenum & "&view=" & view & "&limit=" & limit & "&part=" & Session("EquivRS")("PT2") & "&Rate=" & strRate & "&discount=" & strDiscount & """>" & Session("EquivRS")("PT2") & "</A>"
  		else
			 	response.write "&nbsp;"
			end if
      response.write "</TD>"
      
      ' Equivalent
	
			response.write "<TD ALIGN=RIGHT BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
'			if Session("EquivRS")("ppnp") <> 0 then
'        response.write "<A HREF=""Part2_Query.asp?rt=" & pagenum & "&view=" & view & "&limit=" & limit & "&part=" & Session("EquivRS")("PPNP") & "&Rate=" & strRate & "&discount=" & strDiscount & """>" & Session("EquivRS")("ppnp") & "</A>"
'  		else
			  response.write "&nbsp;"
'			end if
      response.write "</TD>"
      
      ' Part Restriction Code
	
   		response.write "<TD BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
			if Session("EquivRS")("C3") <> "" then
        %>
			  &nbsp;&nbsp;&nbsp;&nbsp;<A HREF="Part_CodePage.asp" TARGET="codes" onclick="openit('Part_CodePage.asp','Vertical');return false;"><%=Session("EquivRS")("C3")%></A>&nbsp;&nbsp;<A HREF="Part_CodePage.asp" TARGET="codes" onclick="openit('codepage.asp','Vertical');return false;"><%=session("EquivRS")("c4")%></A>
				<%
      else
				response.write "&nbsp;"
      end if
      response.write "</TD>"
      
    end if
        
    ' US List Price
    			
  	response.write "<TD ALIGN=RIGHT BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
	  if Session("EquivRS")("list_price") <> 0 then
	    response.write FormatNumber(Session("EquivRS")("list_price") / 100)
  	else
	    response.write "&nbsp;"
  	end if
	  response.write "</TD>"

    ' Local Price

    if strDiscount <> 100 or strRate <> 1 then
      response.write "<TD ALIGN=RIGHT BGCOLOR=""#EEEEEE"" NOWRAP CLASS=SMALL>"
		  if Session("EquivRS")("list_price") <> 0 then
			  response.write FormatNumber(((Session("EquivRS")("list_price") / 100) * strRate) * ((100 - strDiscount) / 100))
  		else
	  	  response.write "&nbsp;"
  		end if
      response.write "</TD>"
	  end if
        
    ' Unit of Measure
    
  	response.write "<TD ALIGN=""CENTER"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
		if Session("EquivRS")("unit_of_measure") <> "" THEN
      response.write Session("EquivRS")("unit_of_measure")
 		else
			response.write "EA"
		end if
    response.write "</TD>"

    ' Short Description
    
  	response.write "<TD BGCOLOR=""#FFFFFF"" CLASS=SMALL>"
    if Session("EquivRS")("short_description") <> "" then
      tempStr = UCase(Replace(Session("EquivRS")("short_description")," ,",", "))
      tempStr = Replace(tempStr,",",", ")
      tempStr = Replace(tempStr,"  "," ")
      response.write Replace(tempStr, UCase(Session("Part")), "<FONT COLOR=""Red"">" & UCase(Session("Part")) & "</FONT>")
      tempStr = null
    else
	  response.write "&nbsp;"
  	end if
    response.write "</TD>"
    
    if dept_id = 20 then
    
			' UPC Code
      
      response.write "<TD ALIGN=""CENTER"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
  		if Session("EquivRS")("upc_code") <> "" THEN
        response.write Session("EquivRS")("upc_code")
    	else
  			response.write "&nbsp;"
  		end if
      response.write "</TD>"
    
      ' Weight
           
			response.write "<TD ALIGN=""RIGHT"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
 			if Session("EquivRS")("weight") <> "" and (isnumeric(Session("EquivRS")("weight")) and Session("EquivRS")("weight") <> 0) THEN
        response.write FormatNumberFloat(CStr(Session("EquivRS")("weight")),3)
        if UCase(Session("EquivRS")("weight_code")) = "KILOGRAM" then
          response.write " " & Replace(Session("EquivRS")("weight_code"),"KILOGRAM","kg")
          response.write "<BR>" & FormatNumberFloat(CStr((CDbl(Session("EquivRS")("weight")) / 0.4535924)),3) & "&nbsp;&nbsp;lb"
        elseif UCase(Session("EquivRS")("weight_code")) = "GRAM" then
          response.write " " & Replace(Session("EquivRS")("weight_code"),"GRAM"," g")        
          response.write "<BR>" & FormatNumberFloat(CStr((CDbl(Session("EquivRS")("weight")) * 0.035)),3) & "&nbsp;&nbsp;oz"
        elseif Session("EquivRS")("weight_code") <> "" then
          response.write " " & Session("EquivRS")("weight_code")
        end if  
  		else
				response.write "&nbsp;<BR>&nbsp;"
 			end if
      response.write "</TD>"

      ' CE Status
      
			response.write "<TD ALIGN=""CENTER"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
 			if Session("EquivRS")("ce_status") <> "" THEN
        response.write Session("EquivRS")("ce_status")
   		else
				response.write "&nbsp;"
 			end if
      response.write "</TD>"
      
			' Country of Origin
      
      response.write "<TD ALIGN=""CENTER"" BGCOLOR=""#FFFFFF"" NOWRAP CLASS=SMALL>"
 			if Session("EquivRS")("origin") <> "" THEN
        response.write Session("EquivRS")("origin")
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
				response.write "<A HREF=""Part2_Results.asp?view=" & view & "&whatpage=" & i & """><FONT CLASS=NavLeftHighLight1>&nbsp;" & cstr(i) & "&nbsp;</FONT></A>&nbsp;"
			else
				response.write "<A HREF=""Part2_Results.asp?view=" & view & "&whatpage=" & i & """><FONT CLASS=NavTopHighLight>&nbsp;" & cstr(i) & "&nbsp;</FONT></A>&nbsp;"
			end if
		next
		response.write "</TD>"
	end sub

' --------------------------------------------------------------------------------------  
%>

<!--#include virtual="/include/Pop-Up.asp" -->

