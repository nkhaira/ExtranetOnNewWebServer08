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

Set Session("EquivRS") = nothing          ' Reset this since you are not in Part2_Query.asp and there is no need to keep the old recordset.

Call Connect_SiteWide

%>
<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->
<%

'response.buffer = true

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
Dim ErrorString
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
ServiceOption(0) = Translate("Repair, Standard Price",Login_Language,conn)
PartNumber(1) ="1231580"
ServiceOption(1) = Translate("Performance Test",Login_Language,conn)

PartNumber(2) ="2132535"
ServiceOption(2) = Translate("Calibration, Traceable, with Data",Login_Language,conn)

PartNumber(3) ="1007020"
ServiceOption(3) = Translate("Calibration, Traceable, no Data",Login_Language,conn)
PartNumber(4) ="1015530"
ServiceOption(4) = Translate("Calibration, Z540 Traceable, with Data",Login_Language,conn)
PartNumber(5) ="1259210"
ServiceOption(5) = Translate("Calibration, Z540 Traceable, no Data",Login_Language,conn)
PartNumber(6) ="1025160"
ServiceOption(6) = Translate("Calibration, Stds Lab, Certified, with Data",Login_Language,conn)
PartNumber(7) ="1230640"
ServiceOption(7) = Translate("Calibration, Artifact",Login_Language,conn)
PartNumber(8) ="1024360"
ServiceOption(8) = Translate("Calibration, Accredited",Login_Language,conn)

PartNumber(9) ="2132558"
ServiceOption(9) = Translate("Calibration, Traceable, with Data",Login_Language,conn)

PartNumber(10) ="1259800"
ServiceOption(10) = Translate("Calibration, Traceable, no Data",Login_Language,conn)
PartNumber(11) ="1256480"
ServiceOption(11) = Translate("Calibration, Z540 Traceable, with Data",Login_Language,conn)
PartNumber(12)="1258910"
ServiceOption(12)= Translate("Calibration, Z540 Traceable, no Data",Login_Language,conn)
PartNumber(13)="1256990"
ServiceOption(13)= Translate("Calibration, Accredited",Login_Language,conn)
PartNumber(14)="1024830"
ServiceOption(14)= Translate("Agreement, Extended Warranty",Login_Language,conn)
PartNumber(15)="1028820"
ServiceOption(15)=Translate("Agreement, Calibration, Traceable, no Data",Login_Language,conn)
PartNumber(16)="1259170"
ServiceOption(16)=Translate("Agreement, Calibration, Z540 Traceable, with Data",Login_Language,conn)
PartNumber(17)="1258730"
ServiceOption(17)=Translate("Agreement, Calibration, Z540 Traceable, no Data",Login_Language,conn)
PartNumber(18)="1259340"
ServiceOption(18)=Translate("Agreement, Calibration, Accredited",Login_Language,conn)
PartNumber(19)="1257890"
ServiceOption(19)=Translate("Agreement, Gold Priority Support",Login_Language,conn)
PartNumber(20)="1540600"
ServiceOption(20)=Translate("Agreement, Calibration, Artifact",Login_Language,conn)
PartNumber(21)="1663069"
ServiceOption(21)=Translate("Agreement, Z540 Care 3 Year, Z540 w/ Data Calibration and Repair",Login_Language,conn)
PartNumber(22)="1663078"
ServiceOption(22)=Translate("Agreement, Z540 Care 5 Year, Z540 w/ Data Calibration and Repair",Login_Language,conn)

PartNumber(23)="1613014"
ServiceOption(23)=Translate("Agreement, Fluke Care 5 Year Accredited Calibration and Repair",Login_Language,conn)
PartNumber(24)="1613006"
ServiceOption(24)=Translate("Agreement, Fluke Care 3 Year Accredited Calibration and Repair",Login_Language,conn)

' PartNumberMax = last array number

PartNumberMax = 24

strRate = Session("StrRate")
strDiscount = Session("StrDiscount")
limit = Session("Perpage")
view = Session("view")

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("US Repair and Calibration Service Options Database",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("US Repair and Calibration Service Options Database",Login_Language,conn) & "</FONT>"

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

response.write "<FONT CLASS=Heading3>" & Translate("US Repair and Calibration Service Options Database",Login_Language,conn) & "</FONT><BR>"
response.write "<FONT CLASS=Heading4>" & Translate("Search Results",Login_Language,conn) & "</FONT><BR><BR>"

response.write "<FONT CLASS=Medium>"

'Call WriteDiscountInfo

IF Session("RS").eof THEN

  response.write "<UL>"
  response.write "<LI><FONT COLOR=""#FF0000"">" & Translate("Sorry, No Service Options match the search criteria you have entered.",Login_Language,conn) & "</FONT></LI>"
  response.write "</UL>"
	response.write "<UL>"
  response.write Translate("Click on [New Search] to enter new search criteria.",Login_Language,conn)
  response.write "</UL>"

else

  Model_Noun = Replace(Session("RS")("name")," ,",", ")
  
  response.write "<UL>"
  response.write "<LI>" & Translate("Click on the <FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;Back&nbsp;&nbsp;</FONT> button of your browser to return to the previous model list or click on <FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;New Search&nbsp;&nbsp;</FONT> to enter new search criteria.",Login_Language,conn) & "</LI><BR><BR>"
  response.write "<LI>" & Translate("Fluke Model Name",Login_language,conn) & " : <FONT COLOR=""#FF0000"">" &  Model_Noun & "</FONT></B></LI>"
  response.write "<LI>" & Translate("Oracle Item Number",Login_Language,conn) & " : <FONT COLOR=""#FF0000"">" &  mid(Session("RS")("pfid"),1,instr(1,Session("RS")("pfid"),"-")-1) & "</FONT></B></LI>"
  response.write "</UL>"

  Call WriteNavigation
    
  Response.write "<B>" & Translate("Standard Price Repair bundled with a Performance Check or Calibration Service Option",Login_Language,conn) & "</B></BR><BR>"

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
	
'	response.write "</TABLE></TD></TR></TABLE>"
	response.write "</TABLE>"
  Call Nav_Border_end

  ' Ala Carte Service Options
  
  Response.write "<BR><B>" & Translate("Al&aacute; Carte Service Options",Login_Language,conn) & "</B> (<FONT COLOR=""Gray"">" & Translate("Gray",Login_Language,conn) & "</FONT> " & Translate("Service Options are not to be quoted Al&aacute; Carte, but provided <U>only</U> as reference information.",Login_Language,conn) & ")<BR><BR>"

  Session("RS").MoveFirst  
  x=0
	Call WriteTableHeaders	    
	while ((not Session("RS").EOF) AND (x <= limit))
		Call WriteRecordsToPage_AlaCart
		Session("RS").MoveNext
		x=x+1
	wend
	
'	response.write "</TABLE></TD></TR></TABLE>"
	response.write "</TABLE>"
  Call Nav_Border_end


end if

Call WriteNavigation

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
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

end sub

' --------------------------------------------------------------------------------------

sub WriteTableHeaders

  Call Nav_Border_Begin
  response.write "<TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH=""100%"">"
  response.write "<TR>"
  response.write "<TD Class=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER WIDTH=""10%""><NOBR>" & Translate("Model Number",Login_Language,conn) & "<BR>" & Translate("Service Option",Login_Language,conn) & "</B></TD>"
  response.write "<TD CLASS=SmallBoldGold BGCOLOR=""#000000"">" & Translate("Service Option Description",Login_Language,conn) & "</FONT></TD>"
  response.write "<TD Class=SmallBoldGold BGCOLOR=""#000000"" ALIGN=CENTER WIDTH=""5%"">" & Translate("US List",Login_Language,conn) & "<BR>" & Translate("Price",Login_Language,conn) & "</B></FONT></TD>"
  response.write "</TR>"

end sub

' --------------------------------------------------------------------------------------

Sub WriteServiceOptions

   tempPartNumber = mid(Session("RS")("pfid"),instr(1,Session("RS")("pfid"),"-")+1)
   xOption=0
   While xOption <= PartNumberMax
    if cStr(tempPartNumber) = PartNumber(xOption) then
      response.write ServiceOption(xOption)
      xOption=PartNumberMax
    end if
    xOption=xOption + 1
   Wend
   
  End Sub

' --------------------------------------------------------------------------------------

Sub WriteRecordsToPage_Bundled

  response.write "<TR>"
	response.write "<TD Class=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"" NOWRAP>"
  response.write mid(Session("RS")("pfid"), 1, instr(1,Session("RS")("pfid"),"-")-1) & "-" & PartNumber(0) & "<BR>" & Session("RS")("pfid")
  response.write "</TD>"
  
  response.write "<TD Class=Small BGCOLOR=""#FFFFFF"">"
  if Session("RS")("short_description") <> "" then
    response.write "<FONT COLOR=""#FF0000"">" & Translate("Standard Price Repair",Login_Language,conn) & "</FONT> " & Translate("and",Login_Language,conn) & " <BR>"
    call WriteServiceOptions
	else
	  response.write "&nbsp;"
	end if
  response.write "</TD>"
  
	response.write "<TD Class=Small ALIGN=right BGCOLOR=""#FFFFFF"" VALIGN=MIDDLE NOWRAP>"
  ' N/A
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

  response.write "</TD>"
	response.write "</TR>"

end sub

' --------------------------------------------------------------------------------------

sub WriteRecordsToPage_AlaCart

	response.write "<TR>"
  
  if instr(1,Session("RS")("short_description"),"Repair") = 1 or Session("RS")("C3") = "99" then
	  response.write "<TD CLASS=Small BGCOLOR=""Silver"" ALIGN=""RIGHT"" NOWRAP>"
  else
	  response.write "<TD Class=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"" NOWRAP>"
  end if
        
  response.write Session("RS")("pfid")
  response.write "</TD>"

  if instr(1,Session("RS")("short_description"),"Repair") = 1 or Session("RS")("C3") = "99" then
	  response.write "<TD Class=Small BGCOLOR=""Silver"">"
  else
	  response.write "<TD Class=Small BGCOLOR=""#FFFFFF"">"
  end if

  if Session("RS")("short_description") <> "" then
    call WriteServiceOptions
    if not isnull(Session("RS")("C3")) or Session("RS")("C3") = "99" then
       response.write " <FONT COLOR=""#FF0000"">(" & Translate("One-Time Calibration",Login_Language,conn) & ")</FONT>"
    end if
  else
   response.write "&nbsp;"
  end if

  response.write "</TD>"

  if instr(1,Session("RS")("short_description"),"Repair") = 1 or Session("RS")("C3") = "99" then
	  response.write "<TD Class=Small BGCOLOR=""Silver"" ALIGN=""RIGHT"" NOWRAP>"
  else
	  response.write "<TD Class=Small BGCOLOR=""#FFFFFF"" ALIGN=""RIGHT"" NOWRAP>"
  end if

  if isnull(Session("RS")("list_price")) or (Session("RS")("list_price")) = "" then
    response.write "N / A"
  elseif cdbl(Session("RS")("list_price")) = 0 Then
    response.write "T &amp; M"
  elseIf cdbl(Session("RS")("list_price")) <> 0 THEN
  	response.write FormatNumber(Session("RS")("list_price") / 100)
  else
    response.write "&nbsp;"
  end if

  response.write "</TD>"
	response.write "</TR>"

end sub

' --------------------------------------------------------------------------------------

Sub WriteNavigation

  if Session("partpages") > 1 then
  	if request("Whatpage") <> "" then
  		pagenum=request("Whatpage")
  	else
  		pagenum=1
  	end if
  		
  	response.write "<UL>" & vbCrLf
    if pagenum=1 then
      
      response.write "<BR>"
      Call Nav_Border_Begin
      
  		response.write "<TABLE><TR>" & vbCrLf
  		Call WritePages
  		response.write "<TD CLASS=Normal>"
      response.write "&nbsp;<A HREF=""Part_Results.asp?view=" & view & "&whatpage=" & pagenum+1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>"
      response.write "&nbsp;&nbsp;"
      response.write "<A HREF="""
      response.write "RnC_QueryForm.asp"
      response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
  		response.write "</TD></TR></TABLE>" & vbCrLf

      Call Nav_Border_End
      response.write "<BR>"      

  		Session("rs").movefirst
  	else
  		if Cint(pagenum) = Cint(Session("partpages")) then

        response.write "<BR>"
        Call Nav_Border_Begin

  			response.write "<TABLE><TR><TD CLASS=NORMAL>"
        response.write "<A HREF=""Part_Results.asp?view=" & view & "&whatpage=" & pagenum-1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>"
        response.write "&nbsp;&nbsp;"
  			response.write "</TD>"
  			Call WritePages
        response.write "<TD CLASS=Normal>"
        response.write "<A HREF="""
        response.write "RnC_QueryForm.asp"
        response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"      
  			response.write "</TD></TR></TABLE>"

        Call Nav_Border_End
        response.write "<BR>"      

  		else

        response.write "<BR>"
        Call Nav_Border_Begin

  			response.write "<TABLE><TR><TD CLASS=Normal>"
        response.write "<A HREF=""Part_Results.asp?view=" & view & "&whatpage=" & pagenum-1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>"      
        response.write "&nbsp;&nbsp;"
  			response.write "</TD>"
  			Call WritePages
  			response.write "<TD CLASS=Normal>"
        response.write "&nbsp;<A HREF=""Part_Results.asp?view=" & view & "&whatpage=" & pagenum+1 & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>"      
        response.write "&nbsp;&nbsp;"
        response.write "<A HREF="""
        response.write "RnC_QueryForm.asp"
        response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
  			response.write "</TD></TR></TABLE>"

        Call Nav_Border_End
        response.write "<BR>"      

  		end if
  		Session("rs").movefirst
  		Session("rs").move limit*(pagenum-1)
  	end if

  else
  	if not session("rs").eof then
  		session("rs").movefirst
  	end if

    response.write "<BR>"
    Call Nav_Border_Begin

  	response.write "<TABLE><TR><TD CLASS=Normal>"
    response.write "<A HREF="""  
    response.write "RnC_QueryForm.asp"
    response.write "?view=" & view & "&Discount=" & strDiscount & "&rate=" & strRate & "&limit=" & limit & """><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("New Search",Login_Language,conn) & "&nbsp;&nbsp;</FONT></A>"
  	response.write "</TD>"
  	response.write "</TR></TABLE>" 

    Call Nav_Border_End
    response.write "<BR>"      

  end if

end sub

sub writepages

  response.write "<TD CLASS=NORMAL>"

	for i = 1 to Session("partpages")
'		if i=26 then
'			response.write "</TD>"
'			exit sub
'		end if
		if i = cint(pagenum) then
			response.write "<A HREF=""RnC_Results.asp?view=" & view & "&whatpage=" & i & """><FONT CLASS=NavLeftHighLight1>&nbsp;" & cstr(i) & "&nbsp;</FONT></A>&nbsp;"
		else
			response.write "<A HREF=""RnC_Results.asp?view=" & view & "&whatpage=" & i & """><FONT CLASS=NavTopHighLight>&nbsp;" & cstr(i) & "&nbsp;</FONT></A>&nbsp;"
		end if
	next

	response.write "</TD>"
	
end sub

%>
