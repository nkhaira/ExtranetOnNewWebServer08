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

Dim TypeOfSearch
Dim TempPartType
Dim strRate
Dim strPart
Dim strDiscount
Dim strInvalid
Dim strStrip
Dim strview
Dim region
Dim dept_id

Dim Site_ID

Site_ID = Request("Site_ID")
	
' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

if request("Returned") <> "" then
  limit = Cint(request("Returned"))
else
	limit = 25
end if

strInvalid = ""

Call VerifyRate
Call VerifyDiscount
Call VerifyDept_ID

if Request("Part") <> "" then
	Call CheckValidity
else
  if Session("LimitView") = false then
  	strInvalid = "<LI>" & Translate("Please enter a 6, 7 or 12-digit valid Fluke part number or a portion of that part number using the wild card *, or Model Noun.",Login_Language,conn) & "</LI>"
  else  
  	strInvalid = "<LI>" & Translate("Please enter a 6, 7 or 12-digit valid Fluke part number or a portion of that part number using the wild card *.",Login_Language,conn) & "</LI>"
  end if  
end if

if strInvalid <> "" then

	Session("ErrorString") = strInvalid
  Session("view")        = strview
  Session("StrRate")     = strRate
  Session("StrDiscount") = strDiscount
  Session("Perpage")     = limit
  Session("part")        = strPart
  Session("dept_id")     = dept_id

	response.redirect ("parterror.asp")

end if

Call ReplaceWildCards

Call connect_parts

if tempPartType = true then
  if dept_id = 20 then
    SQL = "SELECT * FROM vcturbo_product_family WHERE vcturbo_product_family.dept_id=" & dept_id & " AND vcturbo_product_family.pfid " & TypeOfSearch & " ORDER BY vcturbo_product_family.name"
  else
    SQL = "SELECT * FROM vcturbo_product_family WHERE vcturbo_product_family.dept_id=" & dept_id & " AND vcturbo_product_family.pfid " & TypeOfSearch & " ORDER BY vcturbo_product_family.pfid"
  end if  
elseif tempPartType = false then
  if dept_id = 20 then
    SQL = "SELECT * FROM vcturbo_product_family WHERE vcturbo_product_family.dept_id=" & dept_id & " AND vcturbo_product_family.short_description " & TypeOfSearch  & " ORDER BY vcturbo_product_family.name"  
  else
    SQL = "SELECT * FROM vcturbo_product_family WHERE vcturbo_product_family.dept_id=" & dept_id & " AND vcturbo_product_family.short_description " & TypeOfSearch  & " ORDER BY vcturbo_product_family.pfid"  
  end if      
end if

Set Session("rs") = Server.CreateObject("ADODB.Recordset")
Session("rs").open SQL,DBConn,3,1,1

if limit > Session("rs").recordcount then
	numpages = 1
else
	numpages = Session("rs").recordcount \ limit
	numpages = cint(numpages)
	extrapage = Session("rs").recordcount MOD limit
	if extrapage > 0 then
		numpages = numpages + 1
	end if
end if

if numpages > 25 then
		numpages = 25
end if

Session("view")        = strview
Session("StrRate")     = strRate
Session("StrDiscount") = strDiscount
Session("Perpage")     = limit
Session("partpages")   = numpages
Session("part")        = strPart
Session("dept_id")     = dept_id

response.redirect ("partresults.asp?whatpage=1")

' --------------------------------------------------------------------------------------
' Subroutines
' --------------------------------------------------------------------------------------
	
sub VerifyRate
	if Request("rate") <> "" then
		strRate = Request("rate")
		if NOT IsNumeric(strRate) then
      strRate    = ""
			strInvalid = "<LI>" & Translate("The Exchange Rate that you have specified, must be numeric and in the format of 0.00 to US dollars.",Login_Language,conn) & "</LI>"
		end if
	else
		strRate = 1
	end if
end sub

' --------------------------------------------------------------------------------------

sub VerifyDiscount
	if Request("discount") <> ""  then
		strDiscount = Request("discount")
		if NOT IsNumeric(strDiscount) then
      strDiscount = ""
			strInvalid  = "<LI>" & Translate("The Discount from US List Price that you have specified, must be numeric and in the format of 00 (percent).",Login_Language,conn) & "</LI>"
		end if 
	else
		strDiscount = 100
	end if 
end sub

' --------------------------------------------------------------------------------------

sub VerifyDept_ID
	if Request("dept_id") <> "" then
		dept_id = Request("dept_id")
		if NOT IsNumeric(dept_id) then
      dept_id    = ""
			strInvalid = "<LI>" & Translate("The Department ID that you have specified, must be numeric.",Login_Language,conn) & "</LI>"
		end if
	else
		dept_id = 1
	end if
end sub

' --------------------------------------------------------------------------------------

sub CheckValidity

   	'=======MAKE SURE ITS NUMERIC OR NUMERIC WITH WILDCARDS OR KEY WORD======
    
		tempPart = Replace(Request("Part"), "*", "") 		'remove all the *
    tempPart = Replace(tempPart,"-","")             'remove all dashes    
    tempPart = Replace(tempPart," ","")             'remove all spaces

    if tempPart = ""  then
      if Session("LimitView") = false then
  			strInvalid = "<LI>" & Translate("The Fluke part number that you have specified must be a 6, 7 or 12-digit valid part number or a portion of that part number using the wild card *, or a Key Word.",Login_Language,conn) & "</LI>"
      else
  			strInvalid = "<LI>" & Translate("The Fluke part number that you have specified must be a 6, 7 or 12-digit valid part number or a portion of that part number using the wild card *.",Login_Language,conn) & "</LI>"        
      end if  
		elseif IsNumeric(tempPart) then
      tempPartType = true  ' Numeric Part Number
    elseif Not IsNumeric(tempPart) then
      tempPartType = false ' Alpha Key WordNumeric Part
		end if
		
		if strInvalid = "" and tempPartType = true then 'no errors so far for numeric

			'=====VERifY THAT THE WILDCARDS ARE AT EITHER END OF THE STRING====

			tempPart= Request("Part")			

			if Left(tempPart, 1) = "*" then
				tempPart=Right(tempPart, (Len(tempPart)-1)) 	'pull off the right most char
			end if
			
			if Right(tempPart, 1) = "*" then
				tempPart=Left(tempPart, (Len(tempPart)-1)) 	'pull off the left most char
			end if

      tempPart = Replace(tempPart,"-","")             'remove all dashes    
      tempPart = Replace(tempPart," ","")             'remove all spaces
			
			if NOT IsNumeric(tempPart) then
        strInvalid="<LI>" & Translate("The wild card * must be at the beginning or end of the part number.",Login_Language,conn) & "</LI>"
      end if        

    elseif strInvalid = "" and tempPartType = false then 'no errors so far for alpha key word

			tempPart= Request("Part")			

			if Left(tempPart, 1) = "*" then
				tempPart=Right(tempPart, (Len(tempPart)-1)) 	'pull off the right most char
			end if
			
			if Right(tempPart, 1) = "*" then
				tempPart=Left(tempPart, (Len(tempPart)-1)) 	'pull off the left most char
			end if
			
			if instr(1,tempPart,"*") > 0 then
        strInvalid="<LI>" & Translate("The wild card * cannot be used in a Key Word Search.",Login_Language,conn) & "</LI>"
      end if
      
      if Session("LimitView") = true then
        if not isnumeric(tempPart) then        
          strInvalid = "<LI>" & Translate("The Fluke part number that you have specified must be a 6, 7 or 12-digit valid part number or a portion of that part number using the wild card *, and cannot contain alpha characters.",Login_Language,conn) & "</LI>"
        end if
      end if            
		end if

end sub

' --------------------------------------------------------------------------------------

sub ReplaceWildCards

  if tempPartType = false then   ' Alpha always uses pre and post wild cards so strip users additions, if present

		strPart = Replace(Request("Part"), "*", "")
    strPart = Replace(strPart,chr(34),"")             'Remove Double Quotes
    strPart = Replace(strPart,chr(39),"")             'Remove Single Quotes    
  	TypeOfSearch = "LIKE " & "'%" & strPart & "%'"

  elseif tempPartType = true then                      ' Numeric

  	if InStr(Request("Part"), "*") <> 0 then

  		'======HAS WILDCARDS(*)-(USE 'PART LIKE' IN SQL)=========
  		strPart = Replace(Request("Part"), "*", "%")
      strPart = Replace(strPart,"-","")               'remove all dashes    
      strPart = Replace(strPart," ","")               'remove all spaces       		
    	TypeOfSearch = "LIKE " & "'" & strPart & "'"

  	else

  		'======NO WILDCARDS-(USE 'PART =' IN SQL)===============
  		strPart = Request("Part")
      strPart = Replace(strPart,"-","")               'remove all dashes    
      strPart = Replace(strPart," ","")               'remove all spaces       		
  		TypeOfSearch = " = " & strPart
  	end if

  end if

end sub

' --------------------------------------------------------------------------------------
%>

