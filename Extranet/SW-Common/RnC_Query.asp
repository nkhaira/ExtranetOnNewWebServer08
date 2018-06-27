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

%>
<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->
<%

Dim TypeOfSearch
Dim TempPartType
Dim strRate
Dim strPart
Dim strDiscount
Dim ErrorString
Dim strStrip
Dim strview
Dim region
	
strview = Request("view")
region = Request("region")

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

if not isblank(request("Limit")) then
  limit = Cint(request("Limit"))
else
	limit = 25
end if

LimitView = Session("LimitView")
ErrorString = ""

Call VerifyRate
Call VerifyDiscount
Call ReplaceWildCards
Call CheckValidity

if not isblank(ErrorString) then
  Session("view")        = strview
  Session("StrRate")     = strRate
  Session("StrDiscount") = strDiscount
  Session("Perpage")     = limit
  Session("ErrorString") = ErrorString
  
  response.redirect ("RnC_QueryForm.asp")
end if

Call Connect_Parts

if tempPartType = true then
 SQL = "SELECT * FROM vcturbo_product_family WHERE pfid " & TypeOfSearch & " ORDER BY vcturbo_product_family.short_description DESC"
elseif tempPartType = false then
 SQL = "SELECT * FROM vcturbo_product_family WHERE short_description " & TypeOfSearch  & " ORDER BY vcturbo_product_family.short_description DESC"  
end if

'response.write SQL

Set Session("rs") = Server.CreateObject("ADODB.Recordset")
Session("rs").open SQL,DBConn,3,1,1

if limit > Session("rs").recordcount then
	numpages=1
else
	numpages=Session("rs").recordcount \ limit
	numpages=cint(numpages)
	extrapage = Session("rs").recordcount MOD limit
	if extrapage > 0 then
		numpages=numpages + 1
	end if
end if

if numpages > 25 then
		numpages = 25
end if

Session("view")=strview
Session("StrRate")=strRate
Session("StrDiscount")=strDiscount
Session("Perpage")=limit
Session("partpages")=numpages
Session("region")=region

Call Disconnect_Sitewide
'Call Disconnect_Parts

response.redirect ("/SW-Common/RnC_Results.asp?whatpage=1")

' --------------------------------------------------------------------------------------
' Subroutines
' --------------------------------------------------------------------------------------

Sub VerifyRate
	IF Request("rate") <> "" THEN
		strRate = Request("rate")
		IF NOT IsNUmeric(strRate) THEN
      strRate = ""
			ErrorString="<LI>" & Translate("The Exchange Rate that you have specified, must be numeric and in the format of 0.00 to US dollars.",Login_Language,conn) & "</LI>"
		END IF
	ELSE
		strRate = 1
	END IF
End Sub

Sub VerifyDiscount
	IF Request("discount") <> ""  THEN
		strDiscount = Request("discount")
		IF NOT IsNumeric(strDiscount) THEN
      strDiscount = ""
			ErrorString="<LI>" & Translate("The Discount from US List Price that you have specified, must be numeric and in the format of 00 (percent).",Login_Language,conn)
		END IF 
	ELSE
		strDiscount=100
	END IF 
End Sub

Sub CheckValidity

   	'=======MAKE SURE ITS NUMERIC OR NUMERIC WITH WILDCARDS OR KEY WORD======
    
		tempPart = Replace(Request("Part"), "*", "") 		'remove all the *
    tempPart = Replace(tempPart,"-","")             'remove all dashes    
    tempPart = Replace(tempPart," ","")             'remove all spaces

    IF tempPart = ""  THEN
			ErrorString = "<LI>" & Translate("The Fluke part number that you have specified must be a 6, 7 or 12-digit valid part number or a portion of that part number using the wild card *, or a Key Word.",Login_Language,conn) & "</LI>"
		ELSEIF IsNumeric(tempPart) THEN
      tempPartType = true  ' Numeric Part Number
    ELSEIF Not IsNumeric(tempPart) THEN
      tempPartType = false ' Alpha Key WordNumeric Part
		END IF
		
		IF ErrorString = "" and tempPartType = true then 'no errors so far for numeric

			'=====VERIFY THAT THE WILDCARDS ARE AT EITHER END OF THE STRING====

			tempPart= Request("Part")			

			IF Left(tempPart, 1) = "*" THEN
				tempPart=Right(tempPart, (Len(tempPart)-1)) 	'pull off the right most char
			END IF
			
			IF Right(tempPart, 1) = "*" THEN
				tempPart=Left(tempPart, (Len(tempPart)-1)) 	'pull off the left most char
			END IF

      tempPart = Replace(tempPart,"-","")             'remove all dashes    
      tempPart = Replace(tempPart," ","")             'remove all spaces
			
			IF NOT IsNumeric(tempPart) THEN
        ErrorString="<LI>" & Translate("The wild card * must be at the beginning or end of the part number.",Login_Language,conn) & "</LI>"
      END IF        

    ELSEIF ErrorString = "" and tempPartType = false then 'no errors so far for alpha key word

			tempPart= Request("Part")			

			IF Left(tempPart, 1) = "*" THEN
				tempPart=Right(tempPart, (Len(tempPart)-1)) 	'pull off the right most char
			END IF
			
			IF Right(tempPart, 1) = "*" THEN
				tempPart=Left(tempPart, (Len(tempPart)-1)) 	'pull off the left most char
			END IF
			
			IF instr(1,tempPart,"*") > 0 THEN
        ErrorString="<LI>" & Translate("The wild card * cannot be used in a Key Word Search.",Login_Language,conn) & "</LI>"
      END IF        

		END IF

End Sub

Sub ReplaceWildCards

  IF tempPartType = false then   ' Alpha always uses pre and post wild cards so strip users additions, if present

		strPart = Replace(Request("Part"), "*", "")
    strPart = Replace(strPart,chr(34),"")             'Remove Double Quotes
    strPart = Replace(strPart,chr(39),"")             'Remove Single Quotes    
  	TypeOfSearch = "LIKE " & "'" & strPart & "-%'"

  ELSEIF tempPartType = true then                      ' Numeric

  	IF InStr(Request("Part"), "*") <> 0 THEN

  		'======HAS WILDCARDS(*)-(USE 'PART LIKE' IN SQL)=========
  		strPart = Replace(Request("Part"), "*", "%")
      strPart = Replace(strPart,"-","")               'remove all dashes    
      strPart = Replace(strPart," ","")               'remove all spaces       		
    	TypeOfSearch = "LIKE " & "'" & strPart & "-%'"

  	ELSE

  		'======NO WILDCARDS-(USE 'PART =' IN SQL)===============
  		strPart = Request("Part")
      strPart = Replace(strPart,"-","")               'remove all dashes    
      strPart = Replace(strPart," ","")               'remove all spaces       		
  		TypeOfSearch = "LIKE " & "'" & strPart & "-%'"
  	END IF

  END IF

End Sub
%>

