
<!-- #include virtual="/connections/connections_parts.asp" -->

<%

' This script is the same as prtQuery.asp, however opens a second record set for
' Replaced By or Equivalent, so that user may return back to the original recordset data.
' When modifying this script, ensure that prtQuery.asp is also updated where appropriate.

Dim TypeOfSearch
Dim TempPartType
Dim strRate
Dim strPart
Dim strDiscount
Dim strInvalid
Dim strStrip
Dim strview
Dim region
Dim rt
	
strview = Request("view")
region = Request("region")
rt = Request("rt")

Sub VerifyRate
	IF Request("rate") <> "" THEN
		strRate = Request("rate")
		IF NOT IsNUmeric(strRate) THEN
      strRate = ""
			strInvalid="<LI>The Exchange Rate that you have specified, must be numeric and in the format of 0.00 to US dollars.</LI>"
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
			strInvalid="<LI>The Discount from US List Price that you have specified, must be numeric and in the format of 00 (percent)."
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
			strInvalid = "<LI>The Fluke part number that you have specified must be a 6, 7 or 12-digit valid part number or a portion of that part number using the wild card <FONT SIZE=2 COLOR=""#FF0000"">*</FONT>, or a Key Word.</LI>"
		ELSEIF IsNumeric(tempPart) THEN
      tempPartType = true  ' Numeric Part Number
    ELSEIF Not IsNumeric(tempPart) THEN
      tempPartType = false ' Alpha Key WordNumeric Part
		END IF
		
		IF strInvalid = "" and tempPartType = true then 'no errors so far for numeric

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
        strInvalid="<LI>The wild card <FONT SIZE=2 COLOR=""#FF0000"">*</FONT> must be at the beginning or end of the part number.</LI>"
      END IF        

    ELSEIF strInvalid = "" and tempPartType = false then 'no errors so far for alpha key word

			tempPart= Request("Part")			

			IF Left(tempPart, 1) = "*" THEN
				tempPart=Right(tempPart, (Len(tempPart)-1)) 	'pull off the right most char
			END IF
			
			IF Right(tempPart, 1) = "*" THEN
				tempPart=Left(tempPart, (Len(tempPart)-1)) 	'pull off the left most char
			END IF
			
			IF instr(1,tempPart,"*") > 0 THEN
        strInvalid="<LI>The wild card <FONT SIZE=2 COLOR=""#FF0000"">*</FONT> cannot be used in a Key Word Search.</LI>"
      END IF        

		END IF

End Sub

Sub ReplaceWildCards

  IF tempPartType = false then   ' Alpha always uses pre and post wild cards so strip users additions, if present

		strPart = Replace(Request("Part"), "*", "")
    strPart = Replace(strPart,chr(34),"")             'Remove Double Quotes
    strPart = Replace(strPart,chr(39),"")             'Remove Single Quotes    
  	TypeOfSearch = "LIKE " & "'%" & strPart & "%'"

  ELSEIF tempPartType = true then                      ' Numeric

  	IF InStr(Request("Part"), "*") <> 0 THEN

  		'======HAS WILDCARDS(*)-(USE 'PART LIKE' IN SQL)=========
  		strPart = Replace(Request("Part"), "*", "%")
      strPart = Replace(strPart,"-","")               'remove all dashes    
      strPart = Replace(strPart," ","")               'remove all spaces       		
    	TypeOfSearch = "LIKE " & "'" & strPart & "'"

  	ELSE

  		'======NO WILDCARDS-(USE 'PART =' IN SQL)===============
  		strPart = Request("Part")
      strPart = Replace(strPart,"-","")               'remove all dashes    
      strPart = Replace(strPart," ","")               'remove all spaces       		
  		TypeOfSearch = " = " & strPart
  	END IF

  END IF

End Sub

'=============================================subs========

'start

	if request("Returned") <> "" then
		limit=Cint(request("Returned"))
	else
		limit=25
	end if

strInvalid=""

Call VerifyRate
Call VerifyDiscount

IF Request("Part") <> "" THEN
	Call CheckValidity
ELSE
	strInvalid = "<LI>Please enter a 6, 7 or 12-digit valid Fluke part number or a portion of that part number using the wild card <FONT SIZE=2 COLOR=""#FF0000"">*</FONT></LI>"  
END IF


IF strInvalid <> "" THEN
	Session("ErrorString") = strInvalid
  Session("view")=strview
  Session("StrRate")=strRate
  Session("StrDiscount")=strDiscount
  Session("Region") = region
  Session("ePerpage")=limit
  Session("part")=strPart
	response.redirect ("parterror.asp")
END IF

Call ReplaceWildCards

Call connect_parts
'Set DBConn2 = Server.CreateObject("ADODB.Connection")
'DBConn2.Open strConnectionString_Parts

if tempPartType = true then
 SQL = "SELECT * FROM vcturbo_product_family WHERE pfid " & TypeOfSearch & " ORDER BY pfid"
elseif tempPartType = false then
 SQL = "SELECT * FROM vcturbo_product_family WHERE short_description " & TypeOfSearch  '& " ORDER BY pfid"  
end if

'response.write SQL

Set equivRS = Server.CreateObject("ADODB.Recordset")
'Set Session("equivRS") = Server.CreateObject("ADODB.Recordset")
equivRS.Open SQL,strConnectionString_Parts,3,1,1

if limit > equivRS.recordcount then
	numpages=1
else
	numpages=equivRS.recordcount \ limit
	numpages=cint(numpages)
	extrapage = equivRS.recordcount MOD limit
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
Session("ePerpage")=limit
Session("ePartpages")=numpages
Session("region")=region
Session("rt") = rt

Set Session("equivRS") = equivRS
Set equivRS = Nothing

response.redirect ("equivresults.asp?rt=" & rt & "&whatpage=1")
	
%>

