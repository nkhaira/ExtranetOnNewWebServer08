<%
'
'  String Functions
'  Author:  Kelly Whitlock
'  Date:    3/17/2000

' --------------------------------------------------------------------------------------

Function stripHTML(strHTML)
  
  if not isblank(strHTML) then

    'Strips the HTML tags from strHTML
  
    Dim objRegExp, strOutput
    Set objRegExp = New Regexp
  
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = "<(.|\n)+?>"
  
    'Replace all HTML tag matches with the empty string
    'Convert <BR> and <P> to vbCrLf
    strHTML = Replace(Replace(strHTML,"<BR>",vbCrLf),"<br>",vbCrLf)
    strHTML = Replace(Replace(strHTML,"<P>",vbCrLf & vbCrLf),"<P>",vbCrLf & vbCrLf)
    strOutput = objRegExp.Replace(strHTML, "")
    
    'Replace all < and > with &lt; and &gt;
    strOutput = Replace(strOutput, "<", "&lt;")
    strOutput = Replace(strOutput, ">", "&gt;")
    
    stripHTML = strOutput    'Return the value of strOutput
    
    Set objRegExp = Nothing

  else  
    stripHTML = stripHTML    'Return the value of orignal string
  end if   

End Function

' --------------------------------------------------------------------------------------
' Used with Subscription Service body text rendering
' --------------------------------------------------------------------------------------

function Un_HTML(str)
	' we need to:
	' replace multiple space with one space
	' replace "vbcrlf " with vbcrlf
	Dim st1,st2,tstr
	
	if isblank(str) OR isnull(str) then
		Un_HTML = str
		Exit Function
	End if
	
	str = RestoreQuote(str)
	
	' replace vbcrlf with space (assume wrapped text)
	str = Replace(str,vbcrlf," ")
	' replace each <BR> with placeholeer
	str = Replace(str,"<BR>","$$")
	' replace each <P> with placeholeer
	str = Replace(str,"<P>","$$$$")
	' force all combinations of "li>" to uppercase
	str = Replace(str,"li>","LI>")
	str = Replace(str,"Li>","LI>")
	str = Replace(str,"lI>","LI>")
	' replace <LI> with "vbcrlf- "
	str = Replace(str,"<LI>",vbcrlf & "- ")
	' replace </LI> with vbcrlf
	str = Replace(str,"</LI>",vbcrlf)
	
	' now get the all the other tags (replace with nothing)
	st1 = Instr(str,"<")	
	if st1 > 0 then
		st2 = Instr(st1,str,">")
	Else
		st2 = 0
	End if
	While st2 > st1
		tstr = Mid(str,st1,st2-st1+1)
		str = Replace(str,tstr,"")
		st1 = Instr(str,"<")
		if st1 > 0 then
			st2 = Instr(st1,str,">")
		Else
			st2 = 0
		End if
	Wend
	
	' eliminate leading spaces on lines
	while Instr(str,vbcrlf & " ") > 0
		str = Replace(str,vbcrlf & " ",vbcrlf)
	Wend
	
	' eliminate double vbcrlf
	while Instr(str,vbcrlf & vbcrlf) > 0
		str = Replace(str,vbcrlf & vbcrlf,vbcrlf)
	Wend
	
	' replace each placeholeer with vbcrlf
	str = Replace(str,"$$",vbcrlf)
	
	' some leading spaces are creeping in
	str = Replace(str,vbcrlf & " ",vbcrlf)
	
	' eliminate multiple spaces
	while Instr(str,"  ") > 0
		str = Replace(str,"  "," ")
	Wend
	Un_HTML = str
End function

' --------------------------------------------------------------------------------------
    
function Decode_Key(User_Site, User_Account, User_Asset, User_Key)

  on error resume next
  Master_Key = ((2 * CInt(User_Site) + (CInt(CInt(User_Account)/2)) + (CInt(CInt(User_Asset)/3)))) + CInt(User_Asset)
  if err.number <> 0 then
    on error goto 0
    Decode_Key = False
  end if
  on error goto 0
  
  if (User_Site > 0 and User_Account > 0 and User_Asset > 0) and (CLng(Master_Key) = CLng(User_Key)) then
    Decode_Key = True
  else
    Decode_Key = False
  end if

end function

' --------------------------------------------------------------------------------------

function Encode_Key(User_Site, User_Account, User_Asset)

    Encode_Key = ((2 * CInt(User_Site) + (Cint(CInt(User_Account)/2)) + (CInt(Cint(User_Asset)/3)))) + CInt(User_Asset)
    
end function

' --------------------------------------------------------------------------------------

function Internet_Safe_Chr(str)

  Dim TempStr
  TempStr = str
  if not isnull(TempStr) and not isblank(TempStr) then
    TempStr = Replace(TempStr,"+","%20")
    TempStr = Replace(TempStr,"\n", vbCrLf)
'    TempStr = Replace(TempStr,vbCrLf,"%0D%0A") ' CRLF
    TempStr = Replace(TempStr,"\r\n","%0D%0A") ' CRLF    
    TempStr = Replace(TempStr,"\r", "%0D")     ' Carrage Return
    TempStr = Replace(TempStr,"\n", "%0A")     ' Line Feed
  end if
  
  Internet_Safe_Chr = TempStr

end function

' --------------------------------------------------------------------------------------

function FormatFullName(FirstName, MiddleName, LastName)

  Dim TempStr
  TempStr = ""
  
  if not isblank(FirstName) then
    TempStr = TempStr & FirstName
  end if
  
  if not isblank(MiddleName) then
    if len(Trim(MiddleName)) > 1 and instr(1,MiddleName,".") = 0 then
      TempStr = TempStr & " " & MiddleName
    end if
  end if
  
  if not isblank(LastName) then
    TempStr = TempStr & " " & LastName
  end if
  
  FormatFullName = TempStr
  
end function
      
' --------------------------------------------------------------------------------------

function FormatPhone(MyString)

  Dim tempstr, tempchr, tempout
  
  tempstr = MyString
  tempchr = 0
  tempout = ""
  
  if not isblank(Trim(tempstr)) then
  
    for x = 1 to len(tempstr)
    
      tempchr = UCase(mid(tempstr,x,1))
      tempasc = asc(tempchr)

      select case tempasc
        case 48,49,50,51,52,53,54,55,56,57
          tempout = tempout & tempchr        
        case 65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90
          select case tempchr
            case "A","B","C"
              tempout = tempout & "2"
            case "D","E","F"
              tempout = tempout & "3"
            case "G","H","I"
              tempout = tempout & "4"
            case "J","K","L"
              tempout = tempout & "5"
            case "M","N","O"
              tempout = tempout & "6"
            case "P","Q","R","S"
              tempout = tempout & "7"
            case "T","U","V"
              tempout = tempout & "8"
            case "W","X","Y","Z"
              tempout = tempout & "9"
           end select
      end select

    next
    
    if len(tempout) = 7 then  
      tempout = Mid(tempout,1,3) & "." & Mid(tempout,4,4)
    elseif len(tempout) = 10 then
      tempout = Mid(tempout,1,3) & "." & Mid(tempout,4,3) & "." & Mid(tempout,7,4)
    elseif len(tempout) = 11 then
      tempout = Mid(tempout,1,1) & "." & Mid(tempout,2,3) & "." & Mid(tempout,5,3) & "." & Mid(tempout,8,4)    
    else
      tempout = tempstr  
    end if
    
    FormatPhone = tempout

  else
    FormatPhone = ""
  end if    
  
end function

' --------------------------------------------------------------------------------------
' NOTE - this function is copied into subscription email tool!

function IsBlank(MyString)

  Dim TempString
  if isEmpty(MyString) then
    TempString = ""
  elseif  isNull(MyString) then  
    TempString = ""
  elseif isObject(MyString) then
  IsBlank = false
'    if not MyString then
'      TempString = ""
'    else
'      TempString = MyString
'    end if
  else
    TempString = MyString
  end if

  
  if isEmpty(TempString) then
    IsBlank = true
  elseif isNull(TempString) then
      IsBlank = true
  elseif Len(TempString) = 0 then
      IsBlank = true
  elseif Len(Trim(TempString & " ")) = 0 then
      IsBlank = true
  elseif UCase(Trim(TempString & " ")) = "NULL" then
      IsBlank = true
  elseif UCase(Trim(TempString & " ")) = "NO VALUE" then
      IsBlank = true
  elseif isObject(TempString) then
      IsBlank = false  
  else
      IsBlank = false
  end if
  
end function

' --------------------------------------------------------------------------------------

Function KillComma(str)
  KillComma = replace(str, ",","")
End Function    

' --------------------------------------------------------------------------------------

Function KillCrLf(str)

  Dim tempstr
  
  tempstr = replace(str,Chr(13),"")
  tempstr = replace(tempstr,Chr(10),"")
  KillCrLf = tempstr
End Function

' --------------------------------------------------------------------------------------

Function KillQuote(str)
  Dim TempStr
  TempStr = Str
  TempStr = replace(str, """", "")
  TempStr = replace(TempStr, "&rdquo;", "")
  TempStr = replace(TempStr, "&quot;", "")
  TempStr = replace(TempStr, "'", "")
  TempStr = replace(TempStr, "&acute;", "")  
  killquote = replace(TempStr, "&rsquo;","")
End Function    

' --------------------------------------------------------------------------------------
' SQL Statements do not like ' in Strings, replace with acute and restore
' --------------------------------------------------------------------------------------
' NOTE - this function is copied into subscription email tool!

Function ReplaceQuote(str)
  Dim TempStr
  TempStr = str
  if not isnull(TempStr) and not isblank(TempStr) then
    TempStr      = replace(TempStr,"'","&acute;")
    TempStr      = replace(TempStr,"&rsquo;","&acute;")
    ReplaceQuote = replace(TempStr,"&rdquo;","&quot;")
  else
    ReplaceQuote = TempStr
  end if    
End Function

' --------------------------------------------------------------------------------------

Function ReplaceRSQuote(str)
  Dim TempStr
  TempStr = str
  if not isnull(TempStr) and not isblank(TempStr) then
    TempStr            = replace(TempStr,"&acute;", "'")
    ReplaceRSQuote = replace(TempStr,"&rsquo;", "'")    
  else
    ReplaceRSQuote = TempStr
  end if   
End Function

' --------------------------------------------------------------------------------------

Function RestoreQuote(str)
  Dim TempStr
  TempStr = str
  if not isnull(TempStr) and not isblank(TempStr) then
    TempStr      = replace(TempStr,"&acute;","'")
    TempStr      = replace(TempStr,"&rsquo;","'")
    TempStr      = replace(TempStr,"&quot;","'")
    RestoreQuote = replace(TempStr,"&rdquo;","'")    
  else
    RestoreQuote = TempStr
  end if  
End Function

' --------------------------------------------------------------------------------------

Function FormatNumberFloat(strNumber,DecimalPlaces)

  Dim tempstr
  Dim IntegerPart
  Dim DecimalPart
  Dim DecimalPosition
  
  tempstr = CStr(strNumber)
  if DecimalPlaces > 10 then DecimalPlaces = 10
  if DecimalPlaces <  0 then DecimalPlaces =  0
  
  DecimalPosition = instr(1,tempstr,".")
  
  if DecimalPosition > 0 then    
    IntegerPart = mid(tempstr,1,DecimalPosition -1)
    if DecimalPlaces > 0 then IntegerPart = IntegerPart & "."
    DecimalPart = mid(mid(tempstr,DecimalPosition +1),1,DecimalPlaces)
  else
    IntegerPart = tempstr
    DecimalPart = ""    
  end if
  
  DecimalPart = DecimalPart & mid("0000000000",1,DecimalPlaces - Len(DecimalPart))

  FormatNumberFloat = IntegerPart & DecimalPart

end function

' --------------------------------------------------------------------------------------

Function FormatDate(iDateFormat, TempDate)

  Dim strDay
  Dim strMonth
  
  if Len(DatePart("d", TempDate)) = 1 then
  	strDay = "0" & DatePart("d", TempDate)
  else
  	strDay = DatePart("d", TempDate)
  end if
  
  if Len(DatePart("m", TempDate)) = 1 then
  	strMonth = "0" & DatePart("m", TempDate)
  else
  	strMonth = DatePart("m", TempDate)
  end if
  
  if iDateFormat = 1 then
  	FormatDate = strMonth & "/" & strDay & "/" & DatePart("yyyy", TempDate)
  elseif iDateFormat = 2 then
  	FormatDate = strDay & "/" & strMonth & "/" & DatePart("yyyy", TempDate)
  else
  	FormatDate = DatePart("yyyy", TempDate) & "/" & strMonth & "/" & strDay
  end if

end function

' --------------------------------------------------------------------------------------

Function FormatDateText(iDateFormat, TempDate)

  Dim strDay
  Dim strMonth
  Months="January,February,March,April,May,June,July,August,September,October,November,December"
  arrayMonths = Split(Months,",")
 
  if Len(DatePart("d", TempDate)) = 1 then
  	strDay = "0" & DatePart("d", TempDate)
  else
  	strDay = DatePart("d", TempDate)
  end if
  
  if DatePart("m", TempDate) < 1 or DatePart("m", TempDate) > 12 then
  	strMonth = "Invalid Month"
  else
  	strMonth = arrayMonths(DatePart("m",TempDate)-1)
  end if
  
  if iDateFormat = 1 then
  	FormatDateText = strMonth & ", " & strDay & "," & DatePart("yyyy", TempDate)
  elseif iDateFormat = 2 then
  	FormatDateText = strDay & " " & strMonth & " " & DatePart("yyyy", TempDate)
  else
  	FormatDateText = DatePart("yyyy", TempDate) & " " & strMonth & " " & strDay
  end if

end function

' --------------------------------------------------------------------------------------

function ProperCase(strInput)

  if not isnull(strInput) and not isblank(strInput) then
  
    Dim iPosition ' Our current position in the string (First character = 1)
    Dim iSpace    ' The position of the next space after our iPosition
    Dim strOutput ' Our temporary string used to build the function's output
  
    iPosition = 1
  
    do while InStr(iPosition, strInput, " ", 1) <> 0
    
      iSpace = InStr(iPosition, strInput, " ", 1)
      
      myWord = Mid(strInput,iPosition,iSpace - iPosition)
      
      ' Check for numeric value in word or word is already capitalized, if True capitalize entire word (Typically Model Noun)

      ContainsNumber = False
      for x = 1 to len(myWord)
        if isnumeric(Mid(myWord,x,1)) then
          ContainsNumber = True
          exit for
        elseif ASC(Mid(myWord,x,1)) >= 65 and ASC(Mid(myWord,x,1)) <= 90 then
          if x+1 <= len(myword) then
            if ASC(Mid(myWord,x+1,1)) >= 65 and ASC(Mid(myWord,x+1,1)) <= 90 then
              ContainsNumber = True
              exit for
            end if
          end if
        end if
      next
      
      if ContainsNumber = False then   
        strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
        strOutput = strOutput & LCase(Mid(strInput, iPosition + 1, iSpace - iPosition))
      else
        strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
        strOutput = strOutput & UCase(Mid(strInput, iPosition + 1, iSpace - iPosition))
      end if  
      iPosition = iSpace + 1
    loop
  
    myWord = Mid(strInput,iPosition)  
      
    ContainsNumber = False
    for x = 1 to len(myWord)
      if isnumeric(Mid(myWord,x,1)) then
        ContainsNumber = True
        exit for
      end if
    next
    
    if ContainsNumber = False then   
      strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
      strOutput = strOutput & LCase(Mid(strInput, iPosition + 1))
    else
      strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
      strOutput = strOutput & UCase(Mid(strInput, iPosition + 1))
    end if

  else
    strOutput = strInput
  end if  

  ProperCase = strOutput
  
end function

' --------------------------------------------------------------------------------------

function Highlight_Keyword(strInput,Keyword, Highlight_Style)

  if not isblank(Trim(strInput)) and not isblank(Trim(Keyword)) then

    if instr(1,LCase(strInput),LCase(Keyword)) > 0 then

      if instr(1,Highlight_Style,"#") = 1 then
        Highlight_Tag = "<FONT COLOR=""" & Highlight_Style & """>"
      else
        Highlight_Tag = "<SPAN CLASS=""" & Highlight_Style & """>"
      end if
      
      Dim iPosition ' Current position in the string (First character = 1 of Keyword)
      Dim iLength   ' Length of Keyword String
      Dim iPointer  ' The position of the next Keyword after our iPosition
      Dim strOutput ' Temporary string used to build the function's output

      iPosition = 1
      iPointer  = 1
      strOutput = ""
      iLength   = len(Keyword)

      do while InStr(iPointer, LCase(strInput), LCase(Keyword)) <> 0

        iPosition = InStr(iPointer, LCase(strInput), LCase(Keyword))
        if iPointer < iPosition then
          strOutput = strOutput & mid(strInput,iPointer,iPosition - iPointer)
        end if
        strOutput = strOutput & Highlight_Tag
        strOutput = strOutput & Mid(strInput, iPosition, iLength)
        if instr(1,Highlight_Style,"#") = 1 then
          strOutput = strOutput & "</FONT>"
        else
          strOutput = strOutput & "</SPAN>"
        end if  
        iPointer  = iPosition + iLength

      loop

      strOutput = strOutput & Mid(strInput,iPointer)
      
    else
      strOutput = strInput
    end if
  else
    strOutput = strInput
  end if

  Highlight_Keyword = strOutput

end function
%>